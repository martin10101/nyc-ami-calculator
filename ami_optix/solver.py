﻿import pandas as pd
from ortools.sat.python import cp_model
import itertools
import copy
import time
from typing import Dict, List, Any, Optional

def _calculate_waami_from_assignments(assignments: List[Dict[str, Any]]) -> float:
    """Calculates the WAAMI from a list of assignment dictionaries using integer arithmetic."""
    if not assignments:
        return 0.0

    total_sf_int = sum(int(unit['net_sf'] * 100) for unit in assignments)
    if total_sf_int == 0:
        return 0.0

    total_ami_sf_scaled = sum(
        int(unit['net_sf'] * 100) * int(unit['assigned_ami'] * 10000)
        for unit in assignments
    )
    return (total_ami_sf_scaled / total_sf_int) / 10000

def _get_bands_from_assignments(assignments: List[Dict[str, Any]]) -> List[int]:
    """Derives the unique AMI bands used in a set of assignments."""
    if not assignments:
        return []
    used_bands = {int(round(u['assigned_ami'] * 100)) for u in assignments}
    return sorted(list(used_bands))

def _build_metrics(assignments: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Compute reusable metrics for a scenario."""
    total_sf = sum(float(u['net_sf']) for u in assignments) or 0.0
    band_stats = {}
    for unit in assignments:
        band = int(round(unit['assigned_ami'] * 100))
        stats = band_stats.setdefault(band, {'band': band, 'units': 0, 'net_sf': 0.0})
        stats['units'] += 1
        stats['net_sf'] += float(unit['net_sf'])
    band_mix = []
    for band in sorted(band_stats):
        stats = band_stats[band]
        share = (stats['net_sf'] / total_sf) if total_sf else 0.0
        band_mix.append({**stats, 'share_of_sf': share})
    revenue_score = sum(float(u['net_sf']) * float(u['assigned_ami']) for u in assignments)
    return {
        'total_units': len(assignments),
        'total_sf': total_sf,
        'revenue_score': revenue_score,
        'waami_percent': _calculate_waami_from_assignments(assignments) * 100,
        'band_mix': band_mix,
        'sf_at_40_band': sum(float(u['net_sf']) for u in assignments if int(round(u['assigned_ami'] * 100)) <= 40),
    }

def calculate_premium_scores(df: pd.DataFrame, dev_preferences: Dict[str, Any]) -> pd.DataFrame:
    df_norm = df.copy()
    weights = dev_preferences['premium_score_weights']
    if 'balcony' not in df_norm.columns:
        df_norm['balcony'] = 0
    else:
        df_norm['balcony'] = df_norm['balcony'].apply(lambda x: 1 if pd.notna(x) and str(x).lower() not in ['false', '0', 'no', ''] else 0)
    for col in ['floor', 'net_sf', 'bedrooms', 'balcony']:
        if col in df_norm.columns:
            if df_norm[col].max() > df_norm[col].min():
                df_norm[f'{col}_norm'] = (df_norm[col] - df_norm[col].min()) / (df_norm[col].max() - df_norm[col].min())
            else:
                df_norm[f'{col}_norm'] = 0
        else:
            df_norm[f'{col}_norm'] = 0
    df['premium_score'] = (
        df_norm['floor_norm'] * weights['floor'] +
        df_norm['net_sf_norm'] * weights['net_sf'] +
        df_norm['bedrooms_norm'] * weights['bedrooms'] +
        df_norm['balcony_norm'] * weights['balcony']
    )
    return df

def _solve_single_scenario(df_affordable: pd.DataFrame, bands_to_test: List[int], total_affordable_sf: float, optimization_rules: Dict[str, Any]) -> Dict[str, Any]:
    bands_to_test = [band for band in bands_to_test if band != 50]
    if not bands_to_test:
        return {"status": "NO_SOLUTION"}
    waami_cap_basis_points = int(optimization_rules['waami_cap_percent'] * 100)
    bands_basis_points = [int(b * 100) for b in bands_to_test]
    sf_coeffs_int = (df_affordable['net_sf'] * 100).astype(int)
    total_sf_int = int(sf_coeffs_int.sum())
    model = cp_model.CpModel()
    num_units = len(df_affordable)
    num_bands = len(bands_to_test)
    x = [[model.NewBoolVar(f'x_{i}_{j}') for j in range(num_bands)] for i in range(num_units)]
    for i in range(num_units):
        model.AddExactlyOne(x[i])
    total_ami_sf_expr = sum(
        sum(x[i][j] * bands_basis_points[j] for j in range(num_bands)) * sf_coeffs_int.iloc[i]
        for i in range(num_units)
    )
    max_waami_scaled = waami_cap_basis_points * total_sf_int
    total_ami_sf_var = model.NewIntVar(0, max_waami_scaled, 'total_ami_sf_var')
    model.Add(total_ami_sf_var == total_ami_sf_expr)
    waami_floor_percent = optimization_rules.get('waami_floor')
    if waami_floor_percent:
        waami_floor_basis_points = int(waami_floor_percent * 100)
        min_waami_scaled = waami_floor_basis_points * total_sf_int
        model.Add(total_ami_sf_var >= min_waami_scaled)
    deep_affordability_threshold = optimization_rules.get('deep_affordability_sf_threshold', 10000)
    if total_affordable_sf >= deep_affordability_threshold:
        low_band_indices = [j for j, band in enumerate(bands_to_test) if band <= 40]
        if low_band_indices:
            low_band_assignments = []
            for i in range(num_units):
                for j in low_band_indices:
                    low_band_assignments.append(x[i][j] * sf_coeffs_int.iloc[i])
            min_share = optimization_rules.get('deep_affordability_min_share', 0.2)
            max_share = optimization_rules.get('deep_affordability_max_share')
            min_required_sf = int(total_sf_int * min_share)
            if min_required_sf == 0:
                min_required_sf = int(total_sf_int * 0.2)
            model.Add(sum(low_band_assignments) >= min_required_sf)
            if max_share is not None:
                upper_sf = int(total_sf_int * max_share)
                model.Add(sum(low_band_assignments) <= upper_sf)
    model.Maximize(total_ami_sf_var)
    solver = cp_model.CpSolver()
    solver.parameters.num_workers = 1
    solver.parameters.random_seed = 0
    time_limit = optimization_rules.get('scenario_time_limit_seconds')
    if time_limit:
        solver.parameters.max_time_in_seconds = time_limit
    try:
        status = solver.Solve(model)
    except (SystemExit, KeyboardInterrupt):
        return {"status": "INTERRUPTED"}
    if status != cp_model.OPTIMAL and status != cp_model.FEASIBLE:
        return {"status": "NO_SOLUTION"}
    optimal_total_ami_sf = solver.Value(total_ami_sf_var)
    model.Add(total_ami_sf_var == optimal_total_ami_sf)
    premium_scores_int = (df_affordable['premium_score'] * 1000).astype(int)
    premium_alignment_expr = sum(
        sum(x[i][j] * bands_basis_points[j] for j in range(num_bands)) * premium_scores_int.iloc[i]
        for i in range(num_units)
    )
    model.Maximize(premium_alignment_expr)
    try:
        status = solver.Solve(model)
    except (SystemExit, KeyboardInterrupt):
        return {"status": "INTERRUPTED"}
    if status != cp_model.OPTIMAL and status != cp_model.FEASIBLE:
        return {"status": "NO_SOLUTION_IN_PASS_2"}
    def _extract_assignments():
        extracted = []
        for i in range(num_units):
            for j in range(num_bands):
                if solver.Value(x[i][j]):
                    unit_data = df_affordable.iloc[i].to_dict()
                    unit_data['assigned_ami'] = bands_to_test[j] / 100.0
                    extracted.append(unit_data)
                    break
        return extracted
    best_assignments = _extract_assignments()
    premium_optimal = solver.Value(premium_alignment_expr)
    model.Add(premium_alignment_expr == premium_optimal)

    lex_failed = False
    for unit_idx in range(num_units):
        assignment_index_expr = sum(j * x[unit_idx][j] for j in range(num_bands))
        model.Minimize(assignment_index_expr)
        status = solver.Solve(model)
        if status != cp_model.OPTIMAL and status != cp_model.FEASIBLE:
            lex_failed = True
            break
        model.Add(assignment_index_expr == solver.Value(assignment_index_expr))
    assignments = best_assignments if lex_failed else _extract_assignments()
    final_waami = _calculate_waami_from_assignments(assignments)
    metrics = _build_metrics(assignments)
    return {
        "status": "OPTIMAL",
        "waami": final_waami,
        "assignments": assignments,
        "bands": _get_bands_from_assignments(assignments),
        "metrics": metrics,
        "revenue_score": metrics['revenue_score'],
    }

def find_optimal_scenarios(
    df_affordable: pd.DataFrame,
    config: Dict[str, Any],
    relaxed_floor: float = None,
    diagnostics: Optional[List[Dict[str, Any]]] = None,
) -> Dict[str, Any]:
    optimization_rules = copy.deepcopy(config['optimization_rules'])
    if relaxed_floor:
        optimization_rules['waami_floor'] = relaxed_floor
    if 'deep_affordability_min_share' not in optimization_rules:
        optimization_rules['deep_affordability_min_share'] = 0.2
    dev_preferences = config['developer_preferences']
    df_with_scores = calculate_premium_scores(df_affordable, dev_preferences)
    total_affordable_sf = df_with_scores['net_sf'].sum()
    small_project_unit_threshold = optimization_rules.get('small_project_unit_threshold', 10)
    deep_aff_sf_threshold = optimization_rules.get('deep_affordability_sf_threshold', 10000)
    is_small_project = (total_affordable_sf < deep_aff_sf_threshold) or (len(df_with_scores) <= small_project_unit_threshold)
    potential_bands = optimization_rules.get('potential_bands', [])
    max_bands = optimization_rules.get('max_bands_per_scenario', 3)
    band_combos = list(itertools.combinations(potential_bands, 2)) + list(itertools.combinations(potential_bands, max_bands))
    band_combos = [sorted(combo) for combo in band_combos]
    waami_cap = optimization_rules.get('waami_cap_percent', 60)
    base_max_combo_checks = optimization_rules.get('max_band_combo_checks')
    effective_max_combo_checks = base_max_combo_checks
    if is_small_project:
        effective_max_combo_checks = optimization_rules.get('small_project_combo_allowance', base_max_combo_checks)
    band_combos = [combo for combo in band_combos if min(combo) <= waami_cap]

    priority_raw = list(optimization_rules.get('priority_band_combos', []))
    if is_small_project:
        small_priority_raw = optimization_rules.get('small_project_priority_band_combos', [])
        seen = {tuple(sorted(p)) for p in priority_raw}
        priority_raw.extend([combo for combo in small_priority_raw if tuple(sorted(combo)) not in seen])
        seen.update(tuple(sorted(combo)) for combo in small_priority_raw)
    priority_set = {tuple(sorted(p)) for p in priority_raw}
    priority_combos = []
    remaining_combos = []
    for combo in band_combos:
        if tuple(combo) in priority_set:
            priority_combos.append(combo)
        else:
            remaining_combos.append(combo)

    def _combo_sort_key(combo):
        mean = sum(combo) / len(combo)
        spread = max(combo) - min(combo)
        penalty = 0
        if 70 in combo and max(combo) >= 100:
            penalty += 50
        return (abs(mean - waami_cap), spread + penalty, len(combo), -max(combo))

    band_combos = priority_combos + sorted(remaining_combos, key=_combo_sort_key)
    notes = []
    max_unique = optimization_rules.get('max_unique_scenarios', 25)
    if is_small_project:
        max_unique = optimization_rules.get('small_project_max_unique_scenarios', max_unique)
    unique_results: Dict[tuple, Dict[str, Any]] = {}
    combos_checked = 0
    truncated_for_combo_limit = False
    interrupted = False
    for combo in band_combos:
        if effective_max_combo_checks and combos_checked >= effective_max_combo_checks:
            truncated_for_combo_limit = True
            break
        combo_start = time.perf_counter()
        combos_checked += 1
        result = _solve_single_scenario(df_with_scores, list(combo), total_affordable_sf, optimization_rules)
        combo_duration = time.perf_counter() - combo_start
        if diagnostics is not None:
            diagnostics.append({
                'combo': combo,
                'status': result.get('status'),
                'elapsed_sec': combo_duration,
                'combos_checked': combos_checked,
                'unique_scenarios_so_far': len(unique_results),
            })
        if result.get('status') == 'INTERRUPTED':
            interrupted = True
            break
        if result['status'] != 'OPTIMAL':
            continue
        result['premium_score'] = sum(u['premium_score'] * u['assigned_ami'] for u in result['assignments'])
        canonical = tuple(sorted((u['unit_id'], u['assigned_ami']) for u in result['assignments']))
        result['canonical_assignments'] = canonical
        result['source_combo'] = combo
        existing = unique_results.get(canonical)
        if existing:
            existing_score = (
                existing['waami'],
                existing['metrics']['revenue_score'],
                existing['premium_score'],
            )
            new_score = (
                result['waami'],
                result['metrics']['revenue_score'],
                result['premium_score'],
            )
            if new_score <= existing_score:
                continue
        unique_results[canonical] = result
        if max_unique and len(unique_results) >= max_unique:
            break
    if interrupted:
        notes.append("Solver interrupted before completing all band combinations (time limit or worker shutdown).")
    if truncated_for_combo_limit and effective_max_combo_checks and (not max_unique or len(unique_results) < max_unique):
        notes.append(
            f"Search stopped after evaluating {combos_checked} band mixes (configured limit: {effective_max_combo_checks}). Additional scenarios may be omitted."
        )
    if not unique_results:
        return {"scenarios": {}, "notes": ["The solver could not find any optimal solutions given the project constraints."]}
    sorted_results = sorted(
        unique_results.values(),
        key=lambda x: (x['waami'], x['metrics']['revenue_score'], x['premium_score']),
        reverse=True,
    )
    reporting_floor = optimization_rules.get('waami_floor')
    if reporting_floor is None:
        reporting_floor = 0.58
    reporting_floor = max(0.58, reporting_floor)
    epsilon = 1e-9
    filtered_results = [r for r in sorted_results if r['waami'] + epsilon >= reporting_floor]
    if filtered_results:
        if len(filtered_results) < len(sorted_results):
            notes.append(
                f"Excluded {len(sorted_results) - len(filtered_results)} scenario(s) with WAAMI below {reporting_floor*100:.2f}%."
            )
        sorted_results = filtered_results
    else:
        notes.append(
            f"No scenario reached the {reporting_floor*100:.2f}% WAAMI floor; returning the closest-feasible configurations for review."
        )
    scenarios: Dict[str, Dict[str, Any]] = {}
    selected_assignments = set()
    three_band_results = [r for r in sorted_results if len(r['bands']) >= 3]
    two_band_results = [r for r in sorted_results if len(r['bands']) == 2]
    multi_band_results = [r for r in sorted_results if len(r['bands']) >= 2]
    def _register(name: str, scenario: Dict[str, Any]):
        if scenario:
            scenarios[name] = scenario
            selected_assignments.add(scenario['canonical_assignments'])
        else:
            notes.append(f"No scenario available for '{name.replace('_', ' ')}'.")
    def _pick_from_list(candidates: List[Dict[str, Any]]):
        for candidate in candidates:
            if candidate['canonical_assignments'] in selected_assignments:
                continue
            return candidate
        return None
    absolute_best = _pick_from_list(three_band_results)
    if not absolute_best:
        absolute_best = _pick_from_list(multi_band_results)
        if absolute_best:
            notes.append("No 3-band configuration met the constraints; using the best available multi-band scenario.")
    if not absolute_best:
        absolute_best = _pick_from_list(sorted_results)
        if absolute_best:
            notes.append("Only single-band configurations satisfied the constraints; presenting the top-scoring outcome.")
    _register('absolute_best', absolute_best)
    best_3_band = _pick_from_list(three_band_results)
    if best_3_band:
        _register('best_3_band', best_3_band)
    else:
        notes.append("No viable 3-band scenario distinct from the absolute best could be found.")
    best_2_band = _pick_from_list(two_band_results)
    if best_2_band:
        _register('best_2_band', best_2_band)
    else:
        notes.append("No viable 2-band solution met the WAAMI floor.")
    alternative = _pick_from_list(three_band_results)
    if not alternative:
        alternative = _pick_from_list(multi_band_results)
        if alternative:
            notes.append("No additional 3-band scenario was available; using the best remaining multi-band option.")
    if alternative:
        _register('alternative', alternative)
    else:
        notes.append("No viable alternative scenario with a different unit assignment mix could be found.")
    revenue_sorted = sorted(
        sorted_results,
        key=lambda x: (x['metrics']['revenue_score'], x['waami'], x['premium_score']),
        reverse=True,
    )
    revenue_three_band = [r for r in revenue_sorted if len(r['bands']) >= 3 and r['canonical_assignments'] not in selected_assignments]
    revenue_multi_band = [r for r in revenue_sorted if len(r['bands']) >= 2 and r['canonical_assignments'] not in selected_assignments]
    client_oriented = _pick_from_list(revenue_three_band)
    if not client_oriented:
        client_oriented = _pick_from_list(revenue_multi_band)
        if client_oriented:
            notes.append("Client-oriented scenario falls back to the best remaining multi-band configuration.")
    if client_oriented:
        _register('client_oriented', client_oriented)
    else:
        notes.append("Client-oriented scenario unavailable; no remaining band mix satisfied the WAAMI floor.")
    return {"scenarios": scenarios, "notes": notes}


