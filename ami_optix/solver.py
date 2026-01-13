import pandas as pd
from ortools.sat.python import cp_model
import itertools
import copy
import time
import math
from typing import Dict, List, Any, Optional

from ami_optix.overrides import ProjectOverrides


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


def _build_share_constraints(optimization_rules: Dict[str, Any]) -> Optional[Dict[str, float]]:
    min_share = optimization_rules.get('deep_affordability_min_share')
    max_share = optimization_rules.get('deep_affordability_max_share')
    threshold_band = optimization_rules.get('low_band_band_threshold', 40)
    if min_share is None and max_share is None:
        return None
    constraints = {'band_threshold': threshold_band}
    if min_share is not None:
        constraints['min_share'] = float(min_share)
    if max_share is not None:
        constraints['max_share'] = float(max_share)
    return constraints


def _get_bands_from_assignments(assignments: List[Dict[str, Any]]) -> List[int]:
    """Derives the unique AMI bands used in a set of assignments."""
    if not assignments:
        return []
    used_bands = {int(round(float(u['assigned_ami']) * 100)) for u in assignments}
    return sorted(list(used_bands))


def _build_metrics(assignments: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Compute reusable metrics for a scenario."""
    total_sf = sum(float(u['net_sf']) for u in assignments) or 0.0
    band_stats = {}
    low_band_units = 0
    low_band_sf = 0.0
    for unit in assignments:
        band = int(round(float(unit['assigned_ami']) * 100))
        stats = band_stats.setdefault(band, {'band': band, 'units': 0, 'net_sf': 0.0})
        stats['units'] += 1
        stats['net_sf'] += float(unit['net_sf'])
        if band <= 40:
            low_band_units += 1
            low_band_sf += float(unit['net_sf'])
    band_mix = []
    for band in sorted(band_stats):
        stats = band_stats[band]
        share = (stats['net_sf'] / total_sf) if total_sf else 0.0
        band_mix.append({**stats, 'share_of_sf': share})
    revenue_score = sum(float(u['net_sf']) * float(u['assigned_ami']) for u in assignments)
    low_band_share = (low_band_sf / total_sf) if total_sf else 0.0
    return {
        'total_units': len(assignments),
        'total_sf': total_sf,
        'revenue_score': revenue_score,
        'waami_percent': _calculate_waami_from_assignments(assignments) * 100,
        'band_mix': band_mix,
        'sf_at_40_band': sum(float(u['net_sf']) for u in assignments if int(round(float(u['assigned_ami']) * 100)) <= 40),
        'low_band_units': low_band_units,
        'low_band_sf': low_band_sf,
        'low_band_share': low_band_share,
        'total_monthly_rent': 0.0,
        'total_annual_rent': 0.0,
    }


def calculate_premium_scores(df: pd.DataFrame, dev_preferences: Dict[str, Any]) -> pd.DataFrame:
    df_norm = df.copy()
    weights = dev_preferences['premium_score_weights']
    if 'balcony' not in df_norm.columns:
        df_norm['balcony'] = 0
    else:
        df_norm['balcony'] = df_norm['balcony'].apply(
            lambda x: 1 if pd.notna(x) and str(x).lower() not in ['false', '0', 'no', ''] else 0
        )
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


def _assignments_to_canonical(assignments: List[Dict[str, Any]]) -> tuple:
    return tuple(sorted(
        (str(unit['unit_id']), int(round(float(unit['assigned_ami']) * 100)))
        for unit in assignments
    ))


def _solve_single_scenario(
    df_affordable: pd.DataFrame,
    bands_to_test: List[int],
    total_affordable_sf: float,
    optimization_rules: Dict[str, Any],
    share_constraints: Optional[Dict[str, float]] = None,
    unit_band_rules: Optional[Dict[int, List[int]]] = None,
    unit_min_band: Optional[Dict[int, int]] = None,
) -> Dict[str, Any]:
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

    allowed_band_rules = unit_band_rules or {}
    min_band_rules = unit_min_band or {}
    for i in range(num_units):
        allowed_bands = allowed_band_rules.get(i)
        min_band_value = min_band_rules.get(i)
        for j in range(num_bands):
            band_value = bands_to_test[j]
            if allowed_bands is not None and band_value not in allowed_bands:
                model.Add(x[i][j] == 0)
            if min_band_value is not None and band_value < min_band_value:
                model.Add(x[i][j] == 0)

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

    low_band_threshold = share_constraints.get('band_threshold', 40) if share_constraints else 40
    low_band_indices = [j for j, band in enumerate(bands_to_test) if band <= low_band_threshold]
    min_share = share_constraints.get('min_share') if share_constraints else None
    max_share = share_constraints.get('max_share') if share_constraints else None
    deep_affordability_threshold = optimization_rules.get('deep_affordability_sf_threshold', 10000)
    if min_share is None and max_share is None and total_affordable_sf >= deep_affordability_threshold:
        min_share = optimization_rules.get('deep_affordability_min_share', 0.2)
        max_share = optimization_rules.get('deep_affordability_max_share')

    if low_band_indices:
        low_band_sf_expr = sum(
            x[i][j] * sf_coeffs_int.iloc[i]
            for i in range(num_units)
            for j in low_band_indices
        )
        low_band_var = model.NewIntVar(0, total_sf_int, 'low_band_sf')
        model.Add(low_band_var == low_band_sf_expr)
        if min_share is not None:
            min_required_sf = math.ceil(min_share * total_sf_int)
            model.Add(low_band_var >= min_required_sf)
        if max_share is not None:
            upper_sf = math.floor(max_share * total_sf_int)
            model.Add(low_band_var <= upper_sf)
    elif min_share not in (None, 0.0):
        # No low-band options available for this combo; infeasible.
        return {"status": "NO_SOLUTION"}

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
        "canonical_assignments": _assignments_to_canonical(assignments),
    }


def find_optimal_scenarios(
    df_affordable: pd.DataFrame,
    config: Dict[str, Any],
    relaxed_floor: float = None,
    diagnostics: Optional[List[Dict[str, Any]]] = None,
    project_overrides: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    optimization_rules = copy.deepcopy(config['optimization_rules'])
    if relaxed_floor:
        optimization_rules['waami_floor'] = relaxed_floor

    overrides = ProjectOverrides.from_dict(project_overrides)
    solver_overrides = overrides.to_solver_payload(df_affordable)

    share_constraints = _build_share_constraints(optimization_rules)

    dev_preferences = copy.deepcopy(config['developer_preferences'])
    if solver_overrides.get('premium_weights'):
        weights = dev_preferences.get('premium_score_weights', {}).copy()
        for key, value in solver_overrides['premium_weights'].items():
            if key in weights:
                weights[key] = float(value)
        weight_sum = sum(weights.values())
        if weight_sum > 0:
            weights = {k: v / weight_sum for k, v in weights.items()}
        dev_preferences['premium_score_weights'] = weights

    df_with_scores = calculate_premium_scores(df_affordable, dev_preferences)
    total_affordable_sf = df_with_scores['net_sf'].sum()

    band_whitelist = solver_overrides.get('band_whitelist')
    potential_bands = optimization_rules.get('potential_bands', [])
    if band_whitelist:
        allowed = set(int(b) for b in band_whitelist)
        potential_bands = [band for band in potential_bands if band in allowed]
        for band in sorted(allowed):
            if band not in potential_bands:
                potential_bands.append(band)
        potential_bands = sorted(set(potential_bands))

    unit_band_rules = solver_overrides.get('unit_band_rules') or {}
    unit_min_band = solver_overrides.get('unit_min_band') or {}

    small_project_unit_threshold = optimization_rules.get('small_project_unit_threshold', 10)
    deep_aff_sf_threshold = optimization_rules.get('deep_affordability_sf_threshold', 10000)
    is_small_project = (
        (total_affordable_sf < deep_aff_sf_threshold) or
        (len(df_with_scores) <= small_project_unit_threshold)
    )
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
    if overrides.notes:
        notes.extend(overrides.notes)

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
        result = _solve_single_scenario(
            df_with_scores,
            list(combo),
            total_affordable_sf,
            optimization_rules,
            share_constraints=share_constraints,
            unit_band_rules=unit_band_rules,
            unit_min_band=unit_min_band,
        )
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
        canonical = result['canonical_assignments']
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

    # --- Dynamic WAAMI Threshold Filtering ---
    # If a 60%+ scenario exists, allow scenarios within 1% of best to show as alternatives
    # This shows alternative band mixes even if they have slightly lower WAAMI
    if sorted_results:
        best_waami = sorted_results[0]['waami']
        dynamic_thresholds = [0.60, 0.50, 0.40, 0.30]  # Standard AMI thresholds
        dynamic_floor = None
        for threshold in dynamic_thresholds:
            if best_waami >= threshold:
                # Allow 1% tolerance below the threshold for alternative scenarios
                dynamic_floor = threshold - 0.01
                break
        if dynamic_floor:
            pre_filter_count = len(sorted_results)
            sorted_results = [r for r in sorted_results if r['waami'] >= dynamic_floor - 0.001]
            if len(sorted_results) < pre_filter_count:
                notes.append(
                    f"Filtered {pre_filter_count - len(sorted_results)} scenario(s) below {(dynamic_floor + 0.01)*100:.0f}% WAAMI threshold (with 1% tolerance for alternatives)."
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
    revenue_three_band = [
        r for r in revenue_sorted
        if len(r['bands']) >= 3 and r['canonical_assignments'] not in selected_assignments
    ]
    revenue_multi_band = [
        r for r in revenue_sorted
        if len(r['bands']) >= 2 and r['canonical_assignments'] not in selected_assignments
    ]
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
