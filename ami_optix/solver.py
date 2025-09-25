import pandas as pd
from ortools.sat.python import cp_model
import itertools
import copy

def _calculate_waami_from_assignments(assignments):
    """
    Calculates the WAAMI from a list of assignment dictionaries using integer arithmetic.
    This serves as a consistent final calculator and verifier for the system.

    Args:
        assignments (list): A list of unit assignment dictionaries. Each dict must
                            have 'net_sf' and 'assigned_ami' (as a decimal, e.g., 0.6).

    Returns:
        float: The calculated WAAMI.
    """
    if not assignments:
        return 0.0

    # Use integer math to prevent floating point inaccuracies.
    # Scale SF by 100 (to handle up to 2 decimal places).
    # Scale AMI by 10000 (to handle basis points, e.g., 0.6 -> 6000).
    total_sf_int = sum(int(unit['net_sf'] * 100) for unit in assignments)

    if total_sf_int == 0:
        return 0.0

    total_ami_sf_scaled = sum(
        int(unit['net_sf'] * 100) * int(unit['assigned_ami'] * 10000)
        for unit in assignments
    )

    # WAAMI = (total_ami_sf_scaled / total_sf_int) / 10000
    # The final result is a float, but it's derived from precise integer calculations.
    return (total_ami_sf_scaled / total_sf_int) / 10000

def _get_bands_from_assignments(assignments):
    """
    Derives the unique AMI bands used in a set of assignments, ensuring the
    UI is always truthful to the solver's actual output.
    """
    if not assignments:
        return []
    # Convert from decimal (0.4) to percent (40) and round to avoid float issues
    used_bands = {int(round(u['assigned_ami'] * 100)) for u in assignments}
    return sorted(list(used_bands))

def calculate_premium_scores(df, dev_preferences):
    """
    Calculates a normalized premium score for each unit based on developer preferences.
    These scores are used for tie-breaking between scenarios with the same WAAMI.

    Args:
        df (pd.DataFrame): The DataFrame of affordable units.
        dev_preferences (dict): The developer_preferences portion of the config.

    Returns:
        pd.DataFrame: The DataFrame with an added 'premium_score' column.
    """
    df_norm = df.copy()
    weights = dev_preferences['premium_score_weights']

    # Ensure balcony column exists and is numeric, default to 0 if not
    if 'balcony' not in df_norm.columns:
        df_norm['balcony'] = 0
    else:
        df_norm['balcony'] = df_norm['balcony'].apply(lambda x: 1 if pd.notna(x) and str(x).lower() not in ['false', '0', 'no', ''] else 0)


    # Normalize each component from 0 to 1 to give them equal footing before applying weights
    for col in ['floor', 'net_sf', 'bedrooms', 'balcony']:
        if col in df_norm.columns:
            if df_norm[col].max() > df_norm[col].min():
                df_norm[f'{col}_norm'] = (df_norm[col] - df_norm[col].min()) / (df_norm[col].max() - df_norm[col].min())
            else:
                df_norm[f'{col}_norm'] = 0
        else: # If a premium column like 'floor' or 'balcony' wasn't in the input
             df_norm[f'{col}_norm'] = 0


    # Calculate the final weighted score
    df['premium_score'] = (
        df_norm['floor_norm'] * weights['floor'] +
        df_norm['net_sf_norm'] * weights['net_sf'] +
        df_norm['bedrooms_norm'] * weights['bedrooms'] +
        df_norm['balcony_norm'] * weights['balcony']
    )
    return df

def _solve_single_scenario(df_affordable, bands_to_test, total_affordable_sf, optimization_rules):
    """
    Internal function to solve for the optimal assignment using a two-pass
    lexicographical approach. It first maximizes WAAMI, then maximizes the
    premium score as a secondary objective.
    """
    # --- Integer-based Setup ---
    waami_cap_basis_points = int(optimization_rules['waami_cap_percent'] * 100)
    bands_basis_points = [int(b * 100) for b in bands_to_test]
    sf_coeffs_int = (df_affordable['net_sf'] * 100).astype(int)
    total_sf_int = int(sf_coeffs_int.sum())

    model = cp_model.CpModel()
    num_units = len(df_affordable)
    num_bands = len(bands_to_test)

    # --- Decision Variables ---
    x = [[model.NewBoolVar(f'x_{i}_{j}') for j in range(num_bands)] for i in range(num_units)]

    # --- Constraints ---
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
            low_band_assignments = [x[i][j] for i in range(num_units) for j in low_band_indices]
            min_required_low_band_units = 20 * num_units
            model.Add(100 * sum(low_band_assignments) >= min_required_low_band_units)

    # --- Pass 1: Maximize WAAMI ---
    model.Maximize(total_ami_sf_var)
    solver = cp_model.CpSolver()
    # Use a single worker to keep CP-SAT output deterministic across runs.
    solver.parameters.num_workers = 1
    solver.parameters.random_seed = 0
    status = solver.Solve(model)

    if status != cp_model.OPTIMAL and status != cp_model.FEASIBLE:
        return {"status": "NO_SOLUTION"}

    # Lock in the optimal WAAMI from the first pass
    optimal_total_ami_sf = solver.Value(total_ami_sf_var)
    model.Add(total_ami_sf_var == optimal_total_ami_sf)

    # --- Pass 2: Maximize Premium Score Alignment ---
    premium_scores_int = (df_affordable['premium_score'] * 1000).astype(int)
    premium_alignment_expr = sum(
        sum(x[i][j] * bands_basis_points[j] for j in range(num_bands)) * premium_scores_int.iloc[i]
        for i in range(num_units)
    )
    model.Maximize(premium_alignment_expr)

    status = solver.Solve(model)

    if status != cp_model.OPTIMAL and status != cp_model.FEASIBLE:
        # This fallback should rarely be hit, but indicates an issue if the
        # second pass fails after the first one succeeded.
        return {"status": "NO_SOLUTION_IN_PASS_2"}

    def _extract_assignments():
        extracted = []
        for i in range(num_units):
            for j in range(num_bands):
                if solver.Value(x[i][j]) == 1:
                    unit_data = df_affordable.iloc[i].to_dict()
                    unit_data['assigned_ami'] = bands_to_test[j] / 100.0
                    extracted.append(unit_data)
                    break
        return extracted

    optimal_premium_alignment = solver.Value(premium_alignment_expr)
    best_assignments = _extract_assignments()

    # --- Pass 3: Enumerate optimal assignments for deterministic ordering ---
    model.Add(premium_alignment_expr == optimal_premium_alignment)

    class _OptimalSolutionCollector(cp_model.CpSolverSolutionCallback):
        def __init__(self, x_vars, bands, frame):
            super().__init__()
            self._x_vars = x_vars
            self._bands = bands
            self._frame = frame
            self.best_assignment = None
            self.best_canonical = None

        def on_solution_callback(self):
            extracted = []
            for unit_idx in range(num_units):
                for band_idx in range(num_bands):
                    if self.Value(self._x_vars[unit_idx][band_idx]):
                        unit_data = self._frame.iloc[unit_idx].to_dict()
                        unit_data['assigned_ami'] = self._bands[band_idx] / 100.0
                        extracted.append(unit_data)
                        break
            canonical = tuple(sorted((unit['unit_id'], unit['assigned_ami']) for unit in extracted))
            if self.best_canonical is None or canonical < self.best_canonical:
                self.best_canonical = canonical
                self.best_assignment = extracted

    collector = _OptimalSolutionCollector(x, bands_to_test, df_affordable)
    solver.parameters.enumerate_all_solutions = True
    model.ClearObjective()
    solver.Solve(model, collector)
    solver.parameters.enumerate_all_solutions = False

    if collector.best_assignment:
        assignments = collector.best_assignment
    else:
        assignments = best_assignments

    final_waami = _calculate_waami_from_assignments(assignments)

    return {
        "status": "OPTIMAL",
        "waami": final_waami,
        "assignments": assignments,
        "bands": _get_bands_from_assignments(assignments)
    }

def find_optimal_scenarios(df_affordable, config, relaxed_floor=None):
    """
    Main orchestrator for the solver. It finds the best scenarios by running
    the lexicographical solver across all band combinations, then selects a
    distinct 'absolute_best' and 'alternative' scenario from the results.
    """
    optimization_rules = copy.deepcopy(config['optimization_rules'])
    if relaxed_floor:
        optimization_rules['waami_floor'] = relaxed_floor

    dev_preferences = config['developer_preferences']

    df_with_scores = calculate_premium_scores(df_affordable, dev_preferences)
    total_affordable_sf = df_with_scores['net_sf'].sum()

    potential_bands = optimization_rules.get('potential_bands', [])
    max_bands = optimization_rules.get('max_bands_per_scenario', 3)

    band_combos = list(itertools.combinations(potential_bands, 2)) + \
                  list(itertools.combinations(potential_bands, max_bands))

    # --- Run Solver for All Combinations ---
    all_results = []
    for combo in band_combos:
        result = _solve_single_scenario(df_with_scores, list(combo), total_affordable_sf, optimization_rules)
        if result['status'] == 'OPTIMAL':
            # Calculate premium score for sorting and add canonical representation for uniqueness check
            result['premium_score'] = sum(u['premium_score'] * u['assigned_ami'] for u in result['assignments'])
            result['canonical_assignments'] = tuple(sorted((u['unit_id'], u['assigned_ami']) for u in result['assignments']))
            all_results.append(result)

    if not all_results:
        return {"scenarios": {}, "notes": ["The solver could not find any optimal solutions given the project constraints."]}

    # --- Filter for Unique Scenarios and Sort ---
    # We only want to present scenarios that are meaningfully different.
    unique_results = {result['canonical_assignments']: result for result in all_results}.values()
    sorted_results = sorted(list(unique_results), key=lambda x: (x['waami'], x['premium_score']), reverse=True)

    scenarios = {}
    notes = []

    # --- Select Scenarios for Presentation ---
    # 1. Absolute Best: The highest WAAMI, with premium score as a tie-breaker.
    scenarios['absolute_best'] = [sorted_results[0]]

    # 2. Alternative: The next-best scenario that has a different assignment set.
    best_assignments = scenarios['absolute_best'][0]['canonical_assignments']
    for result in sorted_results[1:]:
        if result['canonical_assignments'] != best_assignments:
            scenarios['alternative'] = [result]
            break

    if 'alternative' not in scenarios:
        notes.append("No viable alternative scenario with a different unit assignment mix could be found.")

    # 3. Best 2-Band: The best-performing scenario that only uses two bands.
    two_band_scenarios = [s for s in sorted_results if len(s['bands']) == 2]
    if two_band_scenarios:
        # The list is already sorted, so the first one is the best.
        scenarios['best_2_band'] = two_band_scenarios[0]
    else:
        notes.append("No viable 2-band solution was found that could meet the project's financial and compliance constraints.")

    return {"scenarios": scenarios, "notes": notes}






