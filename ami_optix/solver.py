import pandas as pd
from ortools.sat.python import cp_model
import itertools
import copy

# A large integer scaling factor to handle floating-point arithmetic for currency and percentages.
# This avoids precision issues when dealing with constraints in the CP-SAT solver.
SCALE_FACTOR = 100_000_000

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
    Internal function to solve for the optimal assignment for one specific set of bands.
    This is the core of the "Specialist Chef".
    """
    waami_cap = optimization_rules['waami_cap_percent'] / 100.0
    bands = [b / 100.0 for b in bands_to_test]  # Convert bands from % to decimal

    # Scale SF to integers to avoid float issues in the solver, preserving 2 decimal places.
    sf_coeffs_int = (df_affordable['net_sf'] * 100).astype(int)
    total_sf_int = int(sf_coeffs_int.sum())

    model = cp_model.CpModel()
    num_units = len(df_affordable)
    num_bands = len(bands)

    # --- Decision Variables ---
    x = [[model.NewBoolVar(f'x_{i}_{j}') for j in range(num_bands)] for i in range(num_units)]

    # --- Constraints ---
    # 1. Each unit must be assigned exactly one AMI band.
    for i in range(num_units):
        model.AddExactlyOne(x[i])

    # 2. The WAAMI must be strictly less than the cap.
    # We use integer arithmetic with scaling factors to avoid floating point errors.
    cap_scaled = int(waami_cap * total_sf_int)
    total_ami_sf_scaled = model.NewIntVar(0, cap_scaled, 'total_ami_sf_scaled')

    bands_scaled = [int(b * SCALE_FACTOR) for b in bands]

    # Expression for the sum of (unit_sf * assigned_band_ami) across all units.
    # Both sf and ami are scaled to integers before being passed to the solver.
    total_ami_expr_scaled = sum(
        sum(x[i][j] * bands_scaled[j] for j in range(num_bands)) * sf_coeffs_int.iloc[i]
        for i in range(num_units)
    )

    # We need to divide the expression by SCALE_FACTOR to match the scale of total_ami_sf_scaled
    model.Add(total_ami_sf_scaled * SCALE_FACTOR == total_ami_expr_scaled)
    # The inclusive inequality allows solutions exactly at the cap.
    model.Add(total_ami_sf_scaled <= cap_scaled)

    # Add WAAMI floor constraint if a relaxed search is triggered
    waami_floor = optimization_rules.get('waami_floor')
    if waami_floor:
        floor_scaled = int(waami_floor * total_sf_int)
        model.Add(total_ami_sf_scaled >= floor_scaled)

    # 3. Deep Affordability Constraint (if applicable)
    deep_affordability_threshold = optimization_rules.get('deep_affordability_sf_threshold', 10000)
    if total_affordable_sf >= deep_affordability_threshold:
        # Find indices of bands that are at or below 40% AMI
        low_band_indices = [j for j, band in enumerate(bands_to_test) if band <= 40]

        if low_band_indices:
            # Sum of boolean variables for assignments to low bands
            low_band_assignments = [x[i][j] for i in range(num_units) for j in low_band_indices]

            # The number of units assigned to low bands must be >= 20% of total units.
            # We multiply by 100 to avoid floating-point comparisons.
            min_required_low_band_units = 20 * num_units
            model.Add(100 * sum(low_band_assignments) >= min_required_low_band_units)

    # --- Objective Function ---
    # Maximize the WAAMI by maximizing the scaled total AMI SF.
    model.Maximize(total_ami_sf_scaled)

    # --- Solve ---
    solver = cp_model.CpSolver()
    # Using more workers can speed up the search on multi-core machines.
    solver.parameters.num_workers = 8
    status = solver.Solve(model)

    # --- Process Results ---
    if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
        assignments = []
        for i in range(num_units):
            for j in range(num_bands):
                if solver.Value(x[i][j]) == 1:
                    unit_data = df_affordable.iloc[i].to_dict()
                    unit_data['assigned_ami'] = bands[j]
                    assignments.append(unit_data)
                    break

        # Calculate final WAAMI from the solver's objective value
        final_waami = solver.Value(total_ami_sf_scaled) / total_sf_int

        return {
            "status": "OPTIMAL",
            "waami": final_waami,
            "assignments": assignments,
            "bands": sorted(bands_to_test)
        }
    else:
        return {"status": "NO_SOLUTION"}

def find_optimal_scenarios(df_affordable, config, relaxed_floor=None):
    """
    Main orchestrator for the solver module. It runs the solver in two modes,
    filters for a 2-band option, and returns the results along with
    explanatory notes.
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

    scenarios = {}
    notes = []

    # --- Run 1: Absolute Best (WAAMI Maximization) ---
    absolute_best_results = []
    for combo in band_combos:
        result = _solve_single_scenario(df_with_scores, list(combo), total_affordable_sf, optimization_rules)
        if result['status'] == 'OPTIMAL':
            result['premium_score'] = sum(u['premium_score'] * u['assigned_ami'] for u in result['assignments'])
            # Add a canonical representation for robust comparison
            result['canonical_assignments'] = tuple(sorted((u['unit_id'], u['assigned_ami']) for u in result['assignments']))
            absolute_best_results.append(result)

    if absolute_best_results:
        absolute_best_results.sort(key=lambda x: (x['waami'], x['premium_score']), reverse=True)
        scenarios["absolute_best"] = absolute_best_results

    # --- Run 2: Client Oriented (Preference-Weighted) ---
    client_oriented_results = []
    for combo in band_combos:
        result = _solve_preference_weighted_scenario(df_with_scores, list(combo), total_affordable_sf, optimization_rules)
        if result['status'] == 'OPTIMAL':
            result['premium_score'] = sum(u['premium_score'] * u['assigned_ami'] for u in result['assignments'])
            result['canonical_assignments'] = tuple(sorted((u['unit_id'], u['assigned_ami']) for u in result['assignments']))
            client_oriented_results.append(result)

    if client_oriented_results:
        client_oriented_results.sort(key=lambda x: (x['waami'], x['premium_score']), reverse=True)
        # Only add the client-oriented scenario if it's meaningfully different from the absolute best
        if scenarios.get("absolute_best") and client_oriented_results[0]['canonical_assignments'] != scenarios["absolute_best"][0]['canonical_assignments']:
            scenarios["client_oriented"] = client_oriented_results
        else:
            notes.append("The 'Client Oriented' scenario was not shown because its optimal solution was identical to the 'Absolute Best' scenario.")

    # --- Find Best 2-Band Scenario from the 'Absolute Best' results ---
    if scenarios.get("absolute_best"):
        two_band_scenarios = [s for s in scenarios["absolute_best"] if len(s['bands']) == 2]
        if two_band_scenarios:
            scenarios["best_2_band"] = two_band_scenarios[0]
        else:
            notes.append("No viable 2-band solution was found that could meet the project's financial and compliance constraints.")

    return {"scenarios": scenarios, "notes": notes}

def _solve_preference_weighted_scenario(df_affordable, bands_to_test, total_affordable_sf, optimization_rules):
    """
    A separate solver function that finds an optimal assignment by balancing
    the WAAMI with a preference for assigning higher rents to more premium units.
    """
    waami_cap = optimization_rules['waami_cap_percent'] / 100.0
    bands = [b / 100.0 for b in bands_to_test]

    # Scale SF to integers to avoid float issues in the solver, preserving 2 decimal places.
    sf_coeffs_int = (df_affordable['net_sf'] * 100).astype(int)
    total_sf_int = int(sf_coeffs_int.sum())

    model = cp_model.CpModel()
    num_units = len(df_affordable)
    num_bands = len(bands)

    x = [[model.NewBoolVar(f'x_{i}_{j}') for j in range(num_bands)] for i in range(num_units)]

    for i in range(num_units):
        model.AddExactlyOne(x[i])

    cap_scaled = int(waami_cap * total_sf_int)
    total_ami_sf_scaled = model.NewIntVar(0, cap_scaled, 'total_ami_sf_scaled')
    bands_scaled = [int(b * SCALE_FACTOR) for b in bands]

    total_ami_expr_scaled = sum(
        sum(x[i][j] * bands_scaled[j] for j in range(num_bands)) * sf_coeffs_int.iloc[i]
        for i in range(num_units)
    )
    model.Add(total_ami_sf_scaled * SCALE_FACTOR == total_ami_expr_scaled)
    model.Add(total_ami_sf_scaled <= cap_scaled)

    # Add WAAMI floor constraint if a relaxed search is triggered
    waami_floor = optimization_rules.get('waami_floor')
    if waami_floor:
        floor_scaled = int(waami_floor * total_sf_int)
        model.Add(total_ami_sf_scaled >= floor_scaled)

    # --- Multi-Part Objective Function ---
    # This objective balances maximizing revenue with aligning higher rents to premium units.
    # We create a weighted objective where the premium score alignment is the primary driver.

    # 1. Premium Score Alignment Component
    # We scale the 0-1 premium score to make it a significant integer.
    premium_scores_int = (df_affordable['premium_score'] * 1000).astype(int)
    premium_alignment_expr = sum(
        sum(x[i][j] * bands_scaled[j] for j in range(num_bands)) * premium_scores_int.iloc[i]
        for i in range(num_units)
    )

    # 2. Combined Objective
    # The premium alignment term is weighted heavily to make it the primary objective.
    # The WAAMI-maximization term (total_ami_expr_scaled) acts as the tie-breaker.
    # The weight (2000) is a heuristic chosen to be larger than any likely SF value.
    WEIGHT = 2000
    final_objective = (premium_alignment_expr * WEIGHT) + total_ami_expr_scaled
    model.Maximize(final_objective)

    solver = cp_model.CpSolver()
    solver.parameters.num_workers = 8
    status = solver.Solve(model)

    if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
        assignments = []
        for i in range(num_units):
            for j in range(num_bands):
                if solver.Value(x[i][j]) == 1:
                    unit_data = df_affordable.iloc[i].to_dict()
                    unit_data['assigned_ami'] = bands[j]
                    assignments.append(unit_data)
                    break

        final_waami = solver.Value(total_ami_sf_scaled) / total_sf_int

        return {
            "status": "OPTIMAL",
            "waami": final_waami,
            "assignments": assignments,
            "bands": sorted(bands_to_test)
        }
    else:
        return {"status": "NO_SOLUTION"}
