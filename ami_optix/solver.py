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

    # Fix 4: Round SF to integers to avoid float issues in the solver.
    sf_coeffs_int = (df_affordable['net_sf'] * 100).round().astype(int)
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

    # Fix 2: Deep Affordability Constraint based on SF, not unit count.
    deep_affordability_threshold = optimization_rules.get('deep_affordability_sf_threshold', 10000)
    if total_affordable_sf >= deep_affordability_threshold:
        low_band_indices = [j for j, band in enumerate(bands_to_test) if band <= 40]
        if low_band_indices:
            low_sf = sum(
                x[i][j] * sf_coeffs_int.iloc[i]
                for i in range(num_units)
                for j in low_band_indices
            )
            # The SF of units assigned to low bands must be >= 20% of total affordable SF.
            model.Add(100 * low_sf >= 20 * total_sf_int)

    # --- Objective Function ---
    # Fix 1: Maximize the WAAMI to get it as tight to the cap as possible.
    model.Maximize(total_ami_sf_scaled)

    # --- Solve ---
    solver = cp_model.CpSolver()
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

        final_waami = solver.Value(total_ami_sf_scaled) / total_sf_int

        # Fix 5: Derive the bands from the actual assignments.
        used_bands = sorted({int(u['assigned_ami'] * 100) for u in assignments})

        return {
            "status": "OPTIMAL",
            "waami": final_waami,
            "assignments": assignments,
            "bands": used_bands
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
        # Sort by WAAMI (desc) and then premium score (desc) to find the best
        absolute_best_results.sort(key=lambda x: (x['waami'], x['premium_score']), reverse=True)
        scenarios["absolute_best"] = absolute_best_results

        # --- Run 2: Client Oriented (Lexicographical) ---
        # This pass finds the solution that maximizes the premium score, while maintaining the optimal WAAMI from the first pass.
        best_waami = scenarios["absolute_best"][0]['waami']
        best_bands = scenarios["absolute_best"][0]['bands']

        client_oriented_result = _solve_lexicographical_scenario(df_with_scores, best_bands, total_affordable_sf, optimization_rules, best_waami)
        if client_oriented_result['status'] == 'OPTIMAL':
            client_oriented_result['premium_score'] = sum(u['premium_score'] * u['assigned_ami'] for u in client_oriented_result['assignments'])
            client_oriented_result['canonical_assignments'] = tuple(sorted((u['unit_id'], u['assigned_ami']) for u in client_oriented_result['assignments']))

            # Only add the client-oriented scenario if it's meaningfully different from the absolute best
            if client_oriented_result['canonical_assignments'] != scenarios["absolute_best"][0]['canonical_assignments']:
                scenarios["client_oriented"] = [client_oriented_result] # Wrap in list to match structure
            else:
                notes.append("The 'Client Oriented' scenario was not shown because its optimal solution was identical to the 'Absolute Best' scenario.")

    # --- Find Best 2-Band and Alternative Scenarios from the 'Absolute Best' results ---
    if scenarios.get("absolute_best"):
        abs_best_scenario = scenarios["absolute_best"][0]

        # Find a suitable alternative scenario that is different from the absolute best
        for alt in scenarios["absolute_best"][1:]:
            if alt['canonical_assignments'] != abs_best_scenario['canonical_assignments']:
                scenarios["alternative"] = alt
                break

        # Find the best 2-band option
        two_band_scenarios = [s for s in scenarios["absolute_best"] if len(s['bands']) == 2]
        if two_band_scenarios:
            # Ensure the 2-band is also different from the absolute best, if possible
            best_2_band = two_band_scenarios[0]
            if best_2_band['canonical_assignments'] != abs_best_scenario['canonical_assignments']:
                 scenarios["best_2_band"] = best_2_band
            # If the best 2-band IS the absolute best, see if there's another 2-band option
            elif len(two_band_scenarios) > 1 and two_band_scenarios[1]['canonical_assignments'] != abs_best_scenario['canonical_assignments']:
                 scenarios["best_2_band"] = two_band_scenarios[1]
            else: # Otherwise, don't show a redundant 2-band scenario
                notes.append("The best 2-band solution was identical to the 'Absolute Best' scenario.")

        else:
            notes.append("No viable 2-band solution was found that could meet the project's financial and compliance constraints.")

    return {"scenarios": scenarios, "notes": notes}

def _solve_lexicographical_scenario(df_affordable, bands_to_test, total_affordable_sf, optimization_rules, optimal_waami):
    """
    Second-pass solver that locks the WAAMI to the optimal value and then
    maximizes for the premium score alignment.
    """
    bands = [b / 100.0 for b in bands_to_test]

    sf_coeffs_int = (df_affordable['net_sf'] * 100).round().astype(int)
    total_sf_int = int(sf_coeffs_int.sum())

    model = cp_model.CpModel()
    num_units = len(df_affordable)
    num_bands = len(bands)

    x = [[model.NewBoolVar(f'x_{i}_{j}') for j in range(num_bands)] for i in range(num_units)]

    for i in range(num_units):
        model.AddExactlyOne(x[i])

    # --- Constraint: Lock the WAAMI ---
    # The total scaled AMI*SF must equal the optimal value found in the first pass.
    # We allow a tiny tolerance for floating point conversion issues.
    optimal_waami_scaled = int(optimal_waami * total_sf_int)
    total_ami_sf_scaled = model.NewIntVar(optimal_waami_scaled, optimal_waami_scaled, 'total_ami_sf_scaled')

    bands_scaled = [int(b * SCALE_FACTOR) for b in bands]
    total_ami_expr_scaled = sum(
        sum(x[i][j] * bands_scaled[j] for j in range(num_bands)) * sf_coeffs_int.iloc[i]
        for i in range(num_units)
    )
    model.Add(total_ami_sf_scaled * SCALE_FACTOR == total_ami_expr_scaled)

    # --- Objective: Maximize Premium Alignment ---
    premium_scores_int = (df_affordable['premium_score'] * 1000).astype(int)
    premium_alignment_expr = sum(
        sum(x[i][j] * bands_scaled[j] for j in range(num_bands)) * premium_scores_int.iloc[i]
        for i in range(num_units)
    )
    model.Maximize(premium_alignment_expr)

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

        used_bands = sorted({int(u['assigned_ami'] * 100) for u in assignments})
        return {
            "status": "OPTIMAL",
            "waami": optimal_waami, # WAAMI is fixed
            "assignments": assignments,
            "bands": used_bands
        }
    else:
        return {"status": "NO_SOLUTION"}
