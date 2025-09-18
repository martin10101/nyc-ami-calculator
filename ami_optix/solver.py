import pandas as pd
from ortools.sat.python import cp_model
import itertools

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

    model = cp_model.CpModel()
    num_units = len(df_affordable)
    num_bands = len(bands)

    # --- Decision Variables ---
    # x[i][j] is a boolean variable, true if unit i is assigned band j
    x = [[model.NewBoolVar(f'x_{i}_{j}') for j in range(num_bands)] for i in range(num_units)]

    # --- Constraints ---
    # 1. Each unit must be assigned exactly one AMI band.
    for i in range(num_units):
        model.AddExactlyOne(x[i])

    # 2. The WAAMI must be strictly less than the cap.
    # We use integer arithmetic with a scaling factor to avoid floating point errors.
    total_ami_sf_scaled = model.NewIntVar(0, int(waami_cap * total_affordable_sf * SCALE_FACTOR), 'total_ami_sf_scaled')

    # Pre-calculate scaled values
    bands_scaled = [int(b * SCALE_FACTOR) for b in bands]

    # Expression for the sum of (unit_sf * assigned_band_ami) across all units, scaled.
    total_ami_expr_scaled = sum(
        sum(x[i][j] * bands_scaled[j] for j in range(num_bands)) * df_affordable['net_sf'].iloc[i]
        for i in range(num_units)
    )

    model.Add(total_ami_sf_scaled == total_ami_expr_scaled)
    # The strict inequality is handled by making the upper bound of the variable one less than the cap.
    model.Add(total_ami_sf_scaled < int(waami_cap * total_affordable_sf * SCALE_FACTOR))

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
        # The objective value is already scaled by (total_sf * SCALE_FACTOR), so we just reverse that.
        final_waami = solver.Value(total_ami_sf_scaled) / (total_affordable_sf * SCALE_FACTOR)

        return {
            "status": "OPTIMAL",
            "waami": final_waami,
            "assignments": assignments,
            "bands": sorted(bands_to_test)
        }
    else:
        return {"status": "NO_SOLUTION"}

def find_optimal_scenarios(df_affordable, config):
    """
    Main orchestrator for the solver module. It calculates premium scores,
    generates band combinations, runs the solver for each, and returns the
    best results.

    Args:
        df_affordable (pd.DataFrame): The validated DataFrame of affordable units.
        config (dict): The full application configuration dictionary.

    Returns:
        list: A list of the best-found scenarios, sorted by WAAMI and premium score.
              Returns an empty list if no solutions are found.
    """
    optimization_rules = config['optimization_rules']
    dev_preferences = config['developer_preferences']

    # First, calculate premium scores for tie-breaking later.
    df_with_scores = calculate_premium_scores(df_affordable, dev_preferences)
    total_affordable_sf = df_with_scores['net_sf'].sum()

    potential_bands = optimization_rules.get('potential_bands', [30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130])
    max_bands = optimization_rules.get('max_bands_per_scenario', 3)

    # Generate all 2-band and 3-band combinations to test.
    band_combos = list(itertools.combinations(potential_bands, 2)) + \
                  list(itertools.combinations(potential_bands, max_bands))

    all_results = []
    for combo in band_combos:
        result = _solve_single_scenario(df_with_scores, list(combo), total_affordable_sf, optimization_rules)
        if result['status'] == 'OPTIMAL':
            # Calculate the aggregate premium score for the entire scenario for tie-breaking.
            premium_score = sum(u['premium_score'] * u['assigned_ami'] for u in result['assignments'])
            result['premium_score'] = premium_score
            all_results.append(result)

    if not all_results:
        return []

    # Sort results: Primary key is WAAMI (desc), secondary is Premium Score (desc).
    all_results.sort(key=lambda x: (x['waami'], x['premium_score']), reverse=True)

    return all_results
