import pandas as pd

def run_compliance_checks(df_assignments, nyc_rules):
    """
    Runs all validation checks from the config file against a scenario's assignments.

    Args:
        df_assignments (pd.DataFrame): A DataFrame of the units with their final AMI assignments.
        nyc_rules (dict): The 'nyc_rules' portion of the configuration.

    Returns:
        list: A list of dictionaries, where each dictionary is a compliance check result.
    """
    checks = nyc_rules['validation_checks']
    results = []

    # Helper to map bedroom counts (integers) to the string keys used in rules_config.yml
    bedroom_map = {
        0: 'studio',
        1: 'one_bedroom',
        2: 'two_bedroom',
        3: 'three_bedroom',
        4: 'four_bedroom'
    }

    # 1. Unit Size Minimum Check
    size_minima_passed = True
    # Ensure a 'size_minima' key exists before proceeding
    if 'size_minima' in checks:
        for _, unit in df_assignments.iterrows():
            bedroom_key = bedroom_map.get(int(unit['bedrooms']))
            if bedroom_key:
                min_sf = checks['size_minima'].get(bedroom_key, 0)
                if min_sf > 0 and unit['net_sf'] < min_sf:
                    size_minima_passed = False
                    results.append({
                        "check": "Unit Size Minimum",
                        "status": "FLAGGED",
                        "details": f"Unit {unit['unit_id']} ({int(unit['bedrooms'])} BR) is {unit['net_sf']} SF, below the required {min_sf} SF."
                    })
    if size_minima_passed:
        results.append({"check": "Unit Size Minimum", "status": "PASS", "details": "All units meet minimum size requirements."})


    # 2. Building Mix Checks
    total_units = len(df_assignments)
    if total_units > 0 and 'mix_checks' in checks:
        studio_units = len(df_assignments[df_assignments['bedrooms'] == 0])
        two_br_plus_units = len(df_assignments[df_assignments['bedrooms'] >= 2])

        # Max Studio Percentage Check
        max_studio_percent = checks['mix_checks']['max_studio_percent']
        studio_percent = (studio_units / total_units) * 100
        if studio_percent > max_studio_percent:
            results.append({
                "check": "Max Studio Percentage",
                "status": "FLAGGED",
                "details": f"Project has {studio_percent:.1f}% studios, exceeding the {max_studio_percent}% maximum."
            })
        else:
            results.append({"check": "Max Studio Percentage", "status": "PASS", "details": f"{studio_percent:.1f}% studios is within the {max_studio_percent}% limit."})

        # Min 2+ Bedroom Percentage Check
        min_two_br_plus_percent = checks['mix_checks']['min_two_br_plus_percent']
        two_br_plus_percent = (two_br_plus_units / total_units) * 100
        if two_br_plus_percent < min_two_br_plus_percent:
            results.append({
                "check": "Min 2+ Bedroom Percentage",
                "status": "FLAGGED",
                "details": f"Project has {two_br_plus_percent:.1f}% 2+ BR units, below the required {min_two_br_plus_percent}%."
            })
        else:
            results.append({"check": "Min 2+ Bedroom Percentage", "status": "PASS", "details": f"{two_br_plus_percent:.1f}% 2+ BR units meets the {min_two_br_plus_percent}% requirement."})
    else:
        # If there are no units or no mix_checks config, the checks are not applicable or pass by default.
        results.append({"check": "Max Studio Percentage", "status": "PASS", "details": "N/A."})
        results.append({"check": "Min 2+ Bedroom Percentage", "status": "PASS", "details": "N/A."})


    return results
