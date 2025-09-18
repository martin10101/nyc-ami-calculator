import sys
import json
import pandas as pd
import numpy as np

from ami_optix.parser import Parser
from ami_optix.config_loader import load_config
from ami_optix.solver import find_optimal_scenarios
from ami_optix.validator import run_compliance_checks

def default_converter(o):
    """
    A default converter for json.dumps to handle numpy types that pandas uses.
    """
    if isinstance(o, np.integer):
        return int(o)
    if isinstance(o, np.floating):
        return float(o)
    if isinstance(o, np.ndarray):
        return o.tolist()
    # Let the default error handler raise the TypeError
    raise TypeError

def main(file_path):
    """
    Main function to orchestrate the entire optimization process.
    It coordinates the Parser, Solver, and Validator modules.
    """
    try:
        # 1. Load Configuration from the external YAML file
        config = load_config()

        # 2. Parse and Validate the input spreadsheet
        parser = Parser(file_path)
        df_affordable = parser.get_affordable_units()

        # 3. Find Optimal Scenarios using the Solver
        scenarios = find_optimal_scenarios(df_affordable, config)

        if not scenarios:
            print(json.dumps({"error": "No optimal solution found for any band combination. The project may be unworkable with the current constraints."}))
            return

        # 4. Select top scenarios and run compliance checks
        scenario1 = scenarios[0]

        # Find a good, different alternative scenario
        scenario2 = None
        for res in scenarios[1:]:
            if set(res['bands']) != set(scenario1['bands']):
                scenario2 = res
                break
        if not scenario2 and len(scenarios) > 1:
            scenario2 = scenarios[1]

        # Convert assignment list to DataFrame for validation
        df_scenario1_assignments = pd.DataFrame(scenario1['assignments'])
        compliance_report = run_compliance_checks(df_scenario1_assignments, config['nyc_rules'])

        # 5. Assemble the final output
        output = {
            "scenario1": {
                "waami": scenario1['waami'],
                "bands": scenario1['bands'],
                "assignments": scenario1['assignments'],
            },
            "complianceReport": compliance_report,
            "projectSummary": {
                "totalAffordableSF": df_affordable['net_sf'].sum(),
                "totalAffordableUnits": len(df_affordable),
            }
        }
        if scenario2:
            output["scenario2"] = {
                "waami": scenario2['waami'],
                "bands": scenario2['bands'],
                "assignments": scenario2['assignments'],
            }

        # Use the custom converter to handle numpy types during JSON serialization
        print(json.dumps(output, indent=2, default=default_converter))

    except (FileNotFoundError, ValueError, IOError) as e:
        print(json.dumps({"error": str(e)}))
        sys.exit(1)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(f"Usage: python {sys.argv[0]} <path_to_file.csv_or_xlsx>")
        sys.exit(1)

    file_path_arg = sys.argv[1]
    main(file_path_arg)
