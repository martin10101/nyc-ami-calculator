import sys
import json
import pandas as pd
import numpy as np

from ami_optix.parser import Parser
from ami_optix.config_loader import load_config
from ami_optix.solver import find_optimal_scenarios
from ami_optix.validator import run_compliance_checks

def default_converter(o):
    """A default converter for json.dumps to handle numpy types."""
    if isinstance(o, np.integer):
        return int(o)
    if isinstance(o, np.floating):
        return float(o)
    if isinstance(o, np.ndarray):
        return o.tolist()
    raise TypeError

def main(file_path):
    """
    Main function to orchestrate the entire optimization process.
    """
    try:
        config = load_config()
        parser = Parser(file_path)
        df_affordable = parser.get_affordable_units()

        scenarios_dict = find_optimal_scenarios(df_affordable, config)

        if not scenarios_dict.get("absolute_best"):
            print(json.dumps({"error": "No optimal solution found. The project may be unworkable with the current constraints."}))
            return

        # --- Process Scenarios ---
        # Scenario 1: Absolute Best
        s1_abs_best = scenarios_dict["absolute_best"][0]

        # Scenario 2: A good alternative to the best
        s2_alternative = None
        for res in scenarios_dict["absolute_best"][1:]:
            if set(res['bands']) != set(s1_abs_best['bands']):
                s2_alternative = res
                break
        if not s2_alternative and len(scenarios_dict["absolute_best"]) > 1:
            s2_alternative = scenarios_dict["absolute_best"][1]

        # Scenario 3: The 'Client Oriented' scenario
        s3_client_oriented = None
        if scenarios_dict.get("client_oriented"):
            s3_client_oriented = scenarios_dict["client_oriented"][0]

        # --- Compliance & Output ---
        compliance_report = run_compliance_checks(pd.DataFrame(s1_abs_best['assignments']), config['nyc_rules'])

        output = {
            "scenario_absolute_best": {
                "waami": s1_abs_best['waami'],
                "bands": s1_abs_best['bands'],
                "assignments": s1_abs_best['assignments'],
            },
            "compliance_report": compliance_report,
            "project_summary": {
                "total_affordable_sf": df_affordable['net_sf'].sum(),
                "total_affordable_units": len(df_affordable),
            }
        }
        if s2_alternative:
            output["scenario_alternative"] = {
                "waami": s2_alternative['waami'],
                "bands": s2_alternative['bands'],
                "assignments": s2_alternative['assignments'],
            }
        if s3_client_oriented:
            output["scenario_client_oriented"] = {
                "waami": s3_client_oriented['waami'],
                "bands": s3_client_oriented['bands'],
                "assignments": s3_client_oriented['assignments'],
            }

        # Scenario 4: The 'Best 2-Band' scenario
        s4_best_2_band = scenarios_dict.get("best_2_band")
        if s4_best_2_band:
            output["scenario_best_2_band"] = {
                "waami": s4_best_2_band['waami'],
                "bands": s4_best_2_band['bands'],
                "assignments": s4_best_2_band['assignments'],
            }

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
