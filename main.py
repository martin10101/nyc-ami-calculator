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

        # --- Standard Search ---
        solver_results = find_optimal_scenarios(df_affordable, config)
        scenarios = solver_results.get("scenarios", {})
        notes = solver_results.get("notes", [])

        # --- "Smart Search" Fallback Logic ---
        num_scenarios_found = sum(1 for v in scenarios.values() if v)
        unit_count = len(df_affordable)
        threshold = config['optimization_rules'].get('small_project_unit_threshold', 10)

        if num_scenarios_found < 2 and unit_count <= threshold:
            relaxed_floor_pct = config['optimization_rules'].get('relaxed_search_waami_floor', 59.7)
            notes.append(f"Standard search yielded only {num_scenarios_found} result(s). Performing a 'Relaxed Search' with a WAAMI target floor of {relaxed_floor_pct}%.")

            relaxed_solver_results = find_optimal_scenarios(df_affordable, config, relaxed_floor=(relaxed_floor_pct / 100.0))

            # Merge the new results, being careful not to overwrite the original absolute_best
            scenarios.update(relaxed_solver_results.get("scenarios", {}))
            notes.extend(relaxed_solver_results.get("notes", []))


        if not scenarios.get("absolute_best"):
            return {"error": "No optimal solution found. The project may be unworkable with the current constraints.", "analysis_notes": notes}

        # --- Compliance & Output ---
        s1_assignments = scenarios["absolute_best"][0]["assignments"]
        compliance_report = run_compliance_checks(pd.DataFrame(s1_assignments), config['nyc_rules'])

        output = {
            "project_summary": {
                "total_affordable_sf": df_affordable['net_sf'].sum(),
                "total_affordable_units": len(df_affordable)
            },
            "analysis_notes": notes,
            "compliance_report": compliance_report,
        }

        # Flatten the scenarios dictionary for consumers like the narrator and report generator
        if scenarios.get("absolute_best"):
            output["scenario_absolute_best"] = scenarios["absolute_best"][0]
        if scenarios.get("alternative"):
            output["scenario_alternative"] = scenarios["alternative"][0]
        if scenarios.get("best_2_band"):
            output["scenario_best_2_band"] = scenarios["best_2_band"]

        return {
            "results": output,
            "original_headers": parser.mapped_headers
        }

    except (FileNotFoundError, ValueError, IOError) as e:
        return {"error": str(e)}

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(f"Usage: python {sys.argv[0]} <path_to_file.csv_or_xlsx>")
        sys.exit(1)

    file_path_arg = sys.argv[1]
    result = main(file_path_arg)

    print(json.dumps(result, indent=2, default=default_converter))

    if "error" in result:
        sys.exit(1)
