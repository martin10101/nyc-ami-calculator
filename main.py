import sys
import json
import pandas as pd
import numpy as np

from ami_optix.parser import Parser
from ami_optix.config_loader import load_config
from ami_optix.solver import find_optimal_scenarios
from ami_optix.validator import run_compliance_checks


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
            print(json.dumps({"error": "No optimal solution found. The project may be unworkable with the current constraints.", "analysis_notes": notes}))
            return

        # --- Process Final Scenarios ---
        s1_abs_best = scenarios["absolute_best"][0]
        s2_alternative = next((s for s in scenarios["absolute_best"][1:] if set(s['bands']) != set(s1_abs_best['bands'])), scenarios["absolute_best"][1] if len(scenarios["absolute_best"]) > 1 else None)
        s3_client_oriented = scenarios.get("client_oriented")[0] if scenarios.get("client_oriented") else None
        s4_best_2_band = scenarios.get("best_2_band")

        # --- Compliance & Output ---
        compliance_report = run_compliance_checks(pd.DataFrame(s1_abs_best['assignments']), config['nyc_rules'])

        output = {
            "project_summary": {"total_affordable_sf": df_affordable['net_sf'].sum(), "total_affordable_units": len(df_affordable)},
            "analysis_notes": notes,
            "compliance_report": compliance_report,
            "scenario_absolute_best": {"waami": s1_abs_best['waami'], "bands": s1_abs_best['bands'], "assignments": s1_abs_best['assignments']},
        }
        if s2_alternative:
            output["scenario_alternative"] = {"waami": s2_alternative['waami'], "bands": s2_alternative['bands'], "assignments": s2_alternative['assignments']}
        if s3_client_oriented:
            output["scenario_client_oriented"] = {"waami": s3_client_oriented['waami'], "bands": s3_client_oriented['bands'], "assignments": s3_client_oriented['assignments']}
        if s4_best_2_band:
            output["scenario_best_2_band"] = {"waami": s4_best_2_band['waami'], "bands": s4_best_2_band['bands'], "assignments": s4_best_2_band['assignments']}

        return output

    except (FileNotFoundError, ValueError, IOError) as e:
        return {"error": str(e)}

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(f"Usage: python {sys.argv[0]} <path_to_file.csv_or_xlsx>")
        sys.exit(1)

    file_path_arg = sys.argv[1]
    result = main(file_path_arg)

    print(json.dumps(result, indent=2))

    if "error" in result:
        sys.exit(1)
