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
        if np.isnan(o):
            return None  # Convert NaN to None for JSON compatibility
        return float(o)
    if isinstance(o, np.ndarray):
        return o.tolist()
    if pd.isna(o):  # Handle pandas NaN values
        return None
    raise TypeError(f"Object of type {type(o)} is not JSON serializable")

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

        # Clean data for JSON serialization
        def clean_data(data):
            """Recursively clean data to ensure JSON serialization"""
            if isinstance(data, dict):
                return {k: clean_data(v) for k, v in data.items()}
            elif isinstance(data, list):
                return [clean_data(item) for item in data]
            elif isinstance(data, np.integer):
                return int(data)
            elif isinstance(data, np.floating):
                return None if np.isnan(data) else float(data)
            elif pd.isna(data):
                return None
            else:
                return data

        output = {
            "project_summary": {
                "total_affordable_sf": float(df_affordable['net_sf'].sum()), 
                "total_affordable_units": int(len(df_affordable))
            },
            "analysis_notes": notes,
            "compliance_report": clean_data(compliance_report),
            "scenario_absolute_best": clean_data({
                "waami": float(s1_abs_best['waami']), 
                "bands": s1_abs_best['bands'], 
                "assignments": s1_abs_best['assignments']
            }),
        }
        if s2_alternative:
            output["scenario_alternative"] = clean_data({
                "waami": float(s2_alternative['waami']), 
                "bands": s2_alternative['bands'], 
                "assignments": s2_alternative['assignments']
            })
        if s3_client_oriented:
            output["scenario_client_oriented"] = clean_data({
                "waami": float(s3_client_oriented['waami']), 
                "bands": s3_client_oriented['bands'], 
                "assignments": s3_client_oriented['assignments']
            })
        if s4_best_2_band:
            output["scenario_best_2_band"] = clean_data({
                "waami": float(s4_best_2_band['waami']), 
                "bands": s4_best_2_band['bands'], 
                "assignments": s4_best_2_band['assignments']
            })

        return output

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
