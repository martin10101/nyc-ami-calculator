import sys
import json
import uuid
import time
from typing import List, Dict, Any
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
    """Main function to orchestrate the entire optimization process."""
    try:
        analysis_id = uuid.uuid4().hex[:8]
        analysis_started = time.perf_counter()

        config = load_config()
        parser = Parser(file_path)
        df_affordable = parser.get_affordable_units()

        solver_diagnostics: List[Dict[str, Any]] = []
        base_index = len(solver_diagnostics)
        solver_results = find_optimal_scenarios(df_affordable, config, diagnostics=solver_diagnostics)
        for entry in solver_diagnostics[base_index:]:
            entry['phase'] = 'standard'

        scenarios = solver_results.get("scenarios", {})
        notes = solver_results.get("notes", [])

        # --- "Smart Search" Fallback Logic ---
        num_scenarios_found = len([s for s in scenarios.values() if s])
        unit_count = len(df_affordable)
        threshold = config['optimization_rules'].get('small_project_unit_threshold', 10)

        if num_scenarios_found < 2 and unit_count <= threshold:
            relaxed_floor_pct = config['optimization_rules'].get('relaxed_search_waami_floor', 59.7)
            notes.append(
                f"Standard search yielded only {num_scenarios_found} result(s). Performing a 'Relaxed Search' with a WAAMI target floor of {relaxed_floor_pct}%."
            )

            base_index = len(solver_diagnostics)
            relaxed_solver_results = find_optimal_scenarios(
                df_affordable,
                config,
                relaxed_floor=(relaxed_floor_pct / 100.0),
                diagnostics=solver_diagnostics,
            )
            for entry in solver_diagnostics[base_index:]:
                entry['phase'] = 'relaxed'

            relaxed_scenarios = relaxed_solver_results.get("scenarios", {})
            for name, scenario in relaxed_scenarios.items():
                scenarios.setdefault(name, scenario)
            notes.extend(relaxed_solver_results.get("notes", []))

        if not scenarios.get("absolute_best"):
            return {
                "error": "No optimal solution found. The project may be unworkable with the current constraints.",
                "analysis_notes": notes,
                "analysis_id": analysis_id,
                "solver_diagnostics": solver_diagnostics,
            }

        # --- Compliance & Output ---
        s1_assignments = scenarios["absolute_best"]["assignments"]
        compliance_report = run_compliance_checks(pd.DataFrame(s1_assignments), config['nyc_rules'])

        unique_scenario_count = len([s for s in scenarios.values() if s])
        analysis_duration = time.perf_counter() - analysis_started
        analysis_meta = {
            "analysis_id": analysis_id,
            "duration_sec": analysis_duration,
            "solver_combination_count": len(solver_diagnostics),
            "solver_unique_scenarios": unique_scenario_count,
            "timestamp": time.time(),
            "truncated": any("Search stopped after" in note for note in notes),
        }

        output = {
            "project_summary": {
                "total_affordable_sf": df_affordable['net_sf'].sum(),
                "total_affordable_units": len(df_affordable),
            },
            "analysis_notes": notes,
            "compliance_report": compliance_report,
            "scenarios": scenarios,
            "solver_diagnostics": solver_diagnostics,
            "analysis_meta": analysis_meta,
        }

        # Flatten for existing consumers
        if scenarios.get("absolute_best"):
            output["scenario_absolute_best"] = scenarios["absolute_best"]
        if scenarios.get("client_oriented"):
            output["scenario_client_oriented"] = scenarios["client_oriented"]
        if scenarios.get("best_3_band"):
            output["scenario_best_3_band"] = scenarios["best_3_band"]
        if scenarios.get("best_2_band"):
            output["scenario_best_2_band"] = scenarios["best_2_band"]
        if scenarios.get("alternative"):
            output["scenario_alternative"] = scenarios["alternative"]

        return {
            "results": output,
            "original_headers": parser.mapped_headers,
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
