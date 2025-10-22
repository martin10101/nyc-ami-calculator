import os
import sys
import json
import uuid
import time
import copy
from typing import List, Dict, Any
import pandas as pd
import numpy as np

from ami_optix.parser import Parser
from ami_optix.config_loader import load_config
from ami_optix.solver import find_optimal_scenarios
from ami_optix.validator import run_compliance_checks
from ami_optix.rent_calculator import load_rent_schedule, compute_rents_for_assignments


def default_converter(o):
    """A default converter for json.dumps to handle numpy types."""
    if isinstance(o, np.integer):
        return int(o)
    if isinstance(o, np.floating):
        return float(o)
    if isinstance(o, np.ndarray):
        return o.tolist()
    raise TypeError


DEFAULT_UTILITIES = {"cooking": "na", "heat": "na", "hot_water": "na"}


def _sanitize_utilities(payload: Dict[str, Any] | None) -> Dict[str, str]:
    sanitized = DEFAULT_UTILITIES.copy()
    if isinstance(payload, dict):
        for key in sanitized:
            value = payload.get(key)
            if isinstance(value, str) and value:
                sanitized[key] = value.strip()
    return sanitized


def _maybe_load_rent_schedule(candidate_path: str | None) -> tuple:
    """
    Attempts to load the rent calculator workbook.
    Returns (schedule, error_message, resolved_path).
    """
    paths = []
    if candidate_path:
        paths.append(candidate_path)
    default_path = os.path.join(os.getcwd(), "2025 AMI Rent Calculator Unlocked.xlsx")
    if default_path not in paths:
        paths.append(default_path)

    for path in paths:
        if not path or not os.path.exists(path):
            continue
        try:
            return load_rent_schedule(path), None, path
        except Exception as exc:
            last_error = str(exc)
            continue
    return None, locals().get("last_error"), None


def _apply_rent_metrics(schedule, scenarios, utilities):
    for scenario in scenarios.values():
        if not scenario:
            continue
        assignments, rent_totals = compute_rents_for_assignments(
            schedule, scenario['assignments'], utilities
        )
        scenario['assignments'] = assignments
        scenario['rent_breakdown'] = rent_totals
        metrics = scenario.get('metrics', {})
        metrics['total_monthly_rent'] = rent_totals['net_monthly']
        metrics['total_annual_rent'] = rent_totals['net_annual']
        metrics['gross_monthly_rent'] = rent_totals['gross_monthly']
        metrics['gross_annual_rent'] = rent_totals['gross_annual']
        metrics['allowance_monthly_total'] = rent_totals['allowances_monthly']
        metrics['allowance_annual_total'] = rent_totals['allowances_annual']
        metrics['allowance_breakdown'] = rent_totals['allowances_breakdown']
        scenario['metrics'] = metrics


def main(file_path, utilities=None, overrides=None, rent_calculator_path=None):
    """Main function to orchestrate the entire optimization process."""
    try:
        analysis_id = uuid.uuid4().hex[:8]
        analysis_started = time.perf_counter()

        utilities_clean = _sanitize_utilities(utilities if isinstance(utilities, dict) else None)
        overrides_payload = overrides if isinstance(overrides, dict) else None
        rent_schedule, rent_schedule_error, rent_schedule_path = _maybe_load_rent_schedule(rent_calculator_path)

        config = load_config()
        parser = Parser(file_path)
        df_affordable = parser.get_affordable_units()

        solver_diagnostics: List[Dict[str, Any]] = []
        base_index = len(solver_diagnostics)
        solver_results = find_optimal_scenarios(
            df_affordable,
            config,
            diagnostics=solver_diagnostics,
            project_overrides=overrides_payload,
        )
        for entry in solver_diagnostics[base_index:]:
            entry['phase'] = 'standard'

        scenarios = solver_results.get("scenarios", {})
        notes = solver_results.get("notes", [])

        max_share_cap = config['optimization_rules'].get('deep_affordability_max_share')
        min_share_cap = config['optimization_rules'].get('deep_affordability_min_share')
        if (not scenarios.get("absolute_best")) and max_share_cap is not None and min_share_cap is not None:
            widen_step = config['optimization_rules'].get('deep_affordability_widen_step', 0.005)
            widen_cap_limit = config['optimization_rules'].get('deep_affordability_widen_cap', 0.4)
            widened_solution_found = False
            if widen_step > 0:
                candidate_cap = max_share_cap + widen_step
                while candidate_cap <= widen_cap_limit and not widened_solution_found:
                    attempt_config = copy.deepcopy(config)
                    attempt_config['optimization_rules']['deep_affordability_max_share'] = candidate_cap
                    notes.append(
                        f"No solution satisfied the deep-affordability cap of {max_share_cap*100:.1f}% share. Retrying with a cap of {candidate_cap*100:.2f}%."
                    )
                    base_index = len(solver_diagnostics)
                    attempt_results = find_optimal_scenarios(
                        df_affordable,
                        attempt_config,
                        diagnostics=solver_diagnostics,
                        project_overrides=overrides_payload,
                    )
                    for entry in solver_diagnostics[base_index:]:
                        entry['phase'] = 'deep_affordability_widened'
                        entry['widened_cap'] = candidate_cap
                    attempt_scenarios = attempt_results.get("scenarios", {})
                    if attempt_scenarios.get("absolute_best"):
                        notes.append(
                            f"40% share cap widened to {candidate_cap*100:.2f}% to satisfy the deep-affordability requirement."
                        )
                        notes.extend(attempt_results.get("notes", []))
                        scenarios = attempt_scenarios
                        config = attempt_config
                        widened_solution_found = True
                        break
                    else:
                        notes.extend(attempt_results.get("notes", []))
                        candidate_cap = round(candidate_cap + widen_step, 10)
            if not widened_solution_found:
                notes.append(
                    f"No solution satisfied the deep-affordability cap of {max_share_cap*100:.1f}% share even after widening; retrying without that constraint."
                )
                relaxed_config = copy.deepcopy(config)
                relaxed_config['optimization_rules'].pop('deep_affordability_max_share', None)
                base_index = len(solver_diagnostics)
                solver_results = find_optimal_scenarios(
                    df_affordable,
                    relaxed_config,
                    diagnostics=solver_diagnostics,
                    project_overrides=overrides_payload,
                )
                for entry in solver_diagnostics[base_index:]:
                    entry['phase'] = 'deep_affordability_relaxed'
                scenarios = solver_results.get("scenarios", {})
                notes.extend(solver_results.get("notes", []))

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
                project_overrides=overrides_payload,
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

        if rent_schedule:
            _apply_rent_metrics(rent_schedule, scenarios, utilities_clean)
        elif rent_schedule_error:
            notes.append(f"Rent calculator warning: {rent_schedule_error}")

        unique_scenario_count = len([s for s in scenarios.values() if s])
        analysis_duration = time.perf_counter() - analysis_started
        analysis_meta = {
            "analysis_id": analysis_id,
            "duration_sec": analysis_duration,
            "solver_combination_count": len(solver_diagnostics),
            "solver_unique_scenarios": unique_scenario_count,
            "timestamp": time.time(),
            "truncated": any(
                ("Search stopped after" in note) or ("Solver interrupted" in note) for note in notes
            ),
        }

        output = {
            "project_summary": {
                "total_affordable_sf": df_affordable['net_sf'].sum(),
                "total_affordable_units": len(df_affordable),
                "utility_selections": utilities_clean,
            },
            "analysis_notes": notes,
            "compliance_report": compliance_report,
            "scenarios": scenarios,
            "solver_diagnostics": solver_diagnostics,
            "analysis_meta": analysis_meta,
        }

        best_metrics = scenarios["absolute_best"].get('metrics', {}) if scenarios.get("absolute_best") else {}
        output["project_summary"].update({
            "forty_percent_units": best_metrics.get('low_band_units'),
            "forty_percent_sf": best_metrics.get('low_band_sf'),
            "forty_percent_share": best_metrics.get('low_band_share'),
            "total_monthly_rent": best_metrics.get('total_monthly_rent'),
            "total_annual_rent": best_metrics.get('total_annual_rent'),
            "total_gross_monthly_rent": best_metrics.get('gross_monthly_rent'),
            "total_gross_annual_rent": best_metrics.get('gross_annual_rent'),
            "total_rent_deductions_monthly": best_metrics.get('allowance_monthly_total'),
            "total_rent_deductions_annual": best_metrics.get('allowance_annual_total'),
        })

        if rent_schedule_path:
            output["rent_workbook"] = {"source_path": rent_schedule_path}

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
