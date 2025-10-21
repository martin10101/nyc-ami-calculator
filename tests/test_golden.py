import pytest
import pandas as pd
from main import main as run_ami_optix_analysis

def test_decatur_golden_truth():
    """
    Tests the entire analysis pipeline against the known "golden truth"
    for the Decatur project. It asserts that the WAAMI and band assignments
    match the established correct answer.
    """
    # Path to the golden file
    file_path = "tests/test_data/Decatur_for_testing.csv"

    # Run the full analysis
    analysis_output = run_ami_optix_analysis(file_path)
    assert "error" not in analysis_output, f"Analysis failed: {analysis_output.get('error')} | Notes: {analysis_output.get('analysis_notes', [])}"

    results = analysis_output.get("results", {})
    scenario_s1 = results.get("scenario_absolute_best")
    assert scenario_s1, "Absolute Best scenario was not found."

    # Assert WAAMI is correct to 6 decimal places
    # The solver finds the true optimum which is exactly 0.6
    assert abs(scenario_s1['waami'] - 0.6) < 1e-6

    # Assert correct bands were used
    assert set(scenario_s1['bands']) == {40, 60, 90}

    # Assert specific unit assignments are correct
    assignments = {str(u['unit_id']): int(u['assigned_ami'] * 100) for u in scenario_s1['assignments']}

    # Ground truth for Decatur project (updated for 20–21% clamp)
    expected_assignments = {
        '2A': 40, '3A': 60, '4A': 60, '5A': 60, '6E': 60,
        '2B': 40, '3B': 60, '4B': 60, '5B': 60, '6B': 90,
        '2D': 40, '3D': 60, '4D': 60, '5D': 60, '6D': 90,
    }

    assert assignments == expected_assignments, f"Assignments did not match. Got: {assignments}"

def test_beach_115_golden_truth():
    """
    Tests the entire analysis pipeline against the known "golden truth"
    for the Beach 115 Street project.
    """
    file_path = "tests/test_data/169_Beach_115_Street_for_testing.csv"

    analysis_output = run_ami_optix_analysis(file_path)
    assert "error" not in analysis_output, f"Analysis failed: {analysis_output.get('error')} | Notes: {analysis_output.get('analysis_notes', [])}"

    results = analysis_output.get("results", {})
    scenario_s1 = results.get("scenario_absolute_best")
    assert scenario_s1, "Absolute Best scenario was not found."

    assert abs(scenario_s1['waami'] - 0.6) < 1e-6
    assert set(scenario_s1['bands']) == {40, 60, 90}

    assignments = {str(u['unit_id']): int(u['assigned_ami'] * 100) for u in scenario_s1['assignments']}

    expected_assignments = {
        '201': 40, '202': 40, '203': 40, '204': 40, '205': 40, '206': 40, '207': 40, '208': 40, '209': 40, '210': 40, '211': 40, '212': 40, '301': 40, '302': 40, '303': 40, '304': 40, '305': 40, '306': 40, '307': 40, '308': 40, '309': 40, '310': 60, '311': 90, '312': 90, '401': 90, '402': 90, '403': 90, '404': 90, '405': 90, '406': 90, '407': 90, '408': 90, '409': 90, '410': 90, '411': 90, '412': 90
    }

    assert assignments == expected_assignments, f"Assignments did not match. Got: {assignments}"

