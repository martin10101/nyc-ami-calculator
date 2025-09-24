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
    assert set(scenario_s1['bands']) == {30, 60, 120}

    # Assert specific unit assignments are correct
    assignments = {str(u['unit_id']): int(u['assigned_ami'] * 100) for u in scenario_s1['assignments']}

    # Ground truth for Decatur project
    expected_assignments = {
        '2A': 30, '3A': 30, '4A': 60, '5A': 120, '6E': 60, '2B': 30, '3B': 30,
        '4B': 30, '5B': 60, '6B': 120, '2D': 30, '3D': 30, '4D': 30, '5D': 120, '6D': 120
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
    assert set(scenario_s1['bands']) == {30, 120}

    assignments = {str(u['unit_id']): int(u['assigned_ami'] * 100) for u in scenario_s1['assignments']}

    expected_assignments = {
        '201': 30, '202': 30, '203': 30, '204': 30, '205': 30, '206': 30,
        '207': 30, '208': 30, '209': 30, '210': 30, '211': 30, '212': 30,
        '301': 30, '302': 30, '303': 30, '304': 30, '305': 30, '306': 30,
        '307': 30, '308': 30, '309': 30, '310': 30, '311': 30, '312': 30,
        '401': 120, '402': 120, '403': 120, '404': 120, '405': 120, '406': 120,
        '407': 120, '408': 120, '409': 120, '410': 120, '411': 120, '412': 120
    }

    assert assignments == expected_assignments, f"Assignments did not match. Got: {assignments}"