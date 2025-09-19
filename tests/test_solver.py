import pytest
import pandas as pd
from ami_optix.solver import calculate_premium_scores, find_optimal_scenarios

@pytest.fixture
def sample_config():
    """Provides a sample config for solver tests."""
    return {
        'developer_preferences': {
            'premium_score_weights': {
                'floor': 0.45,
                'net_sf': 0.30,
                'bedrooms': 0.15,
                'balcony': 0.10
            }
        },
        'optimization_rules': {
            'waami_cap_percent': 60.0,
            'max_bands_per_scenario': 3,
            'potential_bands': [40, 80] # Simplified for testing
        }
    }

@pytest.fixture
def sample_affordable_df():
    """Provides a simple DataFrame of affordable units for testing."""
    data = {
        'unit_id': ['1A', '2B'],
        'bedrooms': [1, 2],
        'net_sf': [600, 800],
        'floor': [1, 2],
        'balcony': [0, 1], # Add balcony data to make premium score calculation more robust
        'client_ami': [1.0, 1.0] # Not used by solver, but present from parser
    }
    return pd.DataFrame(data)

def test_calculate_premium_scores(sample_affordable_df, sample_config):
    """Tests that premium scores are calculated and the more premium unit gets a higher score."""
    df = calculate_premium_scores(sample_affordable_df, sample_config['developer_preferences'])
    assert 'premium_score' in df.columns
    # Unit 2B is on a higher floor, has more bedrooms, more SF, and a balcony, so its score must be higher.
    assert df.loc[df['unit_id'] == '2B', 'premium_score'].iloc[0] > df.loc[df['unit_id'] == '1A', 'premium_score'].iloc[0]
    # In this specific case, since 2B is max on all metrics and 1A is min, 2B score should be 1.0
    assert df.loc[df['unit_id'] == '2B', 'premium_score'].iloc[0] == 1.0


def test_find_optimal_scenarios_success(sample_affordable_df, sample_config):
    """
    Tests the main solver orchestrator on a simple, deterministic problem.
    The solver should assign the higher 80% band to the more premium unit ('2B')
    to maximize the WAAMI.
    """
    # Increase the WAAMI cap for this test to make both outcomes valid, forcing the solver to optimize.
    sample_config['optimization_rules']['waami_cap_percent'] = 70.0
    solver_results = find_optimal_scenarios(sample_affordable_df, sample_config)
    scenarios = solver_results["scenarios"]

    assert scenarios.get("absolute_best"), "Solver should have found an absolute_best solution."

    top_scenario = scenarios["absolute_best"][0]
    assert top_scenario['status'] == 'OPTIMAL'
    assert top_scenario['bands'] == [40, 80]

    assignments = top_scenario['assignments']
    unit_2b_assignment = next(u for u in assignments if u['unit_id'] == '2B')
    unit_1a_assignment = next(u for u in assignments if u['unit_id'] == '1A')

    # Assert that the more valuable unit gets the higher AMI band
    assert unit_2b_assignment['assigned_ami'] == 0.80
    assert unit_1a_assignment['assigned_ami'] == 0.40

    # Check that the WAAMI is calculated correctly
    expected_waami = ((800 * 0.80) + (600 * 0.40)) / (800 + 600)
    assert abs(top_scenario['waami'] - expected_waami) < 1e-9

    # Check that the best_2_band scenario was found
    assert "best_2_band" in scenarios
    assert len(scenarios["best_2_band"]["bands"]) == 2


def test_no_solution_found(sample_affordable_df, sample_config):
    """
    Tests that the solver returns an empty dictionary when no solution is possible.
    """
    # Set an impossibly low WAAMI cap
    sample_config['optimization_rules']['waami_cap_percent'] = 30.0

    solver_results = find_optimal_scenarios(sample_affordable_df, sample_config)
    scenarios = solver_results["scenarios"]

    assert not scenarios.get("absolute_best"), "Solver should not find an absolute_best solution."
    assert not scenarios.get("client_oriented"), "Solver should not find a client_oriented solution."

def test_deep_affordability_constraint(sample_config):
    """
    Tests that the deep affordability constraint is applied when total SF > 10,000.
    """
    # 10 units, total SF = 10,010, which is over the threshold
    data = {
        'unit_id': [f'U{i}' for i in range(10)],
        'bedrooms': [1] * 10,
        'net_sf': [1001] * 10,
        'floor': [1] * 10,
        'balcony': [0] * 10,
        'client_ami': [1.0] * 10
    }
    df = pd.DataFrame(data)

    # Use bands that include a low-affordability option
    sample_config['optimization_rules']['potential_bands'] = [40, 100]
    sample_config['optimization_rules']['deep_affordability_sf_threshold'] = 10000

    solver_results = find_optimal_scenarios(df, sample_config)
    scenarios = solver_results["scenarios"]

    assert scenarios.get("absolute_best"), "A solution should be found."

    top_scenario = scenarios["absolute_best"][0]

    # Count units assigned to the 40% AMI band
    low_band_units = [u for u in top_scenario['assignments'] if u['assigned_ami'] <= 0.40]

    # 20% of 10 units is 2. The constraint should force at least 2 units into the 40% band.
    assert len(low_band_units) >= 2

def test_client_oriented_scenario_logic(sample_config):
    """
    Tests that the client-oriented scenario prioritizes premium scores, even at a
    slight WAAMI cost compared to the absolute best.
    """
    # Unit A: High premium score, medium SF
    # Unit B: Low premium score, high SF
    data = {
        'unit_id': ['A', 'B'],
        'bedrooms': [3, 1], # Contributes to premium
        'net_sf': [800, 850], # B has higher SF
        'floor': [10, 1], # A is on a higher floor
        'balcony': [1, 0], # A has a balcony
        'client_ami': [1.0, 1.0]
    }
    df = pd.DataFrame(data)

    # With these weights, Unit A will have a much higher premium score.
    # The absolute_best solver should give the 100% band to Unit B (more SF).
    # The client_oriented solver should give the 100% band to Unit A (more premium).
    sample_config['optimization_rules']['potential_bands'] = [50, 100]
    sample_config['optimization_rules']['waami_cap_percent'] = 80.0

    solver_results = find_optimal_scenarios(df, sample_config)
    scenarios = solver_results["scenarios"]

    abs_best_assignments = {u['unit_id']: u['assigned_ami'] for u in scenarios['absolute_best'][0]['assignments']}
    client_oriented_assignments = {u['unit_id']: u['assigned_ami'] for u in scenarios['client_oriented'][0]['assignments']}

    # Assert that the absolute_best scenario maximized SF * AMI
    assert abs_best_assignments['B'] == 1.0
    assert abs_best_assignments['A'] == 0.5

    # Assert that the client_oriented scenario was nudged to prefer the premium unit
    assert client_oriented_assignments['A'] == 1.0
    assert client_oriented_assignments['B'] == 0.5

def test_waami_floor_constraint(sample_affordable_df, sample_config):
    """Tests that the waami_floor constraint correctly finds solutions above a certain floor."""
    sample_config['optimization_rules']['waami_cap_percent'] = 80.0
    sample_config['optimization_rules']['potential_bands'] = [40, 80]

    # With a floor of 75%, the only possible solution is to assign 80% to both units.
    solver_results = find_optimal_scenarios(sample_affordable_df, sample_config, relaxed_floor=0.75)

    assert solver_results["scenarios"].get("absolute_best"), "A solution should have been found."
    top_scenario = solver_results["scenarios"]["absolute_best"][0]

    # The WAAMI should be exactly 80%
    assert abs(top_scenario['waami'] - 0.80) < 1e-9

def test_finds_exact_60_percent_solution(sample_config):
    """
    Tests that the solver can find a solution exactly at the 60% cap,
    using the user-provided 5-unit ground truth case.
    """
    data = {
        'unit_id': ['2A', '3A', '4A', '5A', '6E'],
        'bedrooms': [0, 0, 0, 0, 0],
        'net_sf': [370, 370, 370, 370, 309],
        'floor': [2, 3, 4, 5, 6],
        'client_ami': [0.6] * 5
    }
    df = pd.DataFrame(data)

    sample_config['optimization_rules']['potential_bands'] = [40, 60, 100]

    solver_results = find_optimal_scenarios(df, sample_config)
    scenarios = solver_results["scenarios"]

    assert scenarios.get("absolute_best"), "A solution should be found."
    top_scenario = scenarios["absolute_best"][0]

    # Check that the WAAMI is exactly 0.6, allowing for tiny float precision errors
    assert abs(top_scenario['waami'] - 0.600000) < 1e-9
