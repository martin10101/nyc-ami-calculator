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
    scenarios = find_optimal_scenarios(sample_affordable_df, sample_config)

    assert len(scenarios) > 0, "Solver should have found at least one solution."

    top_scenario = scenarios[0]
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


def test_no_solution_found(sample_affordable_df, sample_config):
    """
    Tests that the solver returns an empty list when no solution is possible.
    """
    # Set an impossibly low WAAMI cap
    sample_config['optimization_rules']['waami_cap_percent'] = 30.0

    scenarios = find_optimal_scenarios(sample_affordable_df, sample_config)

    assert scenarios == [], "Solver should return an empty list for an unsolvable problem."
