import pytest
import pandas as pd
from main import main as run_ami_optix_analysis
from ami_optix.solver import calculate_premium_scores, find_optimal_scenarios

@pytest.fixture
def sample_config():
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
            'potential_bands': [40, 80]
        }
    }

@pytest.fixture
def sample_affordable_df():
    data = {
        'unit_id': ['1A', '2B'],
        'bedrooms': [1, 2],
        'net_sf': [600, 800],
        'floor': [1, 2],
        'balcony': [0, 1],
        'client_ami': [1.0, 1.0]
    }
    return pd.DataFrame(data)

def test_calculate_premium_scores(sample_affordable_df, sample_config):
    df = calculate_premium_scores(sample_affordable_df, sample_config['developer_preferences'])
    assert 'premium_score' in df.columns
    assert df.loc[df['unit_id'] == '2B', 'premium_score'].iloc[0] > df.loc[df['unit_id'] == '1A', 'premium_score'].iloc[0]
    assert df.loc[df['unit_id'] == '2B', 'premium_score'].iloc[0] == 1.0

def test_find_optimal_scenarios_success(sample_affordable_df, sample_config):
    sample_config['optimization_rules']['waami_cap_percent'] = 70.0
    solver_results = find_optimal_scenarios(sample_affordable_df, sample_config)
    scenarios = solver_results["scenarios"]

    assert scenarios.get("absolute_best")
    top_scenario = scenarios["absolute_best"]
    assert top_scenario['status'] == 'OPTIMAL'
    assert top_scenario['bands'] == [40, 80]

    assignments = top_scenario['assignments']
    unit_2b_assignment = next(u for u in assignments if u['unit_id'] == '2B')
    unit_1a_assignment = next(u for u in assignments if u['unit_id'] == '1A')
    assert unit_2b_assignment['assigned_ami'] == 0.80
    assert unit_1a_assignment['assigned_ami'] == 0.40

    expected_waami = ((800 * 0.80) + (600 * 0.40)) / (800 + 600)
    assert abs(top_scenario['waami'] - expected_waami) < 1e-9

    best_2 = scenarios.get("best_2_band")
    if best_2:
        assert len(best_2['bands']) == 2
    else:
        assert any('No viable 2-band solution' in note for note in solver_results['notes'])


def test_client_oriented_present(sample_affordable_df, sample_config):
    sample_config['optimization_rules']['potential_bands'] = [40, 60, 80]
    solver_results = find_optimal_scenarios(sample_affordable_df, sample_config)
    scenarios = solver_results["scenarios"]

    client = scenarios.get("client_oriented")
    if client:
        assert len(client['bands']) >= 2
    else:
        assert any('Client-oriented scenario unavailable' in note for note in solver_results['notes'])


def test_no_solution_found(sample_affordable_df, sample_config):
    sample_config['optimization_rules']['waami_cap_percent'] = 30.0
    solver_results = find_optimal_scenarios(sample_affordable_df, sample_config)
    scenarios = solver_results["scenarios"]

    assert not scenarios.get("absolute_best")
    assert not scenarios.get("alternative")

def test_deep_affordability_constraint(sample_config):
    data = {
        'unit_id': [f'U{i}' for i in range(10)],
        'bedrooms': [1] * 10,
        'net_sf': [1001] * 10,
        'floor': [1] * 10,
        'balcony': [0] * 10,
        'client_ami': [1.0] * 10
    }
    df = pd.DataFrame(data)

    sample_config['optimization_rules']['potential_bands'] = [40, 100]
    sample_config['optimization_rules']['deep_affordability_sf_threshold'] = 10000

    solver_results = find_optimal_scenarios(df, sample_config)
    scenarios = solver_results["scenarios"]

    assert scenarios.get("absolute_best")
    top_scenario = scenarios["absolute_best"]

    low_band_units = [u for u in top_scenario['assignments'] if u['assigned_ami'] <= 0.40]
    assert len(low_band_units) >= 2


def test_lexicographical_tie_breaking_with_premium_score(sample_config):
    data = {
        'unit_id': ['A', 'B'],
        'bedrooms': [3, 1],
        'net_sf': [1000, 1000],
        'floor': [10, 1],
        'balcony': [1, 0],
        'client_ami': [1.0, 1.0]
    }
    df = pd.DataFrame(data)

    sample_config['optimization_rules']['potential_bands'] = [40, 60]
    sample_config['optimization_rules']['waami_cap_percent'] = 50.0

    solver_results = find_optimal_scenarios(df, sample_config)
    scenarios = solver_results["scenarios"]

    abs_best_assignments = {u['unit_id']: u['assigned_ami'] for u in scenarios['absolute_best']['assignments']}
    assert abs_best_assignments['A'] == 0.60
    assert abs_best_assignments['B'] == 0.40


def test_waami_floor_constraint(sample_affordable_df, sample_config):
    sample_config['optimization_rules']['waami_cap_percent'] = 80.0
    sample_config['optimization_rules']['potential_bands'] = [40, 80]

    solver_results = find_optimal_scenarios(sample_affordable_df, sample_config, relaxed_floor=0.75)

    assert solver_results["scenarios"].get("absolute_best")
    top_scenario = solver_results["scenarios"]["absolute_best"]
    assert abs(top_scenario['waami'] - 0.80) < 1e-9


def test_finds_exact_60_percent_solution(sample_config):
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

    assert scenarios.get("absolute_best")
    top_scenario = scenarios["absolute_best"]
    assert abs(top_scenario['waami'] - 0.600000) < 1e-9


def test_multi_band_preferred_and_two_band_missing(sample_config):
    data = {
        'unit_id': ['A1', 'A2', 'A3'],
        'bedrooms': [1, 1, 1],
        'net_sf': [700, 700, 700],
        'floor': [2, 3, 4],
        'balcony': [0, 0, 0],
        'client_ami': [0.6, 0.6, 0.6]
    }
    df = pd.DataFrame(data)

    rules = sample_config['optimization_rules']
    rules['waami_cap_percent'] = 60.0
    rules['waami_floor'] = 0.60
    rules['potential_bands'] = [40, 60, 80]

    solver_results = find_optimal_scenarios(df, sample_config)
    scenarios = solver_results['scenarios']

    assert scenarios.get('absolute_best')
    assert len(scenarios['absolute_best']['bands']) == 3
    best_2 = scenarios.get('best_2_band')
    if best_2:
        assert len(best_2['bands']) == 2
    else:
        assert any('No viable 2-band solution' in note for note in solver_results['notes'])
    assert all(len(s['bands']) >= 2 for s in scenarios.values())


@pytest.mark.parametrize("workbook", [
    'Test.xlsx',
    '1004 Wodycrest Avenue (1).xlsx',
    '1530 Bergen Street (1).xlsx',
    '169 Beach 115 Street (1).xlsx',
    '212 West 231 Street (1).xlsx',
    '2675 Decatur Avenue (1).xlsx'
])
def test_sample_workbooks_have_unique_valid_scenarios(workbook):
    result = run_ami_optix_analysis(workbook)
    assert 'error' not in result, f"Analysis failed for {workbook}: {result.get('error')}"
    scenarios = result['results']['scenarios']
    waami_floor = 0.58 - 1e-9
    seen_assignments = set()
    for scenario in scenarios.values():
        if not scenario:
            continue
        assert scenario['waami'] >= waami_floor
        assert len(scenario['bands']) >= 2
        canonical = tuple(sorted((unit['unit_id'], unit['assigned_ami']) for unit in scenario['assignments']))
        assert canonical not in seen_assignments
        seen_assignments.add(canonical)
