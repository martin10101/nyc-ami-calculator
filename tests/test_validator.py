import pytest
import pandas as pd
from ami_optix.validator import run_compliance_checks

@pytest.fixture
def sample_nyc_rules():
    """Provides a sample nyc_rules config for testing."""
    return {
        'validation_checks': {
            'size_minima': {
                'studio': 400,
                'one_bedroom': 575,
                'two_bedroom': 775
            },
            'mix_checks': {
                'max_studio_percent': 25.0,
                'min_two_br_plus_percent': 50.0
            }
        }
    }

def test_all_checks_pass(sample_nyc_rules):
    """Tests a scenario where all compliance checks should pass."""
    data = {
        'unit_id': ['101', '102', '201', '202'],
        'bedrooms': [0, 1, 2, 2],
        'net_sf': [450, 600, 800, 810]
    }
    df = pd.DataFrame(data)
    results = run_compliance_checks(df, sample_nyc_rules)

    statuses = {r['check']: r['status'] for r in results}
    assert statuses['Unit Size Minimum'] == 'PASS'
    assert statuses['Max Studio Percentage'] == 'PASS'
    assert statuses['Min 2+ Bedroom Percentage'] == 'PASS'

def test_unit_size_flagged(sample_nyc_rules):
    """Tests that a unit below the minimum size is flagged."""
    data = {
        'unit_id': ['101'],
        'bedrooms': [1],
        'net_sf': [500]  # Too small for a 1BR (min 575)
    }
    df = pd.DataFrame(data)
    results = run_compliance_checks(df, sample_nyc_rules)

    flagged_check = next(r for r in results if r['check'] == 'Unit Size Minimum')
    assert flagged_check['status'] == 'FLAGGED'
    assert 'below the required 575 SF' in flagged_check['details']

def test_studio_mix_flagged(sample_nyc_rules):
    """Tests that a high percentage of studios is flagged."""
    data = {
        'unit_id': ['101', '102', '201'],
        'bedrooms': [0, 0, 1], # 2 out of 3 are studios (66.7%)
        'net_sf': [450, 460, 600]
    }
    df = pd.DataFrame(data)
    results = run_compliance_checks(df, sample_nyc_rules)

    flagged_check = next(r for r in results if r['check'] == 'Max Studio Percentage')
    assert flagged_check['status'] == 'FLAGGED'
    assert 'exceeding the 25.0% maximum' in flagged_check['details']

def test_two_br_mix_flagged(sample_nyc_rules):
    """Tests that a low percentage of 2+ BR units is flagged."""
    data = {
        'unit_id': ['101', '102', '201', '202'],
        'bedrooms': [0, 1, 1, 1], # 0% are 2+ BR
        'net_sf': [450, 600, 610, 620]
    }
    df = pd.DataFrame(data)
    results = run_compliance_checks(df, sample_nyc_rules)

    flagged_check = next(r for r in results if r['check'] == 'Min 2+ Bedroom Percentage')
    assert flagged_check['status'] == 'FLAGGED'
    assert 'below the required 50.0%' in flagged_check['details']

def test_empty_dataframe(sample_nyc_rules):
    """Tests that the validator handles an empty DataFrame gracefully."""
    df = pd.DataFrame(columns=['unit_id', 'bedrooms', 'net_sf'])
    results = run_compliance_checks(df, sample_nyc_rules)

    statuses = {r['check']: r['status'] for r in results}
    assert statuses['Unit Size Minimum'] == 'PASS'
    assert statuses['Max Studio Percentage'] == 'PASS'
    assert statuses['Min 2+ Bedroom Percentage'] == 'PASS'
    assert 'N/A' in (next(r for r in results if r['check'] == 'Max Studio Percentage')['details'])
