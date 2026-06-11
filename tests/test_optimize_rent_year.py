"""/api/optimize must honor the per-request rent_roll_year (fix 2026-06).

Background: Manual Calculate (/api/evaluate) always sent rent_roll_year and
priced correctly, but Find Optimal Scenarios (/api/optimize) ignored the
Excel year dropdown and silently used the server-global default (2025).
The client compared 2026-priced manual options against 2025-priced optimizer
output and concluded the optimizer was broken.

Known net rents for a 40% AMI 2-bedroom under the standard utility set
(electricity tenant pays, electric stove, ccASHP heat, electric hot water):
  2024 -> $1,099   2025 -> $1,069   2026 -> $1,131
"""

import pytest

from app import app

UTILITIES = {
    'electricity': 'tenant_pays',
    'cooking': 'electric',
    'heat': 'electric_ccashp',
    'hot_water': 'electric_other',
}

EXPECTED_40_2BR_NET = {
    2024: 1099.0,
    2025: 1069.0,
    2026: 1131.0,
}


def _payload(rent_roll_year=None):
    units = [
        {'unit_id': f'U{i}', 'bedrooms': 2, 'net_sf': 100, 'floor': i, 'balcony': False}
        for i in range(1, 11)
    ]
    payload = {
        'program': 'MIH',
        'mih_option': 'Option 1',
        'mih_residential_sf': 1000,
        'mih_max_band_percent': 135,
        'utilities': UTILITIES,
        'units': units,
    }
    if rent_roll_year is not None:
        payload['rent_roll_year'] = rent_roll_year
    return payload


def _net_rent_of_a_40_band_unit(data):
    """Pull the monthly net rent of any 40%-band unit from absolute_best."""
    scenarios = data.get('scenarios') or {}
    best = scenarios.get('absolute_best') or {}
    for unit in best.get('assignments') or []:
        if float(unit.get('assigned_ami') or 0) <= 0.4 + 1e-12:
            rent = unit.get('monthly_rent')
            if rent is not None:
                return round(float(rent), 2)
    return None


@pytest.mark.parametrize('year', [2024, 2025, 2026])
def test_optimize_honors_requested_rent_roll_year(year):
    client = app.test_client()
    resp = client.post('/api/optimize', json=_payload(rent_roll_year=year))
    assert resp.status_code == 200
    data = resp.get_json()
    assert data['success'] is True

    assert data.get('rent_roll_year_used') == year, (
        f"Server reports year {data.get('rent_roll_year_used')}, requested {year}"
    )
    assert any(f'priced on the {year} rent roll' in str(n) for n in data.get('notes') or []), (
        'Missing the year note in response notes'
    )

    net_40_2br = _net_rent_of_a_40_band_unit(data)
    assert net_40_2br is not None, 'No 40%-band unit with a rent found in absolute_best'
    assert net_40_2br == EXPECTED_40_2BR_NET[year], (
        f'40% 2BR net rent {net_40_2br} does not match the {year} table '
        f'({EXPECTED_40_2BR_NET[year]})'
    )


def test_optimize_without_year_keeps_legacy_default_behavior():
    """Old add-in versions send no rent_roll_year — the server must fall back
    to its global default (2025) exactly as before this fix."""
    client = app.test_client()
    resp = client.post('/api/optimize', json=_payload(rent_roll_year=None))
    assert resp.status_code == 200
    data = resp.get_json()
    assert data['success'] is True

    assert data.get('rent_roll_year_used') == 2025

    net_40_2br = _net_rent_of_a_40_band_unit(data)
    assert net_40_2br == EXPECTED_40_2BR_NET[2025]


def test_optimize_unknown_year_falls_back_with_warning():
    """A year the server doesn't have must fall back to the default and SAY so
    — never silently price on a different table."""
    client = app.test_client()
    resp = client.post('/api/optimize', json=_payload(rent_roll_year=2031))
    assert resp.status_code == 200
    data = resp.get_json()
    assert data['success'] is True

    # Fallback lands on the default calculator (2025).
    assert data.get('rent_roll_year_used') == 2025
    notes_text = ' | '.join(str(n) for n in data.get('notes') or [])
    assert 'rent_roll_year not found' in notes_text or 'Warning' in notes_text
