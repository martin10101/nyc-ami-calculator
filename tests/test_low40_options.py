"""Tests for the low-40 options ladder (client request 2026-06).

Covers the three new solver capabilities (all default-off):
  - low_band_unit_count: pin the exact number of units at bands <= 40%
  - differ_from: require the solution to differ from prior layouts
  - low_band_floor_tiebreak: among equal-rent optima, sink 40% units to
    the lowest floors (free — 40% rent never depends on floor)

Plus the API-level ladder: /api/optimize should return multiple distinct
LOW 40 SHARE options, never outcome-duplicates.
"""

import pandas as pd
import pytest

from ami_optix.solver import find_max_revenue_scenario


def _basic_config():
    return {
        'developer_preferences': {
            'premium_score_weights': {
                'floor': 0.45,
                'net_sf': 0.30,
                'bedrooms': 0.15,
                'balcony': 0.10,
            }
        },
        'optimization_rules': {
            'waami_cap_percent': 70.0,
            'max_bands_per_scenario': 2,
            'potential_bands': [40, 80],
            'share_thresholds': [
                {'band_threshold': 40, 'min_share': 0.2, 'max_share': 0.6, 'denominator': 'affordable'},
            ],
        },
    }


def _six_identical_units():
    # Identical SF + bedrooms means identical rents, so WHICH units land at
    # 40% is a pure tie — perfect for testing the tie-break and differ_from.
    return pd.DataFrame({
        'unit_id': [f'U{i}' for i in range(1, 7)],
        'bedrooms': [1] * 6,
        'net_sf': [100.0] * 6,
        'floor': [1, 2, 3, 4, 5, 6],
        'balcony': [0] * 6,
        'client_ami': [1.0] * 6,
    })


def _rents_for(df):
    # 40% pays $800, 80% pays $1600 per unit (cents).
    return {
        40: [80000] * len(df),
        80: [160000] * len(df),
    }


def test_rent_max_unconstrained_uses_fewest_40_units():
    df = _six_identical_units()
    result = find_max_revenue_scenario(
        df, _basic_config(), rent_by_band_cents=_rents_for(df), waami_floor=0.0,
    )
    assert result and result['status'] == 'OPTIMAL'
    low_units = [u for u in result['assignments'] if u['assigned_ami'] <= 0.40]
    # min_share 0.2 of 600sf = 120sf -> two 100sf units required; rent-max
    # never gives up a third unit's 80% rent voluntarily.
    assert len(low_units) == 2


def test_low_band_unit_count_pins_exact_count():
    df = _six_identical_units()
    result = find_max_revenue_scenario(
        df, _basic_config(), rent_by_band_cents=_rents_for(df), waami_floor=0.0,
        low_band_unit_count=3,
    )
    assert result and result['status'] == 'OPTIMAL'
    low_units = [u for u in result['assignments'] if u['assigned_ami'] <= 0.40]
    assert len(low_units) == 3


def test_low_band_unit_count_infeasible_returns_none():
    df = _six_identical_units()
    # 5 units at 40% = 500sf = 83% of affordable SF, over the 60% max share.
    result = find_max_revenue_scenario(
        df, _basic_config(), rent_by_band_cents=_rents_for(df), waami_floor=0.0,
        low_band_unit_count=5,
    )
    assert result is None


def test_differ_from_forces_a_different_layout():
    df = _six_identical_units()
    config = _basic_config()
    rents = _rents_for(df)

    first = find_max_revenue_scenario(
        df, config, rent_by_band_cents=rents, waami_floor=0.0,
        low_band_floor_tiebreak=True,
    )
    assert first and first['status'] == 'OPTIMAL'

    second = find_max_revenue_scenario(
        df, config, rent_by_band_cents=rents, waami_floor=0.0,
        differ_from=[(list(first['canonical_assignments']), 2)],
    )
    assert second and second['status'] == 'OPTIMAL'

    first_map = {u: b for u, b in first['canonical_assignments']}
    second_map = {u: b for u, b in second['canonical_assignments']}
    changed = sum(1 for u in first_map if first_map[u] != second_map.get(u))
    assert changed >= 2
    # Identical units -> the alternative layout costs zero rent.
    assert second['rent_score'] == first['rent_score']


def test_floor_tiebreak_sinks_40_units_to_lowest_floors():
    df = _six_identical_units()
    result = find_max_revenue_scenario(
        df, _basic_config(), rent_by_band_cents=_rents_for(df), waami_floor=0.0,
        low_band_floor_tiebreak=True,
    )
    assert result and result['status'] == 'OPTIMAL'
    low_floors = sorted(
        int(u['floor']) for u in result['assignments'] if u['assigned_ami'] <= 0.40
    )
    # The two 40% units must be the two lowest floors.
    assert low_floors == [1, 2]


def test_floor_tiebreak_never_costs_rent():
    df = _six_identical_units()
    config = _basic_config()
    rents = _rents_for(df)
    plain = find_max_revenue_scenario(
        df, config, rent_by_band_cents=rents, waami_floor=0.0,
    )
    tiebroken = find_max_revenue_scenario(
        df, config, rent_by_band_cents=rents, waami_floor=0.0,
        low_band_floor_tiebreak=True,
    )
    assert plain['rent_score'] == tiebroken['rent_score']


def _ladder_units():
    """20 varied units so several distinct low-40 layouts exist:
    different unit counts can hit the [10%, 10.5%] window, and different
    bedroom mixes at 40% produce genuinely different rents."""
    units = []
    sf_by_tier = {
        0: [400, 410, 420, 430, 440],
        1: [500, 510, 520, 530, 540],
        2: [600, 610, 620, 630, 640],
        3: [700, 710, 720, 730, 740],
    }
    i = 0
    for bedrooms, sfs in sf_by_tier.items():
        for sf in sfs:
            i += 1
            units.append({
                'unit_id': f'L{i}',
                'bedrooms': bedrooms,
                'net_sf': sf,
                'floor': i,
                'balcony': False,
            })
    return units


def test_api_ladder_returns_multiple_distinct_low40_options():
    from app import app

    client = app.test_client()
    units = _ladder_units()
    payload = {
        'program': 'MIH',
        'mih_option': 'Option 1',
        'mih_residential_sf': 41000,
        'mih_max_band_percent': 135,
        'utilities': {'electricity': 'na', 'cooking': 'na', 'heat': 'na', 'hot_water': 'na'},
        'units': units,
    }
    resp = client.post('/api/optimize', json=payload)
    assert resp.status_code == 200
    data = resp.get_json()
    assert data['success'] is True
    scenarios = data.get('scenarios') or {}

    if 'low_40_share' not in scenarios:
        pytest.skip('Rent calculator unavailable in this environment; low-40 group not generated.')

    low40_keys = sorted(k for k in scenarios if k.startswith('low_40_share'))
    extra_keys = [k for k in low40_keys if k != 'low_40_share']
    assert len(extra_keys) >= 2, (
        f"Expected at least 2 additional low-40 options for a 20-unit varied building, "
        f"got {low40_keys}"
    )

    residential_sf = 41000.0
    seen_outcomes = set()
    for key in low40_keys:
        scenario = scenarios[key]
        assignments = scenario.get('assignments') or []
        assert assignments, f"{key} has no assignments"

        # Every option must respect the hard WAAMI cap...
        total_sf = sum(float(u['net_sf']) for u in assignments)
        waami = sum(float(u['net_sf']) * float(u['assigned_ami']) for u in assignments) / total_sf
        assert waami <= 0.60 + 1e-9, f"{key} WAAMI {waami} exceeds cap"

        # ...and hug the minimum 40% window (10% floor, narrow ceiling).
        low_sf = sum(
            float(u['net_sf']) for u in assignments
            if float(u['assigned_ami']) <= 0.4 + 1e-12
        )
        low_share = low_sf / residential_sf
        assert low_share >= 0.10 - 1e-9, f"{key} 40-share {low_share} below 10% floor"
        assert low_share <= 0.155 + 1e-9, f"{key} 40-share {low_share} far above the floor window"

        # Outcome distinctness: band unit-counts + monthly rent must differ.
        band_counts = {}
        for u in assignments:
            band = int(round(float(u['assigned_ami']) * 100))
            band_counts[band] = band_counts.get(band, 0) + 1
        rent_monthly = round(float((scenario.get('rent_totals') or {}).get('net_monthly') or 0.0), 2)
        outcome = (tuple(sorted(band_counts.items())), rent_monthly)
        assert outcome not in seen_outcomes, f"{key} duplicates another low-40 option's outcome"
        seen_outcomes.add(outcome)

    # The extra options must carry a client-readable description line.
    for key in extra_keys:
        tradeoffs = scenarios[key].get('tradeoffs') or []
        assert tradeoffs and any('low-40' in str(t).lower() or '40% ami' in str(t).lower() for t in tradeoffs), (
            f"{key} is missing its description line"
        )
