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


def test_api_fewest_40_units_family():
    """Rachel's concept (2026-06-11): developers minimize the NUMBER of
    apartments at 40% AMI, not the SF share. The fewest_40_units family must
    pin the true minimum unit count over the FULL legal window and offer
    several options at that count.

    For _ladder_units() with residential_sf=41000: required 40-SF = 4,100.
    Largest-first greedy: 740+730+720+710+700 = 3,600 (5 units), +640 =
    4,240 (6 units) >= 4,100 -> the minimum is 6 apartments.
    """
    from app import app

    client = app.test_client()
    payload = {
        'program': 'MIH',
        'mih_option': 'Option 1',
        'mih_residential_sf': 41000,
        'mih_max_band_percent': 135,
        'utilities': {'electricity': 'na', 'cooking': 'na', 'heat': 'na', 'hot_water': 'na'},
        'units': _ladder_units(),
    }
    resp = client.post('/api/optimize', json=payload)
    assert resp.status_code == 200
    data = resp.get_json()
    assert data['success'] is True
    scenarios = data.get('scenarios') or {}

    if 'low_40_share' not in scenarios:
        pytest.skip('Rent calculator unavailable in this environment.')

    fewest = scenarios.get('fewest_40_units')
    assert fewest, f"fewest_40_units missing; keys: {sorted(scenarios.keys())}"

    residential_sf = 41000.0

    def _forty_band_stats(scenario):
        assignments = scenario.get('assignments') or []
        n = sum(1 for u in assignments if float(u['assigned_ami']) <= 0.4 + 1e-12)
        low_sf = sum(
            float(u['net_sf']) for u in assignments
            if float(u['assigned_ami']) <= 0.4 + 1e-12
        )
        total_sf = sum(float(u['net_sf']) for u in assignments)
        waami = sum(float(u['net_sf']) * float(u['assigned_ami']) for u in assignments) / total_sf
        return n, low_sf, waami

    n, low_sf, waami = _forty_band_stats(fewest)
    assert n == 6, f"fewest_40_units used {n} apartments at 40%; minimum is 6"
    assert low_sf / residential_sf >= 0.10 - 1e-9
    assert low_sf / residential_sf <= 0.125 + 1e-9, (
        "fewest option must stay inside the FULL legal window"
    )
    assert waami <= 0.60 + 1e-9

    # The family must offer more than one option (band variety at the
    # minimum count, or at worst one option at minimum+1).
    family_keys = sorted(k for k in scenarios if k.startswith('fewest_40_units'))
    assert len(family_keys) >= 2, f"Expected >=2 fewest-40 options, got {family_keys}"

    seen_outcomes = set()
    for key in family_keys:
        scenario = scenarios[key]
        kn, klow_sf, kwaami = _forty_band_stats(scenario)
        assert kn in (6, 7), f"{key} used {kn} units at 40%; expected the minimum (6) or 6+1"
        assert kwaami <= 0.60 + 1e-9, f"{key} WAAMI {kwaami} exceeds cap"

        band_counts = {}
        for u in scenario.get('assignments') or []:
            band = int(round(float(u['assigned_ami']) * 100))
            band_counts[band] = band_counts.get(band, 0) + 1
        rent_monthly = round(float((scenario.get('rent_totals') or {}).get('net_monthly') or 0.0), 2)
        outcome = (tuple(sorted(band_counts.items())), rent_monthly)
        assert outcome not in seen_outcomes, f"{key} duplicates another option's outcome"
        seen_outcomes.add(outcome)

        # Strategy line ("Why:") lives in description since 2026-06-11;
        # tradeoffs is reserved for real rule relaxations on edge scenarios.
        description = str(scenario.get('description') or '')
        assert 'apartments at 40%' in description.lower() or 'minimum' in description.lower(), (
            f"{key} is missing its strategy description line: {description!r}"
        )


def test_api_mid_40_share_fills_the_middle_of_the_window():
    """Client request 2026-06-11: the low-40 group hugs the 10% floor and the
    unconstrained best often runs to the 12.5% ceiling, leaving 10.5-11.5%
    unexplored. mid_40_share must land inside that middle range."""
    from app import app

    client = app.test_client()
    payload = {
        'program': 'MIH',
        'mih_option': 'Option 1',
        'mih_residential_sf': 41000,
        'mih_max_band_percent': 135,
        'utilities': {'electricity': 'na', 'cooking': 'na', 'heat': 'na', 'hot_water': 'na'},
        'units': _ladder_units(),
    }
    resp = client.post('/api/optimize', json=payload)
    assert resp.status_code == 200
    data = resp.get_json()
    scenarios = data.get('scenarios') or {}

    if 'low_40_share' not in scenarios:
        pytest.skip('Rent calculator unavailable in this environment.')

    mid = scenarios.get('mid_40_share')
    assert mid, f"mid_40_share missing; keys: {sorted(scenarios.keys())}"

    assignments = mid.get('assignments') or []
    low_sf = sum(
        float(u['net_sf']) for u in assignments
        if float(u['assigned_ami']) <= 0.4 + 1e-12
    )
    share = low_sf / 41000.0
    assert 0.105 - 1e-9 <= share <= 0.115 + 1e-9, (
        f"mid_40_share 40-band share {share*100:.2f}% outside [10.5%, 11.5%]"
    )

    # Must be a distinct layout from the floor-hugging and ceiling options.
    others = {
        tuple(map(tuple, (scenarios[k].get('canonical_assignments') or [])))
        for k in scenarios if k != 'mid_40_share'
    }
    mid_canon = tuple(map(tuple, (mid.get('canonical_assignments') or [])))
    assert mid_canon not in others

    # Strategy line lives in description since 2026-06-11.
    assert 'mid-range' in str(mid.get('description') or '').lower()
