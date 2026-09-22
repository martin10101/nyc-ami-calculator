"""Floor-spread rule (Fix 5 of the 2026-09-02 client feedback).

HPD reviewer on Building D: options 1-4 had every 40% unit on floors 4-13
and none on the upper stories (14-19). HPD Design Guidelines 2026 s4.1.4-6
require bands to be "distributed throughout the building ... vertically";
there is no numeric rule, so we encode the reviewer's test: split the floors
spanned by the pool into thirds; every band with >= 3 units must have a unit
in each third.

Opt-in (`floor_spread` payload field). Absent = byte-identical legacy output.
When NO combo can satisfy the rule the API falls back honestly: rule off,
note added, per-scenario reviewer view still shows the truth.
"""
import pandas as pd
import pytest

from ami_optix.solver import (
    find_max_revenue_scenario,
    floor_spread_rule_from,
    floor_spread_summary,
    floor_thirds,
)


# ---------------------------------------------------------------------------
# Helpers: thirds + rule normalization
# ---------------------------------------------------------------------------

def _df(floors):
    return pd.DataFrame({
        'unit_id': [f'U{i}' for i in range(len(floors))],
        'bedrooms': [1] * len(floors),
        'net_sf': [500.0] * len(floors),
        'floor': floors,
        'client_ami': [0.6] * len(floors),
    })


def test_thirds_building_d_split_matches_reviewer():
    # Pool floors 3..19 -> 3-7 / 8-13 / 14-19 (remainder to the upper thirds).
    thirds = floor_thirds(_df(list(range(3, 20))))
    assert [(t['min_floor'], t['max_floor']) for t in thirds] == [(3, 7), (8, 13), (14, 19)]
    assert [t['label'] for t in thirds] == ['lower', 'middle', 'upper']


def test_thirds_even_split_and_membership():
    thirds = floor_thirds(_df([1, 2, 3, 4, 5, 6]))
    assert [(t['min_floor'], t['max_floor']) for t in thirds] == [(1, 2), (3, 4), (5, 6)]
    assert thirds[0]['indices'] == [0, 1]
    assert thirds[2]['indices'] == [4, 5]


def test_thirds_need_three_distinct_floors_and_a_floor_column():
    assert floor_thirds(_df([1, 1, 2, 2])) is None
    assert floor_thirds(_df([None, None, None])) is None
    assert floor_thirds(_df([1, 2, 3]).drop(columns=['floor'])) is None


def test_thirds_ignore_units_without_a_floor():
    thirds = floor_thirds(_df([1, 2, 3, None]))
    assert sum(len(t['indices']) for t in thirds) == 3


def test_rule_normalization():
    assert floor_spread_rule_from(None) is None
    assert floor_spread_rule_from(False) is None
    assert floor_spread_rule_from(True) == {'min_units_per_band': 3}
    assert floor_spread_rule_from({'min_units_per_band': 2}) == {'min_units_per_band': 2}
    assert floor_spread_rule_from({'min_units_per_band': 'x'}) == {'min_units_per_band': 3}
    assert floor_spread_rule_from({'min_units_per_band': 0}) == {'min_units_per_band': 1}


def test_summary_flags_the_reviewer_complaint():
    thirds = floor_thirds(_df(list(range(3, 20))))
    # 40% units only on floors 4-13 (Building D options 1-4), 70% everywhere.
    assignments = []
    for f in range(3, 20):
        assignments.append({'unit_id': f'U{f}', 'assigned_ami': 0.4 if 4 <= f <= 13 else 0.7, 'floor': f})
    s = floor_spread_summary(assignments, thirds)
    assert s['satisfied'] is False
    assert any('40% AMI has no unit on the upper floors (14-19)' == m for m in s['missing'])
    upper = next(t for t in s['thirds'] if t['label'] == 'upper')
    assert upper['by_band']['40'] == 0
    assert upper['by_band']['70'] == 6


def test_summary_exempts_small_bands():
    thirds = floor_thirds(_df([1, 2, 3, 4, 5, 6]))
    assignments = [
        {'unit_id': 'a', 'assigned_ami': 0.4, 'floor': 1},
        {'unit_id': 'b', 'assigned_ami': 0.4, 'floor': 2},   # only 2 units at 40 -> exempt
        {'unit_id': 'c', 'assigned_ami': 0.8, 'floor': 2},
        {'unit_id': 'd', 'assigned_ami': 0.8, 'floor': 3},
        {'unit_id': 'e', 'assigned_ami': 0.8, 'floor': 5},
        {'unit_id': 'f', 'assigned_ami': 0.8, 'floor': 6},
    ]
    assert floor_spread_summary(assignments, thirds)['satisfied'] is True


# ---------------------------------------------------------------------------
# Solver: the rule actually binds (and stays rent-neutral)
# ---------------------------------------------------------------------------

def _config(rule=None):
    rules = {
        'waami_cap_percent': 70.0,
        'max_bands_per_scenario': 2,
        'potential_bands': [40, 80],
        'share_thresholds': [
            {'band_threshold': 40, 'min_share': 0.2, 'max_share': 0.6, 'denominator': 'affordable'},
        ],
    }
    if rule is not None:
        rules['floor_spread'] = rule
    return {
        'developer_preferences': {
            'premium_score_weights': {'floor': 0.45, 'net_sf': 0.30, 'bedrooms': 0.15, 'balcony': 0.10}
        },
        'optimization_rules': rules,
    }


def _six_identical_units():
    return pd.DataFrame({
        'unit_id': [f'U{i}' for i in range(1, 7)],
        'bedrooms': [1] * 6,
        'net_sf': [100.0] * 6,
        'floor': [1, 2, 3, 4, 5, 6],
        'balcony': [0] * 6,
        'client_ami': [1.0] * 6,
    })


def _rents(df):
    return {40: [80000] * len(df), 80: [160000] * len(df)}


def test_rule_off_sinks_forty_low_and_breaks_spread():
    df = _six_identical_units()
    r = find_max_revenue_scenario(df, _config(), _rents(df), waami_floor=0.5, low_band_floor_tiebreak=True)
    assert r is not None
    floors40 = sorted(int(u['floor']) for u in r['assignments'] if u['assigned_ami'] <= 0.4)
    assert floors40 == [1, 2]  # legacy tie-break: 40% sinks to the lowest floors
    s = floor_spread_summary(r['assignments'], floor_thirds(df))
    assert s['satisfied'] is False  # 80% band (4 units) has nothing on floors 1-2


def test_rule_on_spreads_without_losing_rent():
    df = _six_identical_units()
    off = find_max_revenue_scenario(df, _config(), _rents(df), waami_floor=0.5, low_band_floor_tiebreak=True)
    on = find_max_revenue_scenario(df, _config(True), _rents(df), waami_floor=0.5, low_band_floor_tiebreak=True)
    assert on is not None
    assert on['rent_score'] == off['rent_score']  # floors are rent-neutral
    s = floor_spread_summary(on['assignments'], floor_thirds(df))
    assert s['satisfied'] is True, s['missing']
    # 40% units no longer bunched at the bottom: their average floor is near the pool average (3.5).
    floors40 = [int(u['floor']) for u in on['assignments'] if u['assigned_ami'] <= 0.4]
    assert abs(sum(floors40) / len(floors40) - 3.5) <= 1.0


def test_rule_infeasible_gives_no_scenario_for_that_combo():
    # Only one unit on the top floor; with min_units_per_band=1 every band
    # needs a unit there -> 2 bands can never both comply.
    df = pd.DataFrame({
        'unit_id': [f'U{i}' for i in range(7)],
        'bedrooms': [1] * 7,
        'net_sf': [100.0] * 7,
        'floor': [1, 1, 1, 2, 2, 2, 3],
        'balcony': [0] * 7,
        'client_ami': [1.0] * 7,
    })
    off = find_max_revenue_scenario(df, _config(), _rents(df), waami_floor=0.5)
    on = find_max_revenue_scenario(df, _config({'min_units_per_band': 1}), _rents(df), waami_floor=0.5)
    assert off is not None
    assert on is None


# ---------------------------------------------------------------------------
# End-to-end through /api/optimize (Building D shape, MIH Option 1)
# ---------------------------------------------------------------------------

def _pool_units():
    typed = ([0.4] * 8) + ([0.7] * 11) + ([0.8] * 6)
    units = []
    for i in range(25):
        units.append({
            'unit_id': f'S-{i+1}',
            'bedrooms': [0, 1, 1, 2, 2][i % 5],
            'net_sf': 380.0 + (i * 23) % 420,
            'floor': 3 + (i % 17),
            'balcony': i % 3 == 0,
            'client_ami': typed[i],
        })
    return units


def _post(units, extra):
    from app import app
    client = app.test_client()
    pool_sf = sum(u['net_sf'] for u in units)
    body = {
        'program': 'MIH',
        'mih_option': 'Option 1',
        'mih_residential_sf': pool_sf / 0.254,
        'utilities': {'electricity': 'na', 'cooking': 'na', 'heat': 'na', 'hot_water': 'na'},
        'units': units,
    }
    body.update(extra)
    resp = client.post('/api/optimize', json=body)
    assert resp.status_code == 200
    return resp.get_json()


def _optimized(data):
    return {k: v for k, v in (data.get('scenarios') or {}).items() if v and k != 'original'}


def test_api_absent_field_has_no_floor_spread_metadata():
    data = _post(_pool_units(), {})
    assert data['success'] is True
    assert 'floor_spread' not in data['project_summary']
    assert all('floor_spread' not in sc for sc in _optimized(data).values())


def test_api_rule_on_every_scenario_spreads_bands_across_thirds():
    data = _post(_pool_units(), {'floor_spread': True})
    assert data['success'] is True
    fs = data['project_summary']['floor_spread']
    assert fs['requested'] is True
    assert fs['applied'] is True
    assert [(t['min_floor'], t['max_floor']) for t in fs['thirds']] == [(3, 7), (8, 13), (14, 19)]
    scen = _optimized(data)
    assert scen
    for key, sc in scen.items():
        summary = sc.get('floor_spread')
        assert summary, f'{key}: missing reviewer view'
        assert summary['satisfied'] is True, f'{key}: {summary["missing"]}'
        # The reviewer's exact complaint: 40% must reach the upper floors.
        upper = next(t for t in summary['thirds'] if t['label'] == 'upper')
        if upper['by_band'].get('40') is not None and sum(
            1 for u in sc['assignments'] if int(round(float(u['assigned_ami']) * 100)) == 40
        ) >= 3:
            assert upper['by_band']['40'] >= 1, f'{key}: no 40% unit on the upper floors'
    assert any('floor-spread' in n.lower() or 'floor spread' in n.lower() for n in data.get('notes') or [])


def test_api_falls_back_honestly_when_rule_cannot_be_met():
    # One unit on the top floor + "every band needs a unit in every third"
    # (min_units_per_band=1) -> no combo can comply -> rule off, note, truth shown.
    units = []
    floors = [1] * 5 + [2] * 5 + [3]
    for i, f in enumerate(floors):
        units.append({'unit_id': f'F-{i+1}', 'bedrooms': 1, 'net_sf': 400.0 + 5 * (i % 3), 'floor': f, 'client_ami': 0.6})
    data = _post(units, {'floor_spread': {'min_units_per_band': 1}})
    assert data['success'] is True, data.get('error')
    fs = data['project_summary']['floor_spread']
    assert fs['requested'] is True
    assert fs['applied'] is False
    assert 'could not be satisfied' in fs['reason']
    assert any('could not be satisfied' in n for n in data.get('notes') or [])
    scen = _optimized(data)
    assert scen
    assert all(sc['floor_spread']['satisfied'] is False for sc in scen.values())


def test_api_rule_skipped_without_floor_data():
    units = [dict(u) for u in _pool_units()]
    for u in units:
        u.pop('floor', None)
    data = _post(units, {'floor_spread': True})
    assert data['success'] is True
    fs = data['project_summary']['floor_spread']
    assert fs['requested'] is True
    assert fs['applied'] is False
    assert 'floor' in fs['reason'].lower()
