"""40% Rules - ribbon "40% Rules" menu (owner request 2026-09-22, Fix 6).

The owner decides WHICH apartments carry the 40% label; the optimizer still
finds the best rent within those decisions. Four independent, optional rules:
pin units at 40%, keep units out of 40%, restrict 40% to bedroom types, cap
40% units per floor. Absent field = the program decides (legacy, identity).
"""
import pandas as pd
import pytest

from ami_optix.solver import find_max_revenue_scenario
from app import _normalize_forty_rules


# ---------------------------------------------------------------------------
# Payload normalization
# ---------------------------------------------------------------------------

def test_normalize_full_payload():
    got = _normalize_forty_rules({
        'pin_units': ['2A', ' 3A ', '2A', None],
        'exclude_units': ['5B'],
        'bedrooms_allowed': [2, '1', 2.0],
        'max_per_floor': '2',
        'floors_allowed': [7, '5', 6.0, 5],
    })
    assert got == {
        'pin_units': ['2A', '3A'], 'exclude_units': ['5B'], 'bedrooms_allowed': [1, 2],
        'max_per_floor': 2, 'floors_allowed': [5, 6, 7],
    }


def test_normalize_floors_only():
    got = _normalize_forty_rules({'floors_allowed': [3, 4]})
    assert got['floors_allowed'] == [3, 4]
    assert got['pin_units'] == [] and got['bedrooms_allowed'] is None and got['max_per_floor'] is None


def test_normalize_absent_or_empty_means_program_decides():
    assert _normalize_forty_rules(None) is None
    assert _normalize_forty_rules('x') is None
    assert _normalize_forty_rules({}) is None
    assert _normalize_forty_rules({'pin_units': [], 'exclude_units': [], 'max_per_floor': 0}) is None


# ---------------------------------------------------------------------------
# Solver: max 40% units per floor
# ---------------------------------------------------------------------------

def _config(extra_rules=None):
    rules = {
        'waami_cap_percent': 70.0,
        'max_bands_per_scenario': 2,
        'potential_bands': [40, 80],
        # exactly 3 of 6 identical units at 40%
        'share_thresholds': [
            {'band_threshold': 40, 'min_share': 0.5, 'max_share': 0.5, 'denominator': 'affordable'},
        ],
    }
    rules.update(extra_rules or {})
    return {
        'developer_preferences': {
            'premium_score_weights': {'floor': 0.45, 'net_sf': 0.30, 'bedrooms': 0.15, 'balcony': 0.10}
        },
        'optimization_rules': rules,
    }


def _units(floors):
    n = len(floors)
    return pd.DataFrame({
        'unit_id': [f'U{i}' for i in range(1, n + 1)],
        'bedrooms': [1] * n,
        'net_sf': [100.0] * n,
        'floor': floors,
        'balcony': [0] * n,
        'client_ami': [1.0] * n,
    })


def _rents(df):
    return {40: [80000] * len(df), 80: [160000] * len(df)}


def _forty_by_floor(result):
    out = {}
    for u in result['assignments']:
        if u['assigned_ami'] <= 0.4:
            out[int(u['floor'])] = out.get(int(u['floor']), 0) + 1
    return out


def test_per_floor_cap_binds_and_is_rent_neutral():
    df = _units([1, 1, 1, 2, 2, 2])
    off = find_max_revenue_scenario(df, _config(), _rents(df), waami_floor=0.5, low_band_floor_tiebreak=True)
    assert off is not None
    assert _forty_by_floor(off) == {1: 3}   # legacy: sinks all three to floor 1
    capped = find_max_revenue_scenario(df, _config({'forty_max_per_floor': 2}), _rents(df), waami_floor=0.5, low_band_floor_tiebreak=True)
    assert capped is not None
    assert max(_forty_by_floor(capped).values()) <= 2
    assert capped['rent_score'] == off['rent_score']


def test_per_floor_cap_too_tight_gives_no_scenario():
    df = _units([1, 1, 1, 2, 2, 2])
    # 3 units at 40% over 2 floors with at most 1 per floor is impossible.
    assert find_max_revenue_scenario(df, _config({'forty_max_per_floor': 1}), _rents(df), waami_floor=0.5) is None


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


def _forty_ids(sc):
    return {str(u['unit_id']) for u in sc['assignments'] if int(round(float(u['assigned_ami']) * 100)) == 40}


def test_api_absent_field_has_no_forty_rules_metadata():
    data = _post(_pool_units(), {})
    assert data['success'] is True
    assert 'forty_rules' not in data['project_summary']
    assert not any('40% rules' in n for n in data.get('notes') or [])


def test_api_pins_and_exclusions_are_obeyed_in_every_scenario():
    units = _pool_units()
    pins = ['S-4', 'S-9']       # both 2 BR
    excl = ['S-1', 'S-2']
    data = _post(units, {'forty_rules': {'pin_units': pins, 'exclude_units': excl}})
    assert data['success'] is True, data.get('error')
    scen = _optimized(data)
    assert scen
    for key, sc in scen.items():
        ids = _forty_ids(sc)
        assert set(pins) <= ids, f'{key}: pinned units missing from 40%: {set(pins) - ids}'
        assert not (set(excl) & ids), f'{key}: excluded unit at 40%'
    fr = data['project_summary']['forty_rules']
    assert fr['pin_units'] == pins and fr['exclude_units'] == excl
    assert 'pinned at 40%: S-4, S-9' in fr['summary']
    assert any('Built with your 40% rules' in n for n in data['notes'])


def test_api_bedroom_preference_honored_when_the_window_allows_it():
    units = _pool_units()
    data = _post(units, {'forty_rules': {'bedrooms_allowed': [2]}})
    assert data['success'] is True, data.get('error')
    beds = {u['unit_id']: u['bedrooms'] for u in units}
    scen = _optimized(data)
    assert scen
    fr = data['project_summary']['forty_rules']
    # The 2 BR units alone can reach the 10% floor here, so the wish is fully
    # honored (restrict) or every 2 BR is 40% (force) - either way only 2 BR at 40%.
    assert fr['level'] in ('restrict', 'force'), fr
    for key, sc in scen.items():
        for uid in _forty_ids(sc):
            assert beds[uid] == 2, f'{key}: {uid} ({beds[uid]} BR) at 40% despite the 2 BR preference'
    assert '2 BR' in fr['summary']
    assert any('Built with your 40% rules' in n for n in data['notes'])


def test_api_floor_preference_honored_when_the_window_allows_it():
    units = _pool_units()
    allowed = [3, 4, 5, 6, 7, 8, 9, 10]   # the lower half; every floor here holds 2 units
    data = _post(units, {'forty_rules': {'floors_allowed': allowed}})
    assert data['success'] is True, data.get('error')
    floors = {u['unit_id']: u['floor'] for u in units}
    scen = _optimized(data)
    assert scen
    fr = data['project_summary']['forty_rules']
    assert fr['level'] == 'restrict', fr
    for key, sc in scen.items():
        for uid in _forty_ids(sc):
            assert floors[uid] in allowed, f'{key}: {uid} on floor {floors[uid]} at 40% despite the floor preference'
    assert '40% only on floor(s) 3, 4, 5, 6, 7, 8, 9, 10' in fr['summary']


def test_api_floor_preference_too_small_is_honored_partially_not_refused():
    # Floor 19 holds one unit (S-17, ~1.4% of residential SF): far below the
    # 10% floor. The wish becomes "that unit IS 40%, the program fills the
    # rest" - no error, no silent ignore.
    units = _pool_units()
    data = _post(units, {'forty_rules': {'floors_allowed': [19]}})
    assert data['success'] is True, data.get('error')
    fr = data['project_summary']['forty_rules']
    assert fr['level'] == 'force', fr
    assert fr['preferred_units'] == ['S-17']
    scen = _optimized(data)
    assert scen
    for key, sc in scen.items():
        assert 'S-17' in _forty_ids(sc), f'{key}: preferred unit S-17 not at 40%'
        assert len(_forty_ids(sc)) > 1, f'{key}: the program should fill the rest of the 40% band'
    assert 'preferred apartment(s) (floor(s) 19) are 40%' in fr['summary']
    assert any('Built with your 40% rules' in n for n in data['notes'])
    assert not any('could not be honored' in n for n in data['notes'])


def test_api_floor_rule_that_blocks_a_third_skips_floor_spread_honestly():
    units = _pool_units()
    data = _post(units, {'forty_rules': {'floors_allowed': [3, 4, 5, 6, 7, 8, 9, 10]}, 'floor_spread': True})
    assert data['success'] is True, data.get('error')
    fs = data['project_summary']['floor_spread']
    assert fs['requested'] is True and fs['applied'] is False
    assert 'upper (14-19)' in fs['reason']
    assert any('Floor-spread rule skipped' in n and '40% Rules' in n for n in data['notes'])


def test_api_max_per_floor_is_obeyed_in_every_scenario():
    units = _pool_units()   # 25 units on 17 floors -> some floors hold 2 units
    data = _post(units, {'forty_rules': {'max_per_floor': 1}})
    assert data['success'] is True, data.get('error')
    floors = {u['unit_id']: u['floor'] for u in units}
    scen = _optimized(data)
    assert scen
    for key, sc in scen.items():
        per_floor = {}
        for uid in _forty_ids(sc):
            per_floor[floors[uid]] = per_floor.get(floors[uid], 0) + 1
        assert max(per_floor.values()) <= 1, f'{key}: more than one 40% unit on a floor: {per_floor}'


def test_api_rules_stack_with_floor_spread_and_band_picker():
    units = _pool_units()
    data = _post(units, {
        'forty_rules': {'pin_units': ['S-4'], 'max_per_floor': 1},
        'floor_spread': True,
        'allowed_bands': [40, 70, 80, 90, 100],
    })
    assert data['success'] is True, data.get('error')
    scen = _optimized(data)
    assert scen
    for key, sc in scen.items():
        assert 'S-4' in _forty_ids(sc)
        assert sc['floor_spread']['satisfied'] is True, key
        assert 60 not in {int(round(float(u['assigned_ami']) * 100)) for u in sc['assignments']}


def test_api_pin_outside_preferred_floors_still_wins_and_preference_still_applies():
    # Pin S-4 (floor 6) AND prefer floors 9-12. The pin is an order: S-4 is
    # 40% everywhere. The wish still shapes the rest: no other 40% unit lands
    # off floors 9-12 at the restrict level.
    units = _pool_units()
    data = _post(units, {'forty_rules': {'pin_units': ['S-4'], 'floors_allowed': [9, 10, 11, 12]}})
    assert data['success'] is True, data.get('error')
    floors = {u['unit_id']: u['floor'] for u in units}
    fr = data['project_summary']['forty_rules']
    assert fr['level'] in ('restrict', 'force'), fr
    scen = _optimized(data)
    assert scen
    for key, sc in scen.items():
        ids = _forty_ids(sc)
        assert 'S-4' in ids, f'{key}: pinned S-4 not at 40%'
        if fr['level'] == 'restrict':
            others = {u for u in ids if u != 'S-4'}
            assert all(floors[u] in (9, 10, 11, 12) for u in others), f'{key}: 40% off the preferred floors: {others}'
    assert 'pinned at 40%: S-4' in fr['summary']


def test_api_unknown_unit_is_a_clean_error():
    data = _post(_pool_units(), {'forty_rules': {'pin_units': ['NOPE-1']}})
    assert data['success'] is False
    assert 'NOPE-1' in data['error']


def test_api_pin_and_exclude_same_unit_is_a_clean_error():
    data = _post(_pool_units(), {'forty_rules': {'pin_units': ['S-4'], 'exclude_units': ['S-4']}})
    assert data['success'] is False
    assert 'both pinned' in data['error']


def test_api_too_many_pins_is_a_clean_error():
    units = _pool_units()
    # 20 of 25 pool units ~ 20% of residential SF, above the 12.5% + 5-point slide ceiling.
    data = _post(units, {'forty_rules': {'pin_units': [u['unit_id'] for u in units[:20]]}})
    assert data['success'] is False
    assert 'may be at most' in data['error']


def test_api_bedroom_preference_too_small_is_honored_partially():
    units = _pool_units()
    for u in units:
        u['bedrooms'] = 1
    units[0]['bedrooms'] = 3   # a single 3 BR cannot reach the 10% floor alone
    data = _post(units, {'forty_rules': {'bedrooms_allowed': [3]}})
    assert data['success'] is True, data.get('error')
    fr = data['project_summary']['forty_rules']
    assert fr['level'] == 'force'
    for key, sc in _optimized(data).items():
        assert 'S-1' in _forty_ids(sc), f'{key}: the only 3 BR should be 40%'


def test_api_keep_out_leaving_too_little_sf_is_a_clean_error():
    # Keep OUT is an order, not a wish: excluding almost everything is refused.
    units = _pool_units()
    data = _post(units, {'forty_rules': {'exclude_units': [u['unit_id'] for u in units[:22]]}})
    assert data['success'] is False
    assert "Keep OUT" in data['error']
