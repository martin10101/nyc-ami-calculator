"""Band picker (ribbon "AMI Bands" menu) - client feedback 2026-09-02, Fix 4.

Server side = lock 2 of 3: the optional ``allowed_bands`` payload field can
only NARROW the program's candidate bands, never widen them past the option
caps (100 for MIH Option 1, 135 for Option 4). An unusable selection is a
clean, plain-English error - never a silent run with different bands.

Absent field == byte-identical legacy behavior (identity-tested).
"""
import copy

import pytest

from app import _build_program_config, _normalize_allowed_bands


def _base_config():
    return {
        'developer_preferences': {
            'premium_score_weights': {
                'floor': 0.45, 'net_sf': 0.30, 'bedrooms': 0.15, 'balcony': 0.10,
            }
        },
        'optimization_rules': {
            'potential_bands': [40, 60, 70, 80, 90, 100],
            'deep_affordability_min_share': 0.20,
            'deep_affordability_max_share': 0.21,
        },
    }


def _bands(config):
    return list(config['optimization_rules']['potential_bands'])


# ---------------------------------------------------------------------------
# Payload normalization
# ---------------------------------------------------------------------------

def test_normalize_accepts_percents_fractions_and_strings():
    assert _normalize_allowed_bands([40, 0.7, "80", 100.0]) == [40, 70, 80, 100]


def test_normalize_absent_or_junk_means_default():
    assert _normalize_allowed_bands(None) is None
    assert _normalize_allowed_bands([]) is None
    assert _normalize_allowed_bands("40,70") is None
    assert _normalize_allowed_bands(["x", None, -1, 0]) is None


# ---------------------------------------------------------------------------
# Identity: no field -> exactly today's config
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("program,opt", [("UAP", None), ("MIH", "Option 1"), ("MIH", "Option 4")])
def test_absent_field_is_identity(program, opt):
    legacy = _build_program_config(_base_config(), program, mih_option=opt, mih_residential_sf=100000)
    with_none = _build_program_config(_base_config(), program, mih_option=opt, mih_residential_sf=100000, allowed_bands=None)
    a = copy.deepcopy(legacy['optimization_rules'])
    b = copy.deepcopy(with_none['optimization_rules'])
    a.pop('allowed_bands_source', None)
    b.pop('allowed_bands_source', None)
    assert a == b
    assert with_none['optimization_rules']['allowed_bands_source'] == 'default'


@pytest.mark.parametrize("program,opt,full", [
    ("UAP", None, [40, 60, 70, 80, 90, 100]),
    ("MIH", "Option 1", [40, 60, 70, 80, 90, 100]),
    ("MIH", "Option 4", [40, 60, 70, 80, 90, 100, 110, 120, 130, 135]),
])
def test_selecting_every_band_equals_default(program, opt, full):
    legacy = _build_program_config(_base_config(), program, mih_option=opt, mih_residential_sf=100000)
    picked = _build_program_config(_base_config(), program, mih_option=opt, mih_residential_sf=100000, allowed_bands=full)
    assert _bands(picked) == _bands(legacy)
    assert picked['optimization_rules']['allowed_bands_source'] == 'picker'


# ---------------------------------------------------------------------------
# Narrowing
# ---------------------------------------------------------------------------

def test_option1_can_exclude_60():
    cfg = _build_program_config(_base_config(), 'MIH', mih_option='Option 1', mih_residential_sf=100000,
                                allowed_bands=[40, 70, 80, 90, 100])
    assert _bands(cfg) == [40, 70, 80, 90, 100]
    assert 60 not in _bands(cfg)


def test_uap_can_exclude_60():
    cfg = _build_program_config(_base_config(), 'UAP', allowed_bands=[40, 70, 80])
    assert _bands(cfg) == [40, 70, 80]


def test_option1_ignores_bands_above_client_cap():
    # 120 is never legal for Option 1 (client cap 100, 2026-05-18): the picker
    # cannot widen; the band is dropped and reported, not honored.
    cfg = _build_program_config(_base_config(), 'MIH', mih_option='Option 1', mih_residential_sf=100000,
                                allowed_bands=[40, 70, 120])
    assert _bands(cfg) == [40, 70]
    assert cfg['optimization_rules']['allowed_bands_ignored'] == [120]


def test_option4_can_drop_40_and_keep_high_bands():
    cfg = _build_program_config(_base_config(), 'MIH', mih_option='Option 4', mih_residential_sf=100000,
                                allowed_bands=[70, 90, 110, 135])
    assert _bands(cfg) == [70, 90, 110, 135]


def test_lower_workbook_cap_still_applies_under_picker():
    cfg = _build_program_config(_base_config(), 'MIH', mih_option='Option 1', mih_residential_sf=100000,
                                mih_max_band_percent=80, allowed_bands=[40, 60, 70, 80, 90, 100])
    assert _bands(cfg) == [40, 60, 70, 80]


# ---------------------------------------------------------------------------
# Refusals (clean errors, never a silent different run)
# ---------------------------------------------------------------------------

def test_option1_cannot_drop_40():
    with pytest.raises(ValueError, match=r"40% AMI is required for MIH Option 1"):
        _build_program_config(_base_config(), 'MIH', mih_option='Option 1', mih_residential_sf=100000,
                              allowed_bands=[60, 70, 80])


def test_uap_cannot_drop_40():
    with pytest.raises(ValueError, match=r"40% AMI is required for UAP"):
        _build_program_config(_base_config(), 'UAP', allowed_bands=[60, 70, 80])


def test_option4_needs_a_band_at_or_below_70():
    with pytest.raises(ValueError, match=r"at or below 70% AMI"):
        _build_program_config(_base_config(), 'MIH', mih_option='Option 4', mih_residential_sf=100000,
                              allowed_bands=[80, 90, 100, 110])


def test_single_band_selection_rejected():
    with pytest.raises(ValueError, match=r"fewer than 2 usable bands"):
        _build_program_config(_base_config(), 'MIH', mih_option='Option 1', mih_residential_sf=100000,
                              allowed_bands=[40])


def test_selection_entirely_above_cap_rejected():
    with pytest.raises(ValueError):
        _build_program_config(_base_config(), 'MIH', mih_option='Option 1', mih_residential_sf=100000,
                              allowed_bands=[110, 120])


# ---------------------------------------------------------------------------
# End-to-end through /api/optimize (Building D shape: 25-unit pool)
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


def _post(extra):
    from app import app
    client = app.test_client()
    units = _pool_units()
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


def _bands_used(data):
    used = set()
    for key, sc in (data.get('scenarios') or {}).items():
        if key == 'original' or not sc:
            continue
        for u in sc.get('assignments') or []:
            used.add(int(round(float(u['assigned_ami']) * 100)))
    return used


def test_api_option1_excluding_60_never_uses_60():
    data = _post({'allowed_bands': [40, 70, 80, 90, 100]})
    assert data['success'] is True
    used = _bands_used(data)
    assert used, 'expected optimized scenarios'
    assert 60 not in used, f'60% band leaked into scenarios: {sorted(used)}'
    assert 40 in used
    br = data['project_summary']['band_rules']
    assert br['source'] == 'picker'
    assert br['allowed_bands'] == [40, 70, 80, 90, 100]
    assert any('Built with your band rules' in n for n in data.get('notes') or [])


def test_api_unusable_selection_is_a_clean_error():
    data = _post({'allowed_bands': [60, 70]})
    assert data['success'] is False
    assert '40% AMI is required' in data['error']


def test_api_absent_field_echoes_default_and_offers_all_bands():
    data = _post({})
    assert data['success'] is True
    br = data['project_summary']['band_rules']
    assert br['source'] == 'default'
    assert br['allowed_bands'] == [40, 60, 70, 80, 90, 100]
    assert not any('Built with your band rules' in n for n in data.get('notes') or [])
