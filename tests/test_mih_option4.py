"""MIH Option 4 (Workforce Option) rules — client-confirmed 2026-08-04.

Official rules (ZR 23-154(d), NYC Council MIH page; client confirmed):
  - NO 40% AMI requirement (that window belongs to Option 1)
  - >=5% of residential SF at <=70% AMI
  - >=5% at <=90% AMI (encoded cumulatively: >=10% at <=90 since <=70 counts)
  - weighted average AMI <= 115%
  - bands up to 135% AMI allowed (the 100% hard cap is Option 1 only)
  - at most 4 income bands per scenario

Regression context: before this fix the server applied Option 1's
[10%, 12.5%]-at-40% window AND the 100% band cap to Option 4, which forced
~22 phantom 40% AMI units and cost ~$60K/mo on a real Workforce project
(701 Myrtle, 2026-08).
"""

import pytest

from app import _build_program_config


def _base_config():
    return {
        'developer_preferences': {
            'premium_score_weights': {
                'floor': 0.45, 'net_sf': 0.30, 'bedrooms': 0.15, 'balcony': 0.10,
            }
        },
        'optimization_rules': {},
    }


# ---------------------------------------------------------------------------
# _build_program_config unit tests
# ---------------------------------------------------------------------------

def test_option4_has_no_forty_requirement():
    config = _build_program_config(
        _base_config(), 'MIH', mih_option='Option 4',
        mih_residential_sf=100000, mih_max_band_percent=135,
    )
    rules = config['optimization_rules']
    thresholds = rules['share_thresholds']
    assert not any(int(t.get('band_threshold', 0)) <= 40 for t in thresholds), (
        f"Option 4 must have NO <=40%% AMI requirement; got {thresholds}"
    )
    # Workforce set-asides survive.
    t70 = next(t for t in thresholds if int(t['band_threshold']) == 70)
    t90 = next(t for t in thresholds if int(t['band_threshold']) == 90)
    assert float(t70['min_share']) == pytest.approx(0.05)
    assert float(t90['min_share']) == pytest.approx(0.10)  # cumulative 5+5
    assert rules['waami_cap_percent'] == pytest.approx(115.0)
    assert rules['max_bands_per_scenario'] == 4


def test_option4_allows_bands_to_135():
    config = _build_program_config(
        _base_config(), 'MIH', mih_option='Option 4',
        mih_residential_sf=100000, mih_max_band_percent=135,
    )
    bands = config['optimization_rules']['potential_bands']
    for required in (110, 120, 130, 135):
        assert required in bands, f"Option 4 must allow the {required}%% band; got {bands}"
    assert 50 not in bands


def test_option4_band_cap_tops_out_at_135():
    config = _build_program_config(
        _base_config(), 'MIH', mih_option='Option 4',
        mih_residential_sf=100000, mih_max_band_percent=175,
    )
    bands = config['optimization_rules']['potential_bands']
    assert max(bands) == 135


def test_option4_respects_lower_workbook_cap():
    config = _build_program_config(
        _base_config(), 'MIH', mih_option='Option 4',
        mih_residential_sf=100000, mih_max_band_percent=100,
    )
    bands = config['optimization_rules']['potential_bands']
    assert max(bands) == 100


def test_option1_unchanged_forty_window_and_100_cap():
    config = _build_program_config(
        _base_config(), 'MIH', mih_option='Option 1',
        mih_residential_sf=100000, mih_max_band_percent=135,
    )
    rules = config['optimization_rules']
    t40 = next(t for t in rules['share_thresholds'] if int(t['band_threshold']) == 40)
    assert float(t40['min_share']) == pytest.approx(0.10)
    assert float(t40['max_share']) == pytest.approx(0.125)
    assert rules['waami_cap_percent'] == pytest.approx(60.0)
    # 100 hard cap still applies to Option 1 even when the workbook sends 135.
    assert max(rules['potential_bands']) == 100


# ---------------------------------------------------------------------------
# API-level Option 4 behavior
# ---------------------------------------------------------------------------

def _workforce_units():
    """14 varied units. With residential_sf=15000 the pool is 30.4%% of
    residential, mirroring a real Workforce project's affordable share."""
    units = []
    sf_by_tier = {0: [330, 340, 350], 1: [450, 470, 490, 510, 530], 2: [660, 680, 700, 720, 740, 760]}
    i = 0
    for bedrooms, sfs in sf_by_tier.items():
        for sf in sfs:
            i += 1
            units.append({
                'unit_id': f'W{i}', 'bedrooms': bedrooms, 'net_sf': sf,
                'floor': i, 'balcony': False,
            })
    return units


def _forty_share(scenario, residential_sf):
    sf = sum(
        float(u['net_sf']) for u in (scenario.get('assignments') or [])
        if float(u['assigned_ami']) <= 0.4 + 1e-12
    )
    return sf / residential_sf


def _bands_of(scenario):
    return sorted({
        int(round(float(u['assigned_ami']) * 100))
        for u in (scenario.get('assignments') or [])
    })


def _waami(scenario):
    a = scenario.get('assignments') or []
    t = sum(float(u['net_sf']) for u in a)
    return sum(float(u['net_sf']) * float(u['assigned_ami']) for u in a) / t if t else 0.0


def test_api_option4_no_forced_forties_and_high_bands():
    from app import app

    residential_sf = 15000.0
    client = app.test_client()
    resp = client.post('/api/optimize', json={
        'program': 'MIH',
        'mih_option': 'Option 4',
        'mih_residential_sf': residential_sf,
        'mih_max_band_percent': 135,
        'utilities': {'electricity': 'na', 'cooking': 'na', 'heat': 'na', 'hot_water': 'na'},
        'units': _workforce_units(),
    })
    assert resp.status_code == 200
    data = resp.get_json()
    assert data['success'] is True
    scenarios = {k: v for k, v in (data.get('scenarios') or {}).items()
                 if v and k != 'original' and v.get('assignments')}

    if not scenarios or not any((s.get('rent_totals') or {}).get('net_monthly') for s in scenarios.values()):
        pytest.skip('Rent calculator unavailable in this environment.')

    assert len(scenarios) >= 3, f"Expected >=3 Option 4 scenarios, got {sorted(scenarios)}"

    for key, s in scenarios.items():
        bands = _bands_of(s)
        assert len(bands) <= 4, f"{key} uses {len(bands)} bands (max 4): {bands}"
        assert _waami(s) <= 1.15 + 1e-9, f"{key} WAAMI {_waami(s)} exceeds 115%%"
        # Workforce set-asides (cumulative, of residential SF).
        a = s.get('assignments') or []
        sf70 = sum(float(u['net_sf']) for u in a if float(u['assigned_ami']) <= 0.70 + 1e-9)
        sf90 = sum(float(u['net_sf']) for u in a if float(u['assigned_ami']) <= 0.90 + 1e-9)
        assert sf70 / residential_sf >= 0.05 - 1e-9, f"{key} misses 5%% at <=70"
        assert sf90 / residential_sf >= 0.10 - 1e-9, f"{key} misses 10%% at <=90"

    # The old bug forced >=10%% of residential SF to 40%% AMI in EVERY scenario.
    # With no 40%% requirement, rent-max scenarios should not sink SF to 40.
    assert any(_forty_share(s, residential_sf) < 0.10 - 1e-9 for s in scenarios.values()), (
        "Every scenario still carries >=10%% at 40%% AMI - the Option 1 window is leaking into Option 4"
    )

    # The old bug capped bands at 100. The Workforce average (115) is only
    # reachable with high bands; at least one scenario must use a band > 100.
    assert any(max(_bands_of(s)) > 100 for s in scenarios.values()), (
        "No scenario uses a band above 100%% - the Option 1 hard cap is leaking into Option 4"
    )

    # The 40%%-centric families must not be generated for Option 4.
    for missing in ('low_40_share', 'max_40_share', 'mid_40_share'):
        assert missing not in scenarios, f"{missing} should not exist for Option 4"
    assert not any(k.startswith('fewest_40_units') for k in scenarios), (
        "fewest_40_units family should not exist for Option 4"
    )

    # RECOMMENDED for a no-forty option = the straight rent-max scenario.
    # (The Option 1 "fewest at 40%" picker would otherwise crown a low-rent
    # scenario that merely happens to use 40s as ballast, and Excel's
    # ApplyBestScenario would write that to the MIH page.)
    rk = data.get('recommended_key')
    assert rk in scenarios, f"recommended_key {rk!r} missing from scenarios"
    def _income(s):
        return float((s.get('rent_totals') or {}).get('net_monthly') or 0.0)
    best_income = max(_income(s) for s in scenarios.values())
    assert _income(scenarios[rk]) == pytest.approx(best_income), (
        f"recommended {rk} earns {_income(scenarios[rk])}, best is {best_income}"
    )


def test_api_option1_regression_still_has_forty_window():
    """Option 1 behavior must be byte-for-byte the same concept as before:
    the 40%% window, the 100 cap, and the fewest-40 family all present."""
    from app import app

    client = app.test_client()
    resp = client.post('/api/optimize', json={
        'program': 'MIH',
        'mih_option': 'Option 1',
        'mih_residential_sf': 15000,
        'mih_max_band_percent': 135,
        'utilities': {'electricity': 'na', 'cooking': 'na', 'heat': 'na', 'hot_water': 'na'},
        'units': _workforce_units(),
    })
    assert resp.status_code == 200
    data = resp.get_json()
    assert data['success'] is True
    scenarios = {k: v for k, v in (data.get('scenarios') or {}).items()
                 if v and k != 'original' and v.get('assignments')}

    if not scenarios or not any((s.get('rent_totals') or {}).get('net_monthly') for s in scenarios.values()):
        pytest.skip('Rent calculator unavailable in this environment.')

    residential_sf = 15000.0
    for key, s in scenarios.items():
        assert max(_bands_of(s)) <= 100, f"Option 1 {key} uses a band above 100"
        assert _forty_share(s, residential_sf) >= 0.10 - 1e-9, f"Option 1 {key} below the 40%% floor"
        assert _waami(s) <= 0.60 + 1e-9, f"Option 1 {key} WAAMI over 60%%"
