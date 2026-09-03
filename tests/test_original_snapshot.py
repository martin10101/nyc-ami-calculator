"""Original Scenario must reflect the user's true input (units[].original_ami
snapshot) when the add-in provides it, and fall back to the live AMI column
(client_ami) for older add-ins that don't send the field.

Context: after ApplyBestScenario writes a scenario into the workbook's AMI
column, a re-run used to present that previous OUTPUT as "YOUR ORIGINAL
INPUT" (Building D client report, 2026-09-02).
"""
import pytest


def _pool_units(with_snapshot):
    """25-unit affordable pool. client_ami simulates a previously APPLIED
    scenario (40/60/90); original_ami is what the user actually typed
    (40/70/80 - Building D's shape: no 60% anywhere)."""
    applied = ([0.4] * 8) + ([0.6] * 9) + ([0.9] * 8)
    typed = ([0.4] * 8) + ([0.7] * 11) + ([0.8] * 6)
    units = []
    for i in range(25):
        u = {
            'unit_id': f'S-{i+1}',
            'bedrooms': [0, 1, 1, 2, 2][i % 5],
            'net_sf': 380.0 + (i * 23) % 420,
            'floor': 3 + (i % 17),
            'balcony': i % 3 == 0,
            'client_ami': applied[i],
        }
        if with_snapshot:
            u['original_ami'] = typed[i]
        units.append(u)
    return units


def _run(units):
    from app import app
    client = app.test_client()
    pool_sf = sum(u['net_sf'] for u in units)
    resp = client.post('/api/optimize', json={
        'program': 'MIH',
        'mih_option': 'Option 1',
        'mih_residential_sf': pool_sf / 0.254,
        'utilities': {'electricity': 'na', 'cooking': 'na', 'heat': 'na', 'hot_water': 'na'},
        'units': units,
    })
    assert resp.status_code == 200
    data = resp.get_json()
    assert data['success'] is True
    return data


def _original_ami_by_unit(data):
    orig = (data.get('scenarios') or {}).get('original')
    if not orig or not orig.get('assignments'):
        pytest.skip('Original scenario unavailable (no rent calculator in this environment).')
    return {u['unit_id']: round(float(u['assigned_ami']), 4) for u in orig['assignments']}


def test_original_uses_snapshot_when_provided():
    units = _pool_units(with_snapshot=True)
    got = _original_ami_by_unit(_run(units))
    want = {u['unit_id']: round(u['original_ami'], 4) for u in units}
    assert got == want, 'Original Scenario must mirror original_ami (the snapshot), not client_ami'
    # The snapshot has no 60% band anywhere - the applied values do.
    assert 0.6 not in got.values()


def test_original_falls_back_to_client_ami_without_snapshot():
    units = _pool_units(with_snapshot=False)
    got = _original_ami_by_unit(_run(units))
    want = {u['unit_id']: round(u['client_ami'], 4) for u in units}
    assert got == want, 'Legacy add-ins (no original_ami field) must behave exactly as before'
