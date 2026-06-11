"""project_summary must expose the effective MIH 40% window so Excel can
render the client's "required vs provided" compliance box (fix 2026-06-11)."""

from app import app


def test_mih_project_summary_includes_low_band_window():
    client = app.test_client()
    units = [
        {'unit_id': f'U{i}', 'bedrooms': 1, 'net_sf': 100, 'floor': i, 'balcony': False}
        for i in range(1, 11)
    ]
    resp = client.post('/api/optimize', json={
        'program': 'MIH',
        'mih_option': 'Option 1',
        'mih_residential_sf': 1000,
        'mih_max_band_percent': 135,
        'utilities': {'electricity': 'na', 'cooking': 'na', 'heat': 'na', 'hot_water': 'na'},
        'units': units,
    })
    assert resp.status_code == 200
    data = resp.get_json()
    assert data['success'] is True

    summary = data.get('project_summary') or {}
    assert summary.get('total_building_sf') == 1000

    # Standard MIH window [10%, 12.5%] — may sit higher only if the
    # floor-walk slid it (not the case for this feasible payload).
    assert abs(float(summary.get('mih_low_band_min_share')) - 0.10) < 1e-9
    assert abs(float(summary.get('mih_low_band_max_share')) - 0.125) < 1e-9


def test_uap_project_summary_has_no_mih_window_fields():
    client = app.test_client()
    resp = client.post('/api/optimize', json={
        'program': 'UAP',
        'utilities': {'electricity': 'na', 'cooking': 'na', 'heat': 'na', 'hot_water': 'na'},
        'units': [
            {'unit_id': '1A', 'bedrooms': 1, 'net_sf': 200, 'floor': 1, 'balcony': False},
            {'unit_id': '1B', 'bedrooms': 1, 'net_sf': 200, 'floor': 1, 'balcony': False},
            {'unit_id': '2A', 'bedrooms': 2, 'net_sf': 400, 'floor': 6, 'balcony': True},
            {'unit_id': '2B', 'bedrooms': 2, 'net_sf': 400, 'floor': 6, 'balcony': True},
        ],
    })
    assert resp.status_code == 200
    data = resp.get_json()
    summary = data.get('project_summary') or {}
    assert 'mih_low_band_min_share' not in summary
    assert 'mih_low_band_max_share' not in summary
