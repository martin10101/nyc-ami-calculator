from app import app


def test_evaluate_rejects_too_many_bands():
    client = app.test_client()
    payload = {
        "program": "UAP",
        "utilities": {"electricity": "na", "cooking": "na", "heat": "na", "hot_water": "na"},
        "units": [
            {"unit_id": "1A", "bedrooms": 1, "net_sf": 500, "assigned_ami": 0.40},
            {"unit_id": "1B", "bedrooms": 1, "net_sf": 500, "assigned_ami": 0.60},
            {"unit_id": "1C", "bedrooms": 1, "net_sf": 500, "assigned_ami": 0.70},
            {"unit_id": "1D", "bedrooms": 1, "net_sf": 500, "assigned_ami": 0.80},
        ],
    }
    resp = client.post("/api/evaluate", json=payload)
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["success"] is False
    assert any("Too many bands" in err for err in data.get("errors", []))


def test_evaluate_mih_requires_j21():
    client = app.test_client()
    payload = {
        "program": "MIH",
        "mih_option": "Option 4",
        "utilities": {"electricity": "na", "cooking": "na", "heat": "na", "hot_water": "na"},
        "units": [{"unit_id": "1A", "bedrooms": 1, "net_sf": 500, "assigned_ami": 0.40}],
    }
    resp = client.post("/api/evaluate", json=payload)
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["success"] is False
    assert any("MIH requires mih_residential_sf" in err for err in data.get("errors", []))


def test_evaluate_echoes_rent_schedule_metadata():
    client = app.test_client()
    payload = {
        "rent_roll_year": 2025,
        "program": "UAP",
        "utilities": {"electricity": "na", "cooking": "na", "heat": "na", "hot_water": "na"},
        "units": [{"unit_id": "1A", "bedrooms": 1, "net_sf": 500, "assigned_ami": 0.40}],
    }
    resp = client.post("/api/evaluate", json=payload)
    assert resp.status_code == 200
    data = resp.get_json()
    assert data.get("rent_roll_year_requested") == 2025
    assert data.get("rent_roll_year_used") == 2025
    assert data.get("calculator_filename")
    assert data.get("rent_schedule_source") in {"request_year", "request_year_fallback"}


def test_evaluate_mih_option1_enforces_40_band_max_share_cap():
    client = app.test_client()
    payload = {
        "program": "MIH",
        "mih_option": "Option 1",
        "mih_residential_sf": 1000,
        "utilities": {"electricity": "na", "cooking": "na", "heat": "na", "hot_water": "na"},
        "units": [
            {"unit_id": "1A", "bedrooms": 1, "net_sf": 250, "assigned_ami": 0.40},
            {"unit_id": "1B", "bedrooms": 1, "net_sf": 250, "assigned_ami": 0.40},
            {"unit_id": "1C", "bedrooms": 1, "net_sf": 250, "assigned_ami": 0.60},
            {"unit_id": "1D", "bedrooms": 1, "net_sf": 250, "assigned_ami": 0.60},
        ],
    }
    resp = client.post("/api/evaluate", json=payload)
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["success"] is False
    assert any("above maximum 10.00%" in err for err in data.get("errors", []))
