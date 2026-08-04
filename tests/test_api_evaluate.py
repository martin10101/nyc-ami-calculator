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


def test_evaluate_mih_option1_enforces_40_band_min_share_floor():
    """MIH Option 1 requires AT LEAST 10% of residential SF at <=40% AMI.

    Payload places 0% of SF at the 40% band, which violates the floor.
    Replaces the prior test that asserted a 10% ceiling (the rule was
    flipped from max to min on 2026-04 per client request).
    """
    client = app.test_client()
    payload = {
        "program": "MIH",
        "mih_option": "Option 1",
        "mih_residential_sf": 1000,
        "utilities": {"electricity": "na", "cooking": "na", "heat": "na", "hot_water": "na"},
        "units": [
            # All units at 60% -> 0% at <=40 band, WAAMI=60.0% (at cap)
            {"unit_id": "1A", "bedrooms": 1, "net_sf": 250, "assigned_ami": 0.60},
            {"unit_id": "1B", "bedrooms": 1, "net_sf": 250, "assigned_ami": 0.60},
            {"unit_id": "1C", "bedrooms": 1, "net_sf": 250, "assigned_ami": 0.60},
            {"unit_id": "1D", "bedrooms": 1, "net_sf": 250, "assigned_ami": 0.60},
        ],
    }
    resp = client.post("/api/evaluate", json=payload)
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["success"] is False
    assert any("below required 10.00%" in err for err in data.get("errors", []))


def test_evaluate_mih_option1_floor_satisfied_passes_share_constraint():
    """Option 1 with exactly 10% at <=40 band passes the share constraint.

    Note: this assignment may still fail other constraints (rent calc, etc.)
    but should NOT raise the share-threshold "below required" error.
    """
    client = app.test_client()
    payload = {
        "program": "MIH",
        "mih_option": "Option 1",
        "mih_residential_sf": 1000,
        "utilities": {"electricity": "na", "cooking": "na", "heat": "na", "hot_water": "na"},
        "units": [
            # 100 SF at 40 (10%), 800 SF at 60, 100 SF at 80; WAAMI=60% at cap.
            {"unit_id": "1A", "bedrooms": 1, "net_sf": 100, "assigned_ami": 0.40},
            {"unit_id": "1B", "bedrooms": 1, "net_sf": 800, "assigned_ami": 0.60},
            {"unit_id": "1C", "bedrooms": 1, "net_sf": 100, "assigned_ami": 0.80},
        ],
    }
    resp = client.post("/api/evaluate", json=payload)
    assert resp.status_code == 200
    data = resp.get_json()
    # Share constraint must NOT be the failure reason if it fails for other reasons.
    for err in (data.get("errors") or []):
        assert "below required 10.00%" not in err, f"Share floor incorrectly violated: {err}"


def test_evaluate_mih_option4_has_no_40_band_floor():
    """MIH Option 4 (Workforce) has NO 40% AMI requirement — client corrected
    the old 'same window as Option 1' rule on 2026-08-04 (ZR 23-154(d)).
    A zero-40s allocation that satisfies the real set-asides (>=5% at <=70,
    >=10% cumulative at <=90, WAAMI <=115%) must pass compliance."""
    client = app.test_client()
    payload = {
        "program": "MIH",
        "mih_option": "Option 4",
        "mih_residential_sf": 1000,
        "mih_max_band_percent": 135,
        "utilities": {"electricity": "na", "cooking": "na", "heat": "na", "hot_water": "na"},
        "units": [
            # 100 SF at 70% = 10% of residential at <=70 AND <=90 (cumulative).
            # 300 SF at 130% -> WAAMI = (100*0.7 + 300*1.3) / 400 = 1.15 (at cap).
            {"unit_id": "1A", "bedrooms": 1, "net_sf": 100, "assigned_ami": 0.70},
            {"unit_id": "1B", "bedrooms": 1, "net_sf": 100, "assigned_ami": 1.30},
            {"unit_id": "1C", "bedrooms": 1, "net_sf": 100, "assigned_ami": 1.30},
            {"unit_id": "1D", "bedrooms": 1, "net_sf": 100, "assigned_ami": 1.30},
        ],
    }
    resp = client.post("/api/evaluate", json=payload)
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["success"] is True, f"zero-40s Workforce allocation must pass: {data.get('errors')}"
    assert not any("<= 40%" in err for err in data.get("errors", []) or [])
