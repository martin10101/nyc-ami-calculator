from app import app


def test_optimize_accepts_project_overrides_and_compare_baseline():
    client = app.test_client()
    payload = {
        "program": "UAP",
        "compare_baseline": True,
        "project_overrides": {
            "premiumWeights": {"floor": 0.8, "net_sf": 0.1, "bedrooms": 0.05, "balcony": 0.05},
            "notes": ["test override payload"],
        },
        "utilities": {"electricity": "na", "cooking": "na", "heat": "na", "hot_water": "na"},
        "units": [
            {"unit_id": "L1", "bedrooms": 1, "net_sf": 200, "floor": 1, "balcony": False},
            {"unit_id": "L2", "bedrooms": 1, "net_sf": 200, "floor": 1, "balcony": False},
            {"unit_id": "H1", "bedrooms": 2, "net_sf": 400, "floor": 6, "balcony": True},
            {"unit_id": "H2", "bedrooms": 2, "net_sf": 400, "floor": 6, "balcony": True},
            {"unit_id": "M1", "bedrooms": 2, "net_sf": 400, "floor": 3, "balcony": False},
            {"unit_id": "M2", "bedrooms": 2, "net_sf": 400, "floor": 3, "balcony": False},
        ],
    }
    resp = client.post("/api/optimize", json=payload)
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["success"] is True
    assert "scenarios" in data
    assert "absolute_best" in data["scenarios"]
    assert "learning" in data
    assert data["learning"]["compare_baseline"] is True
    assert "baseline" in data["learning"]
    assert "learned" in data["learning"]
    assert "diff" in data["learning"]


def test_optimize_without_overrides_does_not_require_learning_fields():
    client = app.test_client()
    payload = {
        "program": "UAP",
        "utilities": {"electricity": "na", "cooking": "na", "heat": "na", "hot_water": "na"},
        "units": [
            {"unit_id": "1A", "bedrooms": 1, "net_sf": 200, "floor": 1, "balcony": False},
            {"unit_id": "1B", "bedrooms": 1, "net_sf": 200, "floor": 1, "balcony": False},
            {"unit_id": "2A", "bedrooms": 2, "net_sf": 400, "floor": 6, "balcony": True},
            {"unit_id": "2B", "bedrooms": 2, "net_sf": 400, "floor": 6, "balcony": True},
            {"unit_id": "3A", "bedrooms": 2, "net_sf": 400, "floor": 3, "balcony": False},
            {"unit_id": "3B", "bedrooms": 2, "net_sf": 400, "floor": 3, "balcony": False},
        ],
    }
    resp = client.post("/api/optimize", json=payload)
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["success"] is True
    assert "scenarios" in data
    assert "learning" not in data


def test_optimize_mih_option1_meets_40_band_min_share_floor():
    """MIH Option 1: every returned scenario must have AT LEAST 10% of
    residential SF at <=40% AMI (rule flipped from max ceiling to min floor
    on 2026-04 per client request).
    """
    client = app.test_client()
    units = []
    for i in range(1, 11):
        units.append({"unit_id": f"U{i}", "bedrooms": 1, "net_sf": 100, "floor": i, "balcony": False})

    payload = {
        "program": "MIH",
        "mih_option": "Option 1",
        "mih_residential_sf": 1000,
        "mih_max_band_percent": 135,
        "utilities": {"electricity": "na", "cooking": "na", "heat": "na", "hot_water": "na"},
        "units": units,
    }

    resp = client.post("/api/optimize", json=payload)
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["success"] is True
    assert "scenarios" in data and data["scenarios"]

    residential_sf = 1000.0
    required = 0.10 * residential_sf
    # Ceiling at 12.5% (added 2026-04 to keep 40 band close to 10%);
    # the floor-walking loop may slide the window up to a max ceiling
    # of 15% + 2.5% = 17.5% if the strict window is infeasible.
    walked_ceiling = 0.175 * residential_sf
    for name, scenario in data["scenarios"].items():
        assignments = scenario.get("assignments") or []
        low_sf = 0.0
        for a in assignments:
            sf = float(a.get("net_sf") or 0.0)
            ami = float(a.get("assigned_ami") or 0.0)
            if ami <= 0.4 + 1e-12:
                low_sf += sf
        assert low_sf + 1e-9 >= required, (
            f"Scenario '{name}' has {low_sf} SF at <=40 band; "
            f"required >= {required} (10% of residential SF)"
        )
        assert low_sf <= walked_ceiling + 1e-9, (
            f"Scenario '{name}' has {low_sf} SF at <=40 band; "
            f"exceeds {walked_ceiling} (17.5% walked-up ceiling)"
        )


def test_optimize_mih_option4_meets_40_band_min_share_floor():
    """MIH Option 4 also requires AT LEAST 10% at <=40 AMI band (added 2026-04)."""
    client = app.test_client()
    units = []
    for i in range(1, 11):
        units.append({"unit_id": f"U{i}", "bedrooms": 1, "net_sf": 100, "floor": i, "balcony": False})

    payload = {
        "program": "MIH",
        "mih_option": "Option 4",
        "mih_residential_sf": 1000,
        "mih_max_band_percent": 135,
        "utilities": {"electricity": "na", "cooking": "na", "heat": "na", "hot_water": "na"},
        "units": units,
    }

    resp = client.post("/api/optimize", json=payload)
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["success"] is True
    assert "scenarios" in data and data["scenarios"]

    residential_sf = 1000.0
    required = 0.10 * residential_sf
    walked_ceiling = 0.175 * residential_sf
    for name, scenario in data["scenarios"].items():
        assignments = scenario.get("assignments") or []
        low_sf = 0.0
        for a in assignments:
            sf = float(a.get("net_sf") or 0.0)
            ami = float(a.get("assigned_ami") or 0.0)
            if ami <= 0.4 + 1e-12:
                low_sf += sf
        assert low_sf + 1e-9 >= required, (
            f"Option 4 scenario '{name}' has {low_sf} SF at <=40 band; "
            f"required >= {required} (10% of residential SF)"
        )
        assert low_sf <= walked_ceiling + 1e-9, (
            f"Option 4 scenario '{name}' has {low_sf} SF at <=40 band; "
            f"exceeds {walked_ceiling} (17.5% walked-up ceiling)"
        )


def test_optimize_mih_caps_bands_at_100_regardless_of_payload():
    """MIH band cap. Client decision 2026-05-18: even if the client workbook
    still sends mih_max_band_percent >= 110 (Option 1 OR Option 4), the server
    must hard-cap returned scenarios at 100% AMI. No assignment in any returned
    scenario may have assigned_ami > 1.00.
    """
    client = app.test_client()
    units = [
        {"unit_id": f"U{i}", "bedrooms": 1, "net_sf": 100, "floor": i, "balcony": False}
        for i in range(1, 11)
    ]
    for option in ("Option 1", "Option 4"):
        payload = {
            "program": "MIH",
            "mih_option": option,
            "mih_residential_sf": 1000,
            "mih_max_band_percent": 135,  # Old default; server should cap to 100.
            "utilities": {"electricity": "na", "cooking": "na", "heat": "na", "hot_water": "na"},
            "units": units,
        }
        resp = client.post("/api/optimize", json=payload)
        assert resp.status_code == 200, f"{option}: HTTP {resp.status_code}"
        data = resp.get_json()
        assert data.get("success") is True, f"{option}: {data}"
        scenarios = data.get("scenarios") or {}
        for name, scenario in scenarios.items():
            for a in (scenario.get("assignments") or []):
                ami = float(a.get("assigned_ami") or 0.0)
                assert ami <= 1.00 + 1e-9, (
                    f"{option} scenario '{name}' contains assigned_ami={ami} "
                    f"(unit {a.get('unit_id')}) - exceeds 100% cap"
                )

