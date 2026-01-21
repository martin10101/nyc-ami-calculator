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

