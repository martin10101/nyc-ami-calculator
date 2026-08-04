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
        # 'original' is the client's input snapshot, not a solver-generated
        # scenario — it's intentionally exempt from compliance checks so users
        # can see what they had even if it violates the rules.
        if name == "original" or (scenario or {}).get("tier") == "reference":
            continue
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


def test_optimize_mih_option4_has_no_forced_40_band_floor():
    """MIH Option 4 (Workforce) has NO 40% AMI requirement — client corrected
    the old 'same window as Option 1' rule on 2026-08-04 (ZR 23-154(d)).
    The optimizer must not force 10% of residential SF into 40% AMI, and the
    40%-centric scenario families must not be generated."""
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
    solver = {
        name: s for name, s in data["scenarios"].items()
        if s and name != "original" and (s or {}).get("tier") != "reference"
        and s.get("assignments")
    }
    assert solver, "expected at least one solver scenario"

    # The old bug forced >=10% of residential SF into 40% AMI in EVERY
    # scenario. Rent-max with no 40% requirement should leave it unforced.
    def _low_sf(s):
        return sum(
            float(a.get("net_sf") or 0.0) for a in (s.get("assignments") or [])
            if float(a.get("assigned_ami") or 0.0) <= 0.4 + 1e-12
        )

    assert any(_low_sf(s) < 0.10 * residential_sf - 1e-9 for s in solver.values()), (
        "Every Option 4 scenario still carries >=10% at 40% AMI — "
        "the Option 1 window is leaking into Option 4"
    )

    # 40%-centric families are Option 1 concepts; they must not appear.
    for name in solver:
        assert not name.startswith("fewest_40_units"), name
        assert name not in ("low_40_share", "max_40_share", "mid_40_share"), name


def test_optimize_mih_band_caps_option1_100_option4_135():
    """MIH band caps, per option.
    - Option 1 (client decision 2026-05-18): hard-cap at 100% AMI even if the
      workbook sends 135. Unchanged.
    - Option 4 / Workforce (client correction 2026-08-04): bands up to 135%
      are integral to the option (avg 115% is unreachable without them) —
      the 100 cap must NOT apply.
    """
    client = app.test_client()
    units = [
        {"unit_id": f"U{i}", "bedrooms": 1, "net_sf": 100, "floor": i, "balcony": False}
        for i in range(1, 11)
    ]

    def _run(option):
        payload = {
            "program": "MIH",
            "mih_option": option,
            "mih_residential_sf": 1000,
            "mih_max_band_percent": 135,
            "utilities": {"electricity": "na", "cooking": "na", "heat": "na", "hot_water": "na"},
            "units": units,
        }
        resp = client.post("/api/optimize", json=payload)
        assert resp.status_code == 200, f"{option}: HTTP {resp.status_code}"
        data = resp.get_json()
        assert data.get("success") is True, f"{option}: {data}"
        return {
            name: s for name, s in (data.get("scenarios") or {}).items()
            if s and name != "original" and s.get("assignments")
        }

    # Option 1: never above 100%.
    for name, scenario in _run("Option 1").items():
        for a in scenario.get("assignments") or []:
            ami = float(a.get("assigned_ami") or 0.0)
            assert ami <= 1.00 + 1e-9, (
                f"Option 1 scenario '{name}' contains assigned_ami={ami} "
                f"(unit {a.get('unit_id')}) - exceeds 100% cap"
            )

    # Option 4: bands above 100% must be available AND used (rent-max under a
    # 115% average cannot land entirely at <=100 for this all-equal pool).
    opt4 = _run("Option 4")
    assert any(
        float(a.get("assigned_ami") or 0.0) > 1.00 + 1e-9
        for s in opt4.values() for a in (s.get("assignments") or [])
    ), "No Option 4 scenario uses a band above 100% - the Option 1 cap is leaking into Option 4"

