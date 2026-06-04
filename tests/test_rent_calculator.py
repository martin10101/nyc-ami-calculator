from pathlib import Path

import pytest

from ami_optix.rent_calculator import (
    HAIRCUT_FACTOR,
    compute_rents_for_assignments,
    load_rent_schedule,
)


def test_rent_allowances_reflect_utility_selection():
    workbook = Path(__file__).resolve().parent.parent / "2025 AMI Rent Calculator Unlocked.xlsx"
    schedule = load_rent_schedule(str(workbook))

    base_components = schedule.rent_components(
        0.6,
        2,
        {"electricity": "na", "cooking": "na", "heat": "na", "hot_water": "na"},
    )
    tenant_electric_components = schedule.rent_components(
        0.6,
        2,
        {"electricity": "tenant_pays", "cooking": "na", "heat": "na", "hot_water": "na"},
    )

    assert tenant_electric_components["allowances"]["electricity"]["amount"] > base_components["allowances"]["electricity"]["amount"]
    assert tenant_electric_components["allowance_total"] > base_components["allowance_total"]
    assert tenant_electric_components["net"] == tenant_electric_components["gross"] - tenant_electric_components["allowance_total"]

    enriched, totals = compute_rents_for_assignments(
        schedule,
        [{"assigned_ami": 0.6, "bedrooms": 2}],
        {"electricity": "tenant_pays", "cooking": "na", "heat": "na", "hot_water": "na"},
    )

    assert enriched[0]["allowance_total"] == tenant_electric_components["allowance_total"]
    assert totals["allowances_monthly"] == tenant_electric_components["allowance_total"]
    assert totals["net_monthly"] == tenant_electric_components["net"]


def test_100_pct_ami_rent_has_haircut_applied():
    """NYC HPD caps collectable rent at 100% AMI to 97% of the headline value.
    rent_components must apply the haircut at exactly 100% AMI so every
    downstream rent calculation (solver objective, display, edge scenarios)
    sees the regulatory-capped number."""
    workbook = Path(__file__).resolve().parent.parent / "2025 AMI Rent Calculator Unlocked.xlsx"
    schedule = load_rent_schedule(str(workbook))
    utilities = {"electricity": "na", "cooking": "na", "heat": "na", "hot_water": "na"}

    components = schedule.rent_components(1.0, 1, utilities)

    assert components["haircut_applied"] is True
    assert components["gross_pre_haircut"] > 0
    assert components["gross"] == pytest.approx(components["gross_pre_haircut"] * HAIRCUT_FACTOR)
    assert components["gross"] < components["gross_pre_haircut"]


def test_no_haircut_at_bands_other_than_100():
    """The haircut applies ONLY at 100% AMI. Bands at 40/60/80/90 must return
    headline rents unchanged. This guards against accidentally scaling all
    rents by 0.97."""
    workbook = Path(__file__).resolve().parent.parent / "2025 AMI Rent Calculator Unlocked.xlsx"
    schedule = load_rent_schedule(str(workbook))
    utilities = {"electricity": "na", "cooking": "na", "heat": "na", "hot_water": "na"}

    for ami_band in [0.4, 0.6, 0.8, 0.9]:
        components = schedule.rent_components(ami_band, 1, utilities)
        assert components["haircut_applied"] is False, f"Haircut wrongly applied at {ami_band*100:.0f}% AMI"
        assert components["gross"] == pytest.approx(components["gross_pre_haircut"]), (
            f"gross != gross_pre_haircut at {ami_band*100:.0f}% AMI"
        )


def test_compute_rents_for_assignments_reflects_haircut():
    """Mixed-AMI rent roll: units at 100% must have 3% lower rent than they
    would without the haircut; units at other bands are unaffected. This
    proves the haircut flows through compute_rents_for_assignments and into
    the rent_totals.net_monthly used by the rent-first re-ranking."""
    workbook = Path(__file__).resolve().parent.parent / "2025 AMI Rent Calculator Unlocked.xlsx"
    schedule = load_rent_schedule(str(workbook))
    utilities = {"electricity": "na", "cooking": "na", "heat": "na", "hot_water": "na"}

    headline_60 = schedule.rent_components(0.6, 1, utilities)["gross_pre_haircut"]
    headline_100 = schedule.rent_components(1.0, 1, utilities)["gross_pre_haircut"]

    enriched, totals = compute_rents_for_assignments(
        schedule,
        [
            {"assigned_ami": 0.6, "bedrooms": 1},
            {"assigned_ami": 1.0, "bedrooms": 1},
        ],
        utilities,
    )

    assert enriched[0]["gross_rent"] == pytest.approx(headline_60)
    assert enriched[1]["gross_rent"] == pytest.approx(headline_100 * HAIRCUT_FACTOR)

    expected_total = headline_60 + (headline_100 * HAIRCUT_FACTOR)
    assert totals["gross_monthly"] == pytest.approx(expected_total)
