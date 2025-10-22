from pathlib import Path

from ami_optix.rent_calculator import (
    compute_rents_for_assignments,
    load_rent_schedule,
)


def test_rent_allowances_reflect_utility_selection():
    workbook = Path(__file__).resolve().parent.parent / "2025 AMI Rent Calculator Unlocked.xlsx"
    schedule = load_rent_schedule(str(workbook))

    base_components = schedule.rent_components(
        0.6,
        2,
        {"cooking": "na", "heat": "na", "hot_water": "na"},
    )
    electric_components = schedule.rent_components(
        0.6,
        2,
        {"cooking": "electric", "heat": "na", "hot_water": "na"},
    )

    assert electric_components["allowances"]["cooking"]["amount"] > base_components["allowances"]["cooking"]["amount"]
    assert electric_components["allowance_total"] > base_components["allowance_total"]
    assert electric_components["net"] == electric_components["gross"] - electric_components["allowance_total"]

    enriched, totals = compute_rents_for_assignments(
        schedule,
        [{"assigned_ami": 0.6, "bedrooms": 2}],
        {"cooking": "electric", "heat": "na", "hot_water": "na"},
    )

    assert enriched[0]["allowance_total"] == electric_components["allowance_total"]
    assert totals["allowances_monthly"] == electric_components["allowance_total"]
    assert totals["net_monthly"] == electric_components["net"]
