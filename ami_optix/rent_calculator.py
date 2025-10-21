"""
Utility helpers for loading the annual AMI rent workbook and deriving
monthly/annual net rents based on utility selections.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Tuple
import os
import math

import pandas as pd
from pathlib import Path
from openpyxl import load_workbook

BEDROOM_LABELS = ["studio", "1 BR", "2 BR", "3 BR", "4 BR", "5 BR"]

COOKING_OPTIONS = {
    "electric": "Electric Stove",
    "gas": "Gas Stove",
    "na": "N/A or owner pays",
}

HEAT_OPTIONS = {
    "electric_ccashp": "Electric Heat - Cold Climate Air Source Heat Pump (ccASHP)1",
    "electric_other": "Electric Heat - Other2",
    "gas": "Gas Heat",
    "oil": "Oil Heat",
    "na": "N/A or owner pays",
}

HOT_WATER_OPTIONS = {
    "electric_heat_pump": "Electric Hot Water - Heat Pump",
    "electric_other": "Electric Hot Water - Other",
    "gas": "Gas Hot Water",
    "oil": "Oil Hot Water",
    "na": "N/A or owner pays",
}


@dataclass
class RentSchedule:
    gross_rents: Dict[Tuple[float, str], float]
    allowances: Dict[str, Dict[str, Dict[str, float]]]

    def rent_for(self, ami_percent: float, bedrooms: float, selections: Dict[str, str]) -> float:
        bedroom_label = _normalize_bedroom_label(bedrooms)
        gross = self._gross_rents_lookup(ami_percent, bedroom_label)
        allowance_total = 0.0
        for category, human_choice in selections.items():
            option_label = _resolve_option_label(category, human_choice)
            allowance_total += self._allowance_lookup(category, option_label, bedroom_label)
        return max(gross - allowance_total, 0.0)

    def _gross_rents_lookup(self, ami_percent: float, bedroom_label: str) -> float:
        key = (round(ami_percent, 4), bedroom_label)
        if key not in self.gross_rents:
            raise ValueError(f"Rent table missing entry for {ami_percent*100:.0f}% AMI / {bedroom_label}.")
        return float(self.gross_rents[key])

    def _allowance_lookup(self, category: str, option: str, bedroom_label: str) -> float:
        cat = self.allowances.get(category)
        if not cat or option not in cat:
            return 0.0
        bedroom_allowances = cat[option]
        return float(bedroom_allowances.get(bedroom_label, 0.0))


def _pandas_engine_for(path: str) -> str | None:
    ext = os.path.splitext(path.lower())[1]
    if ext == ".xlsb":
        return "pyxlsb"
    return None


def load_rent_schedule(workbook_path: str) -> RentSchedule:
    if not os.path.exists(workbook_path):
        raise FileNotFoundError(f"Rent calculator workbook not found: {workbook_path}")

    engine = _pandas_engine_for(workbook_path)
    sheet = pd.read_excel(workbook_path, sheet_name="AMI & Rent", header=None, engine=engine)

    allowances = _parse_allowances(sheet)
    gross_rents = _parse_ami_rent_table(sheet)

    return RentSchedule(gross_rents=gross_rents, allowances=allowances)


def _parse_allowances(sheet: pd.DataFrame) -> Dict[str, Dict[str, Dict[str, float]]]:
    allowances: Dict[str, Dict[str, Dict[str, float]]] = {}
    current_category = None
    for col_idx in range(sheet.shape[1]):
        header = sheet.iloc[14, col_idx]
        if isinstance(header, str) and header.strip():
            label = header.strip()
            if label not in ("Project Total",):
                current_category = label

        option = sheet.iloc[15, col_idx]
        if current_category and isinstance(option, str) and option.strip() and option.strip().lower() != "select -->>":
            cleaned_option = option.strip()
            allowances.setdefault(current_category, {})
            bedroom_map = {}
            for offset, bedroom in enumerate(BEDROOM_LABELS):
                value = sheet.iloc[17 + offset, col_idx]
                bedroom_map[bedroom] = float(value) if not pd.isna(value) else 0.0
            allowances[current_category][cleaned_option] = bedroom_map
    return allowances


def _parse_ami_rent_table(sheet: pd.DataFrame) -> Dict[Tuple[float, str], float]:
    rents: Dict[Tuple[float, str], float] = {}
    current_ami = None
    for idx in range(sheet.shape[0]):
        cell = sheet.iloc[idx, 2]
        marker = sheet.iloc[idx, 3] if sheet.shape[1] > 3 else None

        if isinstance(cell, (int, float)) and isinstance(marker, str) and marker.strip().lower() == "of ami":
            current_ami = float(cell)
            continue

        if current_ami is None:
            continue

        if isinstance(cell, str):
            label = cell.strip()
            if label in BEDROOM_LABELS:
                gross_rent = sheet.iloc[idx, 6] if sheet.shape[1] > 6 else None
                if pd.isna(gross_rent):
                    raise ValueError(f"Missing gross rent for {current_ami*100:.0f}% AMI / {label}.")
                rents[(round(current_ami, 4), label)] = float(gross_rent)
    return rents


def _normalize_bedroom_label(bedrooms: float) -> str:
    try:
        bedrooms_int = int(round(float(bedrooms)))
    except (TypeError, ValueError):
        bedrooms_int = 0
    if bedrooms_int <= 0:
        return "studio"
    if bedrooms_int >= 5:
        return "5 BR"
    return f"{bedrooms_int} BR"


def _resolve_option_label(category: str, choice_key: str) -> str:
    mappings = {
        "cooking": COOKING_OPTIONS,
        "heat": HEAT_OPTIONS,
        "hot_water": HOT_WATER_OPTIONS,
    }
    if category not in mappings:
        raise ValueError(f"Unknown utility category '{category}'.")
    options = mappings[category]
    if choice_key not in options:
        raise ValueError(f"Unsupported selection '{choice_key}' for category '{category}'.")
    return options[choice_key]


def compute_rents_for_assignments(
    schedule: RentSchedule,
    assignments: List[Dict[str, float]],
    utilities: Dict[str, str],
) -> Tuple[List[Dict[str, float]], float, float]:
    updated_assignments = []
    total_monthly = 0.0
    for unit in assignments:
        ami = float(unit.get("assigned_ami"))
        bedrooms = unit.get("bedrooms", 0)
        monthly = schedule.rent_for(ami, bedrooms, utilities)
        annual = monthly * 12.0
        enriched = dict(unit)
        enriched["monthly_rent"] = round(monthly, 2)
        enriched["annual_rent"] = round(annual, 2)
        updated_assignments.append(enriched)
        total_monthly += monthly
    return updated_assignments, round(total_monthly, 2), round(total_monthly * 12.0, 2)


def save_rent_workbook_with_utilities(
    source_path: str,
    utilities: Dict[str, str],
    destination_path: str,
) -> str:
    """
    Persist utility selections into the rent workbook so Excel shows the same
    allowance choices the user picked in the UI.
    """
    if not source_path or not os.path.exists(source_path):
        raise FileNotFoundError(f"Rent workbook not found: {source_path}")

    source_ext = Path(source_path).suffix.lower()
    if source_ext == ".xlsb":
        raise ValueError("Writing utility selections into .xlsb workbooks is not supported.")

    keep_vba = source_ext == ".xlsm"
    workbook = load_workbook(source_path, keep_vba=keep_vba)
    sheet = workbook["AMI & Rent"]

    mapping = [
        ("cooking", 17, 3, COOKING_OPTIONS),
        ("heat", 17, 5, HEAT_OPTIONS),
        ("hot_water", 17, 10, HOT_WATER_OPTIONS),
    ]

    for key, row, col, options in mapping:
        selection_key = utilities.get(key, "na")
        label = options.get(selection_key, options["na"])
        sheet.cell(row=row, column=col).value = label

    workbook.save(destination_path)
    return destination_path
