"""
Helpers to interpret per-project override payloads and convert them into
solver-friendly constraint structures.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Any, Optional, Iterable


def _clean_band_value(value: Any) -> Optional[int]:
    try:
        band = int(round(float(value)))
        if band <= 0:
            return None
        return band
    except (TypeError, ValueError):
        return None


def _clean_unit_id(value: Any) -> Optional[str]:
    if value is None:
        return None
    text = str(value).strip()
    return text if text else None


def _expand_floor_rule(rule: Dict[str, Any]) -> Iterable[int]:
    if "floors" in rule and isinstance(rule["floors"], list):
        for val in rule["floors"]:
            try:
                yield int(val)
            except (TypeError, ValueError):
                continue
    else:
        min_floor = rule.get("minFloor")
        max_floor = rule.get("maxFloor")
        if min_floor is None and max_floor is None:
            return
        try:
            min_floor = int(min_floor)
            max_floor = int(max_floor if max_floor is not None else min_floor)
        except (TypeError, ValueError):
            return
        if min_floor > max_floor:
            min_floor, max_floor = max_floor, min_floor
        for floor in range(min_floor, max_floor + 1):
            yield floor


@dataclass
class ProjectOverrides:
    band_whitelist: Optional[List[int]] = None
    fixed_units: Dict[str, List[int]] = field(default_factory=dict)
    floor_minimums: Dict[int, int] = field(default_factory=dict)
    premium_weights: Optional[Dict[str, float]] = None
    notes: List[str] = field(default_factory=list)

    @classmethod
    def from_dict(cls, payload: Optional[Dict[str, Any]]) -> "ProjectOverrides":
        if not payload:
            return cls()

        overrides = cls()

        whitelist = payload.get("bandWhitelist")
        if isinstance(whitelist, list):
            cleaned = [_clean_band_value(b) for b in whitelist]
            overrides.band_whitelist = sorted({b for b in cleaned if b})

        fixed_units = payload.get("fixedUnits")
        if isinstance(fixed_units, list):
            for item in fixed_units:
                unit_id = _clean_unit_id(item.get("unitId"))
                if not unit_id:
                    continue
                bands = item.get("bands")
                band = item.get("band")
                collected: List[int] = []
                if isinstance(bands, list):
                    collected.extend([b for b in (_clean_band_value(v) for v in bands) if b])
                if band is not None:
                    cleaned_band = _clean_band_value(band)
                    if cleaned_band:
                        collected.append(cleaned_band)
                if collected:
                    overrides.fixed_units[unit_id] = sorted({int(b) for b in collected})

        floor_rules = payload.get("floorMinimums")
        if isinstance(floor_rules, list):
            for rule in floor_rules:
                min_band = _clean_band_value(rule.get("minBand"))
                if not min_band:
                    continue
                for floor in _expand_floor_rule(rule):
                    overrides.floor_minimums[int(floor)] = min_band

        premium = payload.get("premiumWeights")
        if isinstance(premium, dict):
            cleaned_weights = {}
            for key, value in premium.items():
                try:
                    cleaned_weights[str(key)] = float(value)
                except (TypeError, ValueError):
                    continue
            if cleaned_weights:
                overrides.premium_weights = cleaned_weights

        notes = payload.get("notes")
        if isinstance(notes, list):
            overrides.notes = [str(note) for note in notes if note]

        return overrides

    def to_solver_payload(self, df_affordable) -> Dict[str, Any]:
        unit_band_rules: Dict[int, List[int]] = {}
        if self.fixed_units:
            unit_id_series = df_affordable["unit_id"].astype(str).str.strip()
            for idx, unit_id in enumerate(unit_id_series):
                if unit_id in self.fixed_units:
                    unit_band_rules[idx] = self.fixed_units[unit_id]

        unit_min_band: Dict[int, int] = {}
        if self.floor_minimums and "floor" in df_affordable.columns:
            for idx, floor_value in enumerate(df_affordable["floor"]):
                try:
                    floor_int = int(round(float(floor_value)))
                except (TypeError, ValueError):
                    continue
                if floor_int in self.floor_minimums:
                    unit_min_band[idx] = self.floor_minimums[floor_int]

        return {
            "band_whitelist": self.band_whitelist,
            "unit_band_rules": unit_band_rules,
            "unit_min_band": unit_min_band,
            "premium_weights": self.premium_weights,
            "notes": self.notes,
        }
