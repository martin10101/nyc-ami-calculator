import os
import io
from typing import List, Dict, Tuple

import numpy as np
import pandas as pd
from pulp import (
    LpProblem,
    LpMaximize,
    LpVariable,
    lpSum,
    LpStatus,
    value,
    PULP_CBC_CMD,
)

# Optional: LLM explanations in the Excel output (safe to leave unset)
try:
    import google.generativeai as genai

    genai.configure(api_key=os.getenv("GEMINI_API_KEY"))
except ImportError:
    genai = None

DEFAULT_TIMELIMIT = int(os.getenv("MILP_TIMELIMIT_SEC", "30"))
ALLOWED_BANDS = [0.40, 0.60, 0.70, 0.80, 0.90, 1.00]

# Fuzzy header matching -> canonical column names
FLEX_HEADERS = {
    "NET SF": {"net sf", "netsf", "sqft", "area", "square feet"},
    "AMI": {"ami", "aff ami", "assigned_ami"},
    "FLOOR": {"floor", "story", "level"},
    "APT": {"apt", "unit", "apartment"},
    "BED": {"bed", "beds", "bedroom", "br"},
    "BALCONY": {"balcony", "balconies", "terrace"},
}


def coerce_ami(x):
    if pd.isna(x):
        return np.nan
    s = str(x).strip().upper().replace("%", "").replace("AMI", "")
    if s in {"Y", "YES", "TRUE", "1", "✓", "✔", "X"}:
        return 0.60
    try:
        v = float(s) / 100 if float(s) > 1 else float(s)
        return v if v in ALLOWED_BANDS else np.nan
    except Exception:
        return np.nan


def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    # Normalize incoming headers to a stripped, lowercased, no-space form
    df.columns = [
        c.strip().lower().replace(" ", "").replace("_", "").replace(".", "")
        for c in df.columns
    ]
    # Build a rename map to canonical titles (keys of FLEX_HEADERS)
    ren = {}
    for target, variants in FLEX_HEADERS.items():
        for c in df.columns:
            if c in [v.lower().replace(" ", "").replace("_", "") for v in variants]:
                ren[c] = target
    df = df.rename(columns=ren)

    if "AMI" in df.columns:
        df["AMI"] = df["AMI"].apply(coerce_ami)
    if "FLOOR" in df.columns:
        df["FLOOR"] = pd.to_numeric(df["FLOOR"], errors="coerce")
    if "BED" in df.columns:
        df["BED"] = pd.to_numeric(df["BED"], errors="coerce")
    if "NET SF" in df.columns:
        df["NET SF"] = pd.to_numeric(df["NET SF"], errors="coerce")
    if "BALCONY" in df.columns:
        df["BALCONY"] = df["BALCONY"].apply(
            lambda x: 1 if str(x).strip().upper() in {"Y", "YES", "1", "TRUE"} else 0
        )

    # Require NET SF for optimization rows
    return df.dropna(subset=["NET SF"]) if "NET SF" in df.columns else df


def calculate_revenue_weight(row: pd.Series) -> float:
    # Developer pref: higher floor/beds/SF/balcony = higher value
    return (
        (row.get("FLOOR", 1) * 0.3)
        + (row.get("BED", 0) * 0.4)
        + (row.get("NET SF", 500) / 1000 * 0.2)
        + (row.get("BALCONY", 0) * 0.1)
    )


def validate_selection(df: pd.DataFrame, required_sf: float) -> Dict:
    aff = df[df["AMI"].notna()] if "AMI" in df.columns else pd.DataFrame()
    total_aff_sf = aff["NET SF"].sum() if "NET SF" in aff.columns else 0.0
    if total_aff_sf < required_sf:
        return {
            "valid": False,
            "message": f"Affordable SF {total_aff_sf:.2f} < required {required_sf:.2f}",
        }
    # Add other validations (bed mix, vertical, horizontal, balcony) from docs - warn only
    return {"valid": True}


def milp_assign_ami(
    aff: pd.DataFrame, bands: List[float], required_40_pct: float = 0.20
) -> Tuple[np.ndarray, Dict]:
    """
    Assign each affordable unit a single AMI band subject to:
      - each unit gets exactly one band
      - at least `required_40_pct` of total affordable NET SF at 40% AMI
      - cap 40% AMI at 21% of total affordable NET SF
      - weighted average AMI <= 60%
      - objective balances higher AMI on higher-value units while staying near 60%
    """
    n = len(aff)
    prob = LpProblem("AMI_Assignment", LpMaximize)

    # Decision vars: assign[i][j] = 1 if unit i gets band j
    assign = [[LpVariable(f"assign_{i}_{j}", cat="Binary") for j in range(len(bands))] for i in range(n)]

    # Constraints
    for i in range(n):
        prob += lpSum(assign[i]) == 1  # one band per unit

    sf = aff["NET SF"].to_numpy()
    total_sf = sf.sum()

    if 0.40 in bands:
        sf_40 = lpSum(assign[i][bands.index(0.40)] * sf[i] for i in range(n))
        prob += sf_40 >= required_40_pct * total_sf
        prob += sf_40 <= 0.21 * total_sf  # (adjust if your policy differs)

    wavg = lpSum(
        lpSum(assign[i][j] * bands[j] * sf[i] for j in range(len(bands)))
        for i in range(n)
    ) / total_sf
    prob += wavg <= 0.60

    # Objective: Max avg + revenue (high bands on high-value) - |wavg - 0.60|
    revenue = aff.apply(calculate_revenue_weight, axis=1).to_numpy()
    obj = (
        wavg * 100
        - (0.60 - wavg) * 50
        + lpSum(
            lpSum(assign[i][j] * bands[j] * revenue[i] for j in range(len(bands)))
            for i in range(n)
        )
        / n
    )
    prob += obj

    # ---- FIX: use a CBC solver instance rather than passing timeLimit into solve() ----
    # If you're on an older PuLP, replace timeLimit=... with maxSeconds=...
    solver = PULP_CBC_CMD(timeLimit=DEFAULT_TIMELIMIT, msg=False)
    prob.solve(solver=solver)

    if LpStatus[prob.status] != "Optimal":
        return None, {}

    assigned = np.array(
        [bands[j] for i in range(n) for j in range(len(bands)) if value(assign[i][j]) == 1]
    )
    metrics = {
        "wavg": float(np.dot(assigned, sf) / total_sf),
        "pct40": float(sf[assigned == 0.40].sum() / total_sf) if 0.40 in bands else 0.0,
        "bands_count": int(len(set(assigned))),
    }
    return assigned, metrics


def generate_scenarios(
    df: pd.DataFrame, required_sf: float, prefs: Dict
) -> Dict:
    df = normalize_df(df)

    valid = validate_selection(df, required_sf)
    if not valid["valid"]:
        raise ValueError(valid["message"])

    # Work only on rows with an AMI value present (the affordable set)
    aff = df[df["AMI"].notna()].copy()

    from itertools import combinations

    # Enumerate band subsets that always include 40% AMI
    band_subsets = [
        [0.40] + list(combo)
        for r in range(1, 3)
        for combo in combinations([b for b in ALLOWED_BANDS if b > 0.40], r)
    ]

    scen: Dict[str, Dict] = {}
    for bands in band_subsets:
        assigned, metrics = milp_assign_ami(
            aff, bands, prefs.get("required_40_pct", 0.20)
        )
        if assigned is not None:
            scen[f"Opt{len(scen) + 1}"] = {"assigned": assigned, "metrics": metrics}

    if not scen:
        raise ValueError("No valid scenarios")

    # Rank by a simple score: max wavg, close to 20% at 40% AMI, and prefer ~2 bands
    def score(m):
        return m["wavg"] * 10 - abs(m["pct40"] - 0.20) * 5 - (m["bands_count"] - 2) * 2

    top_labels = sorted(scen, key=lambda k: score(scen[k]["metrics"]), reverse=True)[:2]
    return {k: scen[k] for k in top_labels}, aff


def build_outputs(
    df: pd.DataFrame, scen: Dict, aff: pd.DataFrame, prefs: Dict
) -> List[Tuple[str, bytes]]:
    outputs: List[Tuple[str, bytes]] = []

    for label, data in scen.items():
        out_df = df.copy()
        out_df.loc[aff.index, "AMI"] = data["assigned"]  # override

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            out_df.to_excel(writer, sheet_name="Master", index=False)
            aff.assign(AMI=data["assigned"]).to_excel(
                writer, sheet_name="Affordable Breakdown", index=False
            )
            metrics_df = pd.DataFrame([data["metrics"]])
            metrics_df.to_excel(writer, sheet_name="Summary", index=False)

            if genai:
                model = genai.GenerativeModel("gemini-1.5-pro")
                prompt = f"Explain why this AMI assignment: {data['metrics']}, prefs: {prefs}"
                expl = model.generate_content(prompt).text
                pd.DataFrame({"Explanation": [expl]}).to_excel(
                    writer, sheet_name="Logic", index=False
                )

        buf.seek(0)
        outputs.append((f"{label}.xlsx", buf.getvalue()))

    # (Optional) write-back hook—left to the FastAPI layer if you want it
    if prefs.get("write_back"):
        pass

    return outputs
