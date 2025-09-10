import os
from typing import List, Dict, Tuple
import numpy as np
import pandas as pd
from pulp import LpProblem, LpMaximize, LpVariable, lpSum, LpStatus, value

try:
    import google.generativeai as genai
    genai.configure(api_key=os.getenv("GEMINI_API_KEY"))  # Optional LLM explanations
except ImportError:
    genai = None

DEFAULT_TIMELIMIT = int(os.getenv("MILP_TIMELIMIT_SEC", "30"))
ALLOWED_BANDS = [0.40, 0.60, 0.70, 0.80, 0.90, 1.00]
FLEX_HEADERS = {  # Fuzzy matching
    "NET SF": {"net sf", "netsf", "sqft", "area", "square feet"},
    "AMI": {"ami", "aff ami", "assigned_ami"},
    "FLOOR": {"floor", "story", "level"},
    "APT": {"apt", "unit", "apartment"},
    "BED": {"bed", "beds", "bedroom", "br"},
    "BALCONY": {"balcony", "balconies", "terrace"},
}

def coerce_ami(x):
    if pd.isna(x): return np.nan
    s = str(x).strip().upper().replace('%', '').replace('AMI', '')
    if s in {'Y', 'YES', 'TRUE', '1', '✓', '✔', 'X'}: return 0.60
    try:
        v = float(s) / 100 if float(s) > 1 else float(s)
        return v if v in ALLOWED_BANDS else np.nan
    except:
        return np.nan

def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [c.strip().lower().replace(' ', '').replace('_', '').replace('.', '') for c in df.columns]
    ren = {}
    for target, variants in FLEX_HEADERS.items():
        for c in df.columns:
            if c in [v.lower() for v in variants]:
                ren[c] = target
    df = df.rename(columns=ren)
    if "AMI" in df.columns:
        df["AMI"] = df["AMI"].apply(coerce_ami)
    if "FLOOR" in df.columns:
        df["FLOOR"] = pd.to_numeric(df["FLOOR"], errors='coerce')
    if "BED" in df.columns:
        df["BED"] = pd.to_numeric(df["BED"], errors='coerce')
    if "NET SF" in df.columns:
        df["NET SF"] = pd.to_numeric(df["NET SF"], errors='coerce')
    if "BALCONY" in df.columns:
        df["BALCONY"] = df["BALCONY"].apply(lambda x: 1 if str(x).strip().upper() in {'Y', 'YES', '1', 'TRUE'} else 0)
    return df.dropna(subset=["NET SF"])

def calculate_revenue_weight(row):
    # Developer pref: higher floor/beds/SF/balcony = higher value
    return (row.get("FLOOR", 1) * 0.3) + (row.get("BED", 0) * 0.4) + (row.get("NET SF", 500) / 1000 * 0.2) + (row.get("BALCONY", 0) * 0.1)

def validate_selection(df: pd.DataFrame, required_sf: float) -> Dict:
    aff = df[df["AMI"].notna()]
    total_aff_sf = aff["NET SF"].sum()
    if total_aff_sf < required_sf:
        return {"valid": False, "message": f"Affordable SF {total_aff_sf:.2f} < required {required_sf:.2f}"}
    # Add other validations (bed mix, vertical, horizontal, balcony) from docs - warn only
    return {"valid": True}

def milp_assign_ami(aff: pd.DataFrame, bands: List[float], required_40_pct: float = 0.20) -> Tuple[np.ndarray, Dict]:
    n = len(aff)
    prob = LpProblem("AMI_Assignment", LpMaximize)
    assign = [[LpVariable(f"assign_{i}_{j}", cat='Binary') for j in range(len(bands))] for i in range(n)]
    # Constraints
    for i in range(n):
        prob += lpSum(assign[i]) == 1  # One band per unit
    sf = aff["NET SF"].to_numpy()
    total_sf = sf.sum()
    sf_40 = lpSum(assign[i][bands.index(0.40)] * sf[i] for i in range(n) if 0.40 in bands)
    prob += sf_40 >= required_40_pct * total_sf
    prob += sf_40 <= 0.21 * total_sf
    wavg = lpSum(lpSum(assign[i][j] * bands[j] * sf[i] for j in range(len(bands))) for i in range(n)) / total_sf
    prob += wavg <= 0.60
    # Objective: Max avg + revenue (high bands on high-value) - |wavg-0.60|
    revenue = aff.apply(calculate_revenue_weight, axis=1).to_numpy()
    obj = wavg * 100 - (0.60 - wavg) * 50 + lpSum(lpSum(assign[i][j] * bands[j] * revenue[i] for j in range(len(bands))) for i in range(n)) / n
    prob += obj
    prob.solve(timeLimit=DEFAULT_TIMELIMIT)
    if LpStatus[prob.status] != 'Optimal':
        return None, {}
    assigned = np.array([bands[j] for i in range(n) for j in range(len(bands)) if value(assign[i][j]) == 1])
    metrics = {
        "wavg": np.dot(assigned, sf) / total_sf,
        "pct40": (sf[assigned == 0.40].sum() / total_sf) if 0.40 in bands else 0,
        "bands_count": len(set(assigned)),
    }
    return assigned, metrics

def generate_scenarios(df: pd.DataFrame, required_sf: float, prefs: Dict) -> Dict:
    df = normalize_df(df)
    valid = validate_selection(df, required_sf)
    if not valid["valid"]:
        raise ValueError(valid["message"])
    aff = df[df["AMI"].notna()].copy()  # Pre-selected
    from itertools import combinations
    band_subsets = [ [0.40] + list(combo) for r in range(1, 3) for combo in combinations([b for b in ALLOWED_BANDS if b > 0.40], r) ]
    scen = {}
    for bands in band_subsets:
        assigned, metrics = milp_assign_ami(aff, bands, prefs.get("required_40_pct", 0.20))
        if assigned is not None:
            scen[f"Opt{len(scen)+1}"] = {"assigned": assigned, "metrics": metrics}
    if not scen:
        raise ValueError("No valid scenarios")
    # Select top 2 by score (max wavg, close to 20% 40, revenue prefs)
    def score(m): return m["wavg"] * 10 - abs(m["pct40"] - 0.20) * 5 - (m["bands_count"] - 2) * 2
    top_labels = sorted(scen, key=lambda k: score(scen[k]["metrics"]), reverse=True)[:2]
    return {k: scen[k] for k in top_labels}, aff

def build_outputs(df: pd.DataFrame, scen: Dict, aff: pd.DataFrame, prefs: Dict) -> List[Tuple[str, bytes]]:
    outputs = []
    for label, data in scen.items():
        out_df = df.copy()
        out_df.loc[aff.index, "AMI"] = data["assigned"]  # Override
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            out_df.to_excel(writer, sheet_name="Master", index=False)
            aff.assign(AMI=data["assigned"]).to_excel(writer, sheet_name="Affordable Breakdown", index=False)
            metrics_df = pd.DataFrame([data["metrics"]])
            metrics_df.to_excel(writer, sheet_name="Summary", index=False)
            if genai:  # Optional LLM logic
                model = genai.GenerativeModel('gemini-1.5-pro')
                prompt = f"Explain why this AMI assignment: {data['metrics']}, prefs: {prefs}"
                expl = model.generate_content(prompt).text
                pd.DataFrame({"Explanation": [expl]}).to_excel(writer, sheet_name="Logic", index=False)
        buf.seek(0)
        outputs.append((f"{label}.xlsx", buf.getvalue()))
    # Write-back option: Add to original if prefs['write_back']
    if prefs.get('write_back'):
        # Assume original is uploaded as bytes; in app, handle file write-back
        pass  # Implement in app.py
    return outputs
