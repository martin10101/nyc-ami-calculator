# ami_core.py
import math
from typing import Optional, Tuple, Dict

import numpy as np
import pandas as pd

# ---------------------------
# Fuzzy headers & value parsing
# ---------------------------

# --- REPLACE these helpers in ami_core.py ---

FLEX_HEADERS = {
    "NET SF": {"netsf","net sf","net_sf","net s.f.","sf","sqft","sq ft","square feet","net area","area"},
    "AMI": {"ami","aime","aff ami","affordable ami","assigned_ami","aff_ami","aff-ami"},
    "FLOOR": {"floor","fl","story","level"},
    "APT": {"apt","apartment","unit","unit #","apt#","apt no","apartment #"},
    "BED": {"bed","beds","bedroom","br"},
    "AFF": {"aff","affordable","selected","ami_selected","is_affordable","target_ami"},
}

def _coerce_percent(x):
    if pd.isna(x): return np.nan
    if isinstance(x,(int,float)):
        v=float(x); return v/100.0 if v>1.0 else v
    s=str(x).strip()
    sU=s.upper()
    sU=sU.replace("AIME","").replace("AMI","").replace("%","").strip()
    # treat checkmarks / x / yes as "selected"
    if sU in {"Y","YES","TRUE","1","✓","✔","X"}:
        return 0.60   # placeholder if someone asks for a numeric; not used by solver
    try:
        v=float(sU); return v/100.0 if v>1.0 else v
    except:
        return np.nan

def _normalize_headers(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    lc = {c: c.lower().replace(" ", "").replace(".", "").replace("_","") for c in df.columns}
    ren = {}
    for target, variants in FLEX_HEADERS.items():
        for c in df.columns:
            if lc[c] in variants: ren[c]=target
    if ren: df = df.rename(columns=ren)
    if "NET SF" in df.columns: df["NET SF"] = pd.to_numeric(df["NET SF"], errors="coerce")
    if "FLOOR" in df.columns: df["FLOOR"] = pd.to_numeric(df["FLOOR"], errors="coerce")
    if "BED" in df.columns: df["BED"] = pd.to_numeric(df["BED"], errors="coerce")
    if "AMI" in df.columns: df["AMI_RAW"] = df["AMI"]  # keep original cell for “selected” detection
    if "AMI" in df.columns: df["AMI"] = df["AMI"].apply(_coerce_percent)
    return df

def _selected_for_ami_mask(df: pd.DataFrame) -> pd.Series:
    """
    A row is selected if the AMI cell is *non-empty* (any value at all: 0.6, 60%, x, ✓, Yes, etc.),
    OR if there is an explicit Affordable/Selected flag set to Yes/True/1.
    We don't trust the numeric; we only use this as a yes/no selection.
    """
    # explicit yes/no flag takes priority
    for c in df.columns:
        key = str(c).strip().lower()
        if key in {"aff","affordable","selected","ami_selected","is_affordable","target_ami"}:
            return df[c].astype(str).str.strip().str.lower().isin({"1","true","yes","y","✓","✔","x"}) & df["NET SF"].notna()

    # otherwise: any non-empty AMI cell means selected
    if "AMI_RAW" in df.columns:
        return df["AMI_RAW"].astype(str).str.strip().ne("").fillna(False) & df["NET SF"].notna()

    # last resort: if an AMI numeric exists, use non-null numeric
    if "AMI" in df.columns:
        return df["AMI"].notna() & df["NET SF"].notna()

    return pd.Series(False, index=df.index)

def _sf_avg(ami: np.ndarray, sf: np.ndarray) -> float:
    return float(np.dot(ami, sf)/sf.sum()) if sf.sum() > 0 else 0.0


# ---------------------------
# Exact 40% SF picker (20–21% band)
# ---------------------------

def _choose_40_mitm(sf: np.ndarray, low_share=0.20, high_share=0.21) -> np.ndarray:
    """
    Meet-in-the-middle exact subset search. Returns a boolean mask choosing
    units whose SF sum falls in [low_share, high_share] of total affordable SF.
    """
    n = len(sf)
    total = float(sf.sum()); lo = total*low_share; hi = total*high_share
    idx = np.arange(n)
    A = idx[: n//2]; B = idx[n//2:]
    sums_a = []
    for m in range(1<<len(A)):
        s=0.0
        for i in range(len(A)):
            if m & (1<<i): s += float(sf[A[i]])
        sums_a.append((s,m))
    sums_b = []
    for m in range(1<<len(B)):
        s=0.0
        for i in range(len(B)):
            if m & (1<<i): s += float(sf[B[i]])
        sums_b.append((s,m))
    sums_b.sort(key=lambda x: x[0])
    bvals = [x[0] for x in sums_b]
    import bisect
    # try exact-in-band
    for s_a, m_a in sums_a:
        L = bisect.bisect_left(bvals, lo - s_a)
        R = bisect.bisect_right(bvals, hi - s_a)
        if L < R:
            mid = (L + R - 1)//2
            s_b, m_b = sums_b[mid]
            mask = np.zeros(n, dtype=bool)
            for i in range(len(A)):
                if m_a & (1<<i): mask[A[i]] = True
            for i in range(len(B)):
                if m_b & (1<<i): mask[B[i]] = True
            return mask
    # else choose closest to mid
    best=None; mid_target = (lo+hi)/2
    for s_a, m_a in sums_a:
        pos = bisect.bisect_left(bvals, mid_target - s_a)
        for cand in (pos-1,pos,pos+1):
            if 0 <= cand < len(bvals):
                s_b, m_b = sums_b[cand]
                tot = s_a + s_b
                gap = max(0.0, lo - tot, tot - hi)
                if (best is None) or (gap < best[0]):
                    best = (gap, m_a, m_b)
    if best:
        _, m_a, m_b = best
        mask = np.zeros(n, dtype=bool)
        for i in range(len(A)):
            if m_a & (1<<i): mask[A[i]] = True
        for i in range(len(B)):
            if m_b & (1<<i): mask[B[i]] = True
        return mask
    return np.zeros(n, dtype=bool)

def choose_40pct_subset_by_sf(
    sf: np.ndarray,
    floors: np.ndarray,
    low_share: float = 0.20,
    high_share: float = 0.21,
    prefer_high_floors_for_40s: bool = True,
) -> np.ndarray:
    """
    Returns boolean mask for 40% units. Uses exact MITM for n<=34; greedy+swaps otherwise.
    Preserves the [lo,hi] band during any swaps/biasing.
    """
    n = len(sf)
    total = float(sf.sum()); lo = total*low_share; hi = total*high_share

    if n <= 34:
        mask = _choose_40_mitm(sf, low_share, high_share)
    else:
        # Greedy + swaps fallback for large sets
        order = np.argsort(sf)
        chosen=[]; s=0.0
        for i in order:
            if s<lo: chosen.append(i); s+=float(sf[i])
        not_chosen=[i for i in order if i not in chosen]
        improved=True
        while s>hi and improved:
            improved=False
            for c in sorted(chosen, key=lambda k: sf[k], reverse=True):
                for j in not_chosen:
                    cand = s - float(sf[c]) + float(sf[j])
                    if lo <= cand <= hi:
                        chosen.remove(c); chosen.append(j)
                        not_chosen.remove(j); not_chosen.append(c)
                        s=cand; improved=True; break
                if improved: break
        mask = np.zeros(n, dtype=bool); mask[np.array(chosen, dtype=int)] = True

    # Bias 40% toward higher floors while keeping band intact
    if prefer_high_floors_for_40s and (floors is not None) and not np.all(np.isnan(floors)):
        cur = float(sf[mask].sum())
        chosen = np.where(mask)[0].tolist()
        others = np.where(~mask)[0].tolist()
        chosen.sort(key=lambda i: floors[i] if not np.isnan(floors[i]) else -1)
        others.sort(key=lambda i: -(floors[i] if not np.isnan(floors[i]) else -1))
        for c in chosen:
            for j in others:
                cand = cur - float(sf[c]) + float(sf[j])
                if lo <= cand <= hi and (np.isnan(floors[c]) or np.isnan(floors[j]) or floors[j] > floors[c]):
                    mask[c]=False; mask[j]=True; cur=cand; break
    return mask

# ---------------------------
# Average balancer 59–60%
# ---------------------------

def balance_average(
    sf: np.ndarray,
    floors: np.ndarray,
    base_mask_40: np.ndarray,
    avg_low: float = 0.59,
    avg_high: float = 0.60,
    raise_bottom_first: bool = True,
    cap_target: float = 0.600 - 1e-6,
) -> np.ndarray:
    assigned = np.full(len(sf), 0.60, dtype=float)
    assigned[base_mask_40] = 0.40
    def cur(): return _sf_avg(assigned, sf)

    idx = np.where(~base_mask_40)[0].tolist()
    if raise_bottom_first and (floors is not None):
        idx.sort(key=lambda i: ((floors[i] if not np.isnan(floors[i]) else 10**9), -sf[i]))
    else:
        idx.sort(key=lambda i: -sf[i])

    while cur() < cap_target:
        progressed=False
        for i in idx:
            if cur() >= cap_target: break
            if assigned[i] < 1.0 - 1e-12:
                assigned[i] = round(assigned[i] + 0.1, 1)
                progressed=True
        if not progressed: break

    if cur() > avg_high + 1e-12:
        raised = [i for i in idx if assigned[i] > 0.60]
        raised.sort(key=lambda i: sf[i])
        for i in raised:
            old = assigned[i]
            assigned[i] = round(assigned[i] - 0.1, 1)
            if avg_low <= cur() <= avg_high: break
            assigned[i] = old

    final = cur()
    if not (avg_low - 1e-12 <= final <= avg_high + 1e-12):
        raise RuntimeError(f"Average could not be balanced into band: {final:.6f}")
    return assigned

# ---------------------------
# Scenarios (with optional mix rules)
# ---------------------------

def generate_scenarios(
    aff_df: pd.DataFrame,
    low_share=0.20, high_share=0.21,
    avg_low=0.59, avg_high=0.60,
    require_family_at_40: bool = False,
    spread_40_max_per_floor: Optional[int] = None,
    exempt_top_k_floors: int = 0,
) -> Dict[str, Dict]:
    sf = aff_df["NET SF"].to_numpy(float)
    floors = aff_df["FLOOR"].to_numpy(float) if "FLOOR" in aff_df.columns else np.full(len(aff_df), np.nan)
    beds = aff_df["BED"].to_numpy(float) if "BED" in aff_df.columns else np.full(len(aff_df), np.nan)

    def postprocess_spread(mask40: np.ndarray, max_per_floor: Optional[int], hi_share: float) -> np.ndarray:
        if max_per_floor is None or np.all(np.isnan(floors)): return mask40
        total=float(sf.sum()); lo=total*low_share; hi=total*hi_share
        cur=float(sf[mask40].sum())
        fl = pd.Series(floors).fillna(-1).astype(int).to_numpy()
        m = mask40.copy()
        changed=True
        while changed:
            changed=False
            counts={}
            for i,f in enumerate(fl):
                if m[i]: counts[f] = counts.get(f,0)+1
            viol=[f for f,c in counts.items() if c>max_per_floor]
            if not viol: break
            for vf in viol:
                idx_40=[i for i in range(len(m)) if m[i] and fl[i]==vf]
                if not idx_40: continue
                victim=max(idx_40, key=lambda i: sf[i])
                cands=[j for j in range(len(m)) if (not m[j]) and fl[j]!=vf]
                cands.sort(key=lambda j: sf[j])
                swapped=False
                for j in cands:
                    nt = cur - float(sf[victim]) + float(sf[j])
                    if lo <= nt <= hi:
                        m[victim]=False; m[j]=True; cur=nt; changed=True; swapped=True; break
        return m

    def build(prefer_high_for_40: bool, shallow_extra: bool, protect_top: int, need_family: bool):
        hi_share = min(high_share + (0.002 if shallow_extra else 0.0), 0.22)
        m40 = choose_40pct_subset_by_sf(sf, floors, low_share, hi_share, prefer_high_for_40)
        # require at least one 2BR at 40%
        if need_family and not np.any(m40 & (beds >= 2)):
            fam_idx = np.where(beds >= 2)[0]
            if len(fam_idx) > 0:
                fam = fam_idx[np.argmin(sf[fam_idx])]
                if not m40[fam]:
                    total=float(sf[m40].sum()); lo=float(sf.sum())*low_share; hi=float(sf.sum())*hi_share
                    drop = np.where(m40)[0][np.argmax(sf[m40])]
                    cand = total - float(sf[drop]) + float(sf[fam])
                    if lo <= cand <= hi: m40[drop]=False; m40[fam]=True
        # protect top floors best-effort
        if protect_top and not np.all(np.isnan(floors)):
            top_cut = np.nanmax(floors) - protect_top + 1
            if not np.isnan(top_cut):
                total=float(sf[m40].sum()); lo=float(sf.sum())*low_share; hi=float(sf.sum())*hi_share
                top_40=[i for i in np.where(m40)[0] if floors[i] >= top_cut]
                non_top=[i for i in np.where(~m40)[0] if floors[i] < top_cut]
                top_40.sort(key=lambda i: -sf[i]); non_top.sort(key=lambda i: sf[i])
                for a in top_40:
                    for b in non_top:
                        cand = total - float(sf[a]) + float(sf[b])
                        if lo <= cand <= hi: m40[a]=False; m40[b]=True; total=cand; break
        # spread cap
        m40 = postprocess_spread(m40, spread_40_max_per_floor, hi_share)
        # balance average
        a = balance_average(sf, floors, m40, avg_low, avg_high, True, avg_high-1e-6)
        return m40, a

    A_m40, A_ass = build(True,  False, exempt_top_k_floors, require_family_at_40)
    B_m40, B_ass = build(True,  True,  exempt_top_k_floors, True)
    C_m40, C_ass = build(False, False, max(exempt_top_k_floors,2), require_family_at_40)

    def metrics(mask40, assigned):
        total=float(sf.sum()); sf40=float(sf[mask40].sum())
        return {"aff_sf_total": total, "sf_at_40": sf40,
                "pct40": (sf40/total*100.0) if total else 0.0,
                "wavg": _sf_avg(assigned, sf)*100.0}

    out = {
        "A": {"mask40": A_m40, "assigned": A_ass, "metrics": metrics(A_m40, A_ass)},
        "B": {"mask40": B_m40, "assigned": B_ass, "metrics": metrics(B_m40, B_ass)},
        "C": {"mask40": C_m40, "assigned": C_ass, "metrics": metrics(C_m40, C_ass)},
    }

    # Hard compliance guardrails
    for k in out:
        m = out[k]["metrics"]
        if not (20.0 - 1e-6 <= m["pct40"] <= 21.0 + 1e-6):
            raise RuntimeError(f"Scenario {k}: 40% share out of band ({m['pct40']:.3f}%).")
        if not (59.0 - 1e-6 <= m["wavg"] <= 60.0 + 1e-6):
            raise RuntimeError(f"Scenario {k}: Weighted average out of band ({m['wavg']:.3f}%).")
    return out

# ---------------------------
# Public API
# ---------------------------

# --- REPLACE allocate_with_scenarios in ami_core.py ---

def allocate_with_scenarios(
    df: pd.DataFrame,
    low_share=0.20, high_share=0.21, avg_low=0.59, avg_high=0.60,
    require_family_at_40: bool = False,
    spread_40_max_per_floor: Optional[int] = None,
    exempt_top_k_floors: int = 0,
):
    df = _normalize_headers(df)
    if "NET SF" not in df.columns:
        raise ValueError("Missing NET SF column after normalization.")

    sel = _selected_for_ami_mask(df)
    if not sel.any():
        raise ValueError("No AMI-selected rows detected. Fill the AMI column (any value) to mark selected rows.")

    aff = df.loc[sel].copy()

    scen = generate_scenarios(
        aff,
        low_share=low_share, high_share=high_share, avg_low=avg_low, avg_high=avg_high,
        require_family_at_40=require_family_at_40,
        spread_40_max_per_floor=spread_40_max_per_floor,
        exempt_top_k_floors=exempt_top_k_floors,
    )

    # attach scenario assignments to the full table (separate columns)
    full = df.copy()
    for label in ["A","B","C"]:
        full[f"Assigned_AMI_{label}"] = np.nan
        full.loc[sel, f"Assigned_AMI_{label}"] = scen[label]["assigned"]

    # make an “affordable only” breakdown for each scenario
    base_cols = [c for c in ["FLOOR","APT","BED","NET SF","AMI_RAW"] if c in df.columns]
    aff_br = df.loc[sel, base_cols].copy()
    for label in ["A","B","C"]:
        aff_br[f"Assigned_AMI_{label}"] = scen[label]["assigned"]

    # choose “best” scenario for mirror write-back (closest to 20% and 60%)
    def score(m):
        wavg = min(m["wavg"], 60.0)
        pct  = max(min(m["pct40"], 21.0), 20.0)
        return 1000.0 - abs(60.0 - wavg)*50.0 - abs(20.0 - pct)*100.0

    metrics = {k:v["metrics"] for k,v in scen.items()}
    best_label = max(["A","B","C"], key=lambda k: score(metrics[k]))
    best_assigned = scen[best_label]["assigned"]

    # build a mirror of the original sheet:
    # overwrite AMI for selected rows with the chosen scenario,
    # preserve every other column and row order; append a totals block.
    mirror = df.copy()
    # If original had AMI_RAW and AMI, prefer to overwrite the visible AMI column if present,
    # otherwise create it.
    target_col = "AMI" if "AMI" in mirror.columns else "AMI"
    mirror[target_col] = mirror.get(target_col, np.nan)
    mirror.loc[sel, target_col] = best_assigned  # decimals: 0.4, 0.6, etc.

    # append a small totals block (blank row + 4 lines)
    total_sf = float(aff["NET SF"].sum())
    sf40 = float(aff["NET SF"].to_numpy()[np.isclose(best_assigned, 0.40)].sum())
    pct40 = (sf40 / total_sf * 100.0) if total_sf else 0.0
    wavg = _sf_avg(best_assigned, aff["NET SF"].to_numpy()) * 100.0

    footer = pd.DataFrame({
        list(mirror.columns)[0]: [
            "", "Affordable SF total", "SF at 40% AMI", "% at 40% AMI", "Weighted Avg AMI"
        ],
        list(mirror.columns)[1] if len(mirror.columns) > 1 else "Value": [
            "", f"{total_sf:,.2f}", f"{sf40:,.2f}", f"{pct40:.3f}%", f"{wavg:.3f}%"
        ]
    })
    mirror_out = pd.concat([mirror, footer], ignore_index=True)

    return full, aff_br, metrics, mirror_out, best_label
