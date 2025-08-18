# ami_core.py  — upgraded with:
# - Max 3 bands constraint (auto-merge)
# - Vertical placement preference (objective tie-break)
# - Dominance filter (Pareto best scenarios only)
# - Exact optimizer (PuLP) with graceful fallback
#
# Backward compatible with your existing app endpoints.

import numpy as np
import pandas as pd
from typing import Optional, Dict, List, Tuple

# ----------------------------
# Header normalization + select
# ----------------------------

FLEX_HEADERS = {
    "NET SF": {"netsf","net sf","net_sf","net s.f.","sf","sqft","sq ft","square feet","net area","area","area(sf)"},
    "AMI": {"ami","aime","aff ami","affordable ami","assigned_ami","aff_ami","aff-ami"},
    "FLOOR": {"floor","fl","story","level"},
    "APT": {"apt","apartment","unit","unit #","apt#","apt no","apartment #","apartment"},
    "BED": {"bed","beds","bedroom","br","bedrooms"},
    "AFF": {"aff","affordable","selected","ami_selected","is_affordable","target_ami"},
    "SIGNED_AMI": {"signed_ami","signedami","assigned_ami","assignedami"},
}

def _coerce_percent(x):
    """Normalize things like 0.6 / 60 / '60%' / 'x' / 'yes' => float in [0,1] or NaN.
       We do NOT trust this for assignment; it's just for selection if needed."""
    if pd.isna(x): return np.nan
    if isinstance(x,(int,float)):
        v=float(x); return v/100.0 if v>1.0 else v
    s=str(x).strip()
    sU=s.upper().replace("AIME","").replace("AMI","").replace("%","").strip()
    if sU in {"Y","YES","TRUE","1","✓","✔","X"}:
        return 0.60
    try:
        v=float(sU); return v/100.0 if v>1.0 else v
    except:
        return np.nan

def _normalize_headers(df: pd.DataFrame) -> pd.DataFrame:
    d = df.copy()
    d.columns = [str(c).strip() for c in d.columns]
    lc = {c: c.lower().replace(" ", "").replace(".", "").replace("_","") for c in d.columns}
    ren = {}
    for target, variants in FLEX_HEADERS.items():
        for c in d.columns:
            if lc[c] in variants: ren[c]=target
    if ren: d = d.rename(columns=ren)
    # keep original AMI cell for selection
    if "AMI" in d.columns: d["AMI_RAW"] = d["AMI"]
    if "NET SF" in d.columns: d["NET SF"] = pd.to_numeric(d["NET SF"], errors="coerce")
    if "FLOOR" in d.columns: d["FLOOR"] = pd.to_numeric(d["FLOOR"], errors="coerce")
    if "BED" in d.columns: d["BED"] = pd.to_numeric(d["BED"], errors="coerce")
    if "AMI" in d.columns: d["AMI"] = d["AMI"].apply(_coerce_percent)
    return d

def _selected_for_ami_mask(d: pd.DataFrame) -> pd.Series:
    """
    Which rows are PRESELECTED by the client?
    Priority:
      1) Dedicated flag column (Affordable/Selected/etc) → Yes/True/1/✓/✔/X.
      2) 'SIGNED_AMI' style column → any non-empty cell.
      3) AMI column → treat only X/✓/✔/YES/Y/1 or numeric near 0.6 (0.55–0.65) as 'selected'.
    Always require NET SF present.
    """
    # 1) explicit flag
    for c in d.columns:
        key=str(c).strip().lower()
        if key in {"aff","affordable","selected","ami_selected","is_affordable","target_ami"}:
            return d[c].astype(str).str.strip().str.lower().isin(
                {"1","true","yes","y","✓","✔","x"}
            ) & d["NET SF"].notna()
    # 2) signed_ami style
    for c in d.columns:
        if str(c).strip().lower() in FLEX_HEADERS["SIGNED_AMI"]:
            return d[c].astype(str).str.strip().ne("").fillna(False) & d["NET SF"].notna()
    # 3) AMI RAW near 0.6 or explicit ticks
    if "AMI_RAW" in d.columns:
        raw = d["AMI_RAW"].astype(str).str.strip()
        yesish = raw.str.upper().isin({"X","✓","✔","YES","Y","1"})
        def _num_is_06(s):
            try:
                v = float(str(s).upper().replace("%","").replace("AMI","").replace("AIME","").strip())
                v = v/100.0 if v>1.0 else v
                return 0.55 <= v <= 0.65
            except:
                return False
        near06 = raw.apply(_num_is_06)
        return (yesish | near06) & d["NET SF"].notna()
    # fallback
    return pd.Series(False, index=d.index)

# ----------------------------
# Math helpers
# ----------------------------

def _sf_avg(ami: np.ndarray, sf: np.ndarray) -> float:
    return float(np.dot(ami, sf)/sf.sum()) if sf.sum()>0 else 0.0

def _unique_bands(amis: np.ndarray) -> List[float]:
    return sorted({round(float(x),2) for x in amis.tolist()})

# --- 40% selectors (MITM for small n; greedy+swaps for large) ---

def _choose_40_mitm(sf: np.ndarray, low_share=0.20, high_share=0.21) -> np.ndarray:
    import bisect
    n=len(sf); total=float(sf.sum()); lo=total*low_share; hi=total*high_share
    idx=np.arange(n); A=idx[:n//2]; B=idx[n//2:]
    sums_a=[]
    for m in range(1<<len(A)):
        s=0.0
        for i in range(len(A)):
            if m & (1<<i): s += float(sf[A[i]])
        sums_a.append((s,m))
    sums_b=[]
    for m in range(1<<len(B)):
        s=0.0
        for i in range(len(B)):
            if m & (1<<i): s += float(sf[B[i]])
        sums_b.append((s,m))
    sums_b.sort(key=lambda x:x[0]); bvals=[x[0] for x in sums_b]
    # exact band
    for s_a, m_a in sums_a:
        L=bisect.bisect_left(bvals, lo - s_a)
        R=bisect.bisect_right(bvals, hi - s_a)
        if L<R:
            mid=(L+R-1)//2; s_b, m_b = sums_b[mid]
            mask=np.zeros(n,dtype=bool)
            for i in range(len(A)):
                if m_a & (1<<i): mask[A[i]] = True
            for i in range(len(B)):
                if m_b & (1<<i): mask[B[i]] = True
            return mask
    # closest to mid
    best=None; mid_target=(lo+hi)/2
    for s_a, m_a in sums_a:
        pos=bisect.bisect_left(bvals, mid_target - s_a)
        for cand in (pos-1,pos,pos+1):
            if 0<=cand<len(bvals):
                s_b, m_b = sums_b[cand]; tot = s_a + s_b
                gap = max(0.0, lo - tot, tot - hi)
                if (best is None) or (gap < best[0]): best=(gap,m_a,m_b)
    if best:
        _, m_a, m_b = best
        mask=np.zeros(n,dtype=bool)
        for i in range(len(A)):
            if m_a & (1<<i): mask[A[i]] = True
        for i in range(len(B)):
            if m_b & (1<<i): mask[B[i]] = True
        return mask
    return np.zeros(n,dtype=bool)

def _choose_40_large(sf: np.ndarray, low_share=0.20, high_share=0.21) -> np.ndarray:
    total=float(sf.sum()); lo=total*low_share; hi=total*high_share
    order=np.argsort(sf)  # small → large
    chosen=[]; s=0.0
    for i in order:
        if s < lo:
            chosen.append(i); s += float(sf[i])
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
    m=np.zeros(len(sf),dtype=bool); m[np.array(chosen,dtype=int)]=True
    return m

def choose_40pct_subset_by_sf(sf: np.ndarray, floors: np.ndarray, low_share=0.20, high_share=0.21, prefer_high=True):
    n=len(sf)
    if n<=34:
        mask=_choose_40_mitm(sf, low_share, high_share)
    else:
        mask=_choose_40_large(sf, low_share, high_share)

    # Optional bias: push 40% set toward higher floors while staying inside band
    total=float(sf.sum()); lo=total*low_share; hi=total*high_share
    if prefer_high and floors is not None and not np.all(np.isnan(floors)):
        cur=float(sf[mask].sum())
        chosen=np.where(mask)[0].tolist()
        others=np.where(~mask)[0].tolist()
        chosen.sort(key=lambda i: floors[i] if not np.isnan(floors[i]) else -1)
        others.sort(key=lambda i: -(floors[i] if not np.isnan(floors[i]) else -1))
        for c in chosen:
            for j in others:
                cand=cur - float(sf[c]) + float(sf[j])
                if lo <= cand <= hi and (np.isnan(floors[c]) or np.isnan(floors[j]) or floors[j] > floors[c]):
                    mask[c]=False; mask[j]=True; cur=cand; break
    return mask

def ensure_band(mask40: np.ndarray, sf: np.ndarray, low_share: float, high_share: float) -> np.ndarray:
    """Repair 40% mask into [low, high] share by SF with minimal changes."""
    m = mask40.copy()
    total = float(sf.sum()); lo = total*low_share; hi = total*high_share
    s = float(sf[m].sum())

    # Too low → add smallest outsiders
    if s < lo - 1e-9:
        outsiders = [i for i in np.where(~m)[0]]
        outsiders.sort(key=lambda i: sf[i])
        for j in outsiders:
            cand = s + float(sf[j])
            m[j] = True; s = cand
            if s >= lo - 1e-9: break

    # Too high → swaps/trim
    if s > hi + 1e-9:
        chosen = [i for i in np.where(m)[0]]
        outsiders = [i for i in np.where(~m)[0]]
        chosen.sort(key=lambda i: sf[i], reverse=True)
        outsiders.sort(key=lambda i: sf[i])

        changed=True
        while s > hi + 1e-9 and changed:
            changed=False
            for c in list(chosen):
                for j in list(outsiders):
                    cand = s - float(sf[c]) + float(sf[j])
                    if lo - 1e-9 <= cand <= hi + 1e-9:
                        m[c]=False; m[j]=True; s=cand
                        chosen.remove(c); outsiders.remove(j)
                        changed=True; break
                if changed: break

        if s > hi + 1e-9:
            for c in list(chosen):
                if s - float(sf[c]) >= lo - 1e-9:
                    m[c]=False; s -= float(sf[c])
                if s <= hi + 1e-9: break

    final = float(sf[m].sum())
    if not (lo - 1e-9 <= final <= hi + 1e-9):
        raise RuntimeError(f"40% share out of band after repair ({final/total*100.0:.3f}%).")
    return m

# ----------------------------
# Assignment (averages) – heuristic path
# ----------------------------

def balance_average(
    sf: np.ndarray,
    floors: np.ndarray,
    mask40: np.ndarray,
    avg_low=0.59, avg_high=0.60,
    prefer_bottom: bool=True,
    target: Optional[float]=None,
    max_tier: float=1.0,
    preseed: Optional[List[Tuple[float,float]]] = None,   # [(tier, share_of_non40_sf), ...]
) -> np.ndarray:
    """
    Assign 60→70→80→90→100 on non-40% units to reach avg in [avg_low, avg_high] *without ever exceeding 60%*.
    - max_tier caps the highest AMI allowed on non-40% units (e.g., 0.70 for strict).
    - preseed allows R-variants to set a slice of premium units to 0.8/0.9/1.0 first.
    """
    assigned = np.full(len(sf), 0.60, dtype=float)
    assigned[mask40] = 0.40
    non40_idx = np.where(~mask40)[0]

    def cur_avg() -> float:
        return _sf_avg(assigned, sf)

    # ---------- helper to push average DOWN safely when we overshoot ----------
    def clamp_down_to(max_avg: float):
        if cur_avg() <= max_avg + 1e-12:
            return
        raised = [i for i in non40_idx if assigned[i] > 0.60]
        raised.sort(key=lambda i: (assigned[i], sf[i]))
        for i in reversed(raised):
            while assigned[i] > 0.60 + 1e-12 and cur_avg() > max_avg + 1e-12:
                assigned[i] = round(assigned[i] - 0.1, 1)

    # ---------- optional preseed for relaxed scenarios ----------
    if preseed:
        order = sorted(
            non40_idx,
            key=lambda i: (-(floors[i] if floors is not None and not np.isnan(floors[i]) else -1), -sf[i])
        )
        non40_sf_total = float(sf[non40_idx].sum())
        used = set()
        for tier, share in preseed:
            target_sf = non40_sf_total * max(0.0, min(share, 1.0))
            acc = 0.0
            for i in order:
                if i in used: continue
                if tier > max_tier: continue
                assigned[i] = tier
                used.add(i)
                acc += float(sf[i])
                if acc >= target_sf - 1e-9: break
        clamp_down_to(avg_high - 1e-6)

    # ---------- main “raise” loop (never aims above avg_high) ----------
    target_avg = min(target if target is not None else (avg_high - 1e-6), avg_high - 1e-6)

    if floors is not None and not np.all(np.isnan(floors)):
        idx = sorted(non40_idx, key=lambda i: ((floors[i] if not np.isnan(floors[i]) else 1e9), -sf[i]))
    else:
        idx = sorted(non40_idx, key=lambda i: -sf[i])

    while cur_avg() < target_avg - 1e-12:
        progressed = False
        for i in idx:
            if cur_avg() >= target_avg - 1e-12:
                break
            if assigned[i] < 1.0 - 1e-12 and assigned[i] < max_tier - 1e-12:
                old = assigned[i]
                assigned[i] = round(min(old + 0.1, max_tier), 1)
                if cur_avg() > avg_high + 1e-12:
                    assigned[i] = old
                else:
                    progressed = True
        if not progressed:
            break

    if cur_avg() > avg_high + 1e-12:
        clamp_down_to(avg_high - 1e-6)

    attempts = 0
    while cur_avg() < avg_low - 1e-12 and attempts < 3:
        for i in idx:
            old = assigned[i]
            if old < 1.0 - 1e-12 and old < max_tier - 1e-12:
                assigned[i] = round(min(old + 0.1, max_tier), 1)
                if cur_avg() > avg_high + 1e-12:
                    assigned[i] = old
        attempts += 1

    if not (avg_low - 1e-12 <= cur_avg() <= avg_high + 1e-12):
        if cur_avg() > avg_high + 1e-12:
            clamp_down_to(avg_high - 1e-6)

    if not (avg_low - 1e-12 <= cur_avg() <= avg_high + 1e-12):
        raise RuntimeError(f"Weighted average out of band after balance ({cur_avg()*100:.3f}%).")

    return assigned

# ----------------------------
# Power upgrade: exact optimizer (PuLP) — optional
# ----------------------------

def _try_exact_optimize(sf: np.ndarray, floors: np.ndarray, mask40: np.ndarray,
                        avg_low: float, avg_high: float,
                        low_share: float, high_share: float,
                        allowed_bands=(0.40,0.60,0.70,0.80,0.90,1.00),
                        max_bands:int=3,
                        alpha:float=0.02, beta:float=0.02) -> Optional[np.ndarray]:
    """
    Solve an exact assignment via MILP if PuLP is available.
    Objective ~ maximize average AMI + vertical preference (higher floors for high AMI; lower floors for 40%).
    Returns np.ndarray of AMIs or None if solver not available/fails.
    """
    try:
        import pulp
    except Exception:
        return None

    n = len(sf)
    I = list(range(n))
    bands = sorted(set(allowed_bands))
    # preload masks
    is40_fixed = mask40.copy()

    # Problem
    prob = pulp.LpProblem("AMI_Assignment", pulp.LpMaximize)

    # Variables: x[i,b] binary for each unit+band; y[b] band used
    x = {(i,b): pulp.LpVariable(f"x_{i}_{int(b*100)}", 0, 1, pulp.LpBinary) for i in I for b in bands}
    y = {b: pulp.LpVariable(f"y_{int(b*100)}", 0, 1, pulp.LpBinary) for b in bands}

    # Each unit exactly one band
    for i in I:
        prob += pulp.lpSum(x[i,b] for b in bands) == 1

    # Lock 40% units
    for i in I:
        if is40_fixed[i]:
            prob += x[i,0.40] == 1
            for b in bands:
                if b != 0.40: prob += x[i,b] == 0
        else:
            prob += x[i,0.40] == 0  # non-40% cannot be 40 in this stage

    # 40% share band by SF
    total_sf = float(sf.sum())
    sf_40 = pulp.lpSum(sf[i] * x[i,0.40] for i in I)
    prob += sf_40 >= total_sf * low_share - 1e-6
    prob += sf_40 <= total_sf * high_share + 1e-6

    # Weighted average AMI in [avg_low, avg_high]
    wavg_num = pulp.lpSum(sf[i]*pulp.lpSum(b * x[i,b] for b in bands) for i in I)
    prob += wavg_num >= total_sf * avg_low - 1e-6
    prob += wavg_num <= total_sf * avg_high + 1e-6

    # Band count <= max_bands
    for b in bands:
        for i in I:
            prob += x[i,b] <= y[b]
    prob += pulp.lpSum(y[b] for b in bands) <= max_bands

    # Vertical preference scoring
    # avg_hi_floor ≈ sum(floor_i * sum_{b>=0.70} x[i,b]) / (#hi units) — we linearize by just summing floors * x
    Floors = np.nan_to_num(floors, nan=0.0)
    hi_term = pulp.lpSum(Floors[i]*pulp.lpSum(x[i,b] for b in bands if b>=0.70) for i in I)
    lo40_term = pulp.lpSum(Floors[i]*x[i,0.40] for i in I)

    # Objective: maximize average + alpha*hi - beta*low40
    # To keep scale similar, divide the vertical terms by (#units) if >0
    denom = max(1, n)
    prob += (wavg_num / total_sf) + alpha * (hi_term/denom) - beta * (lo40_term/denom)

    # Solve
    prob.solve(pulp.PULP_CBC_CMD(msg=False, timeLimit=10))

    if pulp.LpStatus[prob.status] not in ("Optimal","Not Solved","Optimal Infeasible","Infeasible"):
        return None
    if pulp.LpStatus[prob.status] != "Optimal":
        # If not optimal but feasible solution exists, extract anyway
        pass

    # Build resulting AMI vector
    assigned = np.zeros(n, dtype=float)
    for i in I:
        for b in bands:
            if pulp.value(x[i,b]) >= 0.5:
                assigned[i] = b
                break
        if assigned[i] == 0.0:
            return None
    return assigned

# ----------------------------
# Band cap + vertical score + dominance
# ----------------------------

def enforce_max_bands(amis: np.ndarray, sf: np.ndarray, max_bands: int, targets: dict) -> np.ndarray:
    """
    Reduce the number of bands by merging the smallest-SF band into the nearest neighbor
    while keeping the 40%-share and average constraints valid.
    """
    amis = np.array(amis, dtype=float)
    sf = np.array(sf, dtype=float)

    def ok(aa):
        total = sf.sum()
        wavg = float((aa * sf).sum() / total) if total else 0
        # 40% band range if applicable
        if targets.get("pct40_low") is not None and targets.get("pct40_high") is not None:
            forty_share = float(sf[np.isclose(aa,0.40)].sum())/total if total else 0.0
            if not (targets["pct40_low"]-1e-12 <= forty_share <= targets["pct40_high"]+1e-12):
                return False
        # average window if applicable
        if targets.get("avg_low") is not None and targets.get("avg_high") is not None:
            if not (targets["avg_low"]-1e-12 <= wavg <= targets["avg_high"]+1e-12):
                return False
        return True

    def unique_bands(aa):
        return sorted(set([round(x,2) for x in aa.tolist()]))

    while len(unique_bands(amis)) > max_bands:
        bands = unique_bands(amis)
        band_sums = {b: float(sf[np.isclose(amis,b)].sum()) for b in bands}
        bmin = min(bands, key=lambda b: band_sums[b] if band_sums[b]>0 else 1e18)
        neighbors = [b for b in bands if b != bmin]
        target = min(neighbors, key=lambda b: abs(b - bmin))
        merged = amis.copy()
        merged[np.isclose(merged, bmin)] = target
        if ok(merged):
            amis = merged
        else:
            others = sorted(neighbors, key=lambda b: abs(b - bmin))
            merged_ok = False
            for alt in others[1:]:
                merged2 = amis.copy()
                merged2[np.isclose(merged2, bmin)] = alt
                if ok(merged2):
                    amis = merged2; merged_ok=True; break
            if not merged_ok:
                break
    return amis

def vertical_score(floors: np.ndarray, amis: np.ndarray, alpha=0.02, beta=0.02) -> float:
    floors = np.array(floors, dtype=float)
    amis = np.array(amis, dtype=float)
    if floors.size == 0:
        return 0.0
    mask40 = np.isclose(amis, 0.40)
    mask_hi = amis >= 0.70
    avg_40 = float(floors[mask40].mean()) if mask40.any() else 0.0
    avg_hi = float(floors[mask_hi].mean()) if mask_hi.any() else 0.0
    return alpha * avg_hi - beta * avg_40

def _scenario_metrics(assigned: np.ndarray, sf: np.ndarray, floors: np.ndarray) -> Dict[str,float]:
    total = float(sf.sum())
    sf40 = float(sf[np.isclose(assigned,0.40)].sum())
    pct40 = (sf40/total*100.0) if total else 0.0
    wavg = _sf_avg(assigned, sf)*100.0
    vscore = vertical_score(floors, assigned)
    bands_count = len(_unique_bands(assigned))
    return {"aff_sf_total": total, "sf_at_40": sf40, "pct40": pct40, "wavg": wavg,
            "vscore": vscore, "bands_count": bands_count}

def _is_dominated(a: Dict, b: Dict) -> bool:
    """Return True if scenario a is dominated by scenario b."""
    better_or_equal = (
        (b["wavg"] >= a["wavg"] - 1e-9) and
        (b["vscore"] >= a["vscore"] - 1e-9) and
        (b["bands_count"] <= a["bands_count"])
    )
    strictly_better = (
        (b["wavg"] > a["wavg"] + 1e-6) or
        (b["vscore"] > a["vscore"] + 1e-6) or
        (b["bands_count"] < a["bands_count"])
    )
    return better_or_equal and strictly_better

def _pareto_front(named: Dict[str,Dict]) -> Dict[str,Dict]:
    items = list(named.items())
    keep = []
    for i,(ki,vi) in enumerate(items):
        dominated = False
        for j,(kj,vj) in enumerate(items):
            if i==j: continue
            if _is_dominated(vi["metrics_ext"], vj["metrics_ext"]):
                dominated = True; break
        if not dominated:
            keep.append((ki,vi))
    # epsilon filter by top average (keep near-best)
    if keep:
        top = max(v["metrics_ext"]["wavg"] for _,v in keep)
        eps = 0.05  # 0.05% points
        keep = [(k,v) for (k,v) in keep if v["metrics_ext"]["wavg"] >= top - eps]
    return dict(keep)

# ----------------------------
# Scenarios (S1–S3 strict, R1–R3 relaxed) + exact path
# ----------------------------

def _order_for_raise(floors: np.ndarray, sf: np.ndarray) -> List[int]:
    if floors is not None and not np.all(np.isnan(floors)):
        return sorted(range(len(sf)), key=lambda i: ((floors[i] if not np.isnan(floors[i]) else 1e9), -sf[i]))
    return sorted(range(len(sf)), key=lambda i: -sf[i])

def generate_scenarios(
    aff_df: pd.DataFrame,
    low_share=0.20, high_share=0.21,
    avg_low=0.59, avg_high=0.60,
    require_family_at_40: bool = False,
    spread_40_max_per_floor: Optional[int] = None,
    exempt_top_k_floors: int = 0,
    enforce_max_bands_count: int = 3,
    try_exact: bool = True
) -> Dict[str, Dict]:

    sf = aff_df["NET SF"].to_numpy(float)
    floors = aff_df["FLOOR"].to_numpy(float) if "FLOOR" in aff_df.columns else np.full(len(aff_df), np.nan)
    beds = aff_df["BED"].to_numpy(float) if "BED" in aff_df.columns else np.full(len(aff_df), np.nan)
    total=float(sf.sum())

    def postprocess_spread(mask40: np.ndarray, max_per_floor: Optional[int], hi_share: float) -> np.ndarray:
        if max_per_floor is None or np.all(np.isnan(floors)): 
            return mask40
        m = mask40.copy()
        lo = total*low_share; hi = total*hi_share
        s = float(sf[m].sum())
        fl = pd.Series(floors).fillna(-1).astype(int).to_numpy()
        changed=True
        while changed:
            changed=False
            counts={}
            for i,f in enumerate(fl):
                if m[i]: counts[f]=counts.get(f,0)+1
            viol=[f for f,c in counts.items() if c>max_per_floor]
            if not viol: break
            for vf in viol:
                idx_40=[i for i in range(len(m)) if m[i] and fl[i]==vf]
                if not idx_40: continue
                victim=max(idx_40, key=lambda i: sf[i])
                cands=[j for j in range(len(m)) if (not m[j]) and fl[j]!=vf]
                cands.sort(key=lambda j: sf[j])
                for j in cands:
                    cand=s - float(sf[victim]) + float(sf[j])
                    if lo <= cand <= hi:
                        m[victim]=False; m[j]=True; s=cand; changed=True; break
        return m

    def build(mask_pref_high: bool, hi_adj: float, need_family: bool, protect_top: int,
              strict_cap_70: bool = False,
              relaxed_seed: Optional[List[Tuple[float,float]]] = None,
              spread_cap: Optional[int] = None,
              label: str = "S?"):
        hi_share = min(high_share + hi_adj, 0.22)
        # 1) base 40% set
        m40 = choose_40pct_subset_by_sf(sf, floors, low_share, hi_share, mask_pref_high)
        # 2) require ≥1 family (2BR+) at 40% if requested
        if need_family and not np.any(m40 & (beds >= 2)):
            fam_idx = np.where(beds >= 2)[0]
            if len(fam_idx):
                fam = int(fam_idx[np.argmin(sf[fam_idx])])
                if not m40[fam]:
                    lo = total*low_share; hi = total*hi_share
                    chosen = np.where(m40)[0]
                    drop = int(chosen[np.argmax(sf[chosen])]) if len(chosen) else None
                    if drop is not None:
                        cand = float(sf[m40].sum()) - float(sf[drop]) + float(sf[fam])
                        if lo <= cand <= hi:
                            m40[drop]=False; m40[fam]=True
        # 3) protect top floors
        if protect_top and not np.all(np.isnan(floors)):
            top_cut = np.nanmax(floors) - protect_top + 1
            if not np.isnan(top_cut):
                lo = total*low_share; hi = total*hi_share
                top_40 = [i for i in np.where(m40)[0] if floors[i] >= top_cut]
                pool   = [i for i in np.where(~m40)[0] if floors[i] <  top_cut]
                top_40.sort(key=lambda i: -sf[i]); pool.sort(key=lambda i: sf[i])
                s = float(sf[m40].sum())
                for a in top_40:
                    for b in pool:
                        cand = s - float(sf[a]) + float(sf[b])
                        if lo <= cand <= hi:
                            m40[a]=False; m40[b]=True; s=cand; break
        # 4) repair into band
        m40 = ensure_band(m40, sf, low_share, hi_share)
        # 5) spread cap per floor
        m40 = postprocess_spread(m40, spread_cap, hi_share)

        # 6) assignment — exact MILP if possible, else heuristic
        assigned = None
        if try_exact:
            assigned = _try_exact_optimize(
                sf=sf, floors=floors, mask40=m40,
                avg_low=avg_low, avg_high=avg_high,
                low_share=low_share, high_share=hi_share,
                allowed_bands=(0.40,0.60,0.70,0.80,0.90,1.00),
                max_bands=enforce_max_bands_count,
                alpha=0.02, beta=0.02
            )
        if assigned is None:
            assigned = balance_average(
                sf, floors, m40, avg_low, avg_high,
                True, avg_high-1e-6,
                max_tier=(0.70 if strict_cap_70 else 1.00),
                preseed=relaxed_seed
            )
            # enforce band cap post-heuristic
            assigned = enforce_max_bands(assigned, sf, max_bands=enforce_max_bands_count,
                                         targets={"pct40_low":low_share,"pct40_high":hi_share,
                                                  "avg_low":avg_low,"avg_high":avg_high})

        metrics = _scenario_metrics(assigned, sf, floors)
        return {"mask40": m40, "assigned": assigned, "metrics": metrics, "label": label}

    # STRICT set (S1–S3): cap non-40 at 70% AMI to keep solutions tight/near-identical
    S1 = build(mask_pref_high=True,  hi_adj=0.000, need_family=False,                protect_top=0, strict_cap_70=True,  label="S1")
    S2 = build(mask_pref_high=True,  hi_adj=0.000, need_family=True,                 protect_top=0, strict_cap_70=True,  label="S2")
    S3 = build(mask_pref_high=False, hi_adj=0.000, need_family=False,                protect_top=max(exempt_top_k_floors,0),
               strict_cap_70=True, spread_cap=spread_40_max_per_floor, label="S3")

    # RELAXED set (R1–R3): seed some high tiers on premium units, then rebalance
    R1 = build(mask_pref_high=True,  hi_adj=0.001, need_family=False, protect_top=0,
               strict_cap_70=False, relaxed_seed=[(0.80, 0.03)], label="R1")
    R2 = build(mask_pref_high=True,  hi_adj=0.001, need_family=False, protect_top=0,
               strict_cap_70=False, relaxed_seed=[(0.90, 0.02),(0.80,0.03)], label="R2")
    R3 = build(mask_pref_high=True,  hi_adj=0.001, need_family=False, protect_top=0,
               strict_cap_70=False, relaxed_seed=[(1.00, 0.01),(0.90,0.02),(0.80,0.03)], label="R3")

    # Guardrails (compliance)
    out_raw = {"S1": S1, "S2": S2, "S3": S3, "R1": R1, "R2": R2, "R3": R3}
    for k,v in out_raw.items():
        m=v["metrics"]
        if not (20.0 - 1e-6 <= m["pct40"] <= 21.0 + 1e-6):
            raise RuntimeError(f"Scenario {k}: 40% share out of band ({m['pct40']:.3f}%).")
        if not (59.0 - 1e-6 <= m["wavg"] <= 60.0 + 1e-6):
            raise RuntimeError(f"Scenario {k}: Weighted average out of band ({m['wavg']:.3f}%).")
        # add extended metrics used by dominance
        v["metrics_ext"] = dict(m, vscore=vertical_score(floors, v["assigned"]),
                                   bands_count=len(_unique_bands(v["assigned"])))

    # Dominance filtering — keep only Pareto-best (but ensure at least 3 scenarios remain)
    out = _pareto_front(out_raw)
    if len(out) < 3:
        out = out_raw  # fall back to all six (already compliance-safe), still with band cap in place

    return out

# ----------------------------
# Public entry: return scenarios + mirror
# ----------------------------

def allocate_with_scenarios(
    df: pd.DataFrame,
    low_share=0.20, high_share=0.21, avg_low=0.59, avg_high=0.60,
    require_family_at_40: bool = False,
    spread_40_max_per_floor: Optional[int] = None,
    exempt_top_k_floors: int = 0,
):
    d = _normalize_headers(df)
    if "NET SF" not in d.columns:
        raise ValueError("Missing NET SF column after normalization.")

    sel = _selected_for_ami_mask(d)
    if not sel.any():
        raise ValueError("No AMI-selected rows detected. Put ANY value in the AMI column for selected rows (e.g., 0.6, x, ✓).")

    aff = d.loc[sel].copy()
    scen = generate_scenarios(
        aff,
        low_share=low_share, high_share=high_share, avg_low=avg_low, avg_high=avg_high,
        require_family_at_40=require_family_at_40,
        spread_40_max_per_floor=spread_40_max_per_floor,
        exempt_top_k_floors=exempt_top_k_floors,
        enforce_max_bands_count=3,
        try_exact=True
    )

    # Attach scenario assignments to the full table (keys are whatever survive Pareto; stable order)
    labels = list(scen.keys())
    full = d.copy()
    for label in labels:
        full[f"Assigned_AMI_{label}"] = np.nan
        full.loc[sel, f"Assigned_AMI_{label}"] = scen[label]["assigned"]

    # Affordable-only breakdown (one wide table)
    base_cols = [c for c in ["FLOOR","APT","BED","NET SF","AMI_RAW"] if c in d.columns]
    aff_br = d.loc[sel, base_cols].copy()
    for label in labels:
        aff_br[f"Assigned_AMI_{label}"] = scen[label]["assigned"]

    # Choose the “best” scenario to mirror back into the original AMI column:
    def score_ext(m):
        # prioritize tight compliance (already met), then higher overall AMI (up to 60), then vertical score,
        # then fewer bands
        return (m["wavg"]*10.0) + (m.get("vscore",0.0)) - (m.get("bands_count",3)-1)*0.5

    metrics = {k: dict(scen[k]["metrics"],
                       vscore=scen[k]["metrics_ext"]["vscore"],
                       bands_count=scen[k]["metrics_ext"]["bands_count"]) for k in labels}
    best_label = max(labels, key=lambda k: score_ext(metrics[k]))
    best_assigned = scen[best_label]["assigned"]

    # Mirror of original: overwrite AMI for selected rows; append totals
    mirror = d.copy()
    target_col = "AMI" if "AMI" in mirror.columns else "AMI"
    mirror[target_col] = mirror.get(target_col, np.nan)
    mirror.loc[sel, target_col] = best_assigned

    total_sf = float(aff["NET SF"].sum())
    sf40 = float(aff["NET SF"].to_numpy()[np.isclose(best_assigned, 0.40)].sum())
    pct40 = (sf40 / total_sf * 100.0) if total_sf else 0.0
    wavg = _sf_avg(best_assigned, aff["NET SF"].to_numpy()) * 100.0

    footer = pd.DataFrame({
        list(mirror.columns)[0]: ["","Affordable SF total","SF at 40% AMI","% at 40% AMI","Weighted Avg AMI","Best Scenario"],
        list(mirror.columns)[1] if len(mirror.columns)>1 else "Value": ["", f"{total_sf:,.2f}", f"{sf40:,.2f}", f"{pct40:.3f}%", f"{wavg:.3f}%", best_label]
    })
    mirror_out = pd.concat([mirror, footer], ignore_index=True)

    # Export metrics as simple dict of scenario -> metrics (w/ vscore & bands)
    metrics_out = {k: dict(v["metrics"], vscore=v["metrics_ext"]["vscore"], bands_count=v["metrics_ext"]["bands_count"]) for k,v in scen.items()}

    return full, aff_br, metrics_out, mirror_out, best_label
