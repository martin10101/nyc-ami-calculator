# top of file
import os
# ...
DEFAULT_MILP_TIMELIMIT = int(os.getenv("MILP_TIMELIMIT_SEC", "10"))

# in _try_exact_optimize(...):
    # Phase 1
    prob1, x1 = build_lp(maximize_wavg=True)
    prob1.solve(pulp.PULP_CBC_CMD(msg=False, timeLimit=DEFAULT_MILP_TIMELIMIT))
    # Phase 2
    prob2, x2 = build_lp(maximize_wavg=False, wavg_floor=wavg_num_opt)
    prob2.solve(pulp.PULP_CBC_CMD(msg=False, timeLimit=DEFAULT_MILP_TIMELIMIT))

# public entry
def allocate_with_scenarios(
    df: pd.DataFrame,
    low_share=0.20, high_share=0.21, avg_low=0.59, avg_high=0.60,
    require_family_at_40: bool = False,
    spread_40_max_per_floor: Optional[int] = None,
    exempt_top_k_floors: int = 0,
    return_top_k: int = 3,
    use_milp: bool = True  # NEW
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
        try_exact=use_milp,                # << toggle here
        return_top_k=return_top_k
    )
    # ... (rest identical to the previous full version you pasted)
# ami_core.py — unified, optimized allocator
# - Any band allowed: {0.40,0.60,0.70,0.80,0.90,1.00}
# - 20–21% SF at 40%, weighted average 59–60% (target 60.00)
# - ≤3 bands per scenario
# - Lexicographic MILP (maximize average, then developer layout)
# - Heuristic fallback if PuLP missing
# - Dedupe scenarios, keep best top_k

from typing import Optional, Dict, List, Tuple
import numpy as np
import pandas as pd

# =========================
# Header normalization
# =========================
FLEX_HEADERS = {
    "NET SF": {"netsf","net sf","net_sf","net s.f.","sf","sqft","sq ft","squarefeet","square feet","netarea","area","area(sf)"},
    "AMI": {"ami","aime","aff ami","assigned_ami","aff_ami","aff-ami"},
    "FLOOR": {"floor","fl","story","level"},
    "APT": {"apt","apartment","unit","unit#","unitno","unit no","apt#","apt no"},
    "BED": {"bed","beds","bedroom","br","bedrooms"},
    "SIGNED_AMI": {"signed_ami","signedami","assignedami"},
    "AFF": {"aff","affordable","selected","ami_selected","is_affordable","target_ami"},
}

def _coerce_percent(x):
    if pd.isna(x): return np.nan
    if isinstance(x,(int,float)):
        v=float(x); return v/100.0 if v>1.0 else v
    s=str(x).strip()
    sU=s.upper().replace("AIME","").replace("AMI","").replace("%","").strip()
    if sU in {"Y","YES","TRUE","1","✓","✔","X"}: return 0.60
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
    if "AMI" in d.columns: d["AMI_RAW"] = d["AMI"]
    if "NET SF" in d.columns: d["NET SF"] = pd.to_numeric(d["NET SF"], errors="coerce")
    if "FLOOR" in d.columns:  d["FLOOR"]  = pd.to_numeric(d["FLOOR"],  errors="coerce")
    if "BED" in d.columns:    d["BED"]    = pd.to_numeric(d["BED"],    errors="coerce")
    if "AMI" in d.columns:    d["AMI"]    = d["AMI"].apply(_coerce_percent)
    return d

def _selected_for_ami_mask(d: pd.DataFrame) -> pd.Series:
    # 1) explicit yes/flag
    for c in d.columns:
        if str(c).strip().lower() in FLEX_HEADERS["AFF"]:
            return d[c].astype(str).str.strip().str.lower().isin(
                {"1","true","yes","y","✓","✔","x"}
            ) & d["NET SF"].notna()
    # 2) signed_ami present
    for c in d.columns:
        if str(c).strip().lower() in FLEX_HEADERS["SIGNED_AMI"]:
            return d[c].astype(str).str.strip().ne("").fillna(False) & d["NET SF"].notna()
    # 3) AMI_RAW near 0.6 or ticked
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
    return pd.Series(False, index=d.index)

# =========================
# Math helpers
# =========================
def _sf_avg(ami: np.ndarray, sf: np.ndarray) -> float:
    return float(np.dot(ami, sf)/sf.sum()) if sf.sum()>0 else 0.0

def _unique_bands(amis: np.ndarray) -> List[float]:
    return sorted({round(float(x),2) for x in amis.tolist()})

# =========================
# 40% selection by SF band
# =========================
def _choose_40_mitm(sf: np.ndarray, low_share=0.20, high_share=0.21) -> np.ndarray:
    import bisect
    n=len(sf); total=float(sf.sum()); lo=total*low_share; hi=total*high_share
    idx=np.arange(n); A=idx[:n//2]; B=idx[n//2:]
    sums_a=[]; sums_b=[]
    for m in range(1<<len(A)):
        s=0.0
        for i in range(len(A)):
            if m&(1<<i): s+=float(sf[A[i]])
        sums_a.append((s,m))
    for m in range(1<<len(B)):
        s=0.0
        for i in range(len(B)):
            if m&(1<<i): s+=float(sf[B[i]])
        sums_b.append((s,m))
    sums_b.sort(key=lambda x:x[0]); bvals=[x[0] for x in sums_b]
    # exact in band
    for sa, ma in sums_a:
        L=bisect.bisect_left(bvals, lo-sa); R=bisect.bisect_right(bvals, hi-sa)
        if L<R:
            mid=(L+R-1)//2; sb, mb = sums_b[mid]
            mask=np.zeros(n,bool)
            for i in range(len(A)):
                if ma&(1<<i): mask[A[i]]=True
            for i in range(len(B)):
                if mb&(1<<i): mask[B[i]]=True
            return mask
    # closest to mid
    best=None; mid=(lo+hi)/2
    for sa,ma in sums_a:
        pos=bisect.bisect_left(bvals, mid-sa)
        for c in (pos-1,pos,pos+1):
            if 0<=c<len(bvals):
                sb,mb=sums_b[c]; tot=sa+sb
                gap=max(0.0, lo-tot, tot-hi)
                if best is None or gap<best[0]: best=(gap,ma,mb)
    if best:
        _,ma,mb=best; mask=np.zeros(n,bool)
        for i in range(len(A)):
            if ma&(1<<i): mask[A[i]]=True
        for i in range(len(B)):
            if mb&(1<<i): mask[B[i]]=True
        return mask
    return np.zeros(n,bool)

def _choose_40_large(sf: np.ndarray, low_share=0.20, high_share=0.21) -> np.ndarray:
    total=float(sf.sum()); lo=total*low_share; hi=total*high_share
    order=np.argsort(sf)  # small→large
    chosen=[]; s=0.0
    for i in order:
        if s<lo:
            chosen.append(i); s+=float(sf[i])
    not_chosen=[i for i in order if i not in chosen]
    improved=True
    while s>hi and improved:
        improved=False
        for c in sorted(chosen, key=lambda k: sf[k], reverse=True):
            for j in not_chosen:
                cand=s - float(sf[c]) + float(sf[j])
                if lo<=cand<=hi:
                    chosen.remove(c); chosen.append(j)
                    not_chosen.remove(j); not_chosen.append(c)
                    s=cand; improved=True; break
            if improved: break
    m=np.zeros(len(sf),bool); m[np.array(chosen,dtype=int)]=True
    return m

def choose_40pct_subset_by_sf(sf: np.ndarray, floors: np.ndarray, low_share=0.20, high_share=0.21) -> np.ndarray:
    return _choose_40_mitm(sf, low_share, high_share) if len(sf)<=34 else _choose_40_large(sf, low_share, high_share)

def ensure_band(mask40: np.ndarray, sf: np.ndarray, low_share: float, high_share: float) -> np.ndarray:
    m = mask40.copy()
    total = float(sf.sum()); lo = total*low_share; hi = total*high_share
    s = float(sf[m].sum())

    if s < lo - 1e-9:
        outsiders = [i for i in np.where(~m)[0]]
        outsiders.sort(key=lambda i: sf[i])
        for j in outsiders:
            cand = s + float(sf[j])
            m[j] = True; s = cand
            if s >= lo - 1e-9: break

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

def push_40_down(m40: np.ndarray, sf: np.ndarray, floors: np.ndarray,
                 low_share: float, high_share: float) -> np.ndarray:
    if floors is None or np.all(np.isnan(floors)): 
        return m40
    total = float(sf.sum()); lo = total*low_share; hi = total*high_share
    s = float(sf[m40].sum())

    chosen = np.where(m40)[0].tolist()
    others = np.where(~m40)[0].tolist()

    chosen.sort(key=lambda i: (floors[i] if not np.isnan(floors[i]) else 9e9), reverse=True)
    others.sort(key=lambda i: (floors[i] if not np.isnan(floors[i]) else -9e9))

    m = m40.copy()
    for c in chosen:
        for j in list(others):
            if np.isnan(floors[c]) or np.isnan(floors[j]) or floors[j] >= floors[c]:
                continue
            cand = s - float(sf[c]) + float(sf[j])
            if lo - 1e-9 <= cand <= hi + 1e-9:
                m[c] = False; m[j] = True; s = cand
                others.remove(j); others.append(c)
                break
    return m

# =========================
# Heuristic assignment
# =========================
def balance_average(sf: np.ndarray, floors: np.ndarray, mask40: np.ndarray,
                    avg_low=0.59, avg_high=0.60, target: Optional[float]=None,
                    max_tier: float=1.0, preseed: Optional[List[Tuple[float,float]]]=None) -> np.ndarray:
    assigned = np.full(len(sf), 0.60, dtype=float)
    assigned[mask40] = 0.40
    non40_idx = np.where(~mask40)[0]

    def cur_avg() -> float: return _sf_avg(assigned, sf)

    # Optional seed (for relaxed runs)
    if preseed:
        order = sorted(
            non40_idx,
            key=lambda i: (-(floors[i] if floors is not None and not np.isnan(floors[i]) else -1), -sf[i])
        )
        non40_sf_total = float(sf[non40_idx].sum())
        used=set()
        for tier, share in preseed:
            target_sf = non40_sf_total * max(0.0, min(share, 1.0))
            acc=0.0
            for i in order:
                if i in used: continue
                if tier > max_tier: continue
                assigned[i] = tier
                used.add(i)
                acc += float(sf[i])
                if acc >= target_sf - 1e-9: break
        # clamp down if overshoot
        if cur_avg() > avg_high + 1e-12:
            raised = [i for i in non40_idx if assigned[i] > 0.60]
            raised.sort(key=lambda i: (assigned[i], sf[i]))
            for i in reversed(raised):
                while assigned[i] > 0.60 + 1e-12 and cur_avg() > avg_high + 1e-12:
                    assigned[i] = round(assigned[i] - 0.1, 1)

    target_avg = avg_high  # aim for true ceiling

    # raise order: lower floors first, then larger SF
    if floors is not None and not np.all(np.isnan(floors)):
        idx = sorted(non40_idx, key=lambda i: ((floors[i] if not np.isnan(floors[i]) else 1e9), -sf[i]))
    else:
        idx = sorted(non40_idx, key=lambda i: -sf[i])

    while cur_avg() < target_avg - 1e-12:
        progressed = False
        for i in idx:
            if cur_avg() >= target_avg - 1e-12: break
            if assigned[i] < 1.0 - 1e-12 and assigned[i] < max_tier - 1e-12:
                old = assigned[i]
                assigned[i] = round(min(old + 0.1, max_tier), 1)
                if cur_avg() > avg_high + 1e-12:
                    assigned[i] = old
                else:
                    progressed = True
        if not progressed: break

    # fine clamp if drifted over
    if cur_avg() > avg_high + 1e-12:
        raised = [i for i in non40_idx if assigned[i] > 0.60]
        raised.sort(key=lambda i: (assigned[i], sf[i]))
        for i in reversed(raised):
            while assigned[i] > 0.60 + 1e-12 and cur_avg() > avg_high + 1e-12:
                assigned[i] = round(assigned[i] - 0.1, 1)

    # if still under low bound, try one more raise pass
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
        raise RuntimeError(f"Weighted average out of band after balance ({cur_avg()*100:.3f}%).")
    return assigned

# =========================
# Lexicographic MILP (PuLP)
# =========================
def _try_exact_optimize(sf: np.ndarray, floors: np.ndarray, mask40: np.ndarray,
                        avg_low: float, avg_high: float,
                        low_share: float, high_share: float,
                        allowed_bands=(0.40,0.60,0.70,0.80,0.90,1.00),
                        max_bands:int=3,
                        alpha:float=0.02, beta:float=0.02, gamma:float=0.001) -> Optional[np.ndarray]:
    try:
        import pulp
    except Exception:
        return None

    n = len(sf)
    I = list(range(n))
    bands = sorted(set(allowed_bands))
    is40 = mask40.copy()
    Floors = np.nan_to_num(floors, nan=0.0)
    total_sf = float(sf.sum())

    def build_lp(maximize_wavg=True, wavg_floor=None):
        prob = pulp.LpProblem("AMI_Assignment", pulp.LpMaximize)
        x = {(i,b): pulp.LpVariable(f"x_{i}_{int(b*100)}", 0, 1, pulp.LpBinary) for i in I for b in bands}
        y = {b: pulp.LpVariable(f"y_{int(b*100)}", 0, 1, pulp.LpBinary) for b in bands}

        for i in I:
            prob += pulp.lpSum(x[i,b] for b in bands) == 1

        for i in I:
            if is40[i]:
                prob += x[i,0.40] == 1
                for b in bands:
                    if b != 0.40: prob += x[i,b] == 0
            else:
                prob += x[i,0.40] == 0

        sf40 = pulp.lpSum(sf[i]*x[i,0.40] for i in I)
        prob += sf40 >= total_sf * low_share - 1e-6
        prob += sf40 <= total_sf * high_share + 1e-6

        wavg_num = pulp.lpSum(sf[i]*pulp.lpSum(b*x[i,b] for b in bands) for i in I)
        prob += wavg_num >= total_sf*avg_low - 1e-6
        prob += wavg_num <= total_sf*avg_high + 1e-6

        for b in bands:
            for i in I:
                prob += x[i,b] <= y[b]
        prob += pulp.lpSum(y[b] for b in bands) <= max_bands

        if maximize_wavg:
            prob += wavg_num  # Phase 1: maximize average
        else:
            if wavg_floor is not None:
                prob += wavg_num >= wavg_floor - 1e-6  # keep max average
            hi_term   = pulp.lpSum(Floors[i]*pulp.lpSum(x[i,b] for b in bands if b>=0.70) for i in I)
            lo40_term = pulp.lpSum(Floors[i]*x[i,0.40] for i in I)
            # tiny preference for putting higher bands on larger SF
            sf_hi = pulp.lpSum(sf[i]*pulp.lpSum(x[i,b] for b in bands if b>=0.70) for i in I)
            denom = max(1, n)
            prob += alpha * (hi_term/denom) - beta * (lo40_term/denom) + gamma * (sf_hi / float(total_sf))
        return prob, x

    # Phase 1 — max average
    prob1, x1 = build_lp(maximize_wavg=True)
    prob1.solve(pulp.PULP_CBC_CMD(msg=False, timeLimit=12))
    if pulp.LpStatus[prob1.status] not in ("Optimal","Not Solved","Optimal Infeasible","Infeasible"):
        return None
    wavg_num_opt = float(pulp.value(prob1.objective))

    # Phase 2 — keep average, optimize layout
    prob2, x2 = build_lp(maximize_wavg=False, wavg_floor=wavg_num_opt)
    prob2.solve(pulp.PULP_CBC_CMD(msg=False, timeLimit=12))
    if pulp.LpStatus[prob2.status] not in ("Optimal","Not Solved","Optimal Infeasible","Infeasible"):
        return None

    assigned = np.zeros(n, dtype=float)
    for i in I:
        for b in bands:
            if pulp.value(x2[i,b]) >= 0.5:
                assigned[i] = b; break
        if assigned[i] == 0.0: return None
    return assigned

# =========================
# Scenario scoring & pruning
# =========================
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

def _bands_count(arr): return len(_unique_bands(np.asarray(arr)))

def _scenario_score(metrics, vscore, bands_count):
    score = metrics["wavg"] * 20.0  # heavy weight on average
    score += vscore * 1.0
    score -= max(0, bands_count - 3) * 2.0
    return score

def _fingerprint(assigned: np.ndarray, mask40: np.ndarray) -> str:
    bands = sorted({round(float(x),2) for x in assigned.tolist()})
    hist = tuple(int(np.isclose(assigned, b).sum()) for b in bands)
    fp = (tuple(np.where(mask40)[0].tolist()), tuple(bands), hist,
          round(float((assigned*1.0).mean()),6))
    return str(fp)

def dedupe_scenarios(scenarios: dict) -> dict:
    seen=set(); out={}
    for name, blob in scenarios.items():
        fp = _fingerprint(blob["assigned"], blob["mask40"])
        if fp in seen: continue
        seen.add(fp); out[name]=blob
    return out

def prune_to_top_k(scenarios: dict, top_k: int = 3) -> dict:
    if not scenarios: return scenarios
    scored=[]
    for name, blob in scenarios.items():
        m = blob["metrics"]; v=blob.get("vscore",0.0); b=blob.get("bands_count",3)
        scored.append(( _scenario_score(m,v,b), name ))
    scored.sort(reverse=True)
    keep={n:scenarios[n] for _,n in scored[:max(1, top_k)]}
    return keep

# =========================
# Scenario generation
# =========================
def generate_scenarios(aff_df: pd.DataFrame,
                       low_share=0.20, high_share=0.21,
                       avg_low=0.59, avg_high=0.60,
                       require_family_at_40: bool = False,
                       spread_40_max_per_floor: Optional[int] = None,
                       exempt_top_k_floors: int = 0,
                       try_exact: bool = True,
                       return_top_k: int = 3) -> Dict[str, Dict]:
    sf = aff_df["NET SF"].to_numpy(float)
    floors = aff_df["FLOOR"].to_numpy(float) if "FLOOR" in aff_df.columns else np.full(len(aff_df), np.nan)
    beds = aff_df["BED"].to_numpy(float) if "BED" in aff_df.columns else np.full(len(aff_df), np.nan)
    total=float(sf.sum())

    def post_spread(mask40: np.ndarray) -> np.ndarray:
        if spread_40_max_per_floor is None or np.all(np.isnan(floors)): 
            return mask40
        m = mask40.copy()
        lo = total*low_share; hi = total*high_share
        s = float(sf[m].sum())
        fl = pd.Series(floors).fillna(-1).astype(int).to_numpy()
        changed=True
        while changed:
            changed=False
            counts={}
            for i,f in enumerate(fl):
                if m[i]: counts[f]=counts.get(f,0)+1
            viol=[f for f,c in counts.items() if c>spread_40_max_per_floor]
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

    def build(label:str, need_family: bool, protect_top:int):
        # 1) select 40% SF band
        m40 = choose_40pct_subset_by_sf(sf, floors, low_share, high_share)
        # 2) require ≥1 family (2BR+) at 40%
        if need_family and not np.any(m40 & (beds >= 2)):
            fam_idx = np.where(beds >= 2)[0]
            if len(fam_idx):
                fam = int(fam_idx[np.argmin(sf[fam_idx])])
                if not m40[fam]:
                    lo = total*low_share; hi = total*high_share
                    chosen = np.where(m40)[0]
                    if len(chosen):
                        drop = int(chosen[np.argmax(sf[chosen])])
                        cand = float(sf[m40].sum()) - float(sf[drop]) + float(sf[fam])
                        if lo <= cand <= hi: m40[drop]=False; m40[fam]=True
        # 3) protect top K floors from 40%
        if protect_top and not np.all(np.isnan(floors)):
            top_cut = np.nanmax(floors) - protect_top + 1
            if not np.isnan(top_cut):
                lo = total*low_share; hi = total*high_share
                top_40 = [i for i in np.where(m40)[0] if floors[i] >= top_cut]
                pool   = [i for i in np.where(~m40)[0] if floors[i] <  top_cut]
                top_40.sort(key=lambda i: -sf[i]); pool.sort(key=lambda i: sf[i])
                s = float(sf[m40].sum())
                for a in top_40:
                    for b in pool:
                        cand = s - float(sf[a]) + float(sf[b])
                        if lo <= cand <= hi:
                            m40[a]=False; m40[b]=True; s=cand; break
        # 4) repair band, spread cap, push down
        m40 = ensure_band(m40, sf, low_share, high_share)
        m40 = post_spread(m40)
        m40 = push_40_down(m40, sf, floors, low_share, high_share)

        # 5) assign bands — exact first, then heuristic
        assigned=None
        if try_exact:
            assigned = _try_exact_optimize(
                sf=sf, floors=floors, mask40=m40,
                avg_low=avg_low, avg_high=avg_high,
                low_share=low_share, high_share=high_share,
                allowed_bands=(0.40,0.60,0.70,0.80,0.90,1.00),
                max_bands=3, alpha=0.02, beta=0.02, gamma=0.001
            )
        if assigned is None:
            assigned = balance_average(
                sf, floors, m40, avg_low, avg_high,
                target=avg_high, max_tier=1.00, preseed=None
            )
            # enforce ≤3 bands greedily by merging smallest band if needed
            assigned = enforce_max_bands(
                assigned, sf, max_bands=3,
                targets={"pct40_low":low_share,"pct40_high":high_share,"avg_low":avg_low,"avg_high":avg_high}
            )

        # metrics
        total_sf = float(sf.sum()); sf40 = float(sf[np.isclose(assigned,0.40)].sum())
        pct40 = (sf40/total_sf*100.0) if total_sf else 0.0
        wavg = _sf_avg(assigned, sf)*100.0
        vscore = vertical_score(floors, assigned)
        bands_count = len(_unique_bands(assigned))
        blob = {"label": label, "mask40": m40, "assigned": assigned,
                "metrics": {"aff_sf_total": total_sf, "sf_at_40": sf40, "pct40": pct40, "wavg": wavg},
                "vscore": vscore, "bands_count": bands_count}
        # compliance guard
        if not (20.0-1e-6 <= pct40 <= 21.0+1e-6):
            raise RuntimeError(f"Scenario {label}: 40% share out of band ({pct40:.3f}%).")
        if not (59.0-1e-6 <= wavg <= 60.0+1e-6):
            raise RuntimeError(f"Scenario {label}: Weighted average out of band ({wavg:.3f}%).")
        return blob

    # Three distinct variants (policies differ only by the 40% family/top-floor rules)
    S1 = build("S1", need_family=False, protect_top=0)
    S2 = build("S2", need_family=True,  protect_top=0)
    S3 = build("S3", need_family=False, protect_top=max(exempt_top_k_floors,0))

    out_raw = {"S1": S1, "S2": S2, "S3": S3}
    out_raw = dedupe_scenarios(out_raw)
    out_top = prune_to_top_k(out_raw, top_k=max(1, return_top_k))
    if not out_top: out_top = out_raw
    return out_top

def enforce_max_bands(amis: np.ndarray, sf: np.ndarray, max_bands: int, targets: dict) -> np.ndarray:
    amis = np.array(amis, dtype=float)
    sf = np.array(sf, dtype=float)

    def ok(aa):
        total = sf.sum()
        wavg = float((aa * sf).sum() / total) if total else 0
        forty_share = float(sf[np.isclose(aa,0.40)].sum())/total if total else 0.0
        return (targets["pct40_low"]-1e-12 <= forty_share <= targets["pct40_high"]+1e-12) and \
               (targets["avg_low"]-1e-12  <= wavg       <= targets["avg_high"]+1e-12)

    def unique_bands(aa): return sorted(set([round(x,2) for x in aa.tolist()]))

    while len(unique_bands(amis)) > max_bands:
        bands = unique_bands(amis)
        band_sums = {b: float(sf[np.isclose(amis,b)].sum()) for b in bands}
        bmin = min(bands, key=lambda b: band_sums[b] if band_sums[b]>0 else 1e18)
        neighbors = [b for b in bands if b != bmin]
        target = min(neighbors, key=lambda b: abs(b - bmin))
        merged = amis.copy(); merged[np.isclose(merged, bmin)] = target
        if ok(merged):
            amis = merged
        else:
            others = sorted(neighbors, key=lambda b: abs(b - bmin))
            merged_ok=False
            for alt in others[1:]:
                merged2 = amis.copy(); merged2[np.isclose(merged2, bmin)] = alt
                if ok(merged2): amis=merged2; merged_ok=True; break
            if not merged_ok: break
    return amis

# =========================
# Public entry
# =========================
def allocate_with_scenarios(
    df: pd.DataFrame,
    low_share=0.20, high_share=0.21, avg_low=0.59, avg_high=0.60,
    require_family_at_40: bool = False,
    spread_40_max_per_floor: Optional[int] = None,
    exempt_top_k_floors: int = 0,
    return_top_k: int = 3  # set 2 if you only want two scenarios
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
        try_exact=True, return_top_k=return_top_k
    )

    labels = list(scen.keys())
    full = d.copy()
    for label in labels:
        full[f"Assigned_AMI_{label}"] = np.nan
        full.loc[sel, f"Assigned_AMI_{label}"] = scen[label]["assigned"]

    base_cols = [c for c in ["FLOOR","APT","BED","NET SF","AMI_RAW"] if c in d.columns]
    aff_br = d.loc[sel, base_cols].copy()
    for label in labels:
        aff_br[f"Assigned_AMI_{label}"] = scen[label]["assigned"]

    def score_ext(m, vs, bc): return (m["wavg"]*20.0) + vs - max(0, bc-3)*2.0
    metrics = {k: dict(scen[k]["metrics"]) for k in labels}
    best_label = max(labels, key=lambda k: score_ext(metrics[k], scen[k]["vscore"], scen[k]["bands_count"]))
    best_assigned = scen[best_label]["assigned"]

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

    metrics_out = {k: scen[k]["metrics"] for k in labels}
    return full, aff_br, metrics_out, mirror_out, best_label
