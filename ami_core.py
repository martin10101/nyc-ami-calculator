import numpy as np
import pandas as pd
from typing import Optional, Dict

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
}

def _coerce_percent(x):
    """Turn things like 0.6 / 60 / '60%' / 'x' / 'yes' into a float 0..1 or NaN.
       We do NOT trust the value for assignment; it only helps selection if needed."""
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

    # keep original AMI cell for selection, overwrite AMI with normalized numeric (optional)
    if "AMI" in d.columns: d["AMI_RAW"] = d["AMI"]
    if "NET SF" in d.columns: d["NET SF"] = pd.to_numeric(d["NET SF"], errors="coerce")
    if "FLOOR" in d.columns: d["FLOOR"] = pd.to_numeric(d["FLOOR"], errors="coerce")
    if "BED" in d.columns: d["BED"] = pd.to_numeric(d["BED"], errors="coerce")
    if "AMI" in d.columns: d["AMI"] = d["AMI"].apply(_coerce_percent)
    return d

def _selected_for_ami_mask(d: pd.DataFrame) -> pd.Series:
    """Row is selected if:
       - there is an explicit Affordable/Selected flag (yes/true/1/x/✓), OR
       - the AMI cell is non-empty (any text/number)."""
    for c in d.columns:
        key=str(c).strip().lower()
        if key in {"aff","affordable","selected","ami_selected","is_affordable","target_ami"}:
            return d[c].astype(str).str.strip().str.lower().isin({"1","true","yes","y","✓","✔","x"}) & d["NET SF"].notna()

    if "AMI_RAW" in d.columns:
        return d["AMI_RAW"].astype(str).str.strip().ne("").fillna(False) & d["NET SF"].notna()

    if "AMI" in d.columns:
        return d["AMI"].notna() & d["NET SF"].notna()

    return pd.Series(False, index=d.index)

# ----------------------------
# Math helpers
# ----------------------------

def _sf_avg(ami: np.ndarray, sf: np.ndarray) -> float:
    return float(np.dot(ami, sf)/sf.sum()) if sf.sum()>0 else 0.0

# 40% picker (MITM for small N, greedy+swaps for large)
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
    # closest to middle
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
    # pairwise swaps to pull back if s > hi
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

    # optional bias: push 40% toward higher floors while staying inside band
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

    # Too high → try swaps, then trim largest while staying ≥ lo
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

def balance_average(
    sf: np.ndarray,
    floors: np.ndarray,
    mask40: np.ndarray,
    avg_low=0.59, avg_high=0.60,
    prefer_bottom: bool=True,
    target: Optional[float]=None
) -> np.ndarray:
    """Assign 60→70→80→90→100 on non-40% units to reach avg in [avg_low, avg_high]."""
    assigned=np.full(len(sf),0.60,dtype=float)
    assigned[mask40]=0.40
    tgt = (target if target is not None else (avg_high-1e-6))
    def cur(): return _sf_avg(assigned, sf)

    idx=np.where(~mask40)[0].tolist()
    if floors is not None:
        # prefer bottom floors (and bigger SF) for higher AMI to favor revenue
        idx.sort(key=lambda i: ((floors[i] if not np.isnan(floors[i]) else 1e9), -sf[i]))
    else:
        idx.sort(key=lambda i: -sf[i])

    while cur() < tgt:
        progressed=False
        for i in idx:
            if cur() >= tgt: break
            if assigned[i] < 1.0 - 1e-12:
                assigned[i]=round(assigned[i]+0.1,1); progressed=True
        if not progressed: break

    # If we overshot, gently lower the smallest-SF ones first
    if cur() > avg_high + 1e-12:
        raised=[i for i in idx if assigned[i]>0.60]
        raised.sort(key=lambda i: sf[i])
        for i in raised:
            old=assigned[i]; assigned[i]=round(assigned[i]-0.1,1)
            if avg_low <= cur() <= avg_high: break
            assigned[i]=old

    if not (avg_low - 1e-12 <= cur() <= avg_high + 1e-12):
        raise RuntimeError(f"Weighted average out of band after balance ({cur()*100:.3f}%).")
    return assigned

# ----------------------------
# Scenarios + public API
# ----------------------------

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

    def pack(mask40, assigned):
        sf40=float(sf[mask40].sum())
        return {
            "mask40": mask40,
            "assigned": assigned,
            "metrics": {
                "aff_sf_total": total,
                "sf_at_40": sf40,
                "pct40": (sf40/total*100.0) if total else 0.0,
                "wavg": _sf_avg(assigned, sf)*100.0
            }
        }

    def build(prefer_high_for_40: bool, shallow_extra: bool, protect_top: int, need_family: bool):
        hi_share = min(high_share + (0.002 if shallow_extra else 0.0), 0.22)

        # 1) pick a 40% set
        m40 = choose_40pct_subset_by_sf(sf, floors, low_share, hi_share, prefer_high_for_40)

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

        # 3) protect top floors (best effort)
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

        # 4) **repair** back into band (critical fix)
        m40 = ensure_band(m40, sf, low_share, hi_share)

        # 5) optional spread cap; stays in-band by using hi_share inside
        m40 = postprocess_spread(m40, spread_40_max_per_floor, hi_share)

        # 6) balance to 59–60%
        assigned = balance_average(sf, floors, m40, avg_low, avg_high, True, avg_high-1e-6)
        return pack(m40, assigned)

    A = build(prefer_high_for_40=True,  shallow_extra=False, protect_top=0,                          need_family=False)
    B = build(prefer_high_for_40=True,  shallow_extra=False, protect_top=0,                          need_family=require_family_at_40)
    C = build(prefer_high_for_40=False, shallow_extra=False, protect_top=max(exempt_top_k_floors,0), need_family=False)

    out = {"A": A, "B": B, "C": C}

    # hard guardrails
    for k,v in out.items():
        m=v["metrics"]
        if not (20.0 - 1e-6 <= m["pct40"] <= 21.0 + 1e-6):
            raise RuntimeError(f"Scenario {k}: 40% share out of band ({m['pct40']:.3f}%).")
        if not (59.0 - 1e-6 <= m["wavg"] <= 60.0 + 1e-6):
            raise RuntimeError(f"Scenario {k}: Weighted average out of band ({m['wavg']:.3f}%).")
    return out

def allocate_with_scenarios(
    df: pd.DataFrame,
    low_share=0.20, high_share=0.21, avg_low=0.59, avg_high=0.60,
    require_family_at_40: bool = False,
    spread_40_max_per_floor: Optional[int] = None,
    exempt_top_k_floors: int = 0,
):
    """Main entry: normalize, select affordable rows, build scenarios, and return
       (full-table assignments, affordable breakdown, metrics, mirror_of_original, best_label)."""
    df = _normalize_headers(df)
    if "NET SF" not in df.columns:
        raise ValueError("Missing NET SF column after normalization.")

    sel = _selected_for_ami_mask(df)
    if not sel.any():
        raise ValueError("No AMI-selected rows detected. Put ANY value in the AMI column for selected rows (e.g., 0.6, x, ✓).")

    aff = df.loc[sel].copy()
    scen = generate_scenarios(
        aff,
        low_share=low_share, high_share=high_share, avg_low=avg_low, avg_high=avg_high,
        require_family_at_40=require_family_at_40,
        spread_40_max_per_floor=spread_40_max_per_floor,
        exempt_top_k_floors=exempt_top_k_floors,
    )

    # Attach scenario assignments to the full table (A/B/C columns)
    full = df.copy()
    for label in ["A","B","C"]:
        full[f"Assigned_AMI_{label}"] = np.nan
        full.loc[sel, f"Assigned_AMI_{label}"] = scen[label]["assigned"]

    # Affordable-only breakdown
    base_cols = [c for c in ["FLOOR","APT","BED","NET SF","AMI_RAW"] if c in df.columns]
    aff_br = df.loc[sel, base_cols].copy()
    for label in ["A","B","C"]:
        aff_br[f"Assigned_AMI_{label}"] = scen[label]["assigned"]

    # Pick best scenario (closest to 20% and 60%)
    def score(m):
        wavg = min(m["wavg"], 60.0)
        pct  = max(min(m["pct40"], 21.0), 20.0)
        return 1000.0 - abs(60.0 - wavg)*50.0 - abs(20.0 - pct)*100.0
    metrics = {k:v["metrics"] for k,v in scen.items()}
    best_label = max(["A","B","C"], key=lambda k: score(metrics[k]))
    best_assigned = scen[best_label]["assigned"]

    # Mirror of original: overwrite AMI for selected rows; append totals block
    mirror = df.copy()
    target_col = "AMI" if "AMI" in mirror.columns else "AMI"
    mirror[target_col] = mirror.get(target_col, np.nan)
    mirror.loc[sel, target_col] = best_assigned

    total_sf = float(aff["NET SF"].sum())
    sf40 = float(aff["NET SF"].to_numpy()[np.isclose(best_assigned, 0.40)].sum())
    pct40 = (sf40 / total_sf * 100.0) if total_sf else 0.0
    wavg = _sf_avg(best_assigned, aff["NET SF"].to_numpy()) * 100.0

    footer = pd.DataFrame({
        list(mirror.columns)[0]: ["","Affordable SF total","SF at 40% AMI","% at 40% AMI","Weighted Avg AMI"],
        list(mirror.columns)[1] if len(mirror.columns)>1 else "Value": ["", f"{total_sf:,.2f}", f"{sf40:,.2f}", f"{pct40:.3f}%", f"{wavg:.3f}%"]
    })
    mirror_out = pd.concat([mirror, footer], ignore_index=True)

    return full, aff_br, metrics, mirror_out, best_label
