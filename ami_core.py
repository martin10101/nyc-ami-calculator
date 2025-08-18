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
    # Some client files have this:
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
# Assignment (averages)
# ----------------------------

def _order_for_raise(floors: np.ndarray, sf: np.ndarray) -> List[int]:
    if floors is not None and not np.all(np.isnan(floors)):
        return sorted(range(len(sf)), key=lambda i: ((floors[i] if not np.isnan(floors[i]) else 1e9), -sf[i]))
    return sorted(range(len(sf)), key=lambda i: -sf[i])

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
        """Lower the highest AMI non-40 units first (smallest SF last) until <= max_avg."""
        if cur_avg() <= max_avg + 1e-12:
            return
        raised = [i for i in non40_idx if assigned[i] > 0.60]
        # Lower *largest* AMI first, and within same tier, smallest SF first
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
                if tier > max_tier: continue  # never exceed configured cap
                assigned[i] = tier
                used.add(i)
                acc += float(sf[i])
                if acc >= target_sf - 1e-9: break
        # If that pushed us above 60, clamp back down *before* continuing.
        clamp_down_to(avg_high - 1e-6)

    # ---------- main “raise” loop (never aims above avg_high) ----------
    target_avg = min(target if target is not None else (avg_high - 1e-6), avg_high - 1e-6)

    # Order to raise: lower floors first & larger sf (dev preference)
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
                # Try a +0.1 bump but don't allow resulting avg > avg_high
                old = assigned[i]
                assigned[i] = round(min(old + 0.1, max_tier), 1)
                if cur_avg() > avg_high + 1e-12:
                    assigned[i] = old  # revert if it breaches the cap
                else:
                    progressed = True
        if not progressed:
            break  # cannot raise further without breaching the cap

    # ---------- final safety: place inside [avg_low, avg_high] ----------
    if cur_avg() > avg_high + 1e-12:
        clamp_down_to(avg_high - 1e-6)

    # If still below lower bound, try minimal ups that keep <= avg_high
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
        # Last resort: clamp down if we're 1–3 bps over
        if cur_avg() > avg_high + 1e-12:
            clamp_down_to(avg_high - 1e-6)

    if not (avg_low - 1e-12 <= cur_avg() <= avg_high + 1e-12):
        raise RuntimeError(f"Weighted average out of band after balance ({cur_avg()*100:.3f}%).")

    return assigned


# ----------------------------
# Scenarios (S1–S3 strict, R1–R3 relaxed)
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

    def build(mask_pref_high: bool, hi_adj: float, need_family: bool, protect_top: int,
              strict_cap_70: bool = False,
              relaxed_seed: Optional[List[Tuple[float,float]]] = None,
              spread_cap: Optional[int] = None):
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
        # 6) balance average with caps & seeds
        assigned = balance_average(
            sf, floors, m40, avg_low, avg_high,
            True, avg_high-1e-6,
            max_tier=(0.70 if strict_cap_70 else 1.00),
            preseed=relaxed_seed
        )
        sf40=float(sf[m40].sum())
        return {
            "mask40": m40,
            "assigned": assigned,
            "metrics": {
                "aff_sf_total": total,
                "sf_at_40": sf40,
                "pct40": (sf40/total*100.0) if total else 0.0,
                "wavg": _sf_avg(assigned, sf)*100.0
            }
        }

    # STRICT set (S1–S3): cap non-40 at 70% AMI to keep solutions tight/near-identical
    S1 = build(mask_pref_high=True,  hi_adj=0.000, need_family=False,                protect_top=0, strict_cap_70=True)
    S2 = build(mask_pref_high=True,  hi_adj=0.000, need_family=True,                 protect_top=0, strict_cap_70=True)
    S3 = build(mask_pref_high=False, hi_adj=0.000, need_family=False,                protect_top=max(exempt_top_k_floors,0),
               strict_cap_70=True, spread_cap=spread_40_max_per_floor)

    # RELAXED set (R1–R3): seed some high tiers on premium units, then rebalance
    # shares are of non-40% SF; tune to your taste
    R1 = build(mask_pref_high=True,  hi_adj=0.001, need_family=False, protect_top=0,
               strict_cap_70=False, relaxed_seed=[(0.80, 0.06)])
    R2 = build(mask_pref_high=True,  hi_adj=0.002, need_family=False, protect_top=0,
               strict_cap_70=False, relaxed_seed=[(0.90, 0.03),(0.80,0.05)])
    R3 = build(mask_pref_high=True,  hi_adj=0.002, need_family=False, protect_top=0,
               strict_cap_70=False, relaxed_seed=[(1.00, 0.02),(0.90,0.03),(0.80,0.05)])

    out = {"S1": S1, "S2": S2, "S3": S3, "R1": R1, "R2": R2, "R3": R3}

    # guardrails
    for k,v in out.items():
        m=v["metrics"]
        if not (20.0 - 1e-6 <= m["pct40"] <= 21.0 + 1e-6):
            raise RuntimeError(f"Scenario {k}: 40% share out of band ({m['pct40']:.3f}%).")
        if not (59.0 - 1e-6 <= m["wavg"] <= 60.0 + 1e-6):
            raise RuntimeError(f"Scenario {k}: Weighted average out of band ({m['wavg']:.3f}%).")
    return out

# ----------------------------
# Public entry: return 6 scenarios + mirror
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
    )

    # Attach scenario assignments to the full table (S1..R3)
    full = d.copy()
    labels = list(scen.keys())
    for label in labels:
        full[f"Assigned_AMI_{label}"] = np.nan
        full.loc[sel, f"Assigned_AMI_{label}"] = scen[label]["assigned"]

    # Affordable-only breakdown (one wide table)
    base_cols = [c for c in ["FLOOR","APT","BED","NET SF","AMI_RAW"] if c in d.columns]
    aff_br = d.loc[sel, base_cols].copy()
    for label in labels:
        aff_br[f"Assigned_AMI_{label}"] = scen[label]["assigned"]

    # Choose the “best” scenario to mirror back into the original AMI column:
    def score(m):
        # prioritize compliance tightness, then higher overall AMI (developer revenue proxy)
        closeness = 1000.0 - abs(60.0 - m["wavg"])*50.0 - abs(20.0 - m["pct40"])*100.0
        return closeness + m["wavg"]*0.01  # tiny tie-breaker toward higher revenue
    metrics = {k:v["metrics"] for k,v in scen.items()}
    best_label = max(labels, key=lambda k: score(metrics[k]))
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
        list(mirror.columns)[0]: ["","Affordable SF total","SF at 40% AMI","% at 40% AMI","Weighted Avg AMI"],
        list(mirror.columns)[1] if len(mirror.columns)>1 else "Value": ["", f"{total_sf:,.2f}", f"{sf40:,.2f}", f"{pct40:.3f}%", f"{wavg:.3f}%"]
    })
    mirror_out = pd.concat([mirror, footer], ignore_index=True)

    return full, aff_br, metrics, mirror_out, best_label
