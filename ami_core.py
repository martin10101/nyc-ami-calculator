# ami_core.py — unified, optimized allocator
# - Any band allowed: {0.40,0.60,0.70,0.80,0.90,1.00}
# - 20–21% SF at 40%, weighted average 59–60% (target 60.00)
# - ≤3 bands per scenario
# - Lexicographic MILP (maximize average, then developer layout)
# - Heuristic fallback if PuLP missing
# - Dedupe scenarios, keep best top_k

import os
from typing import Optional, Dict, List, Tuple
import numpy as np
import pandas as pd

DEFAULT_MILP_TIMELIMIT = int(os.getenv("MILP_TIMELIMIT_SEC", "10"))

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
    """Convert various AMI representations to decimal format"""
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
    """Normalize column headers to standard format"""
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
    """
    Determine which rows are selected by the client for AMI assignment.
    
    Key fix: ANY non-blank value in AMI column means unit is selected for affordable housing.
    This includes numbers like 0.6, 0.7, text like "Yes", "✓", "X", etc.
    """
    # 1) explicit yes/flag (affordable indicator)
    for c in d.columns:
        if str(c).strip().lower() in FLEX_HEADERS["AFF"]:
            return d[c].astype(str).str.strip().str.lower().isin(
                {"1","true","yes","y","✓","✔","x"}
            ) & d["NET SF"].notna()
    
    # 2) dedicated signed_ami: any non-empty value signals selection
    for c in d.columns:
        if str(c).strip().lower() in FLEX_HEADERS["SIGNED_AMI"]:
            return d[c].astype(str).str.strip().ne("").fillna(False) & d["NET SF"].notna()
    
    # 3) Use AMI column directly as selection flag: any non-empty value marks selection.
    if "AMI" in d.columns:
        return d["AMI"].apply(lambda x: (not pd.isna(x)) and (str(x).strip() != "")).fillna(False) & d["NET SF"].notna()
    
    # 4) Legacy fallback: AMI_RAW near 0.6 or ticked
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
    
    # Default: none selected
    return pd.Series(False, index=d.index)

# =========================
# Math helpers
# =========================
def _sf_avg(ami: np.ndarray, sf: np.ndarray) -> float:
    """Calculate SF-weighted average AMI"""
    return float(np.dot(ami, sf)/sf.sum()) if sf.sum()>0 else 0.0

def _adjust_wavg_down(assigned: np.ndarray, sf: np.ndarray, avg_high: float) -> np.ndarray:
    """
    Reduce the weighted average AMI by incrementally lowering the highest bands until
    the overall SF-weighted average does not exceed avg_high.
    """
    bands_allowed = [0.40, 0.60, 0.70, 0.80, 0.90, 1.00]
    cur_avg = _sf_avg(assigned, sf)
    
    # Iterate until the average is within bounds or no further reductions are possible
    while cur_avg > avg_high + 1e-6:
        # sort bands present in descending order
        for band in sorted(set(assigned), reverse=True):
            if band <= 0.60:
                continue  # don't lower 60% or 40% bands further
            idxs = np.where(np.isclose(assigned, band))[0]
            for idx in idxs:
                # find the next lower allowed band below current
                lower_candidates = [b for b in bands_allowed if b < band]
                if not lower_candidates:
                    continue
                # choose the closest lower band
                new_band = max(lower_candidates)
                assigned[idx] = new_band
                cur_avg = _sf_avg(assigned, sf)
                # If we've achieved the target, break out
                if cur_avg <= avg_high + 1e-6:
                    return assigned
        # if we exit the loop without reducing the average, break to avoid infinite loop
        break
    return assigned

def _unique_bands(amis: np.ndarray) -> List[float]:
    """Get unique AMI bands used"""
    return sorted({round(float(x),2) for x in amis.tolist()})

# =========================
# 40% selection by SF band
# =========================
def _choose_40_mitm(sf: np.ndarray, low_share=0.20, high_share=0.21) -> np.ndarray:
    """
    Choose units for 40% AMI using meet-in-the-middle algorithm to get 20-21% of SF
    """
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
    
    sums_b.sort(key=lambda x: x[0])
    best_diff = float('inf')
    best_mask = 0
    
    for sa, ma in sums_a:
        target_lo = lo - sa
        target_hi = hi - sa
        # Find the closest sum in B that gets us within [lo, hi]
        i = bisect.bisect_left(sums_b, (target_lo, ))
        while i < len(sums_b):
            sb, mb = sums_b[i]
            if sb > target_hi: break
            total_s = sa + sb
            diff = abs(total_s - (lo + hi)/2)  # Prefer midpoint
            if diff < best_diff:
                best_diff = diff
                best_mask = ma | (mb << len(A))
            i += 1
    
    mask = np.zeros(n, dtype=bool)
    for i in range(n):
        if best_mask & (1 << i): mask[i] = True
    return mask

# =========================
# Exact optimization (MILP)
# =========================
def _try_exact_optimize(sf: np.ndarray, floors: np.ndarray, beds: np.ndarray, low_share=0.20, high_share=0.21,
                        avg_low=0.59, avg_high=0.60, require_family_at_40=False, spread_40_max_per_floor=None,
                        exempt_top_k_floors=0):
    """Try exact MILP optimization"""
    try:
        import pulp
    except ImportError:
        return None  # Fallback to heuristic if PuLP not available

    n = len(sf)
    bands = [0.40, 0.60, 0.70, 0.80, 0.90, 1.00]
    total_sf = sf.sum()

    # Build LP model
    def build_lp(maximize_wavg=True, wavg_floor=None):
        prob = pulp.LpProblem("AMI_Allocation", pulp.LpMaximize)
        x = pulp.LpVariable.dicts("x", ((i, j) for i in range(n) for j in range(len(bands))),
                                  cat='Binary')
        
        # Each unit assigned to exactly one band
        for i in range(n):
            prob += pulp.lpSum(x[i, j] for j in range(len(bands))) == 1
        
        # 20-21% SF at 40%
        sf_40 = pulp.lpSum(sf[i] * x[i, 0] for i in range(n))
        prob += sf_40 >= total_sf * low_share
        prob += sf_40 <= total_sf * high_share
        
        # Weighted average 59-60%
        wavg = pulp.lpSum(bands[j] * sf[i] * x[i, j] for i in range(n) for j in range(len(bands))) / total_sf
        prob += wavg >= avg_low
        prob += wavg <= avg_high
        
        if maximize_wavg:
            prob += wavg
        else:
            # Secondary objective: minimize floor numbers for 40% AMI units
            # This places 40% units on LOWER floors (corrected from original)
            prob += -pulp.lpSum(floors[i] * x[i, 0] for i in range(n))
            if wavg_floor is not None:
                prob += wavg >= wavg_floor
        
        # Optional constraints
        if require_family_at_40:
            prob += pulp.lpSum(x[i, 0] for i in range(n) if beds[i] >= 2) >= 1
        
        unique_floors = np.unique(floors)
        top_floors = sorted(unique_floors)[-exempt_top_k_floors:] if exempt_top_k_floors > 0 else []
        for f in unique_floors:
            floor_units = [i for i in range(n) if floors[i] == f]
            if spread_40_max_per_floor is not None and f not in top_floors:
                prob += pulp.lpSum(x[i, 0] for i in floor_units) <= spread_40_max_per_floor
        
        return prob, x

    # Phase 1: Maximize wavg within bounds
    prob1, x1 = build_lp(maximize_wavg=True)
    prob1.solve(pulp.PULP_CBC_CMD(msg=False, timeLimit=DEFAULT_MILP_TIMELIMIT))
    if pulp.LpStatus[prob1.status] != 'Optimal':
        return None
    wavg_opt = pulp.value(prob1.objective)

    # Phase 2: Minimize floor numbers for 40% units given wavg >= optimal
    prob2, x2 = build_lp(maximize_wavg=False, wavg_floor=wavg_opt)
    prob2.solve(pulp.PULP_CBC_CMD(msg=False, timeLimit=DEFAULT_MILP_TIMELIMIT))
    if pulp.LpStatus[prob2.status] != 'Optimal':
        return None

    assigned = np.array([bands[j] for i in range(n) for j in range(len(bands)) if pulp.value(x2[i, j]) == 1])
    return assigned

# =========================
# Heuristic fallback
# =========================
def _heuristic_assign(sf: np.ndarray, mask40: np.ndarray, target_avg=0.60, max_bands=3, floors: Optional[np.ndarray] = None) -> np.ndarray:
    """
    Heuristic assignment when the exact MILP is unavailable.
    """
    assigned = np.full(len(sf), 0.60)
    assigned[mask40] = 0.40
    current_avg = _sf_avg(assigned, sf)

    bands_high = [0.70, 0.80, 0.90, 1.00]
    remaining = ~mask40
    rem_idx = np.where(remaining)[0]
    
    # Prefer assigning higher bands to higher floors if floor data available
    if floors is not None and len(floors) == len(sf):
        rem_idx = rem_idx[np.argsort(floors[rem_idx])][::-1]
    else:
        np.random.shuffle(rem_idx)

    while current_avg < target_avg and len(bands_high) > 0:
        for i in rem_idx:
            if current_avg >= target_avg:
                break
            # choose band that best moves the average towards the target
            assigned[i] = min(bands_high, key=lambda b: abs((current_avg * sf.sum() - assigned[i] * sf[i] + b * sf[i]) / sf.sum() - target_avg))
            current_avg = _sf_avg(assigned, sf)

    return assigned

# =========================
# Vertical score
# =========================
def vertical_score(floors: np.ndarray, assigned: np.ndarray):
    """Calculate vertical score - higher score means 40% units on higher floors (bad)"""
    forty = assigned == 0.40
    if not np.any(forty):
        return 0
    weighted = np.sum(floors[forty]) / np.sum(forty)
    return weighted

# =========================
# Band enforcement
# =========================
def enforce_max_bands(assigned: np.ndarray, sf: np.ndarray, max_bands: int = 3, 
                     targets: Dict[str, float] = None) -> np.ndarray:
    """Enforce maximum number of bands constraint"""
    if targets is None:
        targets = {"pct40_low": 0.20, "pct40_high": 0.21, "avg_low": 0.59, "avg_high": 0.60}
    
    unique_bands = _unique_bands(assigned)
    if len(unique_bands) <= max_bands:
        return assigned
    
    # If too many bands, consolidate by merging similar bands
    while len(_unique_bands(assigned)) > max_bands:
        bands = _unique_bands(assigned)
        # Find two closest bands to merge
        min_diff = float('inf')
        merge_from, merge_to = None, None
        for i in range(len(bands)):
            for j in range(i+1, len(bands)):
                diff = abs(bands[i] - bands[j])
                if diff < min_diff:
                    min_diff = diff
                    merge_from, merge_to = bands[i], bands[j]
        
        # Merge the closer band to the farther one
        if merge_from and merge_to:
            assigned[np.isclose(assigned, merge_from)] = merge_to
    
    return assigned

# =========================
# Scenario deduplication
# =========================
def dedupe_scenarios(scenarios: List[Dict]) -> List[Dict]:
    """Remove duplicate scenarios"""
    unique = []
    for scenario in scenarios:
        is_duplicate = False
        for existing in unique:
            # Compare assigned arrays
            if np.allclose(scenario["assigned"], existing["assigned"], atol=1e-6):
                is_duplicate = True
                break
        if not is_duplicate:
            unique.append(scenario)
    return unique

def prune_to_top_k(scenarios: List[Dict], k: int = 3) -> List[Dict]:
    """Keep only top k scenarios based on score"""
    def score_scenario(s):
        m = s["metrics"]
        vs = s["vscore"]
        bc = s["bands_count"]
        # Higher weighted average AMI is better; lower vscore (lower floors for 40%) is better
        return (m["wavg"] * 20.0) - vs - max(0, bc - 3) * 2.0
    
    scenarios.sort(key=score_scenario, reverse=True)
    return scenarios[:k]

# =========================
# Generate scenarios
# =========================
def generate_scenarios(aff: pd.DataFrame, low_share=0.20, high_share=0.21, avg_low=0.59, avg_high=0.60,
                       require_family_at_40=False, spread_40_max_per_floor=None, exempt_top_k_floors=0,
                       try_exact=True, return_top_k=3):
    """Generate multiple AMI allocation scenarios"""
    sf = aff["NET SF"].to_numpy()
    floors = aff.get("FLOOR", np.zeros(len(sf))).to_numpy()
    beds = aff.get("BED", np.zeros(len(sf))).to_numpy()

    def build_scenario(label, need_family, protect_top):
        # Select 40% mask
        m40 = _choose_40_mitm(sf, low_share, high_share)
        if need_family:
            # Ensure at least one 2BR+ in 40%
            if not np.any((beds >= 2) & m40):
                family_idx = np.where(beds >= 2)[0]
                if len(family_idx) == 0:
                    raise ValueError("No family units available for 40% requirement.")
                # Swap a non-family in m40 with a family out
                non_family_in = np.where(~(beds >= 2) & m40)[0]
                if len(non_family_in) > 0:
                    swap_in = np.random.choice(family_idx[~m40[family_idx]])
                    swap_out = np.random.choice(non_family_in)
                    m40[swap_out] = False
                    m40[swap_in] = True
        
        # Assign bands
        assigned = None
        if try_exact:
            assigned = _try_exact_optimize(sf, floors, beds, low_share, high_share, avg_low, avg_high,
                                           require_family_at_40=need_family, spread_40_max_per_floor=spread_40_max_per_floor,
                                           exempt_top_k_floors=protect_top)
        
        # Fallback to heuristic when exact optimizer is disabled or fails
        if assigned is None:
            assigned = _heuristic_assign(sf, m40, avg_high, floors=floors)
        
        # Enforce maximum band count and maintain compliance within target ranges
        assigned = enforce_max_bands(assigned, sf, max_bands=3, targets={
            "pct40_low": low_share, "pct40_high": high_share, "avg_low": avg_low, "avg_high": avg_high
        })
        
        # If the weighted average exceeds the upper bound, step down high bands until it fits
        assigned = _adjust_wavg_down(assigned, sf, avg_high)

        # Calculate metrics
        total_sf = float(sf.sum())
        sf40 = float(sf[np.isclose(assigned, 0.40)].sum())
        pct40 = (sf40/total_sf*100.0) if total_sf else 0.0
        wavg = _sf_avg(assigned, sf)*100.0
        vscore = vertical_score(floors, assigned)
        bands_count = len(_unique_bands(assigned))
        
        return {
            "label": label, 
            "mask40": m40, 
            "assigned": assigned,
            "metrics": {"aff_sf_total": total_sf, "sf_at_40": sf40, "pct40": pct40, "wavg": wavg},
            "vscore": vscore, 
            "bands_count": bands_count
        }

    # Generate different scenario variants
    scenarios = []
    
    # Base scenarios
    try:
        scenarios.append(build_scenario("S1", False, 0))
    except Exception:
        pass
    
    try:
        scenarios.append(build_scenario("S2", require_family_at_40, exempt_top_k_floors))
    except Exception:
        pass
    
    try:
        scenarios.append(build_scenario("S3", False, exempt_top_k_floors))
    except Exception:
        pass

    # Deduplicate and prune
    scenarios = dedupe_scenarios(scenarios)
    scenarios = prune_to_top_k(scenarios, return_top_k)
    
    return {s["label"]: s for s in scenarios}

# =========================
# Main allocation function
# =========================
def allocate_with_scenarios(d: pd.DataFrame, require_family_at_40=False, spread_40_max_per_floor=None,
                           exempt_top_k_floors=0, return_top_k=3, use_milp=True):
    """
    Main function to allocate AMI bands with multiple scenarios
    """
    # Normalize headers
    d = _normalize_headers(d)
    
    # Get selected affordable units
    sel = _selected_for_ami_mask(d)
    aff = d.loc[sel].copy()
    
    if len(aff) == 0:
        raise ValueError("No units selected for AMI allocation")
    
    # Generate scenarios with error handling
    try:
        scen = generate_scenarios(
            aff, 
            require_family_at_40=require_family_at_40,
            spread_40_max_per_floor=spread_40_max_per_floor,
            exempt_top_k_floors=exempt_top_k_floors,
            try_exact=use_milp,
            return_top_k=return_top_k
        )
    except Exception:
        # Fallback to heuristic if MILP fails
        if use_milp:
            scen = generate_scenarios(
                aff,
                require_family_at_40=require_family_at_40,
                spread_40_max_per_floor=spread_40_max_per_floor,
                exempt_top_k_floors=exempt_top_k_floors,
                try_exact=False,
                return_top_k=return_top_k
            )
        else:
            raise
    
    if not scen:
        raise ValueError("No valid scenarios could be generated")
    
    # Build output
    labels = list(scen.keys())
    full = d.copy()
    for label in labels:
        full[f"Assigned_AMI_{label}"] = np.nan
        full.loc[sel, f"Assigned_AMI_{label}"] = scen[label]["assigned"]

    base_cols = [c for c in ["FLOOR","APT","BED","NET SF","AMI_RAW"] if c in d.columns]
    aff_br = d.loc[sel, base_cols].copy()
    for label in labels:
        aff_br[f"Assigned_AMI_{label}"] = scen[label]["assigned"]

    def score_ext(m, vs, bc):
        # Higher weighted average AMI is better; lower vscore (lower floors for 40%) is better.
        return (m["wavg"]*20.0) - vs - max(0, bc-3)*2.0
    
    metrics = {k: dict(scen[k]["metrics"]) for k in labels}
    best_label = max(labels, key=lambda k: score_ext(metrics[k], scen[k]["vscore"], scen[k]["bands_count"]))
    best_assigned = scen[best_label]["assigned"]

    # Create mirror table with best scenario
    mirror = d.copy()
    target_col = "AMI" if "AMI" in mirror.columns else "AMI"
    mirror[target_col] = mirror.get(target_col, np.nan)
    mirror.loc[sel, target_col] = best_assigned

    # Add summary footer
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

