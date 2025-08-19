# ami_core.py — unified, optimized allocator for UAP
# - Bands: {0.40,0.60,0.70,0.80,0.90,1.00}
# - If AFA <10k SF: all 0.60
# - Else: >=20% SF at 0.40, max weighted avg <=0.60 (closest to 0.60)
# - <=3 bands per scenario
# - MILP: lex max avg, then variant objectives
# - Heuristic fallback
# - Distinct variants: low floors 40%, max avg, small units 40%

import os
from typing import Optional, Dict, List, Tuple
import numpy as np
import pandas as pd

DEFAULT_MILP_TIMELIMIT = int(os.getenv("MILP_TIMELIMIT_SEC", "10"))

# ... (keep FLEX_HEADERS, _coerce_percent, _normalize_headers, _selected_for_ami_mask as is)

# Math helpers (keep _sf_avg, _unique_bands)

# 40% selection (MITM for heuristic, but MILP integrates it)
def _choose_40_mitm(sf: np.ndarray, low_share=0.20, high_share=0.201) -> np.ndarray:
    # ... (keep as is, but adjust high_share for tolerance)

# MILP for assignment
def _try_exact_optimize(sf: np.ndarray, floors: np.ndarray, beds: np.ndarray, sfs: np.ndarray, low_share=0.20, high_share=0.201,
                        avg_high=0.60, require_family_at_40=False, spread_40_max_per_floor=None,
                        exempt_top_k_floors=0, objective_mode='max_avg'):
    try:
        import pulp
    except ImportError:
        return None

    n = len(sf)
    bands = [0.40, 0.60, 0.70, 0.80, 0.90, 1.00]
    total_sf = sf.sum()

    def build_lp(maximize_wavg=True, wavg_floor=None):
        sense = pulp.LpMaximize if maximize_wavg else pulp.LpMinimize if 'min' in objective_mode else pulp.LpMaximize
        prob = pulp.LpProblem("AMI_Allocation", sense)
        x = pulp.LpVariable.dicts("x", ((i, j) for i in range(n) for j in range(len(bands))),
                                  cat='Binary')
        
        # Each unit one band
        for i in range(n):
            prob += pulp.lpSum(x[i, j] for j in range(len(bands))) == 1
        
        # >=20% SF at 40%
        sf_40 = pulp.lpSum(sf[i] * x[i, 0] for i in range(n))
        prob += sf_40 >= total_sf * low_share
        prob += sf_40 <= total_sf * high_share
        
        # Weighted avg <=60%
        wavg = pulp.lpSum(bands[j] * sf[i] * x[i, j] for i in range(n) for j in range(len(bands))) / total_sf
        prob += wavg <= avg_high
        
        if maximize_wavg:
            prob += wavg  # max
        else:
            if objective_mode == 'low_floors':
                prob += pulp.lpSum(floors[i] * x[i, 0] for i in range(n))  # min floor sum for 40%
            elif objective_mode == 'small_units':
                prob += pulp.lpSum(sf[i] * x[i, 0] for i in range(n))  # min SF sum for 40%
            elif objective_mode == 'max_avg':  # already in phase 1
                pass
            if wavg_floor is not None:
                prob += wavg >= wavg_floor
        
        # Constraints (keep family, spread, exempt as is)

        return prob, x

    # Phase 1: Max wavg <=0.60
    prob1, x1 = build_lp(maximize_wavg=True)
    prob1.solve(pulp.PULP_CBC_CMD(msg=False, timeLimit=DEFAULT_MILP_TIMELIMIT))
    if pulp.LpStatus[prob1.status] != 'Optimal':
        return None
    wavg_opt = pulp.value(prob1.objective)

    # Phase 2: Variant objective at wavg >= opt
    prob2, x2 = build_lp(maximize_wavg=False, wavg_floor=wavg_opt)
    prob2.solve(pulp.PULP_CBC_CMD(msg=False, timeLimit=DEFAULT_MILP_TIMELIMIT))
    if pulp.LpStatus[prob2.status] != 'Optimal':
        return None

    assigned = np.array([bands[j] for i in range(n) for j in range(len(bands)) if pulp.value(x2[i, j]) == 1])
    return assigned

# Heuristic (updated for variants)
def _heuristic_assign(sf: np.ndarray, mask40: np.ndarray, floors: np.ndarray, beds: np.ndarray, target_avg=0.60, objective_mode='max_avg'):
    assigned = np.full(len(sf), 0.60)
    assigned[mask40] = 0.40
    current_avg = _sf_avg(assigned, sf)
    
    bands_high = [0.70, 0.80, 0.90, 1.00]
    remaining = ~mask40
    rem_idx = np.argsort(floors[remaining]) if 'low_floors' in objective_mode else np.argsort(sf[remaining]) if 'small_units' in objective_mode else np.random.shuffle(np.where(remaining)[0])
    
    while current_avg < target_avg and len(bands_high) > 0:
        for i in rem_idx:
            if current_avg >= target_avg:
                break
            for b in sorted(bands_high, reverse=True):  # prefer high to push avg
                new_avg = (current_avg * sf.sum() - 0.60 * sf[i] + b * sf[i]) / sf.sum()
                if new_avg <= target_avg:
                    assigned[i] = b
                    current_avg = new_avg
                    break
    
    return assigned

# Generate scenarios
def generate_scenarios(aff: pd.DataFrame, low_share=0.20, high_share=0.201, avg_high=0.60,
                       require_family_at_40=False, spread_40_max_per_floor=None, exempt_top_k_floors=0,
                       try_exact=True, return_top_k=3):
    total_aff_sf = aff["NET SF"].sum()
    if total_aff_sf < 10000:
        assigned = np.full(len(aff), 0.60)
        metrics = {"S1": {"aff_sf_total": total_aff_sf, "sf_at_40": 0.0, "pct40": 0.0, "wavg": 60.0}}
        return {"S1": {"assigned": assigned, "metrics": metrics["S1"], "vscore": 0, "bands_count": 1}}

    sf = aff["NET SF"].to_numpy()
    floors = aff.get("FLOOR", np.ones(len(sf))).to_numpy()  # assume 1 low
    beds = aff.get("BED", np.zeros(len(sf))).to_numpy()

    def build(label, obj_mode, need_family):
        m40 = _choose_40_mitm(sf, low_share, high_share)
        if need_family and not np.any((beds >= 2) & m40):
            # Adjust mask for family (as before)
            pass
        
        if try_exact:
            assigned = _try_exact_optimize(sf, floors, beds, sf, low_share, high_share, avg_high,
                                           require_family_at_40=need_family, spread_40_max_per_floor=spread_40_max_per_floor,
                                           exempt_top_k_floors=exempt_top_k_floors, objective_mode=obj_mode)
        if assigned is None:
            assigned = _heuristic_assign(sf, m40, floors, beds, avg_high, objective_mode=obj_mode)
        assigned = enforce_max_bands(assigned, sf, max_bands=3, targets={"pct40_low": low_share, "pct40_high": high_share, "avg_high": avg_high})

        # Metrics and guard (updated for >=20%, <=60%)
        total_sf = float(sf.sum()); sf40 = float(sf[np.isclose(assigned,0.40)].sum())
        pct40 = (sf40/total_sf*100.0) if total_sf else 0.0
        wavg = _sf_avg(assigned, sf)*100.0
        if not (20.0 <= pct40 <= 20.1):
            raise RuntimeError(f"Scenario {label}: 40% share out of band ({pct40:.3f}%).")
        if wavg > 60.0 + 1e-6:
            raise RuntimeError(f"Scenario {label}: Weighted average exceeds 60% ({wavg:.3f}%).")
        vscore = -np.sum(floors[np.isclose(assigned,0.40)]) / np.sum(np.isclose(assigned,0.40)) if np.any(np.isclose(assigned,0.40)) else 0  # negative for low floors better
        bands_count = len(_unique_bands(assigned))
        return {"label": label, "assigned": assigned, "metrics": {"aff_sf_total": total_sf, "sf_at_40": sf40, "pct40": pct40, "wavg": wavg}, "vscore": vscore, "bands_count": bands_count}

    S1 = build("S1", 'low_floors', require_family_at_40)
    S2 = build("S2", 'max_avg', require_family_at_40)
    S3 = build("S3", 'small_units', require_family_at_40)

    out_raw = {"S1": S1, "S2": S2, "S3": S3}
    out_raw = dedupe_scenarios(out_raw)
    out_top = prune_to_top_k(out_raw, top_k=return_top_k)
    return out_top or out_raw

# ... (keep enforce_max_bands, dedupe_scenarios, prune_to_top_k as in previous fix, but update score to favor close to 60%, low vscore, low bands)

# Public entry (keep, but add mirror_out write in app.py)
def allocate_with_scenarios(...):
    # ... (keep core, but integrate new generate_scenarios)
