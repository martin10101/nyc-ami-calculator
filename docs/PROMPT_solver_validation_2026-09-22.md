# Session prompt: prove the solver still finds the best legal result (2026-09-22)

Paste everything below this line into a new Claude Code session opened in
`C:\Users\MLFLL\Downloads\nyc-ami-calculator-1` (branch `feature/excel-agent-foundation`).

---

You are validating the AMI Optix optimizer after a week of new features. The owner
needs a hard, numbers-based answer to one question: **does the solver still return
the best legal result (maximum rent, fewest apartments at 40%, closest to the 10%
floor) with every combination of the new features on and off, and can it beat a
human's hand-built option?** Do not change solver behavior without the owner's
explicit approval; adding tests, harnesses, and a report is expected.

## What triggered this

Running MIH Option 1 in Excel on a real building: the recommended option had
**4 units at 40% covering 11.19%** of residential SF. With a "40% Rules" floor
preference set, the run returned **4 units at 40% covering 11.05%** - same count,
closer to the 10% floor. The owner's worry: either the unconstrained search missed
the 11.05% layout (a search gap), or the recommendation no longer prefers "closest
to 10%". Determine which. Note the design: RECOMMENDED = fewest apartments at 40%,
then MAXIMUM rent, tie-break tighter 40% SF (app.py, search
`RECOMMENDED option (client direction 2026-06-12)`). So an 11.05% layout with lower
rent is demoted by design; an 11.05% layout with >= rent that the unconstrained run
never produced is a bug. Ask the owner for that building's file to reproduce; also
reproduce the pattern on synthetic pools.

Second observation: with "Spread Across Floors" on, the 40% units satisfy the
lower/middle/upper test with the minimum (often one unit on the upper floors),
not an even distribution. HPD has no numeric rule (see
`docs/fixes/FIX_floor-spread-rule.md` and memory note on HPD band distribution
research). Measure it, and propose an even-spread tie-break that costs no rent -
propose only; implement only if the owner says yes.

## Read first (in this order)

1. `docs/PLAN_client_feedback_five_fixes_2026-09-02.md` - the client feedback and
   the golden Building D baseline (99 units, 25 affordable on floors 3-19, input
   8@40 / 11@70 / 6@80, WAAMI 0.5997, recommended rent $45,374/mo, 40% window
   10-12.5% of residential SF). Rachel's own hand-built options for Building D:
   8 units at 40%, bands 40/60/80 = $44,933/mo and 40/70/80 = $44,801/mo - the
   program must match or beat these.
2. `docs/fixes/FIX_fewest-40-units-family.md`, `FIX_min-count-40-frontier.md` -
   the fewest-40 family and the RECOMMENDED rule (the owner's core mental model:
   owners want the least apartments at 40%).
3. `docs/fixes/FIX_band-picker-ribbon.md` (Fix 4, payload `allowed_bands`),
   `FIX_floor-spread-rule.md` (Fix 5, payload `floor_spread`, default scope =
   40% band only), `FIX_forty-rules-ribbon.md` (Fix 6, payload `forty_rules`:
   pins / Keep OUT are firm; floors / bedrooms are preferences with a
   restrict -> force -> drop ladder; `max_per_floor` cap).
4. `docs/fixes/FIX_mih-option4-workforce-rules.md`, `FIX_mih-100-band-cap.md`,
   `FIX_mih-40-ami-floor-10pct.md`, `FIX_strict-waami-cap-enforcement.md`.
5. Code: `ami_optix/solver.py` (`_solve_single_scenario` is the single choke
   point every ladder uses; `find_optimal_scenarios`, `find_max_revenue_scenario`,
   `floor_thirds`, `floor_spread_summary`); `app.py` `/api/optimize`
   (`_build_program_config`, `_apply_allowed_bands`, the floor-spread block, the
   40% Rules block with `_fr_apply_level`, `_primary_solve`, the MIH floor-walk,
   the fewest-40 ladder `_f40_solve`, the frontier `_fr_solve`, the RECOMMENDED
   block). Search budgets to be aware of: `max_revenue_combo_checks` (12 in the
   fewest ladder, 36 in the frontier), `max_band_combo_checks`,
   `scenario_time_limit_seconds` (3 s, cut to 1 s for rent-max), `max_unique_scenarios`.

## Rules every output must satisfy (MIH Option 1 unless stated)

- 40% band = 10.0% to 12.5% of residential SF (`mih_residential_sf`); the server
  slides the window up in 0.1% steps only when nothing fits (up to 15-17.5%).
- Weighted average AMI <= 60% (integer method, 1e-9 tolerance); reporting floor 59.1%.
- Bands from {40, 60, 70, 80, 90, 100}; never 50; never above 100 for Option 1;
  at most 3 bands per scenario. Option 4: avg <= 115%, >= 5% at <= 70%, >= 10% at
  <= 90%, bands to 135, at most 4 bands, no 40% requirement. UAP: 20-21% of
  affordable SF at <= 40%.
- Rent per unit = band + bedrooms + utilities from the rent calculator for the
  selected year (floors never change rent). Rent totals must match
  `compute_rents_for_assignments` exactly.
- Fewest apartments at 40% first; RECOMMENDED = fewest, then max rent, then
  tighter 40% SF. Tighter-but-lower-rent layouts appear demoted
  (`tight_40_footprint_N`).
- New features must only NARROW: with `allowed_bands`, `floor_spread`, or
  `forty_rules` set, the recommended rent can never exceed the unconstrained
  run's recommended rent at the same 40% count; and every scenario must still
  satisfy the window, cap, bands, and the feature's own rule (echoed in
  `project_summary.band_rules` / `floor_spread` / `forty_rules`).
- Absent feature fields = byte-identical legacy behavior.

## Test material in the repo

- Synthetic Building D shape: `_pool_units()` in `tests/test_forty_rules.py`
  (25 units, floors 3-19, bedrooms 0/1/1/2/2 cycle, net SF 380-800). Reuse it.
- Real pools as CSV: `tests/test_data/Decatur_for_testing.csv`,
  `tests/test_data/169_Beach_115_Street_for_testing.csv` (golden truths in
  `tests/test_golden.py`), plus `test_data*.csv`, `user_project.csv` at repo root.
- Real building workbooks at repo root: `1004 Wodycrest Avenue (1).xlsx`,
  `1530 Bergen Street (1).xlsx`, `169 Beach 115 Street (1).xlsx`,
  `212 West 231 Street (1).xlsx`, `2675 Decatur Avenue (1).xlsx`, `Test.xlsx`,
  `Unit Schedule - UAP.xlsm`; parser: `ami_optix/parser.py`.
- Add-in golden workbooks: `tools/excel-agent/assets/workbooks/MIH_golden.xlsb`,
  `UAP_golden.xlsm` (read .xlsb with pyxlsb or ask the owner to export .xlsx).
- Rent calculators: `tools/excel-agent/assets/rent-roll-years/{2024,2025,2026}/`
  and `rent_calculators/`; the server auto-seeds them. Send `rent_roll_year`
  in payloads (2026 is what the client uses now).
- The real Building D `.xlsb` and the 2026-09-02 "244 rents matched" harness were
  NOT committed - ask the owner for the file (`Documents\3320 Atlantic - MIH
  v10.xlsb` on Rachel's PC) if you want the true benchmark.
- Existing suites: `tests/test_band_picker.py`, `test_floor_spread.py`,
  `test_forty_rules.py`, `test_low40_options.py`, `test_mih_option4.py`,
  `test_project_summary_compliance.py`, `test_original_snapshot.py`. Run all with
  `python -m pytest tests/ -q -p no:cacheprovider` (about 8 minutes, 145 tests).

## How to call the solver

```python
from app import app
client = app.test_client()
resp = client.post('/api/optimize', json={
    'program': 'MIH', 'mih_option': 'Option 1', 'rent_roll_year': 2026,
    'mih_residential_sf': pool_sf / 0.254,
    'utilities': {'electricity': 'na', 'cooking': 'na', 'heat': 'na', 'hot_water': 'na'},
    'units': units,                       # unit_id, bedrooms, net_sf, floor, client_ami
    # optional: 'allowed_bands': [40, 70, 80, 90, 100],
    #           'floor_spread': True | {'scope': 'all'},
    #           'forty_rules': {'pin_units': [...], 'exclude_units': [...],
    #                           'bedrooms_allowed': [2, 3], 'floors_allowed': [9, 10],
    #                           'max_per_floor': 1},
})
data = resp.get_json()   # data['scenarios'][key]['assignments' | 'rent_totals' | 'metrics'], data['recommended_key'], data['notes'], data['project_summary']
```

Set `AMI_OPTIX_SOLVER_WORKERS=1` for determinism. Each request takes 15-60 s.

## What to build

1. **An independent reference solver** for pools of <= 14 units: enumerate every
   band assignment over the allowed bands, keep the legal ones (window, cap,
   band count), price them with the same rent schedule, and compute the true
   optimum for: max rent; minimum 40% count; at that count, max rent; and the
   tightest 40% share within a rent tolerance. Compare with the API's
   `absolute_best`, `fewest_40_units*`, `recommended_key`, `tight_40_footprint_*`.
   Any case where the API's recommended has the same count but less rent than the
   reference optimum, or where a tighter-at-equal-rent layout exists that the API
   never offered, is a finding. Put it in `tests/test_solver_reference.py`.
2. **A feature matrix sweep** in `tools/solver_validation/run_matrix.py`: pools of
   12, 25 (Building D shape), 30, 60, 99 units plus the real CSV/xlsx pools;
   features {none} x {floor_spread off / on} x {allowed_bands all / no-60 /
   [40,70,80] / [40,60,80]} x {forty_rules none / floors subset / bedroom subset /
   2 pins / max_per_floor 1 / floors + pins}. For every run record: recommended
   key, rent, 40% count, 40% share, WAAMI, bands, level echoed, runtime, and
   check every invariant above. Run each configuration twice to confirm
   determinism.
3. **The owner's case**: reproduce "same count, tighter share under a constraint"
   on synthetic pools; explain each instance as design (lower rent) or gap
   (search budget / frontier missed it). If it is a gap, quantify how often and
   propose the smallest budget or generator change.
4. **Floor-spread distribution**: for every floor_spread-on run, count 40% units
   per third; report how often the upper third gets exactly one. Propose an
   even-spread tie-break at equal rent (do not implement without approval) and
   measure its rent cost with a prototype behind a flag if cheap.
5. **Beat the human**: on Building D (shape now, real file if provided), confirm
   the program's 8-unit options match or beat $44,933 (40/60/80) and $44,801
   (40/70/80) at the same count, and that the recommended is the tightest at
   the top rent.

## Deliverables

- `docs/REPORT_solver_validation_2026-09.md`: plain-English verdict first
  (three sentences an owner can read), then the tables, then every finding with
  a reproduction snippet and a proposed fix. No fix applied to solver/app
  behavior without the owner's go; tests and tooling may be committed on a branch
  `test/solver-validation` merged fast-forward into `feature/excel-agent-foundation`.
- Keep the Python suite green. Do not touch `excel-addin/` or the deploy script.
- IMPORTANT: pushing `feature/excel-agent-foundation` auto-deploys the server on
  Render. Tests and tools are safe to push; any `app.py` / `solver.py` change is
  live for the client the moment it is pushed - so do not push behavior changes.
- The owner reads plain English. Lead every message with the answer, keep the
  numbers in tables, and stop when done.
