# FIX: Floor-spread rule + reviewer view (Fix 5 of the 2026-09-02 client feedback)

**Branch:** `fix/floor-spread`
**Date:** 2026-09-22
**Risk:** Moderate. This is the ONLY fix in the set that changes unit
assignments. Server + solver + VBA + ribbon XML. Opt-in payload field:
absent/false = byte-identical legacy output. No EventHooks / manual-block /
Ctrl+Z changes.

## What the client asked for

HPD's reviewer rejected Building D options 1-4: every 40% unit sat on
floors 4-13, none on the upper stories (14-19). Research
(`reference_hpd_band_distribution_research`): no numeric rule exists; HPD
Design Guidelines 2026 s4.1.4-4.1.6 require bands "distributed throughout
the building ... vertically"; ZR 27-16(b) lets HPD disapprove segregating
layouts. We encode the reviewer's own test.

## The rule (server-side, configurable)

- Split the floors spanned by the affordable pool into lower / middle /
  upper thirds by DISTINCT floor number (stacking charts are per story),
  remainder to the upper thirds: pool floors 3-19 -> 3-7 / 8-13 / 14-19,
  exactly the reviewer's split.
- Every band with >= 3 units (`min_units_per_band`, default 3) must place at
  least one unit in each third. 1-2 unit bands are exempt ("to the maximum
  extent feasible"). Units without a floor value are never constrained.
- Rent-neutral: regulated rent is band + bedrooms only. Cost can only come
  through the SF-share quotas, and the results show the rent either way.
- Tie-break change while the rule is ON: among equal-rent optima the 40%
  units are no longer sunk to the lowest floors; their average floor is
  kept as close as possible to the pool's average floor.
- Honest fallback: if NO band mix can satisfy the rule, the request reruns
  once without it, `project_summary.floor_spread.applied = false` with a
  reason, a note says so, and every scenario still carries its true
  reviewer view (`scenario.floor_spread.satisfied = false`).

## What shipped

**Solver** (`ami_optix/solver.py`): `floor_thirds`, `floor_spread_rule_from`,
`floor_spread_summary`; the constraint inside `_solve_single_scenario`
(the one choke point every ladder uses) driven by
`optimization_rules['floor_spread']`; avg-floor tie-break when ON.

**Server** (`app.py`): parses `floor_spread` (true / {min_units_per_band});
applies the rule only when the pool has >= 3 distinct floors; primary solve
wrapped in `_primary_solve()` for the fallback; per-scenario reviewer view
(Original Scenario included) added only when the rule was requested;
`project_summary.floor_spread = {requested, applied, reason,
min_units_per_band, thirds[]}`.

**Ribbon** (`customUI14.xml`): `toggleButton` "Spread Across Floors" in the
"Bands & Floors" group.

**VBA**: `AMI_Optix_Bands.GetFloorSpreadEnabled / SetFloorSpreadEnabled`
(hidden name `AMI_Optix_FloorSpread`; unset = ON for MIH, OFF for UAP);
ribbon callbacks; payload always sends `floor_spread: true|false`;
ResultsWriter header line "Floor Spread: ON - every band with 3+ apartments
has one on the lower 3-7, middle 8-13, upper 14-19 floors" (or the skip /
fallback reason), and a per-scenario "Floor Spread (HPD reviewer view)"
table: rows upper/middle/lower with floor ranges, one column per band,
"Rule satisfied: Yes/NO" plus the exact missing placements in red.

## Default ON for MIH (owner-facing decision, easy to flip)

The complaint came from an MIH reviewer and the owner asked for the rule
to "respect the floors"; shipping it OFF would leave Rachel's next MIH run
identical to the rejected one. UAP stays OFF until UAP floor expectations
are confirmed. One click on the ribbon toggle turns it off per workbook.
This means the first MIH run after the update WILL differ from today's
output in unit placement (rent within the shown delta); the sandbox QA
compares both.

## Verification

- `tests/test_floor_spread.py` (14 tests): thirds split (Building D ->
  3-7/8-13/14-19; even split; < 3 floors -> None; units without floors
  ignored); rule normalization; reviewer-view summary reproduces the
  rejection of options 1-4 and exempts small bands; solver: rule OFF sinks
  40% to floors 1-2 and breaks the spread, rule ON spreads at the SAME
  rent; infeasible combo -> no scenario; API: absent field -> no metadata
  (identity), rule ON -> every scenario satisfied and 40% reaches the upper
  third, impossible rule -> honest fallback with note, no floor data ->
  skipped with reason.
- Full suite green (see commit message for the count).
- Sandbox Excel QA: MIH file, run with toggle ON (default): header line
  present, each scenario shows the reviewer table with "Yes", 40% appears
  on upper floors, rent close to the previous run (delta visible);
  toggle OFF -> run -> header says nothing about floors, tables gone,
  output matches the pre-update build; Apply, Ctrl+Z, Manual Calculate,
  year change unaffected.

## Deploy

Same one-liner as Fix 4 (module swap + ribbon XML replace), markers
updated to this version. Rollback = re-pin the previous commit.
