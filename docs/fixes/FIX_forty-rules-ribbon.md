# FIX: 40% Rules - the owner decides which apartments are 40% (Fix 6)

**Branch:** `feat/forty-rules`
**Date:** 2026-09-22
**Risk:** Moderate, additive. Server + solver (one new constraint) + VBA +
ribbon XML. Absent payload field = byte-identical legacy behavior. No
EventHooks / manual-block / Ctrl+Z changes.

## What the owner asked for

"An option where the client can decide how many 40% units go on each
floor, or which units (two bedroom, whatever the case is), and it still
looks for the maximum rent. Optional, with a clear option to go back to
letting the program decide."

## What shipped

**Ribbon** (`customUI14.xml`): `dynamicMenu` "40% Rules" in the Bands &
Floors group, rebuilt on every drop:

- Pin selected units at 40% / Keep selected units OUT of 40% / Remove
  rule from selected units - the user selects unit rows on the MIH / UAP
  sheet, then clicks. Rows are mapped to unit ids through the same
  DataReader the run uses (`unit("row")`); a selection on another sheet or
  with no affordable units is refused with a message.
- Bedroom types allowed at 40% (Studio / 1 / 2 / 3 / 4+ BR; default all;
  at least one must stay checked).
- Floors allowed for 40% (one checkbox per floor of the affordable pool,
  default all; at least one must stay checked; "Allow all floors" resets).
  Added 2026-09-22 after the owner's first test: "which floors" was the
  original ask. If the floor rule keeps 40% off a whole floor third, the
  floor-spread test is skipped with a note instead of failing.
- Max 40% units per floor (No limit / 1 / 2 / 3 / 4), radio-style.
- Show current 40% rules; Clear all 40% rules (program decides).

**VBA** (new `AMI_Optix_FortyRules.bas`): per-workbook storage in four
hidden defined names (`AMI_Optix_Forty_Pins/Excludes/Bedrooms/Floors/PerFloor`),
selection mapping, rule edits, payload JSON, run preflight (pinned /
excluded ids must exist in this run; no unit both pinned and excluded -
stops the run with a clear message). Callbacks in `AMI_Optix_Ribbon`;
payload field `forty_rules` in `BuildAPIPayloadV2` only when a rule is
set; results header "40% Rules: ..." (`WriteFortyRulesLine`, survives
Manual Calculate / year change).

**Server** (`app.py`): `_normalize_forty_rules`; pins -> `fixedUnits`
[40]; exclusions and units outside the allowed bedroom types ->
`fixedUnits` [every allowed band except 40]; both through the existing
`project_overrides` path every solver call already receives. Per-floor
cap -> `optimization_rules['forty_max_per_floor']`, enforced in
`_solve_single_scenario` (new constraint: sum of <=40% units on a floor
<= cap). Checked BEFORE solving against the 40% share window: pinned SF
above the ceiling (12.5% + the 5-point slide for MIH; UAP widen cap) or
eligible SF below the floor return a clean `success:false` error. Echo in
`project_summary.forty_rules` (+ `summary` string) and a note "Built with
your 40% rules: ...". If the run finds nothing, a note says the rules may
be too tight.

**Objective unchanged:** best rent within the rules; ranking, fewest-40
logic, floor spread and band picker all still apply (stacking tested).

## Orders vs preferences (owner clarification 2026-09-22)

On a fresh client file the AMI column is a placeholder (e.g. 60% on every
affordable row) that only marks the pool; the program decides the real
bands. The owner wants to say, BEFORE running, "put the 40% units on
floor 9, 2 BR / 3 BR" as a wish that never breaks the run:

- **Orders (firm):** Pin selected units at 40%; Keep selected units OUT.
  Impossible orders are refused with a plain error before solving.
- **Preferences (wish):** Floors for 40%, Bedroom types for 40%. Server
  ladder, tried in order until the solver finds scenarios:
  1. `restrict` - 40% only from the preferred apartments (when their SF can
     reach the 40% floor);
  2. `force` - every preferred apartment IS 40%, the program fills the rest
     (when their SF is under the ceiling);
  3. `drop` - preference and per-floor cap set aside, pins / Keep OUT kept,
     with a note "could not be honored".
  `project_summary.forty_rules.level` + `summary` say which level ran; the
  results header prints it. Floor spread is skipped only at `restrict` when
  the chosen floors miss a whole third.

## Pool vs rules (the owner's first question)

Which apartments are AFFORDABLE is Rachel's input: any row with a value in
the AMI column is in the pool; a blank AMI row is market rate and the
program never touches it. Which BAND each pooled apartment gets is the
program's decision on every run. 40% Rules are inputs set before Run that
override that decision for the pooled apartments they name. Selecting a
blank-AMI row and clicking Pin is refused with a message naming the rows
and explaining why (a market-rate unit cannot become 40%).

## Verification

- `tests/test_forty_rules.py` (19 tests): normalization; per-floor cap
  binds and is rent-neutral, too-tight cap -> no scenario; API: absent
  field -> no metadata (identity); pins + exclusions obeyed in every
  scenario; 2 BR-only filter obeyed; max 1 per floor obeyed; rules stack
  with floor spread + band picker; clean errors for unknown unit, unit
  both pinned and excluded, too many pins, too little eligible SF.
- Full suite green (see commit); VBA compiles (Test-VbaCompile.ps1);
  ribbon XML parses.
- Sandbox Excel QA: select two 2 BR rows on the MIH sheet -> 40% Rules >
  Pin -> run -> header shows "40% Rules: pinned at 40%: ..." and those
  units are 40% in every option; Keep OUT on another row -> never 40%;
  Bedroom types -> untick Studio and 1 BR -> run -> only 2 BR+ at 40%;
  Max per floor 1 -> no floor has two 40% units; Clear all -> header line
  gone, output as before; selecting rows on the AMI Scenarios sheet and
  clicking Pin is refused with a message.

## Deploy

Same one-liner (module swap + ribbon XML replace); adds
`AMI_Optix_FortyRules`; markers updated. Rollback = re-pin.
