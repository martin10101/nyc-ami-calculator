# FIX: "YOUR ORIGINAL INPUT" always shows what the user really typed

**Branch:** `fix/original-input-snapshot`
**Date:** 2026-09-02
**Risk:** Moderate scope, additive only. Server + VBA. No ribbon change,
no EventHooks change, no manual-block change.

## Symptom (client, Building D)

"Option 11: That was not my original input. I didn't have any units at 60% AMI."
Verified: her true input was 8@40/11@70/6@80; the sheet showed the previous
run's applied scenario (40/60/90) labeled YOUR ORIGINAL INPUT. Root cause:
the server builds the Original Scenario from the workbook's live AMI column,
which ApplyBestScenario had overwritten.

## Change

**Server (app.py, Original-scenario builder):** prefer `units[].original_ami`
when present; fall back to `client_ami` (legacy add-ins → byte-identical
behavior).

**VBA (new module `AMI_Optix_Baseline.bas`):** very-hidden sheet
`AMI_Optix_Baseline` in the data workbook stores unit_id -> AMI as last typed,
plus a signature of the last PROGRAM write to the AMI column.

- Before each run (hook in Main after ReadUnitData): current column signature
  == last-program-write signature -> user changed nothing -> keep baseline;
  different -> user hand-edited (or first run) -> current column becomes the
  new baseline. Units get `original_ami`; BuildAPIPayloadV2 (API.bas) sends it.
- After each program write (hooks at the end of ApplyBestScenario,
  ApplyCanonicalAssignmentsToDataSheet, ApplyScenarioByKey, and
  ClearProgramAmiColumn): RecordProgramWrite stores the new signature, so a
  year-change clear or an apply never masquerades as user input.
- Every entry point swallows its own errors: baseline failure can never block
  a run or an apply; worst-case failure mode is today's (pre-fix) behavior.

## Rerun/edit semantics (owner-approved)

apply -> rerun: baseline kept (truthful Original). User edits any AMI cell
after an apply: that edited column becomes the new baseline (their new input).
First run on a workbook: current column captured as baseline.

## Verification

- tests/test_original_snapshot.py (new): snapshot honored (Building D shape:
  applied 40/60/90 vs typed 40/70/80 - Original must show typed, no 60% band);
  legacy fallback exact.
- Full suite: 89 passed.
- Sandbox Excel QA: run -> apply -> rerun -> Original equals first input;
  hand-edit one AMI -> rerun -> Original equals edited column; year change ->
  baseline survives; Ctrl+Z / Manual Calculate unaffected.

## Deploy

Module swap now carries 6 modules (adds Baseline, Main, API). Server deploys
via Render on push (inert for old add-ins). Staged: sandbox -> secondary ->
Rachel.
