# FIX: AMI column shows % consistently after Manual Calculate

**Branch:** `fix/ami-percent-on-manual-calc`
**Cut from:** `feature/excel-agent-foundation` @ `c3056bb`
**Date:** 2026-06-17
**Risk:** Low. One added block in `ManualCalculateScenario`; writes only the AMI cells of units just read, only on the Manual Calculate button, never while typing.

## Symptom (client)

After the Ctrl+Z fix, a number typed into a **non-preformatted** AMI cell displayed as `50` instead of `50%`. Cells already formatted as percent (the highlighted/yellow ones) showed `50%`; blank/General cells showed `50`. Cosmetic only — every consumer already reads a raw `50` as `50%`.

## Why

The Ctrl+Z fix made AMI edits passive (no write while typing), which removed the old on-edit reformat (`50` -> `0.5` + `"0%"`). That reformat was the undo killer, so it had to go. Result: a typed value keeps whatever format the cell already had.

## Change ([excel-addin/src/AMI_Optix_ResultsWriter.bas](excel-addin/src/AMI_Optix_ResultsWriter.bas), `ManualCalculateScenario`)

At the Manual Calculate button path (`preserveAppliedScenario = False`), after units are read, normalize + format the AMI cells of those units:

```vba
For each read unit:
    dataSheet.Cells(unit.row, amiCol).Value = unit.client_ami   ' normalized 0.5
    dataSheet.Cells(unit.row, amiCol).NumberFormat = "0%"        ' shows 50%
```

`GetDataSheet()` / `GetAMIColumn()` are set by `ReadUnitData`; each unit carries `row` + normalized `client_ami`. `Application.EnableEvents = False` around the writes. Only valid-AMI rows are touched (blank / market-rate rows aren't in `units`).

Why this is safe for Ctrl+Z: Manual Calculate already writes the scenario block, so Excel's session undo is cleared at this point regardless. We do NOT format while typing — that's what preserves native undo. **Run MIH/UAP** already normalizes + formats the AMI column via `ApplyBestScenario`, so it needed no change.

## Behavior

- Type AMI values (Ctrl+Z works) → click **Manual Calculate** → the AMI column displays consistently as `%`.
- Optional complementary nicety: **File → Options → Advanced → Enable automatic percent entry** makes typing `50` show as `50%` immediately (Excel-native, no write, undo preserved).

## Verification

- `Sub`/`Function` balance: 68/68.
- Only AMI cells of read units are written; events suppressed; runs only on the button (not on edit).

## Deploy

VBA-only (one `.bas`) → standard `.xlam` refresh (module swap).
