# FIX: AMI normalize-on-edit must touch only the AMI column (paste corruption)

**Branch:** `fix/paste-normalize-ami-column-only`
**Cut from:** `feature/excel-agent-foundation` @ `966059e`
**Date:** 2026-06-16
**Risk:** Low — narrows one existing call to the AMI column. No change to the Ctrl+Z deferral logic, the manual-block sync, or the solver. `.cls`-only.

## Symptom (client report, screenshot)

After installing the add-in, pasting a unit schedule (or values from another spreadsheet) into an MIH/UAP page corrupted non-AMI columns: numbers showed as percentages. FLOOR `1,2,3,4,5` → `100%, 200%, 3%, 4%, 5%`; BED `1/0/2` → `100%/0%/200%`; NET SF `505,555,432…` → `505%,555%,432%…`.

The displayed values reveal the underlying corruption: any value > 2 was divided by 100 and **every** numeric cell was stamped with a `"0%"` format. `505 → 5.05` shown as `505%`; `3 → 0.03` shown as `3%`; `1` (not > 2) stays `1` but shows as `100%`.

## Root cause ([excel-addin/src/AMI_Optix_AppEvents.cls](excel-addin/src/AMI_Optix_AppEvents.cls), `HandleDataSheetAmiChange`)

Live Sync's `App_SheetChange` calls `HandleDataSheetAmiChange` whenever a change **intersects** the AMI column (`ChangeIsInAMIColumn`). A paste that spans the whole schedule includes the AMI column, so it qualifies. `HandleDataSheetAmiChange` then passed the **entire** changed range to `NormalizeAmiInputCells(target)`:

```vba
Call NormalizeAmiInputCells(target)   ' target = ALL pasted columns
```

`NormalizeAmiInputCells` divides any numeric value > 2 by 100 and sets `NumberFormat = "0%"` on every numeric cell it is given. Handed a multi-column paste, it applied AMI treatment to FLOOR / BED / NET SF and any other numeric column.

Note: Live Sync cannot currently be toggled off (`GetLiveSyncEnabled` is hardwired to `True`), so there was no user-facing way to avoid this short of disabling the add-in.

## Change

Restrict normalization to the AMI column before calling the normalizer:

```vba
Dim headerRow As Long
headerRow = FindHeaderRow(ws)
If headerRow > 0 Then
    Dim amiCol As Long
    amiCol = FindAMIColumn(ws, headerRow)
    If amiCol > 0 Then
        Dim amiData As Range
        Set amiData = ws.Range(ws.Cells(headerRow + 1, amiCol), ws.Cells(ws.Rows.Count, amiCol))
        Dim amiTarget As Range
        Set amiTarget = Intersect(target, amiData)
        If Not amiTarget Is Nothing Then Call NormalizeAmiInputCells(amiTarget)
    End If
End If
```

Uses the same `FindHeaderRow` / `FindAMIColumn` helpers as `ChangeIsInAMIColumn`, so the gating and the normalization now agree on what "the AMI column" is.

## Effect

- Single-cell AMI edit and AMI-only paste/fill: unchanged (still normalized).
- Multi-column paste that includes the AMI column: only the AMI cells are normalized; FLOOR / BED / NET SF and every other column are left exactly as pasted.
- Ctrl+Z behavior is unchanged — the deferred-refresh logic below this block is untouched, and non-AMI cells are no longer written at all (strictly fewer writes).

## Verification

- Logic trace against the reported screenshot: with the guard, only AMI-column cells reach `NormalizeAmiInputCells`; the other columns are never divided or reformatted.
- Manual after `.xlam` refresh: paste a full unit schedule into an MIH and a UAP page — FLOOR/BED/NET SF paste verbatim; typing `60` into an AMI cell still becomes `60%`.

## Deploy

VBA-only (one `.cls`) → standard `.xlam` refresh (module swap).
