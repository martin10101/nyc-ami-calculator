# FIX: "Rent Roll Year" line survives Manual Calculate and year switches

**Branch:** `fix/rent-year-line-persists`
**Cut from:** `feature/excel-agent-foundation` @ `171132d`
**Date:** 2026-06-10
**Risk:** Very low — VBA-only, display-only; three writers now share one helper.

## Symptom

After the rent-year fix shipped, the "Rent Roll Year: NNNN" line appeared under "AMI OPTIMIZATION RESULTS" the first time (after Find Optimal Scenarios) — then **disappeared** when the user pressed Manual Calculate or changed the year via the ribbon, and never updated to the newly selected year.

## Cause

The results sheet's top block is erased (`ClearManualBlock`) and fully rebuilt by **three different writers**, and only one of them wrote the year line:

1. `WriteManualScenarioBlockFromResult` — Find Optimal Scenarios response → had the line (shipped in `171132d`)
2. `WriteManualScenarioBlockFromEvaluate` — Manual Calculate response → missing
3. `RefreshManualWorkingCopyLocalRents` — local rebuild on year switch / AMI cell edits → missing

Whichever writer ran last decided whether the line existed. Manual Calculate and year switches run writers 2/3, wiping the line.

## Fix ([excel-addin/src/AMI_Optix_ResultsWriter.bas](excel-addin/src/AMI_Optix_ResultsWriter.bas))

Two small shared helpers:

- `WriteRentRollYearLine(ws, row, label)` — writes the bold "Rent Roll Year:" line; empty label writes nothing.
- `ResolveRentYearLabelFromResponse(resp)` — reads `rent_roll_year_used` from a server response (both `/api/evaluate` and `/api/manual_calculate` already return it via `resp.update(rent_meta)`, and `/api/optimize` since `171132d`); falls back to the local dropdown year labeled "(local)".

All three writers now emit the line right under the title:

- Optimize writer → server-reported year (refactored to use the helpers).
- Evaluate/Manual Calculate writer → server-reported year from the evaluate response.
- Local refresh writer → `selectedYear` (the dropdown year it priced with, already in scope) labeled **"(local)"**, since that path computes rents inside Excel.

The "(local)" suffix intentionally distinguishes locally-priced blocks from server-priced ones — any future mismatch is visible instead of silent.

## Verification (on client PC after refresh)

1. Find Optimal Scenarios with dropdown on 2026 → "Rent Roll Year: 2026".
2. Press Manual Calculate → line persists, shows the year Manual Calculate used.
3. Switch ribbon year 2026 → 2025 → block refreshes, line reads "2025 (local)".
4. Edit an AMI cell, wait ~2s for the deferred refresh → line still present.
5. Full pytest suite: 75 passed (no Python changes).

## Deploy

VBA-only → standard `.xlam` PowerShell refresh on both PCs. Render rebuild is a no-op.
