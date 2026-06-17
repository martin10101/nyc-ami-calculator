# FIX: AMI-column edits never write to the workbook (restore native Ctrl+Z)

**Branch:** `fix/ami-edit-no-write`
**Cut from:** `feature/excel-agent-foundation` @ `de87e15`
**Date:** 2026-06-17
**Risk:** Low-moderate. Removes two writes from the AMI-edit event path; relies on the read-time normalization that already exists everywhere. No solver/API changes, no `Application.OnUndo`, no custom undo. `.cls` + `.bas` only.

## Symptom (client, both confirmed by isolation test)

Editing the **AMI column** on an MIH/UAP page broke Excel's native Ctrl+Z. Two independent causes, each reproduced:
1. **Immediate:** typing a whole number (e.g. `50`) flipped the cell to `50%` and killed undo *instantly*.
2. **Delayed:** ~2-3 seconds after an AMI edit, undo died on its own.

Non-AMI columns and blank workbooks were unaffected. (One of two PCs didn't show it — explained below.)

## Root cause

Any VBA cell write clears Excel's **session-wide** native undo stack. The AMI-edit handler (`HandleDataSheetAmiChange`) did two writes:
1. `NormalizeAmiInputCells` rewrote `50` → `0.5` and stamped `"0%"` (the instant killer).
2. `AMI_Optix_ScheduleDeferredRefresh` wrote the Manual Working Copy ~2s later (the delayed killer).

Confirmed live via temporary external logging ([DIAG_ctrlz-edit-path-logging.md](DIAG_ctrlz-edit-path-logging.md)). PC difference: the "good" PC either had live-sync events off or Excel's *automatic percent entry* on (so `50` became `0.5` natively, value ≤ 2, no normalize write) — so no write, undo survived.

## Change (Option 1 — refresh on demand; client-approved)

`AMI_Optix_AppEvents.HandleDataSheetAmiChange` is now **passive** — on an AMI edit it updates only the in-memory selection cache; it performs **no workbook write**:
- No immediate normalize write. `NormalizeAmiInputCells` is **removed**.
- No deferred refresh scheduling.

Raw entries are normalized at **read time**, which already happens everywhere they are consumed (`> 2 → /100`): `AMI_Optix_DataReader` (optimizer), `AMI_Optix_RentCalcTables`, `AMI_Optix_ResultsWriter` (Manual block / scenarios), `ParseAmiValue`, `AMI_Optix_VerifyManualRents`. So a cell holding `50` is interpreted as `50%` by the optimizer, Manual Calculate, and rent calcs without any rewrite.

The Manual Working Copy refreshes **on demand only**: the **Manual Calculate** button (`Ribbon_ManualCalculate`, which already cancels any pending refresh) or a solver run (**Run MIH / Run UAP**) — points where a write (and undo reset) is expected because the user initiated it.

Explicitly **not** done: no `Application.OnUndo`, no custom undo system, no optimizer logic change, no form/ribbon change.

## Behavior change to expect

- Editing AMI cells no longer auto-updates the Manual block; click **Manual Calculate** (or run the solver) to refresh it.
- A whole number typed into a non-percent AMI cell stays as typed (e.g. `50`) and is read as `50%`. To make typing `50` display as `50%` natively (no write, undo preserved), enable Excel's **File → Options → Advanced → Enable automatic percent entry** — optional, per PC.

## Verification

- `Sub`/`Function` balance: AppEvents 24/24, EventHooks 17/17.
- No remaining references to `NormalizeAmiInputCells`.
- Read-time normalization confirmed present in all consumer paths.
- Manual after `.xlam` refresh: in an MIH page, change an AMI cell, type values in rows below, press Ctrl+Z repeatedly → every edit undoes. Click Manual Calculate → Manual block matches the AMI column. Run MIH → scenarios compute correctly from the typed AMIs.

## Deploy

VBA-only (`.cls` + `.bas`) → standard `.xlam` refresh (module swap). The temporary `[EDIT]` diagnostic logs are removed; one confirmation breadcrumb remains in `HandleDataSheetAmiChange` (external log only).
