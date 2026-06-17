# DIAG: edit-path logging to pinpoint the Ctrl+Z killer (temporary)

**Branch:** `diag/edit-path-logging`
**Cut from:** `feature/excel-agent-foundation` @ `24cf515`
**Date:** 2026-06-17
**Risk:** None to behavior. Logging only, **external file** (`%TEMP%\AMI_Optix_Debug.log`) via `DebugLog` — no workbook writes (no cells / sheets / names / comments). Confirmed against [AMI_Optix_Main.bas:71-91](../../excel-addin/src/AMI_Optix_Main.bas#L71-L91).

## Purpose

Confirm, with data, which workbook write clears Excel's native undo when editing the AMI column — the **immediate normalize write** or the **2-second Manual-Working-Copy refresh** — and why one office PC breaks Ctrl+Z while another (same `.xlam`) does not. No behavior is changed; the existing writes still happen, they are just logged.

## What was instrumented (all lines prefixed `[EDIT]`, all `force:=True`)

- `AMI_Optix_AppEvents.HandleDataSheetAmiChange` — entry (sheet, target, cell count, `Application.EnableEvents`), whether `NormalizeAmiInputCells` is invoked, and when the deferred refresh is armed.
- `AMI_Optix_AppEvents.NormalizeAmiInputCells` — whether it took the **no-write** path (undo preserved) or the **WRITE** path (clears undo).
- `AMI_Optix_EventHooks.AMI_Optix_ScheduleDeferredRefresh` — armed (+2s).
- `AMI_Optix_EventHooks.AMI_Optix_DoDeferredRefresh` — timer fired, each skip reason, and the **write moment** (Manual Working Copy) that clears undo.

## How to read it

Reproduce the break on the bad PC, then open `%TEMP%\AMI_Optix_Debug.log`. The last `[EDIT] ... CLEARS native undo` line **before** the failed Ctrl+Z is the culprit:
- `NormalizeAmiInputCells: raw>2 ... WRITING` → the immediate normalize write.
- `DoDeferredRefresh: writing Manual Working Copy NOW` → the delayed refresh.

PC difference: if editing the AMI column on the "good" PC produces **no `[EDIT]` lines at all**, its live-sync events aren't firing (e.g. `EnableEvents=False`), which is why its undo survives.

## Next step

Superseded by the Option-1 fix (no AMI-edit writes; read-time normalization; refresh only on Manual Calculate / optimizer). These `[EDIT]` logs are removed/gated there.
