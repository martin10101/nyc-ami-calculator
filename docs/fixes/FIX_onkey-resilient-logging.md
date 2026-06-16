# FIX: OnKey reset is resilient + lifecycle logging (run-time error 1004)

**Branch:** `fix/onkey-resilient-logging`
**Cut from:** `feature/excel-agent-foundation` @ `45dcfae`
**Date:** 2026-06-16
**Risk:** Low — wraps two defensive `OnKey` resets in error handling and adds forced logging. No solver/API/manual-sync changes. `.bas`-only (`AMI_Optix_EventHooks`).

## Symptom (client, screenshot)

`Run-time error '1004': Method 'OnKey' of object '_Application' failed`, debugger stopped on `Application.OnKey "^z"` in `StopAMIOptixEventHooks`. A workbook was open in **Protected View** at the time.

## Root cause ([excel-addin/src/AMI_Optix_EventHooks.bas](excel-addin/src/AMI_Optix_EventHooks.bas))

`Auto_Close` → `StopAMIOptixEventHooks` ran `Application.OnKey "^z"` / `"^+z"` **without** error handling (a deliberate choice: "if the reset fails, Ctrl+Z stays hijacked... we need to know"). But `Application.OnKey` raises 1004 in some Excel states (Protected View active, application not `Ready`, shutdown / no usable window). When that happened the user got an unhandled crash.

Important: **this build never hijacks Ctrl+Z** — nothing assigns `OnKey` to a macro anywhere; the resets in `Start`/`Stop` are only a self-heal for older buggy builds. So a *failed* reset is harmless, and the hard crash was unjustified. `Start` already wrapped its resets in `On Error Resume Next`; `Stop` did not — that asymmetry is the bug.

## Change

- New `SafeResetCtrlZ(caller)` helper: attempts both `OnKey` resets, and on failure logs `caller`, the error number/description, and an Excel-state snapshot — never throws. Used by both `Start` and `Stop`.
- New `EventHookContext()` helper: `wbCount`, `activeWb`, `appReady` snapshot for the log.
- Forced logging (`force:=True`, so it records even with debug logging off) at `Auto_Open` / `Auto_Close` / `StartAMIOptixEventHooks` / `StopAMIOptixEventHooks` enter+exit, and on AppEvents wireup failure.

The original intent ("know if the reset ever fails") is now satisfied by a persistent log line rather than an interactive crash.

## Where the log goes

`%TEMP%\AMI_Optix_Debug.log` (see `GetDebugLogPath`). These lifecycle lines are forced, so no setup is needed; send this file after a repro.

## Why this is safe

- No `OnKey` assignment exists in the add-in, so resilient resets cannot leave Ctrl+Z hijacked.
- Only the event-hook lifecycle is touched; the deferred-refresh / manual-sync logic is unchanged.

## Verification

- `Sub`/`Function` vs `End` balance: 17/17.
- After `.xlam` refresh: open a Protected-View workbook and close Excel — no 1004 dialog; `%TEMP%\AMI_Optix_Debug.log` shows the `Auto_Close`/`Stop` sequence and, if `OnKey` still rejects, a `... OnKey ^z reset FAILED [...] - 1004: ...` line pinpointing the state.

## Deploy

VBA-only (one `.bas`) → standard `.xlam` refresh (module swap).
