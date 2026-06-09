# FIX: Restore native Ctrl+Z everywhere in Excel

**Branch:** `fix/restore-native-ctrl-z`
**Cut from:** `feature/excel-agent-foundation` @ `9ba38f4`
**Date:** 2026-06-09
**Risk:** Medium — touches the global Ctrl+Z key binding and the deferred-refresh timer. Self-heals broken installs on add-in load.

## Symptom

Reported by user on both his dev PC and the client's PC after the previous Ctrl+Z fix shipped:

- **Dev PC:** Only ONE Ctrl+Z works per edit. Excel acts like the undo stack has depth 1. The user can't undo multiple prior actions like they could before the add-in was installed.
- **Client PC:** Bugs vary by cell — sometimes Ctrl+Z works, sometimes it doesn't, and the pattern correlates with which row/column the user is in.
- **Affects every workbook the user opens**, not just AMI calc files. A pristine invoice spreadsheet that has nothing to do with this add-in has broken Ctrl+Z too.

## Root cause — three pieces of global poison

The previous "Ctrl+Z fix" added three mechanisms that all leaked outside the AMI workbook.

### 1. Global Ctrl+Z hijack

`AMI_Optix_ArmCtrlZIntercept` registered `Application.OnKey "^z"` (and `"^+z"`) at add-in load, routing **every Ctrl+Z press in every workbook in the Excel session** through `AMI_Optix_HandleCtrlZ`. The handler did its own custom undo if it had armed state, otherwise called `Application.Undo` as a passthrough. Two problems:

- `Application.Undo` only undoes the most recent action — it doesn't pop the stack. So repeated Ctrl+Z always undid the same single most-recent thing, then nothing. That's the "only one undo works" symptom.
- For workbooks that never had AMI-armed state, the OnKey route still ran (even if it just called `Application.Undo`). Excel's native Ctrl+Z handling is richer than `Application.Undo` — it manages the multi-step stack, redo, etc. Hijacking and re-emitting via `Application.Undo` is a regression even when the handler thinks it's "doing nothing."

`AMI_Optix_ArmUndoForAmiEdit` re-armed the OnKey hijack on every AMI edit (line 176), so even if the add-in unloaded cleanly, any subsequent edit put the poison right back.

### 2. Deferred refresh fired regardless of active workbook

`AMI_Optix_DoDeferredRefresh` was scheduled via `Application.OnTime` for `Now + 2 seconds`. When it fired, it called `RefreshManualWorkingCopyLocalRents` which writes 100+ cells. There was no check that the AMI workbook was still active.

So this sequence broke Ctrl+Z globally:

1. User edits an AMI cell → 2-second timer armed.
2. User switches to a different workbook within those 2 seconds.
3. User does work there, presses Ctrl+Z to undo a typo.
4. Our timer fires and writes 100+ cells into the AMI workbook (or whichever happens to be active), clearing Excel's session-wide undo stack.
5. Their unrelated Ctrl+Z stack is gone.

That matches the "varies by cell location" client-PC symptom — it depends on where the user happened to be standing when the timer fired.

### 3. Silent error swallow at unload

`StopAMIOptixEventHooks` was wrapped in `On Error Resume Next`. If the OnKey release ever threw (object already nothing, COM error during shutdown), the hijack survived the apparent unload. The proc reference now pointed at dead code and Ctrl+Z was permanently broken for the rest of the Excel session. This explained why the breakage outlasted closing and reopening files.

## Fix

All in [excel-addin/src/AMI_Optix_EventHooks.bas](excel-addin/src/AMI_Optix_EventHooks.bas) and [excel-addin/src/AMI_Optix_AppEvents.cls](excel-addin/src/AMI_Optix_AppEvents.cls). No new modules, no `.frm`/`.frx`, no ribbon XML.

### 1. Delete the global Ctrl+Z hijack entirely

Removed:

- `AMI_Optix_ArmCtrlZIntercept`
- `AMI_Optix_HandleCtrlZ`
- `AMI_Optix_UndoLastAmiEdit`
- `AMI_Optix_ArmUndoForAmiEdit` (was only used to feed the now-dead handlers)
- The `g_AMIOptixUndo*` module-level state vars
- Both callers of `ArmUndoForAmiEdit` in `AMI_Optix_AppEvents.cls` (`HandleDataSheetAmiChange` and `HandleManualScenarioEdit`)

Excel's native Ctrl+Z now handles undo for every workbook, including ours. No middleman.

### 2. Self-heal on add-in load

Existing installs already have the broken OnKey registered, so a code-only fix isn't enough on its own — those bindings need to be cleared from the running Excel session. New first block of `StartAMIOptixEventHooks`:

```vba
On Error Resume Next
Application.OnKey "^z"
Application.OnKey "^+z"
On Error GoTo 0
```

Calling `OnKey` with no second argument resets the key to default. Idempotent and safe on fresh sessions. Runs before anything else, so as soon as the fixed `.xlam` loads, the user's global Ctrl+Z is restored.

### 3. Gate the deferred refresh on active workbook

`AMI_Optix_ScheduleDeferredRefresh` now captures `ActiveWorkbook.Name` into a new module-level var `g_AMIOptixDeferredRefreshWorkbook`. `AMI_Optix_DoDeferredRefresh` checks two things before writing:

- Active workbook name must match the one that scheduled the timer.
- That workbook must have an `"AMI Scenarios"` sheet (belt-and-suspenders, since workbook names can collide).

If either check fails, the timer fires harmlessly and exits without touching any cell. The user's undo stack in their other workbook stays intact.

### 4. Cancel deferred refresh on workbook switch

New `App_WorkbookDeactivate` handler in `AMI_Optix_AppEvents.cls` calls `AMI_Optix_CancelDeferredRefresh` the moment the user leaves the AMI workbook. Belt-and-suspenders with the gate above — the gate would catch it, but canceling outright is cleaner.

New `App_WorkbookActivate` handler reschedules the refresh when the user returns to the AMI workbook (only if a refresh was actually pending for *that* workbook).

### 5. Remove `On Error Resume Next` from the OnKey release path in `StopAMIOptixEventHooks`

The release calls now run unconditionally — no silent swallow. Subsequent state cleanup (event object teardown) still uses `On Error Resume Next` because that part is genuinely best-effort, but the critical OnKey reset is no longer suppressible.

## Trade-off

The custom Ctrl+Z fallback that restored a manually-typed-raw AMI value (user types `60` instead of `60%`) is gone. If the user makes that typo and immediately presses Ctrl+Z, native undo will restore `60` but the cell's NumberFormat is still `0%`, so it'll display as `6000%`. They need to retype.

This is rare and acceptable. The alternative — keeping the OnKey hijack — destroyed global Ctrl+Z for every unrelated workbook on the user's machine. Not worth it.

## Verification

Both PCs (dev + client):

1. **Cold-start global Ctrl+Z test.** Close all of Excel. Open a non-AMI workbook. Make 5 edits. Press Ctrl+Z five times. All five must undo.
2. **Cross-workbook test.** Open the AMI workbook. Edit an AMI cell. Within 2 seconds switch to a different workbook and press Ctrl+Z there. Native undo must work; no refresh-write must land in the other workbook.
3. **AMI workbook normal-flow test.** Edit an AMI cell with a clean value (`60%`). Wait 3 seconds. Manual Working Copy must refresh — confirms the deferred refresh still fires when it should.
4. **AMI workbook return test.** Edit an AMI cell. Switch workbooks within 1 second. Come back within 5 seconds. Refresh fires after return.
5. **Self-heal test.** On the client PC where Ctrl+Z is currently broken, install the new `.xlam` via the standard PowerShell refresh one-liner. Without restarting Excel, open any random workbook. Ctrl+Z must work immediately, before any AMI cell is touched.

## Files modified

- [excel-addin/src/AMI_Optix_EventHooks.bas](excel-addin/src/AMI_Optix_EventHooks.bas) — deleted hijack functions, gated deferred refresh on active workbook, added cancel/reschedule helpers, self-heal on load
- [excel-addin/src/AMI_Optix_AppEvents.cls](excel-addin/src/AMI_Optix_AppEvents.cls) — added `App_WorkbookActivate`/`App_WorkbookDeactivate`, removed dead `ArmUndoForAmiEdit` call sites
