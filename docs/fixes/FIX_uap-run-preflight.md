# FIX: UAP — Run preflight popup (utilities reminder before running UAP)

**Branch:** `fix/mih-band-ceiling-and-uap-popup` (combined commit with Fix D.2)
**Cut from:** `feature/excel-agent-foundation` @ `6ee8088`
**Date:** 2026-04-29
**Author:** client request via remote session
**Risk:** Very Low (single ribbon callback, UAP only, MIH untouched)

## Status

| Step | When | Result |
|---|---|---|
| Committed locally | 2026-04-29 — combined with Fix D.2 commit | ✅ |
| Pushed `fix/mih-band-ceiling-and-uap-popup` | _pending_ | ⏳ |
| Fast-forward merged into `feature/excel-agent-foundation` | _pending_ | ⏳ |
| Pushed `feature/excel-agent-foundation` (triggers Render auto-deploy) | _pending_ | ⏳ |
| Client PC refreshed via PS agent | _pending_ | ⏳ |
| Manual test passed on client PC | _pending_ | ⏳ |
| Approved by client | _pending_ | ⏳ |

## Problem

Fix A added a preflight popup for Run MIH (with two checks: Option 1/4 + Utilities). Client requested the same preflight for **Run UAP**, but with a single check (utilities only — UAP has no Option 1/4 selection).

User spec 2026-04-29: *"i forgot beofre for uap can we also add before run uap a popup that just asked if utlitys was added same as mih but without the 1-4 q"*

## Code change

**File:** `excel-addin/src/AMI_Optix_Ribbon.bas`
**Function:** `Ribbon_RunSolverUAP` (lines 85-88 in the original file, 85-104 after this change)

**Before:**
```vba
Public Sub Ribbon_RunSolverUAP(control As IRibbonControl)
    RunOptimizationForProgram "UAP"
    EnsureAMIOptixTabActive
End Sub
```

**After:**
```vba
Public Sub Ribbon_RunSolverUAP(control As IRibbonControl)
    ' Pre-flight reminder before running UAP: confirm Utilities filled in.
    ' Pure reminder - does not auto-detect; the client confirms manually
    ' so they get the prompt every time. Mirrors the MIH preflight (Fix A)
    ' but only checks Utilities (UAP has no Option 1/4 selection).
    Dim msg As String
    msg = "Before running UAP, please confirm:" & vbCrLf & vbCrLf & _
          "  [ ]  Utilities are filled in (Settings > Utilities)" & vbCrLf & vbCrLf & _
          "Click YES if done - UAP will run." & vbCrLf & _
          "Click NO to cancel and complete the missing item."
    If MsgBox(msg, vbYesNo + vbInformation, "Run UAP - Pre-flight") = vbNo Then
        MsgBox "Please fill in Utilities, then click Run UAP again.", _
               vbInformation, "Run UAP cancelled"
        EnsureAMIOptixTabActive
        Exit Sub
    End If

    RunOptimizationForProgram "UAP"
    EnsureAMIOptixTabActive
End Sub
```

**Net change:** +13 insertions, 0 deletions. No other functions or files touched.

`Ribbon_RunSolver` (line 79-83, legacy callback that also runs UAP) was **not** modified — it's not wired to the visible "Run UAP" ribbon button. Only `Ribbon_RunSolverUAP` is.

MIH path (`Ribbon_RunSolverMIH`) is untouched — Fix A's two-check popup stays.

## Manual test (after deploy)

1. Push branch + ff-merge to feature + push feature (triggers Render redeploy — VBA-only change so server functionally unchanged).
2. Run the standard PS one-liner on the client PC.
3. Close + reopen Excel + a UAP workbook.
4. Click **Run UAP** on the AMI Optix ribbon.
5. ✅ Pre-flight popup appears. Title: "Run UAP - Pre-flight". One `[ ]` checklist item (utilities only).
6. Click **No** → reminder popup ("Please fill in Utilities, then click Run UAP again."). UAP does **not** fire.
7. Click **Run UAP** again → click **Yes** → UAP runs as before, scenarios populate normally.
8. Click **X** (close) on the popup → behaves like No.
9. Click **Run MIH** on an MIH workbook → MIH preflight (Fix A) still works, two-check popup shows.

## Rollback

```bash
git checkout feature/excel-agent-foundation
git revert <commit-sha-on-this-branch>
git push origin feature/excel-agent-foundation
```

Then run the standard PS one-liner on the client PC.

## One-line PowerShell command for client-PC update

```
powershell -ExecutionPolicy Bypass -File C:\AMI_Optix_Agent\scripts\Refresh-AmiOptixAgentFromGitHub.ps1 -AgentRoot C:\AMI_Optix_Agent
```

(Same one-liner as every other fix — also covers Fix D.2's column rename + ceiling deploy.)
