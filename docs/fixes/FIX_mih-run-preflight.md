# FIX: MIH — Run preflight popup (reminder before running MIH)

**Branch:** `fix/mih-run-preflight`
**Cut from:** `feature/excel-agent-foundation` @ `46d1148`
**Date:** 2026-04-29
**Author:** client request via remote session
**Risk:** Low (single ribbon callback, MIH only, UAP untouched)

## Status

| Step | When | Result |
|---|---|---|
| Committed locally | 2026-04-29 — sha `_pending_` | ⏳ |
| Pushed `fix/mih-run-preflight` | _pending_ | ⏳ |
| Fast-forward merged into `feature/excel-agent-foundation` | _pending_ | ⏳ |
| Pushed `feature/excel-agent-foundation` (triggers Render auto-deploy) | _pending_ | ⏳ |
| Render deploy triggered | _pending — not waiting per workflow_ | ⏳ |
| Client PC refreshed via PS agent | _pending — user runs the one-liner_ | ⏳ |
| Manual test passed on client PC | _pending — user follows the test checklist below_ | ⏳ |
| Approved by client | _pending_ | ⏳ |

This change is VBA-only — Render serves the Python backend, which is unchanged.

## Problem

The client wants a reminder popup before "Run MIH" actually fires, so
they don't accidentally launch a run when:

1. **Option 1 or Option 4 isn't selected** on the MIH project sheet (a
   project-level setting; whole project is one option or the other), or
2. **Utilities aren't filled in** (Settings → Utilities ribbon dialog).

Per user spec: *"the popup for the run mih will have both q in one pop
up with a little box they can enter x for each that it was done or
left empty if 1 left empty it wont alow it to run till clenit checks
both."* And *"what it doesnt have do do is tell the client where to
find the utility box or 1-4 the client knows where it is its just for
a reminder."*

This is a **behavioral nudge**, not a technical validation. The existing
data-layer check in [excel-addin/src/AMI_Optix_Main.bas:812](../../excel-addin/src/AMI_Optix_Main.bas#L812)
`TryReadMIHInputs` already errors out if the `Prog` sheet is missing —
that's preserved unchanged and runs *after* the popup if the user clicks
Yes.

UAP is unaffected — `Ribbon_RunSolverUAP` is its own callback at lines 85-88.

## Code change

**File:** `excel-addin/src/AMI_Optix_Ribbon.bas`
**Function:** `Ribbon_RunSolverMIH` (lines 90-93 in the original file, 90-109 after this change)

**Before:**
```vba
Public Sub Ribbon_RunSolverMIH(control As IRibbonControl)
    RunOptimizationForProgram "MIH"
    EnsureAMIOptixTabActive
End Sub
```

**After:**
```vba
Public Sub Ribbon_RunSolverMIH(control As IRibbonControl)
    ' Pre-flight reminder before running MIH: confirm Option 1/4 selected
    ' and Utilities filled in. Pure reminder - does not auto-detect; the
    ' client confirms manually so they get the prompt every time.
    Dim msg As String
    msg = "Before running MIH, please confirm:" & vbCrLf & vbCrLf & _
          "  [ ]  Option 1 or Option 4 is selected on the MIH sheet" & vbCrLf & _
          "  [ ]  Utilities are filled in (Settings > Utilities)" & vbCrLf & vbCrLf & _
          "Click YES if both are done - MIH will run." & vbCrLf & _
          "Click NO to cancel and complete the missing item(s)."
    If MsgBox(msg, vbYesNo + vbInformation, "Run MIH - Pre-flight") = vbNo Then
        MsgBox "Please select Option 1 or 4 and fill in Utilities, then click Run MIH again.", _
               vbInformation, "Run MIH cancelled"
        EnsureAMIOptixTabActive
        Exit Sub
    End If

    RunOptimizationForProgram "MIH"
    EnsureAMIOptixTabActive
End Sub
```

**Net change:** +13 insertions, 0 deletions. No other functions or files touched.

### Why a `MsgBox` and not a real two-checkbox UserForm?

Native VBA `MsgBox` cannot render interactive checkboxes — that requires
a UserForm (`.frm` + `.frx`), which the PowerShell agent does NOT
auto-deploy (per [docs/EXCEL_AGENT_FOUNDATION.md:8-18](../EXCEL_AGENT_FOUNDATION.md#L8-L18)).
Adding a UserForm would force a one-time manual VBA-editor step on the
client PC, which violates the "PowerShell-deployable changes only" rule
locked in for this project.

The `MsgBox` compromise renders the two checkboxes as visual `[ ]`
markers in the message body and uses YES/NO buttons. Functionally
identical to two literal checkboxes (both must be confirmed → run;
otherwise → cancel + reminder), but visually less interactive. Per user
spec, the popup is a pure reminder — if the client clicks YES without
actually doing the items, that's their problem.

ASCII characters are used throughout the message text (no em-dashes,
no Unicode arrows) to avoid garbled rendering in some Excel locales.

## Manual test (before merging back)

1. Push branch → ff-merge → push `feature/excel-agent-foundation` → run
   the standard one-line PS command on client PC.
2. Close + reopen Excel + a MIH workbook (e.g., `230 Kent_Unit Schedule - MIH v3.xlsm`).
3. Click **Run MIH** on the AMI Optix ribbon.
4. ✅ Pre-flight popup appears. Title bar: "Run MIH - Pre-flight".
   Body lists the two `[ ]` checklist items + Yes/No instructions.
5. Click **No** → reminder popup appears ("Please select Option 1 or 4
   and fill in Utilities, then click Run MIH again."). Run MIH does
   **not** fire. AMI Optix tab stays active.
6. Click **Run MIH** again → click **Yes** → MIH runs as before;
   scenarios populate normally.
7. Click **X** (close) on the popup → behaves like No (Run MIH cancels,
   reminder shows).
8. Click **Run UAP** on a UAP workbook → **no popup**, UAP runs as
   before. Confirms UAP is fully untouched.
9. Confirm `TryReadMIHInputs` still fires its existing "Prog sheet
   missing" error if you click Yes on the popup but the Prog sheet
   really is missing. (Just preserving existing behavior — should not
   change.)

## Rollback

```bash
git checkout feature/excel-agent-foundation
git revert <commit-sha-on-fix-branch>
```

Then re-run the PS one-liner on the client PC to regenerate the staged
`.xlam` from the (now-reverted) source.

## One-line PowerShell command for client-PC update

Run on the client PC (cmd.exe, Run dialog, or PowerShell):

```
powershell -ExecutionPolicy Bypass -File C:\AMI_Optix_Agent\scripts\Refresh-AmiOptixAgentFromGitHub.ps1 -AgentRoot C:\AMI_Optix_Agent
```

This pulls the latest `feature/excel-agent-foundation` from GitHub,
runs preflight, rebuilds the staged `.xlam`, deploys it to
`AppData\Roaming\Microsoft\AddIns\`, and runs the acceptance suite.

## Deploy notes

- Render auto-deploy is ON. Pushing `feature/excel-agent-foundation`
  will trigger a Render deploy. The change is VBA-only — Render serves
  no VBA — so the deploy adds nothing functional, but it does happen.
- Client-PC rollout is via PowerShell agent only; no manual VBA copy /
  paste in the Excel VBA editor.
