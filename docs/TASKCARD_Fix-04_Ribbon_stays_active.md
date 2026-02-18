# Task Card — Fix-04 Ribbon stays active (no collapsing)

## Goal
- Prevent the **AMI Optix** ribbon tab from collapsing/unselecting after ribbon interactions.
- Ensure dropdown interactions don’t use modal UI that can disrupt ribbon focus.

## Acceptance Criteria (from `docs/FIX_REQUIREMENTS.md`)
- Selecting AMI ribbon tab behaves like normal Excel tabs: it stays active while user works.
- Dropdown selection and button clicks must not cause the tab to collapse.
- Must work consistently for all ribbon buttons (MIH + UAP).

## Success Criteria
- After clicking any AMI Optix control (Run UAP/MIH, Live Sync toggle, Manual Calculate, dropdown selection, etc.), the active ribbon tab remains **AMI Optix**.
- Rent roll dropdown selection does not show a modal message box.

## Files to Change
- `excel-addin/customUI/customUI14.xml`
- `excel-addin/src/AMI_Optix_Ribbon.bas`
- `docs/CODEX_LEDGER.md` (work log update)

## Functions / Entrypoints
- RibbonX `onLoad` callback: `Ribbon_OnLoad`
- Helper: `EnsureAMIOptixTabActive`
- Ribbon callbacks (onAction): `Ribbon_*` subs in `AMI_Optix_Ribbon.bas`

## Proposed Patch (logic-level)
1) **Capture Ribbon UI on load**
   - Add `onLoad="Ribbon_OnLoad"` to the Ribbon XML so VBA can store the `IRibbonUI` handle.

2) **Re-activate AMI Optix tab after actions**
   - Add `EnsureAMIOptixTabActive` that calls `IRibbonUI.ActivateTab "tabAMIOptix"` (best-effort, no errors surfaced).
   - Call `EnsureAMIOptixTabActive` at the end of each `onAction` ribbon callback so the tab stays selected after the action completes (even if the action shows modal UI).

3) **Dropdown callback: no modal UI**
   - Remove `MsgBox` from `Ribbon_SelectRentRoll` (keep selection state + optional debug log).

## Risk Pre-mortem
- **Risk:** If `IRibbonUI` is not available (unexpected Office version/reference issues), calls to activate the tab could error.
  - **Mitigation:** Use best-effort `On Error Resume Next` and keep behavior unchanged when ribbon UI isn’t available.
- **Risk:** Re-activating the tab could feel “sticky” if invoked too broadly.
  - **Mitigation:** Only call after AMI Optix ribbon actions (not on worksheet events / selection changes).

## Test Plan
- Manual (Excel):
  1. Open Excel workbook with AMI Optix add-in installed.
  2. Click **AMI Optix** tab.
  3. Click **Run UAP** then click a worksheet cell → tab remains **AMI Optix**.
  4. Click **Run MIH** then click a worksheet cell → tab remains **AMI Optix**.
  5. Use **Select Rent Roll** dropdown then click a worksheet cell → tab remains **AMI Optix** and no modal selection MsgBox appears.
  6. Toggle **Live Sync** and click worksheet cells → tab remains **AMI Optix**.

## Rollback Plan
- Revert the fix commit(s) and rebuild the `.xlam` with the prior ribbon XML/module versions.

