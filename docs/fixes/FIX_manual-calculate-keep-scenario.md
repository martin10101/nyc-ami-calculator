# FIX: Manual Calculate keeps the displayed scenario (no swap) + protects Ctrl+Z

**Branch:** `fix/manual-calculate-keep-scenario`
**Cut from:** `feature/excel-agent-foundation`
**Date:** 2026-06-15
**Risk:** Low — one VBA handler, reusing the already-tested keep-displayed/in-place path. No solver, no server, no change to Live Sync or the undo/debounce machinery.

## Symptoms (client, 2026-06-15)

1. **Manual Calculate swapped the manual scenario.** Pressing the button replaced the displayed scenario (e.g. the recommended fewest-40, 8 units) with a different, higher-40%-count layout.
2. **Ctrl+Z locked after Manual Calculate** — including for edits made after the press.

## Root cause

The manual block is a "working copy" of the **MIH input sheet**: `Ribbon_ManualCalculate` → `ManualCalculateScenario(programNorm)` (preserve=False) → `ReadUnitData()` reads the MIH AMI column and **rebuilds** the block from it ([DataReader.bas:22](excel-addin/src/AMI_Optix_DataReader.bas#L22)). But after Run MIH the block is set to the **recommended** scenario, which differs from the raw MIH input. So pressing the button reverted the recommended → MIH input ("swap to higher 40% count"). The server endpoint is purely deterministic ([app.py:2632](app.py#L2632)) — it evaluates the AMIs sent, never optimizes — so the swap was entirely the VBA choosing the wrong source.

Ctrl+Z: Excel's undo is session-wide and is wiped by **any** VBA cell write. The button did a full clear-and-rebuild immediately (no debounce), wiping undo on press; and a debounced refresh armed by a prior edit could fire ~2s later and wipe it again while the user worked.

## Change ([excel-addin/src/AMI_Optix_Ribbon.bas](excel-addin/src/AMI_Optix_Ribbon.bas))

`Ribbon_ManualCalculate`:
1. Calls `ManualCalculateScenario(programNorm, True)` — `preserveAppliedScenario:=True`, the same proven in-place path used by the rent-roll-year change: it re-prices the manual block's **own** displayed bands instead of rebuilding from the MIH input. The displayed scenario stays; only rents refresh.
2. Calls `AMI_Optix_CancelDeferredRefresh` first, so a debounced refresh armed by a recent edit can't fire after the press and re-clear the undo stack.

MIH-column edits still flow into the block via Live Sync, so "reflect MIH edits" is preserved without any new baseline/snapshot state.

## Deliberately NOT changed (Ctrl+Z safety)

- The 2-second edit debounce and AMI input-normalization are untouched — that's what makes Ctrl+Z work today.
- Honest limit: the button press still clears *prior* undo history (Excel wipes the session stack on any macro write; a recompute must write). The fix guarantees **edits after the press stay undoable**, which is the regression being reported.

## Known follow-up (not in scope)

With this minimal version, editing a single MIH AMI cell still triggers Live Sync's full rebuild-from-input, which can pull other units toward the input. Not reported by the client; left alone to avoid risking the working Ctrl+Z path. If it surfaces, the follow-up is to make MIH edits apply surgically (only the edited unit) to the displayed block.

## Verified

- Static checks: Ribbon proc balance (41/41 Sub, 23/23 Function); button passes `True`; old preserve-False call removed; `AMI_Optix_CancelDeferredRefresh` confirmed Public.
- Reuses the keep-displayed path already verified by the rent-roll-year fix.

## Deploy

VBA-only → standard `.xlam` refresh (GitHub Refresh button or the from-GitHub PowerShell one-liner). No Render deploy.
