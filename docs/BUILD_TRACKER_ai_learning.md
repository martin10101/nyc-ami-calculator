# Build Tracker: AI Learning (Preferences + Logging)

Branch: `feature/ai-learning`  
Rule: do **not** merge to `main` until approved.

## What “Learning” Means (so we don’t break the solver)

**Hard rules stay hard.** Learning must NOT change:
- UAP vs MIH program rules (never mixed)
- Allowed bands (no 50%)
- WAAMI caps/floors (including Max Revenue floor `58.9%` only for that scenario)
- Deep affordability/share constraints
- Band caps (UAP max 100%, MIH caps per option)
- Max bands per scenario

**Learning only adjusts soft preferences**, like:
- What “premium” means (floor vs SF vs beds vs balcony weighting)
- Scenario ranking/recommendation (never silently relax constraints)

## Goal (Client Value)

Over time, the tool should:
- Prefer placing higher AMI bands on the kinds of units your office consistently treats as “premium”
- Keep results compliant, but closer to what your team actually chooses
- Log exactly what changed (baseline vs learned) so behavior is auditable

## Modes (per profile: UAP / MIH Option 1 / MIH Option 4)

1. **OFF**: default solver behavior.
2. **SHADOW**: runs learned solve + baseline compare; **applies baseline** to the sheet; logs diff.
3. **ON**: uses learned preferences; can optionally compare against baseline and log diff.

## Storage (per-client / office)

Current approach: **shared-folder event log** (one JSON file per event; no shared-file appends).

- Default: `%USERPROFILE%\Documents\AMI_Optix_Learning`
- Can be set to a shared drive/path (e.g. `Z:\AMI_Optix_Learning` or `\\server\share\AMI_Optix_Learning`)

## What we log (append-only events)

- `solver_run`: mode, weights sent, scenario count, baseline vs learned WAAMI, changed unit count
- `scenario_applied`: which scenario was applied (AUTO or USER) + unit features + assigned AMI

Planned next:
- `final_selection`: “Final Selection” + “Reason/Notes”
- `manual_edit`: “Validate/Update” custom changes (not on every keystroke)

## Critique (risks + mitigations)

1) **Learning breaks compliance**  
Mitigation: learning only affects premium weights; constraints stay enforced server-side.

2) **Noisy data causes drift**  
Mitigation: bounded weights, minimum-data threshold, and logging + compare-baseline QA.

3) **Multi-user office write conflicts**  
Mitigation: write one JSON file per event with GUID filenames (no file-lock contention).

4) **Black box behavior**  
Mitigation: solver_run logs baseline vs learned WAAMI + count of changed units.

## Implementation Checklist

### Backend (Render)
- [x] Accept `project_overrides` on `/api/optimize` and pass into solver
- [x] Support `compare_baseline` (baseline vs learned diff in one request)
- [x] Ensure Max Revenue scenario respects overrides
- [x] Tests for overrides + compare (`tests/test_api_optimize_learning.py`)

### Excel Add-in (VBA)
- [x] Compute learned premium weights from `scenario_applied` events (USER)
- [x] Send `project_overrides` + `compare_baseline` in optimize payload (when enabled)
- [x] Log `solver_run` and `scenario_applied`
- [x] Implement SHADOW semantics (apply baseline to sheet; show learned scenarios)
- [x] Ribbon “Learning” settings + “Open Logs” buttons (no UserForms)

### QA
- [ ] Verify OFF matches current behavior
- [ ] Verify SHADOW applies baseline + logs diff
- [ ] Verify ON affects only soft ranking, not constraints
