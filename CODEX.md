# CODEX OPERATING PROTOCOL — NYC AMI Calculator

You are working in an existing production repository. Your job is to implement small, isolated fixes without breaking established logic.

## Repo + Base Branch (must be explicit)
- Repo: martin10101/nyc-ami-calculator
- Base branch for all fixes: perf/api-optimize-speed-2026-02-05

“Codex must consult docs/FIX_REQUIREMENTS.md for acceptance criteria.

## Golden Rules
1) One fix per iteration. Do not combine fixes.
2) Minimal diff. No refactors, renames, formatting sweeps, or dependency changes unless explicitly required.
3) Preserve interfaces and Excel sheet structure unless the fix explicitly changes them.
4) Do not change solver scoring/constraints unless the fix explicitly targets solver behavior.
5) If you think you must touch >3 files, STOP and ask for approval.
6) Never rely on ActiveSheet to locate MIH/UAP targets unless explicitly verified safe.

## Required Workflow (every session)
### Step 0 — Read the Ledger (required)
Open docs/CODEX_LEDGER.md and identify:
- the next unchecked fix in the queue,
- the last completed fix,
- any open risks or follow-ups.

### Step 1 — Create a Branch
Create a branch from the base branch:
- fix/<ID>-<short-slug>

### Step 2 — Produce a Task Card (before coding)
Write a Task Card that includes:
- Goal (2–4 bullets)
- Success Criteria
- Files to change (exact paths)
- Functions/entrypoints to change
- Proposed patch (logic-level description)
- Risk pre-mortem (how this could break things + mitigations)
- Test plan (exact steps)
- Rollback plan

Do not code until the Task Card is written.

### Step 3 — Implement
- Only implement what was proposed in the Task Card.
- Keep changes localized.
- Add/adjust tests only when lightweight and aligned with repo.

### Step 4 — Verify
- Run tests if possible.
- If you cannot run tests, provide exact manual test steps and expected outputs.

### Step 5 — GitHub Push + PR (required)
After committing:
1) Push the branch to origin.
2) Open a draft PR targeting the base branch (perf/api-optimize-speed-2026-02-05).
3) Put the Fix ID in the PR title (e.g., "Fix-04: Ribbon stays active").

### Step 6 — Update Ledger (required)
Update docs/CODEX_LEDGER.md with:
- Repo + base branch
- Work branch + commit SHA
- PR number/link (or "no PR")
- Files changed
- Summary of changes
- Tests run + results
- Remaining notes/risks
- Render deploy status (see below)
Then STOP. Do not start the next fix.

---

## Render Deployment Rule (GitHub-backed service)
We want deterministic deployments tied to a commit SHA, not branch-hopping.

### Current setting (do not change)
- Auto-deploy: OFF
- Service ID: srv-d37n6her433s73et1bp0

### Behavior required from Codex
- Do NOT trigger deployments automatically.
- Only record the ready-to-deploy commit SHA in docs/CODEX_LEDGER.md and tell the user which commit/PR to deploy.

Preferred method (manual by user):
- User deploys from Render dashboard when ready.
Optional method (only if user provides deploy hook secret out-of-band):
- Trigger deploy via Render deploy hook with `ref=<commit_sha>` (deploy specific commit).
  - Note: deploying a specific commit may disable auto-deploys until reenabled. (Render docs)

Render MCP note:
- Render MCP currently does NOT support triggering deploys or modifying existing services (except env vars). Therefore do NOT rely on MCP for branch switching or deploy triggers. (Render docs)

---

## Known repo-specific findings (use these to guide safe fixes)
### 1) Scenario apply uses cached DataSheet + AMI column
- ApplyScenarioByKey(...) in excel-addin/src/AMI_Optix_ResultsWriter.bas calls GetDataSheet() and GetAMIColumn()
- Those are cached in excel-addin/src/AMI_Optix_DataReader.bas (m_DataSheet, m_AMICol)
- FindDataSheet() currently prefers ActiveSheet first — this can cause wrong targeting after UI interactions

### 2) Ribbon dropdown currently activates sheets + shows MsgBox
- Ribbon_SelectRentRoll(...) in excel-addin/src/AMI_Optix_Ribbon.bas uses .Activate and MsgBox
- Avoid doing this in dropdown callbacks; it can cause context drift and ribbon focus issues

### 3) Utilities already support variants in the UI + payload
- excel-addin/forms/frmUtilities.frm includes variant codes
- excel-addin/src/AMI_Optix_API.bas BuildAPIPayload sends electricity/cooking/heat/hot_water
- “only 4 utilities” symptom is likely display/mapping, not capture

---

## Ledger Template (create if missing)
(If docs/CODEX_LEDGER.md does not exist, create it using the template in this section.)

# CODEX LEDGER

## Repo + Base Branch
- Repo: martin10101/nyc-ami-calculator
- Base branch: perf/api-optimize-speed-2026-02-05

## Render (optional but recommended)
- Service name:
- Environment:
- Service ID: srv-d37n6her433s73et1bp0
- Auto-deploy setting (On Commit / After CI / Off): OFF
- Deploy hook URL: [REDACTED]

## Fix Queue
- [ ] Fix-01 DataReader stability + remove ribbon sheet activation (prevents wrong-sheet/AMI-col bugs)
- [ ] Fix-04 Ribbon stays active (no collapsing; no modal MsgBox in dropdown callbacks)
- [ ] Fix-03 Rent roll YEAR selector + Manage uploads + mismatch warning (MIH + UAP)
- [ ] Fix-01b Utilities variant breakdown (top-of-sheet; do not collapse variant labels)
- [ ] Fix-02 Solver dedupe identical outcomes + placement tie-break (Python solver)
- [ ] Fix-05 Manual working-copy always-on + two-way sync to MIH/UAP + rent roll + year (requires event modules / sheet code)

## Work Log
### YYYY-MM-DD Fix-XX <title>
- Repo:
- Base branch:
- Work branch:
- Commit:
- PR:
- Files changed:
- Summary:
- Tests run + results:
- Render deploy: (manual; ready commit SHA recorded)
- Notes / risks:
- Next:
