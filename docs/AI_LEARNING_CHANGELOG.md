# AI Learning Changelog (Audit Log)

This file is a human-readable log of what changed and why.  
For the machine-readable event log that drives learning, we will use a separate shared folder (path TBD).

## 2026-01-21
- Created `docs/BUILD_TRACKER_ai_learning.md` to track scope, risks, and implementation checklist.
- Confirmed current working branch: `feature/ai-learning` (commit `ac17889`).
- Confirmed solver already supports **soft** premium-weight overrides via `project_overrides` (not yet exposed through `/api/optimize`).
- Backend: `/api/optimize` now accepts `project_overrides` + optional `compare_baseline`, returning a `learning` diff (baseline vs learned).
- Backend: Max Revenue scenario now respects overrides; baseline/learned snapshots include WAAMI + canonical assignments diff.
- Backend tests: added `tests/test_api_optimize_learning.py` and verified passing.
- Excel/VBA: added `excel-addin/src/AMI_Optix_Learning.bas` (learning settings in registry + JSON event logging + weight computation).
- Excel/VBA: optimize payload supports `project_overrides` + `compare_baseline` (see `excel-addin/src/AMI_Optix_API.bas`).
- Excel/VBA: implemented SHADOW semantics (apply baseline to sheet, but keep learned scenarios visible) and scenario-apply logging for training.
- Ribbon: added Learning controls to `excel-addin/customUI/customUI14.xml` and callbacks in `excel-addin/src/AMI_Optix_Ribbon.bas`.
