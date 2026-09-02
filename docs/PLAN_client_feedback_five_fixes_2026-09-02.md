# PLAN: Five fixes from Rachel's Building D feedback (2026-09-02)

Client feedback (Building D, MIH Option 1): options 1-4 lacked 40% units on upper
stories; 5-6 used a 60% band the owner refuses; 7-10 use more 40% units than
required; option 11 ("YOUR ORIGINAL INPUT") was not her input; remove the
methodology paragraph. Verified against her files 2026-09-02 (all claims correct).
Research basis for the floor rule: no numeric HPD rule exists; HPD Design
Guidelines 2026 §4.1.4-4.1.6 + ZR 27-16(b) anti-segregation clause; see memory
`reference_hpd_band_distribution_research`.

Golden baseline: Building D — 99 units, 25 affordable (pool floors 3-19), input
8@40/11@70/6@80, WAAMI 0.5997, recommended rent $45,374/mo, 40% window 10-12.5%
of residential SF.

Order of implementation: ① → ② → ③ → ④ → ⑤, one branch + one doc + one deploy
each (branch-per-fix workflow). Server changes are opt-in by new payload fields:
absent fields → byte-identical legacy behavior (identity-tested).

---

## Fix ① Remove "HOW THESE OPTIONS ARE BUILT" paragraph

- **Where:** `excel-addin/src/AMI_Optix_ResultsWriter.bas:907` block (VBA only).
- **Change:** delete the 5 text rows from the overview render.
- **Program effect:** none. Cosmetic. Overview block shortens; no keys, no order,
  no server traffic changes.
- **Risk & avoidance:** rows below shift up — verify no code addresses those rows
  absolutely (grep for row anchors near the block); manual QA: run, re-run,
  Manual Calculate, Ctrl+Z on sandbox. Per-PC deploy via module swap; rollback =
  re-pin previous commit.

## Fix ② Relabel the comparison group (options 7-10)

- **Where:** `AMI_Optix_ResultsWriter.bas:515` `groupNames` array + strategy "Why"
  lines; **plus the banner recognizer** (comment at :2862 — parses
  "GROUP: <name>" text) and any other name-matchers.
- **Change:** "MAX RENT / OTHER OPTIONS" → "FOR REFERENCE ONLY — USES MORE 40%
  THAN REQUIRED"; prepend "Comparison only —" to those scenarios' Why lines.
- **Program effect:** display text only; scenario keys (`absolute_best`,
  `max_40_share`...) and ordering unchanged; manual-block sync uses
  `recommended_key`, not names.
- **Risk & avoidance:** KNOWN TRAP: banner text is parsed elsewhere. Mitigation:
  keep the `GROUP: ` prefix contract, update every recognizer in the same
  commit (grep all `.bas` for the old strings first), then sandbox QA of live
  sync + Ctrl+Z + Apply. VBA-only; per-PC rollback as ①.

## Fix ③ Original-input snapshot (kills the false "YOUR ORIGINAL INPUT")

- **Where:** VBA (`AMI_Optix_Automation.bas` payload build + new hidden storage;
  small addition inside ApplyBestScenario write path) + server
  (`app.py` original-scenario builder, ~line 2168).
- **Change:**
  - VBA: hidden very-hidden sheet `AMI_Optix_Baseline` storing unit_id → AMI
    snapshot + a hash of the last program-written AMI column. Run logic: no
    snapshot → capture current column as baseline. Snapshot exists → if current
    column hash == last-program-write hash, user changed nothing → keep
    baseline; else user hand-edited → refresh baseline (rerun-aware, per owner
    direction). Payload gains `units[].original_ami`.
  - Server: `scenarios['original']` built from `original_ami` when present,
    else legacy `client_ami` (back-compat).
- **Program effect:** "YOUR ORIGINAL INPUT" becomes permanently truthful across
  apply→rerun cycles. No other scenario, rent, or constraint changes.
- **Risk & avoidance:** does NOT touch EventHooks/Ctrl+Z; ApplyBestScenario gains
  only a post-write hash-save. Old add-in + new server → legacy path (identity
  test). Golden test: Building D asserts original == 8@40/11@70/6@80 after an
  apply→rerun cycle. Hidden sheet is xlObjectVeryHidden, ignored by DataReader.

## Fix ④ Band picker (program- and option-aware)

- **Where:** VBA settings block (data-validation cells, no ribbon XML / no
  .frm — PowerShell-deployable constraint) + payload `allowed_bands` +
  server validation/clamp + solver (allowed-band filter already exists at
  `solver.py:576`).
- **Change:**
  - Picker menu GENERATED from the server-side per-program rule set, never
    hand-typed: MIH O1 → 40/60/70/80/90/100 (40 locked ON; hard cap 100,
    client rule 2026-05-18 untouched); MIH O4 → 40-135 menu (validation
    requires ≥1 enabled band ≤70 and ≤90-depth reachable); UAP → 40-100
    (40 locked).
  - Option/rerun awareness: preflight (extends existing run-preflight pattern)
    re-reads program + Prog!K4 option on EVERY run; if picker state is stale
    for the current option (e.g. user flipped O1↔O4), the run STOPS with a
    clear message and the menu re-renders — never silently proceeds.
  - Server: intersects `allowed_bands` with its own per-option candidate list
    and hard caps (can only narrow, never widen); infeasible selection → 400
    with explicit reason, surfaced verbatim in Excel.
  - Results header echoes: "Built with your band rules: ...".
- **Program effect:** scenarios use only permitted bands; possibly fewer
  scenarios; rent may be lower than unconstrained (her explicit choice, echoed
  in header). No picker → identical output to today.
- **Risk & avoidance:** three locks (picker validation, server clamp, solver
  domain); client caps live server-side so no workbook can widen them;
  identity test (field absent), all-bands test (== identity), exclusion test
  (Building D minus 60% → no 0.6 anywhere, 10-12.5% window + WAAMI≤60 still
  pass), O4 feasibility test (blocking bands ≤70 → clean error, no silent
  relax).

## Fix ⑤ Floor-spread rule + Reviewer View

- **Where:** VBA DataReader (send `units[].floor`, column exists in MIH sheet) +
  settings toggle + solver constraint + results renderer (new band×floor table).
- **Change:**
  - Solver hard constraint (when enabled AND floor data present): split floors
    spanned by the affordable pool into thirds; every band with ≥3 units must
    place ≥1 unit in each third; optional avg-floor tolerance knob for the 40%
    band. If NO scenario satisfies it, retry without and label results
    "floor-spread could not be satisfied — shown without it" (honest fallback,
    mirrors the existing 40%-window floor-walk pattern).
  - Reviewer View: per-scenario band×floor table in results — exactly what the
    HPD reviewer sees in the stacking charts.
- **Program effect:** the ONLY fix that changes assignments: a few units swap
  bands vertically. Rent ≈ unchanged (regulated rent = band+bedrooms; floors
  are rent-neutral; only SF-quota interplay can cost a little — delta shown,
  never hidden). Toggle OFF or no floor column → identical to today.
- **Risk & avoidance:** ships default OFF; enabled on sandbox first with
  Building D asserting: every band ≥3 units present in each third, 40% band no
  longer confined to floors 4-13, rent within stated delta of $45,374, all
  5 HPD compliance tests + 10-12.5% window still green. MIH-only until UAP
  floor expectations are confirmed. Server+solver opt-in as always.

---

## Cross-cutting safety (applies to every fix)

1. **Identity harness:** all repo MIH/UAP golden workbooks + Building D payload
   through old vs new server with new fields absent → JSON must be identical.
2. **Compliance suite:** existing pytest (incl. 5 HPD MIH tests, Option 4 suite,
   87-test baseline) stays green; each fix adds its own tests.
3. **Rollout:** server first (inert without new fields) → sandbox PC → secondary
   PC (1 day) → Rachel. Per-PC module-swap deploy; rollback = re-pin.
4. **Rollback layers:** git revert single commit / Render redeploy previous /
   per-PC re-pin. Each fix independently revertible.
5. **Untouchables:** EventHooks (Ctrl+Z, paste-normalize), manual-block sync,
   year-change preserve logic — no edits, additive only.
6. **Manual QA checklist** (tests/manual_qa_plan.md) on sandbox before any
   client-facing machine: run, rerun, option flip O1↔O4, year change, Manual
   Calculate, Apply, Ctrl+Z, copy-paste.

Approval gate: owner approves this plan → implement ① and ② first (VBA-only,
cosmetic), then ③, ④, ⑤ each behind its own branch/doc/tests.
