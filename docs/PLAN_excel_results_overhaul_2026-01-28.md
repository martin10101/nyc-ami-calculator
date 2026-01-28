# AMI Optix (Excel Add-in + API) — Results Overhaul Plan
**Date:** 2026-01-28  
**Scope:** Fix Excel write-back issues, improve scenario layout, add “strict vs edge” scenario generation, and replace noisy in-sheet solver notes with file-based run logs.  
**Non-goal (for now):** Reintroducing API-driven “AI learning” (Render instability). We’ll replace the Learning button with local logging of user-chosen scenarios.

---

## 1) Goals (what success looks like)

### A. Client-facing Excel output is clean
- `AMI Scenarios` sheet **no longer shows** solver/debug notes.
- Scenarios appear **immediately after** the “Scenario Manual (Live Sync)” block (no hard jump to row 125).
- The “best” scenario is **not duplicated** (shown in manual block only; not repeated again as Scenario 1).
- WAAMI/Avg AMI display shows **2 decimals** so values like `59.96%` don’t appear as `60.0%`.

### B. Solver returns ~6 scenarios (3 strict + up to 3 edge)
- **Strict tier (up to 3):** satisfies all “hard compliance” constraints, including:
  - WAAMI cap **never above 60%** (hard rule).
  - WAAMI floor **>= 59.10%** (strict rule).
  - Deep affordability share at/below 40% AMI is **20%–21% by Net SF** (strict rule).
  - **Never** use the **50%** AMI band (hard rule).
  - Any other baked-in solver rules remain unchanged.
- **Edge tier (up to 3):** relaxes constraints **one at a time** to improve rent roll while still:
  - enforcing WAAMI cap **<= 60%** (never changes),
  - never using **50%** AMI band,
  - targeting WAAMI as close to 60% as feasible under the relaxed rule,
  - enforcing an **absolute minimum edge WAAMI floor of 57.50%**.
- Each edge scenario includes a concise “Tradeoffs / Sacrifices” list explaining exactly what was relaxed or violated.

### C. Logs are written to a single append-only file (no file spam)
- Replace “Solver Notes” visibility in Excel with a **single log file** in the Learning log root (often on `Z:\...`).
- Do **not** create a new file per run; instead append entries with timestamps to the same file (JSONL or similar).

---

## 2) Current behavior (what’s wrong and where it lives)

### 2.1 Solver notes are shown in the client sheet
**Where:** `excel-addin/src/AMI_Optix_ResultsWriter.bas` → `CreateScenariosSheet()`  
**Why:** It writes `result("notes")` directly to `AMI Scenarios` starting at `SCENARIOS_START_ROW`.

### 2.2 Scenario output starts at row 125
**Where:** `excel-addin/src/AMI_Optix_ResultsWriter.bas`  
**Why:** Hard constants:
- `MANUAL_BLOCK_HEIGHT = 120`
- `SCENARIOS_START_ROW = 125`

### 2.3 Best scenario duplicated
**Where:** `excel-addin/src/AMI_Optix_ResultsWriter.bas`  
**Why:** Manual block renders “best scenario”; then the scenario loop renders it again.

### 2.4 Avg AMI display rounding confusion
**Workbook reality:** `Calculations!C20` is WAAMI and is formatted as `0.0%` (1 decimal).  
**Effect:** `59.96%` displays as `60.0%`.

### 2.5 “Extreme” results appear (e.g., 56% WAAMI)
**Where:** API `app.py` expansion logic merges “fallback” scenarios that do not meet strict WAAMI floor into the main scenario list.  
**Effect:** The Excel “Available Scenarios” list can include very low WAAMI outcomes that were intended as “closest-feasible” fallbacks.

### 2.6 Learning mode instability (Render)
**Where:** `/api/optimize` baseline/compare + expansion loops + large response payload.  
**Decision:** Do not fix this now; replace Excel “Learning” with local logging of user choice.

---

## 3) Design decisions agreed with the user

### 3.1 Logs
- Do **not** show solver notes on `AMI Scenarios`.
- Do write logs to a **single append-only file** in the Learning log folder (often `Z:\...`).
- File format: **JSONL** (one JSON object per line) to keep it parseable later.

### 3.2 Scenario layout
- Remove the hard row 125 jump: start scenarios immediately after manual block.
- Refactor manual-table event detection to avoid hard-coded “rows 1–120”.
- Skip rendering the best scenario in the lower list.

### 3.3 WAAMI formatting
- Apply `0.00%` formatting to Avg AMI / WAAMI display cells in a **safe, minimally invasive** way:
  - only adjust number formats on known/identified cells,
  - never change formulas or values.

### 3.4 Edge scenario tolerances
- Edge tier can relax the deep affordability share rule:
  - allow <=40% share to go slightly **below 20%** (e.g., down to **19.80%**),
  - allow <=40% share max to increase (up to **23–24%** if it materially improves rent),
  - other rules may be relaxed **one at a time**, but:
    - WAAMI cap <= 60% stays hard,
    - 50% AMI band stays forbidden,
    - edge WAAMI floor >= 57.50%.

---

## 4) Implementation plan (no code yet — just the strategy)

### Phase 1 — Excel: logging + layout cleanup (safe, isolated)

#### 4.1 Replace in-sheet solver notes with file logging
**Files to change:**
- `excel-addin/src/AMI_Optix_ResultsWriter.bas`
- `excel-addin/src/AMI_Optix_Learning.bas` (or add `AMI_Optix_Logging.bas`)

**Approach:**
1. Add a function like `AppendRunLog(eventName As String, payloadJson As String)` that:
   - resolves log root via existing Learning setting (`GetLearningLogRootPath()`),
   - uses a single file path like: `GetLearningLogRootPath() & "\ami_optix_runs.jsonl"`,
   - appends a new JSON line with:
     - `timestamp` (ISO-like string),
     - `eventName` (e.g. `"solver_run"`),
     - `workbookName`, `workbookPath` (if available),
     - `program`, `mih_option` if present,
     - `scenario_keys` returned,
     - `notes` (solver notes),
     - any error flags.
2. In `CreateScenariosSheet()`, remove the block that writes “Solver Notes” to the worksheet.
3. Instead, call `AppendRunLog("solver_notes", ...)` with the notes.

**Risk & mitigation:**
- Network path (Z drive) may be unavailable → fallback to local Documents folder and show a single warning (once per session).
- File growth → optional size-based rotation later (not required now).

#### 4.2 Dynamic scenario start row (remove hard row 125)
**Files to change:**
- `excel-addin/src/AMI_Optix_ResultsWriter.bas`
- `excel-addin/src/AMI_Optix_AppEvents.cls`

**Approach:**
1. Make `WriteManualScenarioBlockFromResult()` return the last row it wrote (end row).
   - Same for `WriteManualScenarioBlockFromEvaluate()`.
2. In `CreateScenariosSheet()`, set:
   - `row = manualEndRow + 2` (or similar padding) instead of `SCENARIOS_START_ROW`.
3. Update the “jump/scroll” logic to scroll to `row` (new scenarios start row).

**Event handler refactor (critical):**
- Remove the hard-coded `If target.Row > 120 Then Exit Function`.
- Compute manual-table bounds dynamically:
  - find the manual table header row (the row where col A is “Unit” and col F is “AMI”),
  - treat the manual table as rows until first blank unit-id row (or until a marker like “SCENARIO”).

#### 4.3 Prevent best scenario duplication
**File to change:**
- `excel-addin/src/AMI_Optix_ResultsWriter.bas`

**Approach:**
- Determine `bestKey` via existing `GetBestScenarioKey(result)`.
- When looping scenarios for output, `If scenarioKey = bestKey Then Skip`.

#### 4.4 Scenario ordering (quality-of-life)
**File to change:**
- `excel-addin/src/AMI_Optix_ResultsWriter.bas`
- `excel-addin/src/AMI_Optix_Ribbon.bas` (scenario picker list order)

**Approach:**
- Build a deterministic ordered key list:
  - strict keys first: `absolute_best`, `best_3_band`, `best_2_band`, `alternative`,
  - then edge keys (once added) in numeric order,
  - then any leftovers alphabetically.

#### 4.5 WAAMI display precision (safe formatting only)
**Files to change:**
- `excel-addin/src/AMI_Optix_ResultsWriter.bas` (post-apply)
- maybe `excel-addin/src/AMI_Optix_Main.bas` / scenario apply flow

**Approach (minimally invasive):**
1. After applying a scenario to `UAP` AMI column, call:
   - `Application.Calculate` (or `CalculateFull` if needed, but prefer basic calculate).
2. Apply number formatting `0.00%` to the workbook’s WAAMI display cells:
   - primary: `Calculations!C20` if it exists,
   - plus any cells in `UAP` used-range whose formula references `Calculations!C20` (format those cells too).
3. Never change formulas; only number formats.

---

### Phase 2 — API: strict vs edge scenario generation + tradeoffs

#### 4.6 Stop “edge” scenarios from polluting strict results
**Files to change:**
- `app.py`
- possibly `ami_optix/solver.py`

**Approach:**
- Restructure `/api/optimize` to explicitly produce:
  1) `strict_scenarios` (using current rules, no relaxation merging)
  2) `edge_scenarios` (generated by controlled relaxation passes)
- Keep response backward compatible:
  - still return a combined `scenarios` dictionary for Excel,
  - but also include metadata fields per scenario (tier + tradeoffs).

#### 4.7 Edge scenario algorithm (relax one rule at a time)
**Hard rules (never change):**
- `waami_cap_percent = 60.0`
- forbid band `50`

**Edge generation outline:**
1. Start from strict config.
2. Attempt `edge_1`: relax only **max_share** (e.g. 0.22 → 0.24), keep min_share=0.20 and waami_floor=0.591.
3. Attempt `edge_2`: relax only **min_share** (e.g. 0.20 → 0.198), keep max_share=0.21 and waami_floor=0.591.
4. Attempt `edge_3`: relax **max_share further** OR relax **waami_floor** only if needed, but never below **0.575**.
5. Each step:
   - run a rent-maximizing solve (use `find_max_revenue_scenario`),
   - require uniqueness vs strict scenarios (by canonical assignments),
   - reject if WAAMI < 0.575,
   - capture what changed (relaxation parameters).

#### 4.8 Tradeoffs (client-readable)
**Approach:**
- For each edge scenario, validate its assignments against **strict rules** and collect the failing checks.
- Convert failures into short “tradeoffs” strings (no debug noise).
- Example outputs:
  - “40% share = 19.82% (below 20.00% minimum by 0.18%)”
  - “40% share = 23.40% (above 21.00% maximum by 2.40%)”
  - “WAAMI = 58.70% (below 59.10% floor by 0.40%)”

---

### Phase 3 — Excel UI: beautiful layout + edge section

#### 4.9 Keep the current scenario detail tables, but improve layout
**Requirement:** Each scenario still shows:
- utilities
- WAAMI
- bands used
- rent totals
- allowances
- full per-unit table (gross rent, net rent, annual, etc.)

**Layout plan:**
- Manual block stays at top (green).
- Strict scenarios section header (blue).
- Edge scenarios section header (orange) with “Tradeoffs” bullets near the top of each scenario.
- Consistent spacing, column widths, and number formats.

---

### Phase 4 — Replace “Learning” button with local choice logging (no API)

#### 4.10 Ribbon behavior change
**Files to change:**
- `excel-addin/customUI/customUI14.xml`
- `excel-addin/src/AMI_Optix_Ribbon.bas`
- `excel-addin/src/AMI_Optix_Learning.bas` (or new logging module)

**Approach:**
- Add a new button: “Record Choice” (or repurpose existing “Learning” button).
- On click:
  1. show scenario list + ask user which scenario they chose (1..N),
  2. read the current workbook’s applied AMI assignments from the sheet,
  3. log a JSON record to the **single append-only log file** on `Z:\...`,
  4. include: project name/address if available from known cells, program, chosen scenario key, WAAMI, band mix, rent totals.

**Key point:** No API calls in this learning action.

---

## 5) Timeouts & performance

### Excel → API timeouts
Current `ServerXMLHTTP` receive timeout is ~120s.  
We can increase to ~180–240s to allow more solver work, but only after:
- strict/edge generation is optimized to avoid runaway passes,
- we confirm Render’s request limits for your plan.

### API-side compute controls
Prefer:
- increasing combo checks and per-scenario time limits in a controlled way,
- limiting the number of edge passes (max 3),
- returning only the scenarios we intend to display (avoid shipping huge debug data).

---

## 6) Safety, stability, and security considerations
- Do not log API keys or secrets (ever).
- Logging must fail “softly”: if writing to Z drive fails, continue solver flow and show a single warning.
- Avoid changing user formulas/values; only adjust number formats for display precision.
- Preserve backward compatibility of API response keys used by VBA (`success`, `scenarios`, `notes`, `project_summary`).

---

## 7) Rollout strategy (avoid breaking existing workflows)
1. Implement Phase 1 (Excel cleanup) and validate on `Unit Schedule v7copy.xlsm`.
2. Implement Phase 2 (strict vs edge API changes) behind a feature flag if needed.
3. Update Excel rendering to show edge section + tradeoffs.
4. Replace Learning behavior last.

---

## 8) Git strategy (don’t risk the current working version)
Preferred: create a new git branch (e.g. `feature/results-overhaul-2026-01-28`) and make all changes there.  
This already prevents “messing up existing files” and is cleaner than duplicating folders.

If you still want an on-disk folder for “change artifacts”, we’ll add a dedicated folder under `docs/` for specs/log samples, but code changes should stay in-place on a branch.

