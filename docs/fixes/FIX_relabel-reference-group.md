# FIX: Relabel the extra-40% group "FOR REFERENCE ONLY"

**Branch:** `fix/relabel-reference-group`
**Date:** 2026-09-02
**Risk:** Low. VBA-only, one display string. No server change, no ribbon change.

## Symptom (client, Building D feedback)

Options in the "MAX RENT / OTHER OPTIONS" group use MORE 40% AMI units than the
legal minimum (they exist as what-if comparisons), but the group name reads like
a recommendation. Client: "I don't want that many units at 40% AMI."

## Change

`AMI_Optix_ResultsWriter.bas:515` `groupNames` G3:
"MAX RENT / OTHER OPTIONS" -> "FOR REFERENCE ONLY - USES MORE 40% THAN REQUIRED"
(+ doc-comment noting the label contracts).

## Why this is safe (from the full-source coupling audit)

- Group membership is computed from DATA (40%-unit counts / shares / keys at
  :538-570), never from the label text. The array is consumed at exactly one
  site (:657) and the label text is written in three places (detail banner
  "GROUP: <label>" :368, bare overview row :812, Ribbon picker "--- <label> ---"
  :1591) — and read back by its own text NOWHERE.
- Sheet scanners match only generic prefixes; the new label honors all three
  contracts: does not start with "SCENARIO ", not numeric, does not start
  with "=". ASCII hyphen used (no smart dashes — .bas files are ANSI).
- Old sheets refreshed by the new add-in: banner rows are rewritten on every
  CreateScenariosSheet; stale bare labels in an old overview are ignored by
  the IsNumeric guard, same as before.

## Verification

- Grep: zero remaining references to the old label in excel-addin/.
- Scenario keys, ordering, recommended_key flow: untouched.
- Sandbox Excel QA (with Fix 1): run UAP + MIH, re-run, picker list shows new
  group title, Apply from picker, Ctrl+Z, live-sync edit.

## Deploy

Same staged module-swap as FIX_remove-how-built-legend (ship together).
