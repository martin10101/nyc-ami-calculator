# Fix: Dashboard dependency audit cleanup (2026-08-26)

## What

`npm audit` on `dashboard/` reported **11 vulnerabilities (2 critical, 8 high, 1 moderate)**.
This fix clears 8 of them via in-range patch updates. The remaining 3 are all tied to the
`next` package and are deliberately deferred (see below).

## Changes

- `dashboard/package.json` — added an `overrides` block:
  - `js-yaml: 4.3.1` (patched Jul 31; latest 4.3.2 was published 2026-08-26 — too new per the 6-day rule)
  - `picomatch@^4.0.0: 4.0.5` (patched Jul 2; latest 4.0.7 was published 2026-08-24 — too new)
- `dashboard/package-lock.json` — `npm audit fix` in-range updates:

| Package | Before | After | Severity fixed |
|---|---|---|---|
| tar | 7.4.3 | 7.5.22 | **critical** (hardlink/symlink path traversal) |
| minimatch | 3.1.2 / 9.0.5 | 3.1.5 / 9.0.9 | high (ReDoS) |
| brace-expansion | 1.1.12 / 2.0.2 | 1.1.18 / 2.1.4 | high (DoS) |
| js-yaml | 4.1.0 | 4.3.1 (pinned) | high (proto pollution, DoS) |
| picomatch | 2.3.1 / 4.0.3 | 2.3.2 / 4.0.5 (pinned) | high (ReDoS) |
| nanoid | 3.3.11 | 3.3.18 | high (infinite loop) |
| flatted | 3.3.3 | 3.4.4 | high (proto pollution, DoS) |
| postcss (top-level) | 8.5.6 | 8.5.26 | high (path traversal via sourceMappingURL) |
| ajv | 6.12.6 | 6.15.0 | moderate (ReDoS) |

## Deliberately NOT fixed yet (deferred to ~2026-09-01)

`next@15.5.3` → 15.5.24 (critical: RCE in React flight protocol, GHSA-9qr9-h5gf-34mp)
plus its bundled `postcss@8.4.31` and optional `sharp` — all three clear with the single
`next` bump. **15.5.24 was published 2026-08-25**, one day before this fix, which fails
the 6-day supply-chain rule (never adopt a release less than 6 days old), and no older
fully-patched version exists.

Risk while deferred is low: the dashboard deploys as fully static prerendered pages and
production is served by gunicorn/Flask — no Next.js server process runs, so the
server-side Next.js CVEs are not reachable in production.

**Follow-up:** on or after 2026-09-01, re-check `next@15.5.24` for new advisories
(`npm audit` / GitHub Advisory DB), then bump `next` and `eslint-config-next` to 15.5.24
in `dashboard/package.json`. The two `overrides` pins can be dropped once js-yaml ≥4.3.2
and picomatch ≥4.0.7 are 6+ days old with no advisories.

## Verification

- `npm run build` passes; route sizes and chunk hashes identical to the pre-fix
  Render deploy (2026-08-26 19:23 UTC log) — zero behavior change.
- `npm audit` after fix: 3 vulnerabilities remain (next + bundled postcss + sharp),
  all resolved by the deferred `next` bump.
- All adopted versions were ≥6 days old on the npm registry with no advisories
  as of 2026-08-26.
