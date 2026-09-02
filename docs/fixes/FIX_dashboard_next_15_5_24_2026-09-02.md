# Fix: Dashboard next 15.5.24 bump — clears the critical (2026-09-02)

Follow-up to `FIX_dashboard_dep_audit_2026-08-26.md`. The deferred `next` bump, executed
on schedule once 15.5.24 passed the 6-day rule (published 2026-08-25, adopted at 8 days old).

## Changes

- `next` 15.5.3 → **15.5.24** — clears the **critical** (RCE in React flight protocol,
  GHSA-9qr9-h5gf-34mp) and the rest of the Next.js advisory batch.
- `eslint-config-next` 15.5.3 → 15.5.24 (kept in lockstep).
- `sharp` 0.34.5 → **0.35.4** (published 2026-08-26, 7 days old ✅) — clears the high
  (libvips CVEs); now fixable in-range without touching `next`.
- Dropped the temporary `overrides` pins for js-yaml/picomatch added on 2026-08-26 —
  their newest patches are now past the 6-day window, so normal semver resolution is safe
  again. Lockfile keeps the already-patched 4.3.1 / 4.0.5 until a routine update moves them.

## Remaining (2 advisories, both the same root cause — deferred)

`postcss@8.4.31` **bundled inside next** (high: sourceMappingURL file-read advisories,
including GHSA-fxqj-rqcc-2cmp published after the Aug 26 cleanup) plus `next` itself
flagged moderate for depending on it. The only fix npm offers is **next@16.3.4 — a major
version, published 2026-08-31 (2 days old)**. Deferred because it fails the 6-day rule
and is a breaking-change major that needs its own migration pass.

Risk assessment: the bundled postcss runs only at build time against our own CSS —
the advisories need attacker-controlled CSS/sourceMappingURL input. The dashboard is
static prerendered output served by gunicorn/Flask; no Next server runs in production.

**Follow-up:** on or after ~2026-09-06, evaluate the next 16.x line (16.3.4 or later)
as a deliberate major upgrade with its own fix branch and build/visual verification —
or wait for a 15.5.x release that bumps the bundled postcss, if one appears.

## Verification

- `npm audit`: 11 → 2 (was 3 before this fix; the critical is gone).
- `npm run build` passes; routes unchanged and fully static
  (First Load JS 111 kB vs 110 kB — patched Next runtime, +0.7 kB framework chunk).
- Adopted versions ≥6 days old with no advisories as of 2026-09-02:
  next 15.5.24 (8d), sharp 0.35.4 (7d).
