# Fix: Remove chat-agent tool (security) — 2026-09-02

## What / Why

Deleted `tools/chat-agent/` entirely (added on this branch in `0a855c9` "Add browser
chat agent with OpenAI agentic loop"). A security review found two confirmed
vulnerabilities in it:

1. **HIGH — unauthenticated network-exposed agent server.** `server.js` bound all
   network interfaces with no authentication on any endpoint. Any LAN peer could drive
   the LLM agent via `POST /api/chat`; the agent's tools included `edit_source_file`
   (no path containment — arbitrary file write) and `run_full_refresh` (executes a
   PowerShell script with `-ExecutionPolicy Bypass`) — i.e. remote code execution on
   the machine running the tool. `GET /api/history` also leaked all conversations.
2. **MEDIUM — stored XSS.** `public/index.html` rendered assistant text into
   `innerHTML` unescaped; a poisoned reply persisted to `chat-history.json` would
   execute in the operator's browser on every page load.

Owner decision: the tool was not in use (no `chat-history.json` existed anywhere —
it was never launched), nothing else in the repo references it, and its self-service
purpose isn't currently needed. Removal was chosen over hardening.

## Impact

None on the product. The Flask API, solver, dashboard, and Excel add-in never
referenced `tools/chat-agent/` (verified by repo-wide grep — all matches were
internal to the deleted folder). No deploy/runtime path changes.

## Recovery

The full tool remains in git history (last present at commit `ffadde6` on
`feature/excel-agent-foundation`). If it's ever wanted again, restore with
`git checkout ffadde6 -- tools/chat-agent` — but apply the hardening from the
2026-09-02 security review first: bind 127.0.0.1, per-session auth token,
path containment in file tools, escape-before-render in the UI.

## Rotation assessment (from the same review)

No credential rotation required: the npm dependency vulnerabilities were not
secret leaks, Render env keys were never committed (gitleaks clean on every
commit), and the chat-agent was never run so its network exposure never existed
in practice.
