# FIX: Manual Calculate no longer 500s on market-rate (0% AMI) units

**Branch:** `fix/manual-calc-skip-zero-ami`
**Cut from:** `feature/excel-agent-foundation` @ `6ff86f5`
**Date:** 2026-06-17
**Risk:** Low. One read loop + one write-back line in `RecalculateSolverScenarioRents`. In the normal case (no blank-AMI rows) behavior is identical. No solver/API/undo changes.

## Symptom (client)

After typing AMI values and clicking **Manual Calculate**, a dialog showed `API error 500 / internal server error`. The write-back to the AMI Scenarios sheet still completed.

## Server evidence

```
ValueError: Rent table missing entry for 0% AMI / studio.
  rent_calculator.py _gross_rents_lookup -> raise ValueError(f"... {ami_percent*100:.0f}% AMI / {bedroom_label}.")
POST /api/manual_calculate 500
```

`ami_percent = 0` for a studio (0 BR). A unit with **no AMI** (market-rate / unassigned) was sent to the rent calculator, which has no 0% row. (Not the "raw 50 sent" theory — that would read `5000%`, not `0%`.)

## Root cause ([excel-addin/src/AMI_Optix_ResultsWriter.bas](excel-addin/src/AMI_Optix_ResultsWriter.bas), `RecalculateSolverScenarioRents`)

Manual Calculate re-prices every solver scenario block by reading its rows and POSTing them to `/api/manual_calculate`. That read loop added **every** row, including ones whose AMI cell is blank/zero:

```vba
If IsNumeric(rawAmi) Then amiVal = CDbl(rawAmi)
unit("client_ami") = amiVal   ' 0 when blank
units.Add unit                ' added anyway -> assigned_ami: 0 -> server 500
```

The main Manual Calculate call doesn't hit this because `ReadUnitData` already excludes `ami <= 0` units (so does `VerifyManualRents`). Only this per-scenario re-pricing path lacked the guard. It surfaced now because, after the Ctrl+Z fix, the Manual Calculate **button** (which runs this re-pricing) is the primary refresh trigger.

## Change

In the read loop: normalize `> 2 -> /100`, **skip rows with `ami <= 0`**, and stamp `unit("row")` on each kept unit. In the write-back loop: target `assignRow = units(a)("row")` instead of `headerRow + a`, so rents land on the correct rows even when some were skipped.

```vba
If amiVal <= 0# Then
    dataRow = dataRow + 1            ' market-rate / unassigned: not priceable, skip
Else
    unit("client_ami") = amiVal
    unit("row") = dataRow
    units.Add unit
    dataRow = dataRow + 1
End If
...
assignRow = units(a)("row")           ' was headerRow + a
```

When no row is skipped, `units(a)("row") == headerRow + a`, so existing behavior is unchanged.

## Not changed (and why)

- **Ctrl+Z clears after Manual Calculate** — expected. Manual Calculate is an approved on-demand write point; any VBA write resets Excel's session undo. Edits typed *after* it undo normally. Not a regression.
- **`%` symbol on newly typed MIH cells** — cosmetic, by design: we no longer format the AMI cell while typing (that was the undo killer). Optional follow-up: format the AMI column on Manual Calculate / Run only.
- Server left as-is (a 0% AMI band is genuinely unpriceable; the client shouldn't send it). A server guard could be added as defense-in-depth later.

## Verification

- `Sub`/`Function` balance: 68/68.
- Logic: skipped rows excluded from payload; kept rows write to their stored row; no-skip case identical to before.
- Manual after `.xlam` refresh: a schedule containing a market-rate studio (blank AMI) → Manual Calculate completes with no 500; affordable units still re-price correctly.

## Deploy

VBA-only (one `.bas`) → standard `.xlam` refresh (module swap).
