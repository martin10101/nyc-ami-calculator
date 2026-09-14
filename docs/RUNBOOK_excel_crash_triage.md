# RUNBOOK: "Excel has run into an error" on a client PC

Five-minute triage for the repair-prompt crash dialog. Read-only: the script
changes nothing, writes nothing to the registry, and never launches Excel.

## Run it (client PC, ScreenConnect)

```powershell
irm https://raw.githubusercontent.com/martin10101/nyc-ami-calculator/main/tools/excel-agent/Diagnose-ExcelCrash.ps1 | iex
```

(Use the pinned-commit URL handed out with the support request if one was given.)

It prints a report and saves it to `%TEMP%\ExcelTriage_<stamp>.txt`.

## Read three things

1. **Section 2 - faulting module.** This names the component that broke.
2. **Section 5 - Resiliency.** If `AMI_Optix.xlam` is listed, Excel already
   blamed and disabled the add-in.
3. **Section 7 - installed build.** `NEW` = has `AMI_Optix_Baseline`
   (2026-09 snapshot fix); `OLDER` = pre-snapshot build.

## Faulting module -> cause -> action

| Module in section 2 | Means | Next step |
|---|---|---|
| `VBE7.DLL` / `VBE6.DLL` | VBA engine - an .xlam is crashing | Rename-the-add-in test below; then bisect by re-pinning the deploy to the previous commit |
| `EXCEL.EXE` | Excel itself - usually a damaged workbook | Open the workbook with `excel /safe`; if clean, Save As a new .xlsb |
| `oart.dll`, `d3d*`, `igd*`, `nvd*`, `atig*`, `DXGI` | Graphics/rendering | Excel > Options > Advanced > Display > "Disable hardware graphics acceleration" |
| `mso*.dll`, `KERNELBASE.dll` | Generic Office fault | Office Quick Repair; check C2R build in section 1 against a known-good PC |
| *(no crash events at all)* | Dialog is a handled alert, not a hard crash | Check section 3 (OAlerts) for the exact message and timestamps |

## Why "no AMI logs" does NOT clear the add-in

The add-in loads at Excel **startup** (`Auto_Open` in `AMI_Optix_EventHooks`,
which wires the `AMI_Optix_AppEvents` application event sink). It is live for
every sheet edit and sheet activation whether or not anyone clicks Run Solver.
`DebugLog` only writes once the add-in is running, so a crash at load - or a
crash inside the event sink - leaves no AMI log at all.

## 60-second confirm test (reversible)

1. Close Excel.
2. `ren "$env:APPDATA\Microsoft\AddIns\AMI_Optix.xlam" AMI_Optix.xlam.off`
3. Open the same workbook.
   - No crash -> the add-in is the cause.
   - Still crashes -> the workbook or Excel itself.
4. Rename back to undo.

## If the add-in is the cause

Roll back that PC only by re-pinning the deploy one-liner to the previous
commit (see `reference_xlam_deploy_pipeline`); other PCs are untouched because
each has its own `%APPDATA%\Microsoft\AddIns` copy.
