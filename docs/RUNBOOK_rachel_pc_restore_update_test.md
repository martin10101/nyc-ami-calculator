# RUNBOOK: Rachel's PC - bring AMI Optix back, update it, test it (2026-09-22)

Context: on 2026-09-14 Excel itself was crashing on her PC (add-in proven
innocent). The add-in was renamed off and her Excel settings profile was
deleted (backup `ExcelSettings_backup.reg` on her Desktop). Her copy is the
June 16 build. This runbook restores it, updates it to the current build
(fixes 1-6), and tests it. Do it over ScreenConnect with Excel closed.

## Part A - is Excel itself healthy? (2 minutes)

1. Open Excel, blank workbook, type a few numbers, save it to the Desktop as
   `excel-test.xlsx`, close Excel.
2. If "Excel has run into an error" appears at any point: STOP. Excel still
   needs the IT repair / Office update. Nothing below will help until then.

## Part B - bring the program file back (2 minutes)

Do NOT double-click `ExcelSettings_backup.reg` (owner decision 2026-09-22: her
PC is fixed; restoring the old profile could bring the old errors back). The
add-in registration that lived in that profile is redone by hand in Part D.

3. Open PowerShell (Start, type PowerShell, Enter). Paste, Enter:

```
ren "$env:APPDATA\Microsoft\AddIns\AMI_Optix.xlam.off" AMI_Optix.xlam
Remove-Item "$env:APPDATA\Microsoft\AddIns\AMI_Optix_Autofix.xlam.off" -Force
```

   (If the first line says it cannot find the file, the .xlam was already
   renamed back or deleted - carry on, Part C installs a fresh copy.)

## Part C - update to the current build (3 minutes)

4. Excel closed. PowerShell, paste, Enter (needs the Z: drive mapped):

```
irm https://raw.githubusercontent.com/martin10101/nyc-ami-calculator/46df0e3/tools/excel-agent/Deploy-AmiOptixFixes.ps1 | iex
```

   It closes Excel, downloads the fixed modules and ribbon, patches a copy of
   `Z:\AMI_Optix.xlam`, verifies every piece, backs up the master to
   `Z:\AMI_Optix.xlam.bak`, then installs to Z: and to her PC.
5. Last line must be `SUCCESS - modules + ribbon applied.` If it says FAILED,
   nothing was changed anywhere; copy the red line and send it.

## Part D - register the add-in, then test (12 minutes, on a COPY of a building file)

6. Register the add-in (its registration was in the deleted profile). Open
   Excel, File > Options > Add-ins, bottom "Manage: Excel Add-ins" > Go >
   Browse > paste `%APPDATA%\Microsoft\AddIns` in the address bar > pick
   `AMI_Optix.xlam` > OK. If Excel asks to copy it to the AddIns folder,
   answer No (it is already there). Make sure AMI_Optix is ticked > OK.
   Close and reopen Excel.
7. If the add-in asks for the API key (AMI Optix > API Settings), enter it.
   The key normally survives (it lives outside the deleted profile). Then
   Rent Roll Year: pick the year and click Refresh Rent Tables (Selected
   Year) once.
8. Open Excel. The AMI Optix tab now has a "Bands & Floors" group with
   AMI Bands, Spread Across Floors, 40% Rules. If the WHOLE AMI Optix tab is
   missing, go to Part E.
9. Open a MIH building file, File > Save As `... - test copy.xlsb`, work in
    the copy. Run MIH. Check:
    - the "HOW THESE OPTIONS ARE BUILT" paragraph is gone;
    - one group is labelled FOR REFERENCE ONLY;
    - header has "Floor Spread: ON - the 40% AMI apartments include at least
      one on the lower .., middle .., upper .. floors";
    - each scenario has "Floor Spread (HPD reviewer view)" with
      "Rule satisfied: Yes" and a 40% count above zero on all three rows.
10. Apply the best scenario, then Run MIH again. "YOUR ORIGINAL INPUT" still
    shows the numbers she typed, not the applied ones.
11. AMI Bands: 40% is greyed and ticked; untick 60%; Run MIH. Header shows
    "AMI Bands Allowed: 40%, 70%, 80%, 90%, 100%"; no 60% anywhere. Then
    AMI Bands > Allow all bands.
12. 40% Rules: click two unit rows on the MIH sheet, 40% Rules > Pin selected
    units at 40% (popup names them). Run MIH: those two are 40% in every
    option; header has "40% Rules: pinned at 40%: ...". Then 40% Rules >
    Clear all 40% rules, Run MIH: the line is gone.
13. Untouchables: type in one AMI cell and press Ctrl+Z (it undoes); click
    Manual Calculate once; change the Rent Roll Year once. All as before.
14. Close the test copy without saving. Her real file was never touched.

## Part E - rollback (only if something is wrong)

Excel closed, PowerShell:

```
Copy-Item "Z:\AMI_Optix.xlam.bak" "Z:\AMI_Optix.xlam" -Force
Copy-Item "Z:\AMI_Optix.xlam.bak" "$env:APPDATA\Microsoft\AddIns\AMI_Optix.xlam" -Force
Unblock-File "$env:APPDATA\Microsoft\AddIns\AMI_Optix.xlam"
```

She is back on the June build. Send what you saw.

Notes: the update also refreshes the Z: master, so the other office PCs get
the new build by running the same line (or re-copying from Z:). The web
server is already live and works with old and new add-ins alike.
