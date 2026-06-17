<#
  Deploy-AmiOptixFixes.ps1

  One-shot deploy for the current AMI Optix VBA fixes. Swaps ONLY the changed
  modules into a copy of the Z: master AMI_Optix.xlam, then installs it to this
  PC. The form and ribbon are never touched. The master is backed up and only
  replaced after every module is patched AND verified.

  Modules applied (pinned to commit c6da60c):
    - AMI_Optix_AppEvents      AMI edits never write -> native Ctrl+Z restored
                               (also subsumes the paste FLOOR/BED/NET SF fix)
    - AMI_Optix_ResultsWriter  FEWEST grouping fix; Manual Calculate skips
                               market-rate (0% AMI) rows (fixes API 500) and
                               formats the AMI column as % on the button
    - AMI_Optix_EventHooks     OnKey reset no longer crashes (1004) + logging

  Run on a client PC that has the Z: drive mapped and Excel installed:
    irm <raw-url-to-this-script> | iex

  Safe to run on multiple PCs (idempotent): the first run patches the master,
  later runs re-apply the same modules and refresh each PC's local copy.
#>

$ErrorActionPreference = 'Stop'

# --- Config ---------------------------------------------------------------
$Commit  = 'c6da60cca16cb8bde9c1abb64c4ddbbd35d9ae1b'
$BaseUrl = "https://raw.githubusercontent.com/martin10101/nyc-ami-calculator/$Commit"
$Master  = 'Z:\AMI_Optix.xlam'
$Local   = Join-Path $env:APPDATA 'Microsoft\AddIns\AMI_Optix.xlam'
$tmpXlam = Join-Path $env:TEMP 'AMI_Optix_patch.xlam'

# Each module: component name, source path in repo, local temp file, proof-of-fix marker
$Modules = @(
    @{ Name = 'AMI_Optix_AppEvents';     Path = 'excel-addin/src/AMI_Optix_AppEvents.cls';     Temp = (Join-Path $env:TEMP 'AMI_Optix_AppEvents.cls');     Marker = 'passive, NO workbook write' }
    @{ Name = 'AMI_Optix_ResultsWriter'; Path = 'excel-addin/src/AMI_Optix_ResultsWriter.bas'; Temp = (Join-Path $env:TEMP 'AMI_Optix_ResultsWriter.bas'); Marker = 'cnt <= 0 And' }
    @{ Name = 'AMI_Optix_EventHooks';    Path = 'excel-addin/src/AMI_Optix_EventHooks.bas';    Temp = (Join-Path $env:TEMP 'AMI_Optix_EventHooks.bas');    Marker = 'SafeResetCtrlZ' }
)

function Fail($msg) { Write-Host "FAILED: $msg" -ForegroundColor Red; exit 1 }

if (-not (Test-Path $Master)) { Fail "Master not found at $Master - is the Z: drive mapped on this PC?" }

# --- 1) Download the fixed modules ---------------------------------------
Write-Host 'Downloading fixes...' -ForegroundColor Cyan
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
foreach ($m in $Modules) {
    Invoke-WebRequest -Uri "$BaseUrl/$($m.Path)" -OutFile $m.Temp -UseBasicParsing
    # VBA's VBComponents.Import requires CRLF line endings; GitHub serves LF.
    # With LF the .cls header (VERSION 1.0 CLASS / BEGIN ...) spills into the
    # code body and the module won't compile. Normalize to CRLF and strip BOM.
    $raw = [IO.File]::ReadAllText($m.Temp)
    $raw = $raw -replace "`r`n", "`n" -replace "`r", "`n" -replace "`n", "`r`n"
    [IO.File]::WriteAllText($m.Temp, $raw, (New-Object Text.UTF8Encoding($false)))
    if ($raw -notmatch [regex]::Escape($m.Marker)) {
        Fail "Downloaded $($m.Name) is missing its fix marker."
    }
}

# --- 2) Enable programmatic VBA access (needed to import modules) ---------
Get-ChildItem 'HKCU:\Software\Microsoft\Office' -ErrorAction SilentlyContinue |
  Where-Object { $_.PSChildName -match '^\d+\.\d+$' } |
  ForEach-Object {
    $sec = "HKCU:\Software\Microsoft\Office\$($_.PSChildName)\Excel\Security"
    New-Item -Path $sec -Force | Out-Null
    Set-ItemProperty -Path $sec -Name AccessVBOM -Value 1 -Type DWord
  }

# --- 3) Close Excel so the .xlam isn't locked ----------------------------
Get-Process excel -ErrorAction SilentlyContinue | ForEach-Object { $_.CloseMainWindow() | Out-Null }
Start-Sleep -Seconds 2
Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

# --- 4) Patch a TEMP copy of the master ----------------------------------
Write-Host 'Patching add-in (form & ribbon untouched)...' -ForegroundColor Cyan
Copy-Item $Master $tmpXlam -Force

$xl = $null; $wb = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.AutomationSecurity = 3   # msoAutomationSecurityForceDisable - no macros/prompts on open

    $wb = $xl.Workbooks.Open($tmpXlam)
    $proj = $wb.VBProject

    foreach ($m in $Modules) {
        # remove the existing component (by name), then import the fixed file
        for ($i = $proj.VBComponents.Count; $i -ge 1; $i--) {
            $c = $proj.VBComponents.Item($i)
            if ($c.Name -eq $m.Name) { $proj.VBComponents.Remove($c) }
        }
        $proj.VBComponents.Import($m.Temp) | Out-Null

        # verify the imported module contains the fix AND its file header did
        # not leak into the code body (the classic LF line-ending import failure)
        $cm = $proj.VBComponents.Item($m.Name).CodeModule
        $code = $cm.Lines(1, $cm.CountOfLines)
        if ($code -notmatch [regex]::Escape($m.Marker)) {
            throw "Post-import verification failed for $($m.Name) - fix marker not found."
        }
        if ($code -match 'VERSION 1\.0 CLASS' -or $code -match 'Attribute VB_Name') {
            throw "Import corrupted for $($m.Name) - module header leaked into code (line endings)."
        }
        Write-Host "  patched $($m.Name)" -ForegroundColor DarkGray
    }

    $wb.Save()
    $wb.Close($false)
    $xl.Quit()
}
catch {
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($xl) { try { $xl.Quit() } catch {} }
    Fail "VBA patch failed: $($_.Exception.Message)`nMake sure Excel is installed and not running, then retry."
}
finally {
    if ($wb) { [Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
    if ($xl) { [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null }
}

# --- 5) Back up the master, then promote the patched copy ----------------
Copy-Item $Master "$Master.bak" -Force          # rollback copy on Z:
Copy-Item $tmpXlam $Master -Force               # patched master for all PCs
Copy-Item $tmpXlam $Local  -Force               # this PC's installed add-in
Unblock-File $Local

Write-Host ''
Write-Host 'SUCCESS - both fixes applied.' -ForegroundColor Green
Write-Host "  - Z: master updated:  $Master   (backup: $Master.bak)"
Write-Host "  - This PC updated:    $Local"
Write-Host 'Reopen Excel. Other PCs: just re-copy from Z: (close Excel first).'
