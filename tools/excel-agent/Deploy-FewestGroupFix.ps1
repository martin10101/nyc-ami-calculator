<#
  Deploy-FewestGroupFix.ps1

  One-shot deploy for the "FEWEST group name-net" fix (fix/fewest-group-name-net).
  Swaps ONLY the AMI_Optix_ResultsWriter module into a copy of the Z: master
  AMI_Optix.xlam, then installs it to this PC. The form and ribbon are never
  touched. The master is backed up and only replaced after a verified patch.

  Run on a client PC that has the Z: drive mapped and Excel installed:
    irm <raw-url-to-this-script> | iex

  Safe to run on multiple PCs (idempotent): the first run patches the master,
  later runs re-apply the same module and refresh each PC's local copy.
#>

$ErrorActionPreference = 'Stop'

# --- Config ---------------------------------------------------------------
$BasUrl  = 'https://raw.githubusercontent.com/martin10101/nyc-ami-calculator/99efdd423a06b411e89b6788857a659a762978e5/excel-addin/src/AMI_Optix_ResultsWriter.bas'
$Master  = 'Z:\AMI_Optix.xlam'
$Local   = Join-Path $env:APPDATA 'Microsoft\AddIns\AMI_Optix.xlam'
$ModName = 'AMI_Optix_ResultsWriter'
$Marker  = 'cnt <= 0 And'        # text proving the fix is present
$tmpBas  = Join-Path $env:TEMP 'AMI_Optix_ResultsWriter.bas'
$tmpXlam = Join-Path $env:TEMP 'AMI_Optix_patch.xlam'

function Fail($msg) { Write-Host "FAILED: $msg" -ForegroundColor Red; exit 1 }

if (-not (Test-Path $Master)) { Fail "Master not found at $Master - is the Z: drive mapped on this PC?" }

# --- 1) Download the fixed module ----------------------------------------
Write-Host 'Downloading fix...' -ForegroundColor Cyan
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Invoke-WebRequest -Uri $BasUrl -OutFile $tmpBas -UseBasicParsing
if ((Get-Content $tmpBas -Raw) -notmatch [regex]::Escape($Marker)) { Fail 'Downloaded module is missing the fix marker.' }

# --- 2) Enable programmatic VBA access (needed to import a module) --------
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

    for ($i = $proj.VBComponents.Count; $i -ge 1; $i--) {
        $c = $proj.VBComponents.Item($i)
        if ($c.Name -eq $ModName) { $proj.VBComponents.Remove($c) }
    }
    $proj.VBComponents.Import($tmpBas) | Out-Null

    # verify the imported module actually contains the fix
    $cm = $proj.VBComponents.Item($ModName).CodeModule
    $code = $cm.Lines(1, $cm.CountOfLines)
    if ($code -notmatch [regex]::Escape($Marker)) { throw 'Post-import verification failed - fix marker not found in module.' }

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
Write-Host 'SUCCESS - fix applied.' -ForegroundColor Green
Write-Host "  - Z: master updated:  $Master   (backup: $Master.bak)"
Write-Host "  - This PC updated:    $Local"
Write-Host 'Reopen Excel. Other PCs: just re-copy from Z: (close Excel first).'
