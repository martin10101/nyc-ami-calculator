<#
  Diagnose-ExcelCrash.ps1

  READ-ONLY triage for "Excel has run into an error..." on a client PC.
  Changes NOTHING: no registry writes, no file changes, Excel is never launched.
  Prints a short report and saves a copy to %TEMP%\ExcelTriage_<stamp>.txt

  Run on the client PC (ScreenConnect / remote session):
    irm <raw-url-to-this-script> | iex
#>

$ErrorActionPreference = 'SilentlyContinue'
$report = New-Object System.Collections.Generic.List[string]

function Say($s)     { Write-Host $s; [void]$report.Add([string]$s) }
function Section($s) { Say ''; Say ('=== ' + $s + ' ===') }
function Hit($s)     { Write-Host $s -ForegroundColor Yellow; [void]$report.Add([string]$s) }

$findings = New-Object System.Collections.Generic.List[string]
$since = (Get-Date).AddDays(-21)

Say '=============== AMI OPTIX - EXCEL CRASH TRIAGE (read-only) ==============='
Say ('PC: ' + $env:COMPUTERNAME + '   User: ' + $env:USERNAME + '   Time: ' + (Get-Date))

# ---------------------------------------------------------------- 1. Excel
Section '1. EXCEL VERSION / STATE'
$xlExe = @(
  "$env:ProgramFiles\Microsoft Office\root\Office16\EXCEL.EXE",
  "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\EXCEL.EXE",
  "$env:ProgramFiles\Microsoft Office\Office16\EXCEL.EXE"
) | Where-Object { Test-Path $_ } | Select-Object -First 1
if ($xlExe) {
    Say ('EXCEL.EXE  : ' + $xlExe)
    Say ('Version    : ' + (Get-Item $xlExe).VersionInfo.ProductVersion)
} else {
    Say 'EXCEL.EXE  : not found in the usual paths'
}
$c2r = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
if ($c2r) { Say ('C2R build  : ' + $c2r.VersionToReport + '  (' + $c2r.Platform + ')') }
$proc = @(Get-Process excel)
if ($proc.Count -gt 0) { Say ('RUNNING NOW: ' + $proc.Count + ' excel.exe process(es)') } else { Say 'RUNNING NOW: no' }

# ------------------------------------------------- 2. Crash events (the key)
Section '2. CRASH EVENTS (Application log, 21 days)  <-- THE ANSWER IS USUALLY HERE'
$crashes = @()
try {
    $crashes = @(Get-WinEvent -FilterHashtable @{LogName='Application'; ProviderName='Application Error'; StartTime=$since} -MaxEvents 300 |
                 Where-Object { $_.Message -match 'EXCEL\.EXE' })
} catch { }
if ($crashes.Count -eq 0) {
    Say 'No "Application Error" crash events for EXCEL.EXE in the last 21 days.'
    Say '(If the dialog still appears, it may be a handled alert - see section 3.)'
} else {
    Say ('Found ' + $crashes.Count + ' Excel crash event(s). Newest first:')
    $modTally = @{}
    $i = 0
    foreach ($e in $crashes) {
        $mod = ''
        $exc = ''
        $m = [regex]::Match($e.Message, 'Faulting module name:\s*([^,\r\n]+)')
        if ($m.Success) { $mod = $m.Groups[1].Value.Trim() }
        $m2 = [regex]::Match($e.Message, 'Exception code:\s*(\S+)')
        if ($m2.Success) { $exc = $m2.Groups[1].Value.Trim() }
        if ($mod -ne '') {
            if ($modTally.ContainsKey($mod)) { $modTally[$mod] = $modTally[$mod] + 1 } else { $modTally[$mod] = 1 }
        }
        if ($i -lt 8) {
            Hit ('  ' + $e.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss') + '   module=' + $mod + '   exception=' + $exc)
        }
        $i = $i + 1
    }
    Say ''
    Say 'Faulting module tally (what is actually breaking):'
    $sortedMods = $modTally.Keys | Sort-Object { - $modTally[$_] }
    foreach ($k in $sortedMods) { Hit ('  ' + $modTally[$k] + ' x  ' + $k) }
    $top = $sortedMods | Select-Object -First 1
    if ($top -match 'VBE7|VBE6') {
        [void]$findings.Add('VBA ENGINE is faulting (' + $top + ') -> an .xlam/VBA add-in is the prime suspect.')
    } elseif ($top -match 'oart|d3d|igd|nvd|atig|DXGI') {
        [void]$findings.Add('GRAPHICS module is faulting (' + $top + ') -> disable Excel hardware graphics acceleration.')
    } elseif ($top -match 'EXCEL\.EXE') {
        [void]$findings.Add('EXCEL.EXE itself is faulting -> usually a damaged workbook or damaged Office install.')
    } elseif ($top) {
        [void]$findings.Add('Faulting module = ' + $top + ' -> that name identifies the owning component.')
    }
}

# --------------------------------------------------------- 3. Office alerts
Section '3. OFFICE ALERT DIALOGS (OAlerts log - the exact message the user saw)'
$al = @()
try {
    $al = @(Get-WinEvent -FilterHashtable @{LogName='OAlerts'; StartTime=$since} -MaxEvents 60 |
            Where-Object { $_.Message -match 'Excel' })
} catch { }
if ($al.Count -eq 0) {
    Say 'No Office alert entries (log empty or disabled).'
} else {
    foreach ($e in ($al | Select-Object -First 8)) {
        $txt = ($e.Message -replace '\s+', ' ')
        if ($txt.Length -gt 150) { $txt = $txt.Substring(0, 150) + '...' }
        Say ('  ' + $e.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss') + '  ' + $txt)
    }
}

# ------------------------------------------------------------------- 4. WER
Section '4. WINDOWS ERROR REPORTING (faulting module, independent of section 2)'
$werRoots = @(
  "$env:LOCALAPPDATA\Microsoft\Windows\WER\ReportArchive",
  "$env:LOCALAPPDATA\Microsoft\Windows\WER\ReportQueue"
)
$wers = @()
foreach ($r in $werRoots) {
    if (Test-Path $r) { $wers = $wers + @(Get-ChildItem $r -Recurse -Filter 'Report.wer') }
}
$wers = @($wers | Where-Object { $_.LastWriteTime -gt $since } | Sort-Object LastWriteTime -Descending)
$shown = 0
foreach ($w in $wers) {
    if ($shown -ge 5) { break }
    $txt = Get-Content $w.FullName -Raw
    if ($txt -notmatch 'EXCEL\.EXE') { continue }
    $mod = ''
    $sig = ''
    $m = [regex]::Match($txt, 'Sig\[3\]\.Value=(.+)')
    if ($m.Success) { $mod = $m.Groups[1].Value.Trim() }
    $m2 = [regex]::Match($txt, 'Sig\[6\]\.Value=(.+)')
    if ($m2.Success) { $sig = $m2.Groups[1].Value.Trim() }
    Say ('  ' + $w.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss') + '  faulting=' + $mod + '  exception=' + $sig)
    $shown = $shown + 1
}
if ($shown -eq 0) { Say 'No Excel WER reports in the last 21 days.' }

# ----------------------------------------------- 5. Excel resiliency
Section '5. EXCEL RESILIENCY (has Excel already blamed / disabled something?)'
$resFound = $false
foreach ($ver in @('16.0','15.0','14.0')) {
    $base = "HKCU:\Software\Microsoft\Office\$ver\Excel\Resiliency"
    if (-not (Test-Path $base)) { continue }
    $resFound = $true
    foreach ($sub in @('DisabledItems','CrashingAddinList','DocumentRecovery','StartupItems')) {
        $k = Join-Path $base $sub
        if (-not (Test-Path $k)) { continue }
        $props = Get-ItemProperty $k
        $names = @($props.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' })
        foreach ($n in $names) {
            $val = $n.Value
            $decoded = ''
            if ($val -is [byte[]]) {
                $s = [System.Text.Encoding]::Unicode.GetString($val)
                $decoded = (([regex]::Matches($s, '[\x20-\x7E]{4,}') | ForEach-Object { $_.Value }) -join ' | ')
            } else {
                $decoded = [string]$val
            }
            Hit ('  ' + $sub + ' -> ' + $decoded)
            if ($decoded -match 'AMI_Optix') {
                [void]$findings.Add('Excel has DISABLED/flagged AMI_Optix.xlam (Resiliency\' + $sub + ') - Excel blames the add-in.')
            }
        }
    }
}
if (-not $resFound) { Say 'No Resiliency keys found.' }

# ------------------------------------------------------------- 6. Add-ins
Section '6. ADD-INS LOADING AT STARTUP'
foreach ($ver in @('16.0','15.0','14.0')) {
    $opt = "HKCU:\Software\Microsoft\Office\$ver\Excel\Options"
    if (-not (Test-Path $opt)) { continue }
    $p = Get-ItemProperty $opt
    $p.PSObject.Properties | Where-Object { $_.Name -match '^OPEN' } | ForEach-Object {
        Say ('  ' + $ver + ' ' + $_.Name + ' = ' + $_.Value)
    }
}
Say ''
Say 'COM add-ins (LoadBehavior 3 = loads at startup):'
foreach ($hive in @('HKCU:\Software\Microsoft\Office\Excel\Addins','HKLM:\Software\Microsoft\Office\Excel\Addins','HKLM:\Software\WOW6432Node\Microsoft\Office\Excel\Addins')) {
    if (-not (Test-Path $hive)) { continue }
    Get-ChildItem $hive | ForEach-Object {
        $lp = Get-ItemProperty $_.PSPath
        Say ('  ' + $_.PSChildName + '  LoadBehavior=' + $lp.LoadBehavior + '  ' + $lp.FriendlyName)
    }
}

# ------------------------------------------------------- 7. AMI Optix add-in
Section '7. AMI OPTIX ADD-IN (which build is installed?)'
$local  = Join-Path $env:APPDATA 'Microsoft\AddIns\AMI_Optix.xlam'
$master = 'Z:\AMI_Optix.xlam'
foreach ($pair in @(@('LOCAL ', $local), @('MASTER', $master))) {
    $label = $pair[0]
    $path  = $pair[1]
    if (Test-Path $path) {
        $f = Get-Item $path
        Say ('  ' + $label + ' : ' + $path)
        Say ('           size=' + $f.Length + '  modified=' + $f.LastWriteTime)
    } else {
        Say ('  ' + $label + ' : NOT FOUND (' + $path + ')')
    }
}
if (Test-Path $local) {
    $zoneStream = $local + ':Zone.Identifier'
    if (Test-Path $zoneStream) { Hit '  NOTE: local .xlam is marked "from the internet" (Zone.Identifier present).' }
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $tmp = Join-Path $env:TEMP ('amiopt_probe_' + (Get-Random) + '.zip')
        Copy-Item $local $tmp -Force
        $zip = [System.IO.Compression.ZipFile]::OpenRead($tmp)
        $entry = $zip.Entries | Where-Object { $_.FullName -eq 'xl/vbaProject.bin' }
        if ($entry) {
            $ms = New-Object System.IO.MemoryStream
            $st = $entry.Open()
            $st.CopyTo($ms)
            $st.Close()
            $ascii = [System.Text.Encoding]::ASCII.GetString($ms.ToArray())
            $mods = [regex]::Matches($ascii, 'Module=([A-Za-z0-9_]+)') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique
            if ($mods) { Say ('  VBA modules: ' + ($mods -join ', ')) }
            if ($ascii -match 'AMI_Optix_Baseline') {
                Say '  Build       : NEW (has AMI_Optix_Baseline - 2026-09 snapshot fix)'
            } else {
                Say '  Build       : OLDER (no AMI_Optix_Baseline module)'
            }
        }
        $zip.Dispose()
        Remove-Item $tmp -Force
    } catch {
        Say '  (could not inspect VBA project - file may be locked by Excel)'
    }
}
foreach ($x in @("$env:APPDATA\Microsoft\Excel\XLSTART", "$env:ProgramFiles\Microsoft Office\root\Office16\XLSTART")) {
    if (-not (Test-Path $x)) { continue }
    $items = @(Get-ChildItem $x -File)
    foreach ($it in $items) { Hit ('  XLSTART auto-open file: ' + $it.FullName + '  (' + $it.LastWriteTime + ')') }
}

# ---------------------------------------------------------------- 8. Graphics
Section '8. GRAPHICS / RENDERING SETTINGS'
foreach ($ver in @('16.0','15.0')) {
    $g = "HKCU:\Software\Microsoft\Office\$ver\Common\Graphics"
    if (-not (Test-Path $g)) { continue }
    $gp = Get-ItemProperty $g
    Say ('  ' + $ver + ' DisableHardwareAcceleration = ' + $gp.DisableHardwareAcceleration)
    Say ('  ' + $ver + ' DisableAnimations           = ' + $gp.DisableAnimations)
}

# ---------------------------------------------------------------- 9. Verdict
Section '9. WHAT THIS POINTS AT'
if ($findings.Count -eq 0) {
    Say '  No single clear culprit from the automated checks.'
    Say '  Next: the faulting module name in section 2 is the lead.'
} else {
    foreach ($f in $findings) { Hit ('  * ' + $f) }
}
Say ''
Say '  Fast confirm test (60 seconds, fully reversible):'
Say '    1) Close Excel.'
Say '    2) Rename the add-in:'
Say '         ren "$env:APPDATA\Microsoft\AddIns\AMI_Optix.xlam" AMI_Optix.xlam.off'
Say '    3) Open the same workbook.'
Say '         No crash  = the add-in is the cause.'
Say '         Crash too = the workbook or Excel itself.'
Say '    4) Undo by renaming it back.'

$outFile = Join-Path $env:TEMP ('ExcelTriage_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.txt')
$report -join "`r`n" | Out-File -FilePath $outFile -Encoding utf8
Say ''
Say ('REPORT SAVED: ' + $outFile)
Say '============================= END OF REPORT ============================='
