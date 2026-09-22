<#
  Test-VbaCompile.ps1  (dev-PC tool; needs Excel installed, no Z: drive)

  Real VBA compile check for excel-addin/src: imports every .bas/.cls into a
  throwaway workbook via COM, stubs the Utilities form, runs the VBE's own
  Debug > Compile, and watches for the modal "Compile error" dialog - which it
  records (text + module + line), dismisses, and reports. Validated with
  -InjectError (a deliberately broken module must be reported).

    powershell -ExecutionPolicy Bypass -File tools\excel-agent\Test-VbaCompile.ps1
    powershell -ExecutionPolicy Bypass -File tools\excel-agent\Test-VbaCompile.ps1 -InjectError

  Prints VERDICT: COMPILE_OK or the error. Any Excel it started is killed.
#>
# Parent: runs the worker, watches for the VBE's modal "Compile error" dialog,
# records its text, dismisses it so the worker can continue, and prints a
# verdict. Excel is killed at the end no matter what.
param([switch]$InjectError)
$ErrorActionPreference = 'Continue'
$here   = Split-Path -Parent $MyInvocation.MyCommand.Path
$repo   = (Resolve-Path (Join-Path $here '..\..')).Path
$src    = Join-Path $repo 'excel-addin\src'
$work   = Join-Path $env:TEMP ('vbacheck_' + [guid]::NewGuid().ToString('N'))
$result = Join-Path $work 'result.txt'
$dialog = Join-Path $work 'dialog.txt'
New-Item -ItemType Directory -Force -Path $work | Out-Null

Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices; using System.Collections.Generic;
public static class W {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc p, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr h, EnumProc p, IntPtr l);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  public static string Text(IntPtr h){ var sb=new StringBuilder(2048); GetWindowText(h,sb,2048); return sb.ToString(); }
  public static string Cls(IntPtr h){ var sb=new StringBuilder(256); GetClassName(h,sb,256); return sb.ToString(); }
  public static List<IntPtr> Tops(){ var l=new List<IntPtr>(); EnumWindows((h,p)=>{ l.Add(h); return true; }, IntPtr.Zero); return l; }
  public static List<IntPtr> Kids(IntPtr h){ var l=new List<IntPtr>(); EnumChildWindows(h,(c,p)=>{ l.Add(c); return true; }, IntPtr.Zero); return l; }
}
"@

$excelPidsBefore = @(Get-Process excel -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$args = @('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $here 'Test-VbaCompile.worker.ps1'),'-Src',$src,'-Work',(Join-Path $work 'src'),'-ResultFile',$result,'-DialogFile',$dialog)
if ($InjectError) { $args += '-InjectError' }
$worker = Start-Process powershell -ArgumentList $args -PassThru -WindowStyle Hidden

$deadline = (Get-Date).AddSeconds(180)
$seen = @()
$seenWindows = @()
$debugWindows = @()
while ((Get-Date) -lt $deadline) {
    Start-Sleep -Milliseconds 400
    if ((Test-Path $result) -and ((Get-Content $result -ErrorAction SilentlyContinue) -match 'WORKER_DONE')) { break }
    foreach ($h in [W]::Tops()) {
        if (-not [W]::IsWindowVisible($h)) { continue }
        $wpid = [uint32]0; [W]::GetWindowThreadProcessId($h, [ref]$wpid) | Out-Null
        if ($excelPidsBefore -contains [int]$wpid) { continue }   # not ours
        $isExcel = $false
        try { $isExcel = ((Get-Process -Id ([int]$wpid) -ErrorAction SilentlyContinue).ProcessName -eq 'EXCEL') } catch {}
        if (-not $isExcel) { continue }
        $title = [W]::Text($h)
        $cls = [W]::Cls($h)
        $key = "$cls|$title"
        if (-not ($seenWindows -contains $key)) { $seenWindows += $key; $debugWindows += "WINDOW seen: class=$cls title=[$title]" }
        if ($cls -ne '#32770') { continue }
        if ($title -notmatch 'Microsoft Visual Basic|Microsoft Excel') { continue }
        $texts = @()
        foreach ($k in [W]::Kids($h)) { $t = [W]::Text($k); if ($t) { $texts += ($t -replace "`r`n", ' | ') } }
        $msg = ($texts -join ' || ')
        $seen += "DIALOG [$title]: $msg"
        Add-Content -Path $dialog -Value $msg -Encoding UTF8
        # Dismiss: click the first button (OK) so the worker's Execute() returns.
        $btn = [W]::Kids($h) | Where-Object { [W]::Cls($_) -eq 'Button' } | Select-Object -First 1
        if ($btn) { [W]::PostMessage($btn, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null }  # BM_CLICK
        else { [W]::PostMessage($h, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null }          # WM_CLOSE
        Start-Sleep -Milliseconds 300
    }
}

# Cleanup: kill any Excel we spawned.
Get-Process excel -ErrorAction SilentlyContinue | Where-Object { $excelPidsBefore -notcontains $_.Id } | Stop-Process -Force -ErrorAction SilentlyContinue
if (-not $worker.HasExited) { Stop-Process -Id $worker.Id -Force -ErrorAction SilentlyContinue }

Write-Host '===== VBA COMPILE CHECK ====='
if (Test-Path $result) { Get-Content $result | ForEach-Object { Write-Host $_ } } else { Write-Host 'NO RESULT FILE (worker never started?)' }
foreach ($d in $debugWindows) { Write-Host $d }
foreach ($s in $seen) { Write-Host $s }
if ($seen.Count -eq 0 -and (Test-Path $result) -and ((Get-Content $result) -match 'COMPILE_OK')) { Write-Host 'VERDICT: COMPILE_OK' } else { Write-Host 'VERDICT: SEE ABOVE' }
Remove-Item -Recurse -Force $work -ErrorAction SilentlyContinue
