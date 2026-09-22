# Worker: build a throwaway add-in project from excel-addin/src via Excel COM
# and run the VBE "Compile VBAProject" command. Writes a result file. A compile
# error pops a modal VBE dialog; the parent script dismisses it and records
# its text, then this worker records the module/line the VBE highlighted.
param(
    [string]$Src,
    [string]$Work,
    [string]$ResultFile,
    [string]$DialogFile,
    [switch]$InjectError
)
$ErrorActionPreference = 'Stop'
function Out-Result($s) { Add-Content -Path $ResultFile -Value $s -Encoding UTF8 }

# Programmatic VBA access (same registry flag the deploy script sets on client PCs).
Get-ChildItem 'HKCU:\Software\Microsoft\Office' -ErrorAction SilentlyContinue |
  Where-Object { $_.PSChildName -match '^\d+\.\d+$' } | ForEach-Object {
    $sec = "HKCU:\Software\Microsoft\Office\$($_.PSChildName)\Excel\Security"
    New-Item -Path $sec -Force | Out-Null
    Set-ItemProperty -Path $sec -Name AccessVBOM -Value 1 -Type DWord
  }

# CRLF-normalized copies (VBComponents.Import needs CRLF; repo files are LF).
New-Item -ItemType Directory -Force -Path $Work | Out-Null
$files = Get-ChildItem $Src -File | Where-Object { $_.Extension -in '.bas', '.cls' }
foreach ($f in $files) {
    $raw = [IO.File]::ReadAllText($f.FullName)
    $raw = $raw -replace "`r`n", "`n" -replace "`r", "`n" -replace "`n", "`r`n"
    [IO.File]::WriteAllText((Join-Path $Work $f.Name), $raw, (New-Object Text.UTF8Encoding($false)))
}

$xl = $null; $wb = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Add()
    $proj = $wb.VBProject

    foreach ($f in (Get-ChildItem $Work -File)) {
        $proj.VBComponents.Import($f.FullName) | Out-Null
    }
    # Stub the Utilities form (its .frx binary is not in the repo). Only one
    # member is referenced from code: frmUtilities.btnSave_Click.
    $frm = $proj.VBComponents.Add(3)   # vbext_ct_MSForm
    $frm.Name = 'frmUtilities'
    $frm.CodeModule.AddFromString("Public Sub btnSave_Click()`r`nEnd Sub")
    if ($InjectError) {
        # Self-test of the watchdog: an undeclared variable under Option Explicit.
        $bad = $proj.VBComponents.Add(1)
        $bad.Name = 'ZZ_InjectedError'
        $bad.CodeModule.AddFromString("Option Explicit`r`nPublic Sub Broken()`r`n    undeclaredVariable = 1`r`nEnd Sub")
    }

    Out-Result ("IMPORTED: " + ($proj.VBComponents | ForEach-Object { $_.Name }) -join ', ')

    # The VBE command bars only materialize once its window has been shown.
    try { $proj.Name = 'AMI_Optix_CompileCheck' } catch {}
    try { $xl.VBE.MainWindow.Visible = $true } catch {}
    Start-Sleep -Milliseconds 1500
    # Make OUR project the active one (Compile acts on the active project).
    try { $proj.VBComponents.Item('AMI_Optix_Main').Activate() } catch {}
    try { $proj.VBComponents.Item('AMI_Optix_Main').CodeModule.CodePane.Show() } catch {}
    try { $xl.VBE.ActiveVBProject = $proj } catch {}
    Start-Sleep -Milliseconds 500
    $activeName = ''
    try { $activeName = [string]$xl.VBE.ActiveVBProject.Name } catch { $activeName = '(none)' }
    $allProjects = @()
    try { foreach ($p in $xl.VBE.VBProjects) { $allProjects += [string]$p.Name } } catch {}
    Out-Result ('ACTIVE_PROJECT: ' + $activeName + '   ALL: ' + ($allProjects -join ', '))
    $ctl = $null
    try { $ctl = $xl.VBE.CommandBars.FindControl(1, 578) } catch {}   # Debug > Compile VBAProject
    if ($null -eq $ctl) {
        try {
            foreach ($bar in $xl.VBE.CommandBars) {
                foreach ($c in $bar.Controls) {
                    $cap = ''
                    try { $cap = [string]$c.Caption } catch {}
                    if ($cap -match '^Compi&?le') { $ctl = $c; break }
                }
                if ($ctl) { break }
            }
        } catch {}
    }
    if ($null -eq $ctl) {
        # Last resort: the Debug menu's own controls by caption.
        try {
            $dbg = $xl.VBE.CommandBars.Item('Debug')
            foreach ($c in $dbg.Controls) { if (([string]$c.Caption) -match 'Compile') { $ctl = $c; break } }
        } catch {}
    }
    if ($null -eq $ctl) { Out-Result 'COMPILE_CMD_MISSING'; throw 'compile command not found' }
    Out-Result ('COMPILE_CMD: ' + [string]$ctl.Caption + ' enabled=' + [string]$ctl.Enabled)
    Out-Result 'COMPILE_START'
    $ctl.Execute()
    # The VBE may process the command asynchronously: give any error dialog
    # time to appear (the parent records it in $DialogFile and dismisses it).
    $waitUntil = (Get-Date).AddSeconds(10)
    while ((Get-Date) -lt $waitUntil -and -not (Test-Path $DialogFile)) { Start-Sleep -Milliseconds 250 }
    if (Test-Path $DialogFile) {
        $pane = $null
        try { $pane = $xl.VBE.ActiveCodePane } catch {}
        if ($pane) {
            $sl = 0; $sc = 0; $el = 0; $ec = 0
            try { $pane.GetSelection([ref]$sl, [ref]$sc, [ref]$el, [ref]$ec) } catch {}
            $modName = ''
            $lineText = ''
            try { $modName = $pane.CodeModule.Parent.Name; if ($sl -gt 0) { $lineText = $pane.CodeModule.Lines($sl, 1) } } catch {}
            Out-Result ("COMPILE_ERROR_AT: " + $modName + " line " + $sl + ": " + $lineText.Trim())
        } else {
            Out-Result 'COMPILE_ERROR_AT: (no active code pane)'
        }
    } else {
        Out-Result 'COMPILE_OK'
    }
}
catch {
    Out-Result ("WORKER_EXCEPTION: " + $_.Exception.Message)
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($xl) { try { $xl.Quit() } catch {} }
    Out-Result 'WORKER_DONE'
}
