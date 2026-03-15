param(
    [string]$AgentRoot = 'C:\AMI_Optix_Agent'
)

$ErrorActionPreference = 'Stop'

$modulePath = Join-Path $AgentRoot 'scripts\AmiOptix.Agent.psm1'
if (-not (Test-Path $modulePath)) {
    Write-Output 'SETUP_NEEDED: Workspace not bootstrapped. Run a full refresh first.'
    exit 1
}

Import-Module $modulePath -Force

$configPath   = Join-Path $AgentRoot 'config\agent-config.json'
$manifestPath = Join-Path $AgentRoot 'state\workspace-manifest.json'
$config       = Read-AmiOptixJsonFile -Path $configPath
$manifest     = Read-AmiOptixJsonFile -Path $manifestPath

# ── Step 1: Build (import source .bas files into the .xlam, no tests) ─────────
Write-Output 'Building add-in from source files...'
$build = Invoke-AmiOptixStagedBuild `
    -Config $config `
    -Manifest $manifest `
    -ResultPath (Join-Path $AgentRoot 'artifacts\build-result.json')

if (-not $build.succeeded) {
    Write-Output "BUILD_FAILED"
    Write-Output ($build | ConvertTo-Json -Compress)
    exit 1
}

Write-Output 'Build succeeded. Opening add-in to check compile...'

# ── Step 2: Open the add-in and call the no-op SyntaxCheck_Agent sub ──────────
$addinPath = [string]$config.rebuiltAddinPath
$excel     = $null
$wb        = $null

try {
    $excel = New-AmiOptixExcelApplication
    $wb    = $excel.Workbooks.Open($addinPath)

    $addinName = [string]$wb.Name
    $macroRef  = if ($addinName -match '\s') {
        "'{0}'!AMI_Optix_Diagnostics.SyntaxCheck_Agent" -f $addinName
    } else {
        '{0}!AMI_Optix_Diagnostics.SyntaxCheck_Agent' -f $addinName
    }

    $null = Invoke-AmiOptixExcelMacro -Excel $excel -MacroName $macroRef -Arguments @()
    Write-Output 'COMPILE_OK'
} catch {
    Write-Output "COMPILE_ERROR: $($_.Exception.Message)"
} finally {
    Close-AmiOptixExcelApplication -Workbook $wb -Excel $excel
}
