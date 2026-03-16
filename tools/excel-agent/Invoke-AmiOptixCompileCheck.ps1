param(
    [string]$AgentRoot = 'C:\AMI_Optix_Agent'
)

$ErrorActionPreference = 'Stop'

# Kill any lingering Excel processes so the build can overwrite the add-in
Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 800

# Prefer the repo psm1 (always latest); fall back to deployed copy
$repoModulePath = if ($env:REPO_ROOT) { Join-Path $env:REPO_ROOT 'tools\excel-agent\AmiOptix.Agent.psm1' } else { $null }
$deployedModulePath = Join-Path $AgentRoot 'scripts\AmiOptix.Agent.psm1'

$modulePath = if ($repoModulePath -and (Test-Path $repoModulePath)) { $repoModulePath } else { $deployedModulePath }

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
    $errorMsg = if ($build.error) { $build.error } elseif ($build.compile -and $build.compile.details) { $build.compile.details } else { ($build | ConvertTo-Json -Compress) }
    Write-Output "COMPILE_ERROR: $errorMsg"
    exit 1
}

# The staged build already ran VBA compile (Invoke-AmiOptixVbaCompile) and
# saved the add-in — if it succeeded the code is clean.
Write-Output 'COMPILE_OK'
