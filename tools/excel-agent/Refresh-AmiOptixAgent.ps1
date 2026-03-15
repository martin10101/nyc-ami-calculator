param(
    [string]$AgentRoot = 'C:\AMI_Optix_Agent',
    [string]$RepoRoot = ''
)

$ErrorActionPreference = 'Stop'

function Invoke-AmiOptixScript {
    param(
        [Parameter(Mandatory = $true)][string]$ScriptPath,
        [Parameter(Mandatory = $true)][string]$AgentRoot,
        [string]$RepoRoot = ''
    )

    $arguments = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', $ScriptPath,
        '-AgentRoot', $AgentRoot
    )

    if (-not [string]::IsNullOrWhiteSpace($RepoRoot)) {
        $arguments += @('-RepoRoot', $RepoRoot)
    }

    & powershell @arguments | Out-Host
}

$bootstrapScript = Join-Path $PSScriptRoot 'Bootstrap-AmiOptixAgent.ps1'
Invoke-AmiOptixScript -ScriptPath $bootstrapScript -AgentRoot $AgentRoot -RepoRoot $RepoRoot

$preflightPath = Join-Path $AgentRoot 'state\preflight.json'
if (Test-Path -LiteralPath $preflightPath) {
    $preflight = Get-Content -LiteralPath $preflightPath -Raw | ConvertFrom-Json
    if (-not $preflight.succeeded) {
        throw "Preflight failed. Review $preflightPath before rerunning the refresh."
    }
}

$buildScript = Join-Path $AgentRoot 'scripts\Build-StagedAddin.ps1'
Invoke-AmiOptixScript -ScriptPath $buildScript -AgentRoot $AgentRoot

$buildResultPath = Join-Path $AgentRoot 'artifacts\build-result.json'
if (Test-Path -LiteralPath $buildResultPath) {
    $buildResult = Get-Content -LiteralPath $buildResultPath -Raw | ConvertFrom-Json
    if (-not $buildResult.succeeded) {
        throw "Build failed. Review $buildResultPath before rerunning the refresh."
    }
}

$acceptanceScript = Join-Path $AgentRoot 'scripts\Invoke-AmiOptixAcceptance.ps1'
Invoke-AmiOptixScript -ScriptPath $acceptanceScript -AgentRoot $AgentRoot

$acceptanceResultPath = Join-Path $AgentRoot 'artifacts\acceptance-result.json'
if (Test-Path -LiteralPath $acceptanceResultPath) {
    $acceptanceResult = Get-Content -LiteralPath $acceptanceResultPath -Raw | ConvertFrom-Json
    if (-not $acceptanceResult.succeeded) {
        throw "Acceptance failed. Review $acceptanceResultPath before rerunning the refresh."
    }
}
