param(
    [string]$AgentRoot = 'C:\AMI_Optix_Agent',
    [string]$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
)

$modulePath = Join-Path $PSScriptRoot 'AmiOptix.Agent.psm1'
Import-Module $modulePath -Force

$folders = @(
    $AgentRoot,
    (Join-Path $AgentRoot 'artifacts'),
    (Join-Path $AgentRoot 'build'),
    (Join-Path $AgentRoot 'config'),
    (Join-Path $AgentRoot 'source'),
    (Join-Path $AgentRoot 'source\excel-addin\src'),
    (Join-Path $AgentRoot 'source\excel-addin\forms'),
    (Join-Path $AgentRoot 'source\excel-addin\customUI'),
    (Join-Path $AgentRoot 'staging'),
    (Join-Path $AgentRoot 'state'),
    (Join-Path $AgentRoot 'workbooks'),
    (Join-Path $AgentRoot 'workbooks\runtime')
)

foreach ($folder in $folders) {
    New-Item -ItemType Directory -Force -Path $folder | Out-Null
}

Copy-Item -Path (Join-Path $RepoRoot 'excel-addin\src\*') -Destination (Join-Path $AgentRoot 'source\excel-addin\src') -Recurse -Force
Copy-Item -Path (Join-Path $RepoRoot 'excel-addin\forms\*') -Destination (Join-Path $AgentRoot 'source\excel-addin\forms') -Recurse -Force
Copy-Item -Path (Join-Path $RepoRoot 'excel-addin\customUI\*') -Destination (Join-Path $AgentRoot 'source\excel-addin\customUI') -Recurse -Force

$manifestPath = Join-Path $AgentRoot 'state\workspace-manifest.json'
$configPath = Join-Path $AgentRoot 'config\agent-config.json'
$acceptancePath = Join-Path $AgentRoot 'config\acceptance.json'

$manifest = New-AmiOptixWorkspaceManifest `
    -SourceRoot (Join-Path $AgentRoot 'source') `
    -RepoRoot $RepoRoot `
    -AgentRoot $AgentRoot `
    -ManifestPath $manifestPath

$configTemplatePath = Join-Path $PSScriptRoot 'config\agent-config.template.json'
$acceptanceTemplatePath = Join-Path $PSScriptRoot 'config\acceptance.template.json'

if (-not (Test-Path -LiteralPath $configPath)) {
    Copy-Item -LiteralPath $configTemplatePath -Destination $configPath -Force
}

if (-not (Test-Path -LiteralPath $acceptancePath)) {
    Copy-Item -LiteralPath $acceptanceTemplatePath -Destination $acceptancePath -Force
}

$preflightResult = Test-AmiOptixEnvironment -Config (Read-AmiOptixJsonFile -Path $configPath) -Manifest $manifest -ResultPath (Join-Path $AgentRoot 'state\preflight.json')

Write-Host "AMI Optix agent workspace initialized at $AgentRoot"
Write-Host "Manifest: $manifestPath"
Write-Host "Config:   $configPath"
Write-Host "Suite:    $acceptancePath"
Write-Host "Preflight passed: $($preflightResult.succeeded)"
if ($preflightResult.manualActionRequests.Count -gt 0) {
    Write-Host ''
    Write-Host 'Manual actions required before unattended build/test:'
    foreach ($request in $preflightResult.manualActionRequests) {
        Write-Host "- $($request.blocked): $($request.why)"
    }
}
