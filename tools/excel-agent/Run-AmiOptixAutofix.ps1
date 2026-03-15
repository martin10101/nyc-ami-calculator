param(
    [string]$AgentRoot = 'C:\AMI_Optix_Agent',
    [string[]]$RequestedFixFiles = @()
)

$modulePath = Join-Path $PSScriptRoot 'AmiOptix.Agent.psm1'
Import-Module $modulePath -Force

$configPath = Join-Path $AgentRoot 'config\agent-config.json'
$manifestPath = Join-Path $AgentRoot 'state\workspace-manifest.json'
$stateRoot = Join-Path $AgentRoot 'state'

$config = Read-AmiOptixJsonFile -Path $configPath
$manifest = Read-AmiOptixJsonFile -Path $manifestPath

$dependencyReport = Get-AmiOptixDependencyClassification -Manifest $manifest -RequestedRelativePaths $RequestedFixFiles
Write-AmiOptixJsonFile -Path (Join-Path $stateRoot 'dependency-report.json') -InputObject $dependencyReport

if ($dependencyReport.manualPackageManaged.Count -gt 0) {
    Write-Host 'Blocked by manual/package-managed files:'
    foreach ($entry in $dependencyReport.manualPackageManaged) {
        Write-Host "- $($entry.relativePath): $($entry.reason)"
    }
    exit 2
}

if ($dependencyReport.securitySensitive.Count -gt 0) {
    Write-Host 'Blocked by security-sensitive files:'
    foreach ($entry in $dependencyReport.securitySensitive) {
        Write-Host "- $($entry.relativePath): $($entry.reason)"
    }
    Write-Host 'Explain the exact security change and obtain explicit approval before proceeding.'
    exit 3
}

$preflight = Test-AmiOptixEnvironment -Config $config -Manifest $manifest -ResultPath (Join-Path $stateRoot 'preflight.json')
if (-not $preflight.succeeded) {
    Write-Host 'Preflight failed. See state\preflight.json.'
    exit 4
}

$build = Invoke-AmiOptixStagedBuild -Config $config -Manifest $manifest -ResultPath (Join-Path $AgentRoot 'artifacts\build-result.json')
if (-not $build.succeeded) {
    Write-Host 'Build failed. See artifacts\build-result.json.'
    exit 5
}

$acceptance = Invoke-AmiOptixAcceptanceSuite -Config $config -Manifest $manifest -ResultPath (Join-Path $AgentRoot 'artifacts\acceptance-result.json')
if (-not $acceptance.succeeded) {
    Write-Host 'Acceptance failed or is blocked. See artifacts\acceptance-result.json.'
    exit 6
}

Write-Host 'Preflight, staged build, and acceptance all passed.'
