param(
    [string]$AgentRoot = 'C:\AMI_Optix_Agent'
)

$modulePath = Join-Path $PSScriptRoot 'AmiOptix.Agent.psm1'
Import-Module $modulePath -Force

$configPath = Join-Path $AgentRoot 'config\agent-config.json'
$manifestPath = Join-Path $AgentRoot 'state\workspace-manifest.json'
$resultPath = Join-Path $AgentRoot 'artifacts\acceptance-result.json'

$config = Read-AmiOptixJsonFile -Path $configPath
$manifest = Read-AmiOptixJsonFile -Path $manifestPath
$result = Invoke-AmiOptixAcceptanceSuite -Config $config -Manifest $manifest -ResultPath $resultPath

if ($config.buildTag) {
    Write-Host "Build tag: $($config.buildTag)"
}
Write-Host "Acceptance passed: $($result.succeeded)"
Write-Host "Result: $resultPath"
foreach ($scenario in $result.scenarios) {
    Write-Host "- $($scenario.name): $($scenario.classification) :: $($scenario.details)"
}
if ($result.manualActionRequests.Count -gt 0) {
    Write-Host ''
    Write-Host 'Manual actions required:'
    foreach ($request in $result.manualActionRequests) {
        Write-Host "- $($request.blocked): $($request.why)"
    }
}
