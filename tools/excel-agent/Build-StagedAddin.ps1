param(
    [string]$AgentRoot = 'C:\AMI_Optix_Agent'
)

$modulePath = Join-Path $PSScriptRoot 'AmiOptix.Agent.psm1'
Import-Module $modulePath -Force

$configPath = Join-Path $AgentRoot 'config\agent-config.json'
$manifestPath = Join-Path $AgentRoot 'state\workspace-manifest.json'
$resultPath = Join-Path $AgentRoot 'artifacts\build-result.json'

$config = Read-AmiOptixJsonFile -Path $configPath
$manifest = Read-AmiOptixJsonFile -Path $manifestPath
$result = Invoke-AmiOptixStagedBuild -Config $config -Manifest $manifest -ResultPath $resultPath

if ($config.buildTag) {
    Write-Host "Build tag: $($config.buildTag)"
}
Write-Host "Build succeeded: $($result.succeeded)"
Write-Host "Result: $resultPath"
if ($result.error) {
    Write-Host "Error: $($result.error)"
}
if ($result.compileWarning) {
    Write-Host "Warning: $($result.compileWarning)"
}
if ($result.manualActionRequests.Count -gt 0) {
    foreach ($request in $result.manualActionRequests) {
        Write-Host "- $($request.blocked): $($request.why)"
    }
}
