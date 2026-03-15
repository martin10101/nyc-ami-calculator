param(
    [string]$AgentRoot = 'C:\AMI_Optix_Agent',
    [string]$Branch = 'feature/excel-agent-foundation'
)

$destinationRoot = Join-Path $AgentRoot 'scripts'
New-Item -ItemType Directory -Force -Path $destinationRoot | Out-Null

$files = @(
    'AmiOptix.Agent.psm1',
    'Bootstrap-AmiOptixAgent.ps1',
    'Build-StagedAddin.ps1',
    'Install-AmiOptixAgent.ps1',
    'Invoke-AmiOptixAcceptance.ps1',
    'Run-AmiOptixAutofix.ps1'
)

foreach ($file in $files) {
    $url = "https://raw.githubusercontent.com/martin10101/nyc-ami-calculator/$Branch/tools/excel-agent/$file"
    $destination = Join-Path $destinationRoot $file
    Invoke-WebRequest -Uri $url -UseBasicParsing -OutFile $destination
    Write-Host "Updated $destination"
}

Write-Host ''
Write-Host 'Agent scripts refreshed from GitHub.'
