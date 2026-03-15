param(
    [string]$AgentRoot = 'C:\AMI_Optix_Agent',
    [string]$Branch = 'feature/excel-agent-foundation'
)

$ErrorActionPreference = 'Stop'

function Get-AmiOptixTempPath {
    param(
        [Parameter(Mandatory = $true)][string]$Prefix,
        [string]$Extension = ''
    )

    $name = '{0}-{1:yyyyMMddHHmmssfff}{2}' -f $Prefix, [DateTime]::UtcNow, $Extension
    return Join-Path $env:TEMP $name
}

function Expand-ZipArchiveSafe {
    param(
        [Parameter(Mandatory = $true)][string]$ZipPath,
        [Parameter(Mandatory = $true)][string]$Destination
    )

    try {
        Expand-Archive -LiteralPath $ZipPath -DestinationPath $Destination -Force
        return
    } catch {
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $Destination)
}

$zipUrl = "https://codeload.github.com/martin10101/nyc-ami-calculator/zip/refs/heads/$Branch"
$zipPath = Get-AmiOptixTempPath -Prefix 'ami-optix-agent' -Extension '.zip'
$extractRoot = Get-AmiOptixTempPath -Prefix 'ami-optix-agent-src'

New-Item -ItemType Directory -Force -Path $extractRoot | Out-Null

Write-Host "Downloading $zipUrl"
Invoke-WebRequest -Uri $zipUrl -UseBasicParsing -OutFile $zipPath

Write-Host "Extracting to $extractRoot"
Expand-ZipArchiveSafe -ZipPath $zipPath -Destination $extractRoot

$repoRoot = (Get-ChildItem -LiteralPath $extractRoot -Directory | Select-Object -First 1).FullName
if (-not $repoRoot) {
    throw "Failed to locate extracted repo root under $extractRoot."
}

$refreshScript = Join-Path $repoRoot 'tools\excel-agent\Refresh-AmiOptixAgent.ps1'
if (-not (Test-Path -LiteralPath $refreshScript)) {
    throw "Refresh script not found in extracted zip: $refreshScript"
}

Write-Host ''
Write-Host "Running refresh from snapshot: $repoRoot"
& powershell -NoProfile -ExecutionPolicy Bypass -File $refreshScript -AgentRoot $AgentRoot -RepoRoot $repoRoot | Out-Host
