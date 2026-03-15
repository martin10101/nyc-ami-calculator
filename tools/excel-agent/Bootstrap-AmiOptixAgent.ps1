param(
    [string]$AgentRoot = 'C:\AMI_Optix_Agent',
    [string]$RepoRoot = ''
)

function Resolve-AmiOptixRepoRoot {
    param(
        [Parameter(Mandatory = $true)][string]$AgentRoot,
        [string]$RequestedRepoRoot
    )

    $candidates = New-Object System.Collections.Generic.List[string]
    if (-not [string]::IsNullOrWhiteSpace($RequestedRepoRoot)) {
        $candidates.Add($RequestedRepoRoot)
    }
    $candidates.Add((Join-Path $PSScriptRoot '..\..'))
    $candidates.Add((Join-Path $AgentRoot 'repo-snapshot'))

    foreach ($candidate in $candidates) {
        try {
            $resolvedCandidate = (Resolve-Path -LiteralPath $candidate -ErrorAction Stop).Path
        } catch {
            continue
        }

        $agentModulePath = Join-Path $resolvedCandidate 'tools\excel-agent\AmiOptix.Agent.psm1'
        $sourcePath = Join-Path $resolvedCandidate 'excel-addin\src'
        if ((Test-Path -LiteralPath $agentModulePath) -and (Test-Path -LiteralPath $sourcePath)) {
            return $resolvedCandidate
        }
    }

    throw "Could not resolve a valid repo root. Expected tools\excel-agent and excel-addin\src under the requested path, '$PSScriptRoot\..\..', or '$AgentRoot\repo-snapshot'."
}

function Get-AmiOptixBuildTag {
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot
    )

    $commit = ''
    try {
        $commit = (& git -C $RepoRoot rev-parse --short HEAD 2>$null)
    } catch {
        $commit = ''
    }

    if (-not [string]::IsNullOrWhiteSpace($commit) -and $LASTEXITCODE -eq 0) {
        return [string]$commit.Trim()
    }

    return "snapshot-{0}" -f [DateTime]::UtcNow.ToString('yyyyMMddHHmmss')
}

$resolvedRepoRoot = Resolve-AmiOptixRepoRoot -AgentRoot $AgentRoot -RequestedRepoRoot $RepoRoot
$buildTag = Get-AmiOptixBuildTag -RepoRoot $resolvedRepoRoot

$modulePath = Join-Path $PSScriptRoot 'AmiOptix.Agent.psm1'
Import-Module $modulePath -Force

$folders = @(
    $AgentRoot,
    (Join-Path $AgentRoot 'artifacts'),
    (Join-Path $AgentRoot 'build'),
    (Join-Path $AgentRoot 'config'),
    (Join-Path $AgentRoot 'scripts'),
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

$scriptsRoot = Join-Path $AgentRoot 'scripts'
$sourceRoot = Join-Path $AgentRoot 'source'
$workbooksRoot = Join-Path $AgentRoot 'workbooks'
$rentRollYearsRoot = Join-Path $env:APPDATA 'AMI_Optix\RentRollYears'
$bundledRentRollYearsRoot = Join-Path $resolvedRepoRoot 'tools\excel-agent\assets\rent-roll-years'

Copy-Item -Path (Join-Path $resolvedRepoRoot 'tools\excel-agent\*') -Destination $scriptsRoot -Recurse -Force
Copy-Item -Path (Join-Path $resolvedRepoRoot 'excel-addin\src\*') -Destination (Join-Path $sourceRoot 'excel-addin\src') -Recurse -Force
Copy-Item -Path (Join-Path $resolvedRepoRoot 'excel-addin\forms\*') -Destination (Join-Path $sourceRoot 'excel-addin\forms') -Recurse -Force
Copy-Item -Path (Join-Path $resolvedRepoRoot 'excel-addin\customUI\*') -Destination (Join-Path $sourceRoot 'excel-addin\customUI') -Recurse -Force

$bundledUapSource = Join-Path $resolvedRepoRoot 'tools\excel-agent\assets\workbooks\UAP_golden.xlsm'
$bundledMihSource = Join-Path $resolvedRepoRoot 'tools\excel-agent\assets\workbooks\MIH_golden.xlsb'
if (Test-Path -LiteralPath $bundledUapSource) {
    Copy-Item -LiteralPath $bundledUapSource -Destination (Join-Path $workbooksRoot 'UAP_golden.xlsm') -Force
}
if (Test-Path -LiteralPath $bundledMihSource) {
    Copy-Item -LiteralPath $bundledMihSource -Destination (Join-Path $workbooksRoot 'MIH_golden.xlsb') -Force
}

$seededRentRollYears = New-Object System.Collections.Generic.List[object]
if (Test-Path -LiteralPath $bundledRentRollYearsRoot) {
    foreach ($yearFolder in Get-ChildItem -LiteralPath $bundledRentRollYearsRoot -Directory | Sort-Object Name) {
        $yearLabel = [string]$yearFolder.Name
        $sourceWorkbookPath = Join-Path $yearFolder.FullName ("RentCalculator_{0}.xlsx" -f $yearLabel)
        if (-not (Test-Path -LiteralPath $sourceWorkbookPath)) {
            continue
        }

        $destinationFolder = Join-Path $rentRollYearsRoot $yearLabel
        $destinationWorkbookPath = Join-Path $destinationFolder ("RentCalculator_{0}.xlsx" -f $yearLabel)
        New-Item -ItemType Directory -Force -Path $destinationFolder | Out-Null
        Copy-Item -LiteralPath $sourceWorkbookPath -Destination $destinationWorkbookPath -Force

        $destinationWorkbook = Get-Item -LiteralPath $destinationWorkbookPath
        $seededRentRollYears.Add([ordered]@{
            year = $yearLabel
            path = $destinationWorkbookPath
            size = [int64]$destinationWorkbook.Length
        })
    }
}

$manifestPath = Join-Path $AgentRoot 'state\workspace-manifest.json'
$configPath = Join-Path $AgentRoot 'config\agent-config.json'
$acceptancePath = Join-Path $AgentRoot 'config\acceptance.json'

$manifest = New-AmiOptixWorkspaceManifest `
    -SourceRoot $sourceRoot `
    -RepoRoot $resolvedRepoRoot `
    -AgentRoot $AgentRoot `
    -ManifestPath $manifestPath

$configTemplatePath = Join-Path $resolvedRepoRoot 'tools\excel-agent\config\agent-config.template.json'
$acceptanceTemplatePath = Join-Path $resolvedRepoRoot 'tools\excel-agent\config\acceptance.template.json'

Copy-Item -LiteralPath $configTemplatePath -Destination $configPath -Force
Copy-Item -LiteralPath $acceptanceTemplatePath -Destination $acceptancePath -Force

$config = Read-AmiOptixJsonFile -Path $configPath
$config.agentRoot = $AgentRoot
$config.stagedContainerPath = (Join-Path $AgentRoot 'staging\AMI_Optix_Staged.xlam')
$config.rebuiltAddinPath = (Join-Path $AgentRoot 'build\AMI_Optix_Autofix.xlam')
$config.goldenWorkbooks.uap = (Join-Path $workbooksRoot 'UAP_golden.xlsm')
$config.goldenWorkbooks.mih = (Join-Path $workbooksRoot 'MIH_golden.xlsb')
$config.acceptancePath = $acceptancePath
if ($null -ne $config.PSObject.Properties['buildTag']) {
    $config.buildTag = $buildTag
} else {
    $config | Add-Member -NotePropertyName buildTag -NotePropertyValue $buildTag
}
Write-AmiOptixJsonFile -Path $configPath -InputObject $config

$syncStatePath = Join-Path $AgentRoot 'state\agent-sync.json'
$syncState = [pscustomobject]@{
    generatedAtUtc = [DateTime]::UtcNow.ToString('o')
    buildTag = $buildTag
    repoRoot = $resolvedRepoRoot
    agentRoot = $AgentRoot
    scriptsRoot = $scriptsRoot
    sourceRoot = $sourceRoot
    seededRentRollYears = @($seededRentRollYears.ToArray())
}
Write-AmiOptixJsonFile -Path $syncStatePath -InputObject $syncState

$preflightResult = Test-AmiOptixEnvironment -Config $config -Manifest $manifest -ResultPath (Join-Path $AgentRoot 'state\preflight.json')

Write-Host "AMI Optix agent workspace initialized at $AgentRoot"
Write-Host "Build tag: $buildTag"
Write-Host "Manifest: $manifestPath"
Write-Host "Config:   $configPath"
Write-Host "Suite:    $acceptancePath"
Write-Host "Sync:     $syncStatePath"
Write-Host "Preflight passed: $($preflightResult.succeeded)"
if ($seededRentRollYears.Count -gt 0) {
    Write-Host "Seeded rent-roll years: $((@($seededRentRollYears | ForEach-Object { $_.year }) -join ', '))"
}
if ($preflightResult.manualActionRequests.Count -gt 0) {
    Write-Host ''
    Write-Host 'Manual actions required before unattended build/test:'
    foreach ($request in $preflightResult.manualActionRequests) {
        Write-Host "- $($request.blocked): $($request.why)"
    }
}
