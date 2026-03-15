param(
    [string]$AgentRoot = 'C:\AMI_Optix_Agent'
)

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$localScriptsRoot = Join-Path $AgentRoot 'scripts'
$localConfigRoot = Join-Path $AgentRoot 'config'
$localSourceRoot = Join-Path $AgentRoot 'source'
$localRepoSnapshotRoot = Join-Path $AgentRoot 'repo-snapshot'
$stagingRoot = Join-Path $AgentRoot 'staging'
$workbooksRoot = Join-Path $AgentRoot 'workbooks'

foreach ($folder in @(
    $AgentRoot,
    $localScriptsRoot,
    $localConfigRoot,
    $localSourceRoot,
    $localRepoSnapshotRoot,
    $stagingRoot,
    $workbooksRoot,
    (Join-Path $localSourceRoot 'excel-addin\src'),
    (Join-Path $localSourceRoot 'excel-addin\forms'),
    (Join-Path $localSourceRoot 'excel-addin\customUI')
)) {
    New-Item -ItemType Directory -Force -Path $folder | Out-Null
}

Copy-Item -Path (Join-Path $repoRoot 'tools\excel-agent\*') -Destination $localScriptsRoot -Recurse -Force
Copy-Item -Path (Join-Path $repoRoot 'excel-addin\src\*') -Destination (Join-Path $localSourceRoot 'excel-addin\src') -Recurse -Force
Copy-Item -Path (Join-Path $repoRoot 'excel-addin\forms\*') -Destination (Join-Path $localSourceRoot 'excel-addin\forms') -Recurse -Force
Copy-Item -Path (Join-Path $repoRoot 'excel-addin\customUI\*') -Destination (Join-Path $localSourceRoot 'excel-addin\customUI') -Recurse -Force

$installedAddinSource = Join-Path $env:APPDATA 'Microsoft\AddIns\AMI_Optix.xlam'
$stagedAddinTarget = Join-Path $stagingRoot 'AMI_Optix_Staged.xlam'
$bundledUapSource = Join-Path $repoRoot 'tools\excel-agent\assets\workbooks\UAP_golden.xlsm'
$bundledMihSource = Join-Path $repoRoot 'tools\excel-agent\assets\workbooks\MIH_golden.xlsb'
$uapTarget = Join-Path $workbooksRoot 'UAP_golden.xlsm'
$mihTarget = Join-Path $workbooksRoot 'MIH_golden.xlsb'

$copiedInstalledAddin = $false
$copiedBundledUap = $false
$copiedBundledMih = $false

if (Test-Path -LiteralPath $installedAddinSource) {
    Copy-Item -LiteralPath $installedAddinSource -Destination $stagedAddinTarget -Force
    $copiedInstalledAddin = $true
}

if (Test-Path -LiteralPath $bundledUapSource) {
    Copy-Item -LiteralPath $bundledUapSource -Destination $uapTarget -Force
    $copiedBundledUap = $true
}

if (Test-Path -LiteralPath $bundledMihSource) {
    Copy-Item -LiteralPath $bundledMihSource -Destination $mihTarget -Force
    $copiedBundledMih = $true
}

$repoSnapshotItems = @(
    'CODEX.md',
    'docs',
    'excel-addin',
    'tools\excel-agent'
)

foreach ($item in $repoSnapshotItems) {
    $sourcePath = Join-Path $repoRoot $item
    if (Test-Path -LiteralPath $sourcePath) {
        $destinationPath = Join-Path $localRepoSnapshotRoot $item
        $destinationParent = Split-Path -Parent $destinationPath
        if ($destinationParent) {
            New-Item -ItemType Directory -Force -Path $destinationParent | Out-Null
        }
        Copy-Item -Path $sourcePath -Destination $destinationPath -Recurse -Force
    }
}

$bootstrapScript = Join-Path $localScriptsRoot 'Bootstrap-AmiOptixAgent.ps1'
& powershell -NoProfile -ExecutionPolicy Bypass -File $bootstrapScript -AgentRoot $AgentRoot -RepoRoot $repoRoot | Out-Host

$configPath = Join-Path $AgentRoot 'config\agent-config.json'
if (Test-Path -LiteralPath $configPath) {
    $configJson = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json
    $configJson.agentRoot = $AgentRoot
    $configJson.stagedContainerPath = $stagedAddinTarget
    $configJson.rebuiltAddinPath = (Join-Path $AgentRoot 'build\AMI_Optix_Autofix.xlam')
    $configJson.goldenWorkbooks.uap = $uapTarget
    $configJson.goldenWorkbooks.mih = $mihTarget
    $configJson.acceptancePath = (Join-Path $AgentRoot 'config\acceptance.json')
    $configJson | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $configPath -Encoding UTF8
}

$modulePath = Join-Path $localScriptsRoot 'AmiOptix.Agent.psm1'
Import-Module $modulePath -Force
$updatedConfig = Read-AmiOptixJsonFile -Path $configPath
$manifestPath = Join-Path $AgentRoot 'state\workspace-manifest.json'
$manifest = Read-AmiOptixJsonFile -Path $manifestPath
$updatedPreflight = Test-AmiOptixEnvironment -Config $updatedConfig -Manifest $manifest -ResultPath (Join-Path $AgentRoot 'state\preflight.json')

$nextStepsPath = Join-Path $AgentRoot 'NEXT-STEPS.txt'
$nextSteps = @"
AMI Optix agent install finished.

What was created:
- $AgentRoot
- $localScriptsRoot
- $localSourceRoot
- $localRepoSnapshotRoot

What the installer already copied for you:
- Installed add-in found: $copiedInstalledAddin
- Bundled UAP workbook copied: $copiedBundledUap
- Bundled MIH workbook copied: $copiedBundledMih
- Updated preflight passed: $($updatedPreflight.succeeded)

What you still need to do manually:
1. Open Excel once and enable:
   File > Options > Trust Center > Trust Center Settings > Macro Settings >
   Trust access to the VBA project object model

2. Open your normal AMI Optix add-in once in Excel and save the API key through:
   AMI Optix > API Settings

3. Build the staged test add-in:
   powershell -ExecutionPolicy Bypass -File $AgentRoot\scripts\Build-StagedAddin.ps1 -AgentRoot $AgentRoot

4. Run the guarded acceptance check:
   powershell -ExecutionPolicy Bypass -File $AgentRoot\scripts\Invoke-AmiOptixAcceptance.ps1 -AgentRoot $AgentRoot

5. Run the orchestrator:
   powershell -ExecutionPolicy Bypass -File $AgentRoot\scripts\Run-AmiOptixAutofix.ps1 -AgentRoot $AgentRoot

Important:
- This setup does NOT need Git on the client PC.
- The zip/extracted repo is only the installer source.
- After install, the agent lives under $AgentRoot.
"@

if (-not $copiedInstalledAddin) {
    $nextSteps += @"

Still missing:
- Installed add-in was NOT found at:
  $installedAddinSource
- Put the working add-in here manually:
  $stagedAddinTarget
"@
}

if (-not $copiedBundledUap) {
    $nextSteps += @"

Still missing:
- Bundled UAP workbook was not packaged into the installer.
- Put the UAP workbook here manually:
  $uapTarget
"@
}

if (-not $copiedBundledMih) {
    $nextSteps += @"

Still missing:
- Bundled MIH workbook was not packaged into the installer.
- Put the MIH workbook here manually:
  $mihTarget
"@
}

[System.IO.File]::WriteAllText($nextStepsPath, $nextSteps, [System.Text.UTF8Encoding]::new($false))

Write-Host ''
Write-Host "Install complete."
Write-Host "Read this file next:"
Write-Host $nextStepsPath
