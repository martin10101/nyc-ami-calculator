Set-StrictMode -Version Latest

function Get-AmiOptixRelativePath {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$Path
    )

    $rootPath = (Resolve-Path -LiteralPath $Root).Path.TrimEnd('\') + '\'
    $targetPath = (Resolve-Path -LiteralPath $Path).Path
    $rootUri = [System.Uri]$rootPath
    $targetUri = [System.Uri]$targetPath
    return [System.Uri]::UnescapeDataString($rootUri.MakeRelativeUri($targetUri).ToString()).Replace('/', '\')
}

function Write-AmiOptixJsonFile {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)]$InputObject,
        [int]$Depth = 12
    )

    $parent = Split-Path -Parent $Path
    if ($parent) {
        New-Item -ItemType Directory -Force -Path $parent | Out-Null
    }

    $json = $InputObject | ConvertTo-Json -Depth $Depth
    [System.IO.File]::WriteAllText($Path, $json, [System.Text.UTF8Encoding]::new($false))
}

function Read-AmiOptixJsonFile {
    param(
        [Parameter(Mandatory = $true)][string]$Path
    )

    return Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
}

function New-AmiOptixManualActionRequest {
    param(
        [Parameter(Mandatory = $true)][string]$Classification,
        [Parameter(Mandatory = $true)][string]$Blocked,
        [Parameter(Mandatory = $true)][string]$Why,
        [Parameter(Mandatory = $true)][string[]]$MinimumManualSteps,
        [Parameter(Mandatory = $true)][string]$ContinueWhen
    )

    return [ordered]@{
        classification = $Classification
        blocked = $Blocked
        why = $Why
        minimumManualSteps = @($MinimumManualSteps)
        continueWhen = $ContinueWhen
    }
}

function Test-AmiOptixObjectProperty {
    param(
        [Parameter(Mandatory = $true)]$InputObject,
        [Parameter(Mandatory = $true)][string]$Name
    )

    if ($null -eq $InputObject) {
        return $false
    }

    return $null -ne $InputObject.PSObject.Properties[$Name]
}

function Get-AmiOptixObjectProperty {
    param(
        [Parameter(Mandatory = $true)]$InputObject,
        [Parameter(Mandatory = $true)][string]$Name,
        $Default = $null
    )

    if (Test-AmiOptixObjectProperty -InputObject $InputObject -Name $Name) {
        return $InputObject.$Name
    }

    return $Default
}

function Get-AmiOptixSecuritySensitiveRelativePaths {
    return @(
        'app.py',
        'excel-addin\src\AMI_Optix_API.bas',
        'excel-addin\src\AMI_Optix_Main.bas',
        'excel-addin\src\AMI_Optix_Setup.bas'
    )
}

function Get-AmiOptixEnvironmentDependencies {
    return @(
        [ordered]@{
            type = 'registry'
            path = 'HKCU\Software\VB and VBA Program Settings\AMI_Optix'
            description = 'GetSetting/SaveSetting state for API key, utilities, debug, learning, and rent-roll year.'
        },
        [ordered]@{
            type = 'filesystem'
            path = '%APPDATA%\AMI_Optix\RentRollYears\<YEAR>'
            description = 'Per-user rent calculator workbook storage.'
        },
        [ordered]@{
            type = 'filesystem'
            path = '%APPDATA%\AMI_Optix\RentTablesCache\<YEAR>'
            description = 'Per-user rent table CSV cache.'
        },
        [ordered]@{
            type = 'filesystem'
            path = 'Z:\AMI_Optix\RentRollYears\<YEAR>'
            description = 'Authoritative shared rent calculator workbook source.'
        }
    )
}

function Get-AmiOptixSourceManifest {
    param(
        [Parameter(Mandatory = $true)][string]$SourceRoot,
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][string]$AgentRoot
    )

    $securitySensitive = Get-AmiOptixSecuritySensitiveRelativePaths
    $sourceRootPath = (Resolve-Path -LiteralPath $SourceRoot).Path
    $repoRootPath = (Resolve-Path -LiteralPath $RepoRoot).Path

    $srcEntries = New-Object System.Collections.Generic.List[object]
    Get-ChildItem -LiteralPath (Join-Path $SourceRoot 'excel-addin\src') -File | Sort-Object Name | ForEach-Object {
        $relativePath = Get-AmiOptixRelativePath -Root $SourceRoot -Path $_.FullName
        $extension = $_.Extension.ToLowerInvariant()
        $componentType = if ($extension -eq '.cls') { 'ClassModule' } else { 'StandardModule' }
        $canAutoPatch = $securitySensitive -notcontains $relativePath
        $notes = if ($canAutoPatch) {
            @('Auto-editable by the agent when the requested fix is in scope.')
        } else {
            @('Security-sensitive: agent may import for sync/build, but must not auto-patch without explicit approval.')
        }

        $srcEntries.Add([ordered]@{
            relativePath = $relativePath
            absolutePath = $_.FullName
            componentName = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
            extension = $extension
            componentType = $componentType
            canAutoPatch = $canAutoPatch
            canImport = $true
            notes = $notes
        })
    }

    $formEntries = New-Object System.Collections.Generic.List[object]
    Get-ChildItem -LiteralPath (Join-Path $SourceRoot 'excel-addin\forms') -Filter *.frm -File -ErrorAction SilentlyContinue | Sort-Object Name | ForEach-Object {
        $frxPath = [System.IO.Path]::ChangeExtension($_.FullName, '.frx')
        $hasFrx = Test-Path -LiteralPath $frxPath
        $formEntries.Add([ordered]@{
            relativePath = Get-AmiOptixRelativePath -Root $SourceRoot -Path $_.FullName
            absolutePath = $_.FullName
            companionFrxPath = if ($hasFrx) { $frxPath } else { '' }
            hasCompanionFrx = $hasFrx
            requiresManualBuild = -not $hasFrx
            notes = if ($hasFrx) {
                @('Form assets are complete; form can be re-imported if explicitly allowed.')
            } else {
                @('Companion .frx asset is missing; preserve the form inside the staged container and do not rebuild it from source.')
            }
        })
    }

    $ribbonXmlPath = Join-Path $SourceRoot 'excel-addin\customUI\customUI14.xml'

    return [ordered]@{
        version = 1
        generatedAtUtc = [DateTime]::UtcNow.ToString('o')
        repoRoot = $repoRootPath
        agentRoot = $AgentRoot
        sourceRoot = $sourceRootPath
        stagedContainer = [ordered]@{
            path = (Join-Path $AgentRoot 'staging\AMI_Optix_Staged.xlam')
            description = 'Known-good add-in container that already contains the embedded ribbon and any surviving form assets.'
            required = $true
        }
        workbooks = [ordered]@{
            uap = [ordered]@{
                path = (Join-Path $AgentRoot 'workbooks\UAP_golden.xlsm')
                required = $true
            }
            mih = [ordered]@{
                path = (Join-Path $AgentRoot 'workbooks\MIH_golden.xlsm')
                required = $true
            }
        }
        sourceFiles = [ordered]@{
            src = $srcEntries
            forms = $formEntries
            ribbon = @(
                [ordered]@{
                    relativePath = Get-AmiOptixRelativePath -Root $SourceRoot -Path $ribbonXmlPath
                    absolutePath = $ribbonXmlPath
                    requiresManualBuild = $true
                    notes = @('Ribbon XML lives outside VBProject import/export and must stay embedded inside the staged container.')
                }
            )
        }
        classification = [ordered]@{
            autoEditable = @($srcEntries | Where-Object { $_.canAutoPatch } | ForEach-Object { $_.relativePath })
            environmentManaged = @(Get-AmiOptixEnvironmentDependencies)
            manualPackageManaged = @(
                @($formEntries | Where-Object { $_.requiresManualBuild } | ForEach-Object {
                    [ordered]@{
                        relativePath = $_.relativePath
                        reason = 'Missing companion .frx asset; preserve this form inside the staged container.'
                    }
                }) + @(
                    [ordered]@{
                        relativePath = 'excel-addin\customUI\customUI14.xml'
                        reason = 'Ribbon XML must be embedded into the .xlam package with OfficeRibbonXEditor or equivalent package editing.'
                    }
                )
            )
            securitySensitive = @(
                $securitySensitive | ForEach-Object {
                    [ordered]@{
                        relativePath = $_
                        reason = 'Contains API key/auth/admin/security-relevant logic; require explanation and approval before mutation.'
                    }
                }
            )
        }
        acceptanceDefaults = [ordered]@{
            macros = [ordered]@{
                runUap = 'AMI_Optix_Automation.RunOptimizationUAP_Agent'
                runMih = 'AMI_Optix_Automation.RunOptimizationMIH_Agent'
                diagnostics = 'AMI_Optix_Diagnostics.ShowAMIOptixDiagnostics'
                verifyManualRents = 'AMI_Optix_VerifyManualRents.VerifyManualRentsAPI_Agent'
            }
        }
        notes = @(
            'Use the generated manifest as the source of truth for automated build/test decisions.',
            'Do not treat install docs as authoritative for build composition.',
            'Import modules/classes from source; preserve forms and ribbon/package content from the staged container.'
        )
    }
}

function New-AmiOptixWorkspaceManifest {
    param(
        [Parameter(Mandatory = $true)][string]$SourceRoot,
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][string]$AgentRoot,
        [Parameter(Mandatory = $true)][string]$ManifestPath
    )

    $manifest = Get-AmiOptixSourceManifest -SourceRoot $SourceRoot -RepoRoot $RepoRoot -AgentRoot $AgentRoot
    Write-AmiOptixJsonFile -Path $ManifestPath -InputObject $manifest
    return $manifest
}

function Get-AmiOptixDependencyClassification {
    param(
        [Parameter(Mandatory = $true)]$Manifest,
        [string[]]$RequestedRelativePaths
    )

    $manualLookup = @{}
    foreach ($entry in $Manifest.classification.manualPackageManaged) {
        $manualLookup[$entry.relativePath] = $entry.reason
    }

    $securityLookup = @{}
    foreach ($entry in $Manifest.classification.securitySensitive) {
        $securityLookup[$entry.relativePath] = $entry.reason
    }

    $autoEditable = @($Manifest.classification.autoEditable)
    if (-not $RequestedRelativePaths -or $RequestedRelativePaths.Count -eq 0) {
        return [ordered]@{
            autoEditable = $autoEditable
            environmentManaged = @($Manifest.classification.environmentManaged)
            manualPackageManaged = @($Manifest.classification.manualPackageManaged)
            securitySensitive = @($Manifest.classification.securitySensitive)
        }
    }

    $result = [ordered]@{
        autoEditable = @()
        manualPackageManaged = @()
        securitySensitive = @()
        unknown = @()
    }

    foreach ($relativePath in $RequestedRelativePaths) {
        if ($autoEditable -contains $relativePath) {
            $result.autoEditable += $relativePath
            continue
        }

        if ($manualLookup.ContainsKey($relativePath)) {
            $result.manualPackageManaged += [ordered]@{
                relativePath = $relativePath
                reason = $manualLookup[$relativePath]
            }
            continue
        }

        if ($securityLookup.ContainsKey($relativePath)) {
            $result.securitySensitive += [ordered]@{
                relativePath = $relativePath
                reason = $securityLookup[$relativePath]
            }
            continue
        }

        $result.unknown += $relativePath
    }

    return $result
}

function New-AmiOptixExcelApplication {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try {
        $excel.AutomationSecurity = 1
    } catch {}
    return $excel
}

function Close-AmiOptixExcelApplication {
    param(
        [Parameter()][object]$Workbook,
        [Parameter()][object]$Excel
    )

    if ($Workbook) {
        try { $Workbook.Close($false) } catch {}
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook) } catch {}
    }

    if ($Excel) {
        try { $Excel.Quit() } catch {}
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) } catch {}
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function Test-AmiOptixVbProjectAccess {
    param(
        [Parameter(Mandatory = $true)][object]$Excel
    )

    $workbook = $null
    try {
        $workbook = $Excel.Workbooks.Add()
        $count = $workbook.VBProject.VBComponents.Count
        return [ordered]@{
            succeeded = $true
            details = "VBProject access available ($count components visible)."
        }
    } catch {
        return [ordered]@{
            succeeded = $false
            details = $_.Exception.Message
        }
    } finally {
        if ($workbook) {
            try { $workbook.Close($false) } catch {}
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) } catch {}
        }
    }
}

function Test-AmiOptixEnvironment {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)]$Manifest,
        [string]$ResultPath
    )

    $checks = New-Object System.Collections.Generic.List[object]
    $manualActions = New-Object System.Collections.Generic.List[object]
    $apiKeyRegistryPath = 'HKCU:\Software\VB and VBA Program Settings\AMI_Optix\Settings'
    $apiKeyPresent = $false

    $excel = $null
    try {
        $excel = New-AmiOptixExcelApplication
        $checks.Add([ordered]@{
            name = 'excel_com'
            succeeded = $true
            details = "Excel COM available (version $($excel.Version))."
        })

        $vbeCheck = Test-AmiOptixVbProjectAccess -Excel $excel
        $checks.Add([ordered]@{
            name = 'vbproject_access'
            succeeded = $vbeCheck.succeeded
            details = $vbeCheck.details
        })
        if (-not $vbeCheck.succeeded) {
            $manualActions.Add((New-AmiOptixManualActionRequest `
                -Classification 'manual/package-managed' `
                -Blocked 'VBProject automation is blocked' `
                -Why 'Excel Trust Center is blocking programmatic access to the VBA project object model, so the staged add-in cannot be updated safely.' `
                -MinimumManualSteps @(
                    'Open Excel > File > Options > Trust Center > Trust Center Settings > Macro Settings.',
                    'Enable "Trust access to the VBA project object model".',
                    'Restart Excel and rerun the agent preflight.'
                ) `
                -ContinueWhen 'After VBProject access succeeds, rerun the staged build.'))
        }
    } catch {
        $checks.Add([ordered]@{
            name = 'excel_com'
            succeeded = $false
            details = $_.Exception.Message
        })
        $manualActions.Add((New-AmiOptixManualActionRequest `
            -Classification 'environment-failure' `
            -Blocked 'Excel COM is unavailable' `
            -Why 'The agent cannot build or test the add-in without a working local Excel COM automation host.' `
            -MinimumManualSteps @(
                'Confirm desktop Excel is installed on the client PC.',
                'Launch Excel once manually, close it, and rerun the preflight.'
            ) `
            -ContinueWhen 'After Excel COM succeeds, rerun the staged build.'))
    } finally {
        if ($excel) {
            Close-AmiOptixExcelApplication -Excel $excel
        }
    }

    foreach ($healthCheck in @(
        [ordered]@{ name = 'app_api'; url = $Config.healthEndpoints.appApi; accepted = @('200') },
        [ordered]@{ name = 'openai_models'; url = $Config.healthEndpoints.openAiModels; accepted = @('200', '401', '403') }
    )) {
        try {
            $response = Invoke-WebRequest -Uri $healthCheck.url -Method Head -UseBasicParsing -TimeoutSec 20
            $statusCode = [string][int]$response.StatusCode
        } catch {
            $statusCode = ''
            if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
                $statusCode = [string][int]$_.Exception.Response.StatusCode.value__
            }

            if ($statusCode -eq '') {
                $checks.Add([ordered]@{
                    name = $healthCheck.name
                    succeeded = $false
                    details = $_.Exception.Message
                })
                continue
            }
        }

        $checks.Add([ordered]@{
            name = $healthCheck.name
            succeeded = $healthCheck.accepted -contains $statusCode
            details = "HTTP $statusCode from $($healthCheck.url)"
        })
    }

    $stagedContainerPath = $Config.stagedContainerPath
    $checks.Add([ordered]@{
        name = 'staged_container'
        succeeded = (Test-Path -LiteralPath $stagedContainerPath)
        details = $stagedContainerPath
    })
    if (-not (Test-Path -LiteralPath $stagedContainerPath)) {
        $manualActions.Add((New-AmiOptixManualActionRequest `
            -Classification 'manual/package-managed' `
            -Blocked 'Known-good staged add-in container is missing' `
            -Why 'The build is intentionally in-place only; it will not reconstruct ribbon XML or missing form assets from raw sources.' `
            -MinimumManualSteps @(
                'Place a known-good AMI_Optix .xlam or .xlsm container at the configured stagedContainerPath.',
                'Make sure the file already contains the embedded ribbon and working Utilities form assets.'
            ) `
            -ContinueWhen 'After the staged container exists, rerun the staged build.'))
    }

    foreach ($role in @('uap', 'mih')) {
        $workbookPath = $Config.goldenWorkbooks.$role
        $exists = $workbookPath -and (Test-Path -LiteralPath $workbookPath)
        $checks.Add([ordered]@{
            name = "golden_workbook_$role"
            succeeded = $exists
            details = if ($workbookPath) { $workbookPath } else { 'Not configured.' }
        })

        if (-not $exists) {
            $manualActions.Add((New-AmiOptixManualActionRequest `
                -Classification 'environment-failure' `
                -Blocked "Golden workbook for $role is missing" `
                -Why 'The acceptance suite can only run against predetermined workbook fixtures that are stable across iterations.' `
                -MinimumManualSteps @(
                    "Place the approved $role workbook at the configured path.",
                    'Keep a clean golden copy; the runner will create runtime copies per iteration.'
                ) `
                -ContinueWhen "After the $role golden workbook is present, rerun the acceptance suite."))
        }
    }

    try {
        $apiKeyValue = (Get-ItemProperty -LiteralPath $apiKeyRegistryPath -Name APIKey -ErrorAction Stop).APIKey
        $apiKeyPresent = -not [string]::IsNullOrWhiteSpace([string]$apiKeyValue)
    } catch {
        $apiKeyPresent = $false
    }

    $checks.Add([ordered]@{
        name = 'api_key_registry'
        succeeded = $apiKeyPresent
        details = if ($apiKeyPresent) { 'API key present in registry.' } else { 'API key missing from registry.' }
    })
    if (-not $apiKeyPresent) {
        $manualActions.Add((New-AmiOptixManualActionRequest `
            -Classification 'environment-failure' `
            -Blocked 'AMI Optix API key is not configured' `
            -Why 'Optimization and verification entrypoints prompt for an API key and cannot complete unattended when the registry is empty.' `
            -MinimumManualSteps @(
                'Open Excel and launch the AMI Optix API Settings entrypoint once.',
                'Store the API key in the normal AMI_Optix registry location.',
                'Close Excel and rerun the preflight.'
            ) `
            -ContinueWhen 'After the API key is present in the registry, rerun the acceptance suite.'))
    }

    $checks.Add([ordered]@{
        name = 'appdata_root'
        succeeded = $true
        details = (Join-Path $env:APPDATA 'AMI_Optix')
    })

    $result = [ordered]@{
        generatedAtUtc = [DateTime]::UtcNow.ToString('o')
        succeeded = (-not ($checks | Where-Object { -not $_.succeeded })) -and ($manualActions.Count -eq 0)
        checks = $checks
        manualActionRequests = $manualActions
    }

    if ($ResultPath) {
        Write-AmiOptixJsonFile -Path $ResultPath -InputObject $result
    }

    return $result
}

function Invoke-AmiOptixVbaCompile {
    param(
        [Parameter(Mandatory = $true)][object]$Excel
    )

    try {
        $compileControl = $Excel.VBE.CommandBars.FindControl(1, 578, $true, $true)
        if ($null -eq $compileControl) {
            return [ordered]@{
                succeeded = $false
                canContinue = $true
                details = 'Could not locate the VBA Compile command in the VBE command bars.'
            }
        }

        if ($compileControl.Enabled) {
            $compileControl.Execute()
        }

        return [ordered]@{
            succeeded = $true
            canContinue = $true
            details = 'Compile command executed.'
        }
    } catch {
        return [ordered]@{
            succeeded = $false
            canContinue = $false
            details = $_.Exception.Message
        }
    }
}

function Get-AmiOptixImportEntries {
    param(
        [Parameter(Mandatory = $true)]$Manifest
    )

    return @($Manifest.sourceFiles.src | Where-Object { $_.canImport })
}

function Get-AmiOptixComponentSourceCode {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Extension
    )

    $lines = Get-Content -LiteralPath $Path
    $filtered = New-Object System.Collections.Generic.List[string]
    $skipClassHeader = ($Extension -eq '.cls')
    $insideClassHeader = $false

    foreach ($line in $lines) {
        if ($skipClassHeader) {
            if ($line -match '^BEGIN\s*$') {
                $insideClassHeader = $true
                continue
            }

            if ($insideClassHeader) {
                if ($line -match '^END\s*$') {
                    $insideClassHeader = $false
                }
                continue
            }

            if ($line -match '^VERSION\s+1\.0\s+CLASS$') {
                continue
            }
        }

        if ($line -match '^Attribute VB_') {
            continue
        }

        $filtered.Add($line)
    }

    return ($filtered -join [Environment]::NewLine)
}

function Import-AmiOptixComponent {
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)]$Entry
    )

    $extension = [string]$Entry.extension
    if ($extension -eq '.cls') {
        $component = $Project.VBComponents.Add(2)
        $component.Name = [string]$Entry.componentName
        $code = Get-AmiOptixComponentSourceCode -Path $Entry.absolutePath -Extension $extension
        $component.CodeModule.AddFromString($code)
        return
    }

    $Project.VBComponents.Import($Entry.absolutePath) | Out-Null
}

function Invoke-AmiOptixStagedBuild {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)]$Manifest,
        [string]$ResultPath
    )

    $result = [ordered]@{
        generatedAtUtc = [DateTime]::UtcNow.ToString('o')
        succeeded = $false
        imported = @()
        removed = @()
        compile = $null
        outputPath = $Config.rebuiltAddinPath
        manualActionRequests = @()
    }

    if (-not (Test-Path -LiteralPath $Config.stagedContainerPath)) {
        $result.manualActionRequests = @(
            New-AmiOptixManualActionRequest `
                -Classification 'manual/package-managed' `
                -Blocked 'Staged container is missing' `
                -Why 'The staged build intentionally preserves ribbon/package/form assets from an existing known-good container.' `
                -MinimumManualSteps @(
                    'Place a known-good AMI_Optix staged container at the configured path.',
                    'Ensure it already contains the ribbon XML and Utilities form assets.'
                ) `
                -ContinueWhen 'After the staged container exists, rerun Build-StagedAddin.ps1.'
        )
        if ($ResultPath) {
            Write-AmiOptixJsonFile -Path $ResultPath -InputObject $result
        }
        return $result
    }

    $outputDirectory = Split-Path -Parent $Config.rebuiltAddinPath
    New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
    Copy-Item -LiteralPath $Config.stagedContainerPath -Destination $Config.rebuiltAddinPath -Force

    $excel = $null
    $workbook = $null
    try {
        $excel = New-AmiOptixExcelApplication
        $workbook = $excel.Workbooks.Open($Config.rebuiltAddinPath)
        $project = $workbook.VBProject

        $importEntries = Get-AmiOptixImportEntries -Manifest $Manifest
        $allowedNames = @($importEntries | ForEach-Object { $_.componentName })
        $removed = New-Object System.Collections.Generic.List[string]

        for ($index = $project.VBComponents.Count; $index -ge 1; $index--) {
            $component = $project.VBComponents.Item($index)
            $name = [string]$component.Name
            $type = [int]$component.Type

            $isManagedCodeComponent = (($type -eq 1) -or ($type -eq 2))
            $isAmiOptixComponent = $name -like 'AMI_Optix*'

            if ($isManagedCodeComponent -and ($isAmiOptixComponent -or ($allowedNames -contains $name))) {
                $project.VBComponents.Remove($component)
                $removed.Add($name)
            }
        }

        foreach ($entry in $importEntries) {
            Import-AmiOptixComponent -Project $project -Entry $entry
            $result.imported += $entry.relativePath
        }

        $result.removed = @($removed)
        $result.compile = Invoke-AmiOptixVbaCompile -Excel $excel
        if (-not $result.compile.succeeded -and -not $result.compile.canContinue) {
            throw "VBA compile failed: $($result.compile.details)"
        }

        $workbook.Save()
        $result.succeeded = $true
        if (-not $result.compile.succeeded -and $result.compile.canContinue) {
            $result.compileWarning = 'Compile command was unavailable on this machine; saved the rebuilt add-in and deferred compile validation to the acceptance run.'
        }
    } catch {
        $result.error = $_.Exception.Message
    } finally {
        if ($workbook -or $excel) {
            Close-AmiOptixExcelApplication -Workbook $workbook -Excel $excel
        }
    }

    if ($ResultPath) {
        Write-AmiOptixJsonFile -Path $ResultPath -InputObject $result
    }

    return $result
}

function Invoke-AmiOptixExcelMacro {
    param(
        [Parameter(Mandatory = $true)][object]$Excel,
        [Parameter(Mandatory = $true)][string]$MacroName,
        [object[]]$Arguments = @()
    )

    switch ($Arguments.Count) {
        0 { return $Excel.Run($MacroName) }
        1 { return $Excel.Run($MacroName, $Arguments[0]) }
        2 { return $Excel.Run($MacroName, $Arguments[0], $Arguments[1]) }
        3 { return $Excel.Run($MacroName, $Arguments[0], $Arguments[1], $Arguments[2]) }
        default { throw "Macro invocation with $($Arguments.Count) arguments is not supported by Invoke-AmiOptixExcelMacro." }
    }
}

function Get-AmiOptixWorksheetByName {
    param(
        [Parameter(Mandatory = $true)][object]$Workbook,
        [Parameter(Mandatory = $true)][string]$WorksheetName
    )

    for ($index = 1; $index -le [int]$Workbook.Worksheets.Count; $index++) {
        $worksheet = $Workbook.Worksheets.Item($index)
        if ([string]$worksheet.Name -eq $WorksheetName) {
            return $worksheet
        }
    }

    throw "Worksheet '$WorksheetName' does not exist."
}

function Test-AmiOptixAssertion {
    param(
        [Parameter(Mandatory = $true)][object]$Workbook,
        [Parameter(Mandatory = $true)]$Assertion
    )

    $type = [string]$Assertion.type
    switch ($type) {
        'sheet_exists' {
            try {
                $null = Get-AmiOptixWorksheetByName -Workbook $Workbook -WorksheetName ([string]$Assertion.sheet)
                return [ordered]@{ succeeded = $true; details = "Worksheet '$($Assertion.sheet)' exists." }
            } catch {
                return [ordered]@{ succeeded = $false; details = "Worksheet '$($Assertion.sheet)' does not exist." }
            }
        }
        'cell_contains' {
            try {
                $sheet = Get-AmiOptixWorksheetByName -Workbook $Workbook -WorksheetName ([string]$Assertion.sheet)
                $value = [string]$sheet.Range([string]$Assertion.cell).Text
                $needle = [string]$Assertion.text
                return [ordered]@{
                    succeeded = $value -like "*$needle*"
                    details = "Cell $($Assertion.sheet)!$($Assertion.cell) = '$value'"
                }
            } catch {
                return [ordered]@{ succeeded = $false; details = $_.Exception.Message }
            }
        }
        'sheet_contains_text' {
            try {
                $sheet = Get-AmiOptixWorksheetByName -Workbook $Workbook -WorksheetName ([string]$Assertion.sheet)
                $range = $sheet.UsedRange
                $value = [string]$range.Text
                $needle = [string]$Assertion.text
                return [ordered]@{
                    succeeded = $value -like "*$needle*"
                    details = "Scanned UsedRange for '$needle'."
                }
            } catch {
                return [ordered]@{ succeeded = $false; details = $_.Exception.Message }
            }
        }
        default {
            return [ordered]@{
                succeeded = $false
                details = "Unsupported assertion type '$type'."
            }
        }
    }
}

function Invoke-AmiOptixAcceptanceSuite {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)]$Manifest,
        [string]$ResultPath
    )

    $suite = Read-AmiOptixJsonFile -Path $Config.acceptancePath
    $suiteResults = New-Object System.Collections.Generic.List[object]
    $manualActions = New-Object System.Collections.Generic.List[object]
    $runtimeRoot = Join-Path $Config.agentRoot 'workbooks\runtime'
    New-Item -ItemType Directory -Force -Path $runtimeRoot | Out-Null

    foreach ($scenario in $suite.scenarios) {
        $scenarioEnabled = Get-AmiOptixObjectProperty -InputObject $scenario -Name 'enabled' -Default $true
        if ($scenarioEnabled -eq $false) {
            continue
        }

        $scenarioName = [string](Get-AmiOptixObjectProperty -InputObject $scenario -Name 'name' -Default 'Unnamed scenario')
        $requiresInteractiveUi = [bool](Get-AmiOptixObjectProperty -InputObject $scenario -Name 'requiresInteractiveUi' -Default $false)
        $workbookRole = [string](Get-AmiOptixObjectProperty -InputObject $scenario -Name 'workbookRole' -Default '')
        $macro = [string](Get-AmiOptixObjectProperty -InputObject $scenario -Name 'macro' -Default '')
        $scenarioArguments = Get-AmiOptixObjectProperty -InputObject $scenario -Name 'arguments' -Default @()
        $scenarioAssertions = Get-AmiOptixObjectProperty -InputObject $scenario -Name 'assertions' -Default @()

        $scenarioResult = [ordered]@{
            name = $scenarioName
            succeeded = $false
            classification = 'code-failure'
            details = ''
            assertions = @()
        }

        if ($requiresInteractiveUi -eq $true -and -not $Config.allowInteractiveUi) {
            $scenarioResult.classification = 'manual/automation-gap'
            $scenarioResult.details = 'Scenario is marked as interactive and would block unattended COM automation because the current VBA entrypoint raises modal UI.'
            $suiteResults.Add($scenarioResult)
            $manualActions.Add((New-AmiOptixManualActionRequest `
                -Classification 'manual/automation-gap' `
                -Blocked "Scenario '$scenarioName' requires interactive UI" `
                -Why 'The configured macro shows modal MsgBox/InputBox UI and cannot be exercised unattended without a dedicated automation-safe test mode.' `
                -MinimumManualSteps @(
                    'Either run this scenario manually and record the result, or add a test-safe non-modal entrypoint before enabling unattended execution.',
                    'Keep the production ribbon behavior unchanged unless the requested fix explicitly targets automation support.'
                ) `
                -ContinueWhen 'After a non-modal test path exists, rerun the acceptance suite.'))
            continue
        }

        $workbook = $null
        $excel = $null
        $addin = $null
        $scenarioStep = 'initializing scenario'
        try {
            $scenarioStep = 'opening Excel application'
            $excel = New-AmiOptixExcelApplication
            $scenarioStep = 'opening rebuilt add-in'
            $addin = $excel.Workbooks.Open($Config.rebuiltAddinPath)

            if ($workbookRole) {
                $scenarioStep = "preparing runtime workbook for role '$workbookRole'"
                $goldenPath = $Config.goldenWorkbooks.$workbookRole
                $runtimeExtension = [System.IO.Path]::GetExtension([string]$goldenPath)
                if ([string]::IsNullOrWhiteSpace($runtimeExtension)) {
                    $runtimeExtension = '.xlsm'
                }
                $runtimeWorkbookPath = Join-Path $runtimeRoot ("{0}-{1:yyyyMMddHHmmss}{2}" -f $workbookRole, [DateTime]::UtcNow, $runtimeExtension)
                Copy-Item -LiteralPath $goldenPath -Destination $runtimeWorkbookPath -Force
                $scenarioStep = 'opening runtime workbook'
                $workbook = $excel.Workbooks.Open($runtimeWorkbookPath)
                $scenarioStep = 'activating runtime workbook'
                $workbook.Activate() | Out-Null
            } else {
                $workbook = $addin
            }

            $addinName = [string]$addin.Name
            if ($addinName -match '\s') {
                $macroName = "'{0}'!{1}" -f $addinName, $macro
            } else {
                $macroName = "{0}!{1}" -f $addinName, $macro
            }
            $arguments = @()
            if ($scenarioArguments) {
                $arguments = @($scenarioArguments)
            }

            $scenarioStep = "running macro '$macroName'"
            $null = Invoke-AmiOptixExcelMacro -Excel $excel -MacroName $macroName -Arguments $arguments

            $scenarioStep = 'running assertions'
            $assertionResults = New-Object System.Collections.Generic.List[object]
            foreach ($assertion in @($scenarioAssertions)) {
                $assertionResults.Add((Test-AmiOptixAssertion -Workbook $workbook -Assertion $assertion))
            }

            $scenarioResult.assertions = @($assertionResults)
            $scenarioResult.succeeded = -not ($assertionResults | Where-Object { -not $_.succeeded })
            $scenarioResult.classification = if ($scenarioResult.succeeded) { 'passed' } else { 'code-failure' }
            $scenarioResult.details = if ($scenarioResult.succeeded) { 'Scenario completed and all assertions passed.' } else { 'One or more assertions failed.' }
        } catch {
            $scenarioResult.classification = 'code-failure'
            $scenarioResult.details = "$scenarioStep :: $($_.Exception.Message)"
        } finally {
            if ($workbook -or $excel) {
                Close-AmiOptixExcelApplication -Workbook $workbook -Excel $excel
            }
            if ($addin) {
                try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($addin) } catch {}
            }
        }

        $suiteResults.Add($scenarioResult)
    }

    $failedScenarios = New-Object System.Collections.Generic.List[object]
    foreach ($suiteResult in $suiteResults) {
        if (-not [bool]$suiteResult.succeeded) {
            $failedScenarios.Add($suiteResult)
        }
    }

    $scenarioArray = @($suiteResults.ToArray())
    $manualActionArray = @($manualActions.ToArray())

    $result = [ordered]@{
        generatedAtUtc = [DateTime]::UtcNow.ToString('o')
        succeeded = (($failedScenarios.Count -eq 0) -and ($manualActions.Count -eq 0))
        scenarios = $scenarioArray
        manualActionRequests = $manualActionArray
    }

    if ($ResultPath) {
        Write-AmiOptixJsonFile -Path $ResultPath -InputObject $result
    }

    return $result
}

Export-ModuleMember -Function `
    Read-AmiOptixJsonFile, `
    Write-AmiOptixJsonFile, `
    New-AmiOptixWorkspaceManifest, `
    Get-AmiOptixDependencyClassification, `
    Test-AmiOptixEnvironment, `
    Invoke-AmiOptixStagedBuild, `
    Invoke-AmiOptixAcceptanceSuite
