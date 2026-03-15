# Excel Agent Foundation

This branch adds a PowerShell-first foundation for a client-PC-resident AMI Optix agent that builds and tests against a **staged add-in container** instead of rebuilding the `.xlam` from raw assets.

## Why the staged-container model is required

Two current build dependencies prevent a safe fresh rebuild:

- `excel-addin/customUI/customUI14.xml` is packaged ribbon XML and must be embedded into the add-in outside normal VBA import/export.
- `excel-addin/forms/frmUtilities.frm` references a companion `.frx` file that is not present in source control.

Because of that, the agent:

- imports only the VBA modules/classes into a known-good staged `.xlam`
- preserves ribbon/package/form assets already present in that staged container
- refuses to mutate ribbon XML/package content automatically
- flags missing staged assets as manual actions instead of guessing

## Files added

- `tools/excel-agent/AmiOptix.Agent.psm1`
- `tools/excel-agent/Bootstrap-AmiOptixAgent.ps1`
- `tools/excel-agent/Build-StagedAddin.ps1`
- `tools/excel-agent/Invoke-AmiOptixAcceptance.ps1`
- `tools/excel-agent/Run-AmiOptixAutofix.ps1`
- `tools/excel-agent/config/agent-config.template.json`
- `tools/excel-agent/config/acceptance.template.json`

## What the foundation does

### 1. Generates a local manifest

`Bootstrap-AmiOptixAgent.ps1` creates `C:\AMI_Optix_Agent\state\workspace-manifest.json` and classifies repo assets into:

- auto-editable
- environment-managed
- manual/package-managed
- security-sensitive

The manifest is generated from the current workspace snapshot, not from stale docs.

### 2. Runs preflight checks

The bootstrap/orchestrator validates:

- Excel COM availability
- VBProject trust access
- outbound app API reachability
- outbound OpenAI reachability
- staged container existence
- golden workbook existence
- API key presence in the normal AMI Optix registry location

Blocking items are written as structured manual action requests.

### 3. Builds in place

`Build-StagedAddin.ps1`:

- copies the staged container to a build output path
- removes old AMI Optix modules/classes from the copy
- reimports current module/class files from the local source snapshot
- runs the VBE compile command
- saves the rebuilt add-in

It does **not** rebuild forms or ribbon/package parts.

### 4. Runs a guarded acceptance suite

`Invoke-AmiOptixAcceptance.ps1` reads `acceptance.json` and runs configured scenarios.

Current default suite intentionally marks `RunOptimizationForProgram` and `VerifyManualRentsAPI` as `requiresInteractiveUi` because those entrypoints still raise modal `MsgBox` UI. In unattended mode, the runner classifies them as a manual automation gap instead of hanging Excel.

This is deliberate: the agent should report that unattended execution is blocked, not guess around modal dialogs.

### 5. Enforces security/manual gates

`Run-AmiOptixAutofix.ps1` accepts optional `-RequestedFixFiles`.

If the requested scope includes:

- manual/package-managed files
- security-sensitive files

the run stops before build/test and prints why. Security-sensitive work must be explained and approved before mutation.

## Recommended setup

### Easiest path: one zip + one command

1. Copy this repo/branch to the client PC as a zip and extract it anywhere.
2. Open PowerShell in the extracted folder.
3. Run:
   - `powershell -ExecutionPolicy Bypass -File .\tools\excel-agent\Install-AmiOptixAgent.ps1`
4. The installer creates `C:\AMI_Optix_Agent\`, copies the scripts/source there, tries to copy the already-installed `AMI_Optix.xlam`, copies the bundled UAP/MIH golden workbooks, runs bootstrap, and writes `C:\AMI_Optix_Agent\NEXT-STEPS.txt`.

### After the installer finishes

1. Read `C:\AMI_Optix_Agent\NEXT-STEPS.txt`.
2. If the installer found the installed add-in and bundled workbooks, you do **not** need to copy those manually.
3. Add the API key through the normal Excel UI once.
4. Run:
   - `powershell -ExecutionPolicy Bypass -File C:\AMI_Optix_Agent\scripts\Run-AmiOptixAutofix.ps1 -AgentRoot C:\AMI_Optix_Agent`

## Current limits

- No ribbon/package mutation automation yet
- No form reconstruction when `.frx` is missing
- No unattended execution for VBA entrypoints that still raise modal dialogs
- No automatic security-sensitive mutation

Those are explicit guardrails, not omissions.

## Updating scripts without re-extracting the zip

After the first install, you can refresh the agent scripts directly from GitHub with:

- `powershell -ExecutionPolicy Bypass -File C:\AMI_Optix_Agent\scripts\Refresh-AmiOptixAgentFromGitHub.ps1 -AgentRoot C:\AMI_Optix_Agent`
