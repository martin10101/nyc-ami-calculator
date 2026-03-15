param(
    [string]$AgentRoot  = 'C:\AMI_Optix_Agent',
    [string]$RepoRoot   = '',
    [string]$OpenAiKey  = '',
    [int]   $Port       = 3000
)

$ErrorActionPreference = 'Stop'
$scriptDir = $PSScriptRoot

# ── Resolve RepoRoot ──────────────────────────────────────────────────────────
if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    # Walk up from this script to find the repo root (contains tools\excel-agent)
    $candidate = Split-Path $scriptDir -Parent  # tools\
    $candidate = Split-Path $candidate -Parent  # repo root
    if (Test-Path (Join-Path $candidate 'tools\excel-agent\AmiOptix.Agent.psm1')) {
        $RepoRoot = $candidate
    }
}

# ── Find Node.js (PATH or portable fallback locations) ───────────────────────
$nodeExe = $null
$npmCmd  = $null

$portableCandidates = @(
    'C:\node\node-v20.19.0-win-x64',
    'C:\node\node-v18.20.0-win-x64',
    'C:\node'
)

# Try PATH first
if (Get-Command 'node' -ErrorAction SilentlyContinue) {
    $nodeExe = 'node'
    $npmCmd  = 'npm'
} else {
    foreach ($dir in $portableCandidates) {
        $candidate = Join-Path $dir 'node.exe'
        if (Test-Path $candidate) {
            $nodeExe = $candidate
            $npmCmd  = Join-Path $dir 'npm.cmd'
            # Add to session PATH so child processes can find it
            $env:PATH = "$dir;$env:PATH"
            break
        }
    }
}

if (-not $nodeExe) {
    Write-Host ''
    Write-Host 'ERROR: Node.js not found.' -ForegroundColor Red
    Write-Host 'Run this to install it (no admin needed):' -ForegroundColor Yellow
    Write-Host '  Invoke-WebRequest -Uri "https://nodejs.org/dist/v20.19.0/node-v20.19.0-win-x64.zip" -OutFile "$env:TEMP\node.zip" -UseBasicParsing; Expand-Archive -LiteralPath "$env:TEMP\node.zip" -DestinationPath "C:\node" -Force' -ForegroundColor Yellow
    exit 1
}

$nodeVersion = & $nodeExe --version 2>&1
Write-Host "Node.js found: $nodeVersion"

# ── Install npm dependencies if needed ───────────────────────────────────────
$nodeModules = Join-Path $scriptDir 'node_modules'
if (-not (Test-Path $nodeModules)) {
    Write-Host 'Installing dependencies (first run only)...'
    Push-Location $scriptDir
    try {
        & $npmCmd install --no-fund --no-audit --strict-ssl=false 2>&1 | Out-Host
    } finally {
        Pop-Location
    }
}

# ── OpenAI API key ────────────────────────────────────────────────────────────
if ([string]::IsNullOrWhiteSpace($OpenAiKey)) {
    $OpenAiKey = $env:OPENAI_API_KEY
}
if ([string]::IsNullOrWhiteSpace($OpenAiKey)) {
    $OpenAiKey = Read-Host 'Enter your OpenAI API key (sk-...)'
}
if ([string]::IsNullOrWhiteSpace($OpenAiKey)) {
    Write-Host 'ERROR: OpenAI API key is required.' -ForegroundColor Red
    exit 1
}

# ── Launch server ─────────────────────────────────────────────────────────────
$env:AGENT_ROOT    = $AgentRoot
$env:REPO_ROOT     = $RepoRoot
$env:OPENAI_API_KEY = $OpenAiKey
$env:PORT          = $Port

$url = "http://localhost:$Port"
Write-Host ''
Write-Host "Starting AMI Optix Chat Agent at $url" -ForegroundColor Cyan
Write-Host "Agent root : $AgentRoot"
Write-Host "Repo root  : $RepoRoot"
Write-Host ''
Write-Host 'Opening browser in 2 seconds... (Ctrl+C to stop the server)' -ForegroundColor Gray

# Open browser after a short delay
$null = Start-Job -ScriptBlock {
    param($u)
    Start-Sleep 2
    Start-Process $u
} -ArgumentList $url

$serverScript = Join-Path $scriptDir 'server.js'
& $nodeExe $serverScript
