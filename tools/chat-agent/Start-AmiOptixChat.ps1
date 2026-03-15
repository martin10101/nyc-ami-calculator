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

# ── Check Node.js ─────────────────────────────────────────────────────────────
try {
    $nodeVersion = & node --version 2>&1
    Write-Host "Node.js found: $nodeVersion"
} catch {
    Write-Host ''
    Write-Host 'ERROR: Node.js is not installed or not on PATH.' -ForegroundColor Red
    Write-Host 'Download it from: https://nodejs.org  (LTS version, Windows Installer)' -ForegroundColor Yellow
    Write-Host 'After installing, re-run this script.' -ForegroundColor Yellow
    exit 1
}

# ── Install npm dependencies if needed ───────────────────────────────────────
$nodeModules = Join-Path $scriptDir 'node_modules'
if (-not (Test-Path $nodeModules)) {
    Write-Host 'Installing dependencies (first run only)...'
    Push-Location $scriptDir
    try {
        & npm install --no-fund --no-audit 2>&1 | Out-Host
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
$job = Start-Job -ScriptBlock {
    param($u)
    Start-Sleep 2
    Start-Process $u
} -ArgumentList $url

$serverScript = Join-Path $scriptDir 'server.js'
& node $serverScript
