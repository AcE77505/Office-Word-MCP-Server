param(
    [string]$HostAddress = "0.0.0.0",
    [int]$Port = 8000,
    [string]$Path = "/mcp"
)

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

$pythonCmd = $null
$venvPython = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"

if (Test-Path $venvPython) {
    $pythonCmd = $venvPython
} elseif (Get-Command python -ErrorAction SilentlyContinue) {
    $pythonCmd = "python"
} elseif (Get-Command py -ErrorAction SilentlyContinue) {
    $pythonCmd = "py"
} else {
    Write-Host "Python not found. Install Python 3.11+ or create .venv in this project." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

$env:MCP_TRANSPORT = "streamable-http"
$env:MCP_HOST = $HostAddress
$env:MCP_PORT = "$Port"
$env:MCP_PATH = $Path
$env:DANGEROUSLY_DISABLE_DNS_REBINDING_PROTECTION = "true"

Write-Host "Starting Office Word MCP Server..." -ForegroundColor Cyan
Write-Host "Transport : $env:MCP_TRANSPORT"
Write-Host "Endpoint  : http://$HostAddress`:$Port$Path"
Write-Host "For LAN access, keep HostAddress=0.0.0.0 and allow this port in firewall." -ForegroundColor Yellow
Write-Host "DNS rebinding protection: DISABLED (unsafe; use only on trusted LAN)." -ForegroundColor Yellow
Write-Host ""

& $pythonCmd "$PSScriptRoot\word_mcp_server.py"

if ($LASTEXITCODE -ne 0) {
    Write-Host "Server exited with code: $LASTEXITCODE" -ForegroundColor Red
}

Read-Host "Server stopped. Press Enter to exit"
