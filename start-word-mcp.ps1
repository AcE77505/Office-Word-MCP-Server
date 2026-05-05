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
    Write-Host "未找到 Python。请先安装 Python 3.11+，或在项目目录创建 .venv。" -ForegroundColor Red
    Read-Host "按回车退出"
    exit 1
}

$env:MCP_TRANSPORT = "streamable-http"
$env:MCP_HOST = $HostAddress
$env:MCP_PORT = "$Port"
$env:MCP_PATH = $Path

Write-Host "启动 Office Word MCP Server..." -ForegroundColor Cyan
Write-Host "Transport : $env:MCP_TRANSPORT"
Write-Host "Endpoint  : http://$HostAddress`:$Port$Path"
Write-Host "局域网访问请将 HostAddress 设为 0.0.0.0，并开放防火墙端口。" -ForegroundColor Yellow
Write-Host ""

& $pythonCmd "$PSScriptRoot\word_mcp_server.py"

if ($LASTEXITCODE -ne 0) {
    Write-Host "服务异常退出，退出码: $LASTEXITCODE" -ForegroundColor Red
}

Read-Host "服务已停止。按回车退出"
