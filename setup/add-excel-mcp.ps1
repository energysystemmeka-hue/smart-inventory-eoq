# Adds the Excel MCP Server to Claude Desktop's config.
# Run it while Claude Desktop is fully closed (otherwise the app overwrites the change).
# Usage (in a normal PowerShell window):
#   powershell -ExecutionPolicy Bypass -File .\setup\add-excel-mcp.ps1

$ErrorActionPreference = 'Stop'

# 1. Claude must be closed
if (Get-Process -Name claude -ErrorAction SilentlyContinue | Where-Object { $_.Path -like '*WindowsApps\Claude*' -or $_.Path -like '*AnthropicClaude*' }) {
    Write-Host "Claude Desktop is still running. Right-click the Claude icon in the system tray, choose Quit, then run this again." -ForegroundColor Red
    exit 1
}

# 2. Find uvx
$uvx = (Get-Command uvx -ErrorAction SilentlyContinue).Source
if (-not $uvx) {
    $uvx = Get-ChildItem "$env:LOCALAPPDATA\Microsoft\WinGet\Packages" -Recurse -Filter uvx.exe -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
}
if (-not $uvx) { Write-Host "uvx not found. Install uv first: winget install --id=astral-sh.uv -e" -ForegroundColor Red; exit 1 }

# 3. Folder the server may use: <project>\data
$dataDir = Join-Path (Split-Path $PSScriptRoot -Parent) 'data'
New-Item -ItemType Directory -Force $dataDir | Out-Null

# 4. Find Claude Desktop's config file (newer Store builds keep it under Claude-Data)
$candidates = @()
$candidates += Get-ChildItem "$env:LOCALAPPDATA\Claude-Data" -Directory -Filter 'claude_*' -ErrorAction SilentlyContinue |
    ForEach-Object { Join-Path $_.FullName 'claude_desktop_config.json' }
$candidates += "$env:APPDATA\Claude\claude_desktop_config.json"
$cfgPath = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $cfgPath) { $cfgPath = $candidates[0]; if (-not $cfgPath) { $cfgPath = "$env:APPDATA\Claude\claude_desktop_config.json" } }
New-Item -ItemType Directory -Force (Split-Path $cfgPath) | Out-Null

# 5. Merge the "excel" entry into the existing config (keeps the app's own settings)
if (Test-Path $cfgPath) {
    Copy-Item $cfgPath "$cfgPath.bak" -Force
    $cfg = Get-Content $cfgPath -Raw -Encoding UTF8 | ConvertFrom-Json
} else {
    $cfg = New-Object psobject
}
if (-not $cfg.mcpServers) { $cfg | Add-Member -NotePropertyName mcpServers -NotePropertyValue (New-Object psobject) -Force }
$entry = [pscustomobject]@{ command = $uvx; args = @('excel-mcp-server', 'stdio', '--allow-dir', $dataDir) }
$cfg.mcpServers | Add-Member -NotePropertyName excel -NotePropertyValue $entry -Force

# Write UTF-8 without BOM
[IO.File]::WriteAllText($cfgPath, ($cfg | ConvertTo-Json -Depth 100), (New-Object System.Text.UTF8Encoding($false)))

Write-Host "Done. Excel MCP added to: $cfgPath" -ForegroundColor Green
Write-Host "  uvx:        $uvx"
Write-Host "  allow-dir:  $dataDir"
Write-Host "Now start Claude Desktop again."
