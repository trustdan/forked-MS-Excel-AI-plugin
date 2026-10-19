<#
.SYNOPSIS
  Hardens Excel AI Plugin by creating an isolated offline Node binary with Windows Defender Firewall rules.

.DESCRIPTION
  1. Copies the system node.exe to ./bin/node-offline.exe.
  2. Ensures bin/ is added to .gitignore.
  3. Configures a Windows Defender Firewall Outbound Rule to block all remote IP addresses
     (0.0.0.0-126.255.255.255 and 128.0.0.0-255.255.255.255) for node-offline.exe.
  4. Keeps loopback (127.0.0.0/8) open so the sidecar can communicate with LM Studio and Excel.

.NOTES
  Must be executed in an elevated PowerShell session (Run as Administrator).
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

# 1. Elevate / Check Administrator Rights
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Error "This script requires Administrator privileges. Please re-run in an elevated PowerShell prompt ('Run as Administrator')."
    exit 1
}

# 2. Determine Project Root and bin directory
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptDir
$BinDir = Join-Path $ProjectRoot "bin"
$NodeOfflineExe = Join-Path $BinDir "node-offline.exe"
$GitIgnorePath = Join-Path $ProjectRoot ".gitignore"

Write-Host "==> Excel AI Offline Lockdown Setup" -ForegroundColor Cyan
Write-Host "Project Root: $ProjectRoot"

# 3. Locate system node.exe and copy to ./bin/node-offline.exe
if (-not (Test-Path $BinDir)) {
    New-Item -ItemType Directory -Path $BinDir -Force | Out-Null
    Write-Host "Created directory: $BinDir" -ForegroundColor Green
}

$SystemNode = (Get-Command node.exe -ErrorAction Stop).Source
Write-Host "System node.exe located at: $SystemNode"
Copy-Item -Path $SystemNode -Destination $NodeOfflineExe -Force
Write-Host "Copied to: $NodeOfflineExe" -ForegroundColor Green

# 4. Ensure bin/ is ignored in git
if (Test-Path $GitIgnorePath) {
    $gitIgnoreContent = Get-Content $GitIgnorePath -Raw
    if ($gitIgnoreContent -notmatch "(?m)^bin/") {
        Add-Content -Path $GitIgnorePath -Value "`nbin/"
        Write-Host "Added 'bin/' to .gitignore" -ForegroundColor Green
    }
}

# 5. Create or Replace Windows Firewall Rule
$RuleName = "ExcelAI-Node-Offline-Block"

# Check if old rule exists and remove it
$existingRule = Get-NetFirewallRule -DisplayName $RuleName -ErrorAction SilentlyContinue
if ($existingRule) {
    Write-Host "Removing existing firewall rule '$RuleName'..." -ForegroundColor Yellow
    Remove-NetFirewallRule -DisplayName $RuleName
}

# Block all non-loopback addresses: 0.0.0.0-126.255.255.255 and 128.0.0.0-255.255.255.255
$BlockedRanges = @(
    "0.0.0.0-126.255.255.255",
    "128.0.0.0-255.255.255.255"
)

Write-Host "Configuring outbound Windows Defender Firewall rule for node-offline.exe..." -ForegroundColor Cyan

New-NetFirewallRule `
    -DisplayName $RuleName `
    -Description "Blocks outbound network access outside localhost for Excel AI offline Node runner" `
    -Direction Outbound `
    -Program $NodeOfflineExe `
    -Action Block `
    -RemoteAddress $BlockedRanges `
    -Enabled True | Out-Null

Write-Host "`n✓ Firewall rule '$RuleName' successfully established." -ForegroundColor Green
Write-Host "✓ Outbound traffic from '$NodeOfflineExe' is blocked for all non-loopback IPs." -ForegroundColor Green
Write-Host "✓ Localhost (127.0.0.0/8) remains accessible for LM Studio and Excel taskpane." -ForegroundColor Green
Write-Host "`nYou can now launch the offline sidecar using:" -ForegroundColor Yellow
Write-Host "  pnpm start:offline`n" -ForegroundColor Yellow
