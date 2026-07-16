#!/usr/bin/env pwsh

# Log session start event

$ErrorActionPreference = 'Stop'

# Skip if logging disabled
if ($env:SKIP_LOGGING -ieq 'true') {
    exit 0
}

# Read input from Copilot when stdin is redirected
$inputPayload = ''
if ([Console]::IsInputRedirected) {
    $inputPayload = [Console]::In.ReadToEnd()
}

# Create logs directory if it doesn't exist
$logsDir = Join-Path -Path 'logs' -ChildPath 'copilot'
New-Item -ItemType Directory -Path $logsDir -Force | Out-Null

# Extract timestamp and session info
$timestamp = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
$cwd = (Get-Location).Path

# Log session start
$entry = @{
    timestamp = $timestamp
    event = 'sessionStart'
    cwd = $cwd
}
($entry | ConvertTo-Json -Compress) | Add-Content -Path (Join-Path $logsDir 'session.log')

Write-Output "Session logged"
exit 0