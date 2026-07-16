#!/usr/bin/env pwsh

# Log user prompt submission

$ErrorActionPreference = 'Stop'

# Skip if logging disabled
if ($env:SKIP_LOGGING -ieq 'true') {
    exit 0
}

# Read input from Copilot (contains prompt info) when stdin is redirected
$inputPayload = ''
if ([Console]::IsInputRedirected) {
    $inputPayload = [Console]::In.ReadToEnd()
}

# Create logs directory if it doesn't exist
$logsDir = Join-Path -Path 'logs' -ChildPath 'copilot'
New-Item -ItemType Directory -Path $logsDir -Force | Out-Null

# Extract timestamp
$timestamp = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
$level = if ([string]::IsNullOrWhiteSpace($env:LOG_LEVEL)) { 'INFO' } else { $env:LOG_LEVEL }

# Log prompt (you can parse inputPayload for more details)
$entry = @{
    timestamp = $timestamp
    event = 'userPromptSubmitted'
    level = $level
}
($entry | ConvertTo-Json -Compress) | Add-Content -Path (Join-Path $logsDir 'prompts.log')

exit 0