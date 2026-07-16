#!/usr/bin/env pwsh

# Session Auto-Commit Hook (PowerShell)
# Automatically commits and pushes changes when a Copilot session ends

$ErrorActionPreference = 'Stop'

# Check if SKIP_AUTO_COMMIT is set
if ($env:SKIP_AUTO_COMMIT -ieq 'true') {
    Write-Output "Auto-commit skipped (SKIP_AUTO_COMMIT=true)"
    exit 0
}

# Check if we're in a git repository
git rev-parse --is-inside-work-tree *> $null
if ($LASTEXITCODE -ne 0) {
    Write-Output "Not in a git repository"
    exit 0
}

# Check for uncommitted changes
$status = git status --porcelain
if ($LASTEXITCODE -ne 0) {
    Write-Output "Unable to read repository status"
    exit 0
}

if ([string]::IsNullOrWhiteSpace(($status -join "`n"))) {
    Write-Output "No changes to commit"
    exit 0
}

Write-Output "Auto-committing changes from Copilot session..."

# Stage all changes
git add -A
if ($LASTEXITCODE -ne 0) {
    Write-Output "Staging changes failed"
    exit 0
}

# Create timestamped commit
$timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
git commit -m "auto-commit: $timestamp" --no-verify 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Output "Commit failed"
    exit 0
}

# Attempt to push
git push 2>$null
if ($LASTEXITCODE -eq 0) {
    Write-Output "Changes committed and pushed successfully"
}
else {
    Write-Output "Push failed - changes committed locally"
}

exit 0