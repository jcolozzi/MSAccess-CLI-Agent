---
name: 'Session Auto-Commit'
description: 'Automatically commits and pushes changes when a Copilot coding agent session ends'
tags: ['automation', 'git', 'productivity']
---

# Session Auto-Commit Hook

Automatically commits and pushes changes when a GitHub Copilot coding agent session ends, ensuring your work is always saved and backed up.

## Overview

This hook runs at the end of each Copilot coding agent session and automatically:

- Detects if there are uncommitted changes
- Stages all changes
- Creates a timestamped commit
- Pushes to the remote repository

## Features

- **Automatic Backup**: Never lose work from a Copilot session
- **Timestamped Commits**: Each auto-commit includes the session end time
- **Safe Execution**: Only commits when there are actual changes
- **Error Handling**: Gracefully handles push failures
- **Windows Native**: Uses PowerShell without requiring Bash/WSL

## Installation

1. Copy this hook folder to your repository's `.github/hooks/` directory.

1. On Windows PowerShell, allow script execution for the current process if needed:

  ```powershell
  Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
  ```

1. Commit the hook configuration to your repository's default branch.

## Configuration

The hook is configured in `hooks.json` to run on the `sessionEnd` event:

```json
{
  "version": 1,
  "hooks": {
    "sessionEnd": [
      {
        "type": "command",
        "powershell": ".github/hooks/session-auto-commit/auto-commit.ps1",
        "timeoutSec": 30
      }
    ]
  }
}
```

## How It Works

1. When a Copilot coding agent session ends, the hook executes
2. Checks if inside a Git repository
3. Detects uncommitted changes using `git status`
4. Stages all changes with `git add -A`
5. Creates a commit with format: `auto-commit: YYYY-MM-DD HH:MM:SS`
6. Attempts to push to remote
7. Reports success or failure

## Customization

You can customize the hook by modifying `auto-commit.ps1`:

- **Commit Message Format**: Change the timestamp format or message prefix
- **Selective Staging**: Use specific git add patterns instead of `-A`
- **Branch Selection**: Push to specific branches only
- **Notifications**: Add desktop notifications or Slack messages

## Disabling

To temporarily disable auto-commits:

1. Remove or comment out the `sessionEnd` hook in `hooks.json`
1. Or set an environment variable:

  ```powershell
  $env:SKIP_AUTO_COMMIT = "true"
  ```

## Notes

- The hook uses `--no-verify` to avoid triggering pre-commit hooks
- Failed pushes won't block session termination
- Requires appropriate git credentials configured
- Works with both Copilot coding agent and GitHub Copilot CLI
