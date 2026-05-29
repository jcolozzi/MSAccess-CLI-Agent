# Private/Session.ps1 — COM session management helpers

function Get-RunningComApp {
    <#
    .SYNOPSIS
        Try to attach to an already-running COM application via the Running Object Table.
        Returns the COM object or $null.  Works on Windows PowerShell 5.1 (Desktop);
        gracefully degrades on PowerShell 7+ where [Marshal]::GetActiveObject is unavailable.
    #>
    param(
        [Parameter(Mandatory)][string]$ProgId,
        [Parameter(Mandatory)][string]$ProcessName
    )

    # Fast exit: if the host process isn't running, skip the COM probe entirely
    if (-not (Get-Process -Name $ProcessName -ErrorAction SilentlyContinue)) {
        Write-Verbose "Get-RunningComApp: no $ProcessName process found — skipping ROT lookup."
        return $null
    }

    try {
        $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject($ProgId)
        Write-Verbose "Get-RunningComApp: attached to existing $ProgId instance."
        return $app
    }
    catch [System.Management.Automation.MethodException] {
        # .NET Core / PS7 — GetActiveObject does not exist
        Write-Verbose "Get-RunningComApp: [Marshal]::GetActiveObject unavailable (PowerShell $($PSVersionTable.PSVersion)) — will create new instance."
        return $null
    }
    catch {
        # No ROT entry, or stale/dead entry
        Write-Verbose "Get-RunningComApp: could not attach to $ProgId — $($_.Exception.Message)"
        return $null
    }
}

function Test-AccessAlive {
    <#
    .SYNOPSIS
        Best-effort COM liveness check (does not depend on Visible).
    #>
    if ($null -eq $script:AccessSession.App) { return $false }
    $alive = $false
    try {
        $null = Get-AccessHwnd -App $script:AccessSession.App
        $alive = $true
    }
    catch {
        try {
            $null = $script:AccessSession.App.Version
            $alive = $true
        }
        catch {
            $alive = $false
        }
    }
    return $alive
}

function Get-AccessHwnd {
    <#
    .SYNOPSIS
        Get the Access window handle. Handles hWndAccessApp being a property or method.
    #>
    param($App)

    $h = $App.hWndAccessApp
    if ($h -is [System.Management.Automation.PSMethod]) {
        return [long]$h.Invoke(@())
    }
    return [long]$h
}

function Set-AccessVisibleBestEffort {
    <#
    .SYNOPSIS
        Try to set Access visibility. Never fail startup if unsupported.
    #>
    param([bool]$Visible = $true)
    if ($null -eq $script:AccessSession.App) { return }
    try {
        $script:AccessSession.App.Visible = $Visible
    } catch {
        Write-Verbose "Could not set Access.Visible=$Visible (continuing): $_"
    }
}

function Clear-AccessCaches {
    <#
    .SYNOPSIS
        Clear all VBE/control/CodeModule caches.
    #>
    $script:AccessSession.VbeCodeCache  = @{}
    $script:AccessSession.ControlsCache = @{}
    $script:AccessSession.CmCache       = @{}
}

function Resolve-SessionDbPath {
    <#
    .SYNOPSIS
        Resolve -DbPath: use explicit value if given, else fall back to active session.
        Throws a terminating error if neither is available.
    #>
    param(
        [string]$DbPath,
        [string]$CallerName = 'AccessPOSH'
    )
    if ($DbPath) { return $DbPath }
    if ($script:AccessSession.DbPath) { return $script:AccessSession.DbPath }
    throw "${CallerName}: -DbPath is required (no active session). Open a database first."
}

function Connect-AccessDB {
    <#
    .SYNOPSIS
        Internal: ensure Access COM is running and the requested DB is open.
        Returns the COM Application object.
        Tries to attach to an already-running Access instance (GetObject-first)
        before creating a new one, to prevent duplicate instances and file corruption.
    #>
    param(
        [string]$DbPath,
        [switch]$ForceNewInstance
    )
    if (-not $DbPath) { throw "Connect-AccessDB: -DbPath is required." }

    $resolved = [System.IO.Path]::GetFullPath($DbPath)

    # If we have an existing session, check liveness
    if ($null -ne $script:AccessSession.App) {
        if (-not (Test-AccessAlive)) {
            Write-Verbose 'COM session stale — auto-reconnecting...'
            $script:AccessSession.App     = $null
            $script:AccessSession.DbPath  = $null
            $script:AccessSession.OwnsApp = $false
            Clear-AccessCaches
        }
    }

    # Acquire Access instance if needed (GetObject-first, then New-Object)
    if ($null -eq $script:AccessSession.App) {
        $adopted = $false

        # Try to attach to an existing Access instance via the ROT
        if (-not $ForceNewInstance) {
            $existing = Get-RunningComApp -ProgId 'Access.Application' -ProcessName 'MSACCESS'
            if ($null -ne $existing) {
                # Access is single-DB-per-instance: only adopt if same DB is open
                $existingDb = $null
                try { $existingDb = $existing.CurrentProject.FullName } catch {}

                if ($existingDb -and ($existingDb -eq $resolved)) {
                    Write-Verbose "Adopting existing Access instance (same DB: $resolved)"
                    $script:AccessSession.App     = $existing
                    $script:AccessSession.OwnsApp = $false
                    $script:AccessSession.DbPath  = $resolved
                    $adopted = $true
                    # Suppress dialogs on the adopted instance
                    try {
                        $script:AccessSession.App.DisplayAlerts = $false
                        $script:AccessSession.App.AutomationSecurity = 1
                    } catch {}
                    Set-AccessVisibleBestEffort -Visible $true
                    Clear-AccessCaches
                    Write-Verbose 'Adopted existing Access instance OK'
                } else {
                    Write-Verbose "Existing Access instance has different DB ($existingDb) — creating new instance."
                }
            }
        }

        # Fall back to creating a new instance
        if (-not $adopted) {
            Write-Verbose 'Launching Access.Application...'
            try {
                $script:AccessSession.App = New-Object -ComObject 'Access.Application'
            } catch {
                throw "Failed to create Access.Application COM object. Is Microsoft Access installed? Error: $_"
            }
            $script:AccessSession.OwnsApp = $true
            # Suppress dialogs for non-interactive automation
            try {
                $script:AccessSession.App.DisplayAlerts = $false
                $script:AccessSession.App.AutomationSecurity = 1  # msoAutomationSecurityForceDisable
            } catch {}
            Set-AccessVisibleBestEffort -Visible $true
            Write-Verbose 'Access launched OK'
        }
    }

    # Switch database if needed (skip if we just adopted with the correct DB)
    if ($script:AccessSession.DbPath -ne $resolved) {
        if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) {
            throw "Database file not found: $resolved"
        }

        # Close previous database
        if ($null -ne $script:AccessSession.DbPath) {
            Write-Verbose "Closing previous DB: $($script:AccessSession.DbPath)"
            try {
                $script:AccessSession.App.CloseCurrentDatabase()
            } catch {
                Write-Verbose "Error closing previous DB: $_"
            }
        }

        # Open new database
        Write-Verbose "Opening DB: $resolved"
        try {
            $script:AccessSession.App.OpenCurrentDatabase($resolved)
        } catch {
            if ($_.Exception.Message -match 'already have the database open') {
                Write-Verbose 'DB was already open — syncing state'
            } else {
                throw
            }
        }

        $script:AccessSession.DbPath = $resolved
        Set-AccessVisibleBestEffort -Visible $true
        Clear-AccessCaches
        Write-Verbose 'DB opened OK'
    }

    return $script:AccessSession.App
}

function Get-AccessOfficeVersion {
    <#
    .SYNOPSIS
        Best-effort detection of installed Office version and MSACCESS.EXE path.
    .DESCRIPTION
        Populates $script:AccessSession.OfficeVersion and .MsAccessPath.
        Never throws. Defaults to "16.0" / $null if nothing found.
    #>
    [CmdletBinding()]
    param()

    # Idempotent — skip if already detected
    if ($script:AccessSession.OfficeVersion) {
        Write-Verbose "Office version already detected: $($script:AccessSession.OfficeVersion)"
        return
    }

    # Defaults
    $script:AccessSession.OfficeVersion = '16.0'
    $script:AccessSession.MsAccessPath  = $null

    try {
        $bestVersion = $null
        $bestPath    = $null

        # Registry roots to probe: HKLM 64-bit, HKLM WOW6432 (32-bit on 64-bit OS), HKCU (per-user C2R)
        $probes = @(
            @{ Hive = [Microsoft.Win32.RegistryHive]::LocalMachine;  View = [Microsoft.Win32.RegistryView]::Registry64; Prefix = 'Software\Microsoft\Office' },
            @{ Hive = [Microsoft.Win32.RegistryHive]::LocalMachine;  View = [Microsoft.Win32.RegistryView]::Registry32; Prefix = 'Software\Microsoft\Office' },
            @{ Hive = [Microsoft.Win32.RegistryHive]::CurrentUser;   View = [Microsoft.Win32.RegistryView]::Default;    Prefix = 'Software\Microsoft\Office' }
        )

        foreach ($probe in $probes) {
            try {
                $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey($probe.Hive, $probe.View)
                $officeKey = $baseKey.OpenSubKey($probe.Prefix)
                if (-not $officeKey) {
                    $baseKey.Dispose()
                    continue
                }

                foreach ($subName in $officeKey.GetSubKeyNames()) {
                    # Match version-like subkeys (e.g. "16.0", "15.0")
                    if ($subName -notmatch '^\d+\.\d+$') { continue }

                    try {
                        $installRoot = $officeKey.OpenSubKey("$subName\Access\InstallRoot")
                        if (-not $installRoot) { continue }

                        $pathVal = $installRoot.GetValue('Path')
                        $installRoot.Dispose()

                        if (-not $pathVal) { continue }

                        $exePath = Join-Path $pathVal 'MSACCESS.EXE'
                        if (-not (Test-Path -LiteralPath $exePath -PathType Leaf)) { continue }

                        # Pick highest version
                        $ver = [version]$subName
                        if ((-not $bestVersion) -or ($ver -gt $bestVersion)) {
                            $bestVersion = $ver
                            $bestPath    = $exePath
                            Write-Verbose "Found Access $subName at $exePath"
                        }
                    }
                    catch {
                        # Skip individual version subkey errors
                    }
                }

                $officeKey.Dispose()
                $baseKey.Dispose()
            }
            catch {
                # Skip probe errors
            }
        }

        # If found via InstallRoot, apply and return
        if ($bestVersion) {
            $script:AccessSession.OfficeVersion = "$($bestVersion.Major).$($bestVersion.Minor)"
            $script:AccessSession.MsAccessPath  = $bestPath
            Write-Verbose "Detected Office version $($script:AccessSession.OfficeVersion)"
            return
        }

        # Fallback: App Paths
        Write-Verbose 'InstallRoot not found — trying App Paths fallback'
        try {
            $baseKey    = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
                [Microsoft.Win32.RegistryHive]::LocalMachine,
                [Microsoft.Win32.RegistryView]::Registry64
            )
            $appPathKey = $baseKey.OpenSubKey('Software\Microsoft\Windows\CurrentVersion\App Paths\MSACCESS.EXE')
            if ($appPathKey) {
                $exePath = $appPathKey.GetValue('')  # (Default) value
                $appPathKey.Dispose()

                if ($exePath -and (Test-Path -LiteralPath $exePath -PathType Leaf)) {
                    $script:AccessSession.MsAccessPath = $exePath

                    # Try to parse OfficeNN from path (e.g. "...\Office16\MSACCESS.EXE")
                    if ($exePath -match '\\Office(\d+)\\') {
                        $script:AccessSession.OfficeVersion = "$($Matches[1]).0"
                    }
                    Write-Verbose "App Paths fallback: $exePath (version $($script:AccessSession.OfficeVersion))"
                }
            }
            $baseKey.Dispose()
        }
        catch {
            Write-Warning "App Paths fallback failed: $_"
        }

        Write-Verbose "Final Office version: $($script:AccessSession.OfficeVersion), path: $($script:AccessSession.MsAccessPath)"
    }
    catch {
        Write-Warning "Office version detection failed: $_"
    }
}
