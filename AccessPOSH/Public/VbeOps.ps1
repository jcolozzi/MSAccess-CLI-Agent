# Public/VbeOps.ps1 — VBE/VBA operations: read, write, search, compile, execute

function Get-AccessVbeLine {
    <#
    .SYNOPSIS
        Read a range of lines from a VBE CodeModule.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        Type: module, form, or report.
    .PARAMETER ObjectName
        Name of the module/form/report.
    .PARAMETER StartLine
        First line to read (1-based).
    .PARAMETER Count
        Number of lines to read.
    .EXAMPLE
        Get-AccessVbeLine -DbPath "C:\db.accdb" -ObjectType module -ObjectName "Module1" -StartLine 1 -Count 10
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [ValidateNotNullOrEmpty()]
        [ValidateSet('module','form','report')][string]$ObjectType,
        [ValidateNotNullOrEmpty()]
        [string]$ObjectName,
        [int]$StartLine,
        [int]$Count
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Get-AccessVbeLine'
    if (-not $ObjectType) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-ObjectType is required (module, form, report).'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $ObjectType
            )
        )
    }
    if (-not $ObjectName) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-ObjectName is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $ObjectName
            )
        )
    }
    if (-not $PSBoundParameters.ContainsKey('StartLine')) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-StartLine is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $PSBoundParameters
            )
        )
    }
    if (-not $PSBoundParameters.ContainsKey('Count')) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-Count is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $PSBoundParameters
            )
        )
    }
    $app = Connect-AccessDB -DbPath $DbPath
    $cm = Get-CodeModule -App $app -ObjectType $ObjectType -ObjectName $ObjectName
    $cacheKey = "${ObjectType}:${ObjectName}"
    $allCode = Get-AllModuleCode -CodeModule $cm -CacheKey $cacheKey
    $allLines = $allCode.Split("`n")
    $total = $allLines.Count

    if ($StartLine -lt 1 -or $StartLine -gt $total) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentOutOfRangeException]::new('StartLine', $StartLine, "start_line $StartLine out of range (1-$total)."),
                'OutOfRange',
                [System.Management.Automation.ErrorCategory]::LimitsExceeded,
                $StartLine
            )
        )
    }
    $actual = [math]::Min($Count, $total - $StartLine + 1)
    $result = $allLines[($StartLine - 1) .. ($StartLine - 1 + $actual - 1)] -join "`n"
    return $result.TrimEnd("`r")
}

function Get-AccessVbeProc {
    <#
    .SYNOPSIS
        Extract a procedure by name from a VBE module.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        Type: module, form, or report.
    .PARAMETER ObjectName
        Name of the module/form/report.
    .PARAMETER ProcName
        Name of the procedure (Sub/Function/Property).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-AccessVbeProc -DbPath "C:\db.accdb" -ObjectType module -ObjectName "Module1" -ProcName "MyFunc"
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [ValidateNotNullOrEmpty()]
        [ValidateSet('module','form','report')][string]$ObjectType,
        [ValidateNotNullOrEmpty()]
        [string]$ObjectName,
        [string]$ProcName,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Get-AccessVbeProc'
    if (-not $ObjectType) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-ObjectType is required (module, form, report).'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $ObjectType
            )
        )
    }
    if (-not $ObjectName) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-ObjectName is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $ObjectName
            )
        )
    }
    if (-not $ProcName) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-ProcName is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $ProcName
            )
        )
    }
    $app = Connect-AccessDB -DbPath $DbPath
    $cm = Get-CodeModule -App $app -ObjectType $ObjectType -ObjectName $ObjectName

    try {
        $start = $cm.ProcStartLine($ProcName, 0)  # 0 = vbext_pk_Proc
        $body  = $cm.ProcBodyLine($ProcName, 0)
        $count = $cm.ProcCountLines($ProcName, 0)
    } catch {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.IO.FileNotFoundException]::new("Procedure '$ProcName' not found in '$ObjectName': $_"),
                'ObjectNotFound',
                [System.Management.Automation.ErrorCategory]::ObjectNotFound,
                $ProcName
            )
        )
    }

    $cacheKey = "${ObjectType}:${ObjectName}"
    $allLines = (Get-AllModuleCode -CodeModule $cm -CacheKey $cacheKey).Split("`n")
    $total = $allLines.Count
    $count = [math]::Min($count, $total - $start + 1)
    $code = ($allLines[($start - 1) .. ($start - 1 + $count - 1)] -join "`n").TrimEnd("`r")

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        proc_name  = $ProcName
        start_line = $start
        body_line  = $body
        count      = $count
        code       = $code
    })
}

function Get-AccessVbeModuleInfo {
    <#
    .SYNOPSIS
        Enumerate all procedures in a VBE module with their positions.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        Type: module, form, or report.
    .PARAMETER ObjectName
        Name of the module/form/report.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-AccessVbeModuleInfo -DbPath "C:\db.accdb" -ObjectType module -ObjectName "Module1"
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [ValidateNotNullOrEmpty()]
        [ValidateSet('module','form','report')][string]$ObjectType,
        [ValidateNotNullOrEmpty()]
        [string]$ObjectName,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Get-AccessVbeModuleInfo'
    if (-not $ObjectType) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-ObjectType is required (module, form, report).'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $ObjectType
            )
        )
    }
    if (-not $ObjectName) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-ObjectName is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $ObjectName
            )
        )
    }

    $app = Connect-AccessDB -DbPath $DbPath
    $cm = Get-CodeModule -App $app -ObjectType $ObjectType -ObjectName $ObjectName
    $cacheKey = "${ObjectType}:${ObjectName}"
    $allCode = Get-AllModuleCode -CodeModule $cm -CacheKey $cacheKey
    $allLines = $allCode.Split("`n")
    $total = $allLines.Count

    $procs = [System.Collections.Generic.List[object]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    for ($i = 0; $i -lt $total; $i++) {
        $line = $allLines[$i].Trim()
        if ($line -match '^(?:Public\s+|Private\s+|Friend\s+)?(?:Function|Sub|Property\s+(?:Get|Let|Set))\s+(\w+)') {
            $pname = $Matches[1]
            if (-not $seen.Add($pname)) { continue }
            try {
                $pstart = $cm.ProcStartLine($pname, 0)
                $pbody  = $cm.ProcBodyLine($pname, 0)
                $pcount = $cm.ProcCountLines($pname, 0)
                $pcount = [math]::Min($pcount, $total - $pstart + 1)
                $procs.Add([PSCustomObject][ordered]@{
                    name       = $pname
                    start_line = $pstart
                    body_line  = $pbody
                    count      = $pcount
                })
            } catch {
                $procs.Add([PSCustomObject][ordered]@{
                    name       = $pname
                    start_line = ($i + 1)
                })
            }
        }
    }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        total_lines = $total
        procs       = @($procs)
    })
}

function Set-AccessVbeLine {
    <#
    .SYNOPSIS
        Replace lines in a VBE module. count=0 inserts without deleting. Empty NewCode deletes only.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        Type: module, form, or report.
    .PARAMETER ObjectName
        Name of the module/form/report.
    .PARAMETER StartLine
        First line to replace (1-based).
    .PARAMETER Count
        Number of lines to delete before inserting.
    .PARAMETER NewCode
        Code to insert at StartLine (can be multiline).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-AccessVbeLine -DbPath "C:\db.accdb" -ObjectType module -ObjectName "Module1" -StartLine 5 -Count 3 -NewCode "' replaced lines"
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [ValidateNotNullOrEmpty()]
        [ValidateSet('module','form','report')][string]$ObjectType,
        [ValidateNotNullOrEmpty()]
        [string]$ObjectName,
        [int]$StartLine,
        [int]$Count = 0,
        [string]$NewCode = '',
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Set-AccessVbeLine'
    if (-not $ObjectType) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-ObjectType is required (module, form, report).'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $ObjectType
            )
        )
    }
    if (-not $ObjectName) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-ObjectName is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $ObjectName
            )
        )
    }
    if (-not $PSBoundParameters.ContainsKey('StartLine')) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-StartLine is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $PSBoundParameters
            )
        )
    }

    $app = Connect-AccessDB -DbPath $DbPath
    $cm = Get-CodeModule -App $app -ObjectType $ObjectType -ObjectName $ObjectName
    $total = $cm.CountOfLines

    if ($StartLine -lt 1 -or $StartLine -gt ($total + 1)) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentOutOfRangeException]::new('StartLine', $StartLine, "start_line $StartLine out of range (1-$total)."),
                'OutOfRange',
                [System.Management.Automation.ErrorCategory]::LimitsExceeded,
                $StartLine
            )
        )
    }

    $clamped = $false
    if ($Count -gt 0) {
        $maxCount = $total - $StartLine + 1
        if ($Count -gt $maxCount) { $Count = $maxCount; $clamped = $true }
        $cm.DeleteLines($StartLine, $Count)
    }

    $inserted = 0
    if ($NewCode) {
        $normalized = $NewCode -replace "`r`n", "`n" -replace "`r", "`n" -replace "`n", "`r`n"
        $cm.InsertLines($StartLine, $normalized)
        $inserted = $NewCode.Split("`n").Count
    }

    # Invalidate cache
    $cacheKey = "${ObjectType}:${ObjectName}"
    $script:AccessSession.VbeCodeCache.Remove($cacheKey)

    $newTotal = $cm.CountOfLines
    $clampNote = if ($clamped) { ' (count clamped to module boundary)' } else { '' }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        status        = "Lines $StartLine replaced ($Count deleted, $inserted inserted)$clampNote -> module now has $newTotal lines"
        deleted       = $Count
        inserted      = $inserted
        new_total     = $newTotal
    })
}

function Set-AccessVbeProc {
    <#
    .SYNOPSIS
        Replace an entire procedure by name. Auto-locates via ProcStartLine/ProcCountLines.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        Type: module, form, or report.
    .PARAMETER ObjectName
        Name of the module/form/report.
    .PARAMETER ProcName
        Name of the procedure to replace.
    .PARAMETER NewCode
        New code for the procedure. Empty string deletes it.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-AccessVbeProc -DbPath "C:\db.accdb" -ObjectType module -ObjectName "Module1" -ProcName "OldFunc" -NewCode "Public Sub OldFunc()`r`n  MsgBox ""Hello""`r`nEnd Sub"
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [ValidateNotNullOrEmpty()]
        [ValidateSet('module','form','report')][string]$ObjectType,
        [ValidateNotNullOrEmpty()]
        [string]$ObjectName,
        [string]$ProcName,
        [string]$NewCode = '',
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Set-AccessVbeProc'
    if (-not $ObjectType) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-ObjectType is required (module, form, report).'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $ObjectType
            )
        )
    }
    if (-not $ObjectName) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-ObjectName is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $ObjectName
            )
        )
    }
    if (-not $ProcName) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-ProcName is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $ProcName
            )
        )
    }

    $app = Connect-AccessDB -DbPath $DbPath

    # Close form/report if open in design (avoids COM conflicts with VBE)
    if ($ObjectType -in 'form', 'report') {
        $acObjType = if ($ObjectType -eq 'form') { $script:AC_FORM } else { $script:AC_REPORT }
        try { $app.DoCmd.Close($acObjType, $ObjectName, $script:AC_SAVE_YES) } catch {}
    }

    # Invalidate cache in case CodeModule is stale
    $cacheKey = "${ObjectType}:${ObjectName}"
    $script:AccessSession.CmCache.Remove($cacheKey)

    $cm = Get-CodeModule -App $app -ObjectType $ObjectType -ObjectName $ObjectName
    try {
        $start = $cm.ProcStartLine($ProcName, 0)
        $count = $cm.ProcCountLines($ProcName, 0)
    } catch {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.IO.FileNotFoundException]::new("Procedure '$ProcName' not found in '$ObjectName': $_"),
                'ObjectNotFound',
                [System.Management.Automation.ErrorCategory]::ObjectNotFound,
                $ProcName
            )
        )
    }

    $total = $cm.CountOfLines
    $count = [math]::Min($count, $total - $start + 1)

    $cm.DeleteLines($start, $count)

    $inserted = 0
    if ($NewCode) {
        $normalized = $NewCode -replace "`r`n", "`n" -replace "`r", "`n" -replace "`n", "`r`n"
        $cm.InsertLines($start, $normalized)
        $inserted = $NewCode.Split("`n").Count
    }

    # Invalidate caches
    $script:AccessSession.VbeCodeCache.Remove($cacheKey)
    $script:AccessSession.CmCache.Remove($cacheKey)

    $newTotal = $cm.CountOfLines
    $action = if ($NewCode) { 'replaced' } else { 'deleted' }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        status   = "Proc '$ProcName' $action ($count deleted, $inserted inserted) -> module now has $newTotal lines"
        action   = $action
        proc     = $ProcName
        deleted  = $count
        inserted = $inserted
        new_total = $newTotal
    })
}

function Update-AccessVbeProc {
    <#
    .SYNOPSIS
        Surgically patch code within a procedure using find/replace with whitespace tolerance.
    .DESCRIPTION
        Applies one or more patches (find/replace pairs) to a procedure without rewriting the entire proc.
        Three-layer matching: (1) exact string, (2) whitespace-normalized fallback, (3) context error reporting.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        Type: module, form, or report.
    .PARAMETER ObjectName
        Name of the module/form/report.
    .PARAMETER ProcName
        Name of the procedure to patch.
    .PARAMETER Patches
        Array of hashtables, each with 'find' and 'replace' keys.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Update-AccessVbeProc -DbPath "C:\db.accdb" -ObjectType module -ObjectName "Module1" `
            -ProcName "MyFunc" -Patches @(@{find='MsgBox "Old"'; replace='MsgBox "New"'})
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [ValidateNotNullOrEmpty()]
        [ValidateSet('module','form','report')][string]$ObjectType,
        [ValidateNotNullOrEmpty()]
        [string]$ObjectName,
        [string]$ProcName,
        [array]$Patches,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Update-AccessVbeProc'
    if (-not $ObjectType) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-ObjectType is required (module, form, report).'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $ObjectType
            )
        )
    }
    if (-not $ObjectName) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-ObjectName is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $ObjectName
            )
        )
    }
    if (-not $ProcName) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-ProcName is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $ProcName
            )
        )
    }
    if (-not $Patches) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-Patches is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $Patches
            )
        )
    }

    $app = Connect-AccessDB -DbPath $DbPath

    # Close form/report if open in design (avoids COM conflicts with VBE)
    if ($ObjectType -in 'form', 'report') {
        $acObjType = if ($ObjectType -eq 'form') { $script:AC_FORM } else { $script:AC_REPORT }
        try { $app.DoCmd.Close($acObjType, $ObjectName, $script:AC_SAVE_YES) } catch {}
    }

    # Invalidate cache
    $cacheKey = "${ObjectType}:${ObjectName}"
    $script:AccessSession.CmCache.Remove($cacheKey)

    $cm = Get-CodeModule -App $app -ObjectType $ObjectType -ObjectName $ObjectName

    # Locate procedure — try kind=0 (Sub/Function), fallback to kind=3 (Property)
    $kind = 0
    try {
        $start = $cm.ProcStartLine($ProcName, 0)
        $count = $cm.ProcCountLines($ProcName, 0)
    } catch {
        try {
            $start = $cm.ProcStartLine($ProcName, 3)
            $count = $cm.ProcCountLines($ProcName, 3)
            $kind = 3
        } catch {
            $PSCmdlet.ThrowTerminatingError(
                [System.Management.Automation.ErrorRecord]::new(
                    [System.IO.FileNotFoundException]::new("Procedure '$ProcName' not found in '$ObjectName': $_"),
                    'ObjectNotFound',
                    [System.Management.Automation.ErrorCategory]::ObjectNotFound,
                    $ProcName
                )
            )
        }
    }

    $total = $cm.CountOfLines
    $count = [math]::Min($count, $total - $start + 1)

    # Get current proc code
    $procCode = $cm.Lines($start, $count)
    $backupCode = $procCode

    # Apply patches sequentially
    $applied = 0
    $notFound = @()
    $wsFallbackNotes = @()

    for ($pi = 0; $pi -lt $Patches.Count; $pi++) {
        $patch = $Patches[$pi]
        $findText    = "$($patch['find'])"
        $replaceText = if ($patch.ContainsKey('replace')) { "$($patch['replace'])" } else { '' }

        # Layer 1: Exact match
        if ($procCode.Contains($findText)) {
            $idx = $procCode.IndexOf($findText)
            $procCode = $procCode.Substring(0, $idx) + $replaceText + $procCode.Substring($idx + $findText.Length)
            $applied++
        }
        else {
            # Layer 2: Whitespace-normalized fallback
            $wsMatch = Test-WsNormalizedMatch -ProcCode $procCode -FindText $findText
            if ($null -ne $wsMatch) {
                $codeLines = @($procCode -split "`r?`n")
                $replaceNorm = $replaceText
                if ($replaceNorm -and -not $replaceNorm.EndsWith("`n")) {
                    $replaceNorm += "`r`n"
                }
                $before = if ($wsMatch.start -gt 0) { $codeLines[0..($wsMatch.start - 1)] } else { @() }
                $after  = if ($wsMatch.end -lt ($codeLines.Count - 1)) { $codeLines[($wsMatch.end + 1)..($codeLines.Count - 1)] } else { @() }
                $middle = if ($replaceNorm) { @($replaceNorm.TrimEnd("`r`n")) } else { @() }
                $procCode = ($before + $middle + $after) -join "`r`n"
                $applied++
                $wsFallbackNotes += "patch[$pi]: matched via ws-normalized fallback"
            }
            else {
                # Layer 3: Context error reporting
                $ctx = Get-ClosestMatchContext -ProcCode $procCode -FindText $findText -ProcName $ProcName
                $notFound += "patch[$pi]: not found. $ctx"
            }
        }
    }

    if ($applied -eq 0) {
        $errMsg = "NOOP: no patches matched in '$ProcName'. Errors:`n$($notFound -join "`n")"
        return Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
            status     = $errMsg
            applied    = 0
            total      = $Patches.Count
            not_found  = $notFound
        })
    }

    # Strip Option lines if proc is NOT at the top of the module
    $optionWarnings = @()
    if ($start -gt 5) {
        $optionRe = '^\s*Option\s+(Explicit|Compare\s+\w+)\s*$'
        $cleanLines = @()
        foreach ($line in ($procCode -split "`r?`n")) {
            if ($line -match $optionRe) {
                $optionWarnings += "Stripped misplaced Option line: '$($line.Trim())'"
            } else {
                $cleanLines += $line
            }
        }
        $procCode = $cleanLines -join "`r`n"
    }

    # Replace entire proc with patched code
    try {
        $cm.DeleteLines($start, $count)
        if ($procCode.Trim()) {
            $normalized = $procCode -replace "`r`n", "`n" -replace "`r", "`n" -replace "`n", "`r`n"
            $cm.InsertLines($start, $normalized)
        }
    } catch {
        # Rollback: restore backup
        try { $cm.InsertLines($start, $backupCode) } catch {}
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.InvalidOperationException]::new("Patch failed (rolled back): $_"),
                'InvalidOperation',
                [System.Management.Automation.ErrorCategory]::InvalidOperation,
                $ProcName
            )
        )
    }

    # Invalidate caches
    $script:AccessSession.VbeCodeCache.Remove($cacheKey)
    $script:AccessSession.CmCache.Remove($cacheKey)

    $newTotal = $cm.CountOfLines
    $newCount = try { $cm.ProcCountLines($ProcName, $kind) } catch { 0 }

    # Build result
    $resultParts = @("OK: $applied/$($Patches.Count) patches applied in '$ProcName' ($count -> $newCount lines) -> module now has $newTotal lines")
    if ($wsFallbackNotes.Count -gt 0) { $resultParts += "WS-fallback: $($wsFallbackNotes -join '; ')" }
    if ($optionWarnings.Count -gt 0)  { $resultParts += $optionWarnings }
    if ($notFound.Count -gt 0)        { $resultParts += "Not found:"; $resultParts += $notFound }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        status          = $resultParts -join "`n"
        applied         = $applied
        total           = $Patches.Count
        proc            = $ProcName
        old_lines       = $count
        new_lines       = $newCount
        module_lines    = $newTotal
        ws_fallback     = $wsFallbackNotes
        not_found       = $notFound
    })
}

function Add-AccessVbeCode {
    <#
    .SYNOPSIS
        Append code at the end of a VBE module.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        Type: module, form, or report.
    .PARAMETER ObjectName
        Name of the module/form/report.
    .PARAMETER Code
        Code to append (can be multiline).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Add-AccessVbeCode -DbPath "C:\db.accdb" -ObjectType module -ObjectName "Module1" -Code "Public Sub NewSub()`r`nEnd Sub"
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [ValidateNotNullOrEmpty()]
        [ValidateSet('module','form','report')][string]$ObjectType,
        [ValidateNotNullOrEmpty()]
        [string]$ObjectName,
        [string]$Code,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Add-AccessVbeCode'
    if (-not $ObjectType) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-ObjectType is required (module, form, report).'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $ObjectType
            )
        )
    }
    if (-not $ObjectName) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-ObjectName is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $ObjectName
            )
        )
    }
    if (-not $Code) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-Code is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $Code
            )
        )
    }

    $app = Connect-AccessDB -DbPath $DbPath
    $cm = Get-CodeModule -App $app -ObjectType $ObjectType -ObjectName $ObjectName
    $total = $cm.CountOfLines

    $normalized = $Code -replace "`r`n", "`n" -replace "`r", "`n" -replace "`n", "`r`n"
    $cm.InsertLines($total + 1, $normalized)
    $inserted = $Code.Split("`n").Count

    $cacheKey = "${ObjectType}:${ObjectName}"
    $script:AccessSession.VbeCodeCache.Remove($cacheKey)

    $newTotal = $cm.CountOfLines

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        status    = "$inserted lines appended -> module now has $newTotal lines"
        inserted  = $inserted
        new_total = $newTotal
    })
}

function Find-AccessVbeText {
    <#
    .SYNOPSIS
        Search for text (or regex) in a VBE module.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER ObjectType
        Type: module, form, or report.
    .PARAMETER ObjectName
        Name of the module/form/report.
    .PARAMETER SearchText
        Text or regex pattern to find.
    .PARAMETER MatchCase
        Case-sensitive matching.
    .PARAMETER UseRegex
        Treat SearchText as a regex pattern.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Find-AccessVbeText -DbPath "C:\db.accdb" -ObjectType module -ObjectName "Module1" -SearchText "MsgBox"
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [ValidateNotNullOrEmpty()]
        [ValidateSet('module','form','report')][string]$ObjectType,
        [ValidateNotNullOrEmpty()]
        [string]$ObjectName,
        [string]$SearchText,
        [switch]$MatchCase,
        [switch]$UseRegex,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Find-AccessVbeText'
    if (-not $ObjectType) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-ObjectType is required (module, form, report).'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $ObjectType
            )
        )
    }
    if (-not $ObjectName) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-ObjectName is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $ObjectName
            )
        )
    }
    if (-not $SearchText) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-SearchText is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $SearchText
            )
        )
    }

    $app = Connect-AccessDB -DbPath $DbPath
    $cm = Get-CodeModule -App $app -ObjectType $ObjectType -ObjectName $ObjectName
    $cacheKey = "${ObjectType}:${ObjectName}"
    $allCode = Get-AllModuleCode -CodeModule $cm -CacheKey $cacheKey

    if (-not $allCode) {
        return Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{ found = $false; match_count = 0; matches = @() })
    }

    $matchList = [System.Collections.Generic.List[object]]::new()
    $lineNum = 0
    foreach ($line in $allCode.Split("`n")) {
        $lineNum++
        if (Test-TextMatch -Needle $SearchText -Haystack $line -MatchCase:$MatchCase -UseRegex:$UseRegex) {
            $matchList.Add([PSCustomObject][ordered]@{
                line    = $lineNum
                content = $line.TrimEnd("`r")
            })
        }
    }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        found       = ($matchList.Count -gt 0)
        match_count = $matchList.Count
        matches     = @($matchList)
    })
}

function Search-AccessVbe {
    <#
    .SYNOPSIS
        Search for text (or regex) across ALL VBA modules, forms, and reports.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER SearchText
        Text or regex pattern to find.
    .PARAMETER MatchCase
        Case-sensitive matching.
    .PARAMETER UseRegex
        Treat SearchText as a regex pattern.
    .PARAMETER MaxResults
        Maximum total matches to return (default 100).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Search-AccessVbe -DbPath "C:\db.accdb" -SearchText "DoCmd"
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [string]$SearchText,
        [switch]$MatchCase,
        [switch]$UseRegex,
        [int]$MaxResults = 100,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Search-AccessVbe'
    if (-not $SearchText) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-SearchText is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $SearchText
            )
        )
    }

    $app = Connect-AccessDB -DbPath $DbPath
    $objects = Get-AccessObject -DbPath $DbPath -ObjectType all
    $results = [System.Collections.Generic.List[object]]::new()
    $total = 0
    $truncated = $false

    foreach ($objType in @('module', 'form', 'report')) {
        if ($truncated) { break }
        $names = @($objects.$objType)
        foreach ($objName in $names) {
            if ($truncated) { break }
            try {
                $cm = Get-CodeModule -App $app -ObjectType $objType -ObjectName $objName
                $cacheKey = "${objType}:${objName}"
                $allCode = Get-AllModuleCode -CodeModule $cm -CacheKey $cacheKey
                if (-not $allCode) { continue }

                $objMatches = [System.Collections.Generic.List[object]]::new()
                $lineNum = 0
                foreach ($line in $allCode.Split("`n")) {
                    $lineNum++
                    if (Test-TextMatch -Needle $SearchText -Haystack $line -MatchCase:$MatchCase -UseRegex:$UseRegex) {
                        $objMatches.Add([PSCustomObject][ordered]@{
                            line    = $lineNum
                            content = $line.TrimEnd("`r")
                        })
                        $total++
                        if ($total -ge $MaxResults) { $truncated = $true; break }
                    }
                }
                if ($objMatches.Count -gt 0) {
                    $results.Add([PSCustomObject][ordered]@{
                        object_type = $objType
                        object_name = $objName
                        matches     = @($objMatches)
                    })
                }
            } catch { continue }
        }
    }

    $out = [ordered]@{ total_matches = $total; results = @($results) }
    if ($truncated) { $out['truncated'] = $true }
    Format-AccessOutput -AsJson:$AsJson -Data $out
}

function Search-AccessQuery {
    <#
    .SYNOPSIS
        Search for text (or regex) in the SQL of all queries.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER SearchText
        Text or regex pattern to find.
    .PARAMETER MatchCase
        Case-sensitive matching.
    .PARAMETER UseRegex
        Treat SearchText as a regex pattern.
    .PARAMETER MaxResults
        Maximum results to return (default 100).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Search-AccessQuery -DbPath "C:\db.accdb" -SearchText "Users"
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [string]$SearchText,
        [switch]$MatchCase,
        [switch]$UseRegex,
        [int]$MaxResults = 100,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Search-AccessQuery'
    if (-not $SearchText) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-SearchText is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $SearchText
            )
        )
    }

    $app = Connect-AccessDB -DbPath $DbPath
    $db = $app.CurrentDb()
    $results = [System.Collections.Generic.List[object]]::new()
    $total = 0

    foreach ($qd in $db.QueryDefs) {
        $name = $qd.Name
        if ($name.StartsWith('~')) { continue }
        $sql = $qd.SQL
        if (Test-TextMatch -Needle $SearchText -Haystack $sql -MatchCase:$MatchCase -UseRegex:$UseRegex) {
            $results.Add([PSCustomObject][ordered]@{
                query_name = $name
                sql        = $sql.Trim()
            })
            $total++
            if ($total -ge $MaxResults) { break }
        }
    }

    $out = [ordered]@{ total_matches = $total; results = @($results) }
    if ($total -ge $MaxResults) { $out['truncated'] = $true }
    Format-AccessOutput -AsJson:$AsJson -Data $out
}

function Find-AccessUsage {
    <#
    .SYNOPSIS
        Search for a name across VBA code, query SQL, and form/report control properties.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER SearchText
        Text or regex pattern to find.
    .PARAMETER MatchCase
        Case-sensitive matching.
    .PARAMETER UseRegex
        Treat SearchText as a regex pattern.
    .PARAMETER MaxResults
        Maximum total matches (default 200).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Find-AccessUsage -DbPath "C:\db.accdb" -SearchText "Users"
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [string]$SearchText,
        [switch]$MatchCase,
        [switch]$UseRegex,
        [int]$MaxResults = 200,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Find-AccessUsage'
    if (-not $SearchText) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-SearchText is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $SearchText
            )
        )
    }

    # 1. VBA matches
    $vbaResult = Search-AccessVbe -DbPath $DbPath -SearchText $SearchText -MatchCase:$MatchCase -UseRegex:$UseRegex -MaxResults $MaxResults
    $vbaMatches = [System.Collections.Generic.List[object]]::new()
    foreach ($group in @($vbaResult.results)) {
        foreach ($m in @($group.matches)) {
            $vbaMatches.Add([PSCustomObject][ordered]@{
                object_type = $group.object_type
                object_name = $group.object_name
                line        = $m.line
                content     = $m.content
            })
        }
    }
    $total = $vbaMatches.Count
    $truncated = [bool]$vbaResult.truncated

    # 2. Query matches
    $queryMatches = [System.Collections.Generic.List[object]]::new()
    if (-not $truncated) {
        $remaining = $MaxResults - $total
        $qryResult = Search-AccessQuery -DbPath $DbPath -SearchText $SearchText -MatchCase:$MatchCase -UseRegex:$UseRegex -MaxResults $remaining
        foreach ($q in @($qryResult.results)) { $queryMatches.Add($q) }
        $total += $qryResult.total_matches
        $truncated = [bool]$qryResult.truncated
    }

    # 3. Control property matches — search form/report exports
    $controlMatches = [System.Collections.Generic.List[object]]::new()
    if (-not $truncated) {
        $app = $script:AccessSession.App
        $objects = Get-AccessObject -DbPath $DbPath -ObjectType all
        foreach ($objType in @('form', 'report')) {
            if ($truncated) { break }
            foreach ($objName in @($objects.$objType)) {
                if ($truncated) { break }
                try {
                    $tmp = [System.IO.Path]::GetTempFileName()
                    try {
                        $app.SaveAsText($script:AC_TYPE[$objType], $objName, $tmp)
                        $rawText = (Read-TempFile -Path $tmp).Content
                    } finally {
                        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
                    }
                    foreach ($line in $rawText.Split("`n")) {
                        $stripped = $line.Trim()
                        foreach ($prop in $script:CONTROL_SEARCH_PROPS) {
                            if ($stripped.StartsWith("$prop =")) {
                                $valuePart = $stripped.Substring($prop.Length + 2).Trim()
                                if (Test-TextMatch -Needle $SearchText -Haystack $valuePart -MatchCase:$MatchCase -UseRegex:$UseRegex) {
                                    $controlMatches.Add([PSCustomObject][ordered]@{
                                        object_type = $objType
                                        object_name = $objName
                                        property    = $prop
                                        value       = $valuePart
                                    })
                                    $total++
                                    if ($total -ge $MaxResults) { $truncated = $true }
                                    break
                                }
                            }
                        }
                    }
                } catch { continue }
            }
        }
    }

    $out = [ordered]@{
        search_text     = $SearchText
        vba_matches     = @($vbaMatches)
        query_matches   = @($queryMatches)
        control_matches = @($controlMatches)
        total_matches   = $total
    }
    if ($truncated) { $out['truncated'] = $true }
    Format-AccessOutput -AsJson:$AsJson -Data $out
}

function Invoke-AccessMacro {
    <#
    .SYNOPSIS
        Run an Access macro by name.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER MacroName
        Name of the macro to run.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Invoke-AccessMacro -DbPath "C:\db.accdb" -MacroName "AutoExec"
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [string]$MacroName,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Invoke-AccessMacro'
    if (-not $MacroName) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-MacroName is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $MacroName
            )
        )
    }

    $app = Connect-AccessDB -DbPath $DbPath
    try {
        $app.DoCmd.RunMacro($MacroName)
    } catch {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.InvalidOperationException]::new("Error running macro '$MacroName': $_"),
                'InvalidOperation',
                [System.Management.Automation.ErrorCategory]::InvalidOperation,
                $MacroName
            )
        )
    }

    Format-AccessOutput -AsJson:$AsJson -Data @{
        macro_name = $MacroName
        status     = 'executed'
    }
}

function Invoke-AccessVba {
    <#
    .SYNOPSIS
        Call a VBA Sub/Function via Application.Run or Forms COM access.
    .DESCRIPTION
        Supports two syntaxes:
        - 'ModuleName.ProcName' or 'ProcName' -> Application.Run (standard modules)
        - 'Forms.FormName.Method' -> COM Forms() access (form must be open)
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER Procedure
        Procedure name or qualified path.
    .PARAMETER Arguments
        Arguments to pass to the procedure (max 30 for Application.Run).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Invoke-AccessVba -DbPath "C:\db.accdb" -Procedure "Module1.Calculate" -Arguments @(42)
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [string]$Procedure,
        [object[]]$Arguments,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Invoke-AccessVba'
    if (-not $Procedure) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-Procedure is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $Procedure
            )
        )
    }

    $app = Connect-AccessDB -DbPath $DbPath
    $callArgs = if ($Arguments) { $Arguments } else { @() }

    if ($callArgs.Count -gt 30) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentOutOfRangeException]::new('Arguments', $callArgs.Count, 'Application.Run supports max 30 arguments.'),
                'OutOfRange',
                [System.Management.Automation.ErrorCategory]::LimitsExceeded,
                $callArgs
            )
        )
    }

    # Forms.FormName.Method -> direct COM access
    if ($Procedure.Contains('.')) {
        $parts = $Procedure.Split('.', 3)
        if ($parts[0] -eq 'Forms' -and $parts.Count -eq 3) {
            $formName   = $parts[1]
            $methodName = $parts[2]
            try {
                $form = $app.Forms($formName)
                if ($callArgs.Count -gt 0) {
                    $result = $form.GetType().InvokeMember(
                        $methodName,
                        [System.Reflection.BindingFlags]::InvokeMethod,
                        $null, $form, $callArgs
                    )
                } else {
                    # Try method call, fall back to property
                    try {
                        $result = $form.GetType().InvokeMember(
                            $methodName,
                            [System.Reflection.BindingFlags]::InvokeMethod,
                            $null, $form, @()
                        )
                    } catch {
                        $result = $form.GetType().InvokeMember(
                            $methodName,
                            [System.Reflection.BindingFlags]::GetProperty,
                            $null, $form, @()
                        )
                    }
                }
            } catch {
                $PSCmdlet.ThrowTerminatingError(
                    [System.Management.Automation.ErrorRecord]::new(
                        [System.InvalidOperationException]::new("Error calling Forms('$formName').$methodName : $_. Make sure the form is open."),
                        'InvalidOperation',
                        [System.Management.Automation.ErrorCategory]::InvalidOperation,
                        $formName
                    )
                )
            }

            return Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
                procedure = $Procedure
                result    = $result
                status    = 'executed'
            })
        }
    }

    # Standard Application.Run
    try {
        $result = switch ($callArgs.Count) {
            0  { $app.Run($Procedure) }
            1  { $app.Run($Procedure, $callArgs[0]) }
            2  { $app.Run($Procedure, $callArgs[0], $callArgs[1]) }
            3  { $app.Run($Procedure, $callArgs[0], $callArgs[1], $callArgs[2]) }
            4  { $app.Run($Procedure, $callArgs[0], $callArgs[1], $callArgs[2], $callArgs[3]) }
            5  { $app.Run($Procedure, $callArgs[0], $callArgs[1], $callArgs[2], $callArgs[3], $callArgs[4]) }
            default {
                # Build args array for Invoke
                $invokeArgs = @($Procedure) + $callArgs
                $app.GetType().InvokeMember('Run', [System.Reflection.BindingFlags]::InvokeMethod, $null, $app, $invokeArgs)
            }
        }
    } catch {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.InvalidOperationException]::new("Error running '$Procedure': $_"),
                'InvalidOperation',
                [System.Management.Automation.ErrorCategory]::InvalidOperation,
                $Procedure
            )
        )
    }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        procedure = $Procedure
        result    = $result
        status    = 'executed'
    })
}

function Invoke-AccessEval {
    <#
    .SYNOPSIS
        Evaluate a VBA/Access expression via Application.Eval.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER Expression
        Expression to evaluate (e.g., "Date()", "DLookup(...)").
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Invoke-AccessEval -DbPath "C:\db.accdb" -Expression "Date()"
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [string]$Expression,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Invoke-AccessEval'
    if (-not $Expression) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-Expression is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $Expression
            )
        )
    }

    $app = Connect-AccessDB -DbPath $DbPath
    try {
        $result = $app.Eval($Expression)
    } catch {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.InvalidOperationException]::new("Error evaluating '$Expression': $_"),
                'InvalidOperation',
                [System.Management.Automation.ErrorCategory]::InvalidOperation,
                $Expression
            )
        )
    }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        expression = $Expression
        result     = $result
        status     = 'evaluated'
    })
}

function Test-AccessVbaCompile {
    <#
    .SYNOPSIS
        Compile and save all VBA modules. Returns status and error location if compilation fails.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Test-AccessVbaCompile -DbPath "C:\db.accdb"
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Test-AccessVbaCompile'

    $app = Connect-AccessDB -DbPath $DbPath

    try {
        $app.RunCommand($script:AC_CMD_COMPILE)
    } catch {
        # Try to get error location from VBE
        $errLoc = $null
        try {
            $pane = $app.VBE.ActiveCodePane
            if ($null -ne $pane) {
                $cm = $pane.CodeModule
                $startLine = 0; $startCol = 0; $endLine = 0; $endCol = 0
                $pane.GetSelection([ref]$startLine, [ref]$startCol, [ref]$endLine, [ref]$endCol)
                $errCode = ''
                if ($startLine -gt 0 -and $cm.CountOfLines -ge $startLine) {
                    $errCode = $cm.Lines($startLine, 1)
                }
                $errLoc = [ordered]@{
                    component = $cm.Parent.Name
                    line      = $startLine
                    code      = $errCode
                }
            }
        } catch {}

        $result = [ordered]@{
            status       = 'error'
            error_detail = "VBA compilation error: $_"
        }
        if ($errLoc) { $result['error_location'] = $errLoc }
        return Format-AccessOutput -AsJson:$AsJson -Data $result
    }

    # Invalidate caches — compilation may change module state
    $script:AccessSession.VbeCodeCache = @{}
    $script:AccessSession.CmCache = @{}

    Format-AccessOutput -AsJson:$AsJson -Data @{ status = 'compiled' }
}

function Import-AccessVbaFile {
    <#
    .SYNOPSIS
        Import a .bas (standard module) or .cls (class module) file into an Access
        database via VBComponents.Import. Validates ANSI encoding and auto-converts
        if needed. Replaces any existing component with the same name.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER FilePath
        Path to the .bas or .cls file to import.
    .PARAMETER Force
        Auto-convert non-ANSI files to a temp ANSI copy before importing (default).
        Set -Force:$false to error on non-ANSI files instead of converting.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Import-AccessVbaFile -DbPath "C:\db.accdb" -FilePath "C:\modules\clsHelper.cls"
    .EXAMPLE
        Import-AccessVbaFile -DbPath "C:\db.accdb" -FilePath "C:\modules\modUtils.bas" -AsJson
    .EXAMPLE
        # Import multiple files
        Get-ChildItem "C:\vba\*.cls","C:\vba\*.bas" | ForEach-Object {
            Import-AccessVbaFile -DbPath "C:\db.accdb" -FilePath $_.FullName -AsJson
        }
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [string]$FilePath,
        [switch]$Force = $true,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Import-AccessVbaFile'
    if (-not $FilePath) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-FilePath is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $FilePath
            )
        )
    }
    if (-not (Test-Path -LiteralPath $FilePath -PathType Leaf)) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.IO.FileNotFoundException]::new("File not found: $FilePath"),
                'ObjectNotFound',
                [System.Management.Automation.ErrorCategory]::ObjectNotFound,
                $FilePath
            )
        )
    }

    $FilePath = (Resolve-Path -LiteralPath $FilePath).Path
    $ext = [System.IO.Path]::GetExtension($FilePath).ToLower()
    if ($ext -notin '.bas', '.cls') {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new("Only .bas and .cls files are supported. Got '$ext'."),
                'InvalidArgument',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $FilePath
            )
        )
    }

    # Validate encoding
    $encCheck = Test-VbaFileEncoding -Path $FilePath
    $importPath = $FilePath
    $converted = $false
    $tmpPath = $null

    if (-not $encCheck.IsAnsi) {
        if (-not $Force) {
            $PSCmdlet.ThrowTerminatingError(
                [System.Management.Automation.ErrorRecord]::new(
                    [System.InvalidOperationException]::new("$($encCheck.Reason) Use -Force to auto-convert."),
                    'InvalidOperation',
                    [System.Management.Automation.ErrorCategory]::InvalidOperation,
                    $FilePath
                )
            )
        }
        Write-Verbose "Non-ANSI encoding detected ($($encCheck.Encoding)). Converting to ANSI temp copy."
        $tmpPath = ConvertTo-AnsiTempFile -SourcePath $FilePath
        $importPath = $tmpPath
        $converted = $true
    }

    try {
        $app = Connect-AccessDB -DbPath $DbPath
        $proj = $app.VBE.ActiveVBProject

        # Derive module name from the file's Attribute VB_Name or from filename
        $moduleName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)

        # Remove existing component if present
        try {
            $existing = $proj.VBComponents.Item($moduleName)
            if ($null -ne $existing) {
                $proj.VBComponents.Remove($existing)
                Write-Verbose "Removed existing component: $moduleName"
            }
        } catch {
            # Component doesn't exist — that's fine
        }

        # Import via VBComponents.Import — correctly handles .cls as class module
        $imported = $proj.VBComponents.Import($importPath)
        $typeName = switch ($imported.Type) {
            1 { 'standard_module' }
            2 { 'class_module' }
            default { "type_$($imported.Type)" }
        }

        # Invalidate VBE caches
        $cacheKey = "module:$($imported.Name)"
        $script:AccessSession.VbeCodeCache.Remove($cacheKey)
        $script:AccessSession.CmCache.Remove($cacheKey)

        $result = [ordered]@{
            status       = 'imported'
            name         = $imported.Name
            module_type  = $typeName
            source_file  = $FilePath
            converted    = $converted
        }
        if ($converted) {
            $result['original_encoding'] = $encCheck.Encoding
        }
        Format-AccessOutput -AsJson:$AsJson -Data $result
    } finally {
        if ($tmpPath -and (Test-Path -LiteralPath $tmpPath)) {
            Remove-Item -LiteralPath $tmpPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Test-AccessVbaFileEncoding {
    <#
    .SYNOPSIS
        Check whether a .bas or .cls file has the correct ANSI encoding
        (Windows-1252, no BOM) required by VBComponents.Import.
    .PARAMETER FilePath
        Path to the .bas or .cls file to check.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Test-AccessVbaFileEncoding -FilePath "C:\modules\clsHelper.cls" -AsJson
    .EXAMPLE
        Get-ChildItem "C:\vba\*" -Include *.bas,*.cls | ForEach-Object {
            Test-AccessVbaFileEncoding -FilePath $_.FullName -AsJson
        }
    #>
    [CmdletBinding()]
    param(
        [string]$FilePath,
        [switch]$AsJson
    )

    if (-not $FilePath) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-FilePath is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $FilePath
            )
        )
    }
    if (-not (Test-Path -LiteralPath $FilePath -PathType Leaf)) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.IO.FileNotFoundException]::new("File not found: $FilePath"),
                'ObjectNotFound',
                [System.Management.Automation.ErrorCategory]::ObjectNotFound,
                $FilePath
            )
        )
    }

    $FilePath = (Resolve-Path -LiteralPath $FilePath).Path
    $check = Test-VbaFileEncoding -Path $FilePath

    $result = [ordered]@{
        file     = $FilePath
        is_ansi  = $check.IsAnsi
        encoding = $check.Encoding
    }
    if (-not $check.IsAnsi) {
        $result['reason'] = $check.Reason
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Find-AccessDefinition {
    <#
    .SYNOPSIS
        Go-to-definition for a VBA symbol. Scans standard modules, form code-behind,
        and report code-behind for DECLARATIONS of the given symbol.
    .DESCRIPTION
        Detects: const, enum, enum_member, type, type_field, sub, function, property
        (Get/Let/Set, incl. Default Property), declare (Win32 API), and module-level
        variable. Handles multi-const lines (Const A=1, B=2) and joins VBA line
        continuations (` _` at end of line). Case-insensitive by default.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER Symbol
        Name to resolve (e.g. 'dbAccess', 'ccRed', 'ProcessInvoices').
    .PARAMETER Kinds
        Optional whitelist of definition kinds. Default: all kinds.
        Valid: const, enum, enum_member, type, type_field, sub, function, property, declare, variable.
    .PARAMETER MatchCase
        Case-sensitive symbol matching.
    .PARAMETER ScanTypes
        Which object types to scan. Default: module, form, report.
    .PARAMETER FirstOnly
        Stop after the first match.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Find-AccessDefinition -DbPath "C:\db.accdb" -Symbol "ccRed"
    .EXAMPLE
        Find-AccessDefinition -Symbol "ProcessInvoices" -Kinds function,sub -FirstOnly
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [Parameter(Mandatory)]
        [string]$Symbol,
        [string[]]$Kinds,
        [switch]$MatchCase,
        [ValidateSet('module','form','report')]
        [string[]]$ScanTypes,
        [switch]$FirstOnly,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Find-AccessDefinition'

    # --- Validate kinds -----------------------------------------------------------
    $allValidKinds = @('const','enum','enum_member','type','type_field',
                       'sub','function','property','declare','variable')

    $filterKinds = $null
    if ($Kinds -and $Kinds.Count -gt 0) {
        $invalid = @($Kinds | Where-Object { $_ -notin $allValidKinds })
        if ($invalid.Count -gt 0) {
            $PSCmdlet.ThrowTerminatingError(
                [System.Management.Automation.ErrorRecord]::new(
                    [System.ArgumentException]::new(
                        "Invalid kinds: $($invalid -join ', '). Valid: $($allValidKinds -join ', ')"),
                    'InvalidKinds',
                    [System.Management.Automation.ErrorCategory]::InvalidArgument,
                    $Kinds
                )
            )
        }
        $filterKinds = [System.Collections.Generic.HashSet[string]]::new(
            [string[]]$Kinds,
            [System.StringComparer]::OrdinalIgnoreCase
        )
    }

    if (-not $ScanTypes -or $ScanTypes.Count -eq 0) {
        $ScanTypes = @('module','form','report')
    }

    # --- Compiled regex patterns --------------------------------------------------
    $rxOpts = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase

    $rxProc    = [regex]::new(
        '^\s*(Public\s+|Private\s+|Friend\s+|Global\s+)?(Static\s+)?(Default\s+)?(Sub|Function|Property\s+(Get|Let|Set))\s+(\w+)',
        $rxOpts)
    $rxConst   = [regex]::new(
        '^\s*(Public\s+|Private\s+|Global\s+)?Const\s+',
        $rxOpts)
    $rxEnum    = [regex]::new(
        '^\s*(Public\s+|Private\s+)?Enum\s+(\w+)',
        $rxOpts)
    $rxType    = [regex]::new(
        '^\s*(Public\s+|Private\s+)?Type\s+(\w+)',
        $rxOpts)
    $rxDeclare = [regex]::new(
        '^\s*(Public\s+|Private\s+)?Declare\s+(PtrSafe\s+)?(Sub|Function)\s+(\w+)\s+Lib\s+',
        $rxOpts)
    $rxVar     = [regex]::new(
        '^\s*(Dim|Public|Private|Global)\s+(?!Const\b|Enum\b|Type\b|Sub\b|Function\b|Property\b|Declare\b)',
        $rxOpts)
    $rxEndProc = [regex]::new('^\s*End\s+(Sub|Function|Property)\b', $rxOpts)
    $rxEndEnum = [regex]::new('^\s*End\s+Enum\b', $rxOpts)
    $rxEndType = [regex]::new('^\s*End\s+Type\b', $rxOpts)
    $rxEnumMem = [regex]::new('^\s*(\w+)\s*(=\s*(.+?))?\s*$')
    $rxTypeFld = [regex]::new('^\s*(\w+)(\(.+?\))?\s+As\s+(.+)', $rxOpts)
    $rxScopePfx = [regex]::new('^\s*(Public|Private|Global)\s+', $rxOpts)
    $rxConstAfter = [regex]::new('^\s*(Public\s+|Private\s+|Global\s+)?Const\s+', $rxOpts)
    $rxScopeKw   = [regex]::new('^\s*(Dim|Public|Private|Global)\s+', $rxOpts)

    # --- Name matcher -------------------------------------------------------------
    $nameMatch = if ($MatchCase) {
        { param($c) $c -ceq $Symbol }
    } else {
        { param($c) $c -ieq $Symbol }
    }

    # --- Connect and iterate components -------------------------------------------
    $app = Connect-AccessDB -DbPath $DbPath
    $project = $app.VBE.ActiveVBProject
    $defs = [System.Collections.Generic.List[PSCustomObject]]::new()

    $earlyReturn = $false

    foreach ($comp in $project.VBComponents) {
        if ($earlyReturn) { break }

        $compType = [int]$comp.Type
        # Skip UserForm (type 3)
        if ($compType -eq 3) { continue }

        $objType  = $null
        $objName  = $null
        $compName = $comp.Name

        if ($compType -eq 1) {
            # Standard module
            if ('module' -notin $ScanTypes) { continue }
            $objType = 'module'
            $objName = $compName
        }
        elseif ($compName.StartsWith('Form_')) {
            if ('form' -notin $ScanTypes) { continue }
            $objType = 'form'
            $objName = $compName.Substring(5)
        }
        elseif ($compName.StartsWith('Report_')) {
            if ('report' -notin $ScanTypes) { continue }
            $objType = 'report'
            $objName = $compName.Substring(7)
        }
        else {
            # Class module or other — treat as module
            if ('module' -notin $ScanTypes) { continue }
            $objType = 'module'
            $objName = $compName
        }

        # Read code
        $cm    = $comp.CodeModule
        $total = $cm.CountOfLines
        if ($total -eq 0) { continue }
        $code  = $cm.Lines(1, $total)

        # Join continuations
        $joined = Join-VbaContinuations -Code $code

        # --- State machine --------------------------------------------------------
        $inProc     = $false
        $inEnum     = $false
        $inType     = $false
        $curEnum    = $null
        $curType    = $null

        foreach ($entry in $joined) {
            if ($earlyReturn) { break }

            $physLine = $entry.Line
            $text     = $entry.Text
            $stripped = $text.Trim()

            # Skip blank lines and comments
            if (-not $stripped -or $stripped.StartsWith("'")) { continue }

            # ---- Inside a procedure: skip until End Sub/Function/Property --------
            if ($inProc) {
                if ($rxEndProc.IsMatch($text)) { $inProc = $false }
                continue
            }

            # ---- Procedure declaration -------------------------------------------
            $m = $rxProc.Match($text)
            if ($m.Success) {
                $scope   = ($m.Groups[1].Value).Trim()
                if (-not $scope) { $scope = $null }
                $isStatic  = [bool]$m.Groups[2].Value
                $keyword   = $m.Groups[4].Value
                $procName  = $m.Groups[6].Value

                if ($keyword -imatch '^Property') {
                    $kind    = 'property'
                    $subkind = $keyword
                }
                elseif ($keyword -ieq 'Sub') {
                    $kind    = 'sub'
                    $subkind = $null
                }
                else {
                    $kind    = 'function'
                    $subkind = $null
                }

                if ((& $nameMatch $procName) -and (-not $filterKinds -or $filterKinds.Contains($kind))) {
                    $defn = [PSCustomObject][ordered]@{
                        kind        = $kind
                        object_type = $objType
                        object_name = $objName
                        line        = $physLine
                        declaration = $stripped
                        scope       = $scope
                    }
                    if ($isStatic) {
                        $defn.scope = "Static$(if ($scope) { " $scope" })"
                    }
                    if ($subkind) {
                        $defn | Add-Member -NotePropertyName subkind -NotePropertyValue $subkind
                    }
                    $defs.Add($defn)
                    if ($FirstOnly) { $earlyReturn = $true; break }
                }

                $inProc = $true
                continue
            }

            # ---- Enum block start ------------------------------------------------
            $m = $rxEnum.Match($text)
            if ($m.Success -and -not $inEnum) {
                $scope    = ($m.Groups[1].Value).Trim()
                if (-not $scope) { $scope = $null }
                $enumName = $m.Groups[2].Value
                $inEnum   = $true
                $curEnum  = $enumName

                if ((& $nameMatch $enumName) -and (-not $filterKinds -or $filterKinds.Contains('enum'))) {
                    $defs.Add([PSCustomObject][ordered]@{
                        kind        = 'enum'
                        object_type = $objType
                        object_name = $objName
                        line        = $physLine
                        declaration = $stripped
                        scope       = $scope
                    })
                    if ($FirstOnly) { $earlyReturn = $true; break }
                }
                continue
            }

            # ---- Inside Enum block -----------------------------------------------
            if ($inEnum) {
                if ($rxEndEnum.IsMatch($text)) {
                    $inEnum  = $false
                    $curEnum = $null
                    continue
                }
                # Enum member
                $m = $rxEnumMem.Match($stripped)
                if ($m.Success) {
                    $memName = $m.Groups[1].Value
                    $memVal  = ($m.Groups[3].Value).Trim()
                    if (-not $memVal) { $memVal = $null }

                    if ((& $nameMatch $memName) -and (-not $filterKinds -or $filterKinds.Contains('enum_member'))) {
                        $defs.Add([PSCustomObject][ordered]@{
                            kind        = 'enum_member'
                            object_type = $objType
                            object_name = $objName
                            line        = $physLine
                            declaration = $stripped
                            scope       = $null
                            parent_enum = $curEnum
                            value       = $memVal
                        })
                        if ($FirstOnly) { $earlyReturn = $true; break }
                    }
                }
                continue
            }

            # ---- Type block start ------------------------------------------------
            $m = $rxType.Match($text)
            if ($m.Success -and -not $inType) {
                $scope    = ($m.Groups[1].Value).Trim()
                if (-not $scope) { $scope = $null }
                $typeName = $m.Groups[2].Value
                $inType   = $true
                $curType  = $typeName

                if ((& $nameMatch $typeName) -and (-not $filterKinds -or $filterKinds.Contains('type'))) {
                    $defs.Add([PSCustomObject][ordered]@{
                        kind        = 'type'
                        object_type = $objType
                        object_name = $objName
                        line        = $physLine
                        declaration = $stripped
                        scope       = $scope
                    })
                    if ($FirstOnly) { $earlyReturn = $true; break }
                }
                continue
            }

            # ---- Inside Type block -----------------------------------------------
            if ($inType) {
                if ($rxEndType.IsMatch($text)) {
                    $inType  = $false
                    $curType = $null
                    continue
                }
                # Type field
                $m = $rxTypeFld.Match($stripped)
                if ($m.Success) {
                    $fldName = $m.Groups[1].Value
                    $asType  = ($m.Groups[3].Value).Trim()

                    if ((& $nameMatch $fldName) -and (-not $filterKinds -or $filterKinds.Contains('type_field'))) {
                        $defs.Add([PSCustomObject][ordered]@{
                            kind        = 'type_field'
                            object_type = $objType
                            object_name = $objName
                            line        = $physLine
                            declaration = $stripped
                            scope       = $null
                            parent_type = $curType
                            as_type     = $asType
                        })
                        if ($FirstOnly) { $earlyReturn = $true; break }
                    }
                }
                continue
            }

            # ---- Const -----------------------------------------------------------
            if ($rxConst.IsMatch($text)) {
                $scopeM = $rxScopePfx.Match($text)
                $scope  = if ($scopeM.Success) { $scopeM.Groups[1].Value } else { $null }
                $afterConst = $rxConstAfter.Replace($text, '', 1)
                $parts = Split-TopLevelCommas -Text $afterConst

                foreach ($part in $parts) {
                    if ($earlyReturn) { break }
                    $cm2 = [regex]::Match($part, '(\w+)')
                    if ($cm2.Success) {
                        $constName = $cm2.Groups[1].Value
                        $valM = [regex]::Match($part, '=\s*(.+)')
                        $constVal = if ($valM.Success) { $valM.Groups[1].Value.Trim() } else { $null }

                        if ((& $nameMatch $constName) -and (-not $filterKinds -or $filterKinds.Contains('const'))) {
                            $defs.Add([PSCustomObject][ordered]@{
                                kind        = 'const'
                                object_type = $objType
                                object_name = $objName
                                line        = $physLine
                                declaration = $stripped
                                scope       = $scope
                                value       = $constVal
                            })
                            if ($FirstOnly) { $earlyReturn = $true; break }
                        }
                    }
                }
                continue
            }

            # ---- Declare (Win32 API) ---------------------------------------------
            $m = $rxDeclare.Match($text)
            if ($m.Success) {
                $scope    = ($m.Groups[1].Value).Trim()
                if (-not $scope) { $scope = $null }
                $subkind  = $m.Groups[3].Value
                $declName = $m.Groups[4].Value

                if ((& $nameMatch $declName) -and (-not $filterKinds -or $filterKinds.Contains('declare'))) {
                    $defs.Add([PSCustomObject][ordered]@{
                        kind        = 'declare'
                        object_type = $objType
                        object_name = $objName
                        line        = $physLine
                        declaration = $stripped
                        scope       = $scope
                        subkind     = $subkind
                    })
                    if ($FirstOnly) { $earlyReturn = $true; break }
                }
                continue
            }

            # ---- Variable (module-level only — in_proc is false here) ------------
            $m = $rxVar.Match($text)
            if ($m.Success) {
                $varScope  = $m.Groups[1].Value
                $afterScope = $rxScopeKw.Replace($text, '', 1)
                $parts = Split-TopLevelCommas -Text $afterScope

                foreach ($part in $parts) {
                    if ($earlyReturn) { break }
                    $vm = [regex]::Match($part, '(\w+)')
                    if ($vm.Success) {
                        $varName = $vm.Groups[1].Value
                        $asM = [regex]::Match($part, '\bAs\s+(.+)', $rxOpts)
                        $asType = if ($asM.Success) { $asM.Groups[1].Value.Trim() } else { $null }

                        if ((& $nameMatch $varName) -and (-not $filterKinds -or $filterKinds.Contains('variable'))) {
                            $defs.Add([PSCustomObject][ordered]@{
                                kind        = 'variable'
                                object_type = $objType
                                object_name = $objName
                                line        = $physLine
                                declaration = $stripped
                                scope       = $varScope
                                as_type     = $asType
                            })
                            if ($FirstOnly) { $earlyReturn = $true; break }
                        }
                    }
                }
                continue
            }
        }
    }

    $out = [ordered]@{
        symbol      = $Symbol
        total       = $defs.Count
        definitions = @($defs)
    }
    Format-AccessOutput -AsJson:$AsJson -Data $out
}
