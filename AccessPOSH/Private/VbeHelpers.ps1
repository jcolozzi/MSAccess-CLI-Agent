# Private/VbeHelpers.ps1 — VBE CodeModule access, text matching, context helpers

function Test-TextMatch {
    <#
    .SYNOPSIS
        Internal: match needle against haystack (plain substring or regex).
    #>
    param(
        [string]$Needle,
        [string]$Haystack,
        [bool]$MatchCase = $false,
        [bool]$UseRegex = $false
    )

    if ($UseRegex) {
        $opts = [System.Text.RegularExpressions.RegexOptions]::None
        if (-not $MatchCase) { $opts = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase }
        return [regex]::IsMatch($Haystack, $Needle, $opts)
    }
    if (-not $MatchCase) {
        return $Haystack.IndexOf($Needle, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
    }
    return $Haystack.Contains($Needle)
}

function Get-CodeModule {
    <#
    .SYNOPSIS
        Internal: Get cached VBE CodeModule COM object for a module/form/report.
    #>
    param(
        $App,
        [string]$ObjectType,
        [string]$ObjectName
    )

    if (-not $script:VBE_PREFIX.ContainsKey($ObjectType)) {
        throw "object_type '$ObjectType' does not support VBE. Use 'module', 'form', or 'report'."
    }

    $cacheKey = "${ObjectType}:${ObjectName}"
    $cm = $script:AccessSession.CmCache[$cacheKey]
    if ($null -ne $cm) { return $cm }

    $compName = $script:VBE_PREFIX[$ObjectType] + $ObjectName
    try {
        $project = $App.VBE.VBProjects(1)
        $component = $project.VBComponents($compName)
        $cm = $component.CodeModule
        $script:AccessSession.CmCache[$cacheKey] = $cm
        return $cm
    } catch {
        $script:AccessSession.CmCache.Remove($cacheKey)
        throw "Cannot access CodeModule '$compName'. Is 'Trust access to the VBA project object model' enabled in Access Trust Center? Error: $_"
    }
}

function Get-AllModuleCode {
    <#
    .SYNOPSIS
        Internal: Get full module text using VbeCodeCache.
    #>
    param(
        $CodeModule,
        [string]$CacheKey
    )

    if (-not $script:AccessSession.VbeCodeCache.ContainsKey($CacheKey)) {
        $total = $CodeModule.CountOfLines
        $text = if ($total -gt 0) { $CodeModule.Lines(1, $total) } else { '' }
        $script:AccessSession.VbeCodeCache[$CacheKey] = $text
    }
    return $script:AccessSession.VbeCodeCache[$CacheKey]
}

function Test-WsNormalizedMatch {
    <#
    .SYNOPSIS
        Whitespace-tolerant matching: strips leading whitespace from each line
        and does a sliding-window search. Returns hashtable with start/end 0-based
        line indices, or $null if no match.
    #>
    param(
        [string]$ProcCode,
        [string]$FindText
    )
    $procLines = $ProcCode -split "`r?`n"
    $findLines = @($FindText -split "`r?`n")
    # Remove empty trailing lines from find text
    while ($findLines.Count -gt 0 -and -not $findLines[-1].Trim()) {
        $findLines = $findLines[0..($findLines.Count - 2)]
    }
    if ($findLines.Count -eq 0) { return $null }

    $procStripped = @($procLines | ForEach-Object { $_.TrimStart() })
    $findStripped = @($findLines | ForEach-Object { $_.TrimStart() })
    $window = $findStripped.Count

    for ($i = 0; $i -le ($procStripped.Count - $window); $i++) {
        $match = $true
        for ($j = 0; $j -lt $window; $j++) {
            if ($procStripped[$i + $j] -cne $findStripped[$j]) {
                $match = $false
                break
            }
        }
        if ($match) {
            return @{ start = $i; end = ($i + $window - 1) }
        }
    }
    return $null
}

function Get-ClosestMatchContext {
    <#
    .SYNOPSIS
        When both exact and ws-normalized match fail, finds the most similar line
        using character-level similarity and returns a contextual snippet.
    #>
    param(
        [string]$ProcCode,
        [string]$FindText,
        [string]$ProcName
    )
    $procLines = $ProcCode -split "`r?`n"
    $findLines = @(($FindText -split "`r?`n") | Where-Object { $_.Trim() })
    if ($findLines.Count -eq 0) { return "Empty find text in proc '$ProcName'" }

    $ref = $findLines[0].Trim()
    $bestRatio = 0.0
    $bestIdx = 0

    for ($i = 0; $i -lt $procLines.Count; $i++) {
        $candidate = $procLines[$i].Trim()
        if (-not $candidate) { continue }
        # Longest Common Subsequence ratio (simplified SequenceMatcher)
        $shorter = if ($ref.Length -lt $candidate.Length) { $ref } else { $candidate }
        $longer  = if ($ref.Length -lt $candidate.Length) { $candidate } else { $ref }
        $matchCount = 0
        $usedIdx = -1
        foreach ($ch in $shorter.ToCharArray()) {
            $pos = $longer.IndexOf($ch, $usedIdx + 1)
            if ($pos -ge 0) { $matchCount++; $usedIdx = $pos }
        }
        $ratio = if (($ref.Length + $candidate.Length) -gt 0) { (2.0 * $matchCount) / ($ref.Length + $candidate.Length) } else { 0 }
        if ($ratio -gt $bestRatio) {
            $bestRatio = $ratio
            $bestIdx = $i
        }
    }

    # Build context: 3 lines around best candidate
    $ctxStart = [math]::Max(0, $bestIdx - 1)
    $ctxEnd   = [math]::Min($procLines.Count - 1, $bestIdx + 1)
    $contextLines = @()
    for ($j = $ctxStart; $j -le $ctxEnd; $j++) {
        $marker = if ($j -eq $bestIdx) { '>>>' } else { '   ' }
        $contextLines += "  $marker L$($j + 1): $($procLines[$j].TrimEnd())"
    }
    $pct = [math]::Round($bestRatio * 100)
    $refSnippet = if ($ref.Length -gt 80) { $ref.Substring(0, 80) } else { $ref }
    return "Best match (${pct}% similar) near line $($bestIdx + 1) of '${ProcName}':`n$($contextLines -join "`n")`n  Looking for: '$refSnippet'"
}

function Join-VbaContinuations {
    <#
    .SYNOPSIS
        Join VBA lines ending with ' _' (line continuation). Returns array of
        [PSCustomObject]@{Line=[int]; Text=[string]} where Line is the 1-based
        physical line number of the first line in the joined group.
    .PARAMETER Code
        Full VBA module text (CRLF or LF line endings).
    #>
    param(
        [string]$Code
    )

    if (-not $Code) { return @() }
    $lines = if ($Code.Contains("`r`n")) { $Code.Split("`r`n") } else { $Code.Split("`n") }
    $result = [System.Collections.Generic.List[PSCustomObject]]::new()
    $i = 0
    while ($i -lt $lines.Count) {
        $parts = [System.Collections.Generic.List[string]]::new()
        $parts.Add($lines[$i].TrimEnd())
        $firstLine = $i + 1  # 1-based
        while ($i -lt $lines.Count - 1 -and $lines[$i].TrimEnd().EndsWith(' _')) {
            # Remove trailing ' _' from current part
            $parts[$parts.Count - 1] = $parts[$parts.Count - 1].Substring(0, $parts[$parts.Count - 1].Length - 2)
            $i++
            $parts.Add($lines[$i].Trim())
        }
        $joined = ($parts -join ' ')
        $result.Add([PSCustomObject]@{ Line = [int]$firstLine; Text = $joined })
        $i++
    }
    return @($result)
}

function Split-TopLevelCommas {
    <#
    .SYNOPSIS
        Split a string by commas that are NOT inside parentheses or quoted strings.
        VBA string escaping (doubled quotes) is respected.
    .PARAMETER Text
        The string to split (e.g. "A = 1, B = 2, C = Func(1, 2)").
    #>
    param(
        [string]$Text
    )

    if (-not $Text) { return @() }
    $parts   = [System.Collections.Generic.List[string]]::new()
    $depth   = 0
    $inStr   = $false
    $current = [System.Text.StringBuilder]::new()
    $i       = 0
    $len     = $Text.Length

    while ($i -lt $len) {
        $ch = $Text[$i]
        if ($ch -eq '"') {
            # VBA escapes " as "" inside strings
            if ($inStr -and ($i + 1) -lt $len -and $Text[$i + 1] -eq '"') {
                [void]$current.Append('""')
                $i += 2
                continue
            }
            $inStr = -not $inStr
            [void]$current.Append($ch)
        }
        elseif ($inStr) {
            [void]$current.Append($ch)
        }
        elseif ($ch -eq '(') {
            $depth++
            [void]$current.Append($ch)
        }
        elseif ($ch -eq ')') {
            $depth--
            [void]$current.Append($ch)
        }
        elseif ($ch -eq ',' -and $depth -eq 0) {
            $val = $current.ToString().Trim()
            if ($val) { $parts.Add($val) }
            [void]$current.Clear()
        }
        else {
            [void]$current.Append($ch)
        }
        $i++
    }
    $val = $current.ToString().Trim()
    if ($val) { $parts.Add($val) }
    return @($parts)
}
