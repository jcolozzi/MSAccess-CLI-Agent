# Private/Utilities.ps1 — Value conversion, output formatting, temp file I/O

function ConvertTo-SafeValue {
    <#
    .SYNOPSIS
        Convert COM values to PowerShell-safe types for JSON serialization.
    #>
    param([AllowNull()]$Value)

    if ($null -eq $Value)               { return $null }
    if ($Value -is [System.DBNull])     { return $null }
    if ($Value -is [System.DateTime])   { return $Value.ToString('o') }  # ISO 8601
    if ($Value -is [decimal])           { return [double]$Value }
    if ($Value -is [byte[]])            { return "<binary $($Value.Length) bytes>" }
    return $Value
}

function ConvertTo-CoercedProp {
    <#
    .SYNOPSIS
        Convert string property values to int/bool as needed for COM.
    #>
    param($Value)

    if ($Value -is [int] -or $Value -is [double] -or $Value -is [float] -or $Value -is [bool]) {
        return $Value
    }
    if ($Value -is [string]) {
        $low = $Value.ToLower()
        if ($low -in 'true', 'yes', '-1')  { return $true }
        if ($low -in 'false', 'no', '0')   { return $false }
        $intVal = 0
        if ([int]::TryParse($Value, [ref]$intVal))    { return $intVal }
        $dblVal = 0.0
        if ([double]::TryParse($Value, [ref]$dblVal))  { return $dblVal }
    }
    return $Value
}

function Format-AccessOutput {
    <#
    .SYNOPSIS
        Handle -AsJson switch: convert hashtable/PSCustomObject to JSON or return as-is.
    #>
    param(
        [Parameter(Mandatory)]$Data,
        [switch]$AsJson
    )

    if ($Data -is [hashtable]) {
        $Data = [PSCustomObject]$Data
    }
    if ($AsJson) {
        return $Data | ConvertTo-Json -Depth 10 -Compress
    }
    return $Data
}

function Read-TempFile {
    <#
    .SYNOPSIS
        Read a file exported by Access. Auto-detects encoding (UTF-16 BOM, UTF-8-sig, cp1252).
        Returns [PSCustomObject]@{ Content = [string]; Encoding = [string] }
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    # Check BOM
    $bom = [byte[]]::new(2)
    $fs = [System.IO.File]::OpenRead($Path)
    try {
        $null = $fs.Read($bom, 0, 2)
    } finally {
        $fs.Close()
    }

    # UTF-16 LE or BE BOM
    if (($bom[0] -eq 0xFF -and $bom[1] -eq 0xFE) -or ($bom[0] -eq 0xFE -and $bom[1] -eq 0xFF)) {
        $content = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::Unicode)
        return [PSCustomObject]@{ Content = $content; Encoding = 'utf-16' }
    }

    # Try UTF-8 with BOM
    try {
        $utf8Bom = New-Object System.Text.UTF8Encoding($true, $true)  # throwOnInvalid
        $content = [System.IO.File]::ReadAllText($Path, $utf8Bom)
        return [PSCustomObject]@{ Content = $content; Encoding = 'utf-8-sig' }
    } catch {}

    # Try Windows-1252 (cp1252) — Access default for ANSI modules
    try {
        $cp1252 = [System.Text.Encoding]::GetEncoding(1252)
        $content = [System.IO.File]::ReadAllText($Path, $cp1252)
        return [PSCustomObject]@{ Content = $content; Encoding = 'cp1252' }
    } catch {}

    # Fallback: UTF-8 with replacement
    $content = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
    return [PSCustomObject]@{ Content = $content; Encoding = 'utf-8' }
}

function Write-TempFile {
    <#
    .SYNOPSIS
        Write content for Access to read with LoadFromText.
        Default utf-16 (Access .accdb expects UTF-16LE with BOM).
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]
        [string]$Content,
        [string]$Encoding = 'utf-16'
    )

    switch ($Encoding) {
        'utf-16' {
            [System.IO.File]::WriteAllText($Path, $Content, [System.Text.Encoding]::Unicode)
        }
        'cp1252' {
            $cp1252 = [System.Text.Encoding]::GetEncoding(1252)
            [System.IO.File]::WriteAllText($Path, $Content, $cp1252)
        }
        default {
            [System.IO.File]::WriteAllText($Path, $Content, [System.Text.Encoding]::UTF8)
        }
    }
}
