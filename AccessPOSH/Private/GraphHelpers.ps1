# Private/GraphHelpers.ps1 — Graph analysis helpers for Export-AccessGraph
#
# All functions operate on a per-invocation $GraphState hashtable (no $script: pollution).
# Ported from access-graph-starter/Export-AccessGraph.ps1 into AccessPOSH module.

# Built-in VBA / Access function names — excluded from cross-module call detection.
$script:VBA_BUILTIN_NAMES = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@(
        'asc','ascw','chr','chrw','format','instr','instrb','instrrev','join',
        'lcase','left','len','lenb','ltrim','mid','replace','right','space',
        'split','str','strcomp','strconv','strreverse','trim','rtrim','ucase','val','string',
        'cbool','cbyte','ccur','cdate','cdbl','cdec','cint','clng','clnglng','clngptr','csng','cstr','cvar','cverr',
        'isarray','isdate','isempty','iserror','ismissing','isnull','isnumeric','isobject','typename','vartype',
        'abs','atn','cos','exp','fix','int','log','rnd','round','sgn','sin','sqr','tan',
        'date','dateadd','datediff','datepart','dateserial','datevalue','day','formatdatetime',
        'hour','minute','month','monthname','now','second','time','timeserial','timevalue','timer',
        'weekday','weekdayname','year',
        'inputbox','msgbox',
        'curdir','dir','eof','filecopy','filedatetime','filelen','freefile','getattr','loc','lof','setattr',
        'array','erase','filter','lbound','ubound',
        'appactivate','beep','command','doevents','environ','sendkeys','shell',
        'error',
        'callbyname','createobject','getobject',
        'deletesetting','getsetting','savesetting',
        'hex','oct',
        'choose','iif','nz','partition','qbcolor','randomize','rgb',
        'davg','dcount','dfirst','dlast','dlookup','dmax','dmin','dstdev','dstdevp','dsum','dvar','dvarp',
        'codedb','currentdb','currentuser','eval','guidfromstring','hyperlinkpart','stringfromguid','syscmd'
    ),
    [System.StringComparer]::OrdinalIgnoreCase
)

# ──────────────────────────────────────────────────────────────────────
#  Graph State Factory
# ──────────────────────────────────────────────────────────────────────

function New-GraphState {
    <#
    .SYNOPSIS
        Create a fresh per-invocation graph state hashtable.
    #>
    return @{
        Nodes           = New-Object 'System.Collections.Generic.List[object]'
        Edges           = New-Object 'System.Collections.Generic.List[object]'
        NodeIndex       = @{}
        EdgeIndex       = @{}
        NameTargets     = @{}
        DataNameTargets = @{}
        SqlNodeCache    = @{}
        KnownTableFields = @{}
        Warnings        = New-Object 'System.Collections.Generic.List[object]'
        EdgeId          = 0
        ProcIndex       = @{}
        ProcCallRe      = $null
        ModuleCodeCache = @{}
    }
}

# ──────────────────────────────────────────────────────────────────────
#  Utility Helpers
# ──────────────────────────────────────────────────────────────────────

function Get-GraphObjectId {
    param(
        [string]$Group,
        [string]$Name
    )
    return ('{0}:{1}' -f $Group, $Name)
}

function Test-SystemOrTemporaryName {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $true }
    return ($Name -like 'MSys*' -or $Name -like '~*')
}

function Test-SqlText {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
    return ($Text.Trim() -match '^(?is)\s*(SELECT|INSERT|UPDATE|DELETE|TRANSFORM|PARAMETERS|WITH)\b')
}

function Remove-GraphAccessBrackets {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $Name }
    $trimmed = $Name.Trim()
    if ($trimmed.StartsWith('[') -and $trimmed.EndsWith(']') -and $trimmed.Length -ge 2) {
        return $trimmed.Substring(1, $trimmed.Length - 2)
    }
    return $trimmed
}

function Convert-GraphAccessLiteral {
    param([string]$RawValue)
    if ($null -eq $RawValue) { return $null }
    $value = $RawValue.Trim()
    if ($value -eq 'Null') { return $null }
    if ($value.Length -ge 2 -and $value.StartsWith('"') -and $value.EndsWith('"')) {
        return ($value.Substring(1, $value.Length - 2) -replace '""', '"')
    }
    return $value
}

function Get-GraphTextHash {
    param([string]$Text)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes([string]$Text)
        $hash = $sha.ComputeHash($bytes)
        return ([System.BitConverter]::ToString($hash) -replace '-', '').ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
    }
}

function Get-GraphFileHash {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        return @{ hash = $null; size = 0 }
    }
    $hash = Get-FileHash -LiteralPath $Path -Algorithm SHA256
    $item = Get-Item -LiteralPath $Path
    return @{ hash = $hash.Hash.ToLowerInvariant(); size = $item.Length }
}

function Get-GraphPreviewText {
    param(
        [string]$Text,
        [int]$MaxLength = 180
    )
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $flat = ($Text -replace '\s+', ' ').Trim()
    if ($flat.Length -le $MaxLength) { return $flat }
    return $flat.Substring(0, $MaxLength).TrimEnd() + '...'
}

function Get-GraphSafeFileName {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return '_' }
    return ($Name -replace '[\\/:*?"<>|]', '_')
}

function Format-GraphMetaTitle {
    param([System.Collections.IDictionary]$Meta)
    if (-not $Meta -or $Meta.Count -eq 0) { return '' }

    $parts = New-Object 'System.Collections.Generic.List[string]'
    foreach ($key in ($Meta.Keys | Sort-Object)) {
        $value = $Meta[$key]
        if ($null -eq $value) { continue }
        if ($value -is [System.Collections.IEnumerable] -and -not ($value -is [string])) {
            $joined = ($value | ForEach-Object { $_ }) -join ', '
            $parts.Add(("{0}: {1}" -f $key, $joined))
        }
        else {
            $parts.Add(("{0}: {1}" -f $key, $value))
        }
    }
    return ($parts -join "`n")
}

# ──────────────────────────────────────────────────────────────────────
#  Warning Helper
# ──────────────────────────────────────────────────────────────────────

function Add-GraphWarning {
    param(
        [hashtable]$GraphState,
        [string]$Code,
        [string]$Message,
        [hashtable]$Meta = @{}
    )
    $entry = [pscustomobject][ordered]@{
        code    = $Code
        message = $Message
        meta    = $Meta
    }
    $GraphState.Warnings.Add($entry)
    Write-Warning $Message
}

# ──────────────────────────────────────────────────────────────────────
#  Node & Edge Management
# ──────────────────────────────────────────────────────────────────────

function Add-GraphNode {
    param(
        [hashtable]$GraphState,
        [string]$Id,
        [string]$Label,
        [string]$Group,
        [hashtable]$Meta = @{}
    )

    if ($GraphState.NodeIndex.ContainsKey($Id)) {
        $existing = $GraphState.NodeIndex[$Id]
        foreach ($key in $Meta.Keys) {
            $existing.meta[$key] = $Meta[$key]
        }
        if ($Label) { $existing.label = $Label }
        if ($Group) { $existing.group = $Group }
        $existing.title = Format-GraphMetaTitle -Meta $existing.meta
        return $existing
    }

    $metaCopy = [ordered]@{}
    foreach ($key in $Meta.Keys) { $metaCopy[$key] = $Meta[$key] }

    $node = [pscustomobject][ordered]@{
        id    = $Id
        label = $Label
        group = $Group
        title = Format-GraphMetaTitle -Meta $metaCopy
        meta  = $metaCopy
    }

    $GraphState.Nodes.Add($node)
    $GraphState.NodeIndex[$Id] = $node
    return $node
}

function Add-GraphEdge {
    param(
        [hashtable]$GraphState,
        [string]$From,
        [string]$To,
        [string]$Label,
        [string]$Kind,
        [string]$Arrows = 'to',
        [hashtable]$Meta = @{}
    )

    $metaJson = ConvertTo-Json -InputObject $Meta -Depth 8 -Compress
    $edgeKey = ($From, $To, $Kind, $Label, $Arrows, $metaJson) -join '|'
    if ($GraphState.EdgeIndex.ContainsKey($edgeKey)) { return }

    $GraphState.EdgeId += 1
    $metaCopy = [ordered]@{}
    foreach ($key in $Meta.Keys) { $metaCopy[$key] = $Meta[$key] }

    $edge = [pscustomobject][ordered]@{
        id     = ('e{0}' -f $GraphState.EdgeId)
        from   = $From
        to     = $To
        label  = $Label
        kind   = $Kind
        arrows = $Arrows
        title  = Format-GraphMetaTitle -Meta $metaCopy
        meta   = $metaCopy
    }

    $GraphState.Edges.Add($edge)
    $GraphState.EdgeIndex[$edgeKey] = $edge
}

function Register-GraphNameTarget {
    param(
        [hashtable]$GraphState,
        [string]$Name,
        [string]$NodeId,
        [string]$Group,
        [switch]$IsDataObject
    )

    if ([string]::IsNullOrWhiteSpace($Name)) { return }

    if (-not $GraphState.NameTargets.ContainsKey($Name)) {
        $GraphState.NameTargets[$Name] = New-Object 'System.Collections.Generic.List[object]'
    }
    $GraphState.NameTargets[$Name].Add([pscustomobject]@{
        id    = $NodeId
        group = $Group
        name  = $Name
    })

    if ($IsDataObject) {
        $lcName = $Name.ToLowerInvariant()
        if (-not $GraphState.DataNameTargets.ContainsKey($lcName)) {
            $GraphState.DataNameTargets[$lcName] = New-Object 'System.Collections.Generic.List[object]'
        }
        $GraphState.DataNameTargets[$lcName].Add([pscustomobject]@{
            id    = $NodeId
            group = $Group
            name  = $Name
        })
    }
}

function Get-GraphTargetsByName {
    param(
        [string]$Name,
        [hashtable]$TargetTable
    )

    if ([string]::IsNullOrWhiteSpace($Name)) { return @() }

    $lcName = $Name.ToLowerInvariant()
    if ($TargetTable.ContainsKey($lcName)) {
        return $TargetTable[$lcName].ToArray()
    }

    $unbracketed = (Remove-GraphAccessBrackets -Name $Name).ToLowerInvariant()
    if ($TargetTable.ContainsKey($unbracketed)) {
        return $TargetTable[$unbracketed].ToArray()
    }

    return @()
}

# ──────────────────────────────────────────────────────────────────────
#  SaveAsText Helper
# ──────────────────────────────────────────────────────────────────────

function Save-AccessTextObject {
    <#
    .SYNOPSIS
        Wrap Access COM SaveAsText with error handling. Returns path/hash/size.
    #>
    param(
        $App,
        [int]$Type,
        [string]$Name,
        [string]$Folder
    )

    if (-not (Test-Path -LiteralPath $Folder)) {
        $null = New-Item -ItemType Directory -Force -Path $Folder
    }
    $path = Join-Path $Folder ((Get-GraphSafeFileName -Name $Name) + '.txt')

    try {
        $App.SaveAsText($Type, $Name, $path)
        $hashInfo = Get-GraphFileHash -Path $path
        return [pscustomobject]@{
            path = $path
            hash = $hashInfo.hash
            size = $hashInfo.size
            ok   = $true
        }
    }
    catch {
        Write-Warning ("SaveAsText failed for object '{0}': {1}" -f $Name, $_.Exception.Message)
        return [pscustomobject]@{
            path = $null
            hash = $null
            size = 0
            ok   = $false
        }
    }
}

# ──────────────────────────────────────────────────────────────────────
#  Access Export Tree Parser
# ──────────────────────────────────────────────────────────────────────

function ConvertFrom-AccessExportTree {
    <#
    .SYNOPSIS
        Parse a SaveAsText .txt export file into a nested block tree.
        Returns [PSCustomObject]@{ Root; DesignRoot; Code; RawLines }
    #>
    param([string]$Path)

    $lines = Get-Content -LiteralPath $Path

    $root = [pscustomobject]@{
        Type       = 'Root'
        Properties = @{}
        Children   = New-Object 'System.Collections.Generic.List[object]'
    }

    $stack = New-Object System.Collections.Stack
    $stack.Push($root)

    $inBlob = $false
    $inCode = $false
    $codeLines = New-Object 'System.Collections.Generic.List[string]'

    foreach ($line in $lines) {
        if ($inCode) {
            $codeLines.Add($line)
            continue
        }

        if ($inBlob) {
            if ($line.Trim() -eq 'End') { $inBlob = $false }
            continue
        }

        if ($line -match '^\s*CodeBehind(Form|Report)\b') {
            $inCode = $true
            $codeLines.Add($line)
            continue
        }

        if ($line -match '^\s*Begin\s+(.+?)\s*$') {
            $typeName = $matches[1].Trim()
            $block = [pscustomobject]@{
                Type       = $typeName
                Properties = @{}
                Children   = New-Object 'System.Collections.Generic.List[object]'
            }
            $stack.Peek().Children.Add($block)
            $stack.Push($block)
            continue
        }

        # Anonymous Begin blocks (controls container / defaults block)
        if ($line -match '^\s*Begin\s*$') {
            $block = [pscustomobject]@{
                Type       = '_anonymous'
                Properties = @{}
                Children   = New-Object 'System.Collections.Generic.List[object]'
            }
            $stack.Peek().Children.Add($block)
            $stack.Push($block)
            continue
        }

        if ($line -match '^\s*End\s*$') {
            if ($stack.Count -gt 1) { [void]$stack.Pop() }
            continue
        }

        if ($line -match '^\s*([A-Za-z][A-Za-z0-9_]*)\s*=\s*(.+?)\s*$') {
            $propName = $matches[1]
            $rawValue = $matches[2]

            if ($rawValue.Trim() -eq 'Begin') {
                $stack.Peek().Properties[$propName] = '[BLOB]'
                $inBlob = $true
                continue
            }

            $stack.Peek().Properties[$propName] = Convert-GraphAccessLiteral -RawValue $rawValue
            continue
        }
    }

    $designRoot = $null
    foreach ($child in $root.Children) {
        if ($child.Type -match '^(Form|Report)$') {
            $designRoot = $child
            break
        }
    }

    return [pscustomobject]@{
        Root       = $root
        DesignRoot = $designRoot
        Code       = ($codeLines -join [Environment]::NewLine)
        RawLines   = $lines
    }
}

function Get-GraphFlattenedBlock {
    <#
    .SYNOPSIS
        Stack-based BFS to flatten nested block tree.
    #>
    param($Block)

    $results = New-Object 'System.Collections.Generic.List[object]'
    if ($null -eq $Block) { return $results.ToArray() }

    $stack = New-Object System.Collections.Stack
    $stack.Push($Block)
    while ($stack.Count -gt 0) {
        $current = $stack.Pop()
        if ($null -eq $current.Children) { continue }
        $childCount = $current.Children.Count
        for ($i = $childCount - 1; $i -ge 0; $i--) {
            $child = $current.Children[$i]
            $results.Add($child)
            $stack.Push($child)
        }
    }

    return $results.ToArray()
}

function Get-GraphControlBlock {
    <#
    .SYNOPSIS
        Filter flattened blocks to known control types.
    #>
    param($DesignRoot)

    $knownControlTypes = @(
        'textbox', 'combobox', 'listbox', 'checkbox', 'optionbutton',
        'togglebutton', 'boundobjectframe', 'attachment',
        'subform', 'subreport', 'customcontrol'
    )

    $blocks = Get-GraphFlattenedBlock -Block $DesignRoot
    return @(
        $blocks | Where-Object {
            $typeName = ([string]$_.Type).ToLowerInvariant()
            $_.Properties.ContainsKey('ControlSource') -or
            $_.Properties.ContainsKey('SourceObject') -or
            ($knownControlTypes -contains $typeName)
        }
    )
}

# ──────────────────────────────────────────────────────────────────────
#  Reference Detection
# ──────────────────────────────────────────────────────────────────────

function Find-GraphReferencedDataName {
    <#
    .SYNOPSIS
        Regex scan for known table/query names in text. Case-insensitive.
    #>
    param(
        [string]$Text,
        [string[]]$KnownNames
    )

    $hits = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }

    foreach ($name in $KnownNames) {
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $escaped = [regex]::Escape($name)
        $pattern = "(?is)(?<![\w])(?:\[$escaped\]|$escaped)(?![\w])"
        if ([regex]::IsMatch($Text, $pattern)) {
            [void]$hits.Add($name)
        }
    }

    return [string[]]$hits
}

function Get-GraphFieldReferenceFromControlSource {
    param([string]$ControlSource)

    if ([string]::IsNullOrWhiteSpace($ControlSource)) { return $null }
    $trimmed = $ControlSource.Trim()
    if ($trimmed.StartsWith('=')) { return $null }
    if ($trimmed -match '[\+\-\*\/\&\(\)]') { return $null }

    $parts = $trimmed -split '\.'
    $candidate = $parts[$parts.Count - 1].Trim()
    $candidate = Remove-GraphAccessBrackets -Name $candidate

    if ([string]::IsNullOrWhiteSpace($candidate)) { return $null }
    return $candidate
}

# ──────────────────────────────────────────────────────────────────────
#  SQL Node Helpers
# ──────────────────────────────────────────────────────────────────────

function New-GraphSqlNode {
    <#
    .SYNOPSIS
        Create or find deduplicated SQL node by SHA256 hash.
        Origin stored as array for multi-reference traceability.
    #>
    param(
        [hashtable]$GraphState,
        [string]$SqlText,
        [string]$Origin,
        [string]$SqlFolder
    )

    $hash = Get-GraphTextHash -Text $SqlText
    if ($GraphState.SqlNodeCache.ContainsKey($hash)) {
        $existing = $GraphState.SqlNodeCache[$hash]
        # Append origin as array for multi-reference traceability
        if ($existing.meta.origin -is [string]) {
            if ($existing.meta.origin -ne $Origin) {
                $existing.meta.origin = @($existing.meta.origin, $Origin)
            }
        }
        elseif ($existing.meta.origin -is [array]) {
            if ($existing.meta.origin -notcontains $Origin) {
                $existing.meta.origin = @($existing.meta.origin) + @($Origin)
            }
        }
        return $existing
    }

    if ($SqlFolder) {
        if (-not (Test-Path -LiteralPath $SqlFolder)) {
            $null = New-Item -ItemType Directory -Force -Path $SqlFolder
        }
        $path = Join-Path $SqlFolder ($hash + '.sql')
        if (-not (Test-Path -LiteralPath $path)) {
            Set-Content -LiteralPath $path -Value $SqlText -Encoding UTF8
        }
    }
    else {
        $path = $null
    }

    $nodeId = 'sql:' + $hash.Substring(0, 20)
    $node = Add-GraphNode -GraphState $GraphState -Id $nodeId -Label ('SQL ' + $hash.Substring(0, 8)) -Group 'sql' -Meta @{
        origin    = $Origin
        sqlHash   = $hash
        sqlPath   = $path
        sqlLength = $SqlText.Length
        preview   = Get-GraphPreviewText -Text $SqlText
    }

    $GraphState.SqlNodeCache[$hash] = $node
    return $node
}

function Add-GraphSqlReferenceEdge {
    <#
    .SYNOPSIS
        Find all known table/query names in SQL and add edges.
    #>
    param(
        [hashtable]$GraphState,
        [string]$SqlText,
        [string]$FromNodeId,
        [string]$RelationKind,
        [string]$SqlFolder,
        [string[]]$KnownDataNames
    )

    foreach ($name in (Find-GraphReferencedDataName -Text $SqlText -KnownNames $KnownDataNames)) {
        foreach ($target in (Get-GraphTargetsByName -Name $name -TargetTable $GraphState.DataNameTargets)) {
            Add-GraphEdge -GraphState $GraphState -From $FromNodeId -To $target.id -Label 'uses' -Kind $RelationKind -Arrows 'to' -Meta @{ name = $name }
        }
    }
}

# ──────────────────────────────────────────────────────────────────────
#  Field Node Helper
# ──────────────────────────────────────────────────────────────────────

function New-GraphFieldNode {
    param(
        [hashtable]$GraphState,
        [string]$OwnerNodeId,
        [string]$OwnerGroup,
        [string]$OwnerName,
        [string]$FieldName,
        [bool]$Verified,
        [string]$DataType = $null,
        [string]$FieldNodeMode = 'ReferencedOnly'
    )

    if ($FieldNodeMode -eq 'None') { return $null }

    $nodeId = ('field:{0}:{1}:{2}' -f $OwnerGroup, $OwnerName, $FieldName)
    $node = Add-GraphNode -GraphState $GraphState -Id $nodeId -Label $FieldName -Group 'field' -Meta @{
        ownerId    = $OwnerNodeId
        ownerGroup = $OwnerGroup
        ownerName  = $OwnerName
        fieldName  = $FieldName
        verified   = $Verified
        dataType   = $DataType
    }

    Add-GraphEdge -GraphState $GraphState -From $OwnerNodeId -To $nodeId -Label 'field' -Kind 'field-owner' -Arrows 'to' -Meta @{ owner = $OwnerName; field = $FieldName }
    return $node
}

# ──────────────────────────────────────────────────────────────────────
#  Data Source Resolution
# ──────────────────────────────────────────────────────────────────────

function Resolve-GraphDataSource {
    <#
    .SYNOPSIS
        Resolve RecordSource to named target or SQL node.
    #>
    param(
        [hashtable]$GraphState,
        [string]$OwnerNodeId,
        [string]$OwnerGroup,
        [string]$OwnerName,
        [string]$RecordSource,
        [string]$SqlFolder,
        [string[]]$KnownDataNames
    )

    $result = [pscustomobject]@{
        mode      = 'none'
        raw       = $RecordSource
        targetIds = @()
        targetRef = $null
        sqlNodeId = $null
    }

    if ([string]::IsNullOrWhiteSpace($RecordSource)) { return $result }

    $targets = @(Get-GraphTargetsByName -Name $RecordSource -TargetTable $GraphState.DataNameTargets)
    if ($targets.Count -gt 0) {
        foreach ($target in $targets) {
            Add-GraphEdge -GraphState $GraphState -From $OwnerNodeId -To $target.id -Label 'RecordSource' -Kind 'recordsource' -Arrows 'to' -Meta @{ recordSource = $RecordSource }
        }
        $result.mode = 'named'
        $result.targetIds = @($targets | ForEach-Object { $_.id })
        if ($targets.Count -eq 1) { $result.targetRef = $targets[0] }
        return $result
    }

    if (Test-SqlText -Text $RecordSource) {
        $sqlNode = New-GraphSqlNode -GraphState $GraphState -SqlText $RecordSource -Origin ("{0}:{1}:RecordSource" -f $OwnerGroup, $OwnerName) -SqlFolder $SqlFolder
        Add-GraphEdge -GraphState $GraphState -From $OwnerNodeId -To $sqlNode.id -Label 'RecordSource' -Kind 'recordsource-sql' -Arrows 'to' -Meta @{ preview = Get-GraphPreviewText -Text $RecordSource }
        Add-GraphSqlReferenceEdge -GraphState $GraphState -SqlText $RecordSource -FromNodeId $sqlNode.id -RelationKind 'sql-reference' -SqlFolder $SqlFolder -KnownDataNames $KnownDataNames

        $result.mode = 'sql'
        $result.targetIds = @($sqlNode.id)
        $result.sqlNodeId = $sqlNode.id
        $result.targetRef = [pscustomobject]@{ id = $sqlNode.id; group = 'sql'; name = $sqlNode.label }
        return $result
    }

    Add-GraphWarning -GraphState $GraphState -Code 'UnresolvedRecordSource' -Message ("Could not resolve RecordSource '{0}' on {1} '{2}'." -f $RecordSource, $OwnerGroup, $OwnerName) -Meta @{ owner = $OwnerName; group = $OwnerGroup; recordSource = $RecordSource }
    $result.mode = 'unresolved'
    return $result
}

function Resolve-GraphSourceObject {
    <#
    .SYNOPSIS
        Parse SourceObject string like "Form.X", "Report.Y", or bare name.
    #>
    param([string]$SourceObject)

    if ([string]::IsNullOrWhiteSpace($SourceObject)) { return $null }
    $trimmed = $SourceObject.Trim()
    if ($trimmed -match '^(?i)(Form|Report)\.(.+)$') {
        $targetGroup = $matches[1].ToLowerInvariant()
        $targetName = $matches[2]
        return [pscustomobject]@{
            group = $targetGroup
            name  = $targetName
            id    = Get-GraphObjectId -Group $targetGroup -Name $targetName
        }
    }
    return $null
}

# ──────────────────────────────────────────────────────────────────────
#  Form / Report Edge Analysis
# ──────────────────────────────────────────────────────────────────────

function Add-GraphFormReportEdge {
    <#
    .SYNOPSIS
        Parse form/report SaveAsText export. Extract RecordSource, controls, code heuristics.
    #>
    param(
        [hashtable]$GraphState,
        [string]$ObjectGroup,
        [string]$ObjectName,
        [string]$RawPath,
        [string]$SqlFolder,
        [string[]]$KnownDataNames,
        [string]$FieldNodeMode = 'ReferencedOnly',
        [switch]$DisableCodeHeuristics
    )

    if ([string]::IsNullOrWhiteSpace($RawPath) -or -not (Test-Path -LiteralPath $RawPath)) { return }

    $objectId = Get-GraphObjectId -Group $ObjectGroup -Name $ObjectName
    $parse = ConvertFrom-AccessExportTree -Path $RawPath
    $designRoot = $parse.DesignRoot
    if ($null -eq $designRoot) {
        Add-GraphWarning -GraphState $GraphState -Code 'DesignRootMissing' -Message ("No design root found while parsing {0} '{1}'." -f $ObjectGroup, $ObjectName) -Meta @{ path = $RawPath }
        return
    }

    $recordSource = $null
    if ($designRoot.Properties.ContainsKey('RecordSource')) {
        $recordSource = [string]$designRoot.Properties['RecordSource']
    }

    $resolvedRecordSource = Resolve-GraphDataSource -GraphState $GraphState -OwnerNodeId $objectId -OwnerGroup $ObjectGroup -OwnerName $ObjectName -RecordSource $recordSource -SqlFolder $SqlFolder -KnownDataNames $KnownDataNames

    $controlBlocks = Get-GraphControlBlock -DesignRoot $designRoot
    foreach ($control in $controlBlocks) {
        $controlType = [string]$control.Type
        $controlName = if ($control.Properties.ContainsKey('Name')) { [string]$control.Properties['Name'] } else { '' }

        # SourceObject edges (subforms/subreports)
        if ($control.Properties.ContainsKey('SourceObject')) {
            $sourceObject = [string]$control.Properties['SourceObject']
            $target = Resolve-GraphSourceObject -SourceObject $sourceObject
            if ($null -ne $target) {
                $linkMasterFields = if ($control.Properties.ContainsKey('LinkMasterFields')) { [string]$control.Properties['LinkMasterFields'] } else { $null }
                $linkChildFields = if ($control.Properties.ContainsKey('LinkChildFields')) { [string]$control.Properties['LinkChildFields'] } else { $null }

                Add-GraphEdge -GraphState $GraphState -From $objectId -To $target.id -Label 'SourceObject' -Kind 'sourceobject' -Arrows 'to' -Meta @{
                    controlName      = $controlName
                    controlType      = $controlType
                    sourceObject     = $sourceObject
                    linkMasterFields = $linkMasterFields
                    linkChildFields  = $linkChildFields
                }
            }
        }

        # ControlSource edges (field bindings)
        if ($control.Properties.ContainsKey('ControlSource')) {
            $controlSource = [string]$control.Properties['ControlSource']
            $fieldName = Get-GraphFieldReferenceFromControlSource -ControlSource $controlSource

            if ($fieldName -and $null -ne $resolvedRecordSource.targetRef) {
                $ownerRef = $resolvedRecordSource.targetRef
                $verified = $false

                if ($ownerRef.group -eq 'table') {
                    if ($GraphState.KnownTableFields.ContainsKey($ownerRef.name) -and $GraphState.KnownTableFields[$ownerRef.name].Contains($fieldName)) {
                        $verified = $true
                    }
                }

                $fieldNode = New-GraphFieldNode -GraphState $GraphState -OwnerNodeId $ownerRef.id -OwnerGroup $ownerRef.group -OwnerName $ownerRef.name -FieldName $fieldName -Verified $verified -FieldNodeMode $FieldNodeMode
                if ($null -ne $fieldNode) {
                    Add-GraphEdge -GraphState $GraphState -From $objectId -To $fieldNode.id -Label 'ControlSource' -Kind 'controlsource' -Arrows 'to' -Meta @{
                        controlName   = $controlName
                        controlType   = $controlType
                        controlSource = $controlSource
                    }
                }
            }
            elseif ($null -ne $resolvedRecordSource.targetRef) {
                Add-GraphEdge -GraphState $GraphState -From $objectId -To $resolvedRecordSource.targetRef.id -Label 'ControlExpr' -Kind 'control-expression' -Arrows 'to' -Meta @{
                    controlName   = $controlName
                    controlType   = $controlType
                    controlSource = $controlSource
                }
            }
        }

        # RowSource (ComboBox / ListBox data binding)
        if ($control.Properties.ContainsKey('RowSource')) {
            $rsValue = [string]$control.Properties['RowSource']
            if (-not [string]::IsNullOrWhiteSpace($rsValue)) {
                $rsTargets = @(Get-GraphTargetsByName -Name $rsValue -TargetTable $GraphState.DataNameTargets)
                if ($rsTargets.Count -gt 0) {
                    Add-GraphEdge -GraphState $GraphState -From $objectId -To $rsTargets[0].id -Label 'RowSource' -Kind 'rowsource' -Arrows 'to' -Meta @{
                        controlName = $controlName
                        controlType = $controlType
                        rowSource   = $rsValue
                    }
                }
                elseif (Test-SqlText -Text $rsValue) {
                    $rsSqlNode = New-GraphSqlNode -GraphState $GraphState -SqlText $rsValue -Origin ("RowSource:{0}" -f $controlName) -SqlFolder $SqlFolder
                    Add-GraphEdge -GraphState $GraphState -From $objectId -To $rsSqlNode.id -Label 'RowSource' -Kind 'rowsource' -Arrows 'to' -Meta @{
                        controlName = $controlName
                        controlType = $controlType
                    }
                    Add-GraphSqlReferenceEdge -GraphState $GraphState -SqlText $rsValue -FromNodeId $rsSqlNode.id -RelationKind 'sql-reference' -SqlFolder $SqlFolder -KnownDataNames $KnownDataNames
                }
            }
        }
    }

    if (-not $DisableCodeHeuristics) {
        Add-GraphCodeHeuristicEdge -GraphState $GraphState -OwnerNodeId $objectId -OwnerGroup $ObjectGroup -OwnerName $ObjectName -Text $parse.Code -SqlFolder $SqlFolder -KnownDataNames $KnownDataNames
    }
}

# ──────────────────────────────────────────────────────────────────────
#  Code Heuristic Edge Detection
# ──────────────────────────────────────────────────────────────────────

function Add-GraphCodeHeuristicEdge {
    <#
    .SYNOPSIS
        Scan VBA code for DoCmd calls, type refs, data refs. Create edges.
    #>
    param(
        [hashtable]$GraphState,
        [string]$OwnerNodeId,
        [string]$OwnerGroup,
        [string]$OwnerName,
        [string]$Text,
        [string]$SqlFolder,
        [string[]]$KnownDataNames
    )

    if ([string]::IsNullOrWhiteSpace($Text)) { return }

    # DoCmd patterns
    $patterns = @(
        @{ regex = '(?is)\bDoCmd\.OpenForm\s+"(?<name>(?:[^"]|"")+)"';   group = 'form';   label = 'OpenForm';   kind = 'vba-openform' },
        @{ regex = '(?is)\bDoCmd\.OpenReport\s+"(?<name>(?:[^"]|"")+)"'; group = 'report'; label = 'OpenReport'; kind = 'vba-openreport' },
        @{ regex = '(?is)\bDoCmd\.OpenQuery\s+"(?<name>(?:[^"]|"")+)"';  group = 'query';  label = 'OpenQuery';  kind = 'vba-openquery' },
        @{ regex = '(?is)\bDoCmd\.OpenTable\s+"(?<name>(?:[^"]|"")+)"';  group = 'table';  label = 'OpenTable';  kind = 'vba-opentable' },
        @{ regex = '(?is)\bCurrentDb\s*\(\s*\)\s*\.\s*QueryDefs\s*\(\s*"(?<name>(?:[^"]|"")+)"\s*\)'; group = 'query'; label = 'QueryDefs'; kind = 'vba-querydefs' },
        @{ regex = '(?is)\bDBEngine\s*\(\s*0\s*\)\s*\(\s*0\s*\)\s*\.\s*QueryDefs\s*\(\s*"(?<name>(?:[^"]|"")+)"\s*\)'; group = 'query'; label = 'QueryDefs'; kind = 'vba-querydefs' },
        @{ regex = '(?is)\bDoCmd\.RunMacro\s+"(?<name>(?:[^"]|"")+)"'; group = 'macro'; label = 'RunMacro'; kind = 'vba-runmacro' }
    )

    foreach ($pattern in $patterns) {
        $regexMatches = [regex]::Matches($Text, $pattern.regex)
        foreach ($match in $regexMatches) {
            $name = ($match.Groups['name'].Value -replace '""', '"')
            if ([string]::IsNullOrWhiteSpace($name)) { continue }

            $targetId = Get-GraphObjectId -Group $pattern.group -Name $name
            if ($GraphState.NodeIndex.ContainsKey($targetId)) {
                Add-GraphEdge -GraphState $GraphState -From $OwnerNodeId -To $targetId -Label $pattern.label -Kind $pattern.kind -Arrows 'to' -Meta @{ name = $name }
            }
        }
    }

    # DoCmd.RunSQL
    $sqlMatches = [regex]::Matches($Text, '(?is)\bDoCmd\.RunSQL\s+"(?<sql>(?:[^"]|"")+)"')
    foreach ($match in $sqlMatches) {
        $sqlText = ($match.Groups['sql'].Value -replace '""', '"')
        if ([string]::IsNullOrWhiteSpace($sqlText)) { continue }

        $sqlNode = New-GraphSqlNode -GraphState $GraphState -SqlText $sqlText -Origin ("{0}:{1}:VBA" -f $OwnerGroup, $OwnerName) -SqlFolder $SqlFolder
        Add-GraphEdge -GraphState $GraphState -From $OwnerNodeId -To $sqlNode.id -Label 'RunSQL' -Kind 'vba-runsql' -Arrows 'to' -Meta @{ preview = Get-GraphPreviewText -Text $sqlText }
        Add-GraphSqlReferenceEdge -GraphState $GraphState -SqlText $sqlText -FromNodeId $sqlNode.id -RelationKind 'sql-reference' -SqlFolder $SqlFolder -KnownDataNames $KnownDataNames
    }

    # VBA SourceObject assignment
    $soMatches = [regex]::Matches($Text, '(?im)\.SourceObject\s*=\s*"(?<name>(?:[^"]|"")+)"')
    foreach ($match in $soMatches) {
        $soValue = ($match.Groups['name'].Value -replace '""', '"').Trim()
        if ([string]::IsNullOrWhiteSpace($soValue)) { continue }

        $targetId = $null
        if ($soValue -match '^(?i)(Form|Report)\.(.+)$') {
            $targetId = Get-GraphObjectId -Group ($matches[1].ToLowerInvariant()) -Name $matches[2]
        }
        else {
            $tryForm = Get-GraphObjectId -Group 'form' -Name $soValue
            $tryReport = Get-GraphObjectId -Group 'report' -Name $soValue
            if ($GraphState.NodeIndex.ContainsKey($tryForm)) { $targetId = $tryForm }
            elseif ($GraphState.NodeIndex.ContainsKey($tryReport)) { $targetId = $tryReport }
        }

        if ($targetId -and $GraphState.NodeIndex.ContainsKey($targetId)) {
            Add-GraphEdge -GraphState $GraphState -From $OwnerNodeId -To $targetId -Label 'SourceObject' -Kind 'vba-sourceobject' -Arrows 'to' -Meta @{ sourceObject = $soValue }
        }
    }

    # VBA type-dependency edges: Dim/As, New, qualified member access
    $seenTypeEdges = @{}
    foreach ($targetName in $GraphState.NameTargets.Keys) {
        foreach ($target in $GraphState.NameTargets[$targetName]) {
            if ($target.group -ne 'module') { continue }
            if ($target.id -eq $OwnerNodeId) { continue }

            $escaped = [regex]::Escape($targetName)
            $typePattern = "(?im)(?:\bAs\s+$escaped\b|\bNew\s+$escaped\b|\b$escaped\s*\.)"
            if ([regex]::IsMatch($Text, $typePattern)) {
                $edgeKey = "$OwnerNodeId->$($target.id)"
                if (-not $seenTypeEdges.ContainsKey($edgeKey)) {
                    $seenTypeEdges[$edgeKey] = $true
                    Add-GraphEdge -GraphState $GraphState -From $OwnerNodeId -To $target.id -Label 'uses type' -Kind 'vba-type-ref' -Arrows 'to' -Meta @{ name = $targetName }
                }
            }
        }
    }

    # VBA data-reference edges: scan string literals for table/query names
    if ($KnownDataNames.Count -gt 0) {
        $literalMatches = [regex]::Matches($Text, '"((?:[^"]|"")*)"')
        if ($literalMatches.Count -gt 0) {
            $literalText = ($literalMatches | ForEach-Object { $_.Groups[1].Value -replace '""', '"' }) -join ' '
            $seenDataEdges = @{}
            foreach ($name in (Find-GraphReferencedDataName -Text $literalText -KnownNames $KnownDataNames)) {
                foreach ($target in (Get-GraphTargetsByName -Name $name -TargetTable $GraphState.DataNameTargets)) {
                    $edgeKey = "$OwnerNodeId->$($target.id)"
                    if (-not $seenDataEdges.ContainsKey($edgeKey)) {
                        $seenDataEdges[$edgeKey] = $true
                        Add-GraphEdge -GraphState $GraphState -From $OwnerNodeId -To $target.id -Label 'uses data' -Kind 'vba-data-ref' -Arrows 'to' -Meta @{ name = $name }
                    }
                }
            }
        }
    }

    # Cross-module procedure calls (bare FuncName( or Call SubName)
    if ($null -ne $GraphState.ProcCallRe) {
        $seenCallEdges = @{}
        foreach ($match in $GraphState.ProcCallRe.Matches($Text)) {
            $matched = $match.Value
            $procName = ($matched -replace '(?i)^\s*Call\s+', '').TrimEnd('( ')
            $pnameLower = $procName.ToLowerInvariant()
            $targetIds = $GraphState.ProcIndex[$pnameLower]
            if ($null -eq $targetIds) { continue }
            foreach ($tid in $targetIds) {
                if ($tid -eq $OwnerNodeId) { continue }
                $edgeKey = "$OwnerNodeId->$($tid):call:$pnameLower"
                if (-not $seenCallEdges.ContainsKey($edgeKey)) {
                    $seenCallEdges[$edgeKey] = $true
                    Add-GraphEdge -GraphState $GraphState -From $OwnerNodeId -To $tid -Label 'calls' -Kind 'vba-call' -Arrows 'to' -Meta @{ procedure = $procName }
                }
            }
        }
    }
}

# ──────────────────────────────────────────────────────────────────────
#  Macro Heuristic Edge Detection
# ──────────────────────────────────────────────────────────────────────

function Add-GraphMacroHeuristicEdge {
    <#
    .SYNOPSIS
        Parse macro SaveAsText export line-by-line for Action/Argument pairs.
    #>
    param(
        [hashtable]$GraphState,
        [string]$MacroName,
        [string]$RawPath,
        [string]$SqlFolder,
        [string[]]$KnownDataNames
    )

    if ([string]::IsNullOrWhiteSpace($RawPath) -or -not (Test-Path -LiteralPath $RawPath)) { return }

    $macroId = Get-GraphObjectId -Group 'macro' -Name $MacroName
    $lines = Get-Content -LiteralPath $RawPath

    for ($i = 0; $i -lt $lines.Count; $i += 1) {
        $line = $lines[$i]
        if ($line -notmatch '^\s*Action\s*=\s*"?(?<action>[A-Za-z0-9_]+)"?\s*$') { continue }

        $action = $matches['action']
        $argValue = $null
        for ($j = $i + 1; $j -lt [Math]::Min($i + 8, $lines.Count); $j += 1) {
            if ($lines[$j] -match '^\s*Action\s*=') { break }
            if ($lines[$j] -match '^\s*Argument\s*=\s*(.+?)\s*$') {
                $argValue = Convert-GraphAccessLiteral -RawValue $matches[1]
                break
            }
        }

        switch -Regex ($action) {
            '^OpenForm$' {
                if ($argValue) {
                    $targetId = Get-GraphObjectId -Group 'form' -Name $argValue
                    if ($GraphState.NodeIndex.ContainsKey($targetId)) {
                        Add-GraphEdge -GraphState $GraphState -From $macroId -To $targetId -Label 'OpenForm' -Kind 'macro-openform' -Arrows 'to' -Meta @{ name = $argValue }
                    }
                }
            }
            '^OpenReport$' {
                if ($argValue) {
                    $targetId = Get-GraphObjectId -Group 'report' -Name $argValue
                    if ($GraphState.NodeIndex.ContainsKey($targetId)) {
                        Add-GraphEdge -GraphState $GraphState -From $macroId -To $targetId -Label 'OpenReport' -Kind 'macro-openreport' -Arrows 'to' -Meta @{ name = $argValue }
                    }
                }
            }
            '^OpenQuery$' {
                if ($argValue) {
                    $targetId = Get-GraphObjectId -Group 'query' -Name $argValue
                    if ($GraphState.NodeIndex.ContainsKey($targetId)) {
                        Add-GraphEdge -GraphState $GraphState -From $macroId -To $targetId -Label 'OpenQuery' -Kind 'macro-openquery' -Arrows 'to' -Meta @{ name = $argValue }
                    }
                }
            }
            '^OpenTable$' {
                if ($argValue) {
                    $targetId = Get-GraphObjectId -Group 'table' -Name $argValue
                    if ($GraphState.NodeIndex.ContainsKey($targetId)) {
                        Add-GraphEdge -GraphState $GraphState -From $macroId -To $targetId -Label 'OpenTable' -Kind 'macro-opentable' -Arrows 'to' -Meta @{ name = $argValue }
                    }
                }
            }
            '^RunSQL$' {
                if ($argValue) {
                    $sqlNode = New-GraphSqlNode -GraphState $GraphState -SqlText $argValue -Origin ("macro:{0}" -f $MacroName) -SqlFolder $SqlFolder
                    Add-GraphEdge -GraphState $GraphState -From $macroId -To $sqlNode.id -Label 'RunSQL' -Kind 'macro-runsql' -Arrows 'to' -Meta @{ preview = Get-GraphPreviewText -Text $argValue }
                    Add-GraphSqlReferenceEdge -GraphState $GraphState -SqlText $argValue -FromNodeId $sqlNode.id -RelationKind 'sql-reference' -SqlFolder $SqlFolder -KnownDataNames $KnownDataNames
                }
            }
        }
    }
}

# ──────────────────────────────────────────────────────────────────────
#  Viewer Copy Helper
# ──────────────────────────────────────────────────────────────────────

function Copy-GraphViewer {
    <#
    .SYNOPSIS
        Embed graph.json into the vis.js HTML viewer and write index.html.
    #>
    param(
        [string]$DestinationFolder,
        [string]$GraphJson,
        [switch]$Disabled
    )

    if ($Disabled) { return }

    # Resolve viewer from module Resources/ folder
    # When dot-sourced from Private/, $PSScriptRoot is the Private/ folder
    $moduleRoot = Split-Path $PSScriptRoot -Parent
    $viewerSource = Join-Path $moduleRoot 'Resources' 'access-graph-viewer.html'

    if (-not (Test-Path -LiteralPath $viewerSource)) {
        Write-Warning "Graph viewer HTML not found at $viewerSource — skipping embedded viewer."
        return
    }

    $html = Get-Content -LiteralPath $viewerSource -Raw
    $embedTag = "<script>var EMBEDDED_GRAPH = $GraphJson;</script>"
    $html = $html -replace '<!-- EMBED_GRAPH_DATA -->', $embedTag
    Set-Content -LiteralPath (Join-Path $DestinationFolder 'index.html') -Value $html -Encoding UTF8
}

# ──────────────────────────────────────────────────────────────────────
#  Cross-Module Procedure Call Index
# ──────────────────────────────────────────────────────────────────────

function Build-GraphProcIndex {
    <#
    .SYNOPSIS
        Index all public VBA procedures across standalone modules for cross-module
        call detection.  Populates GraphState.ProcIndex and GraphState.ProcCallRe.
    #>
    param(
        [hashtable]$GraphState
    )

    $procDeclRe = '(?im)^\s*(?:Public\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)'
    $privateProcRe = '(?im)^\s*Private\s+(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)'

    foreach ($node in $GraphState.NodeIndex.Values) {
        if ($node.group -ne 'module') { continue }

        $code = $null
        if ($GraphState.ModuleCodeCache.ContainsKey($node.label)) {
            $code = $GraphState.ModuleCodeCache[$node.label]
        } else {
            $rawPath = $null
            if ($null -ne $node.meta) {
                if ($node.meta -is [System.Collections.IDictionary]) {
                    if ($node.meta.ContainsKey('rawPath')) { $rawPath = $node.meta['rawPath'] }
                } elseif ($null -ne $node.meta.rawPath) {
                    $rawPath = $node.meta.rawPath
                }
            }
            if ($rawPath -and (Test-Path -LiteralPath $rawPath)) {
                $code = Get-Content -LiteralPath $rawPath -Raw
                $GraphState.ModuleCodeCache[$node.label] = $code
            }
        }

        if ([string]::IsNullOrWhiteSpace($code)) { continue }

        # Collect private procedure names to exclude
        $privateNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($m in [regex]::Matches($code, $privateProcRe)) {
            [void]$privateNames.Add($m.Groups[1].Value)
        }

        # Index all public procedures
        $nodeId = $node.id
        foreach ($m in [regex]::Matches($code, $procDeclRe)) {
            $procName = $m.Groups[1].Value
            $pnameLower = $procName.ToLowerInvariant()
            if ($privateNames.Contains($procName)) { continue }
            if ($script:VBA_BUILTIN_NAMES.Contains($pnameLower)) { continue }
            if ($pnameLower.Length -lt 2) { continue }

            if (-not $GraphState.ProcIndex.ContainsKey($pnameLower)) {
                $GraphState.ProcIndex[$pnameLower] = New-Object 'System.Collections.Generic.List[string]'
            }
            $GraphState.ProcIndex[$pnameLower].Add($nodeId)
        }
    }

    # Build compiled regex for call detection
    $procNames = @($GraphState.ProcIndex.Keys | Sort-Object { $_.Length } -Descending)
    if ($procNames.Count -gt 0) {
        $escaped = $procNames | ForEach-Object { [regex]::Escape($_) }
        $alt = $escaped -join '|'
        $GraphState.ProcCallRe = [regex]::new(
            "(?<![\.\w])(?:$alt)\s*\(|\bCall\s+(?:$alt)\b",
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
        )
    }
}

# ──────────────────────────────────────────────────────────────────────
#  Graph Query Helpers (ported from mcp_access/graph_query.py)
# ──────────────────────────────────────────────────────────────────────

function ConvertFrom-GraphJson {
    <#
    .SYNOPSIS
        Load a graph.json file and build adjacency lookup tables for querying.
    .DESCRIPTION
        Parses the JSON produced by Export-AccessGraph and returns a PSCustomObject
        with node dictionaries, edge lists, and outgoing/incoming adjacency maps.
        Used internally by Get-AccessGraphQuery.
    .PARAMETER Path
        Path to the graph.json file.
    .OUTPUTS
        PSCustomObject with properties: Nodes, Edges, OutAdj, InAdj, IdLookup, LabelLookup, Meta
    #>
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw [System.IO.FileNotFoundException]::new(
            "graph.json not found at: $Path. Run Export-AccessGraph first.", $Path)
    }

    $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    $data = ConvertFrom-Json $raw

    # Node hashtable: id → node object
    $nodes = @{}
    # Case-insensitive lookups
    $idLookup = @{}          # lowercase id → node
    $labelLookup = @{}       # lowercase label → [List of nodes]
    # Adjacency lists
    $outAdj = @{}            # node id → [List of outgoing edges]
    $inAdj  = @{}            # node id → [List of incoming edges]

    # Index nodes
    $nodeList = @()
    if ($null -ne $data.nodes) { $nodeList = @($data.nodes) }
    foreach ($n in $nodeList) {
        $nid = [string]$n.id
        $nodes[$nid] = $n
        $idLookup[$nid.ToLower()] = $n

        $labelKey = ([string]$n.label).ToLower()
        if (-not $labelLookup.ContainsKey($labelKey)) {
            $labelLookup[$labelKey] = New-Object 'System.Collections.Generic.List[object]'
        }
        $labelLookup[$labelKey].Add($n)
    }

    # Index edges
    $edgeList = @()
    if ($null -ne $data.edges) { $edgeList = @($data.edges) }
    foreach ($e in $edgeList) {
        $fromId = [string]$e.from
        $toId   = [string]$e.to

        if (-not $outAdj.ContainsKey($fromId)) {
            $outAdj[$fromId] = New-Object 'System.Collections.Generic.List[object]'
        }
        $outAdj[$fromId].Add($e)

        if (-not $inAdj.ContainsKey($toId)) {
            $inAdj[$toId] = New-Object 'System.Collections.Generic.List[object]'
        }
        $inAdj[$toId].Add($e)
    }

    $meta = @{}
    if ($null -ne $data.meta) {
        $data.meta.PSObject.Properties | ForEach-Object { $meta[$_.Name] = $_.Value }
    }

    return [PSCustomObject]@{
        Nodes       = $nodes
        Edges       = $edgeList
        OutAdj      = $outAdj
        InAdj       = $inAdj
        IdLookup    = $idLookup
        LabelLookup = $labelLookup
        Meta        = $meta
    }
}

function Resolve-GraphNode {
    <#
    .SYNOPSIS
        Resolve a user-supplied name to graph node(s).
    .DESCRIPTION
        Priority: exact id match → group:name probe → label match.
        Returns an array of matching nodes (may be empty).
    .PARAMETER Graph
        Graph object returned by ConvertFrom-GraphJson.
    .PARAMETER Name
        User-supplied node name (e.g. 'Customers', 'table:Customers', 'frmMain').
    #>
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Graph,

        [Parameter(Mandatory)]
        [string]$Name
    )

    $low = $Name.Trim().ToLower()
    if ([string]::IsNullOrWhiteSpace($low)) { return @() }

    # 1. Exact id match
    if ($Graph.IdLookup.ContainsKey($low)) {
        return @($Graph.IdLookup[$low])
    }

    # 2. Try group:name for all standard groups
    $groups = @('table','query','form','report','macro','module','field','sql')
    foreach ($g in $groups) {
        $candidate = ('{0}:{1}' -f $g, $Name).ToLower()
        if ($Graph.IdLookup.ContainsKey($candidate)) {
            return @($Graph.IdLookup[$candidate])
        }
    }

    # 3. Label match (case-insensitive)
    if ($Graph.LabelLookup.ContainsKey($low)) {
        $hits = $Graph.LabelLookup[$low]
        if ($hits.Count -gt 0) {
            return @($hits)
        }
    }

    return @()
}

function Resolve-GraphInput {
    <#
    .SYNOPSIS
        Resolve -Graph / -GraphPath / -DbPath to a loaded graph object.
    .DESCRIPTION
        Shared resolution logic for all graph query functions. Accepts a pre-loaded
        graph object (from Import-AccessGraph), a file path, or a database path
        (auto-locates access-graph-out/graph.json next to the DB).
    #>
    param(
        [PSCustomObject]$Graph,
        [string]$GraphPath,
        [string]$DbPath
    )

    if ($null -ne $Graph -and $null -ne $Graph.Nodes) {
        return $Graph
    }

    $resolvedPath = $GraphPath
    if (-not $resolvedPath -and $DbPath) {
        $dbDir = Split-Path ([System.IO.Path]::GetFullPath($DbPath)) -Parent
        $candidate = Join-Path $dbDir 'access-graph-out' 'graph.json'
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            $resolvedPath = $candidate
        }
    }

    if (-not $resolvedPath -or -not (Test-Path -LiteralPath $resolvedPath -PathType Leaf)) {
        $searched = if ($resolvedPath) { $resolvedPath } else { '(no path provided)' }
        throw "graph.json not found. Run Export-AccessGraph first. Searched: $searched"
    }

    $result = ConvertFrom-GraphJson -Path $resolvedPath
    if ($null -eq $result.Nodes -or $null -eq $result.Edges -or $null -eq $result.OutAdj -or $null -eq $result.InAdj) {
        throw "Invalid graph structure from '$resolvedPath': missing required properties."
    }
    return $result
}
