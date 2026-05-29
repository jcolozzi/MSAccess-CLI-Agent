# Public/TableOps.ps1 — Table structure, field properties, indexes

function Get-AccessTableInfo {
    <#
    .SYNOPSIS
        Get the structure of an Access table: fields, types, sizes, record count, linked info.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER TableName
        Name of the table.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-AccessTableInfo -DbPath "C:\db.accdb" -TableName "Users"
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [ValidateNotNullOrEmpty()]
        [string]$TableName,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Get-AccessTableInfo'
    if (-not $TableName) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-TableName is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $TableName
            )
        )
    }

    $app = Connect-AccessDB -DbPath $DbPath
    $db  = $app.CurrentDb()
    try {
        $td = $db.TableDefs($TableName)
    } catch {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.IO.FileNotFoundException]::new("Table '$TableName' not found: $_"),
                'ObjectNotFound',
                [System.Management.Automation.ErrorCategory]::ObjectNotFound,
                $TableName
            )
        )
    }

    $isLinked = [bool]$td.Connect
    $fields = [System.Collections.Generic.List[object]]::new()

    for ($i = 0; $i -lt $td.Fields.Count; $i++) {
        $fld   = $td.Fields($i)
        $ftype = $fld.Type
        $typeName = $script:DAO_FIELD_TYPE[[int]$ftype]
        if (-not $typeName) { $typeName = "Type$ftype" }

        # AutoNumber detection: Long (4) + dbAutoIncrField attribute (16)
        if ($ftype -eq 4 -and ($fld.Attributes -band $script:DB_AUTO_INCR_FIELD)) {
            $typeName = 'AutoNumber'
        }

        $fields.Add([PSCustomObject][ordered]@{
            name     = $fld.Name
            type     = $typeName
            size     = $fld.Size
            required = [bool]$fld.Required
        })
    }

    # Record count (may fail on linked tables)
    $recordCount = -1
    try {
        $recordCount = $td.RecordCount
        if ($recordCount -eq -1) {
            $rs = $db.OpenRecordset("SELECT COUNT(*) AS cnt FROM [$TableName]")
            $recordCount = $rs.Fields(0).Value
            $rs.Close()
        }
    } catch {}

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        table_name   = $TableName
        fields       = @($fields)
        record_count = $recordCount
        is_linked    = $isLinked
        source_table = if ($isLinked) { $td.SourceTableName } else { '' }
        connect      = if ($isLinked) { $td.Connect } else { '' }
    })
}

function New-AccessTable {
    <#
    .SYNOPSIS
        Create an Access table via DAO with full type support, defaults, descriptions, and primary key.
    .PARAMETER DbPath
        Path to the Access database.
    .PARAMETER TableName
        Name for the new table (must not already exist).
    .PARAMETER Fields
        Array of field definitions: @{ name="ID"; type="autonumber"; primary_key=$true },
        @{ name="Name"; type="text"; size=100; required=$true; default="Unknown"; description="User name" }
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        $fields = @(
            @{ name = "ID"; type = "autonumber"; primary_key = $true }
            @{ name = "Name"; type = "text"; size = 100; required = $true }
        )
        New-AccessTable -DbPath "C:\db.accdb" -TableName "Users" -Fields $fields
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [ValidateNotNullOrEmpty()]
        [string]$TableName,
        [object[]]$Fields,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'New-AccessTable'
    if (-not $TableName) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-TableName is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $TableName
            )
        )
    }
    if (-not $Fields -or $Fields.Count -eq 0) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-Fields is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $Fields
            )
        )
    }

    $app = Connect-AccessDB -DbPath $DbPath
    $db  = $app.CurrentDb()

    # Check table doesn't exist
    $existing = for ($i = 0; $i -lt $db.TableDefs.Count; $i++) { $db.TableDefs($i).Name }
    if ($TableName -in $existing) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new("Table '$TableName' already exists."),
                'ResourceAlreadyExists',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $TableName
            )
        )
    }

    $td = $db.CreateTableDef($TableName)
    $pkFields      = [System.Collections.Generic.List[string]]::new()
    $createdFields = [System.Collections.Generic.List[object]]::new()

    foreach ($fdef in $Fields) {
        $name     = $fdef.name
        $ftype    = ($fdef.type ?? 'text').ToLower()
        $size     = [int]($fdef.size ?? 0)
        $required = [bool]($fdef.required)
        $pk       = [bool]($fdef.primary_key)

        $daoType = $script:FIELD_TYPE_MAP[$ftype]
        if ($null -eq $daoType) {
            $validTypes = ($script:FIELD_TYPE_MAP.Keys | Sort-Object -Unique) -join ', '
            $PSCmdlet.ThrowTerminatingError(
                [System.Management.Automation.ErrorRecord]::new(
                    [System.ArgumentException]::new("Unknown field type: '$ftype'. Valid types: $validTypes"),
                    'InvalidArgument',
                    [System.Management.Automation.ErrorCategory]::InvalidArgument,
                    $ftype
                )
            )
        }

        $isAuto = $ftype -in 'autonumber', 'autoincrement'

        # Text needs size
        if ($daoType -eq 10 -and $size -eq 0) { $size = 255 }

        $fld = if ($size -gt 0) {
            $td.CreateField($name, $daoType, $size)
        } else {
            $td.CreateField($name, $daoType)
        }

        if ($isAuto) {
            $fld.Attributes = $fld.Attributes -bor $script:DB_AUTO_INCR_FIELD
        }

        $fld.Required = $required -or $pk

        $td.Fields.Append($fld)

        if ($pk) { $pkFields.Add($name) }

        $createdFields.Add([PSCustomObject][ordered]@{
            name = $name
            type = $ftype
            size = if ($size -gt 0) { $size } else { $null }
        })
    }

    # Create primary key index
    if ($pkFields.Count -gt 0) {
        $idx = $td.CreateIndex('PrimaryKey')
        $idx.Primary = $true
        $idx.Unique  = $true
        foreach ($pkName in $pkFields) {
            $idxFld = $idx.CreateField($pkName)
            $idx.Fields.Append($idxFld)
        }
        $td.Indexes.Append($idx)
    }

    $db.TableDefs.Append($td)
    $db.TableDefs.Refresh()

    # Set defaults and descriptions via field properties (post-creation)
    foreach ($fdef in $Fields) {
        $name = $fdef.name
        if ($null -ne $fdef.default) {
            try {
                Set-FieldProperty -Db $db -TableName $TableName -FieldName $name -PropertyName 'DefaultValue' -Value ([string]$fdef.default)
            } catch {
                Write-Warning "Error setting default for ${TableName}.${name}: $_"
            }
        }
        if ($null -ne $fdef.description) {
            try {
                Set-FieldProperty -Db $db -TableName $TableName -FieldName $name -PropertyName 'Description' -Value $fdef.description
            } catch {
                Write-Warning "Error setting description for ${TableName}.${name}: $_"
            }
        }
    }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        table_name  = $TableName
        fields      = @($createdFields)
        primary_key = @($pkFields)
        status      = 'created'
    })
}

function Edit-AccessTable {
    <#
    .SYNOPSIS
        Add, delete, or rename fields in an existing table via DAO.
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [ValidateNotNullOrEmpty()]
        [string]$TableName,
        [ValidateSet('add_field','delete_field','rename_field')][string]$Action,
        [string]$FieldName,
        [string]$NewName,
        [string]$FieldType = 'text',
        [int]$Size = 0,
        [switch]$Required,
        $Default,
        [string]$Description,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Edit-AccessTable'
    if (-not $TableName) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-TableName is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $TableName
            )
        )
    }
    if (-not $Action) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-Action is required (add_field, delete_field, rename_field).'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $Action
            )
        )
    }
    if (-not $FieldName) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-FieldName is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $FieldName
            )
        )
    }

    $app = Connect-AccessDB -DbPath $DbPath
    $db  = $app.CurrentDb()
    $td  = $db.TableDefs($TableName)

    switch ($Action) {
        'add_field' {
            $ftype   = $FieldType.ToLower()
            $daoType = $script:FIELD_TYPE_MAP[$ftype]
            if ($null -eq $daoType) {
                $validTypes = ($script:FIELD_TYPE_MAP.Keys | Sort-Object -Unique) -join ', '
                $PSCmdlet.ThrowTerminatingError(
                    [System.Management.Automation.ErrorRecord]::new(
                        [System.ArgumentException]::new("Unknown type: '$ftype'. Valid: $validTypes"),
                        'InvalidArgument',
                        [System.Management.Automation.ErrorCategory]::InvalidArgument,
                        $ftype
                    )
                )
            }
            $isAuto = $ftype -in 'autonumber', 'autoincrement'

            if ($daoType -eq 10 -and $Size -eq 0) { $Size = 255 }

            $fld = if ($Size -gt 0) {
                $td.CreateField($FieldName, $daoType, $Size)
            } else {
                $td.CreateField($FieldName, $daoType)
            }

            if ($isAuto) {
                $fld.Attributes = $fld.Attributes -bor $script:DB_AUTO_INCR_FIELD
            }
            $fld.Required = [bool]$Required

            $td.Fields.Append($fld)
            $td.Fields.Refresh()

            if ($null -ne $Default) {
                try { Set-FieldProperty -Db $db -TableName $TableName -FieldName $FieldName -PropertyName 'DefaultValue' -Value ([string]$Default) } catch {}
            }
            if (-not [string]::IsNullOrEmpty($Description)) {
                try { Set-FieldProperty -Db $db -TableName $TableName -FieldName $FieldName -PropertyName 'Description' -Value $Description } catch {}
            }

            $result = [ordered]@{ action = 'field_added'; table = $TableName; field = $FieldName; type = $ftype }
        }
        'delete_field' {
            if ($PSCmdlet.ShouldProcess("field '$FieldName' in table '$TableName'", 'Delete field')) {
                $td.Fields.Delete($FieldName)
                $result = [ordered]@{ action = 'field_deleted'; table = $TableName; field = $FieldName }
            } else {
                $result = [ordered]@{ action = 'skipped'; reason = 'Declined by ShouldProcess (-WhatIf or -Confirm:$false)' }
            }
        }
        'rename_field' {
            if ([string]::IsNullOrEmpty($NewName)) {
                $PSCmdlet.ThrowTerminatingError(
                    [System.Management.Automation.ErrorRecord]::new(
                        [System.ArgumentException]::new('rename_field requires -NewName'),
                        'MissingRequiredParameter',
                        [System.Management.Automation.ErrorCategory]::InvalidArgument,
                        $NewName
                    )
                )
            }
            $fld = $td.Fields($FieldName)
            $fld.Name = $NewName
            $result = [ordered]@{ action = 'field_renamed'; table = $TableName; old_name = $FieldName; new_name = $NewName }
        }
    }

    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Get-AccessFieldProperty {
    <#
    .SYNOPSIS
        Read all DAO properties from a table field.
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [ValidateNotNullOrEmpty()]
        [string]$TableName,
        [string]$FieldName,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Get-AccessFieldProperty'
    if (-not $TableName) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-TableName is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $TableName
            )
        )
    }
    if (-not $FieldName) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-FieldName is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $FieldName
            )
        )
    }

    $app = Connect-AccessDB -DbPath $DbPath
    $db  = $app.CurrentDb()
    $fld = $db.TableDefs($TableName).Fields($FieldName)

    $props = [ordered]@{}
    for ($i = 0; $i -lt $fld.Properties.Count; $i++) {
        try {
            $p   = $fld.Properties($i)
            $val = $p.Value
            if ($val -is [string] -or $val -is [int] -or $val -is [long] -or
                $val -is [double] -or $val -is [float] -or $val -is [bool] -or $null -eq $val) {
                $props[$p.Name] = $val
            }
        } catch {
            # Skip unreadable properties
        }
    }

    $result = [ordered]@{
        table_name  = $TableName
        field_name  = $FieldName
        properties  = $props
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Set-AccessFieldProperty {
    <#
    .SYNOPSIS
        Set or create a DAO property on a table field.
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [ValidateNotNullOrEmpty()]
        [string]$TableName,
        [string]$FieldName,
        [string]$PropertyName,
        $Value,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Set-AccessFieldProperty'
    if (-not $TableName) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-TableName is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $TableName
            )
        )
    }
    if (-not $FieldName) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-FieldName is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $FieldName
            )
        )
    }
    if (-not $PropertyName) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-PropertyName is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $PropertyName
            )
        )
    }
    if (-not $Value) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-Value is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $Value
            )
        )
    }

    $app     = Connect-AccessDB -DbPath $DbPath
    $db      = $app.CurrentDb()
    $fld     = $db.TableDefs($TableName).Fields($FieldName)
    $coerced = ConvertTo-CoercedProp $Value

    # Try updating existing property first
    $actionTaken = 'updated'
    try {
        $fld.Properties($PropertyName).Value = $coerced
    } catch {
        # Property doesn't exist — create it
        $propType = if ($coerced -is [bool]) { 1 }        # dbBoolean
                    elseif ($coerced -is [int])  { 4 }     # dbLong
                    else                         { 10 }    # dbText

        $prop = $fld.CreateProperty($PropertyName, $propType, $coerced)
        $fld.Properties.Append($prop)
        $actionTaken = 'created'
    }

    $result = [ordered]@{
        table_name    = $TableName
        field_name    = $FieldName
        property_name = $PropertyName
        value         = $coerced
        action        = $actionTaken
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Get-AccessIndex {
    <#
    .SYNOPSIS
        List all indexes on an Access table with field details.
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [ValidateNotNullOrEmpty()]
        [string]$TableName,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Get-AccessIndex'
    if (-not $TableName) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-TableName is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $TableName
            )
        )
    }

    $app = Connect-AccessDB -DbPath $DbPath
    $db  = $app.CurrentDb()
    $td  = $db.TableDefs($TableName)

    $indexes = [System.Collections.Generic.List[object]]::new()

    for ($i = 0; $i -lt $td.Indexes.Count; $i++) {
        $idx = $td.Indexes($i)
        $fields = [System.Collections.Generic.List[object]]::new()

        for ($j = 0; $j -lt $idx.Fields.Count; $j++) {
            $f = $idx.Fields($j)
            $fields.Add([ordered]@{
                name  = $f.Name
                order = if ($f.Attributes -band 1) { 'desc' } else { 'asc' }
            })
        }

        $indexes.Add([ordered]@{
            name    = $idx.Name
            fields  = @($fields)
            primary = [bool]$idx.Primary
            unique  = [bool]$idx.Unique
            foreign = [bool]$idx.Foreign
        })
    }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        table_name = $TableName
        count      = $indexes.Count
        indexes    = @($indexes)
    })
}

function Set-AccessIndex {
    <#
    .SYNOPSIS
        Create or delete an index on an Access table.
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [ValidateNotNullOrEmpty()]
        [string]$TableName,
        [ValidateSet('create','delete')][string]$Action,
        [string]$IndexName,
        [array]$Fields,
        [switch]$Primary,
        [switch]$Unique,
        [switch]$AsJson
    )
    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Set-AccessIndex'
    if (-not $TableName) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-TableName is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $TableName
            )
        )
    }
    if (-not $Action) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-Action is required (create, delete).'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $Action
            )
        )
    }
    if (-not $IndexName) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-IndexName is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $IndexName
            )
        )
    }

    $app = Connect-AccessDB -DbPath $DbPath
    $db  = $app.CurrentDb()
    $td  = $db.TableDefs($TableName)

    switch ($Action) {
        'create' {
            if (-not $Fields -or $Fields.Count -eq 0) {
                $PSCmdlet.ThrowTerminatingError(
                    [System.Management.Automation.ErrorRecord]::new(
                        [System.ArgumentException]::new('create requires -Fields'),
                        'MissingRequiredParameter',
                        [System.Management.Automation.ErrorCategory]::InvalidArgument,
                        $Fields
                    )
                )
            }
            $idx = $td.CreateIndex($IndexName)
            $idx.Primary = [bool]$Primary
            $idx.Unique  = [bool]$Unique

            foreach ($fdef in $Fields) {
                if ($fdef -is [string]) {
                    $fname = $fdef
                    $fld   = $idx.CreateField($fname)
                } else {
                    $fname = $fdef['name']
                    $fld   = $idx.CreateField($fname)
                    if ($fdef.ContainsKey('order') -and $fdef['order'] -eq 'desc') {
                        $fld.Attributes = 1   # dbDescending
                    }
                }
                $idx.Fields.Append($fld)
            }

            $td.Indexes.Append($idx)
            $result = [ordered]@{
                action     = 'created'
                table_name = $TableName
                index_name = $IndexName
                fields     = $Fields
                primary    = [bool]$Primary
                unique     = [bool]$Unique
            }
        }
        'delete' {
            $null = $td.Indexes($IndexName)   # verify exists
            $td.Indexes.Delete($IndexName)
            $result = [ordered]@{
                action     = 'deleted'
                table_name = $TableName
                index_name = $IndexName
            }
        }
    }

    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Search-AccessData {
    <#
    .SYNOPSIS
        Full-text search across Text/Memo fields in local tables.
    .DESCRIPTION
        Searches all Text (type 10) and Memo (type 12) fields in local, non-system
        tables for a substring match using Jet LIKE. Optionally filters by table name
        and supports case-sensitive post-filtering.
    .PARAMETER DbPath
        Path to .accdb/.mdb file. Defaults to the current session database.
    .PARAMETER SearchText
        The text to search for (substring match).
    .PARAMETER Tables
        Optional list of table names to restrict the search to.
    .PARAMETER MaxResultsPerTable
        Maximum matches per table (default 50).
    .PARAMETER MaxResultsTotal
        Maximum matches across all tables (default 500).
    .PARAMETER MatchCase
        If set, filters out rows where the match is case-insensitive only.
    .PARAMETER AsJson
        Return results as JSON string.
    #>
    [CmdletBinding()]
    param(
        [string]$DbPath,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$SearchText,

        [string[]]$Tables,

        [int]$MaxResultsPerTable = 50,

        [int]$MaxResultsTotal = 500,

        [switch]$MatchCase,

        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Search-AccessData'
    $app = Connect-AccessDB -DbPath $DbPath

    # ── excerpt helper ──
    $excerptAround = {
        param([string]$Text, [string]$Needle, [int]$Ctx = 40)
        if (-not $Text -or -not $Needle) { return ($Text ?? '') }
        $idx = $Text.IndexOf($Needle, [System.StringComparison]::OrdinalIgnoreCase)
        if ($idx -lt 0) { return $Text.Substring(0, [Math]::Min($Text.Length, $Ctx * 2)) }
        $start = [Math]::Max(0, $idx - $Ctx)
        $end   = [Math]::Min($Text.Length, $idx + $Needle.Length + $Ctx)
        $prefix = if ($start -gt 0) { '...' } else { '' }
        $suffix = if ($end -lt $Text.Length) { '...' } else { '' }
        return "$prefix$($Text.Substring($start, $end - $start))$suffix"
    }

    $db = $null
    try {
        $db = $app.CurrentDb()
        $safeNeedle = $SearchText.Replace("'", "''")

        $results      = [System.Collections.Generic.List[object]]::new()
        $totalMatches  = 0
        $truncated     = $false

        foreach ($td in $db.TableDefs) {
            # stop early if global cap reached
            if ($totalMatches -ge $MaxResultsTotal) {
                $truncated = $true
                break
            }

            $tName = $td.Name

            # skip system / temp tables
            if ($tName -like 'MSys*' -or $tName -like '~*') { continue }

            # skip linked tables (non-empty Connect)
            if ($td.Connect -and $td.Connect -ne '') { continue }

            # filter to requested tables (case-insensitive)
            if ($Tables -and $Tables.Count -gt 0) {
                $match = $false
                foreach ($t in $Tables) {
                    if ($tName -ieq $t) { $match = $true; break }
                }
                if (-not $match) { continue }
            }

            # collect Text (10) and Memo (12) fields
            $textFields = [System.Collections.Generic.List[string]]::new()
            foreach ($fld in $td.Fields) {
                if ($fld.Type -eq 10 -or $fld.Type -eq 12) {
                    $textFields.Add($fld.Name)
                }
            }
            if ($textFields.Count -eq 0) { continue }

            # build WHERE clause
            $clauses = [System.Collections.Generic.List[string]]::new()
            foreach ($fn in $textFields) {
                $safeFn = $fn.Replace(']', ']]')
                $clauses.Add("[$safeFn] LIKE '*$safeNeedle*'")
            }
            $whereClause = $clauses -join ' OR '

            $safeTable = $tName.Replace(']', ']]')
            $perTableCap = [Math]::Min($MaxResultsPerTable, $MaxResultsTotal - $totalMatches)
            if ($perTableCap -le 0) {
                $truncated = $true
                break
            }

            $sql = "SELECT TOP $($perTableCap + 1) * FROM [$safeTable] WHERE $whereClause"

            $rs = $null
            try {
                $rs = $db.OpenRecordset($sql)
            }
            catch {
                Write-Warning "Search-AccessData: skipping table '$tName': $_"
                continue
            }

            if ($null -eq $rs -or $rs.EOF) {
                if ($null -ne $rs) {
                    $rs.Close()
                    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($rs)
                }
                continue
            }

            # read field names from the recordset
            $fieldNames = [System.Collections.Generic.List[string]]::new()
            for ($i = 0; $i -lt $rs.Fields.Count; $i++) {
                $fieldNames.Add($rs.Fields.Item($i).Name)
            }

            $tableHits = [System.Collections.Generic.List[object]]::new()
            $tableCount = 0

            try {
                while (-not $rs.EOF -and $tableCount -lt $perTableCap) {
                    # read all fields into an ordered dict
                    $row = [ordered]@{}
                    foreach ($fn in $fieldNames) {
                        $val = $rs.Fields.Item($fn).Value
                        $row[$fn] = if ($null -eq $val) { $null } else { $val }
                    }

                    # determine which text fields matched
                    $matchedFields = [System.Collections.Generic.List[string]]::new()
                    $firstExcerpt  = ''

                    foreach ($fn in $textFields) {
                        $val = $row[$fn]
                        if ($null -eq $val) { continue }
                        $strVal = [string]$val

                        if ($MatchCase) {
                            if ($strVal.IndexOf($SearchText, [System.StringComparison]::Ordinal) -lt 0) {
                                continue
                            }
                        }
                        else {
                            if ($strVal.IndexOf($SearchText, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
                                continue
                            }
                        }

                        $matchedFields.Add($fn)
                        if (-not $firstExcerpt) {
                            $firstExcerpt = & $excerptAround $strVal $SearchText
                        }
                    }

                    if ($matchedFields.Count -gt 0) {
                        $row['_matched_fields'] = @($matchedFields)
                        $row['_excerpt']        = $firstExcerpt
                        $tableHits.Add($row)
                        $tableCount++
                    }

                    $rs.MoveNext()
                }
            }
            finally {
                $rs.Close()
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($rs)
            }

            if ($tableHits.Count -gt 0) {
                $tableTruncated = $tableCount -ge $perTableCap
                $results.Add([ordered]@{
                    table     = $tName
                    matches   = @($tableHits)
                    count     = $tableHits.Count
                    truncated = $tableTruncated
                })
                $totalMatches += $tableHits.Count
                if ($totalMatches -ge $MaxResultsTotal) {
                    $truncated = $true
                }
            }
        }
    }
    finally {
        if ($null -ne $db) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($db)
        }
    }

    $output = [ordered]@{
        search_text     = $SearchText
        match_case      = [bool]$MatchCase
        total_matches   = $totalMatches
        tables_with_hits = $results.Count
        truncated       = $truncated
        results         = @($results)
    }

    Format-AccessOutput -AsJson:$AsJson -Data $output
}
