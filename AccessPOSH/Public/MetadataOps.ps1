# Public/MetadataOps.ps1 — Linked tables, relationships, references, queries, startup, db properties, tips

function Get-AccessLinkedTable {
    <#
    .SYNOPSIS
        List all linked (attached) tables in the database.
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Get-AccessLinkedTable'
    $app = Connect-AccessDB -DbPath $DbPath
    $db  = $app.CurrentDb()

    $linked = [System.Collections.Generic.List[object]]::new()
    try {
        for ($i = 0; $i -lt $db.TableDefs.Count; $i++) {
            $td   = $db.TableDefs($i)
            $conn = $td.Connect
            if ([string]::IsNullOrEmpty($conn)) {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($td)
                continue
            }

            $name = $td.Name
            if ($name.StartsWith('~') -or $name.StartsWith('MSys')) {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($td)
                continue
            }

            $linked.Add([PSCustomObject][ordered]@{
                name           = $name
                source_table   = $td.SourceTableName
                connect_string = $conn
                is_odbc        = $conn.ToUpper().StartsWith('ODBC;')
            })
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($td)
        }
    } finally {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($db)
    }

    $result = [ordered]@{
        count         = $linked.Count
        linked_tables = @($linked)
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Set-AccessLinkedTable {
    <#
    .SYNOPSIS
        Relink a linked table (or all tables sharing the same connection) to a new data source.
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [ValidateNotNullOrEmpty()]
        [string]$TableName,
        [string]$NewConnect,
        [switch]$RelinkAll,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Set-AccessLinkedTable'
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
    if (-not $NewConnect) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-NewConnect is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $NewConnect
            )
        )
    }
    $app = Connect-AccessDB -DbPath $DbPath
    $db  = $app.CurrentDb()

    # Verify the reference table is actually linked
    $refTd = $db.TableDefs($TableName)
    if ([string]::IsNullOrEmpty($refTd.Connect)) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new("'$TableName' is not a linked table."),
                'InvalidArgument',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $TableName
            )
        )
    }

    $relinked = [System.Collections.Generic.List[object]]::new()

    $relinkOne = {
        param([string]$tName, [string]$oldConn)
        $t = $db.TableDefs($tName)
        $t.Connect = $NewConnect
        $t.RefreshLink()
        $relinked.Add([PSCustomObject][ordered]@{
            name        = $tName
            old_connect = $oldConn
            new_connect = $NewConnect
        })
    }

    if ($RelinkAll) {
        $oldConnect = $refTd.Connect
        $namesToRelink = [System.Collections.Generic.List[object]]::new()
        for ($i = 0; $i -lt $db.TableDefs.Count; $i++) {
            $td = $db.TableDefs($i)
            if ($td.Connect -eq $oldConnect) {
                $namesToRelink.Add([PSCustomObject]@{ Name = $td.Name; Connect = $td.Connect })
            }
        }
        foreach ($item in $namesToRelink) {
            & $relinkOne $item.Name $item.Connect
        }
    } else {
        & $relinkOne $TableName $refTd.Connect
    }

    $result = [ordered]@{
        relinked_count = $relinked.Count
        tables         = @($relinked)
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Get-AccessRelationship {
    <#
    .SYNOPSIS
        List all non-system relationships in the database.
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Get-AccessRelationship'
    $app = Connect-AccessDB -DbPath $DbPath
    $db  = $app.CurrentDb()
    $rels = [System.Collections.Generic.List[object]]::new()

    for ($i = 0; $i -lt $db.Relations.Count; $i++) {
        $rel  = $db.Relations($i)
        $name = $rel.Name
        if ($name.StartsWith('MSys')) { continue }

        $fields = [System.Collections.Generic.List[object]]::new()
        for ($j = 0; $j -lt $rel.Fields.Count; $j++) {
            $fld = $rel.Fields($j)
            $fields.Add([ordered]@{ local = $fld.Name; foreign = $fld.ForeignName })
        }

        $attrs     = [int]$rel.Attributes
        $attrFlags = foreach ($bit in $script:REL_ATTR.Keys) {
            if ($attrs -band $bit) { $script:REL_ATTR[$bit] }
        }
        if ($null -eq $attrFlags) { $attrFlags = @() }

        $rels.Add([PSCustomObject][ordered]@{
            name            = $name
            table           = $rel.Table
            foreign_table   = $rel.ForeignTable
            fields          = @($fields)
            attributes      = $attrs
            attribute_flags = @($attrFlags)
        })
    }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        count         = $rels.Count
        relationships = @($rels)
    })
}

function New-AccessRelationship {
    <#
    .SYNOPSIS
        Create a new relationship between two tables.
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [ValidateNotNullOrEmpty()]
        [string]$Name,
        [string]$Table,
        [string]$ForeignTable,
        [array]$Fields,
        [int]$Attributes = 0,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'New-AccessRelationship'
    if (-not $Name) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-Name is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $Name
            )
        )
    }
    if (-not $Table) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-Table is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $Table
            )
        )
    }
    if (-not $ForeignTable) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-ForeignTable is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $ForeignTable
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

    $rel = $db.CreateRelation($Name, $Table, $ForeignTable, $Attributes)
    try {
        foreach ($fmap in $Fields) {
            $localName   = $fmap['local']
            $foreignName = $fmap['foreign']
            if (-not $localName -or -not $foreignName) {
                $PSCmdlet.ThrowTerminatingError(
                    [System.Management.Automation.ErrorRecord]::new(
                        [System.ArgumentException]::new("Each field mapping must have 'local' and 'foreign' keys."),
                        'InvalidArgument',
                        [System.Management.Automation.ErrorCategory]::InvalidArgument,
                        $fmap
                    )
                )
            }
            $fld = $rel.CreateField($localName)
            $fld.ForeignName = $foreignName
            $rel.Fields.Append($fld)
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($fld)
        }
        if ($PSCmdlet.ShouldProcess("$Table → $ForeignTable", 'Create relationship')) {
            $db.Relations.Append($rel)
        }
    } finally {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($rel)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($db)
    }

    $attrFlags = foreach ($bit in $script:REL_ATTR.Keys) {
        if ($Attributes -band $bit) { $script:REL_ATTR[$bit] }
    }
    if ($null -eq $attrFlags) { $attrFlags = @() }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        name            = $Name
        table           = $Table
        foreign_table   = $ForeignTable
        fields          = @($Fields)
        attributes      = $Attributes
        attribute_flags = @($attrFlags)
        status          = 'created'
    })
}

function Remove-AccessRelationship {
    <#
    .SYNOPSIS
        Delete a relationship by name.
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [ValidateNotNullOrEmpty()]
        [string]$Name,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Remove-AccessRelationship'
    if (-not $Name) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-Name is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $Name
            )
        )
    }
    $app = Connect-AccessDB -DbPath $DbPath
    if ($PSCmdlet.ShouldProcess("relationship '$Name'", 'Remove relationship')) {
        $db = $app.CurrentDb()
        try {
            $db.Relations.Delete($Name)
        } finally {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($db)
        }

        Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
            action = 'deleted'
            name   = $Name
        })
    }
}

function Get-AccessReference {
    <#
    .SYNOPSIS
        List all VBA project references in the database.
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Get-AccessReference'
    $app     = Connect-AccessDB -DbPath $DbPath
    $refsCol = $app.VBE.ActiveVBProject.References
    $refs    = [System.Collections.Generic.List[object]]::new()

    for ($i = 1; $i -le $refsCol.Count; $i++) {
        $ref = $refsCol.Item($i)

        $isBroken = $true
        try { $isBroken = [bool]$ref.IsBroken } catch {}

        $builtIn = $false
        try { $builtIn = [bool]$ref.BuiltIn } catch {}

        $guid = ''
        try { if ($ref.GUID) { $guid = $ref.GUID } } catch {}

        $refs.Add([PSCustomObject][ordered]@{
            name        = $ref.Name
            description = $ref.Description
            full_path   = $ref.FullPath
            guid        = $guid
            major       = [int]$ref.Major
            minor       = [int]$ref.Minor
            is_broken   = $isBroken
            built_in    = $builtIn
        })
    }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        count      = $refs.Count
        references = @($refs)
    })
}

function Set-AccessReference {
    <#
    .SYNOPSIS
        Add or remove a VBA project reference.
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [ValidateSet('add','remove')][string]$Action,
        [string]$Name,
        [string]$RefPath,
        [string]$Guid,
        [int]$Major = 0,
        [int]$Minor = 0,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Set-AccessReference'
    if (-not $Action) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-Action is required (add, remove).'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $Action
            )
        )
    }
    $app  = Connect-AccessDB -DbPath $DbPath
    $refs = $app.VBE.ActiveVBProject.References

    if ($Action -eq 'add') {
        if ($Guid) {
            $ref    = $refs.AddFromGuid($Guid, $Major, $Minor)
            $result = [ordered]@{
                action = 'added'; name = $ref.Name; guid = $Guid; major = $Major; minor = $Minor
            }
        } elseif ($RefPath) {
            $ref    = $refs.AddFromFile($RefPath)
            $result = [ordered]@{
                action = 'added'; name = $ref.Name; full_path = $RefPath
            }
        } else {
            $PSCmdlet.ThrowTerminatingError(
                [System.Management.Automation.ErrorRecord]::new(
                    [System.ArgumentException]::new("Action 'add' requires either -Guid or -RefPath."),
                    'MissingRequiredParameter',
                    [System.Management.Automation.ErrorCategory]::InvalidArgument,
                    $Action
                )
            )
        }
    } else {
        # remove
        if (-not $Name) {
            $PSCmdlet.ThrowTerminatingError(
                [System.Management.Automation.ErrorRecord]::new(
                    [System.ArgumentException]::new("Action 'remove' requires -Name."),
                    'MissingRequiredParameter',
                    [System.Management.Automation.ErrorCategory]::InvalidArgument,
                    $Name
                )
            )
        }
        $found = $null
        for ($i = 1; $i -le $refs.Count; $i++) {
            $ref = $refs.Item($i)
            if ($ref.Name -ieq $Name) { $found = $ref; break }
        }
        if ($null -eq $found) {
            $PSCmdlet.ThrowTerminatingError(
                [System.Management.Automation.ErrorRecord]::new(
                    [System.IO.FileNotFoundException]::new("Reference '$Name' not found."),
                    'ObjectNotFound',
                    [System.Management.Automation.ErrorCategory]::ObjectNotFound,
                    $Name
                )
            )
        }
        try {
            if ($found.BuiltIn) {
                $PSCmdlet.ThrowTerminatingError(
                    [System.Management.Automation.ErrorRecord]::new(
                        [System.InvalidOperationException]::new("Cannot remove built-in reference '$Name'."),
                        'InvalidOperation',
                        [System.Management.Automation.ErrorCategory]::InvalidOperation,
                        $found
                    )
                )
            }
        } catch [System.Management.Automation.PropertyNotFoundException] {}
        $refs.Remove($found)
        $result = [ordered]@{ action = 'removed'; name = $Name }
    }

    # Clear VBE caches
    $script:AccessSession.VbeCodeCache = @{}
    $script:AccessSession.CmCache     = @{}

    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Set-AccessQuery {
    <#
    .SYNOPSIS
        Create, modify, delete, rename, or retrieve SQL for an Access query.
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [ValidateSet('create','modify','delete','rename','get_sql')][string]$Action,
        [string]$QueryName,
        [string]$Sql,
        [string]$NewName,
        [switch]$ConfirmDelete,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Set-AccessQuery'
    if (-not $Action) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-Action is required (create, modify, delete, rename, get_sql).'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $Action
            )
        )
    }
    if (-not $QueryName) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-QueryName is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $QueryName
            )
        )
    }
    $app = Connect-AccessDB -DbPath $DbPath
    $db  = $app.CurrentDb()

    switch ($Action) {
        'create' {
            if (-not $Sql) {
                $PSCmdlet.ThrowTerminatingError(
                    [System.Management.Automation.ErrorRecord]::new(
                        [System.ArgumentException]::new("'create' action requires -Sql."),
                        'MissingRequiredParameter',
                        [System.Management.Automation.ErrorCategory]::InvalidArgument,
                        $Sql
                    )
                )
            }
            $null = $db.CreateQueryDef($QueryName, $Sql)
            $result = [ordered]@{ action = 'created'; query_name = $QueryName; sql = $Sql }
        }
        'modify' {
            if (-not $Sql) {
                $PSCmdlet.ThrowTerminatingError(
                    [System.Management.Automation.ErrorRecord]::new(
                        [System.ArgumentException]::new("'modify' action requires -Sql."),
                        'MissingRequiredParameter',
                        [System.Management.Automation.ErrorCategory]::InvalidArgument,
                        $Sql
                    )
                )
            }
            $qd = $db.QueryDefs($QueryName)
            $qd.SQL = $Sql
            $result = [ordered]@{ action = 'modified'; query_name = $QueryName; sql = $Sql }
        }
        'delete' {
            if (-not $ConfirmDelete) {
                $result = [ordered]@{ error = "Deleting query '$QueryName' requires -ConfirmDelete" }
            } else {
                $null = $db.QueryDefs($QueryName)   # verify exists
                $db.QueryDefs.Delete($QueryName)
                $result = [ordered]@{ action = 'deleted'; query_name = $QueryName }
            }
        }
        'rename' {
            if (-not $NewName) {
                $PSCmdlet.ThrowTerminatingError(
                    [System.Management.Automation.ErrorRecord]::new(
                        [System.ArgumentException]::new("'rename' action requires -NewName."),
                        'MissingRequiredParameter',
                        [System.Management.Automation.ErrorCategory]::InvalidArgument,
                        $NewName
                    )
                )
            }
            $qd = $db.QueryDefs($QueryName)
            $qd.Name = $NewName
            $result = [ordered]@{ action = 'renamed'; old_name = $QueryName; new_name = $NewName }
        }
        'get_sql' {
            $qd = $db.QueryDefs($QueryName)
            $qdType = $script:QUERYDEF_TYPE[[int]$qd.Type]
            if (-not $qdType) { $qdType = "Unknown($($qd.Type))" }
            $result = [ordered]@{ query_name = $QueryName; sql = $qd.SQL; type = $qdType }
        }
    }

    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Get-AccessStartupOption {
    <#
    .SYNOPSIS
        List Access startup/application options from database properties and application settings.
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Get-AccessStartupOption'
    $app = Connect-AccessDB -DbPath $DbPath
    $db  = $app.CurrentDb()

    $options = [System.Collections.Generic.List[object]]::new()

    foreach ($name in $script:STARTUP_PROPS) {
        $val    = $null
        $source = '<not set>'

        try {
            $val    = $db.Properties($name).Value
            $source = 'database'
        } catch {
            try {
                $val    = $app.GetOption($name)
                $source = 'application'
            } catch {}
        }

        $options.Add([ordered]@{
            name   = $name
            value  = $val
            source = $source
        })
    }

    Format-AccessOutput -AsJson:$AsJson -Data ([ordered]@{
        count   = $options.Count
        options = @($options)
    })
}

function Get-AccessDatabaseProperty {
    <#
    .SYNOPSIS
        Read a database property from CurrentDb().Properties or Application.GetOption.
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [ValidateNotNullOrEmpty()]
        [string]$Name,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Get-AccessDatabaseProperty'
    if (-not $Name) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-Name is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $Name
            )
        )
    }
    $app = Connect-AccessDB -DbPath $DbPath
    $db  = $app.CurrentDb()

    try {
        $val = $db.Properties($Name).Value
        $result = [ordered]@{ name = $Name; value = $val; source = 'database' }
        return (Format-AccessOutput -AsJson:$AsJson -Data $result)
    } catch {}

    try {
        $val = $app.GetOption($Name)
        $result = [ordered]@{ name = $Name; value = $val; source = 'application' }
        return (Format-AccessOutput -AsJson:$AsJson -Data $result)
    } catch {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.IO.FileNotFoundException]::new("Property '$Name' not found in CurrentDb().Properties or Application.GetOption."),
                'ObjectNotFound',
                [System.Management.Automation.ErrorCategory]::ObjectNotFound,
                $Name
            )
        )
    }
}

function Set-AccessDatabaseProperty {
    <#
    .SYNOPSIS
        Set or create a database property in CurrentDb().Properties or Application.SetOption.
    #>
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,
        [ValidateNotNullOrEmpty()]
        [string]$Name,
        $Value,
        [int]$PropType = -1,
        [switch]$AsJson
    )

    $DbPath = Resolve-SessionDbPath -DbPath $DbPath -CallerName 'Set-AccessDatabaseProperty'
    if (-not $Name) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-Name is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $Name
            )
        )
    }
    if (-not $PSBoundParameters.ContainsKey('Value')) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new('-Value is required.'),
                'MissingRequiredParameter',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $PSBoundParameters
            )
        )
    }
    $app     = Connect-AccessDB -DbPath $DbPath
    $db      = $app.CurrentDb()
    $coerced = ConvertTo-CoercedProp -Value $Value

    try {
        $db.Properties($Name).Value = $coerced
        $result = [ordered]@{ name = $Name; value = $coerced; source = 'database'; action = 'updated' }
        return (Format-AccessOutput -AsJson:$AsJson -Data $result)
    } catch {}

    try {
        $app.SetOption($Name, $coerced)
        $result = [ordered]@{ name = $Name; value = $coerced; source = 'application'; action = 'updated' }
        return (Format-AccessOutput -AsJson:$AsJson -Data $result)
    } catch {}

    # Create new database property
    if ($PropType -eq -1) {
        if ($coerced -is [bool])                          { $PropType = 1 }
        elseif ($coerced -is [int] -or $coerced -is [long]) { $PropType = 4 }
        else                                               { $PropType = 10 }
    }

    $prop = $db.CreateProperty($Name, $PropType, $coerced)
    $db.Properties.Append($prop)

    $result = [ordered]@{ name = $Name; value = $coerced; source = 'database'; action = 'created' }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Get-AccessTip {
    <#
    .SYNOPSIS
        Return tips and gotchas by topic for working with Access.
    #>
    [CmdletBinding()]
    param(
        [string]$Topic,
        [switch]$AsJson
    )

    $tips = [ordered]@{
        eval = @"
Invoke-AccessEval can query the Access Object Model without new tools:
  Application.IsCompiled - check if VBA is compiled
  SysCmd(10, 2, "formName") - check if form is open
  Application.BrokenReference - True if any ref is broken
  Screen.ActiveForm.Name / Screen.ActiveControl.Name - active form/control
  Forms.Count - number of open forms
  TempVars("x") - session-persistent variables
  DLookup/DCount/DSum - domain aggregate functions
  TypeName(expr) - inspect type
  Eval only works for expressions/functions, NOT statements/Subs.
"@
        controls = @"
Control types for New-AccessControl:
  119 = acCustomControl (ActiveX) - use ClassName for ProgID
  128 = acWebBrowser (native, NOT ActiveX)
  Common: 100=Label, 109=TextBox, 106=ComboBox, 105=ListBox, 104=CommandButton,
          110=CheckBox, 114=SubForm, 122=Image, 101=Rectangle

  FormatConditions: Get-AccessControl / Get-AccessControlDetail show
  format_conditions count. Use VBA via Invoke-AccessVba to read/modify details.
"@
        gotchas = @"
COM & ODBC:
  dbSeeChanges (512) - REQUIRED for DELETE/UPDATE on ODBC linked tables
  LIKE wildcards - use % for ODBC (not *)
  ListBox.Value - use .Column(0) explicitly
  dbAttachSavePWD = 131072 (NOT 65536)
  Multiple JOINs - Access requires nested parentheses

VBA:
  Str() adds leading space - use CStr()
  IIf() evaluates ALL three args (not short-circuit) - use If/Then/Else
  Dim X As New ClassName in a loop only creates ONE instance
  Chr(128) truncates MsgBox - use ChrW(8364) for euro
"@
        sql = @"
Jet SQL DDL:
  YESNO is not valid - use BIT
  DEFAULT not supported in CREATE TABLE - use Set-AccessFieldProperty
  AUTOINCREMENT works as a type
  Use SHORT instead of SMALLINT, LONG instead of INT
  Prefer New-AccessTable over CREATE TABLE SQL

ODBC pass-through:
  QueryDef.Connect limit 255 chars
"@
        vbe = @"
VBE line numbers are 1-based.
ProcCountLines can inflate the last proc count past end - always clamp.
Access must be Visible=True for VBE COM access.
'Trust access to the VBA project object model' must be enabled.
After design operations, close form before accessing VBE CodeModule.
"@
        compile = @"
Test-AccessVbaCompile tips:
  RunCommand(126) shows MsgBox on error - use timeout param.
  Before compiling: Eval('Application.BrokenReference') for broken refs.
  After error: use Get-AccessVbeLine to read problematic code.
"@
        design = @"
Design view + VBE conflict:
  After design ops, form may remain open in Design view.
  Set-AccessVbeProc closes the form (acSaveYes) before VBE access.
  All design operations invalidate caches.

SaveAsText encoding:
  Modules (.bas) - cp1252 (ANSI, no BOM)
  Forms/reports - utf-16 (UTF-16LE with BOM)
"@
    }

    if (-not $Topic -or $Topic.Trim() -eq '') {
        $result = [ordered]@{
            topics = @($tips.Keys)
            hint   = 'Pass -Topic <name> for details. Fuzzy matching supported.'
        }
        return (Format-AccessOutput -AsJson:$AsJson -Data $result)
    }

    $key = $Topic.Trim().ToLower()

    # Exact match
    if ($tips.Contains($key)) {
        $result = [ordered]@{ topic = $key; tip = $tips[$key] }
        return (Format-AccessOutput -AsJson:$AsJson -Data $result)
    }

    # Fuzzy match
    $matched = [System.Collections.Generic.List[object]]::new()
    foreach ($kv in $tips.GetEnumerator()) {
        if ($kv.Key -like "*$key*" -or $kv.Value -like "*$key*") {
            $matched.Add([ordered]@{ topic = $kv.Key; tip = $kv.Value })
        }
    }

    if ($matched.Count -gt 0) {
        if ($matched.Count -eq 1) {
            return (Format-AccessOutput -AsJson:$AsJson -Data $matched[0])
        }
        $result = [ordered]@{ query = $Topic; matches = @($matched) }
        return (Format-AccessOutput -AsJson:$AsJson -Data $result)
    }

    $result = [ordered]@{
        query            = $Topic
        error            = "No tips found matching '$Topic'"
        available_topics = @($tips.Keys)
    }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}
