# Public/GraphOps.ps1 — Access database dependency graph export

function Export-AccessGraph {
    <#
    .SYNOPSIS
        Build a dependency graph of an Access database and output graph.json + interactive HTML viewer.

    .DESCRIPTION
        Scans tables, queries, forms, reports, macros, and modules in an Access database.
        Discovers edges from RecordSource, ControlSource, SourceObject, DoCmd calls, macro actions,
        SQL references, type dependencies, and data references in VBA string literals.
        Outputs a graph.json file and an interactive vis.js HTML viewer.

    .PARAMETER DbPath
        Path to the Access database (.accdb or .mdb).

    .PARAMETER OutDir
        Output directory for graph.json and index.html. Default: .\access-graph-out

    .PARAMETER FieldNodeMode
        Controls field node creation. None = no field nodes; ReferencedOnly = only fields referenced
        by controls; AllTableFields = all table fields. Default: ReferencedOnly

    .PARAMETER DisableCodeHeuristics
        Skip VBA code analysis (DoCmd patterns, type refs, data refs in string literals).

    .PARAMETER DisableMacroHeuristics
        Skip macro action parsing (OpenForm, OpenReport, etc.).

    .PARAMETER RawExportMode
        Controls raw SaveAsText exports. None = temp files only (cleaned up); Debug = keep raw/ folder.
        Default: None

    .PARAMETER SkipViewerCopy
        Skip generating the embedded HTML viewer (index.html).

    .PARAMETER PassThru
        Return the graph object to the pipeline (in addition to writing files).

    .PARAMETER Force
        Override session safety check when AccessPOSH is connected to a different database.

    .PARAMETER Quiet
        Suppress Write-Progress output.

    .EXAMPLE
        Export-AccessGraph -DbPath C:\MyApp.accdb

    .EXAMPLE
        $graph = Export-AccessGraph -DbPath C:\MyApp.accdb -PassThru
        $graph.nodes | Where-Object { $_.group -eq 'table' }
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$DbPath,

        [string]$OutDir = '.\access-graph-out',

        [ValidateSet('None', 'ReferencedOnly', 'AllTableFields')]
        [string]$FieldNodeMode = 'ReferencedOnly',

        [switch]$DisableCodeHeuristics,

        [switch]$DisableMacroHeuristics,

        [ValidateSet('None', 'Debug')]
        [string]$RawExportMode = 'None',

        [switch]$SkipViewerCopy,

        [switch]$PassThru,

        [switch]$Force,

        [switch]$Quiet
    )

    # ── Constants ──
    $AC_TYPE = @{ Table = 0; Query = 1; Form = 2; Report = 3; Macro = 4; Module = 5 }

    # ── Resolve paths ──
    $resolved = [System.IO.Path]::GetFullPath($DbPath)
    if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.IO.FileNotFoundException]::new("Database file not found: $resolved"),
                'DatabaseNotFound', [System.Management.Automation.ErrorCategory]::ObjectNotFound, $resolved))
    }

    if (-not (Test-Path -LiteralPath $OutDir)) {
        $null = New-Item -ItemType Directory -Force -Path $OutDir
    }
    $OutDir = (Resolve-Path -LiteralPath $OutDir).Path

    # ── Hybrid COM session ──
    $app = $null
    $ownedInstance = $false
    $keepRaw = ($RawExportMode -eq 'Debug')

    if ($null -ne $script:AccessSession -and $null -ne $script:AccessSession.App -and (Test-AccessAlive)) {
        $currentDb = $script:AccessSession.DbPath
        $currentResolved = if ($currentDb) { [System.IO.Path]::GetFullPath($currentDb) } else { '' }

        if ($currentResolved -eq $resolved) {
            # Same DB — reuse session
            $app = $script:AccessSession.App
            Write-Verbose 'Reusing active AccessPOSH session.'
        }
        else {
            if (-not $Force) {
                $PSCmdlet.ThrowTerminatingError(
                    [System.Management.Automation.ErrorRecord]::new(
                        [System.InvalidOperationException]::new(
                            "AccessPOSH session is connected to '$currentDb' but you requested '$resolved'. " +
                            "Use -Force to close the active session and open the new database."),
                        'SessionDbMismatch', [System.Management.Automation.ErrorCategory]::ResourceExists, $resolved))
            }
            Write-Verbose "Force: closing active session ($currentDb) to open $resolved"
            Close-AccessDatabase
            $app = Connect-AccessDB -DbPath $resolved
        }
    }
    else {
        # No active session — create via Connect-AccessDB (module-managed lifecycle)
        $app = Connect-AccessDB -DbPath $resolved
    }

    $db = $app.CurrentDb()

    # ── Initialize graph state ──
    $gs = New-GraphState

    # ── Output directories ──
    $rawDir = Join-Path $OutDir 'raw'
    $sqlFolder = if ($keepRaw) { Join-Path $rawDir 'sql' } else { Join-Path ([System.IO.Path]::GetTempPath()) ('AccessGraph_sql_' + [guid]::NewGuid().ToString('N').Substring(0,8)) }

    if ($keepRaw) {
        foreach ($sub in @($rawDir, (Join-Path $rawDir 'queries'), (Join-Path $rawDir 'forms'), (Join-Path $rawDir 'reports'), (Join-Path $rawDir 'macros'), (Join-Path $rawDir 'modules'), $sqlFolder)) {
            if (-not (Test-Path -LiteralPath $sub)) { $null = New-Item -ItemType Directory -Force -Path $sub }
        }
    }
    else {
        if (-not (Test-Path -LiteralPath $sqlFolder)) { $null = New-Item -ItemType Directory -Force -Path $sqlFolder }
    }

    # Track temp paths for cleanup when RawExportMode=None
    $tempPaths = New-Object 'System.Collections.Generic.List[string]'

    try {
        # ──────────────────────────────────────────────────────────────
        #  SCAN PHASE
        # ──────────────────────────────────────────────────────────────

        # ── Tables ──
        if (-not $Quiet) { Write-Progress -Activity 'Export-AccessGraph' -Status 'Scanning tables...' -PercentComplete 5 }

        foreach ($tableDef in $db.TableDefs) {
            $tdName = $null
            try { $tdName = [string]$tableDef.Name } catch { continue }
            if ([string]::IsNullOrWhiteSpace($tdName) -or (Test-SystemOrTemporaryName -Name $tdName)) { continue }

            $tableName = $tdName
            $tableNodeId = Get-GraphObjectId -Group 'table' -Name $tableName

            # Populate known field set
            if (-not $gs.KnownTableFields.ContainsKey($tableName)) {
                $gs.KnownTableFields[$tableName] = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            }
            try {
                foreach ($field in $tableDef.Fields) {
                    [void]$gs.KnownTableFields[$tableName].Add([string]$field.Name)
                }
            }
            catch {
                Add-GraphWarning -GraphState $gs -Code 'FieldEnumFailed' -Message ("Could not enumerate fields for table '{0}': {1}" -f $tableName, $_.Exception.Message)
            }

            # SaveAsText does not support acTable — skip raw export
            $tblConnect = try { [string]$tableDef.Connect } catch { '' }
            $tblSource  = try { [string]$tableDef.SourceTableName } catch { '' }
            $tblFieldCount = try { $tableDef.Fields.Count } catch { 0 }
            Add-GraphNode -GraphState $gs -Id $tableNodeId -Label $tableName -Group 'table' -Meta @{
                connect    = $tblConnect
                sourceTable = $tblSource
                fieldCount = $tblFieldCount
            } | Out-Null
            Register-GraphNameTarget -GraphState $gs -Name $tableName -NodeId $tableNodeId -Group 'table' -IsDataObject

            if ($FieldNodeMode -eq 'AllTableFields') {
                foreach ($field in $tableDef.Fields) {
                    New-GraphFieldNode -GraphState $gs -OwnerNodeId $tableNodeId -OwnerGroup 'table' -OwnerName $tableName -FieldName ([string]$field.Name) -Verified $true -DataType ([string]$field.Type) -FieldNodeMode $FieldNodeMode | Out-Null
                }
            }
        }

        # ── Relationships ──
        if (-not $Quiet) { Write-Progress -Activity 'Export-AccessGraph' -Status 'Scanning relationships...' -PercentComplete 10 }

        foreach ($relation in $db.Relations) {
            $tblId = Get-GraphObjectId -Group 'table' -Name $relation.Table
            $fTblId = Get-GraphObjectId -Group 'table' -Name $relation.ForeignTable
            if (-not $gs.NodeIndex.ContainsKey($tblId)) { continue }
            if (-not $gs.NodeIndex.ContainsKey($fTblId)) { continue }

            $fieldPairs = New-Object 'System.Collections.Generic.List[string]'
            foreach ($field in $relation.Fields) {
                $fieldPairs.Add(("{0} <-> {1}" -f $field.Name, $field.ForeignName))
            }

            Add-GraphEdge -GraphState $gs -From $fTblId -To $tblId -Label 'relation' -Kind 'relation' -Arrows 'none' -Meta @{
                name   = $relation.Name
                fields = ($fieldPairs -join '; ')
            } | Out-Null
        }

        # ── Queries ──
        if (-not $Quiet) { Write-Progress -Activity 'Export-AccessGraph' -Status 'Scanning queries...' -PercentComplete 15 }
        $queryIdx = 0
        foreach ($queryDef in $db.QueryDefs) {
            $qdName = $null
            try { $qdName = [string]$queryDef.Name } catch { continue }
            if ([string]::IsNullOrWhiteSpace($qdName) -or (Test-SystemOrTemporaryName -Name $qdName)) { continue }

            $queryName = $qdName
            $queryIdx++

            $rawInfo = [pscustomobject]@{ path = $null; hash = $null; size = 0 }
            if ($keepRaw) {
                $rawInfo = Save-AccessTextObject -App $app -Type $AC_TYPE.Query -Name $queryName -Folder (Join-Path $rawDir 'queries')
            }

            $sqlText = try { [string]$queryDef.SQL } catch { '' }
            $queryNodeId = Get-GraphObjectId -Group 'query' -Name $queryName
            $qConnect = try { [string]$queryDef.Connect } catch { '' }
            Add-GraphNode -GraphState $gs -Id $queryNodeId -Label $queryName -Group 'query' -Meta @{
                connect    = $qConnect
                sqlHash    = Get-GraphTextHash -Text $sqlText
                sqlPreview = Get-GraphPreviewText -Text $sqlText
                rawPath    = $rawInfo.path
                rawHash    = $rawInfo.hash
                rawSize    = $rawInfo.size
            } | Out-Null
            Register-GraphNameTarget -GraphState $gs -Name $queryName -NodeId $queryNodeId -Group 'query' -IsDataObject
        }

        # ── Forms ──
        $allForms = @($app.CurrentProject.AllForms)
        $formTotal = $allForms.Count
        $formIdx = 0
        foreach ($obj in $allForms) {
            $name = $null
            try { $name = [string]$obj.Name } catch { continue }
            $formIdx++
            $pct = [int](25 + (15 * $formIdx / [Math]::Max($formTotal, 1)))
            if (-not $Quiet) { Write-Progress -Activity 'Export-AccessGraph' -Status "Exporting form ($formIdx/$formTotal): $name" -PercentComplete $pct }

            $folder = if ($keepRaw) { Join-Path $rawDir 'forms' } else { [System.IO.Path]::GetTempPath() }
            $rawInfo = Save-AccessTextObject -App $app -Type $AC_TYPE.Form -Name $name -Folder $folder
            if (-not $keepRaw -and $rawInfo.path) { $tempPaths.Add($rawInfo.path) }

            $nodeId = Get-GraphObjectId -Group 'form' -Name $name
            Add-GraphNode -GraphState $gs -Id $nodeId -Label $name -Group 'form' -Meta @{
                rawPath = $rawInfo.path
                rawHash = $rawInfo.hash
                rawSize = $rawInfo.size
            } | Out-Null
            Register-GraphNameTarget -GraphState $gs -Name $name -NodeId $nodeId -Group 'form'
        }

        # ── Reports ──
        $allReports = @($app.CurrentProject.AllReports)
        $reportTotal = $allReports.Count
        $reportIdx = 0
        foreach ($obj in $allReports) {
            $name = $null
            try { $name = [string]$obj.Name } catch { continue }
            $reportIdx++
            $pct = [int](40 + (10 * $reportIdx / [Math]::Max($reportTotal, 1)))
            if (-not $Quiet) { Write-Progress -Activity 'Export-AccessGraph' -Status "Exporting report ($reportIdx/$reportTotal): $name" -PercentComplete $pct }

            $folder = if ($keepRaw) { Join-Path $rawDir 'reports' } else { [System.IO.Path]::GetTempPath() }
            $rawInfo = Save-AccessTextObject -App $app -Type $AC_TYPE.Report -Name $name -Folder $folder
            if (-not $keepRaw -and $rawInfo.path) { $tempPaths.Add($rawInfo.path) }

            $nodeId = Get-GraphObjectId -Group 'report' -Name $name
            Add-GraphNode -GraphState $gs -Id $nodeId -Label $name -Group 'report' -Meta @{
                rawPath = $rawInfo.path
                rawHash = $rawInfo.hash
                rawSize = $rawInfo.size
            } | Out-Null
            Register-GraphNameTarget -GraphState $gs -Name $name -NodeId $nodeId -Group 'report'
        }

        # ── Macros ──
        $allMacros = @($app.CurrentProject.AllMacros)
        if (-not $Quiet) { Write-Progress -Activity 'Export-AccessGraph' -Status 'Exporting macros...' -PercentComplete 50 }
        foreach ($obj in $allMacros) {
            $name = $null
            try { $name = [string]$obj.Name } catch { continue }

            $rawInfo = [pscustomobject]@{ path = $null; hash = $null; size = 0 }
            if ($keepRaw -or -not $DisableMacroHeuristics) {
                $folder = if ($keepRaw) { Join-Path $rawDir 'macros' } else { [System.IO.Path]::GetTempPath() }
                $rawInfo = Save-AccessTextObject -App $app -Type $AC_TYPE.Macro -Name $name -Folder $folder
                if (-not $keepRaw -and $rawInfo.path) { $tempPaths.Add($rawInfo.path) }
            }

            $nodeId = Get-GraphObjectId -Group 'macro' -Name $name
            Add-GraphNode -GraphState $gs -Id $nodeId -Label $name -Group 'macro' -Meta @{
                rawPath = $rawInfo.path
                rawHash = $rawInfo.hash
                rawSize = $rawInfo.size
            } | Out-Null
            Register-GraphNameTarget -GraphState $gs -Name $name -NodeId $nodeId -Group 'macro'
        }

        # ── Modules ──
        $allModules = @($app.CurrentProject.AllModules)
        if (-not $Quiet) { Write-Progress -Activity 'Export-AccessGraph' -Status 'Exporting modules...' -PercentComplete 55 }
        foreach ($obj in $allModules) {
            $name = $null
            try { $name = [string]$obj.Name } catch { continue }

            $rawInfo = [pscustomobject]@{ path = $null; hash = $null; size = 0 }
            if ($keepRaw -or -not $DisableCodeHeuristics) {
                $folder = if ($keepRaw) { Join-Path $rawDir 'modules' } else { [System.IO.Path]::GetTempPath() }
                $rawInfo = Save-AccessTextObject -App $app -Type $AC_TYPE.Module -Name $name -Folder $folder
                if (-not $keepRaw -and $rawInfo.path) { $tempPaths.Add($rawInfo.path) }
            }

            $nodeId = Get-GraphObjectId -Group 'module' -Name $name
            Add-GraphNode -GraphState $gs -Id $nodeId -Label $name -Group 'module' -Meta @{
                rawPath = $rawInfo.path
                rawHash = $rawInfo.hash
                rawSize = $rawInfo.size
            } | Out-Null
            Register-GraphNameTarget -GraphState $gs -Name $name -NodeId $nodeId -Group 'module'
        }

        # ──────────────────────────────────────────────────────────────
        #  EDGE ANALYSIS PHASE
        # ──────────────────────────────────────────────────────────────

        $knownDataNames = @($gs.DataNameTargets.Keys | Sort-Object { $_.Length } -Descending)

        # ── Query SQL edges ──
        if (-not $Quiet) { Write-Progress -Activity 'Export-AccessGraph' -Status 'Analyzing query edges...' -PercentComplete 60 }
        foreach ($queryDef in $db.QueryDefs) {
            $qdName = $null
            try { $qdName = [string]$queryDef.Name } catch { continue }
            if ([string]::IsNullOrWhiteSpace($qdName) -or (Test-SystemOrTemporaryName -Name $qdName)) { continue }

            $queryNodeId = Get-GraphObjectId -Group 'query' -Name $qdName
            $sqlText = try { [string]$queryDef.SQL } catch { '' }

            foreach ($referencedName in (Find-GraphReferencedDataName -Text $sqlText -KnownNames $knownDataNames)) {
                foreach ($target in (Get-GraphTargetsByName -Name $referencedName -TargetTable $gs.DataNameTargets)) {
                    if ($target.id -eq $queryNodeId) { continue }
                    Add-GraphEdge -GraphState $gs -From $queryNodeId -To $target.id -Label 'uses' -Kind 'query-sql-reference' -Arrows 'to' -Meta @{ name = $referencedName } | Out-Null
                }
            }
        }

        # ── Form edges ──
        $formEdgeIdx = 0
        foreach ($form in $allForms) {
            $name = [string]$form.Name
            $formEdgeIdx++
            $pct = [int](65 + (15 * $formEdgeIdx / [Math]::Max($formTotal, 1)))
            if (-not $Quiet) { Write-Progress -Activity 'Export-AccessGraph' -Status "Analyzing form edges ($formEdgeIdx/$formTotal): $name" -PercentComplete $pct }

            $rawPath = $gs.NodeIndex[(Get-GraphObjectId -Group 'form' -Name $name)].meta.rawPath
            try {
                Add-GraphFormReportEdge -GraphState $gs -ObjectGroup 'form' -ObjectName $name -RawPath $rawPath -SqlFolder $sqlFolder -KnownDataNames $knownDataNames -FieldNodeMode $FieldNodeMode -DisableCodeHeuristics:$DisableCodeHeuristics
            }
            catch {
                Add-GraphWarning -GraphState $gs -Code 'FormEdgeParseFailed' -Message ("Failed to parse form edges for '{0}': {1}" -f $name, $_.Exception.Message)
            }
        }

        # ── Report edges ──
        if (-not $Quiet) { Write-Progress -Activity 'Export-AccessGraph' -Status 'Analyzing report edges...' -PercentComplete 82 }
        foreach ($report in $allReports) {
            $name = [string]$report.Name
            $rawPath = $gs.NodeIndex[(Get-GraphObjectId -Group 'report' -Name $name)].meta.rawPath
            try {
                Add-GraphFormReportEdge -GraphState $gs -ObjectGroup 'report' -ObjectName $name -RawPath $rawPath -SqlFolder $sqlFolder -KnownDataNames $knownDataNames -FieldNodeMode $FieldNodeMode -DisableCodeHeuristics:$DisableCodeHeuristics
            }
            catch {
                Add-GraphWarning -GraphState $gs -Code 'ReportEdgeParseFailed' -Message ("Failed to parse report edges for '{0}': {1}" -f $name, $_.Exception.Message)
            }
        }

        # ── Macro edges ──
        if (-not $DisableMacroHeuristics) {
            if (-not $Quiet) { Write-Progress -Activity 'Export-AccessGraph' -Status 'Analyzing macro edges...' -PercentComplete 87 }
            foreach ($macro in $allMacros) {
                $name = [string]$macro.Name
                $rawPath = $gs.NodeIndex[(Get-GraphObjectId -Group 'macro' -Name $name)].meta.rawPath
                Add-GraphMacroHeuristicEdge -GraphState $gs -MacroName $name -RawPath $rawPath -SqlFolder $sqlFolder -KnownDataNames $knownDataNames
            }
        }

        # ── Module code edges ──
        if (-not $DisableCodeHeuristics) {
            if (-not $Quiet) { Write-Progress -Activity 'Export-AccessGraph' -Status 'Analyzing module code...' -PercentComplete 90 }
            foreach ($module in $allModules) {
                $name = [string]$module.Name
                $rawPath = $gs.NodeIndex[(Get-GraphObjectId -Group 'module' -Name $name)].meta.rawPath
                if ($rawPath -and (Test-Path -LiteralPath $rawPath)) {
                    $text = Get-Content -LiteralPath $rawPath -Raw
                    Add-GraphCodeHeuristicEdge -GraphState $gs -OwnerNodeId (Get-GraphObjectId -Group 'module' -Name $name) -OwnerGroup 'module' -OwnerName $name -Text $text -SqlFolder $sqlFolder -KnownDataNames $knownDataNames
                }
            }
        }

        # ──────────────────────────────────────────────────────────────
        #  OUTPUT PHASE
        # ──────────────────────────────────────────────────────────────

        if (-not $Quiet) { Write-Progress -Activity 'Export-AccessGraph' -Status 'Writing graph output...' -PercentComplete 95 }

        $graph = [pscustomobject][ordered]@{
            meta = [ordered]@{
                database      = $resolved
                generatedAt   = [DateTime]::UtcNow.ToString('o')
                fieldNodeMode = $FieldNodeMode
                stats         = [ordered]@{
                    nodeCount  = $gs.Nodes.Count
                    edgeCount  = $gs.Edges.Count
                    tables     = @($gs.Nodes | Where-Object { $_.group -eq 'table' }).Count
                    queries    = @($gs.Nodes | Where-Object { $_.group -eq 'query' }).Count
                    forms      = @($gs.Nodes | Where-Object { $_.group -eq 'form' }).Count
                    reports    = @($gs.Nodes | Where-Object { $_.group -eq 'report' }).Count
                    macros     = @($gs.Nodes | Where-Object { $_.group -eq 'macro' }).Count
                    modules    = @($gs.Nodes | Where-Object { $_.group -eq 'module' }).Count
                    sqlNodes   = @($gs.Nodes | Where-Object { $_.group -eq 'sql' }).Count
                    fieldNodes = @($gs.Nodes | Where-Object { $_.group -eq 'field' }).Count
                    warnings   = $gs.Warnings.Count
                }
                warnings      = $gs.Warnings.ToArray()
            }
            nodes = $gs.Nodes.ToArray()
            edges = $gs.Edges.ToArray()
        }

        $graphPath = Join-Path $OutDir 'graph.json'
        $graphJson = $graph | ConvertTo-Json -Depth 25
        Set-Content -LiteralPath $graphPath -Value $graphJson -Encoding UTF8

        Copy-GraphViewer -DestinationFolder $OutDir -GraphJson $graphJson -Disabled:$SkipViewerCopy

        if (-not $Quiet) { Write-Progress -Activity 'Export-AccessGraph' -Completed }
        Write-Host ('Graph written to: ' + $graphPath)
        Write-Host ('Nodes: {0}  Edges: {1}  Warnings: {2}' -f $gs.Nodes.Count, $gs.Edges.Count, $gs.Warnings.Count)

        if ($PassThru) {
            return $graph
        }
    }
    finally {
        # Release DAO database object (but NOT the Access app — session manages that)
        if ($db) {
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($db) } catch {}
        }

        # Clean up temp files when RawExportMode=None
        foreach ($tp in $tempPaths) {
            if ($tp -and (Test-Path -LiteralPath $tp)) {
                try { Remove-Item -LiteralPath $tp -Force -ErrorAction SilentlyContinue } catch {}
            }
        }
        if (-not $keepRaw -and $sqlFolder -and (Test-Path -LiteralPath $sqlFolder)) {
            try { Remove-Item -LiteralPath $sqlFolder -Recurse -Force -ErrorAction SilentlyContinue } catch {}
        }
    }
}
