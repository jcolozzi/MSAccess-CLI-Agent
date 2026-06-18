# Public/GraphQueryOps.ps1 — Graph query functions for AccessPOSH
#
# Individual PowerShell-idiomatic functions for querying a graph.json
# produced by Export-AccessGraph.  All accept either -Graph (pipeline
# from Import-AccessGraph) or -GraphPath / -DbPath for direct file access.

# ──────────────────────────────────────────────────────────────────────
#  Import-AccessGraph
# ──────────────────────────────────────────────────────────────────────

function Import-AccessGraph {
    <#
    .SYNOPSIS
        Load a graph.json into an indexed object for pipeline queries.

    .DESCRIPTION
        Wraps the internal ConvertFrom-GraphJson, providing a public entry point
        with DbPath auto-locate support.  The returned object carries pre-built
        adjacency maps (OutAdj / InAdj) and hashtable indexes (IdLookup / LabelLookup),
        enabling O(1) lookups for all downstream query functions.

        Pipe the result to Get-AccessGraphNode, Get-AccessGraphEdge, etc. to avoid
        re-reading the JSON file on every call.

    .PARAMETER Path
        Path to the graph.json file.

    .PARAMETER DbPath
        Path to the Access database.  graph.json is auto-located at
        <db-dir>/access-graph-out/graph.json.

    .EXAMPLE
        $g = Import-AccessGraph -Path .\access-graph-out\graph.json
        $g | Get-AccessGraphNode -Group table

    .EXAMPLE
        Import-AccessGraph -DbPath C:\MyApp.accdb | Get-AccessGraphStats
    #>
    [CmdletBinding(DefaultParameterSetName = 'ByPath')]
    param(
        [Parameter(Mandatory, Position = 0, ParameterSetName = 'ByPath')]
        [string]$Path,

        [Parameter(Mandatory, ParameterSetName = 'ByDbPath')]
        [string]$DbPath
    )

    return Resolve-GraphInput -GraphPath $Path -DbPath $DbPath
}


# ──────────────────────────────────────────────────────────────────────
#  Get-AccessGraphNode
# ──────────────────────────────────────────────────────────────────────

function Get-AccessGraphNode {
    <#
    .SYNOPSIS
        Query nodes from an Access dependency graph.

    .DESCRIPTION
        Filters nodes by exact id, name pattern (wildcards supported), and/or group.
        With no filters returns all nodes.  Outputs raw node objects to the pipeline.

    .PARAMETER Graph
        Pre-loaded graph object (from Import-AccessGraph).

    .PARAMETER GraphPath
        Path to graph.json.  Ignored when -Graph is supplied.

    .PARAMETER Id
        Exact node id (e.g. 'table:Customers').  Case-insensitive.

    .PARAMETER Name
        Node label filter.  Supports wildcards (e.g. 'Cust*').  Case-insensitive.

    .PARAMETER Group
        Filter by one or more node groups (table, query, form, report, macro, module, field, sql).

    .PARAMETER AsJson
        Emit output as a JSON string.

    .EXAMPLE
        Import-AccessGraph -Path .\graph.json | Get-AccessGraphNode -Group table

    .EXAMPLE
        Get-AccessGraphNode -GraphPath .\graph.json -Name "Cust*"

    .EXAMPLE
        Get-AccessGraphNode -GraphPath .\graph.json -Id "form:frmMain"
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline)]
        [PSCustomObject]$Graph,

        [string]$GraphPath,

        [string]$Id,

        [string]$Name,

        [string[]]$Group,

        [switch]$AsJson
    )

    process {
        $g = Resolve-GraphInput -Graph $Graph -GraphPath $GraphPath
        $nodes = $g.Nodes.Values

        # Exact id lookup
        if ($Id) {
            $low = $Id.Trim().ToLower()
            if ($g.IdLookup.ContainsKey($low)) {
                $nodes = @($g.IdLookup[$low])
            } else {
                $nodes = @()
            }
        }

        # Name filter (wildcard)
        if ($Name) {
            $nodes = @($nodes | Where-Object { $_.label -like $Name })
        }

        # Group filter
        if ($Group -and $Group.Count -gt 0) {
            $groupLower = $Group | ForEach-Object { $_.ToLower() }
            $nodes = @($nodes | Where-Object { $_.group.ToLower() -in $groupLower })
        }

        $result = @($nodes)

        if ($AsJson) {
            return ($result | ConvertTo-Json -Depth 10 -Compress)
        }
        return $result
    }
}


# ──────────────────────────────────────────────────────────────────────
#  Get-AccessGraphEdge
# ──────────────────────────────────────────────────────────────────────

function Get-AccessGraphEdge {
    <#
    .SYNOPSIS
        Query edges from an Access dependency graph.

    .DESCRIPTION
        Filters edges by kind, source node, target node, or any-direction node.
        Filters combine with AND semantics.  With no filters returns all edges.

    .PARAMETER Graph
        Pre-loaded graph object (from Import-AccessGraph).

    .PARAMETER GraphPath
        Path to graph.json.

    .PARAMETER Kind
        Filter by one or more edge kinds (e.g. vba-openform, recordsource).

    .PARAMETER From
        Return edges originating FROM this node id.

    .PARAMETER To
        Return edges pointing TO this node id.

    .PARAMETER NodeId
        Return edges FROM or TO this node id (union).

    .PARAMETER IncludeFields
        Include field-owner edges.  By default they are excluded.

    .PARAMETER AsJson
        Emit output as a JSON string.

    .EXAMPLE
        Import-AccessGraph -Path .\graph.json | Get-AccessGraphEdge -Kind vba-openform

    .EXAMPLE
        Get-AccessGraphEdge -GraphPath .\graph.json -NodeId "table:Customers"
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline)]
        [PSCustomObject]$Graph,

        [string]$GraphPath,

        [string[]]$Kind,

        [string]$From,

        [string]$To,

        [string]$NodeId,

        [switch]$IncludeFields,

        [switch]$AsJson
    )

    process {
        $g = Resolve-GraphInput -Graph $Graph -GraphPath $GraphPath

        # Start with targeted lookup when possible, otherwise all edges
        if ($NodeId) {
            $outEdges = if ($g.OutAdj.ContainsKey($NodeId)) { @($g.OutAdj[$NodeId]) } else { @() }
            $inEdges  = if ($g.InAdj.ContainsKey($NodeId))  { @($g.InAdj[$NodeId])  } else { @() }
            $edges = $outEdges + $inEdges
        } elseif ($From) {
            $edges = if ($g.OutAdj.ContainsKey($From)) { @($g.OutAdj[$From]) } else { @() }
        } elseif ($To) {
            $edges = if ($g.InAdj.ContainsKey($To)) { @($g.InAdj[$To]) } else { @() }
        } else {
            $edges = @($g.Edges)
        }

        # Exclude field-owner edges by default
        if (-not $IncludeFields) {
            $edges = @($edges | Where-Object { $_.kind -ine 'field-owner' })
        }

        # Kind filter
        if ($Kind -and $Kind.Count -gt 0) {
            $kindLower = $Kind | ForEach-Object { $_.ToLower() }
            $edges = @($edges | Where-Object { $_.kind.ToLower() -in $kindLower })
        }

        # Additional From/To filter when NodeId was used (already filtered by adjacency)
        # or when combined with Kind
        if ($From -and $NodeId) {
            $edges = @($edges | Where-Object { $_.from -ieq $From })
        }
        if ($To -and $NodeId) {
            $edges = @($edges | Where-Object { $_.to -ieq $To })
        }

        $result = @($edges)

        if ($AsJson) {
            return ($result | ConvertTo-Json -Depth 10 -Compress)
        }
        return $result
    }
}


# ──────────────────────────────────────────────────────────────────────
#  Get-AccessGraphStats
# ──────────────────────────────────────────────────────────────────────

function Get-AccessGraphStats {
    <#
    .SYNOPSIS
        Get summary statistics from an Access dependency graph.

    .DESCRIPTION
        Returns node/edge counts, per-group breakdowns, edge kind distribution,
        top connected nodes (by degree), and orphan count.

    .PARAMETER Graph
        Pre-loaded graph object (from Import-AccessGraph).

    .PARAMETER GraphPath
        Path to graph.json.

    .PARAMETER Top
        Number of highest-degree nodes to include.  Default: 10.

    .PARAMETER AsJson
        Emit output as a JSON string.

    .EXAMPLE
        Import-AccessGraph -Path .\graph.json | Get-AccessGraphStats

    .EXAMPLE
        Get-AccessGraphStats -GraphPath .\graph.json -Top 20 -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline)]
        [PSCustomObject]$Graph,

        [string]$GraphPath,

        [int]$Top = 10,

        [switch]$AsJson
    )

    process {
        $g = Resolve-GraphInput -Graph $Graph -GraphPath $GraphPath

        # Node counts by group
        $nodeCounts = [ordered]@{}
        foreach ($n in $g.Nodes.Values) {
            $grp = [string]$n.group
            if (-not $nodeCounts.Contains($grp)) { $nodeCounts[$grp] = 0 }
            $nodeCounts[$grp]++
        }

        # Edge counts by kind
        $edgeCounts = [ordered]@{}
        foreach ($e in $g.Edges) {
            $kind = [string]$e.kind
            if (-not $edgeCounts.Contains($kind)) { $edgeCounts[$kind] = 0 }
            $edgeCounts[$kind]++
        }
        $sortedEdges = @($edgeCounts.GetEnumerator() |
            Sort-Object -Property Value -Descending |
            ForEach-Object { [PSCustomObject]@{ kind = $_.Key; count = $_.Value } })

        # Degree analysis (exclude field-owner)
        $degree = @{}
        foreach ($e in $g.Edges) {
            if ($e.kind -ieq 'field-owner') { continue }
            $fromId = [string]$e.from
            $toId   = [string]$e.to
            if (-not $degree.ContainsKey($fromId)) { $degree[$fromId] = 0 }
            $degree[$fromId]++
            if (-not $degree.ContainsKey($toId)) { $degree[$toId] = 0 }
            $degree[$toId]++
        }

        $topNodes = @($degree.GetEnumerator() |
            Sort-Object -Property Value -Descending |
            Select-Object -First $Top |
            ForEach-Object {
                $n = $g.Nodes[$_.Key]
                if ($null -ne $n) {
                    [PSCustomObject][ordered]@{
                        id     = $n.id
                        label  = $n.label
                        group  = $n.group
                        degree = $_.Value
                    }
                }
            })

        # Orphan count (no incoming edges, excluding fields)
        $orphanCount = 0
        foreach ($nid in $g.Nodes.Keys) {
            $n = $g.Nodes[$nid]
            if ($n.group -ieq 'field') { continue }
            $hasIncoming = $false
            if ($g.InAdj.ContainsKey($nid)) {
                foreach ($e in $g.InAdj[$nid]) {
                    if ($e.kind -ine 'field-owner') { $hasIncoming = $true; break }
                }
            }
            if (-not $hasIncoming) { $orphanCount++ }
        }

        $result = [PSCustomObject][ordered]@{
            database       = if ($g.Meta.ContainsKey('database')) { $g.Meta['database'] } else { '' }
            generatedAt    = if ($g.Meta.ContainsKey('generatedAt')) { $g.Meta['generatedAt'] } else { '' }
            totalNodes     = $g.Nodes.Count
            totalEdges     = $g.Edges.Count
            orphanCount    = $orphanCount
            nodesByGroup   = [PSCustomObject]$nodeCounts
            edgesByKind    = $sortedEdges
            topConnected   = $topNodes
        }

        if ($AsJson) {
            return ($result | ConvertTo-Json -Depth 10 -Compress)
        }
        return $result
    }
}


# ──────────────────────────────────────────────────────────────────────
#  Find-AccessGraphPath
# ──────────────────────────────────────────────────────────────────────

function Find-AccessGraphPath {
    <#
    .SYNOPSIS
        Find the shortest path between two nodes in the dependency graph.

    .DESCRIPTION
        Uses BFS to find the shortest path.  By default follows edge direction
        (from → to).  Use -Undirected to ignore edge direction.

    .PARAMETER Graph
        Pre-loaded graph object (from Import-AccessGraph).

    .PARAMETER GraphPath
        Path to graph.json.

    .PARAMETER From
        Source node — name or full id (e.g. 'Customers' or 'table:Customers').

    .PARAMETER To
        Target node — name or full id.

    .PARAMETER MaxDepth
        Maximum BFS depth.  Default: 10.

    .PARAMETER Undirected
        Treat edges as undirected (traverse in both directions).

    .PARAMETER AsJson
        Emit output as a JSON string.

    .EXAMPLE
        Import-AccessGraph -Path .\graph.json | Find-AccessGraphPath -From Customers -To frmOrders

    .EXAMPLE
        Find-AccessGraphPath -GraphPath .\graph.json -From "table:Customers" -To "form:frmMain" -Undirected
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline)]
        [PSCustomObject]$Graph,

        [string]$GraphPath,

        [Parameter(Mandatory)]
        [string]$From,

        [Parameter(Mandatory)]
        [string]$To,

        [int]$MaxDepth = 10,

        [switch]$Undirected,

        [switch]$AsJson
    )

    process {
        $g = Resolve-GraphInput -Graph $Graph -GraphPath $GraphPath

        # Resolve node names
        $fromNodes = Resolve-GraphNode -Graph $g -Name $From
        $toNodes   = Resolve-GraphNode -Graph $g -Name $To

        if ($fromNodes.Count -eq 0) {
            Write-Error "Source node '$From' not found in graph."
            return
        }
        if ($toNodes.Count -eq 0) {
            Write-Error "Target node '$To' not found in graph."
            return
        }
        if ($fromNodes.Count -gt 1) {
            $opts = ($fromNodes | ForEach-Object { $_.id }) -join ', '
            Write-Error "Source '$From' is ambiguous: $opts. Use full id."
            return
        }
        if ($toNodes.Count -gt 1) {
            $opts = ($toNodes | ForEach-Object { $_.id }) -join ', '
            Write-Error "Target '$To' is ambiguous: $opts. Use full id."
            return
        }

        $sourceId = [string]$fromNodes[0].id
        $targetId = [string]$toNodes[0].id

        if ($sourceId -ieq $targetId) {
            $result = [PSCustomObject][ordered]@{
                source = $fromNodes[0]
                target = $toNodes[0]
                found  = $true
                path   = @($fromNodes[0])
                edges  = @()
                length = 0
            }
            if ($AsJson) { return ($result | ConvertTo-Json -Depth 10 -Compress) }
            return $result
        }

        # Build adjacency for BFS
        if ($Undirected) {
            # Undirected: build combined adjacency
            $adj = @{}
            foreach ($e in $g.Edges) {
                $fId = [string]$e.from
                $tId = [string]$e.to
                if (-not $adj.ContainsKey($fId)) {
                    $adj[$fId] = New-Object 'System.Collections.Generic.List[object]'
                }
                $adj[$fId].Add([PSCustomObject]@{ Neighbor = $tId; Edge = $e })
                if (-not $adj.ContainsKey($tId)) {
                    $adj[$tId] = New-Object 'System.Collections.Generic.List[object]'
                }
                $adj[$tId].Add([PSCustomObject]@{ Neighbor = $fId; Edge = $e })
            }
        }

        # BFS
        $visited = New-Object 'System.Collections.Generic.HashSet[string]'
        [void]$visited.Add($sourceId)
        $parent = @{}   # child → @{ ParentId; Edge; Depth }
        $frontier = New-Object 'System.Collections.Generic.Queue[object]'
        $frontier.Enqueue([PSCustomObject]@{ Id = $sourceId; Depth = 0 })
        $found = $false

        while ($frontier.Count -gt 0 -and -not $found) {
            $item = $frontier.Dequeue()
            $cur = [string]$item.Id
            $d   = [int]$item.Depth

            if ($d -ge $MaxDepth) { continue }

            if ($Undirected) {
                $neighbors = if ($adj.ContainsKey($cur)) { $adj[$cur] } else { @() }
            } else {
                # Directed: follow outgoing edges only
                $neighbors = if ($g.OutAdj.ContainsKey($cur)) {
                    @($g.OutAdj[$cur] | ForEach-Object {
                        [PSCustomObject]@{ Neighbor = [string]$_.to; Edge = $_ }
                    })
                } else { @() }
            }

            foreach ($pair in $neighbors) {
                $neighbor = [string]$pair.Neighbor
                if (-not $visited.Contains($neighbor)) {
                    [void]$visited.Add($neighbor)
                    $parent[$neighbor] = [PSCustomObject]@{
                        ParentId = $cur
                        Edge     = $pair.Edge
                        Depth    = ($d + 1)
                    }
                    if ($neighbor -ieq $targetId) {
                        $found = $true
                        break
                    }
                    $frontier.Enqueue([PSCustomObject]@{ Id = $neighbor; Depth = ($d + 1) })
                }
            }
        }

        if (-not $found) {
            $result = [PSCustomObject][ordered]@{
                source = $fromNodes[0]
                target = $toNodes[0]
                found  = $false
                path   = @()
                edges  = @()
                length = -1
            }
            if ($AsJson) { return ($result | ConvertTo-Json -Depth 10 -Compress) }
            return $result
        }

        # Reconstruct path
        $pathNodes = New-Object 'System.Collections.Generic.List[object]'
        $pathEdges = New-Object 'System.Collections.Generic.List[object]'
        $cur = $targetId
        while ($cur -ine $sourceId) {
            $n = $g.Nodes[$cur]
            if ($null -ne $n) { $pathNodes.Add($n) } else { Write-Warning "Missing node in graph: $cur" }
            $pInfo = $parent[$cur]
            $pathEdges.Add($pInfo.Edge)
            $cur = [string]$pInfo.ParentId
        }
        $n = $g.Nodes[$sourceId]
        if ($null -ne $n) { $pathNodes.Add($n) } else { Write-Warning "Missing node in graph: $sourceId" }
        $pathNodes.Reverse()
        $pathEdges.Reverse()

        $result = [PSCustomObject][ordered]@{
            source = $fromNodes[0]
            target = $toNodes[0]
            found  = $true
            path   = $pathNodes.ToArray()
            edges  = $pathEdges.ToArray()
            length = $pathEdges.Count
        }

        if ($AsJson) { return ($result | ConvertTo-Json -Depth 10 -Compress) }
        return $result
    }
}


# ──────────────────────────────────────────────────────────────────────
#  Get-AccessGraphImpact
# ──────────────────────────────────────────────────────────────────────

function Get-AccessGraphImpact {
    <#
    .SYNOPSIS
        Find all nodes transitively affected by (or depended on by) a given node.

    .DESCRIPTION
        Performs a BFS traversal from the specified node.

        Downstream (default): follows outgoing edges — "what depends on this?"
        Upstream:             follows incoming edges — "what does this depend on?"

    .PARAMETER Graph
        Pre-loaded graph object (from Import-AccessGraph).

    .PARAMETER GraphPath
        Path to graph.json.

    .PARAMETER NodeId
        Starting node — name or full id (e.g. 'Customers' or 'table:Customers').

    .PARAMETER Direction
        Traversal direction.  Downstream (default) follows from→to edges.
        Upstream follows to→from edges.

    .PARAMETER MaxDepth
        Maximum BFS depth.  Default: 10.

    .PARAMETER IncludeFields
        Include field-owner edges in traversal.  Excluded by default.

    .PARAMETER AsJson
        Emit output as a JSON string.

    .EXAMPLE
        Import-AccessGraph -Path .\graph.json | Get-AccessGraphImpact -NodeId "table:Customers"

    .EXAMPLE
        Get-AccessGraphImpact -GraphPath .\graph.json -NodeId frmMain -Direction Upstream -MaxDepth 5
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline)]
        [PSCustomObject]$Graph,

        [string]$GraphPath,

        [Parameter(Mandatory)]
        [string]$NodeId,

        [ValidateSet('Downstream','Upstream')]
        [string]$Direction = 'Downstream',

        [int]$MaxDepth = 10,

        [switch]$IncludeFields,

        [switch]$AsJson
    )

    process {
        $g = Resolve-GraphInput -Graph $Graph -GraphPath $GraphPath

        # Resolve node
        $hits = Resolve-GraphNode -Graph $g -Name $NodeId
        if ($hits.Count -eq 0) {
            Write-Error "Node '$NodeId' not found in graph."
            return
        }
        if ($hits.Count -gt 1) {
            $opts = ($hits | ForEach-Object { $_.id }) -join ', '
            Write-Error "Node '$NodeId' is ambiguous: $opts. Use full id."
            return
        }

        $startId = [string]$hits[0].id
        $skipFields = -not $IncludeFields

        # Choose adjacency map by direction
        $adjMap = if ($Direction -ieq 'Upstream') { $g.InAdj } else { $g.OutAdj }

        $visited = New-Object 'System.Collections.Generic.HashSet[string]'
        [void]$visited.Add($startId)
        $frontier = New-Object 'System.Collections.Generic.Queue[object]'
        $frontier.Enqueue([PSCustomObject]@{ Id = $startId; Depth = 0 })

        $affected = New-Object 'System.Collections.Generic.List[object]'
        $edgesUsed = New-Object 'System.Collections.Generic.List[object]'

        while ($frontier.Count -gt 0) {
            $item = $frontier.Dequeue()
            $cur = [string]$item.Id
            $d   = [int]$item.Depth

            if ($d -ge $MaxDepth) { continue }

            if (-not $adjMap.ContainsKey($cur)) { continue }

            foreach ($e in $adjMap[$cur]) {
                if ($skipFields -and $e.kind -ieq 'field-owner') { continue }

                $nextId = if ($Direction -ieq 'Upstream') { [string]$e.from } else { [string]$e.to }
                $edgesUsed.Add($e)

                if (-not $visited.Contains($nextId)) {
                    [void]$visited.Add($nextId)
                    $nextNode = $g.Nodes[$nextId]
                    if ($null -ne $nextNode) {
                        $affected.Add($nextNode)
                    }
                    $frontier.Enqueue([PSCustomObject]@{ Id = $nextId; Depth = ($d + 1) })
                }
            }
        }

        # Group by type
        $byGroup = [ordered]@{}
        foreach ($a in $affected) {
            $grp = [string]$a.group
            if (-not $byGroup.Contains($grp)) {
                $byGroup[$grp] = New-Object 'System.Collections.Generic.List[string]'
            }
            $byGroup[$grp].Add([string]$a.label)
        }

        $result = [PSCustomObject][ordered]@{
            node          = $hits[0]
            direction     = $Direction
            maxDepth      = $MaxDepth
            affectedCount = $affected.Count
            byGroup       = [PSCustomObject]$byGroup
            affected      = @($affected)
            edges         = @($edgesUsed)
        }

        if ($AsJson) { return ($result | ConvertTo-Json -Depth 10 -Compress) }
        return $result
    }
}


# ──────────────────────────────────────────────────────────────────────
#  Get-AccessGraphOrphan
# ──────────────────────────────────────────────────────────────────────

function Get-AccessGraphOrphan {
    <#
    .SYNOPSIS
        Find nodes with no incoming edges (unreferenced objects).

    .DESCRIPTION
        Returns nodes that have zero incoming edges (excluding field-owner by default).
        These are objects that nothing else in the database references — potential
        candidates for cleanup or review.

    .PARAMETER Graph
        Pre-loaded graph object (from Import-AccessGraph).

    .PARAMETER GraphPath
        Path to graph.json.

    .PARAMETER Group
        Filter orphans by one or more node groups.

    .PARAMETER IncludeFields
        Include field nodes and field-owner edges in the analysis.

    .PARAMETER AsJson
        Emit output as a JSON string.

    .EXAMPLE
        Import-AccessGraph -Path .\graph.json | Get-AccessGraphOrphan

    .EXAMPLE
        Get-AccessGraphOrphan -GraphPath .\graph.json -Group table,query -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline)]
        [PSCustomObject]$Graph,

        [string]$GraphPath,

        [string[]]$Group,

        [switch]$IncludeFields,

        [switch]$AsJson
    )

    process {
        $g = Resolve-GraphInput -Graph $Graph -GraphPath $GraphPath
        $skipFields = -not $IncludeFields

        $orphans = New-Object 'System.Collections.Generic.List[object]'

        foreach ($nid in $g.Nodes.Keys) {
            $n = $g.Nodes[$nid]

            # Skip field nodes unless requested
            if ($skipFields -and $n.group -ieq 'field') { continue }

            # Group filter
            if ($Group -and $Group.Count -gt 0) {
                $groupLower = $Group | ForEach-Object { $_.ToLower() }
                if ($n.group.ToLower() -notin $groupLower) { continue }
            }

            $hasIncoming = $false
            if ($g.InAdj.ContainsKey($nid)) {
                foreach ($e in $g.InAdj[$nid]) {
                    if ($skipFields -and $e.kind -ieq 'field-owner') { continue }
                    $hasIncoming = $true
                    break
                }
            }

            if (-not $hasIncoming) {
                $orphans.Add($n)
            }
        }

        # Group by type
        $byGroup = [ordered]@{}
        foreach ($o in $orphans) {
            $grp = [string]$o.group
            if (-not $byGroup.Contains($grp)) {
                $byGroup[$grp] = New-Object 'System.Collections.Generic.List[string]'
            }
            $byGroup[$grp].Add([string]$o.label)
        }

        $result = [PSCustomObject][ordered]@{
            orphanCount = $orphans.Count
            byGroup     = [PSCustomObject]$byGroup
            orphans     = @($orphans)
        }

        if ($AsJson) { return ($result | ConvertTo-Json -Depth 10 -Compress) }
        return $result
    }
}
