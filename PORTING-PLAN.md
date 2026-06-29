# MCP-Access v0.7.36 → AccessPOSH Porting Plan

**Date:** 2026-05-29  
**Source:** `mcp_access/` (Python, v0.7.36)  
**Target:** `AccessPOSH/` (PowerShell, v1.0.0)

---

## Executive Summary

AccessPOSH (91 functions) is a PowerShell port of mcp_access (67 tools). AccessPOSH already covers all the original tools **plus** 29 extra functions (security, themes, ribbons, navigation pane, TempVars, subdatasheets, filtered reports, print, application info, import from XML/Database, VBA file import with encoding detection). However, mcp_access v0.7.35–v0.7.36 added **6 new capabilities** that AccessPOSH does not yet have.

The control parser depth bug (v0.7.34) was already fixed in AccessPOSH's `DesignView.ps1`. No backport needed.

---

## Gap Analysis

| # | mcp_access Feature | Python Source | AccessPOSH Equivalent | Status |
|---|---|---|---|---|
| 1 | `ac_clone_object` | `code.py` | — | **Missing** |
| 2 | `ac_search_data` | `sql.py` | — | **Missing** |
| 3 | `ac_manage_tab_order` | `controls.py` | — | **Missing** |
| 4 | `ac_graph_query` | `graph_query.py` | — | **Missing** |
| 5 | `ac_find_definition` | `vbe.py` | `Find-AccessUsage` (usage only) | **Missing** (different purpose) |
| 6 | `_detect_office_install` | `core.py` | — | **Missing** |

---

## Porting Tasks (Ordered by Priority)

### Task 1: `Copy-AccessObject` → `DatabaseOps.ps1`

**Source:** `code.py:ac_clone_object()`  
**Target file:** `Public/DatabaseOps.ps1`  
**PowerShell verb:** `Copy-AccessObject`

**Parameters:**
```powershell
function Copy-AccessObject {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][ValidateSet('form','report','module','class_module','query','macro')][string]$ObjectType,
        [Parameter(Mandatory)][string]$SourceName,
        [Parameter(Mandatory)][string]$TargetName,
        [switch]$Overwrite,
        [switch]$AsJson
    )
}
```

**Algorithm:**
1. Validate `$SourceName -ne $TargetName`
2. Connect to DB via `Connect-AccessDB`
3. Verify source exists via `app.SaveAsText` into temp file (RAW — NO `Remove-BinarySections`)
4. Read temp file via `Read-TempFile` (auto-detect encoding)
5. For `class_module`: normalize `Attribute VB_Name` to `$TargetName`
6. If target exists and `-Overwrite`: delete via `DoCmd.DeleteObject`
7. If target exists and no `-Overwrite`: throw error
8. Import via `Set-AccessCode` (which detects binary sections present and skips restore)
9. Return result object

**Key detail:** Do NOT strip binary sections — they must ride along for faithful cloning.

**Dependencies:** `Connect-AccessDB`, `Read-TempFile`, `Set-AccessCode` (existing)  
**Estimated complexity:** Medium — ~80 lines

---

### Task 2: `Search-AccessData` → `TableOps.ps1`

**Source:** `sql.py:ac_search_data()`  
**Target file:** `Public/TableOps.ps1`  
**PowerShell verb:** `Search-AccessData`

**Parameters:**
```powershell
function Search-AccessData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][string]$SearchText,
        [string[]]$Tables,
        [int]$MaxPerTable = 50,
        [int]$MaxTotal = 500,
        [switch]$MatchCase,
        [switch]$AsJson
    )
}
```

**Algorithm:**
1. Connect to DB, get `CurrentDb`
2. Enumerate `TableDefs`: skip system (`MSys*`, `~temp*`) and linked (`.Connect` non-empty)
3. Optional whitelist filter via `-Tables`
4. Per table: find Text (type 10) and Memo (type 12) fields
5. Build Jet SQL: `SELECT * FROM [table] WHERE [field1] LIKE '*needle*' OR [field2] LIKE '*needle*'`
6. Execute via DAO `OpenRecordset` with `TOP (MaxPerTable + 1)` for truncation detection
7. If `-MatchCase`: post-filter rows Python-side (Jet LIKE is case-insensitive)
8. Build excerpt around first match (40-char context window)
9. Track per-table and total truncation
10. Return grouped results

**Key detail:** Jet uses `*` as wildcard (not `%`). Bracket-escape field names (`]` → `]]`).

**Dependencies:** `Connect-AccessDB` (existing)  
**Estimated complexity:** High — ~120 lines (DAO recordset walking, encoding edge cases)

---

### Task 3: `Set-AccessTabOrder` → `FormReportOps.ps1`

**Source:** `controls.py:ac_manage_tab_order()`  
**Target file:** `Public/FormReportOps.ps1`  
**PowerShell verb:** `Set-AccessTabOrder`

**Parameters:**
```powershell
function Set-AccessTabOrder {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][ValidateSet('form','report')][string]$ObjectType,
        [Parameter(Mandatory)][string]$ObjectName,
        [Parameter(Mandatory)][ValidateSet('get','set','auto_renumber')][string]$Action,
        [string[]]$TabOrder,
        [string]$Section,
        [switch]$AsJson
    )
}
```

**Algorithm:**
1. Open form/report in Design view via `Open-InDesignView`
2. Enumerate controls, skip non-tabbable types: 100 (Label), 101 (Rectangle), 102 (Line), 103 (Image), 114 (PageBreak), 118 (Page)
3. Section resolution: name → enum via `$script:SECTION_MAP` or parse as int
4. **get**: Group by section, sort by TabIndex, return
5. **set**: Validate all names exist, single-pass assignment `ctrl.TabIndex = $idx` (Access auto-renumbers)
6. **auto_renumber**: Group by section, sort by current TabIndex, reassign 0..N-1

**Key constraint:** Do NOT try to set TabIndex >= N (Access rejects). Single-pass assignment is the correct idiom.

**Dependencies:** `Connect-AccessDB`, `Open-InDesignView` (existing)  
**Estimated complexity:** Medium — ~100 lines

---

### Task 4: `Get-AccessGraphQuery` → `GraphOps.ps1`

**Source:** `graph_query.py:ac_graph_query()` + `_Graph` class  
**Target file:** `Public/GraphOps.ps1`  
**PowerShell verb:** `Get-AccessGraphQuery`

**Parameters:**
```powershell
function Get-AccessGraphQuery {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('neighbors','impact','path','orphans','summary')][string]$Action,
        [string]$GraphPath,
        [string]$DbPath,
        [string]$Node,
        [string]$Source,
        [string]$Target,
        [int]$Depth = 1,
        [ValidateSet('in','out','both')][string]$Direction = 'both',
        [string]$Group,
        [switch]$IncludeFields,
        [switch]$AsJson
    )
}
```

**Algorithm:**
1. Locate `graph.json`: explicit path > sibling of DB > error
2. Parse JSON, build adjacency lists (in-edges, out-edges per node)
3. Node resolution: exact id → group:name pattern → label match (case-insensitive)
4. **neighbors**: BFS depth 1–3, collect in/out edges, truncate at 200
5. **impact**: Transitive downstream BFS, group by type
6. **path**: Undirected BFS for shortest path, return node chain + edge chain
7. **orphans**: Filter nodes with zero incoming edges, group by type
8. **summary**: Node/edge counts, top 15 high-degree nodes

**Private helper needed:** `ConvertFrom-GraphJson` in `Private/GraphHelpers.ps1` — loads JSON, builds lookup tables

**Dependencies:** None (pure data structure, no COM needed)  
**Estimated complexity:** High — ~200 lines (BFS, adjacency, 5 action handlers)

---

### Task 5: `Find-AccessDefinition` → `VbeOps.ps1`

**Source:** `vbe.py:ac_find_definition()` + `_scan_declarations()`  
**Target file:** `Public/VbeOps.ps1`  
**PowerShell verb:** `Find-AccessDefinition`

**Parameters:**
```powershell
function Find-AccessDefinition {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][string]$Symbol,
        [string[]]$Kinds,
        [switch]$MatchCase,
        [string[]]$ScanTypes,
        [switch]$FirstOnly,
        [switch]$AsJson
    )
}
```

**Algorithm:**
1. Validate `$Kinds` (const, enum, enum_member, type, type_field, sub, function, property, declare, variable)
2. For each module/form/report in scan scope:
   a. Read full code via `Get-AllModuleCode` or `CodeModule.Lines(1, total)`
   b. Join VBA line continuations (` _` at end of line)
   c. State machine parse:
      - Track `$inProc`, `$inEnum`, `$inType` states
      - Inside proc → skip (local vars excluded)
      - Inside enum → match members
      - Inside type → match fields
      - Module level → match Sub, Function, Property, Const, Enum, Type, Declare, Dim/Public/Private/Global
   d. Match symbol name (case-insensitive by default)
3. Return all matching definitions with kind, location, scope, value

**Key regex patterns:** (port from Python)
- Proc: `^\s*(Public |Private |Friend |Global )?(Static )?(Default )?(Sub|Function|Property\s+(Get|Let|Set))\s+(\w+)`
- Const: `^\s*(Public |Private |Global )?Const\s+`
- Enum: `^\s*(Public |Private )?Enum\s+(\w+)`
- Type: `^\s*(Public |Private )?Type\s+(\w+)`
- Declare: `^\s*(Public |Private )?Declare\s+(PtrSafe\s+)?(Sub|Function)\s+(\w+)\s+Lib\s+`
- Variable: `^\s*(Dim|Public|Private|Global)\s+`

**Private helpers needed:**
- `Join-VbaContinuations` — join lines ending with ` _`
- `Split-TopLevelCommas` — split by commas NOT inside parens/quotes

**Dependencies:** `Connect-AccessDB`, `Get-CodeModule` or `Get-AllModuleCode` (existing)  
**Estimated complexity:** Very High — ~250 lines (state machine, 10 declaration kinds, multi-const/multi-var parsing)

---

### Task 6: `Get-AccessOfficeVersion` → `Private/Session.ps1`

**Source:** `core.py:_detect_office_install()`  
**Target file:** `Private/Session.ps1`  
**PowerShell name:** `Get-AccessOfficeVersion` (private helper)

**Algorithm:**
1. Enumerate `HKLM:\Software\Microsoft\Office\*\Access\InstallRoot` (both native and WOW6432Node)
2. Enumerate `HKCU:\Software\Microsoft\Office\*\Access\InstallRoot` (per-user C2R)
3. For each version (highest first): check if `MSACCESS.EXE` exists at the InstallRoot path
4. Fallback: `HKLM:\Software\Microsoft\Windows\CurrentVersion\App Paths\MSACCESS.EXE`
5. Store in `$script:AccessSession.OfficeVersion` and `$script:AccessSession.MsAccessPath`

**Dependencies:** None  
**Estimated complexity:** Low — ~50 lines

---

## Module Manifest Update

After all functions are added, update `AccessPOSH.psd1`:
- Bump version to `1.1.0`
- Add new functions to `FunctionsToExport`
- Update description to reflect tool count (91 → 97)

---

## Implementation Order

| Phase | Task | New Functions | File(s) Modified | Complexity |
|-------|------|---------------|------------------|------------|
| 1 | Office version detect | `Get-AccessOfficeVersion` (private) | `Private/Session.ps1` | Low |
| 2 | Clone object | `Copy-AccessObject` | `Public/DatabaseOps.ps1` | Medium |
| 3 | Tab order | `Set-AccessTabOrder` | `Public/FormReportOps.ps1` | Medium |
| 4 | Search data | `Search-AccessData` | `Public/TableOps.ps1` | High |
| 5 | Graph query | `Get-AccessGraphQuery` | `Public/GraphOps.ps1`, `Private/GraphHelpers.ps1` | High |
| 6 | Find definition | `Find-AccessDefinition` | `Public/VbeOps.ps1` | Very High |
| 7 | Manifest update | — | `AccessPOSH.psd1` | Trivial |

**Rationale for order:**
- Phase 1 first: private helper with no external dependencies, used by Invoke-AccessDecompile
- Phases 2–3: medium complexity, each self-contained
- Phase 4: high complexity but isolated (table ops only)
- Phase 5: high complexity but pure data structure work (no COM)
- Phase 6 last: most complex, depends on VBE helpers being solid

---

## Test Plan

Each new function needs a Pester test file or additions to existing test files:

| Function | Test File | Test Cases |
|----------|-----------|------------|
| `Copy-AccessObject` | `Tests/DatabaseOps.Tests.ps1` | Clone form, report, module, class_module, query, macro; overwrite; source-not-found; same-name error |
| `Search-AccessData` | `Tests/TableOps.Tests.ps1` | Match found; no match; case-sensitive; table filter; truncation; system table skip |
| `Set-AccessTabOrder` | `Tests/FormReportOps.Tests.ps1` | get; set reorder; auto_renumber; section filter; non-tabbable skipped |
| `Get-AccessGraphQuery` | `Tests/GraphOps.Tests.ps1` | Each action (neighbors, impact, path, orphans, summary); node not found; ambiguous name |
| `Find-AccessDefinition` | `Tests/VbeOps.Tests.ps1` | Each kind (const, enum, sub, function, property, declare, variable, type, enum_member, type_field); multi-const; case sensitivity |

---

## Risk Assessment

| Risk | Mitigation |
|------|-----------|
| COM thread affinity in PowerShell | AccessPOSH already handles this correctly in Session.ps1 |
| DAO Recordset walking (Search-AccessData) | Follow existing Invoke-AccessSQL pattern |
| Graph JSON schema drift | Load from same Export-AccessGraph output — schema is stable |
| VBA line continuation edge cases | Port Python's `_join_continuations` logic exactly |
| Multi-const / multi-var parsing | Port `_split_top_level_commas` carefully — tricky regex |
| ShouldProcess support | Add to Copy-AccessObject and Set-AccessTabOrder (state-changing) |

---

## Rubber Duck Review Findings (2026-05-29)

### Verified OK (Non-Issues)
- **Module loader**: `AccessPOSH.psm1` line 222 auto-discovers all `Public/*.ps1` and `Private/*.ps1` — new files load automatically
- **Binary section handling**: `Set-AccessCode` calls `Restore-BinarySections` only when needed — Copy-AccessObject can rely on this
- **`$script:SECTION_MAP`**: Already exists in `AccessPOSH.psm1` line 121 with full section mappings
- **Error handling pattern**: All functions use `$PSCmdlet.ThrowTerminatingError()` consistently — new functions must follow this
- **Read-TempFile, Connect-AccessDB, Open-InDesignView, Get-CodeModule, Get-AllModuleCode**: All confirmed present and working

### Issues Found and Corrections

#### CRITICAL — Must Address Before Implementation

1. **`Join-VbaContinuations` helper missing**: VBA line continuation (`_` at end of line) joining does not exist anywhere in AccessPOSH. Must implement as a private helper in `Private/VbeHelpers.ps1` before `Find-AccessDefinition` can work. Without this, the state machine parser fails on any multi-line declaration.

2. **Class module `VB_Name` normalization in Copy-AccessObject**: When cloning a `class_module`, the `Attribute VB_Name` attribute in the exported text must be rewritten to the target name. Plan's algorithm step 5 mentions this but didn't show the implementation. Add explicit regex replacement:
   ```powershell
   if ($ObjectType -eq 'class_module') {
       $Code = $Code -replace 'Attribute VB_Name\s*=\s*"[^"]*"', "Attribute VB_Name = `"$TargetName`""
   }
   ```

#### IMPORTANT — Should Address

3. **Naming: `Invoke-AccessGraphQuery` → `Get-AccessGraphQuery`**: Since graph querying is read-only (no side effects), `Get-` is the correct verb per PowerShell conventions. `Invoke-` implies an action with side effects.

4. **Graph JSON schema validation**: `Get-AccessGraphQuery` must validate that the loaded JSON has `nodes` and `edges` arrays. Document the expected schema (output of `Export-AccessGraph`).

5. **PS7 COM RCW loss risk in Search-AccessData**: When walking DAO Recordset fields in a loop, avoid storing intermediate COM objects in variables across iterations. Chain property access each time. Follow existing `Invoke-AccessSQL` pattern.

6. **Section parameter INT parsing in Set-AccessTabOrder**: Must handle both string names and integer values:
   ```powershell
   $sectionInt = $null
   if ([int]::TryParse($Section, [ref]$sectionInt)) {
       # numeric
   } elseif ($script:SECTION_MAP.ContainsKey($Section.ToLower())) {
       $sectionInt = $script:SECTION_MAP[$Section.ToLower()]
   } else {
       throw "Unrecognized section: $Section"
   }
   ```

#### SUGGESTIONS

7. **Manifest function count**: Current manifest lists ~91 in description but may actually export ~97. Verify exact count after all additions.

8. **Macro encoding**: Verify that `Set-AccessCode` writes macros as UTF-16 (not cp1252). From CLAUDE.md: macros are UTF-16 encoded like forms.

### Private Helpers to Add

| Helper | File | Used By |
|--------|------|---------|
| `Join-VbaContinuations` | `Private/VbeHelpers.ps1` | `Find-AccessDefinition` |
| `Split-TopLevelCommas` | `Private/VbeHelpers.ps1` | `Find-AccessDefinition` (multi-const/multi-var) |
| `ConvertFrom-GraphJson` | `Private/GraphHelpers.ps1` | `Get-AccessGraphQuery` |
| `Get-AccessOfficeVersion` | `Private/Session.ps1` | `Invoke-AccessDecompile` |
