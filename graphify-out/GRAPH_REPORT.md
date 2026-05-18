# Graph Report - MSAccess-agent  (2026-05-11)

## Corpus Check
- Corpus is ~39,421 words - fits in a single context window. You may not need a graph.

## Summary
- 254 nodes · 583 edges · 45 communities (27 shown, 18 thin omitted)
- Extraction: 37% EXTRACTED · 63% INFERRED · 0% AMBIGUOUS · INFERRED: 367 edges (avg confidence: 0.8)
- Token cost: 0 input · 0 output

## Community Hubs (Navigation)
- [[_COMMUNITY_Public Module Files|Public Module Files]]
- [[_COMMUNITY_Core Operations Logic|Core Operations Logic]]
- [[_COMMUNITY_VBE & Code Helpers|VBE & Code Helpers]]
- [[_COMMUNITY_Design View & Forms|Design View & Forms]]
- [[_COMMUNITY_Binary & Export Helpers|Binary & Export Helpers]]
- [[_COMMUNITY_COM Session Management|COM Session Management]]
- [[_COMMUNITY_VBE Code Editing|VBE Code Editing]]
- [[_COMMUNITY_UI Automation & Win32|UI Automation & Win32]]
- [[_COMMUNITY_Theme Operations|Theme Operations]]
- [[_COMMUNITY_Text Match Utilities|Text Match Utilities]]
- [[_COMMUNITY_CodeBehind Splitter|CodeBehind Splitter]]
- [[_COMMUNITY_Control Parser|Control Parser]]
- [[_COMMUNITY_VBE Module Info|VBE Module Info]]
- [[_COMMUNITY_AppOps Tests|AppOps Tests]]
- [[_COMMUNITY_DatabaseOps Tests|DatabaseOps Tests]]
- [[_COMMUNITY_ExportTransfer Tests|ExportTransfer Tests]]
- [[_COMMUNITY_FormReportOps Tests|FormReportOps Tests]]
- [[_COMMUNITY_ImportOps Tests|ImportOps Tests]]
- [[_COMMUNITY_MetadataOps Tests|MetadataOps Tests]]
- [[_COMMUNITY_NavigationPane Tests|NavigationPane Tests]]
- [[_COMMUNITY_PrintOps Tests|PrintOps Tests]]
- [[_COMMUNITY_ReportOps Tests|ReportOps Tests]]
- [[_COMMUNITY_RibbonOps Tests|RibbonOps Tests]]
- [[_COMMUNITY_SecurityOps Tests|SecurityOps Tests]]
- [[_COMMUNITY_SubDataSheet Tests|SubDataSheet Tests]]
- [[_COMMUNITY_TableOps Tests|TableOps Tests]]
- [[_COMMUNITY_TempVarOps Tests|TempVarOps Tests]]

## God Nodes (most connected - your core abstractions)
1. `Format-AccessOutput()` - 91 edges
2. `Resolve-SessionDbPath()` - 89 edges
3. `Connect-AccessDB()` - 89 edges
4. `Connect-AccessDB` - 19 edges
5. `Open-InDesignView()` - 12 edges
6. `Save-AndCloseDesign()` - 12 edges
7. `Import-AccessSource()` - 11 edges
8. `Get-CodeModule()` - 10 edges
9. `VbeOps Tests` - 10 edges
10. `Get-AccessObject()` - 9 edges

## Surprising Connections (you probably didn't know these)
- `COM Automation Architecture` --conceptually_related_to--> `Connect-AccessDB`  [INFERRED]
  README.md → AccessPOSH/Private/Session.ps1
- `Get-AccessFormProperty()` --calls--> `ConvertTo-SafeValue()`  [INFERRED]
  AccessPOSH/Public/FormReportOps.ps1 → AccessPOSH/Private/Utilities.ps1
- `Import-AccessVbaFile()` --calls--> `ConvertTo-AnsiTempFile()`  [INFERRED]
  AccessPOSH/Public/VbeOps.ps1 → AccessPOSH/Private/Utilities.ps1
- `ThrowTerminatingError Pattern` --rationale_for--> `DatabaseOps (Public Commands)`  [EXTRACTED]
  AccessPOSH/EVALUATION.md → AccessPOSH/Public/DatabaseOps.ps1
- `COM Object Disposal Pattern` --rationale_for--> `Get-TableSchemaDDL`  [INFERRED]
  AccessPOSH/EVALUATION.md → AccessPOSH/Private/Utilities.ps1

## Hyperedges (group relationships)
- **COM Session Lifecycle Management** — session_connect_accessdb, session_test_alive, session_get_running_com_app, session_clear_caches, session_access_session_state [EXTRACTED 0.95]
- **Design View Open-Edit-Save Workflow** — designview_open, designview_save_close, designview_parse_controls, formreportops, reportops [EXTRACTED 0.95]
- **SaveAsText Export Pipeline** — exporttransfer, binaryhelpers_remove_sections, utilities_read_temp_file, utilities_write_temp_file, utilities_get_table_schema_ddl [EXTRACTED 0.95]
- **VBE Code Mutation Pipeline** — vbeops_set_access_vbe_line, vbeops_set_access_vbe_proc, vbeops_update_access_vbe_proc, vbeops_add_access_vbe_code, pattern_vbe_code_caching [EXTRACTED 0.95]
- **Win32/GDI UI Automation Subsystem** — uiautomation_accessposhui, uiautomation_get_access_screenshot, uiautomation_send_access_click, uiautomation_send_access_keyboard [EXTRACTED 0.95]
- **Cross-Database Search Functions** — vbeops_search_access_vbe, vbeops_search_access_query, vbeops_find_access_usage [EXTRACTED 0.95]

## Communities (45 total, 18 thin omitted)

### Community 0 - "Public Module Files"
Cohesion: 0.08
Nodes (49): Set-FieldProperty(), Resolve-SessionDbPath(), Format-AccessOutput(), Get-AccessApplicationInfo(), Get-AccessFileInfo(), Test-AccessRuntime(), New-AccessForm(), Export-AccessToExcel() (+41 more)

### Community 1 - "Core Operations Logic"
Cohesion: 0.07
Nodes (42): ApplicationOps (Public Commands), Remove-BinarySections, Restore-BinarySections, Set-FieldProperty, DatabaseOps (Public Commands), Open-InDesignView, Save-AndCloseDesign, COM Object Disposal Pattern (+34 more)

### Community 2 - "VBE & Code Helpers"
Cohesion: 0.19
Nodes (24): Connect-AccessDB(), Test-VbaFileEncoding(), Get-AllModuleCode(), Get-ClosestMatchContext(), Get-CodeModule(), Test-TextMatch(), Test-WsNormalizedMatch(), Add-AccessVbeCode() (+16 more)

### Community 3 - "Design View & Forms"
Cohesion: 0.17
Nodes (19): ConvertFrom-ControlBlock(), Get-ParsedControls(), Open-InDesignView(), Save-AndCloseDesign(), ConvertTo-CoercedProp(), Get-AccessControl(), Get-AccessControlDetail(), Get-AccessFormProperty() (+11 more)

### Community 4 - "Binary & Export Helpers"
Cohesion: 0.18
Nodes (16): Get-BinaryBlocks(), Invoke-VbaAfterImport(), Remove-BinarySections(), Restore-BinarySections(), Split-CodeBehind(), ConvertTo-AnsiTempFile(), ConvertTo-SafeValue(), Get-TableSchemaDDL() (+8 more)

### Community 5 - "COM Session Management"
Cohesion: 0.18
Nodes (15): Clear-AccessCaches(), Get-AccessHwnd(), Get-RunningComApp(), Set-AccessVisibleBestEffort(), Test-AccessAlive(), Close-AccessDatabase(), Export-AccessStructure(), Get-AccessObject() (+7 more)

### Community 6 - "VBE Code Editing"
Cohesion: 0.15
Nodes (18): VBE Code Caching Strategy, Whitespace-Normalized Patch Matching, VbeOps Tests, Add-AccessVbeCode, Find-AccessUsage, Find-AccessVbeText, Get-AccessVbeLine, Get-AccessVbeProc (+10 more)

### Community 7 - "UI Automation & Win32"
Cohesion: 0.43
Nodes (7): PrintWindow Screenshot Capture, AccessPOSH Module Tests, UIAutomation Tests, AccessPoshUI (C# Interop Class), Get-AccessScreenshot, Send-AccessClick, Send-AccessKeyboard

### Community 8 - "Theme Operations"
Cohesion: 0.83
Nodes (4): ThemeOps Tests, Get-AccessTheme, Get-AccessThemeList, Set-AccessTheme

## Knowledge Gaps
- **44 isolated node(s):** `MCP-Access Origin (unmateria)`, `AccessPOSH Quality Evaluation Report`, `Parameter Validation Design Decision`, `ThrowTerminatingError Pattern`, `Session Singleton Pattern` (+39 more)
  These have ≤1 connection - possible missing edges or undocumented components.
- **18 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `Format-AccessOutput()` connect `Public Module Files` to `VBE & Code Helpers`, `Design View & Forms`, `Binary & Export Helpers`, `COM Session Management`?**
  _High betweenness centrality (0.102) - this node is a cross-community bridge._
- **Why does `Resolve-SessionDbPath()` connect `Public Module Files` to `VBE & Code Helpers`, `Design View & Forms`, `Binary & Export Helpers`, `COM Session Management`?**
  _High betweenness centrality (0.096) - this node is a cross-community bridge._
- **Why does `Connect-AccessDB()` connect `VBE & Code Helpers` to `Public Module Files`, `Design View & Forms`, `Binary & Export Helpers`, `COM Session Management`?**
  _High betweenness centrality (0.095) - this node is a cross-community bridge._
- **Are the 90 inferred relationships involving `Format-AccessOutput()` (e.g. with `Get-AccessApplicationInfo()` and `Test-AccessRuntime()`) actually correct?**
  _`Format-AccessOutput()` has 90 INFERRED edges - model-reasoned connections that need verification._
- **Are the 88 inferred relationships involving `Resolve-SessionDbPath()` (e.g. with `Get-ParsedControls()` and `Get-AccessApplicationInfo()`) actually correct?**
  _`Resolve-SessionDbPath()` has 88 INFERRED edges - model-reasoned connections that need verification._
- **Are the 84 inferred relationships involving `Connect-AccessDB()` (e.g. with `Get-AccessApplicationInfo()` and `Test-AccessRuntime()`) actually correct?**
  _`Connect-AccessDB()` has 84 INFERRED edges - model-reasoned connections that need verification._
- **Are the 11 inferred relationships involving `Open-InDesignView()` (e.g. with `Get-AccessFormProperty()` and `Set-AccessFormProperty()`) actually correct?**
  _`Open-InDesignView()` has 11 INFERRED edges - model-reasoned connections that need verification._