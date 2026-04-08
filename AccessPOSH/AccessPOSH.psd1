@{
    RootModule        = 'AccessPOSH.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'a1b2c3d4-e5f6-7890-abcd-ef1234567890'
    Author            = 'Access-POSH'
    Description       = 'PowerShell Access Database Automation via COM — port of MCP-Access (59 tools)'
    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        # Database
        'Close-AccessDatabase'
        'New-AccessDatabase'
        'Repair-AccessDatabase'
        'Invoke-AccessDecompile'

        # Objects
        'Get-AccessObject'
        'Get-AccessCode'
        'Set-AccessCode'
        'Remove-AccessObject'
        'Export-AccessStructure'

        # SQL
        'Invoke-AccessSQL'
        'Invoke-AccessSQLBatch'

        # Tables
        'Get-AccessTableInfo'
        'New-AccessTable'
        'Edit-AccessTable'
        'Get-AccessFieldProperty'
        'Set-AccessFieldProperty'
        'Get-AccessIndex'
        'Set-AccessIndex'

        # VBE
        'Get-AccessVbeLine'
        'Get-AccessVbeProc'
        'Get-AccessVbeModuleInfo'
        'Set-AccessVbeLine'
        'Set-AccessVbeProc'
        'Update-AccessVbeProc'
        'Add-AccessVbeCode'

        # Search
        'Find-AccessVbeText'
        'Search-AccessVbe'
        'Search-AccessQuery'
        'Find-AccessUsage'

        # VBA Exec
        'Invoke-AccessMacro'
        'Invoke-AccessVba'
        'Invoke-AccessEval'
        'Test-AccessVbaCompile'

        # Forms
        'New-AccessForm'
        'Get-AccessFormProperty'
        'Set-AccessFormProperty'

        # Controls
        'Get-AccessControl'
        'Get-AccessControlDetail'
        'New-AccessControl'
        'Remove-AccessControl'
        'Set-AccessControlProperty'
        'Set-AccessControlBatch'

        # Metadata
        'Get-AccessLinkedTable'
        'Set-AccessLinkedTable'
        'Get-AccessRelationship'
        'New-AccessRelationship'
        'Remove-AccessRelationship'
        'Get-AccessReference'
        'Set-AccessReference'
        'Set-AccessQuery'
        'Get-AccessStartupOption'
        'Get-AccessDatabaseProperty'
        'Set-AccessDatabaseProperty'
        'Get-AccessTip'

        # Export
        'Export-AccessReport'
        'Copy-AccessData'

        # UI Automation
        'Get-AccessScreenshot'
        'Send-AccessClick'
        'Send-AccessKeyboard'
    )

    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()
}
