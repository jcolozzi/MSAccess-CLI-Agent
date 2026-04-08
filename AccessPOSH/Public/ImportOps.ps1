# Public/ImportOps.ps1 — Import data from external sources (Excel, CSV, text, XML, other databases)

function Import-AccessFromExcel {
    <#
    .SYNOPSIS
        Import an Excel worksheet into an Access table.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][string]$ExcelPath,
        [Parameter(Mandatory)][string]$TableName,
        [string]$SheetName,
        [switch]$HasFieldNames,
        [ValidateSet('xlsx','xls')]
        [string]$SpreadsheetType = 'xlsx',
        [switch]$AsJson
    )
    $app = Connect-AccessDB -DbPath $DbPath
    if (-not (Test-Path $ExcelPath)) { throw "Excel file not found: $ExcelPath" }
    $ExcelPath = (Resolve-Path $ExcelPath).Path
    $typeMap = @{ xlsx = 10; xls = 8 }
    $acType = $typeMap[$SpreadsheetType]
    $rangeArg = if ($SheetName) { $SheetName } else { [System.Reflection.Missing]::Value }
    $app.DoCmd.TransferSpreadsheet(
        0,
        $acType,
        $TableName,
        $ExcelPath,
        [bool]$HasFieldNames,
        $rangeArg
    )
    $result = [ordered]@{ action = 'imported'; source = $ExcelPath; table = $TableName; format = $SpreadsheetType }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Import-AccessFromCSV {
    <#
    .SYNOPSIS
        Import a CSV or delimited text file into an Access table.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string]$TableName,
        [switch]$HasFieldNames,
        [string]$SpecificationName,
        [switch]$AsJson
    )
    $app = Connect-AccessDB -DbPath $DbPath
    if (-not (Test-Path $FilePath)) { throw "File not found: $FilePath" }
    $FilePath = (Resolve-Path $FilePath).Path
    $specArg = if ($SpecificationName) { $SpecificationName } else { [System.Reflection.Missing]::Value }
    $app.DoCmd.TransferText(
        0,
        $specArg,
        $TableName,
        $FilePath,
        [bool]$HasFieldNames
    )
    $result = [ordered]@{ action = 'imported'; source = $FilePath; table = $TableName; format = 'csv' }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Import-AccessFromXML {
    <#
    .SYNOPSIS
        Import an XML file into the Access database.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][string]$XmlPath,
        [ValidateSet('structureonly','dataonly','structureanddata')]
        [string]$ImportOptions = 'structureanddata',
        [switch]$AsJson
    )
    $app = Connect-AccessDB -DbPath $DbPath
    if (-not (Test-Path $XmlPath)) { throw "XML file not found: $XmlPath" }
    $XmlPath = (Resolve-Path $XmlPath).Path
    $optMap = @{ structureonly = 0; structureanddata = 1; dataonly = 2 }
    $app.ImportXML($XmlPath, $optMap[$ImportOptions])
    $result = [ordered]@{ action = 'imported'; source = $XmlPath; import_option = $ImportOptions }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Import-AccessFromDatabase {
    <#
    .SYNOPSIS
        Import a table or query from another Access database.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][string]$SourceDbPath,
        [Parameter(Mandatory)][string]$SourceObject,
        [string]$DestinationTable,
        [ValidateSet('table','query')]
        [string]$ObjectType = 'table',
        [switch]$StructureOnly,
        [switch]$AsJson
    )
    $app = Connect-AccessDB -DbPath $DbPath
    if (-not (Test-Path $SourceDbPath)) { throw "Source database not found: $SourceDbPath" }
    $SourceDbPath = (Resolve-Path $SourceDbPath).Path
    $destName = if ($DestinationTable) { $DestinationTable } else { $SourceObject }
    $objectTypeConst = if ($ObjectType -eq 'query') { 1 } else { 0 }
    $dataOnly = if ($StructureOnly) { $true } else { $false }
    $app.DoCmd.TransferDatabase(
        0,
        'Microsoft Access',
        $SourceDbPath,
        $objectTypeConst,
        $SourceObject,
        $destName,
        $dataOnly
    )
    $result = [ordered]@{ action = 'imported'; source_db = $SourceDbPath; source_object = $SourceObject; destination_table = $destName; object_type = $ObjectType; structure_only = $dataOnly }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}

function Export-AccessToExcel {
    <#
    .SYNOPSIS
        Export an Access table or query to an Excel workbook.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DbPath,
        [Parameter(Mandatory)][string]$ObjectName,
        [Parameter(Mandatory)][string]$ExcelPath,
        [string]$SheetName,
        [switch]$HasFieldNames,
        [ValidateSet('xlsx','xls')]
        [string]$SpreadsheetType = 'xlsx',
        [switch]$AsJson
    )
    $app = Connect-AccessDB -DbPath $DbPath
    $ExcelPath = [System.IO.Path]::GetFullPath($ExcelPath)
    $typeMap = @{ xlsx = 10; xls = 8 }
    $acType = $typeMap[$SpreadsheetType]
    $rangeArg = if ($SheetName) { $SheetName } else { [System.Reflection.Missing]::Value }
    $app.DoCmd.TransferSpreadsheet(
        1,
        $acType,
        $ObjectName,
        $ExcelPath,
        [bool]$HasFieldNames,
        $rangeArg
    )
    $result = [ordered]@{ action = 'exported'; object = $ObjectName; path = $ExcelPath; format = $SpreadsheetType }
    Format-AccessOutput -AsJson:$AsJson -Data $result
}
