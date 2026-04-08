# Tests/ImportOps.Tests.ps1
# Parameter-validation tests for ImportOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\AccessPOSH\AccessPOSH.psd1'
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Import-AccessFromExcel' {
    It 'Has CmdletBinding' {
        (Get-Command Import-AccessFromExcel).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (mandatory)' {
        $p = (Get-Command Import-AccessFromExcel).Parameters['DbPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has ExcelPath parameter (mandatory)' {
        $p = (Get-Command Import-AccessFromExcel).Parameters['ExcelPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has TableName parameter (mandatory)' {
        $p = (Get-Command Import-AccessFromExcel).Parameters['TableName']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has SpreadsheetType with ValidateSet' {
        $p = (Get-Command Import-AccessFromExcel).Parameters['SpreadsheetType']
        $p | Should -Not -BeNullOrEmpty
        $vs = $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] })
        $vs.Count | Should -BeGreaterThan 0
        $vs[0].ValidValues | Should -Contain 'xlsx'
        $vs[0].ValidValues | Should -Contain 'xls'
    }
    It 'Has HasFieldNames switch' {
        $p = (Get-Command Import-AccessFromExcel).Parameters['HasFieldNames']
        $p | Should -Not -BeNullOrEmpty
        $p.SwitchParameter | Should -BeTrue
    }
    It 'Has AsJson switch' {
        (Get-Command Import-AccessFromExcel).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Import-AccessFromCSV' {
    It 'Has CmdletBinding' {
        (Get-Command Import-AccessFromCSV).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (mandatory)' {
        $p = (Get-Command Import-AccessFromCSV).Parameters['DbPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has FilePath parameter (mandatory)' {
        $p = (Get-Command Import-AccessFromCSV).Parameters['FilePath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has TableName parameter (mandatory)' {
        $p = (Get-Command Import-AccessFromCSV).Parameters['TableName']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has SpecificationName parameter (optional)' {
        (Get-Command Import-AccessFromCSV).Parameters['SpecificationName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Import-AccessFromCSV).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Import-AccessFromXML' {
    It 'Has CmdletBinding' {
        (Get-Command Import-AccessFromXML).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (mandatory)' {
        $p = (Get-Command Import-AccessFromXML).Parameters['DbPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has XmlPath parameter (mandatory)' {
        $p = (Get-Command Import-AccessFromXML).Parameters['XmlPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has ImportOptions with ValidateSet' {
        $p = (Get-Command Import-AccessFromXML).Parameters['ImportOptions']
        $p | Should -Not -BeNullOrEmpty
        $vs = $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] })
        $vs.Count | Should -BeGreaterThan 0
        $vs[0].ValidValues | Should -Contain 'structureanddata'
    }
    It 'Has AsJson switch' {
        (Get-Command Import-AccessFromXML).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Import-AccessFromDatabase' {
    It 'Has CmdletBinding' {
        (Get-Command Import-AccessFromDatabase).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (mandatory)' {
        $p = (Get-Command Import-AccessFromDatabase).Parameters['DbPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has SourceDbPath parameter (mandatory)' {
        $p = (Get-Command Import-AccessFromDatabase).Parameters['SourceDbPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has SourceObject parameter (mandatory)' {
        $p = (Get-Command Import-AccessFromDatabase).Parameters['SourceObject']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has ObjectType with ValidateSet' {
        $p = (Get-Command Import-AccessFromDatabase).Parameters['ObjectType']
        $p | Should -Not -BeNullOrEmpty
        $vs = $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] })
        $vs.Count | Should -BeGreaterThan 0
        $vs[0].ValidValues | Should -Contain 'table'
        $vs[0].ValidValues | Should -Contain 'query'
    }
    It 'Has AsJson switch' {
        (Get-Command Import-AccessFromDatabase).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Export-AccessToExcel' {
    It 'Has CmdletBinding' {
        (Get-Command Export-AccessToExcel).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (mandatory)' {
        $p = (Get-Command Export-AccessToExcel).Parameters['DbPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has ObjectName parameter (mandatory)' {
        $p = (Get-Command Export-AccessToExcel).Parameters['ObjectName']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has ExcelPath parameter (mandatory)' {
        $p = (Get-Command Export-AccessToExcel).Parameters['ExcelPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has SpreadsheetType with ValidateSet' {
        $p = (Get-Command Export-AccessToExcel).Parameters['SpreadsheetType']
        $p | Should -Not -BeNullOrEmpty
        $vs = $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] })
        $vs.Count | Should -BeGreaterThan 0
        $vs[0].ValidValues | Should -Contain 'xlsx'
    }
    It 'Has AsJson switch' {
        (Get-Command Export-AccessToExcel).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}
