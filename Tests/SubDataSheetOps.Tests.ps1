# Tests/SubDataSheetOps.Tests.ps1
# Parameter-validation tests for SubDataSheetOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\AccessPOSH\AccessPOSH.psd1'
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Get-AccessSubDataSheet' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessSubDataSheet).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (mandatory)' {
        $p = (Get-Command Get-AccessSubDataSheet).Parameters['DbPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has TableName parameter (mandatory)' {
        $p = (Get-Command Get-AccessSubDataSheet).Parameters['TableName']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has AsJson switch' {
        $p = (Get-Command Get-AccessSubDataSheet).Parameters['AsJson']
        $p | Should -Not -BeNullOrEmpty
        $p.SwitchParameter | Should -BeTrue
    }
}

Describe 'Set-AccessSubDataSheet' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessSubDataSheet).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (mandatory)' {
        $p = (Get-Command Set-AccessSubDataSheet).Parameters['DbPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has TableName parameter (mandatory)' {
        $p = (Get-Command Set-AccessSubDataSheet).Parameters['TableName']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has SubDataSheetName parameter' {
        (Get-Command Set-AccessSubDataSheet).Parameters['SubDataSheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has LinkChildFields parameter' {
        (Get-Command Set-AccessSubDataSheet).Parameters['LinkChildFields'] | Should -Not -BeNullOrEmpty
    }
    It 'Has LinkMasterFields parameter' {
        (Get-Command Set-AccessSubDataSheet).Parameters['LinkMasterFields'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Height parameter' {
        (Get-Command Set-AccessSubDataSheet).Parameters['Height'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        $p = (Get-Command Set-AccessSubDataSheet).Parameters['AsJson']
        $p | Should -Not -BeNullOrEmpty
        $p.SwitchParameter | Should -BeTrue
    }
}
