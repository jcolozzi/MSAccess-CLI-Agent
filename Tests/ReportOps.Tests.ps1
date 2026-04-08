# Tests/ReportOps.Tests.ps1
# Parameter-validation tests for ReportOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\AccessPOSH\AccessPOSH.psd1'
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'New-AccessReport' {
    It 'Has CmdletBinding' {
        (Get-Command New-AccessReport).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (mandatory)' {
        $p = (Get-Command New-AccessReport).Parameters['DbPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has ReportName parameter (mandatory)' {
        $p = (Get-Command New-AccessReport).Parameters['ReportName']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has RecordSource parameter (optional)' {
        (Get-Command New-AccessReport).Parameters['RecordSource'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command New-AccessReport).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Get-AccessGroupLevel' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessGroupLevel).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (mandatory)' {
        $p = (Get-Command Get-AccessGroupLevel).Parameters['DbPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has ReportName parameter (mandatory)' {
        $p = (Get-Command Get-AccessGroupLevel).Parameters['ReportName']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has AsJson switch' {
        (Get-Command Get-AccessGroupLevel).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Set-AccessGroupLevel' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessGroupLevel).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (mandatory)' {
        $p = (Get-Command Set-AccessGroupLevel).Parameters['DbPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has ReportName parameter (mandatory)' {
        $p = (Get-Command Set-AccessGroupLevel).Parameters['ReportName']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has Expression parameter (mandatory)' {
        $p = (Get-Command Set-AccessGroupLevel).Parameters['Expression']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has SortOrder with ValidateSet' {
        $p = (Get-Command Set-AccessGroupLevel).Parameters['SortOrder']
        $p | Should -Not -BeNullOrEmpty
        $vs = $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] })
        $vs.Count | Should -BeGreaterThan 0
        $vs[0].ValidValues | Should -Contain 'ascending'
        $vs[0].ValidValues | Should -Contain 'descending'
    }
    It 'Has GroupHeader switch' {
        (Get-Command Set-AccessGroupLevel).Parameters['GroupHeader'].SwitchParameter | Should -BeTrue
    }
    It 'Has GroupFooter switch' {
        (Get-Command Set-AccessGroupLevel).Parameters['GroupFooter'].SwitchParameter | Should -BeTrue
    }
    It 'Has AsJson switch' {
        (Get-Command Set-AccessGroupLevel).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Remove-AccessGroupLevel' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-AccessGroupLevel).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (mandatory)' {
        $p = (Get-Command Remove-AccessGroupLevel).Parameters['DbPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has ReportName parameter (mandatory)' {
        $p = (Get-Command Remove-AccessGroupLevel).Parameters['ReportName']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has LevelIndex parameter (mandatory)' {
        $p = (Get-Command Remove-AccessGroupLevel).Parameters['LevelIndex']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has AsJson switch' {
        (Get-Command Remove-AccessGroupLevel).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}
