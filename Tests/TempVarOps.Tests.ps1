# Tests/TempVarOps.Tests.ps1
# Parameter-validation tests for TempVarOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\AccessPOSH\AccessPOSH.psd1'
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Get-AccessTempVar' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessTempVar).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (mandatory)' {
        $p = (Get-Command Get-AccessTempVar).Parameters['DbPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has Name parameter (optional)' {
        (Get-Command Get-AccessTempVar).Parameters['Name'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        $p = (Get-Command Get-AccessTempVar).Parameters['AsJson']
        $p | Should -Not -BeNullOrEmpty
        $p.SwitchParameter | Should -BeTrue
    }
}

Describe 'Set-AccessTempVar' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessTempVar).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (mandatory)' {
        $p = (Get-Command Set-AccessTempVar).Parameters['DbPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has Name parameter (mandatory)' {
        $p = (Get-Command Set-AccessTempVar).Parameters['Name']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has Value parameter (mandatory)' {
        $p = (Get-Command Set-AccessTempVar).Parameters['Value']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has AsJson switch' {
        $p = (Get-Command Set-AccessTempVar).Parameters['AsJson']
        $p | Should -Not -BeNullOrEmpty
        $p.SwitchParameter | Should -BeTrue
    }
}

Describe 'Remove-AccessTempVar' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-AccessTempVar).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (mandatory)' {
        $p = (Get-Command Remove-AccessTempVar).Parameters['DbPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has Name parameter (optional)' {
        (Get-Command Remove-AccessTempVar).Parameters['Name'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        $p = (Get-Command Remove-AccessTempVar).Parameters['AsJson']
        $p | Should -Not -BeNullOrEmpty
        $p.SwitchParameter | Should -BeTrue
    }
}
