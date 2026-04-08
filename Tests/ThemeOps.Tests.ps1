# Tests/ThemeOps.Tests.ps1
# Parameter-validation tests for ThemeOps functions (no COM required)

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]param()

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\AccessPOSH\AccessPOSH.psd1'
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Get-AccessTheme' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessTheme).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (mandatory)' {
        $p = (Get-Command Get-AccessTheme).Parameters['DbPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has ObjectName parameter (mandatory)' {
        $p = (Get-Command Get-AccessTheme).Parameters['ObjectName']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has ObjectType with ValidateSet' {
        $p = (Get-Command Get-AccessTheme).Parameters['ObjectType']
        $p | Should -Not -BeNullOrEmpty
        $vs = $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] })
        $vs.Count | Should -BeGreaterThan 0
        $vs[0].ValidValues | Should -Contain 'form'
        $vs[0].ValidValues | Should -Contain 'report'
    }
    It 'Has AsJson switch' {
        (Get-Command Get-AccessTheme).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Set-AccessTheme' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessTheme).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (mandatory)' {
        $p = (Get-Command Set-AccessTheme).Parameters['DbPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has ObjectName parameter (mandatory)' {
        $p = (Get-Command Set-AccessTheme).Parameters['ObjectName']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has ThemeName parameter (mandatory)' {
        $p = (Get-Command Set-AccessTheme).Parameters['ThemeName']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has ObjectType with ValidateSet' {
        $p = (Get-Command Set-AccessTheme).Parameters['ObjectType']
        $vs = $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] })
        $vs[0].ValidValues | Should -Contain 'form'
        $vs[0].ValidValues | Should -Contain 'report'
    }
    It 'Has AsJson switch' {
        (Get-Command Set-AccessTheme).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Get-AccessThemeList' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessThemeList).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (mandatory)' {
        $p = (Get-Command Get-AccessThemeList).Parameters['DbPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has AsJson switch' {
        (Get-Command Get-AccessThemeList).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}
