# Tests/SecurityOps.Tests.ps1
# Parameter-validation tests for SecurityOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\AccessPOSH\AccessPOSH.psd1'
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Test-AccessDatabasePassword' {
    It 'Has CmdletBinding' {
        (Get-Command Test-AccessDatabasePassword).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (mandatory)' {
        $p = (Get-Command Test-AccessDatabasePassword).Parameters['DbPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has AsJson switch' {
        $p = (Get-Command Test-AccessDatabasePassword).Parameters['AsJson']
        $p | Should -Not -BeNullOrEmpty
        $p.SwitchParameter | Should -BeTrue
    }
}

Describe 'Set-AccessDatabasePassword' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessDatabasePassword).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (mandatory)' {
        $p = (Get-Command Set-AccessDatabasePassword).Parameters['DbPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has NewPassword parameter (mandatory)' {
        $p = (Get-Command Set-AccessDatabasePassword).Parameters['NewPassword']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has OldPassword parameter (optional)' {
        (Get-Command Set-AccessDatabasePassword).Parameters['OldPassword'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Set-AccessDatabasePassword).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Remove-AccessDatabasePassword' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-AccessDatabasePassword).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (mandatory)' {
        $p = (Get-Command Remove-AccessDatabasePassword).Parameters['DbPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has CurrentPassword parameter (mandatory)' {
        $p = (Get-Command Remove-AccessDatabasePassword).Parameters['CurrentPassword']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has AsJson switch' {
        (Get-Command Remove-AccessDatabasePassword).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Get-AccessDatabaseEncryption' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessDatabaseEncryption).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter (mandatory)' {
        $p = (Get-Command Get-AccessDatabaseEncryption).Parameters['DbPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has AsJson switch' {
        $p = (Get-Command Get-AccessDatabaseEncryption).Parameters['AsJson']
        $p | Should -Not -BeNullOrEmpty
        $p.SwitchParameter | Should -BeTrue
    }
}
