# Tests/DatabaseOps.Tests.ps1
# Parameter-validation tests for DatabaseOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\AccessPOSH\AccessPOSH.psd1'
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Close-AccessDatabase' {
    It 'Has CmdletBinding' {
        (Get-Command Close-AccessDatabase).CmdletBinding | Should -BeTrue
    }
}

Describe 'New-AccessDatabase' {
    It 'Has CmdletBinding' {
        (Get-Command New-AccessDatabase).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter' {
        (Get-Command New-AccessDatabase).Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Repair-AccessDatabase' {
    It 'Has CmdletBinding' {
        (Get-Command Repair-AccessDatabase).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter' {
        (Get-Command Repair-AccessDatabase).Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Invoke-AccessDecompile' {
    It 'Has CmdletBinding' {
        (Get-Command Invoke-AccessDecompile).CmdletBinding | Should -BeTrue
    }
    It 'Has DbPath parameter' {
        (Get-Command Invoke-AccessDecompile).Parameters['DbPath'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-AccessObject' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessObject).CmdletBinding | Should -BeTrue
    }
    It 'Has ObjectType parameter' {
        (Get-Command Get-AccessObject).Parameters['ObjectType'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-AccessCode' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessCode).CmdletBinding | Should -BeTrue
    }
    It 'Has Name parameter' {
        (Get-Command Get-AccessCode).Parameters['Name'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Set-AccessCode' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessCode).CmdletBinding | Should -BeTrue
    }
    It 'Has Name parameter' {
        (Get-Command Set-AccessCode).Parameters['Name'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Code parameter' {
        (Get-Command Set-AccessCode).Parameters['Code'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Remove-AccessObject' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-AccessObject).CmdletBinding | Should -BeTrue
    }
    It 'Has Name parameter' {
        (Get-Command Remove-AccessObject).Parameters['Name'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ObjectType parameter' {
        (Get-Command Remove-AccessObject).Parameters['ObjectType'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Export-AccessStructure' {
    It 'Has CmdletBinding' {
        (Get-Command Export-AccessStructure).CmdletBinding | Should -BeTrue
    }
    It 'Has OutputPath parameter' {
        (Get-Command Export-AccessStructure).Parameters['OutputPath'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Invoke-AccessSQL' {
    It 'Has CmdletBinding' {
        (Get-Command Invoke-AccessSQL).CmdletBinding | Should -BeTrue
    }
    It 'Throws when -SQL is omitted' {
        { Invoke-AccessSQL -DbPath 'x:\fake.accdb' } | Should -Throw '*-SQL is required*'
    }
}

Describe 'Invoke-AccessSQLBatch' {
    It 'Has CmdletBinding' {
        (Get-Command Invoke-AccessSQLBatch).CmdletBinding | Should -BeTrue
    }
    It 'Throws when -Statements is omitted' {
        { Invoke-AccessSQLBatch -DbPath 'x:\fake.accdb' } | Should -Throw '*-Statements is required*'
    }
}

Describe 'New-AccessDatabase — Parameter Validation' {
    It 'throws when -DbPath is empty' {
        { New-AccessDatabase -DbPath '' } | Should -Throw '*-DbPath is required*'
    }
    It 'throws when -DbPath is not supplied' {
        { New-AccessDatabase } | Should -Throw '*-DbPath is required*'
    }
}

Describe 'Repair-AccessDatabase — Parameter Validation' {
    It 'throws when -DbPath is empty' {
        { Repair-AccessDatabase -DbPath '' } | Should -Throw '*-DbPath is required*'
    }
}

Describe 'Invoke-AccessDecompile — Parameter Validation' {
    It 'throws when -DbPath is empty' {
        { Invoke-AccessDecompile -DbPath '' } | Should -Throw '*-DbPath is required*'
    }
}

Describe 'Get-AccessCode — Parameter Validation' {
    It 'throws when -ObjectType is omitted' {
        { Get-AccessCode -DbPath 'x:\fake.accdb' } | Should -Throw '*-ObjectType is required*'
    }
    It 'throws when -Name is omitted' {
        { Get-AccessCode -DbPath 'x:\fake.accdb' -ObjectType 'module' } | Should -Throw '*-Name is required*'
    }
}

Describe 'Set-AccessCode — Parameter Validation' {
    It 'throws when -ObjectType is omitted' {
        { Set-AccessCode -DbPath 'x:\fake.accdb' } | Should -Throw '*-ObjectType is required*'
    }
    It 'throws when -Code is omitted' {
        { Set-AccessCode -DbPath 'x:\fake.accdb' -ObjectType 'module' -Name 'M1' } | Should -Throw '*-Code is required*'
    }
}

Describe 'Remove-AccessObject — Parameter Validation' {
    It 'throws when -ObjectType is omitted' {
        { Remove-AccessObject -DbPath 'x:\fake.accdb' } | Should -Throw '*-ObjectType is required*'
    }
    It 'throws when -Name is omitted' {
        { Remove-AccessObject -DbPath 'x:\fake.accdb' -ObjectType 'module' } | Should -Throw '*-Name is required*'
    }
}
