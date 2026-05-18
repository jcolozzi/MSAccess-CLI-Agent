# Tests/MetadataOps.Tests.ps1
# Parameter-validation tests for MetadataOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\AccessPOSH\AccessPOSH.psd1'
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Get-AccessLinkedTable' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessLinkedTable).CmdletBinding | Should -BeTrue
    }
}

Describe 'Set-AccessLinkedTable' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessLinkedTable).CmdletBinding | Should -BeTrue
    }
    It 'Has TableName parameter' {
        (Get-Command Set-AccessLinkedTable).Parameters['TableName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has NewConnect parameter' {
        (Get-Command Set-AccessLinkedTable).Parameters['NewConnect'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-AccessRelationship' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessRelationship).CmdletBinding | Should -BeTrue
    }
}

Describe 'New-AccessRelationship' {
    It 'Has CmdletBinding' {
        (Get-Command New-AccessRelationship).CmdletBinding | Should -BeTrue
    }
    It 'Has Name parameter' {
        (Get-Command New-AccessRelationship).Parameters['Name'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Table parameter' {
        (Get-Command New-AccessRelationship).Parameters['Table'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ForeignTable parameter' {
        (Get-Command New-AccessRelationship).Parameters['ForeignTable'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Fields parameter' {
        (Get-Command New-AccessRelationship).Parameters['Fields'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Remove-AccessRelationship' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-AccessRelationship).CmdletBinding | Should -BeTrue
    }
    It 'Throws when -Name is omitted' {
        { Remove-AccessRelationship -DbPath 'x:\fake.accdb' } | Should -Throw '*-Name is required*'
    }
}

Describe 'Get-AccessReference' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessReference).CmdletBinding | Should -BeTrue
    }
}

Describe 'Set-AccessReference' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessReference).CmdletBinding | Should -BeTrue
    }
    It 'Throws when -Action is omitted' {
        { Set-AccessReference -DbPath 'x:\fake.accdb' } | Should -Throw '*-Action is required*'
    }
}

Describe 'Set-AccessQuery' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessQuery).CmdletBinding | Should -BeTrue
    }
    It 'Has QueryName parameter' {
        (Get-Command Set-AccessQuery).Parameters['QueryName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Sql parameter' {
        (Get-Command Set-AccessQuery).Parameters['Sql'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-AccessStartupOption' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessStartupOption).CmdletBinding | Should -BeTrue
    }
}

Describe 'Get-AccessDatabaseProperty' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessDatabaseProperty).CmdletBinding | Should -BeTrue
    }
}

Describe 'Set-AccessDatabaseProperty' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessDatabaseProperty).CmdletBinding | Should -BeTrue
    }
    It 'Has Name parameter' {
        (Get-Command Set-AccessDatabaseProperty).Parameters['Name'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Value parameter' {
        (Get-Command Set-AccessDatabaseProperty).Parameters['Value'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-AccessTip' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessTip).CmdletBinding | Should -BeTrue
    }
}

Describe 'Set-AccessLinkedTable — Parameter Validation' {
    It 'throws when -TableName is omitted' {
        { Set-AccessLinkedTable -DbPath 'x:\fake.accdb' } | Should -Throw '*-TableName is required*'
    }
    It 'throws when -NewConnect is omitted' {
        { Set-AccessLinkedTable -DbPath 'x:\fake.accdb' -TableName 'T' } | Should -Throw '*-NewConnect is required*'
    }
}

Describe 'New-AccessRelationship — Parameter Validation' {
    It 'throws when -Name is omitted' {
        { New-AccessRelationship -DbPath 'x:\fake.accdb' } | Should -Throw '*-Name is required*'
    }
    It 'throws when -Table is omitted' {
        { New-AccessRelationship -DbPath 'x:\fake.accdb' -Name 'Rel1' } | Should -Throw '*-Table is required*'
    }
    It 'throws when -ForeignTable is omitted' {
        { New-AccessRelationship -DbPath 'x:\fake.accdb' -Name 'Rel1' -Table 'T1' } | Should -Throw '*-ForeignTable is required*'
    }
    It 'throws when -Fields is empty' {
        { New-AccessRelationship -DbPath 'x:\fake.accdb' -Name 'Rel1' -Table 'T1' -ForeignTable 'T2' -Fields @() } | Should -Throw '*-Fields is required*'
    }
}

Describe 'Set-AccessQuery — Parameter Validation' {
    It 'throws when -Action is omitted' {
        { Set-AccessQuery -DbPath 'x:\fake.accdb' } | Should -Throw '*-Action is required*'
    }
    It 'throws when -QueryName is omitted' {
        { Set-AccessQuery -DbPath 'x:\fake.accdb' -Action 'create' } | Should -Throw '*-QueryName is required*'
    }
}

Describe 'Set-AccessDatabaseProperty — Parameter Validation' {
    It 'throws when -Name is omitted' {
        { Set-AccessDatabaseProperty -DbPath 'x:\fake.accdb' } | Should -Throw '*-Name is required*'
    }
}
