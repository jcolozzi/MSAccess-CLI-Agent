# Tests/TableOps.Tests.ps1
# Parameter-validation tests for TableOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\AccessPOSH\AccessPOSH.psd1'
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Get-AccessTableInfo' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessTableInfo).CmdletBinding | Should -BeTrue
    }
    It 'Throws when -TableName is omitted' {
        { Get-AccessTableInfo -DbPath 'x:\fake.accdb' } | Should -Throw '*-TableName is required*'
    }
}

Describe 'New-AccessTable' {
    It 'Has CmdletBinding' {
        (Get-Command New-AccessTable).CmdletBinding | Should -BeTrue
    }
    It 'Throws when -TableName is omitted' {
        { New-AccessTable -DbPath 'x:\fake.accdb' } | Should -Throw '*-TableName is required*'
    }
    It 'Throws when -Fields is omitted' {
        { New-AccessTable -DbPath 'x:\fake.accdb' -TableName 'T' } | Should -Throw '*-Fields is required*'
    }
}

Describe 'Edit-AccessTable' {
    It 'Has CmdletBinding' {
        (Get-Command Edit-AccessTable).CmdletBinding | Should -BeTrue
    }
    It 'Has TableName parameter' {
        (Get-Command Edit-AccessTable).Parameters['TableName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Action parameter' {
        (Get-Command Edit-AccessTable).Parameters['Action'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-AccessFieldProperty' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessFieldProperty).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory TableName parameter' {
        (Get-Command Get-AccessFieldProperty).Parameters['TableName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has mandatory FieldName parameter' {
        (Get-Command Get-AccessFieldProperty).Parameters['FieldName'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Set-AccessFieldProperty' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessFieldProperty).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory TableName parameter' {
        (Get-Command Set-AccessFieldProperty).Parameters['TableName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has mandatory FieldName parameter' {
        (Get-Command Set-AccessFieldProperty).Parameters['FieldName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has mandatory PropertyName parameter' {
        (Get-Command Set-AccessFieldProperty).Parameters['PropertyName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has mandatory Value parameter' {
        (Get-Command Set-AccessFieldProperty).Parameters['Value'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-AccessIndex' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessIndex).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory TableName parameter' {
        (Get-Command Get-AccessIndex).Parameters['TableName'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Set-AccessIndex' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessIndex).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory TableName parameter' {
        (Get-Command Set-AccessIndex).Parameters['TableName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has mandatory IndexName parameter' {
        (Get-Command Set-AccessIndex).Parameters['IndexName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has mandatory Fields parameter' {
        (Get-Command Set-AccessIndex).Parameters['Fields'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Edit-AccessTable — Parameter Validation' {
    It 'throws when -TableName is omitted' {
        { Edit-AccessTable -DbPath 'x:\fake.accdb' } | Should -Throw '*-TableName is required*'
    }
    It 'throws when -Action is omitted' {
        { Edit-AccessTable -DbPath 'x:\fake.accdb' -TableName 'T' } | Should -Throw '*-Action is required*'
    }
    It 'throws when -FieldName is omitted' {
        { Edit-AccessTable -DbPath 'x:\fake.accdb' -TableName 'T' -Action 'add_field' } | Should -Throw '*-FieldName is required*'
    }
}

Describe 'Get-AccessFieldProperty — Parameter Validation' {
    It 'throws when -TableName is omitted' {
        { Get-AccessFieldProperty -DbPath 'x:\fake.accdb' } | Should -Throw '*-TableName is required*'
    }
    It 'throws when -FieldName is omitted' {
        { Get-AccessFieldProperty -DbPath 'x:\fake.accdb' -TableName 'T' } | Should -Throw '*-FieldName is required*'
    }
}

Describe 'Set-AccessFieldProperty — Parameter Validation' {
    It 'throws when -TableName is omitted' {
        { Set-AccessFieldProperty -DbPath 'x:\fake.accdb' } | Should -Throw '*-TableName is required*'
    }
    It 'throws when -FieldName is omitted' {
        { Set-AccessFieldProperty -DbPath 'x:\fake.accdb' -TableName 'T' } | Should -Throw '*-FieldName is required*'
    }
    It 'throws when -PropertyName is omitted' {
        { Set-AccessFieldProperty -DbPath 'x:\fake.accdb' -TableName 'T' -FieldName 'F' } | Should -Throw '*-PropertyName is required*'
    }
}

Describe 'Set-AccessIndex — Parameter Validation' {
    It 'throws when -TableName is omitted' {
        { Set-AccessIndex -DbPath 'x:\fake.accdb' } | Should -Throw '*-TableName is required*'
    }
    It 'throws when -Action is omitted' {
        { Set-AccessIndex -DbPath 'x:\fake.accdb' -TableName 'T' } | Should -Throw '*-Action is required*'
    }
    It 'throws when -IndexName is omitted' {
        { Set-AccessIndex -DbPath 'x:\fake.accdb' -TableName 'T' -Action 'create' } | Should -Throw '*-IndexName is required*'
    }
}
