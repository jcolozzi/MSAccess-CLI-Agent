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
    It 'Has mandatory TableName parameter' {
        $p = (Get-Command Get-AccessTableInfo).Parameters['TableName']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -Contain $true
    }
}

Describe 'New-AccessTable' {
    It 'Has CmdletBinding' {
        (Get-Command New-AccessTable).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory TableName parameter' {
        $p = (Get-Command New-AccessTable).Parameters['TableName']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -Contain $true
    }
    It 'Has mandatory Fields parameter' {
        $p = (Get-Command New-AccessTable).Parameters['Fields']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -Contain $true
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
