# Tests/UIAutomation.Tests.ps1
# Parameter-validation tests for UIAutomation functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\AccessPOSH\AccessPOSH.psd1'
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Get-AccessScreenshot' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessScreenshot).CmdletBinding | Should -BeTrue
    }
    It 'Has OutputPath parameter' {
        (Get-Command Get-AccessScreenshot).Parameters['OutputPath'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Send-AccessClick' {
    It 'Has CmdletBinding' {
        (Get-Command Send-AccessClick).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory X parameter' {
        $p = (Get-Command Send-AccessClick).Parameters['X']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -Contain $true
    }
    It 'Has mandatory Y parameter' {
        $p = (Get-Command Send-AccessClick).Parameters['Y']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -Contain $true
    }
}

Describe 'Send-AccessKeyboard' {
    It 'Has CmdletBinding' {
        (Get-Command Send-AccessKeyboard).CmdletBinding | Should -BeTrue
    }
    It 'Has Text parameter' {
        (Get-Command Send-AccessKeyboard).Parameters['Text'] | Should -Not -BeNullOrEmpty
    }
}
