# Tests/VbeOps.Tests.ps1
# Parameter-validation tests for VbeOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\AccessPOSH\AccessPOSH.psd1'
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module AccessPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Get-AccessVbeLine' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessVbeLine).CmdletBinding | Should -BeTrue
    }
    It 'Has ObjectName parameter' {
        (Get-Command Get-AccessVbeLine).Parameters['ObjectName'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-AccessVbeProc' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessVbeProc).CmdletBinding | Should -BeTrue
    }
    It 'Has ObjectName parameter' {
        (Get-Command Get-AccessVbeProc).Parameters['ObjectName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ProcName parameter' {
        (Get-Command Get-AccessVbeProc).Parameters['ProcName'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-AccessVbeModuleInfo' {
    It 'Has CmdletBinding' {
        (Get-Command Get-AccessVbeModuleInfo).CmdletBinding | Should -BeTrue
    }
}

Describe 'Set-AccessVbeLine' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessVbeLine).CmdletBinding | Should -BeTrue
    }
    It 'Has ObjectName parameter' {
        (Get-Command Set-AccessVbeLine).Parameters['ObjectName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has StartLine parameter' {
        (Get-Command Set-AccessVbeLine).Parameters['StartLine'] | Should -Not -BeNullOrEmpty
    }
    It 'Has NewCode parameter' {
        (Get-Command Set-AccessVbeLine).Parameters['NewCode'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Set-AccessVbeProc' {
    It 'Has CmdletBinding' {
        (Get-Command Set-AccessVbeProc).CmdletBinding | Should -BeTrue
    }
    It 'Has ObjectName parameter' {
        (Get-Command Set-AccessVbeProc).Parameters['ObjectName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ProcName parameter' {
        (Get-Command Set-AccessVbeProc).Parameters['ProcName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has NewCode parameter' {
        (Get-Command Set-AccessVbeProc).Parameters['NewCode'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Update-AccessVbeProc' {
    It 'Has CmdletBinding' {
        (Get-Command Update-AccessVbeProc).CmdletBinding | Should -BeTrue
    }
    It 'Has ObjectName parameter' {
        (Get-Command Update-AccessVbeProc).Parameters['ObjectName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ProcName parameter' {
        (Get-Command Update-AccessVbeProc).Parameters['ProcName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Patches parameter' {
        (Get-Command Update-AccessVbeProc).Parameters['Patches'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Add-AccessVbeCode' {
    It 'Has CmdletBinding' {
        (Get-Command Add-AccessVbeCode).CmdletBinding | Should -BeTrue
    }
    It 'Has ObjectName parameter' {
        (Get-Command Add-AccessVbeCode).Parameters['ObjectName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Code parameter' {
        (Get-Command Add-AccessVbeCode).Parameters['Code'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Find-AccessVbeText' {
    It 'Has CmdletBinding' {
        (Get-Command Find-AccessVbeText).CmdletBinding | Should -BeTrue
    }
    It 'Has SearchText parameter' {
        (Get-Command Find-AccessVbeText).Parameters['SearchText'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Search-AccessVbe' {
    It 'Has CmdletBinding' {
        (Get-Command Search-AccessVbe).CmdletBinding | Should -BeTrue
    }
    It 'Has SearchText parameter' {
        (Get-Command Search-AccessVbe).Parameters['SearchText'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Search-AccessQuery' {
    It 'Has CmdletBinding' {
        (Get-Command Search-AccessQuery).CmdletBinding | Should -BeTrue
    }
    It 'Has SearchText parameter' {
        (Get-Command Search-AccessQuery).Parameters['SearchText'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Find-AccessUsage' {
    It 'Has CmdletBinding' {
        (Get-Command Find-AccessUsage).CmdletBinding | Should -BeTrue
    }
    It 'Has SearchText parameter' {
        (Get-Command Find-AccessUsage).Parameters['SearchText'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Invoke-AccessMacro' {
    It 'Has CmdletBinding' {
        (Get-Command Invoke-AccessMacro).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory MacroName parameter' {
        $p = (Get-Command Invoke-AccessMacro).Parameters['MacroName']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -Contain $true
    }
}

Describe 'Invoke-AccessVba' {
    It 'Has CmdletBinding' {
        (Get-Command Invoke-AccessVba).CmdletBinding | Should -BeTrue
    }
    It 'Has Procedure parameter' {
        (Get-Command Invoke-AccessVba).Parameters['Procedure'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Invoke-AccessEval' {
    It 'Has CmdletBinding' {
        (Get-Command Invoke-AccessEval).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory Expression parameter' {
        $p = (Get-Command Invoke-AccessEval).Parameters['Expression']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -Contain $true
    }
}

Describe 'Test-AccessVbaCompile' {
    It 'Has CmdletBinding' {
        (Get-Command Test-AccessVbaCompile).CmdletBinding | Should -BeTrue
    }
}
