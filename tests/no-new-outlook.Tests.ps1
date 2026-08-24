#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1' }
# run with: Invoke-Pester tests

BeforeAll {
    $repo = Split-Path $PSScriptRoot -Parent
}

Describe 'script files' {
    It 'parses <_>' -ForEach 'Remove-NewOutlook.ps1', 'Restore-NewOutlook.ps1', 'Block-NewOutlookAppLocker.ps1' {
        $errs = $null
        [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $repo $_), [ref]$null, [ref]$errs) | Out-Null
        $errs.Count | Should -Be 0
    }
}
