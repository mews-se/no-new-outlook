#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1' }
# run with: Invoke-Pester tests
# the append/strip/sid/cleanup blocks mirror the logic in the scripts; keep them in sync

BeforeDiscovery {
    # the cleanup block needs a real registry, so it only runs on windows
    $onWindows = $PSVersionTable.PSEdition -eq 'Desktop' -or $IsWindows
}

BeforeAll {
    $repo = Split-Path $PSScriptRoot -Parent

    function Remove-EmptyKeys($leaf, $stopAt) {
        $path = $leaf
        while ($path -like "$stopAt\*") {
            $key = Get-Item $path -ErrorAction SilentlyContinue
            if (-not $key -or $key.ValueCount -or $key.SubKeyCount) { break }
            Remove-Item $path -Force
            $path = Split-Path $path -Parent
        }
    }

    function Add-BlockEntry([string]$blocked) {
        $blocked = $blocked.Trim()
        if ($blocked -notmatch '^\[.+\]$') { return '["MS_Outlook"]' }
        if ($blocked -notmatch '"MS_Outlook"') { return $blocked -replace '\]$', ',"MS_Outlook"]' }
        return $null
    }

    function Remove-BlockEntry([string]$blocked) {
        $blocked = $blocked.Trim()
        if ($blocked -notmatch '"MS_Outlook"') { return '<untouched>' }
        $blocked = $blocked -replace '\s*"MS_Outlook"\s*,?', '' -replace ',\s*\]', ']'
        if ($blocked -match '^\[\s*\]$') { return '<deleted>' }
        return $blocked
    }

    $sidPattern = '^S-1-(5-21|12-1)(-\d+){4}$'
}

Describe 'script files' {
    It 'parses <_>' -ForEach 'Remove-NewOutlook.ps1', 'Restore-NewOutlook.ps1' {
        $errs = $null
        [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $repo $_), [ref]$null, [ref]$errs) | Out-Null
        $errs.Count | Should -Be 0
    }
}

Describe 'BlockedOobeUpdaters append' {
    It 'creates the list when the value is missing or malformed' {
        Add-BlockEntry '' | Should -Be '["MS_Outlook"]'
        Add-BlockEntry 'MS_Outlook' | Should -Be '["MS_Outlook"]'
    }
    It 'appends without touching existing entries' {
        Add-BlockEntry '["MS_DevHome"]' | Should -Be '["MS_DevHome","MS_Outlook"]'
        Add-BlockEntry '["Other"] ' | Should -Be '["Other","MS_Outlook"]'
        Add-BlockEntry '["A", "B"]' | Should -Be '["A", "B","MS_Outlook"]'
    }
    It 'leaves the value alone when already present' {
        Add-BlockEntry '["MS_Outlook"]' | Should -BeNullOrEmpty
        Add-BlockEntry '["Other", "MS_Outlook"]' | Should -BeNullOrEmpty
    }
}

Describe 'BlockedOobeUpdaters strip' {
    It 'removes only our entry' {
        Remove-BlockEntry '["MS_Outlook","X"]' | Should -Be '["X"]'
        Remove-BlockEntry '["X","MS_Outlook"]' | Should -Be '["X"]'
        Remove-BlockEntry '["A","MS_Outlook","B"]' | Should -Be '["A","B"]'
        Remove-BlockEntry '["Other", "MS_Outlook"]' | Should -Be '["Other"]'
    }
    It 'drops the value when the list ends up empty' {
        Remove-BlockEntry '["MS_Outlook"]' | Should -Be '<deleted>'
        Remove-BlockEntry '[ "MS_Outlook" ]' | Should -Be '<deleted>'
        Remove-BlockEntry '["MS_Outlook","MS_Outlook"]' | Should -Be '<deleted>'
    }
    It 'ignores foreign values' {
        Remove-BlockEntry '["MS_DevHome"]' | Should -Be '<untouched>'
        Remove-BlockEntry '' | Should -Be '<untouched>'
    }
}

Describe 'empty key cleanup' -Skip:(-not $onWindows) {
    BeforeAll { $root = 'HKCU:\Software\no-new-outlook-tests' }
    BeforeEach {
        Remove-Item $root -Recurse -Force -ErrorAction SilentlyContinue
        New-Item "$root\a\b\c" -Force | Out-Null
    }
    AfterAll { Remove-Item $root -Recurse -Force -ErrorAction SilentlyContinue }

    It 'walks up the empty chain and stops at the stop key' {
        Remove-EmptyKeys "$root\a\b\c" $root
        Test-Path "$root\a" | Should -BeFalse
        Test-Path $root | Should -BeTrue
    }
    It 'keeps a leaf that still holds a value' {
        Set-ItemProperty "$root\a\b\c" -Name Keep -Value 1 -Type DWord
        Remove-EmptyKeys "$root\a\b\c" $root
        Test-Path "$root\a\b\c" | Should -BeTrue
    }
    It 'stops below a parent that still holds a value' {
        Set-ItemProperty "$root\a" -Name Keep -Value 1 -Type DWord
        Remove-EmptyKeys "$root\a\b\c" $root
        Test-Path "$root\a\b" | Should -BeFalse
        Test-Path "$root\a" | Should -BeTrue
    }
    It 'stops below a parent that still holds a subkey' {
        New-Item "$root\a\b\sibling" -Force | Out-Null
        Remove-EmptyKeys "$root\a\b\c" $root
        Test-Path "$root\a\b\c" | Should -BeFalse
        Test-Path "$root\a\b" | Should -BeTrue
    }
    It 'shrugs at a leaf that was never there' {
        { Remove-EmptyKeys "$root\gone\missing" $root } | Should -Not -Throw
        Test-Path $root | Should -BeTrue
    }
}

Describe 'user hive sid filter' {
    It 'matches local and Entra user sids' {
        'S-1-5-21-1004336348-1177238915-682003330-1001' | Should -Match $sidPattern
        'S-1-12-1-3123456789-1234567890-123456789-1234567890' | Should -Match $sidPattern
    }
    It 'rejects everything else under HKEY_USERS' {
        'S-1-5-21-1004336348-1177238915-682003330-1001_Classes' | Should -Not -Match $sidPattern
        'S-1-5-18' | Should -Not -Match $sidPattern
        'S-1-5-19' | Should -Not -Match $sidPattern
        '.DEFAULT' | Should -Not -Match $sidPattern
    }
}
