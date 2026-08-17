# undoes everything Remove-NewOutlook.ps1 changed
# reinstalling the app itself is up to you, see the readme

$ErrorActionPreference = 'Stop'

if ($PSVersionTable.PSVersion.Major -ge 6 -and -not $IsWindows) {
    Write-Warning 'Windows only.'
    exit 1
}
$user = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $user.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning 'Run this from an elevated PowerShell prompt.'
    exit 1
}

# the remove script creates keys on its way in, so removing only the values it
# wrote leaves empty shells behind. take those back out, bottom up, and stop at
# the first key that still holds values or subkeys: it is not ours to delete
function Remove-EmptyKeys($leaf, $stopAt) {
    $path = $leaf
    while ($path -like "$stopAt\*") {
        $key = Get-Item $path -ErrorAction SilentlyContinue
        if (-not $key -or $key.ValueCount -or $key.SubKeyCount) { break }
        Remove-Item $path -Force
        $path = Split-Path $path -Parent
    }
}

Write-Host 'Removing the deprovision markers and the orchestrator block'
foreach ($pfn in 'Microsoft.OutlookForWindows_8wekyb3d8bbwe', 'microsoft.windowscommunicationsapps_8wekyb3d8bbwe') {
    Remove-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Appx\AppxAllUserStore\Deprovisioned\$pfn" -Recurse -Force -ErrorAction SilentlyContinue
}
$oobe = 'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\Orchestrator\UScheduler_Oobe'
$blocked = ([string](Get-ItemProperty $oobe -Name BlockedOobeUpdaters -ErrorAction SilentlyContinue).BlockedOobeUpdaters).Trim()
if ($blocked -match '"MS_Outlook"') {
    $blocked = $blocked -replace '\s*"MS_Outlook"\s*,?', '' -replace ',\s*\]', ']'
    if ($blocked -match '^\[\s*\]$') { Remove-ItemProperty $oobe -Name BlockedOobeUpdaters }
    else { Set-ItemProperty $oobe -Name BlockedOobeUpdaters -Value $blocked -Type String }
}
Remove-Item 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Orchestrator\UScheduler\OutlookUpdate' -Recurse -Force -ErrorAction SilentlyContinue
Remove-EmptyKeys $oobe 'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\Orchestrator'
Remove-EmptyKeys 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Orchestrator\UScheduler' 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate'

Write-Host 'Removing the classic Outlook policies (all logged-in users)'
$sids = (Get-ChildItem Registry::HKEY_USERS | Where-Object { $_.PSChildName -match '^S-1-(5-21|12-1)(-\d+){4}$' }).PSChildName
foreach ($sid in $sids) {
    $u = "Registry::HKEY_USERS\$sid"
    Remove-ItemProperty "$u\Software\Policies\Microsoft\office\16.0\outlook\preferences" -Name NewOutlookMigrationUserSetting -ErrorAction SilentlyContinue
    foreach ($name in 'DoNewOutlookAutoMigration', 'NewOutlookAutoMigrationRetryIntervals', 'HideNewOutlookToggle') {
        Remove-ItemProperty "$u\Software\Policies\Microsoft\office\16.0\outlook\options\general" -Name $name -ErrorAction SilentlyContinue
    }
    Remove-ItemProperty "$u\Software\Microsoft\Office\16.0\Outlook\Options\General" -Name HideNewOutlookToggle -ErrorAction SilentlyContinue

    $policies = "$u\Software\Policies\Microsoft"
    Remove-EmptyKeys "$policies\office\16.0\outlook\preferences" $policies
    Remove-EmptyKeys "$policies\office\16.0\outlook\options\general" $policies
    Remove-EmptyKeys "$u\Software\Microsoft\Office\16.0\Outlook\Options\General" "$u\Software\Microsoft"
}

# the applocker rules from Block-NewOutlookAppLocker.ps1, if they were ever set.
# only the two rules with our ids go: an existing policy keeps everything else
if (Get-Command Get-AppLockerPolicy -ErrorAction SilentlyContinue) {
    $ourIds = 'f6b9a6c1-1f1a-4d5e-9c2b-5a1e0d7c4b01', 'f6b9a6c1-1f1a-4d5e-9c2b-5a1e0d7c4b02'
    $xml = [xml](Get-AppLockerPolicy -Local -Xml)
    $ours = @($xml.SelectNodes('//FilePublisherRule') | Where-Object { $_.Id -in $ourIds })
    if ($ours) {
        Write-Host 'Removing the AppLocker rules'
        foreach ($rule in $ours) { $rule.ParentNode.RemoveChild($rule) | Out-Null }
        # a collection left empty would still block every packaged app it covers
        foreach ($collection in @($xml.SelectNodes('//RuleCollection'))) {
            if (-not $collection.ChildNodes.Count) { $collection.ParentNode.RemoveChild($collection) | Out-Null }
        }
        $tmp = Join-Path $env:TEMP 'no-new-outlook-applocker-restore.xml'
        $xml.Save($tmp)
        Set-AppLockerPolicy -XmlPolicy $tmp
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    }
}

Write-Host ''
Write-Host 'Done. Install new Outlook from the Microsoft Store if you want it back.'
