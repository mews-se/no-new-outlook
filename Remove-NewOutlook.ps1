# removes new outlook for all users and blocks windows from bringing it back
# see https://github.com/mews-se/no-new-outlook for what and why

$ErrorActionPreference = 'Stop'

if ($PSVersionTable.PSEdition -eq 'Core') {
    Write-Warning 'The Appx cmdlets are unreliable in PowerShell 7. Run this in Windows PowerShell (powershell.exe).'
    exit 1
}
$user = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $user.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning 'Run this from an elevated PowerShell prompt.'
    exit 1
}

function Set-RegDword($path, $name, $value) {
    if (-not (Test-Path $path)) { New-Item $path -Force | Out-Null }
    Set-ItemProperty $path -Name $name -Value $value -Type DWord
}

# mail and calendar go too: dead since end of 2024, and their migration flow
# installs new outlook through the store, where the deprovision marker can't block it
$apps = 'Microsoft.OutlookForWindows', 'microsoft.windowscommunicationsapps'

Write-Host 'Removing new Outlook and the dead Mail and Calendar apps for all users'
foreach ($app in $apps) {
    Get-AppxPackage -AllUsers $app -ErrorAction SilentlyContinue |
        Remove-AppxPackage -AllUsers -ErrorAction Continue
}

Write-Host 'Deprovisioning them so Windows Update leaves them out'
Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -in $apps } |
    ForEach-Object { Remove-AppxProvisionedPackage -Online -PackageName $_.PackageName -ErrorAction Continue | Out-Null }
# the marker windows honors across feature updates, set even if the package was never provisioned
$depro = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Appx\AppxAllUserStore\Deprovisioned\Microsoft.OutlookForWindows_8wekyb3d8bbwe'
if (-not (Test-Path $depro)) { New-Item $depro -Force | Out-Null }

Write-Host 'Blocking the update orchestrator installer'
$oobe = 'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\Orchestrator\UScheduler_Oobe'
Remove-Item "$oobe\OutlookUpdate" -Recurse -Force -ErrorAction SilentlyContinue
if (Test-Path $oobe) { Remove-ItemProperty $oobe -Name OutlookUpdate -ErrorAction SilentlyContinue }
if (-not (Test-Path $oobe)) { New-Item $oobe -Force | Out-Null }
# add to the block list without clobbering entries other tools may have put there
$blocked = [string](Get-ItemProperty $oobe -Name BlockedOobeUpdaters -ErrorAction SilentlyContinue).BlockedOobeUpdaters
if ($blocked -notmatch '^\[.+\]$') { $blocked = '["MS_Outlook"]' }
elseif ($blocked -notmatch '"MS_Outlook"') { $blocked = $blocked -replace '\]$', ',"MS_Outlook"]' }
else { $blocked = $null }
if ($blocked) { Set-ItemProperty $oobe -Name BlockedOobeUpdaters -Value $blocked -Type String }
# not in microsoft's docs but community-proven: mark the oobe install job as already done
Set-RegDword 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Orchestrator\UScheduler\OutlookUpdate' workCompleted 1

# per-user settings go into every loaded hive, not just hkcu: elevating from a
# standard account would otherwise land them in the admin's profile
Write-Host 'Telling classic Outlook to stay classic (all logged-in users)'
$sids = (Get-ChildItem Registry::HKEY_USERS | Where-Object { $_.PSChildName -match '^S-1-(5-21|12-1)(-\d+){4}$' }).PSChildName
foreach ($sid in $sids) {
    $u = "Registry::HKEY_USERS\$sid"
    Set-RegDword "$u\Software\Policies\Microsoft\office\16.0\outlook\preferences" NewOutlookMigrationUserSetting 0
    Set-RegDword "$u\Software\Policies\Microsoft\office\16.0\outlook\options\general" DoNewOutlookAutoMigration 0
    Set-RegDword "$u\Software\Policies\Microsoft\office\16.0\outlook\options\general" NewOutlookAutoMigrationRetryIntervals 0
    Set-RegDword "$u\Software\Policies\Microsoft\office\16.0\outlook\options\general" HideNewOutlookToggle 1
    Set-RegDword "$u\Software\Microsoft\Office\16.0\Outlook\Options\General" HideNewOutlookToggle 1
    Remove-ItemProperty "$u\Software\Microsoft\Office\16.0\Outlook\Preferences" -Name UseNewOutlook -ErrorAction SilentlyContinue
}

Write-Host ''
Write-Host 'Done. Settings were applied to every logged-in user profile; accounts that were not logged in need a later run.'
