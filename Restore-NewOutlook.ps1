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

Write-Host 'Removing the classic Outlook policies (all logged-in users)'
$sids = (Get-ChildItem Registry::HKEY_USERS | Where-Object { $_.PSChildName -match '^S-1-(5-21|12-1)(-\d+){4}$' }).PSChildName
foreach ($sid in $sids) {
    $u = "Registry::HKEY_USERS\$sid"
    Remove-ItemProperty "$u\Software\Policies\Microsoft\office\16.0\outlook\preferences" -Name NewOutlookMigrationUserSetting -ErrorAction SilentlyContinue
    foreach ($name in 'DoNewOutlookAutoMigration', 'NewOutlookAutoMigrationRetryIntervals', 'HideNewOutlookToggle') {
        Remove-ItemProperty "$u\Software\Policies\Microsoft\office\16.0\outlook\options\general" -Name $name -ErrorAction SilentlyContinue
    }
    Remove-ItemProperty "$u\Software\Microsoft\Office\16.0\Outlook\Options\General" -Name HideNewOutlookToggle -ErrorAction SilentlyContinue
}

Write-Host ''
Write-Host 'Done. Install new Outlook from the Microsoft Store if you want it back.'
