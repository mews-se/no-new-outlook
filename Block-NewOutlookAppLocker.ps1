# optional, and separate from Remove-NewOutlook.ps1 on purpose: this writes a
# machine-wide applocker policy, which is a bigger hammer than removing an app
#
# it exists because the registry blocks cannot reach every path. office ships its
# own installer, NewOutlookInstaller.exe, which pulls the msix straight from
# res.cdn.office.net and registers it per user. no store, no windows update, no
# provisioning: the deprovision marker, BlockedOobeUpdaters and workCompleted are
# all bypassed, and on the business skus the migration policies are ignored too.
# an applocker rule blocks the package itself, so it stops both the install and
# the app. see the readme for the caveats, licensing included.

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
if (-not (Get-Command Set-AppLockerPolicy -ErrorAction SilentlyContinue)) {
    Write-Warning 'The AppLocker cmdlets are missing on this edition of Windows.'
    exit 1
}

# fixed ids so this script is idempotent and Restore-NewOutlook.ps1 can find
# exactly these two rules again
$allowId = 'f6b9a6c1-1f1a-4d5e-9c2b-5a1e0d7c4b01'
$denyId = 'f6b9a6c1-1f1a-4d5e-9c2b-5a1e0d7c4b02'

$allowRule = @"
<FilePublisherRule Id="$allowId" Name="Allow all signed packaged apps" Description="Without this every packaged app that is not explicitly allowed would be blocked" UserOrGroupSid="S-1-1-0" Action="Allow">
  <Conditions>
    <FilePublisherCondition PublisherName="*" ProductName="*" BinaryName="*">
      <BinaryVersionRange LowSection="0.0.0.0" HighSection="*" />
    </FilePublisherCondition>
  </Conditions>
</FilePublisherRule>
"@

$denyRule = @"
<FilePublisherRule Id="$denyId" Name="Deny new Outlook" Description="Blocks Microsoft.OutlookForWindows however it got installed" UserOrGroupSid="S-1-1-0" Action="Deny">
  <Conditions>
    <FilePublisherCondition PublisherName="CN=MICROSOFT CORPORATION, O=MICROSOFT CORPORATION, L=REDMOND, S=WASHINGTON, C=US" ProductName="MICROSOFT.OUTLOOKFORWINDOWS" BinaryName="*">
      <BinaryVersionRange LowSection="0.0.0.0" HighSection="*" />
    </FilePublisherCondition>
  </Conditions>
</FilePublisherRule>
"@

# the whole policy is built here rather than handed to -Merge, so what gets
# tested below is exactly what gets applied
$xml = [xml](Get-AppLockerPolicy -Local -Xml)
foreach ($old in @($xml.SelectNodes('//FilePublisherRule') | Where-Object { $_.Id -in $allowId, $denyId })) {
    $old.ParentNode.RemoveChild($old) | Out-Null
}
$appx = $xml.AppLockerPolicy.RuleCollection | Where-Object Type -eq 'Appx'
if (-not $appx) {
    $appx = $xml.CreateElement('RuleCollection')
    $appx.SetAttribute('Type', 'Appx')
    $appx.SetAttribute('EnforcementMode', 'Enabled')
    $xml.AppLockerPolicy.AppendChild($appx) | Out-Null
}

# a collection holding any rule blocks every packaged app that is not explicitly
# allowed, start menu and settings included. an empty collection therefore needs
# the allow-everything rule; one that already has rules keeps its own
if (-not $appx.ChildNodes.Count) {
    Write-Host 'No packaged-app rules here yet, adding the allow-everything rule with it'
    $appx.AppendChild($xml.ImportNode(([xml]$allowRule).DocumentElement, $true)) | Out-Null
} else {
    Write-Host 'Packaged-app rules already exist, leaving them alone and adding the deny'
}
$appx.AppendChild($xml.ImportNode(([xml]$denyRule).DocumentElement, $true)) | Out-Null

$candidate = Join-Path $env:TEMP 'no-new-outlook-applocker.xml'
$xml.Save($candidate)

# dry run first: a packaged app that stops running here is a broken desktop, not
# a blocked mail client, so nothing is applied until the blast radius is known
$blocked = @(Get-AppxPackage | Test-AppLockerPolicy -XmlPolicy $candidate -User Everyone |
    Where-Object PolicyDecision -ne 'Allowed')
$unexpected = @($blocked | Where-Object { $_.FilePath -notmatch 'OutlookForWindows' })
if ($unexpected) {
    Remove-Item $candidate -Force -ErrorAction SilentlyContinue
    Write-Host ''
    Write-Warning "Not applying: the policy would also block $($unexpected.Count) other packaged app(s), starting with $($unexpected[0].FilePath)."
    exit 1
}

$backup = Join-Path $env:ProgramData 'no-new-outlook-applocker-backup.xml'
Get-AppLockerPolicy -Local -Xml | Out-File $backup -Encoding utf8
Set-AppLockerPolicy -XmlPolicy $candidate
Remove-Item $candidate -Force -ErrorAction SilentlyContinue

$live = [xml](Get-AppLockerPolicy -Local -Xml)
if (-not @($live.SelectNodes('//FilePublisherRule') | Where-Object Id -eq $denyId)) {
    Write-Warning 'The policy did not stick. Nothing is blocking new Outlook right now.'
    exit 1
}

Write-Host ''
Write-Host "Done. New Outlook is blocked from installing and from running; the previous policy is saved at $backup."
if (-not $blocked) {
    Write-Host 'The package is not installed at the moment, so the rule had nothing to match yet. It applies as soon as anything tries to install it.'
}
Write-Host 'The AppID service starts on its own from the policy. Undo with Restore-NewOutlook.ps1.'
