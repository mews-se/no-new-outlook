# no-new-outlook

![Platform: Windows 10/11][badge-windows]
![Shell: Windows PowerShell][badge-powershell]
[![License: Unlicense](https://img.shields.io/badge/license-Unlicense-blue.svg)](LICENSE)

New Outlook keeps coming back: you uninstall it, an update or a migration wave reinstalls it. This script removes it for every user on the machine and flips every switch Microsoft documents for keeping it away, so you don't have to dig through the same forum threads every few months.

**Use at your own risk.** It removes apps for every user and writes to HKLM and to every logged-in user's registry. It does what the readme says on my machines; nothing is guaranteed on yours, and you get to keep both pieces if it breaks.

## Usage

```
powershell -ExecutionPolicy Bypass -File .\Remove-NewOutlook.ps1
```

Or straight from GitHub, nothing to download:

```
powershell -Command "irm https://raw.githubusercontent.com/mews-se/no-new-outlook/main/Remove-NewOutlook.ps1 | iex"
```

Run either from an elevated prompt. The file form needs the built-in Windows PowerShell (powershell.exe) — the Appx cmdlets are unreliable in PowerShell 7 and the script refuses to run there — while the one-liner starts powershell.exe itself, so it can be pasted into anything elevated. One run covers the machine and every user profile that is logged in at the time; run it again later for anyone who wasn't. Sign out and back in before checking classic Outlook: the "Try the new Outlook" toggle is read at logon, not at launch.

## What it does

Removes new Outlook for all users, deprovisions it, and writes the `Deprovisioned` marker in HKLM — even when the package was never provisioned — so Windows updates don't bring it back. The dead Mail and Calendar apps go too; what remains of them is a migration flow that installs new Outlook through the Store. On builds where new Outlook arrives through a Windows Update orchestrator job instead (Windows 10, and 11 23H2 before the March 2024 update), the job is deleted, `MS_Outlook` is appended to `BlockedOobeUpdaters` without touching other entries, and the job is marked as already completed — that last one is community-proven rather than documented. Classic Outlook gets the policy values that stop the migration machinery, written into every logged-in profile, and the toggle is hidden.

## Undo

```
powershell -ExecutionPolicy Bypass -File .\Restore-NewOutlook.ps1
```

or without downloading:

```
powershell -Command "irm https://raw.githubusercontent.com/mews-se/no-new-outlook/main/Restore-NewOutlook.ps1 | iex"
```

Puts everything back to Microsoft defaults — from what those defaults should be, not from a snapshot — removes the keys it created once they are empty, and drops the AppLocker rules and their service start type if you added them. Same loaded-profile rule as above. It doesn't reinstall anything; new Outlook is on the Microsoft Store if you actually want it back.

## What it can't do

A deliberate install from the Store still works. The Start menu may keep a placeholder "Outlook (new)" pin that installs the app when clicked — unpin it instead. Reinstalling or repairing Microsoft 365 Apps brings new Outlook along: with the Deployment Tool add `<ExcludeApp ID="OutlookForWindows" />`, otherwise run the script again afterwards; same after a Windows reset. In a company tenant, Intune policies and admin-controlled migration beat anything set locally — and if you are the admin, Exchange Online has the real kill switch: `Get-OwaMailboxPolicy | Set-OwaMailboxPolicy -OneWinNativeOutlookEnabled $false` blocks the mailboxes from new Outlook entirely, so even a manual install dead-ends at sign-in.

Microsoft changes these mechanisms every now and then. If new Outlook reappears after some future update, open an issue.

## When that isn't enough

There is a fourth way in that walks past everything above. Classic Outlook ships its own installer, `NewOutlookInstaller.exe`, which pulls the MSIX straight from Microsoft's CDN and registers it per user — no Store, no Windows Update, no provisioning, and on the business SKUs the migration policies are ignored on top of that. Once Microsoft picks the machine for a migration wave, that installer fires at every sign-in and the app is back within minutes of each reboot.

```
powershell -ExecutionPolicy Bypass -File .\Block-NewOutlookAppLocker.ps1
```

or without downloading:

```
powershell -Command "irm https://raw.githubusercontent.com/mews-se/no-new-outlook/main/Block-NewOutlookAppLocker.ps1 | iex"
```

This denies the `Microsoft.OutlookForWindows` package with an AppLocker rule, which stops the install and the app whichever path they take. On a machine caught in a migration wave the block gets exercised at every sign-in: the installer's attempts land as event 8025 in the AppLocker log and go nowhere.

Three things worth knowing. AppLocker treats packaged apps as an allow-list — a collection holding any rule blocks everything not explicitly allowed, so on a machine with no rules yet the script adds an allow-everything rule alongside the deny. It dry-runs the finished policy against every installed package and refuses to apply if more than new Outlook would be caught. And the rules alone block nothing: the AppID service does the enforcing and its trigger start can't be trusted, so the script sets the service to start at boot, starts it right away, and treats a stopped service as a failure. Microsoft supports AppLocker enforcement on Enterprise and Education; it does work on Pro, but you are outside the support matrix there.

## Tests

`Invoke-Pester tests` checks that all three scripts parse. CI runs it on every push.

## License

[The Unlicense](LICENSE) — public domain, do whatever you want with it.

[badge-windows]: https://img.shields.io/badge/platform-Windows%2010%2F11-0078D4.svg?logo=data:image/svg%2bxml;base64,PHN2ZyBmaWxsPSJ3aGl0ZSIgcm9sZT0iaW1nIiB2aWV3Qm94PSIwIDAgMjQgMjQiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyI+PHRpdGxlPldpbmRvd3M8L3RpdGxlPjxwYXRoIGQ9Ik0wLDBIMTEuMzc3VjExLjM3MkgwWk0xMi42MjMsMEgyNFYxMS4zNzJIMTIuNjIzWk0wLDEyLjYyM0gxMS4zNzdWMjRIMFptMTIuNjIzLDBIMjRWMjRIMTIuNjIzIi8+PC9zdmc+
[badge-powershell]: https://img.shields.io/badge/shell-Windows%20PowerShell-5391FE.svg?logo=data:image/svg%2bxml;base64,PHN2ZyBmaWxsPSJ3aGl0ZSIgcm9sZT0iaW1nIiB2aWV3Qm94PSIwIDAgMjQgMjQiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyI+PHRpdGxlPlBvd2VyU2hlbGw8L3RpdGxlPjxwYXRoIGQ9Ik0yMy4xODEgMi45NzRjLjU2OCAwIC45MjMuNDYzLjc5MiAxLjAzNWwtMy42NTkgMTUuOTgyYy0uMTMuNTcyLS42OTcgMS4wMzUtMS4yNjUgMS4wMzVILjgxOWMtLjU2OCAwLS45MjMtLjQ2My0uNzkyLTEuMDM1TDMuNjg2IDQuMDA5Yy4xMy0uNTcyLjY5Ny0xLjAzNSAxLjI2NS0xLjAzNXptLTguMzc1IDkuMzQ2Yy4yNTEtLjM5NC4yMjctLjkwNS0uMDktMS4yNDNMOS4xMjIgNS4xMjVjLS4zOC0uNDA0LTEuMDM3LS40MDctMS40NjYtLjAwMy0uNDI5LjQwMi0uNDY4IDEuMDU2LS4wODggMS40Nmw0LjY2MiA0Ljk2di4xMWwtNy40MiA1LjM3NGMtLjQ1LjMyNy0uNTMzLjk3Ny0uMTg3IDEuNDUzLjM0Ni40NzYuOTkxLjU5NyAxLjQ0LjI3bDguMjI5LTUuOTFjLjI4LS4xOTYuNDM4LS4zNjUuNTE0LS41MnptLTIuNzk2IDQuMzk5YS45MjguOTI4IDAgMDAtLjkzNC45MjNjMCAuNTEuNDE4LjkyMy45MzQuOTIzaDQuNDMzYS45MjguOTI4IDAgMDAuOTM0LS45MjMuOTI4LjkyOCAwIDAwLS45MzQtLjkyM3oiLz48L3N2Zz4=
