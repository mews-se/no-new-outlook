# no-new-outlook

![Platform: Windows 10/11][badge-windows]
![Shell: Windows PowerShell][badge-powershell]
[![License: Unlicense](https://img.shields.io/badge/license-Unlicense-blue.svg)](LICENSE)

New Outlook keeps coming back. You uninstall it, a Windows update reinstalls it. Staying on classic Outlook doesn't help either; one morning it has migrated you anyway. This script removes new Outlook for every user on the machine and flips every switch Microsoft documents for keeping it away, instead of you digging through the same forum threads every few months.

**Use at your own risk.** The script removes apps for every user on the machine and writes to HKLM and to every logged-in user's registry. It does what the readme says on my machines; nothing is guaranteed on yours, and you get to keep both pieces if it breaks.

## Usage

Download or clone the repo, then from an elevated Windows PowerShell prompt:

```
powershell -ExecutionPolicy Bypass -File .\Remove-NewOutlook.ps1
```

Use the built-in Windows PowerShell (powershell.exe), not PowerShell 7 — the Appx cmdlets are unreliable there and the script refuses to run in it.

One run covers the machine and every user profile that is logged in at the time. The per-user Outlook settings can only be written into loaded profiles, so if several people use the machine, log them in first or run the script again later. Sign out and back in (or reboot) before checking classic Outlook: it reads the toggle setting at logon, not at every launch, so the "Try the new Outlook" switch hangs around until the next sign-in.

## What it does

The app itself is removed with `Remove-AppxPackage` for all users and deprovisioned with `Remove-AppxProvisionedPackage`, so new accounts don't get it either. Deprovisioning is the part that makes this stick: Windows records it under the `Deprovisioned` key in HKLM and won't reinstall the app through updates. The script writes that marker even when the package was never provisioned, which also blocks future provisioning.

The old Mail and Calendar apps go too, if they are still around. They stopped working at the end of 2024, and what remains of them is a migration flow that installs new Outlook through the Store pipeline, out of reach of the deprovision marker.

On Windows 10, and Windows 11 23H2 before the March 2024 preview update, new Outlook arrives through a Windows Update orchestrator job instead. The script deletes the `OutlookUpdate` entry and adds `MS_Outlook` to the `BlockedOobeUpdaters` list without touching anything else in it; both are Microsoft's documented blocks for those builds. On top of that it marks the orchestrator's Outlook job as already completed, a trick that isn't in Microsoft's docs but has been proven out by the debloat community.

Classic Outlook gets the policy values that stop the migration machinery, written into every logged-in user profile: `NewOutlookMigrationUserSetting`, `DoNewOutlookAutoMigration` and `NewOutlookAutoMigrationRetryIntervals` all set to 0, and `HideNewOutlookToggle` set to 1 in both the policy and user paths so the "Try the new Outlook" toggle goes away. A leftover `UseNewOutlook` value is cleared in case the toggle was ever used.

## Undo

```
powershell -ExecutionPolicy Bypass -File .\Restore-NewOutlook.ps1
```

This takes every block and policy the remove script set back to Microsoft defaults, removes the keys it created once they are empty, and drops the AppLocker rules if you added them. It works from what those defaults should be, not from a saved snapshot, so a value that some other tool had set before ends up removed rather than put back. Keys that still hold something else are left alone. The per-user settings follow the same loaded-profile rule as above, so run it again for anyone who wasn't logged in. It doesn't reinstall anything. New Outlook is on the Microsoft Store if you actually want it back; Mail and Calendar are gone for good, Microsoft discontinued those.

## What it can't do

A deliberate manual install from the Microsoft Store still works; the script doesn't try to block that. Neither does it stop Office's own installer — see [the next section](#when-that-isnt-enough) for that one. The Start menu may also keep a leftover "Outlook (new)" pin that is just a placeholder — clicking it installs the app from the Store, so unpin it instead. On consumer editions Windows likes to advertise the app in Start's Recommended section as well, which is the same one-click install; the cure there is turning off Start menu recommendations.

Reinstalling Microsoft 365 Apps brings new Outlook along with it these days. If you deploy Office with the Deployment Tool, add `<ExcludeApp ID="OutlookForWindows" />` to the configuration. The plain installer from office.com has no such option, so after an Office reinstall or repair, run the script again. Same thing after a Windows reset or in-place repair install: the protections live in the registry, and a rebuilt Windows starts the cycle over.

If your mailbox lives in a company tenant, Intune policies and admin-controlled migration beat anything set locally. That is between you and your IT department. If you happen to be that department, Exchange Online has the real kill switch: `Get-OwaMailboxPolicy | Set-OwaMailboxPolicy -OneWinNativeOutlookEnabled $false` blocks the mailboxes from new Outlook entirely, so even a manual install dead-ends at sign-in.

Microsoft changes these mechanisms every now and then. If new Outlook reappears after some future update, open an issue.

## When that isn't enough

There is a fourth way in, and it walks past everything above. Classic Outlook ships its own installer, `NewOutlookInstaller.exe`, which pulls the MSIX straight from `res.cdn.office.net` and registers it for the current user. No Store, no Windows Update, no provisioning — so the deprovision marker, `BlockedOobeUpdaters` and the orchestrator job never get a say, and none of them are touched when it happens. On the business SKUs the migration policies are ignored on top of that, so `DoNewOutlookAutoMigration = 0` doesn't stop it either: Microsoft picks the machine for a migration wave, writes `NewOutlookAutoMigrationType = 1` into the user's Outlook key, and at the next sign-in the app is simply back.

If that is happening to you, there is a bigger hammer:

```
powershell -ExecutionPolicy Bypass -File .\Block-NewOutlookAppLocker.ps1
```

It denies the `Microsoft.OutlookForWindows` package family with an AppLocker rule, which stops the install and the app itself whichever path they take. Two things are worth knowing first. AppLocker treats packaged apps as an allow-list, so a rule collection holding any rule blocks every packaged app that isn't explicitly allowed — on a machine with no rules yet the script adds an allow-everything rule alongside the deny, because without it the Start menu and Settings go too. And it dry-runs the finished policy against every installed package before applying it, refusing to touch anything if more than new Outlook would be caught. Microsoft supports AppLocker enforcement on Enterprise and Education; it does work on Pro, but you are outside the support matrix there. `Restore-NewOutlook.ps1` takes the rules back out.

## Tests

`Invoke-Pester tests` runs a small suite: both scripts must parse, the block list handling and the SID filter are exercised against their corner cases, and the empty-key cleanup runs against real keys under `HKCU` (skipped when not on Windows). The registry logic is mirrored in the test file since the scripts themselves only run elevated on Windows.

## License

[The Unlicense](LICENSE) — public domain, do whatever you want with it.

[badge-windows]: https://img.shields.io/badge/platform-Windows%2010%2F11-0078D4.svg?logo=data:image/svg%2bxml;base64,PHN2ZyBmaWxsPSJ3aGl0ZSIgcm9sZT0iaW1nIiB2aWV3Qm94PSIwIDAgMjQgMjQiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyI+PHRpdGxlPldpbmRvd3M8L3RpdGxlPjxwYXRoIGQ9Ik0wLDBIMTEuMzc3VjExLjM3MkgwWk0xMi42MjMsMEgyNFYxMS4zNzJIMTIuNjIzWk0wLDEyLjYyM0gxMS4zNzdWMjRIMFptMTIuNjIzLDBIMjRWMjRIMTIuNjIzIi8+PC9zdmc+
[badge-powershell]: https://img.shields.io/badge/shell-Windows%20PowerShell-5391FE.svg?logo=data:image/svg%2bxml;base64,PHN2ZyBmaWxsPSJ3aGl0ZSIgcm9sZT0iaW1nIiB2aWV3Qm94PSIwIDAgMjQgMjQiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyI+PHRpdGxlPlBvd2VyU2hlbGw8L3RpdGxlPjxwYXRoIGQ9Ik0yMy4xODEgMi45NzRjLjU2OCAwIC45MjMuNDYzLjc5MiAxLjAzNWwtMy42NTkgMTUuOTgyYy0uMTMuNTcyLS42OTcgMS4wMzUtMS4yNjUgMS4wMzVILjgxOWMtLjU2OCAwLS45MjMtLjQ2My0uNzkyLTEuMDM1TDMuNjg2IDQuMDA5Yy4xMy0uNTcyLjY5Ny0xLjAzNSAxLjI2NS0xLjAzNXptLTguMzc1IDkuMzQ2Yy4yNTEtLjM5NC4yMjctLjkwNS0uMDktMS4yNDNMOS4xMjIgNS4xMjVjLS4zOC0uNDA0LTEuMDM3LS40MDctMS40NjYtLjAwMy0uNDI5LjQwMi0uNDY4IDEuMDU2LS4wODggMS40Nmw0LjY2MiA0Ljk2di4xMWwtNy40MiA1LjM3NGMtLjQ1LjMyNy0uNTMzLjk3Ny0uMTg3IDEuNDUzLjM0Ni40NzYuOTkxLjU5NyAxLjQ0LjI3bDguMjI5LTUuOTFjLjI4LS4xOTYuNDM4LS4zNjUuNTE0LS41MnptLTIuNzk2IDQuMzk5YS45MjguOTI4IDAgMDAtLjkzNC45MjNjMCAuNTEuNDE4LjkyMy45MzQuOTIzaDQuNDMzYS45MjguOTI4IDAgMDAuOTM0LS45MjMuOTI4LjkyOCAwIDAwLS45MzQtLjkyM3oiLz48L3N2Zz4=
