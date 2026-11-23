# PPTK for Mac __PPTK_VERSION__: manual installation (IT)

This ZIP installs PPTK on one Mac. It contains **the same PKG that Intune deploys**, so a Mac installed by hand is
later updated centrally by Intune: same files, same LaunchAgent, same receipt, never installed twice.

## What gets installed

| What | Where |
|---|---|
| Keyboard-shortcut helper | `/Applications/PPTK Keys.app` (bundle id `com.pptk.keys`, version __PPTK_VERSION__) |
| Starts the helper at every user's login | `/Library/LaunchAgents/com.pptk.keys.plist` |
| PowerPoint add-in, loaded for every user | `/Library/Application Support/Microsoft/Office365/User Content.localized/Startup/PowerPoint/PPTK.ppam` |
| Version manifest, Email Slides plugin, uninstall and check scripts | `/Library/Application Support/PPTK/` |
| Helper log | `~/Library/Logs/PPTK/PPTKKeys.log` |

## Requirements

- macOS 13 or later, Apple silicon or Intel.
- Microsoft 365 PowerPoint for Mac in `/Applications`.
- An administrator account (the installer asks for its password).

## Install

1. Copy `PPTK-Mac-__PPTK_VERSION__-Installer.zip` to the Mac (for example OneDrive, Teams or a USB drive). Put it in
   `/Users/Shared`, so an administrator can reach it if the employee isn't an admin.
2. Double-click the ZIP to extract `PPTK-Mac-__PPTK_VERSION__-Installer`.
3. Double-click **Install PPTK.command**.
   - If macOS says it can't be opened because the developer can't be verified: open System Settings > Privacy &
     Security, find the message about "Install PPTK.command", click **Open Anyway**, confirm, and open the file again.
     That's Apple's supported way. Don't remove the download mark from files.
   - Alternative, in Terminal: `cd "/Users/Shared/PPTK-Mac-__PPTK_VERSION__-Installer" && bash "Install PPTK.command"`.
4. Enter the administrator password. The script checks the package checksum, installs it, and runs the check.
5. The employee approves once, in System Settings > Privacy & Security:
   - **Accessibility**: turn on **PPTK Keys**;
   - **Input Monitoring**: turn on **PPTK Keys**.
6. **Quit PowerPoint (⌘Q) and open it again.** The **PPTK** tab appears. The first time a shortcut is used, macOS
   asks whether PPTK Keys may control PowerPoint: click **OK**.

No logout or restart is needed. If macOS reports "Background item added: PPTK Keys", leave it on.

## Check

```bash
sudo "/Library/Application Support/PPTK/pptk-detect.sh"
```

`PPTK __PPTK_VERSION__ installed` means everything is in place, including the logged-in user's helper. `--machine`
checks only the machine-wide part.

## Updates

Updates come from Intune: see `INTUNE-MAC.md`. Running a newer ZIP's **Install PPTK.command** also works.
- **PowerPoint is never closed.** An open PowerPoint keeps the add-in version it loaded. The new add-in is used at
  the next PowerPoint start; the keyboard helper updates immediately.
- **Downgrades are refused.** Installing an older version over a newer one changes nothing.
- **Permissions stay granted** when the update is signed by the same Developer ID team.

## Uninstall

Double-click **Uninstall PPTK.command** (or `sudo "/Library/Application Support/PPTK/pptk-uninstall.sh"`).
- It removes the helper, the LaunchAgent, the PPTK add-in (no other add-in), the package files and each user's PPTK
  logs.
- If PowerPoint still lists PPTK under Tools > PowerPoint Add-ins after it restarts, select it and click the minus
  button.
- The shared Email Slides plugin, also used by Instrumenta, is kept.

## Development installs

A Mac that ran the development helper (`~/Library/LaunchAgents/com.pptk.keys.plist`) is migrated automatically:
the old agent and app are moved to `~/Library/Application Support/PPTK/backups/legacy-install-<date>/`, so only one
helper runs. A development add-in (for example "PPTK DEV") is not touched; unload it in PowerPoint so that the PPTK
tab doesn't appear twice.
