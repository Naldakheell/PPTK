# 07 — Release and deployment (Windows) and what it means for Mac

Windows deployment code lives in `src/Deploy/` (committed in `49fe2f2 installation`). Release outputs are in `dist/`, which is git-ignored.

**None of the scripts below run on macOS.** They use PowerShell 5.1, Task Scheduler, HKCU, Authenticode and Win32. Reuse the *concepts* in §3, not the code.

## 1. Windows release structure

### 1.1 Package (built by `src/Deploy/build-package.ps1`)

```
PPTK-Package\
  version.json          schemaVersion 1, packageVersion, SHA-256 + size of every file,
                        build info (commit, dirty flag), requirements, "runtimes" fingerprints
  install.ps1  uninstall.ps1  detect.ps1  deploy-env.ps1
  payload\PPTK.ppam
  payload\PPTKKeys\PPTKKeys.dll, .deps.json, .runtimeconfig.json     (framework-dependent, AnyCPU IL)
  payload\runtime\win-x64\...    official Microsoft .NET 10.0.15 runtime   (1.1.0 universal only)
  payload\runtime\win-arm64\...  official Microsoft .NET 10.0.15 runtime   (1.1.0 universal only)
PPTK-detect.ps1         detect.ps1 with a slim manifest embedded (Intune detection rule)
```

What the build checks:

- **The add-in:**
  - It is built by PowerPoint itself from `src/PPTK-WINDOWS-FINAL.pptm` (save as `.ppam`).
  - Before PowerPoint starts, the VBA inside the PPTM is compared with the exported `.bas`/`.cls`/`.frm` sources. A stale PPTM stops the build.
- **Version consistency:**
  - `PPTK_VERSION` in `src/PPTK/Brand/ModulePPTKBrand.bas` must equal `<Version>` in `PPTKKeys.csproj`.
  - The built add-in must contain that version, and `PPTK_GetVersion()` must return it.
  - The DLL's `--version` must match.
- **Reproducibility:**
  - PPTKKeys files are byte-identical for the same source.
  - The PPAM bytes change on every PowerPoint save, so `version.json` records a `contentFingerprint` that is stable for the same source.
- **No executables except the runtimes:**
  - No apphost `PPTKKeys.exe`. The helper runs as `dotnet.exe exec PPTKKeys.dll`.
  - In the universal package, every `.exe`/`.dll` under `payload/runtime/<rid>/` must be validly signed by Microsoft, and `dotnet.exe` must be built for its architecture (PE machine `0x8664`/`0xAA64`).

### 1.2 Releases in `dist/` (git-ignored)

| Folder | Content |
|---|---|
| `dist/PPTK-1.0.0/` | PPTK 1.0.0 (needs .NET 10 installed), `intunewin/install.intunewin`, `INTUNE.md`, `SHA256SUMS.txt`. Kept as a backup. |
| `dist/PPTK-1.1.0-universal/` | PPTK 1.1.0 with private .NET runtimes for x64 and ARM64, its `.intunewin`, `INTUNE.md`, `SHA256SUMS.txt` |
| `dist/PPTK-Direct-Installer/` | employee ZIP `PPTK-1.1.0-Installer.zip` (the 1.1.0 package plus `Install PPTK.cmd` / `Uninstall PPTK.cmd`, verification, Arabic and English guides); built from the `direct-installer` worktree branch |

### 1.3 Installed layout (per user, no admin rights)

```
%APPDATA%\Microsoft\AddIns\PPTK.ppam
HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\PPTK   Path=PPTK.ppam, AutoLoad=1
HKCU\Software\PPTK                                            InstalledVersion
%LOCALAPPDATA%\PPTK\versions\<v>\{PPTK.ppam, version.json, PPTKKeys\..., runtime\...}   installed + previous
%LOCALAPPDATA%\PPTK\state.json        installed, previous, pending{...}, architecture, helperActivation
%LOCALAPPDATA%\PPTK\state.lock        exclusive lock while any operation runs
%LOCALAPPDATA%\PPTK\transaction.json  undo journal (see src/Deploy/JOURNAL.md)
%LOCALAPPDATA%\PPTK\pending\          staged add-in while PowerPoint is open
%LOCALAPPDATA%\PPTK\logs\             install-*, uninstall-*, helper-<date>.log
Task \PPTKKeys Startup                current user, interactive token, least privilege, logon trigger:
                                      "<versions\<v>\runtime\dotnet.exe>" exec "<versions\<v>\PPTKKeys\PPTKKeys.dll>"
```

## 2. Mechanisms and why they are built this way

| Mechanism | Where | Why |
|---|---|---|
| **Transaction journal + lock** | `install.ps1`, `uninstall.ps1`, `PendingAddinUpdate.cs`, contract in `src/Deploy/JOURNAL.md` | Every change is recorded first. A failed or killed run is undone by the same run or the next one. One shared format, so the scripts and the helper engine can undo each other's interrupted work. |
| **Never touch an open add-in** | `Test-AddinReplaceable`, engine `CanOpenExclusively` | PowerPoint holds the PPAM open. Replacing it under PowerPoint is unsafe, so the update is **staged** in `pending\` and the installer exits 1618 ("retry"). |
| **Pending update applied by the helper** | `PendingUpdateWatcher.cs` + `PendingAddinUpdate.cs` | When every PowerPoint window has closed, the running helper applies the staged version itself. A hidden PowerPoint process is waited for, never killed. |
| **Update notice** | `UpdateNotice.cs`, `NotificationWindow.cs`, `PowerPointWindows.cs` | One notice per version: "A PPTK update is ready. Please close and reopen PowerPoint…". It never takes focus (`WS_EX_NOACTIVATE`) and stays up 10 s. It waits while PowerPoint is not in front, a slideshow is visible (checked live), a dialog or menu is open, or the user typed in the last 5 s. |
| **Helper handoff** | `HelperHandoff.cs`, `HelperWiring.cs` | The new helper takes over the keyboard hook from the old one without a gap. Rule: **hook ownership = mutex ownership**, so two hooks never run at once. Forced termination of a stuck new helper is **disabled** (`DeploymentFlags.ForcedRecoveryEnabled = false`) because it was not validated on a real device. |
| **Graceful helper stop** | `deploy-env.ps1` `Stop-Helper` | The installer posts `WM_CLOSE` to the helper's message window. It never kills the process; if the helper doesn't close, the new helper starts at the next sign-in. |
| **Private runtime per architecture** | `install.ps1` `Get-NativeRid`, `Assert-Runtime` | Native architecture comes from `Win32_Processor` (an x64 PowerShell on ARM64 is emulated and would lie). Only that runtime is staged. Its signatures, PE machine and fingerprint are verified. |
| **Safe test modes** | `deploy-env.ps1` | `-Root` simulates everything with JSON files. `-Root -LiveTestId` uses the real Task Scheduler and HKCU under `PPTK-Test-<id>` names only. Production names are refused in test mode. |

## 3. Concepts reusable for a future Mac distribution (do not start this work yet)

| Windows concept | Mac equivalent to investigate | Status |
|---|---|---|
| Per-user install, no admin | `~/Library/Group Containers/UBF8T346G9.Office/User Content/Add-Ins/` for the `.ppam` (Office for Mac add-in location), `~/Library/Application Support/PPTK/` for the helper | **Unverified.** The upstream Instrumenta pkg (`src/Installers/pkg/InstrumentaInstaller/`) shows the paths Instrumenta used. |
| Scheduled task at logon | LaunchAgent `~/Library/LaunchAgents/<id>.plist` | Not built. The Mac helper is started manually today. |
| `version.json` with hashes + `detect.ps1` | the same manifest idea; a shell or Swift verifier | Concept reusable as-is. |
| Transaction journal + lock | the same idea in a shell or Swift installer | The concept is reusable; `JOURNAL.md` is the format reference. |
| Staged update while PowerPoint is open, applied by the helper | the same flow. PowerPoint for Mac also locks the add-in while running (**unverified**). | Concept reusable. |
| Non-activating notice | `NSPanel` with `.nonactivatingPanel` | Concept reusable. |
| Private runtime | Not needed if the helper stays a native Swift binary. | Mac advantage: no runtime to ship. |
| Authenticode check of shipped binaries | codesign + notarization of the helper, required for Gatekeeper and TCC stability | **Mac blocker to plan.** An unsigned helper loses Accessibility and Input Monitoring grants when its binary changes. |
| Intune Win32 app | Intune for macOS (`.pkg`), Jamf, or a manual `.pkg` | Out of scope now. |
| ASR / Defender lessons | Gatekeeper, TCC prompts (Accessibility, Input Monitoring, Automation of PowerPoint) | Different controls, same lesson: test on a managed device early. |

## 4. Lessons from Windows packaging and employee testing

1. **Corporate security controls decide the architecture.**
   - Defender ASR rule `01443614` (low prevalence) blocked an unsigned `PPTKKeys.exe` nine times on the dev machine.
   - `dotnet.exe` (Microsoft-signed) running `PPTKKeys.dll` was never blocked.
   - Rule `92E97FA1` blocks Win32 `Declare` in Office macros, so the VBA has no Declares.

   On Mac, plan for TCC, Gatekeeper and MDM profiles the same way, before building features that depend on them.
2. **Never trust the process architecture.** On ARM64, PowerShell can be emulated x64 or x86 and report the wrong architecture. On Mac, Rosetta has the same effect: use the native machine query, not the process's own.
3. **Files downloaded from Teams/SharePoint carry Mark-of-the-Web.** Office blocks macros in such add-ins. The direct installer refuses and explains; it never strips the mark.

   On Mac the equivalent is the `com.apple.quarantine` attribute. Plan for it; never remove it silently.
4. **PowerShell `exit` inside a dot-sourced file only leaves that file.** Use `[Environment]::Exit`. (Example of a script-host trap; on Mac, check `set -e` and subshell exit codes in the same spirit.)
5. **Short 8.3 paths** (`C:\Users\NOUFAL~1`) broke relative path computation. Always normalise paths before comparing.
6. **Always test the real artifact.** The package build, the `.intunewin` (decrypted and compared file by file) and the ZIP (extracted and installed) were each verified after they were built, not only their inputs.
7. **Real-device checks found what simulation could not:**
   - The update notice kept waiting for 30 s after a slideshow ended, because the cached COM state was stale. It was fixed with a live window check.
   - The handoff with a real keyboard hook was validated with test names.

## 5. Validation status (Windows)

| Area | Status |
|---|---|
| Installer, uninstaller, detection, rollback, journal, pending staging | Automated, isolated: sim suites; live Task Scheduler/HKCU suite with test names; the built packages installed and uninstalled with test names. |
| Pending-update engine, update notice, handoff | Unit tests plus real processes and windows, test names. A real-device check of the notice and of the handoff with a real hook was done with test names. |
| Intune end-to-end, 1618 retry behavior, ASR on pilot devices | **Not done.** Requires an Intune pilot. |
| ARM64 | Runtime files verified (signature, PE machine). **No real ARM64 execution.** |
| Forced recovery of a stuck new helper | Disabled; not validated. |
