# PPTK for Windows: Intune deployment (universal x64 + ARM64)

One Microsoft Intune **Win32 app** for every Windows device, x64 and ARM64. PPTK installs per user:

- the PowerPoint add-in;
- the PPTKKeys keyboard-shortcut helper, which a per-user scheduled task starts at sign-in.

The package carries its own official Microsoft .NET runtime for both architectures. The installer uses the one that matches the device. **No .NET installation and no administrator rights are needed.**

## 1. Release files

```
PPTK-<VERSION>-universal\
  PPTK-Package\                         <- the Intune content folder
    version.json                        version, SHA-256 of every file, runtime fingerprints
    install.ps1  uninstall.ps1  detect.ps1  deploy-env.ps1
    payload\PPTK.ppam
    payload\PPTKKeys\PPTKKeys.dll         one AnyCPU (MSIL) helper for both architectures
    payload\PPTKKeys\PPTKKeys.deps.json
    payload\PPTKKeys\PPTKKeys.runtimeconfig.json
    payload\runtime\win-x64\...         official Microsoft .NET runtime (dotnet-runtime-<v>-win-x64.zip)
    payload\runtime\win-arm64\...       official Microsoft .NET runtime (dotnet-runtime-<v>-win-arm64.zip)
  PPTK-detect.ps1                       Intune custom detection script (manifest embedded)
  intunewin\install.intunewin           the file to upload to Intune
  SHA256SUMS.txt
  INTUNE.md
```

### How the helper runs

The task starts this command:

```
%LOCALAPPDATA%\PPTK\versions\<v>\runtime\dotnet.exe exec "%LOCALAPPDATA%\PPTK\versions\<v>\PPTKKeys\PPTKKeys.dll"
```

- `dotnet.exe` is Microsoft's own signed host from the official runtime archive.
- **PPTK ships no executable of its own.** Builds with an unsigned `PPTKKeys.exe` were blocked by the ASR "low prevalence" rule on the development device. This design avoids that.

### Package safety checks

The installer accepts executables **only** inside `payload\runtime\win-x64\` and `payload\runtime\win-arm64\`. Before using a runtime, it checks that:

- every `.exe` and `.dll` has a **valid Microsoft Corporation signature**;
- `dotnet.exe` is built for the device architecture (machine type x64 `0x8664`, ARM64 `0xAA64`);
- the runtime contains the .NET framework version PPTKKeys needs;
- the runtime's file fingerprint matches `version.json`.

## 2. Prerequisites

| Item | Value |
|---|---|
| OS | Windows 10 / 11, **x64 or ARM64** (x86 is refused) |
| Office | Microsoft 365 desktop PowerPoint (`HKCU\Software\Microsoft\Office\16.0`) |
| .NET | **None.** It is included (private, per user, per version). |
| User | A signed-in user. Install and uninstall run in that user's context. |
| PowerShell | Windows PowerShell 5.1 in Full Language mode. |
| App control | `dotnet.exe` (Microsoft-signed) and its DLLs must be allowed to run from `%LOCALAPPDATA%\PPTK\versions\`. See limitations. |

## 3. Intune Win32 app settings

| Setting | Value |
|---|---|
| App package file | `intunewin\install.intunewin` |
| Install command | `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1` |
| Uninstall command | `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\uninstall.ps1` |
| Install behavior | **User** |
| Device restart behavior | No specific action |
| Return codes | Intune defaults. Keep **1618 = Retry**. |
| Requirements | Operating system architecture: **x64 and ARM64**; Windows 10 1809+ or Windows 11 |
| Dependencies | **None** (the .NET dependency of 1.0.0 is no longer needed) |
| Detection rules | Custom script: **`PPTK-detect.ps1`** |
| Run script as 32-bit process on 64-bit clients | No |
| Enforce script signature check | No (the scripts are not signed) |

### Exit codes

| Code | Meaning |
|---|---|
| 0 | Installed (or removed) and confirmed, or already installed (no change). |
| 1618 | Retry later. The installed version is unchanged. Install: PowerPoint has the add-in open, so the new version is staged and applied when PowerPoint closes. Uninstall: PowerPoint is running or the helper did not close. |
| 1 | Failed. The previous installation is restored. Also returned for an unsupported architecture, an invalid or unsigned runtime, a downgrade, or SYSTEM context. |

## 4. Architecture detection

1. The installer reads the device's own architecture from `Win32_Processor.Architecture` (9 = x64, 12 = ARM64).
   - It does not use the PowerShell process architecture. On ARM64, Intune may run an x64 or x86 PowerShell under emulation, and that would report the wrong architecture.
   - If WMI returns nothing, it falls back to `PROCESSOR_ARCHITECTURE` in `HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment`.
2. Only that architecture's runtime is copied to `versions\<v>\runtime\`. The architecture is recorded in `%LOCALAPPDATA%\PPTK\state.json` (`architecture`).
3. If Intune starts install or uninstall in a 32-bit PowerShell, the script runs itself again in the native 64-bit PowerShell.
4. Detection checks that:
   - the recorded architecture is the device's;
   - the private runtime matches its fingerprint;
   - the task runs that runtime's `dotnet.exe`.

## 5. What detection checks

The app is detected only when all of these are true:

- `HKCU\Software\PPTK\InstalledVersion` equals the version.
- `%APPDATA%\Microsoft\AddIns\PPTK.ppam` matches its hash.
- The add-in registration exists with `AutoLoad = 1`.
- The helper files match their hashes.
- The private runtime matches its fingerprint for the recorded architecture.
- Exactly one task `\PPTKKeys Startup` exists. It runs as the current user, interactive and least-privilege, has a logon trigger, and runs `versions\<v>\runtime\dotnet.exe exec ...\PPTKKeys.dll`.
- `state.json` says this version is installed, and no update for it is still staged.

## 6. Updates from 1.0.0 and later versions

- **Deploying 1.1.0 over 1.0.0:**
  - Use a new Win32 app that supersedes the 1.0.0 app, with "Uninstall previous version" = **No**.
  - Remove the .NET dependency from the 1.0.0 app's successor; it is no longer needed.
  - The task switches from the installed .NET to the private runtime. The helper is asked to close and is started again.
- **If PowerPoint is open:** nothing is closed.
  - The new version, with its runtime, is staged, and install returns 1618.
  - The employee sees one notice that never takes focus.
  - When PowerPoint closes, the installed helper applies the update. It switches the task to the new version's runtime and hands the keyboard hook to the new helper, which starts on its own runtime.
  - For a 1.0.0 helper the handover itself starts on the installed .NET. From the next sign-in the helper runs on the private runtime.
- **Rollback:** `install.ps1 -Rollback` returns to the previous version together with that version's runtime. A rollback to 1.0.0 returns to the installed .NET.
- **Disk use:** each kept version holds its own runtime, about 75 MB. The installed and previous versions are kept, so expect about 150 MB per user.

## 7. Validation status

### Tested locally (automated, isolated)

All tests used isolated test roots. Real Task Scheduler and registry tests used only `PPTK-Test-<id>` names. The production installation was checked before and after and did not change.

**PPTKKeys unit and integration tests:**
- the update engine switches the task to the new private runtime;
- damaged or wrong-architecture runtimes are rejected;
- the task host is restored on failure;
- the handoff starts the new helper on its own runtime.

**Universal package suite, using the real Microsoft runtimes:**
- x64 and ARM64 runtime selection (architecture simulated per test root);
- unsupported architecture refused;
- a tampered runtime is not detected and is repaired by reinstalling;
- an unsigned file in the runtime is refused;
- x64 binaries offered as the ARM64 runtime are refused;
- an executable outside the runtimes is refused;
- upgrade 1.0.0 to 1.1.0, then rollback;
- staging while PowerPoint is open;
- 1.1.0 to 1.1.1, then rollback;
- uninstall;
- failed-upgrade restore.

**Earlier suites still pass:** install, uninstall, detect, the live Task Scheduler/HKCU adapter, and the package build.

**The built package on this x64 device:**
- installed from the package in live-test mode;
- confirmed by `PPTK-detect.ps1`;
- the task's private x64 `dotnet.exe` runs the helper DLL (`--version`);
- uninstalled.

### Requires validation on a real device / in an Intune pilot

**ARM64 (no ARM64 device was available):**
- the ARM64 runtime has only been checked as files (Microsoft signature, ARM64 machine type, hash). It has **never been executed**;
- the helper on ARM64: keyboard hook, shortcuts, PowerPoint COM, focus detection (window classes such as `PPTFrameClass`), the update notice and the handoff;
- PowerPoint for ARM64 (native ARM64 Office) and the PPTK add-in VBA on it;
- architecture detection through Intune on ARM64. Detection of a real ARM64 processor (`Win32_Processor.Architecture` = 12) was simulated only.

**All devices:**
- Intune user-context install, upgrade from the 1.0.0 app (supersedence), uninstall and detection;
- the 1618 retry behaviour;
- ASR, Defender, WDAC and AppLocker on the production policies, running `dotnet.exe` from `%LOCALAPPDATA%`;
- the time a first install takes. It hashes and signature-checks about 75 MB, roughly 30 s on the development device.

## 8. Known limitations

- **App control:** devices whose WDAC or AppLocker policy blocks executables in user-writable folders will block the private `dotnet.exe`. IT can allow Microsoft-signed `dotnet.exe` (publisher rule). PPTK contains no bypass.
- **No forced recovery during handoff** (`ForcedRecoveryEnabled = false`, not validated).
- **No automatic runtime patching:** the bundled runtime is fixed per PPTK version. .NET security updates reach PPTK with a new PPTK build (rebuild with the newer runtime).
- **Unsigned scripts; Constrained Language mode is not supported.**
- **Per-user scope:** each user is installed and detected separately. Each user also gets their own copy of the runtime.
