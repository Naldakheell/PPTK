# PPTK for Mac: distribution (ZIP for IT, PKG for Intune)

**Core requirement: future updates are published once in Intune and reach every assigned Mac.** One build produces
one PKG. The IT ZIP and the Intune folder both carry **that same PKG** (`com.pptk.mac.pkg`), with the same paths and
the same version. A Mac installed by hand is an ordinary Intune install: Intune detects its version
(`com.pptk.keys` + `CFBundleShortVersionString`) and upgrades it in place. The exact update routine is in
`docs/INTUNE-MAC.md` §5.

```
src/Distribution/mac/
  build-dist.sh            builds the release into dist/mac/PPTK-Mac-<version>[-UNSIGNED-TEST]/
  scripts/common.sh        every path, label and id (inlined into each script at build time)
  scripts/preinstall       downgrade guard, stop helpers, migrate a per-user development install
  scripts/postinstall      permissions, start the helper in logged-in sessions
  scripts/pptk-uninstall.sh
  scripts/pptk-detect.sh   health check (IT, Intune custom attribute)
  zip/Install PPTK.command, zip/Uninstall PPTK.command
  docs/INSTALL-MAC.md, docs/INSTALL-MAC-AR.md, docs/INTUNE-MAC.md   (copied into the release)
  tests/test-dist.sh       validates a built release in a fake root (no real install)
```

## Release output

```
dist/mac/PPTK-Mac-<v>/
  PPTK-<v>.pkg                      the package (Intune "macOS app (PKG)"; also inside the ZIP)
  PPTK-Mac-<v>-Installer.zip        IT: PKG + Install/Uninstall PPTK.command + scripts + EN/AR guide + SHA256SUMS
  PPTK-PPPC.mobileconfig            privacy profile (Accessibility, Automation; Input Monitoring user-approved)
  PPTK-LoginItems.mobileconfig      managed login items: the helper's background item stays on
  intune/                           PKG, both profiles, pptk-uninstall.sh, pptk-detect.sh, INTUNE-MAC.md
  version.json  SHA256SUMS.txt  BUILD-INFO.txt
```

## Installed layout (same for ZIP and Intune)

| Item | Path | Owner |
|---|---|---|
| Helper | `/Applications/PPTK Keys.app` (bundle id `com.pptk.keys`, version = PPTK version) | root |
| LaunchAgent (all users, Aqua sessions) | `/Library/LaunchAgents/com.pptk.keys.plist` | root |
| **Add-in** | `/Library/Application Support/Microsoft/Office365/User Content.localized/Startup/PowerPoint/PPTK.ppam` | root |
| `version.json`, Email Slides plugin, uninstall/detect | `/Library/Application Support/PPTK/` | root |
| Email Slides plugin, per user (copied by the helper) | `~/Library/Application Scripts/com.microsoft.Powerpoint/InstrumentaAppleScriptPlugin.applescript` | user |
| Log | `~/Library/Logs/PPTK/PPTKKeys.log` (rotated at 5 MB) | user |
| Receipt | `com.pptk.mac.pkg` | |

**Why the Startup folder.** These were tested on 2026-11-22, on a managed Mac with the employees' policies:
- **Office's system-wide Startup folder:** PowerPoint loads every add-in in it for every user, at start, with no
  registration, no file-access dialog and no permission prompt. Replacing the file while PowerPoint had it loaded
  was safe; the new version loaded at the next start.
- **An add-in outside Office's folders:** PowerPoint (sandboxed) shows "Grant File Access", and without it reports
  "couldn't load the add-in".
- **Office's per-user folder:** it loads, but rewriting a file there makes macOS ask each user to let the app
  "access data from other apps", which is bad for central updates.

The Startup folder is root-only, so only the PKG changes it, and there is one add-in per Mac.

## Versioning

- One version everywhere: `PPTK_VERSION` in `src/Modules/ModulePPTKBrand.bas`. The build stops when the exported
  add-in's VBA has another version, or lacks the Mac ribbon.
- The same version goes into: PKG version, receipt, `CFBundleShortVersionString` / `CFBundleVersion` of
  `PPTK Keys.app` (Intune detection), `version.json`, and the default expected version of `pptk-detect.sh`.
- `version.json` lists the SHA-256 of every installed file; `pptk-detect.sh` checks them.

## Update and upgrade logic

| Situation | Behaviour |
|---|---|
| Same version again (manual, then Intune) | Reinstalled in place; still one app, one agent, one add-in. |
| Newer version | preinstall stops the helper in each logged-in session. The payload replaces the files (the app is non-relocatable: it always goes to `/Applications`). postinstall starts the helper. |
| PowerPoint open during an update | PowerPoint is never closed. The installer writes the new add-in beside the old one and renames it into place; an open PowerPoint keeps reading the version it loaded and uses the new one at its next start. The helper updates at once. |
| Older version | preinstall exits 1, nothing is changed (`PPTK_ALLOW_DOWNGRADE=1` overrides, for IT rollback). |
| Development install (`~/Library/LaunchAgents/com.pptk.keys.plist`, `pptkkeys-ctl`) | Agent, app and script are moved to `~/Library/Application Support/PPTK/backups/legacy-install-<date>/`; only the system agent remains. |
| Installing onto another volume | Installed into that volume; no helper of the running system is touched. |

## Signing and notarization (required before employees)

The current local certificate "PPTK Local Code Signing" is **self-signed and only trusted on the build Mac**. It is
**not suitable for distribution**:
- Other Macs do not trust it.
- Gatekeeper and notarization reject it.
- A PPPC profile can't rely on it.

`build-dist.sh` never uses it. Without Developer ID it makes an ad-hoc signed, unsigned-PKG build named
`…-UNSIGNED-TEST`, for testing only.

Apple requirements for distribution:
1. **Apple Developer Program**, organisation account (Team ID).
2. **Developer ID Application** certificate: `PPTK Keys.app` is signed with the hardened runtime and the
   `com.apple.security.automation.apple-events` entitlement (`PPTK_APP_SIGN_ID`).
3. **Developer ID Installer** certificate: signs the PKG (`PPTK_PKG_SIGN_ID`).
4. **Notarization**: `xcrun notarytool store-credentials <profile>` once, then `PPTK_NOTARY_PROFILE=<profile>`. The
   build submits the PKG, waits and staples the ticket.
5. Keep the same Team ID for every release. TCC and the PPPC profile match the code requirement
   `identifier "com.pptk.keys" and anchor apple generic and certificate leaf[subject.OU] = <TEAMID> …`, so
   permissions survive updates.

```bash
PPTK_APP_SIGN_ID="Developer ID Application: <Org> (<TEAMID>)" \
PPTK_PKG_SIGN_ID="Developer ID Installer: <Org> (<TEAMID>)" \
PPTK_NOTARY_PROFILE="pptk-notary" \
src/Distribution/mac/build-dist.sh
```

## macOS permissions: what can be managed

| Permission (TCC) | Needed for | Intune / MDM (PPPC profile) | Without a profile |
|---|---|---|---|
| Accessibility | reading the focused PowerPoint element (key context) | **Can be granted** (Allow) | user turns on PPTK Keys |
| Automation: PPTK Keys > PowerPoint | selection read, running PPTK commands | **Can be granted** (AppleEvents Allow) | macOS asks once; user clicks OK |
| Input Monitoring | the keyboard event tap | **Cannot be granted**. A profile can only let a standard user approve it. | user turns on PPTK Keys |

A profile entry is honoured only for an app whose signature matches its code requirement, so in practice a
Developer ID build. Reports say macOS 27 ignores profile-granted Accessibility; check each new macOS release.
Nothing in PPTK bypasses TCC, Gatekeeper or quarantine. Downloaded files keep their quarantine mark; the guides tell
IT to use Control-click > Open or `bash`.

## Build and test

```bash
src/Distribution/mac/build-dist.sh                       # dist/mac/PPTK-Mac-<v>[-UNSIGNED-TEST]
PPTK_VERSION_OVERRIDE=1.2.1 src/Distribution/mac/build-dist.sh --out /tmp/pptk-test   # test-only newer build
src/Distribution/mac/tests/test-dist.sh dist/mac/PPTK-Mac-1.2.0-UNSIGNED-TEST /tmp/pptk-test/PPTK-Mac-1.2.1-UNSIGNED-TEST-VERSIONTEST
```

Requirements on the build Mac:
- Xcode Command Line Tools. Universal binaries are built per architecture and merged with `lipo`, so Xcode isn't needed.
- python3 with `oletools`.
- `/Applications/Microsoft PowerPoint.app`, which supplies PowerPoint's code requirement for the profile.

The add-in input is the exported `bin/PPTK DEV.ppam`, which carries the Mac ribbon. It is renamed to `PPTK.ppam`
in the package, so PowerPoint lists it as "PPTK", like Windows.
