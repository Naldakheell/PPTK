# PPTK for Mac 1.2.0: Microsoft Intune deployment and updates

**Core rule: one PKG for every channel.** The IT ZIP and Intune install the same PKG (identifier
`com.pptk.mac.pkg`), with the same paths and the same version. A Mac installed by hand is therefore an ordinary
Intune install: Intune detects its version and upgrades it in place. Every later update is published once in
Intune and reaches all assigned Macs, with no manual reinstall.

## 1. Files in this folder

| File | Use |
|---|---|
| `PPTK-1.2.0.pkg` | The app to upload ("macOS app (PKG)"). Byte-identical to the PKG in the IT ZIP. |
| `PPTK-PPPC.mobileconfig` | Privacy profile: pre-approves Accessibility and Automation (PowerPoint) for PPTK Keys. |
| `PPTK-LoginItems.mobileconfig` | Managed login items: keeps the PPTK Keys background item on (the user cannot turn it off). |
| `pptk-uninstall.sh` | Uninstall, as an Intune shell script (self-contained). |
| `pptk-detect.sh` | Health check, for IT or as an Intune custom attribute (self-contained). |

## 2. What the PKG installs (same for ZIP and Intune)

| Item | Path |
|---|---|
| Keyboard helper (Intune detection) | `/Applications/PPTK Keys.app`: bundle id `com.pptk.keys`, version = PPTK version |
| Helper start at every login | `/Library/LaunchAgents/com.pptk.keys.plist` |
| PowerPoint add-in | `/Library/Application Support/Microsoft/Office365/User Content.localized/Startup/PowerPoint/PPTK.ppam` |
| Version manifest, Email Slides plugin, uninstall/detect | `/Library/Application Support/PPTK/` |
| Receipt | `com.pptk.mac.pkg` |

The add-in sits in Office's system-wide Startup folder. PowerPoint loads every add-in there for every user when it
starts: no registration, no file-access dialog, no per-user copy. This was tested on a managed Mac with the same
policies. Only root can write there, so only the PKG changes it: one add-in, never a duplicate.

## 3. Prerequisites (once)

1. **Developer ID signing and notarization.** Upload only a build whose `BUILD-INFO.txt` shows Developer ID and
   `notarized true`. Keep the **same Team ID for every release**: macOS keeps the employees' permissions across
   updates only when each version is signed by the same team. An `UNSIGNED-TEST` build is for testing only.
2. **The Microsoft Intune management agent for macOS.** Intune installs it on managed Macs when the first PKG app or
   shell script is assigned.
3. **Profiles, assigned to the same group as the app, before or with it:**
   - `PPTK-PPPC.mobileconfig` (Devices > macOS > Configuration > Create > Templates > **Custom**);
   - `PPTK-LoginItems.mobileconfig` (same way).

## 4. First release: create the Intune app

Intune admin center > Apps > macOS > Add > **macOS app (PKG)**.

| Setting | Value |
|---|---|
| App package file | `PPTK-1.2.0.pkg` |
| Name | PPTK for Mac (keep the same name and app for all versions) |
| Minimum operating system | macOS Ventura 13.0 |
| Pre-install / post-install script | none (the PKG has its own) |
| Detection: included apps | `com.pptk.keys`, version `1.2.0` (filled in from the PKG) |
| **Ignore app version** | **No.** Detection must compare the version, or Intune would never install updates. |
| **Assignment** | **Required**, to the device group (or user group) of PPTK users. Not "Available": see below. |

**Why Required:**
- With a Required assignment, Intune checks the detection rule on every assigned Mac and installs the PKG wherever the
  installed version doesn't match. That includes Macs installed by hand with the ZIP.
- An "Available" assignment updates only copies the user installed from Company Portal, and only with agent
  2606.013 or later. A manual install would never be updated.

**Macs installed by hand with the ZIP:**
- If they have the same version, Intune detects them as installed and does nothing.
- If they're older, Intune installs the newer PKG over them.

## 5. Publishing every future update (the routine)

1. **Build** the new version with Developer ID and notarization (`src/Distribution/mac/README.md`).
   - Raise `PPTK_VERSION` in `ModulePPTKBrand.bas`, for example 1.2.0 → 1.2.1.
   - Export the add-in, then run `build-dist.sh`.
   - The build refuses a mismatch between the add-in's version and the source.
2. **Check the release:**
   - `BUILD-INFO.txt` shows Developer ID, `notarized true` and the new version.
   - Run `tests/test-dist.sh` with the previous and the new release.
3. **Pilot:** assign the update to a small pilot group first (step 4 below), or test the ZIP on one Mac.
4. **Update the existing Intune app.** Don't create a second app: Intune has no supersedence for macOS apps, so
   updates are made by replacing the package in the same app.
   1. Apps > macOS > **PPTK for Mac** > Properties > App information > Edit.
   2. Select the new file, `PPTK-<new version>.pkg`.
   3. Check that the detection rule now shows `com.pptk.keys` with the **new** version, and that **Ignore app
      version = No**.
   4. Save. Keep the assignment **Required**.
5. **Rollout:**
   - At its next check-in, each Mac's Intune agent compares the installed helper version with the new one and
     installs the new PKG.
   - The agent checks in about every 8 hours. Users can trigger it sooner in Company Portal > Devices > Check status.
6. **Watch:** Intune app status ("Installed" or "Failed"), plus optionally the custom attribute from `pptk-detect.sh`.

**What happens on each Mac during an update:**
- **Downgrades are refused.** The installer refuses an older PKG than the installed one (exit code 1) and changes
  nothing.
- **The helper updates immediately.** It is stopped in every logged-in session, replaced, and started again.
- **PowerPoint is never closed.** If PowerPoint is open, it keeps working with the add-in version it loaded. The new
  add-in replaces the file atomically (tested while PowerPoint had the add-in loaded), and PowerPoint uses it at its
  next start.
- **Permissions stay granted.** Accessibility, Input Monitoring and Automation remain when the new version is signed
  by the same Developer ID team.
- **No duplicates:** one app, one LaunchAgent, one add-in file.

**Rollback:** publish the previous PKG only with care, because downgrades are refused by design. For a deliberate
rollback, build the old code with a **higher** version number.

## 6. Privacy and background items

| Item | Intune | Notes |
|---|---|---|
| Accessibility | `PPTK-PPPC.mobileconfig` (Allow) | Needs the Developer ID build (the code requirement in the profile). |
| Automation: PPTK Keys > PowerPoint | `PPTK-PPPC.mobileconfig` (Allow) | Same. |
| Input Monitoring | Cannot be granted (Apple) | The user turns on PPTK Keys once in System Settings > Privacy & Security > Input Monitoring. The profile lets standard users approve it. |
| Background item "PPTK Keys" | `PPTK-LoginItems.mobileconfig` | Keeps it on, so the helper starts at every login. |

**To check in the pilot:**
- Apple's documentation lists the privacy profile as needing a supervised Mac. These Macs are enrolled with
  User-Approved MDM. If the profile doesn't apply, employees approve Accessibility and Automation once themselves.
- Reports say macOS 27 ignores Accessibility granted by profile.

## 7. Uninstall

Intune doesn't uninstall "macOS app (PKG)" apps.
1. Remove the app's assignment.
2. Run `pptk-uninstall.sh` once as an Intune shell script:
   - run as signed-in user: **No**;
   - hide notifications: Yes;
   - frequency: once;
   - retries: 1.

It removes the helper, the LaunchAgent, the add-in from the Startup folder (no other add-in), the package files,
each user's PPTK logs, and the receipt. An open PowerPoint keeps the add-in until it quits.

## 8. Validation status

- **Tested on the build Mac (automated, isolated, the real PKGs):**
  - a 1.2.0 manual install, then 1.2.1 through the simulated Intune detection;
  - no duplicates;
  - an add-in kept open during the upgrade;
  - downgrade refusal, migration of a development install, uninstall.
- **Tested in real PowerPoint on this managed Mac:**
  - loading from the Startup folder;
  - replacing the add-in while PowerPoint had it loaded;
  - the new version loading at the next start.
- **Not yet tested:**
  - Intune itself (upload, detection, update over a manual install);
  - a Developer ID signed, notarized build;
  - the profiles on a managed Mac;
  - an Intel Mac.
