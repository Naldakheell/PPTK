#!/bin/bash
# Validates built PPTK for Mac releases without touching this Mac's real installation.
#   src/Distribution/mac/tests/test-dist.sh <release 1.2.0 folder> <newer release folder>
# Realistic update scenario: version A installed manually (the ZIP's PKG), then version B published through Intune
# (the intune/ folder's PKG), with Intune's detection rule (bundle id + CFBundleShortVersionString) simulated.
# The PKGs are expanded and their real preinstall/postinstall run against a fake root (PPTK_TEST_ROOT) with a stub
# launchctl and one simulated logged-in user. Payload files are placed the way the installer does it (new file, then
# rename), so a process that has the old add-in open (like PowerPoint) keeps reading the old version.
set -u
REL="$(cd "$1" && pwd)"
NEWER="$(cd "$2" && pwd)"
WORK="$(mktemp -d)"
trap 'kill $HOLDER 2>/dev/null; rm -rf "$WORK"' EXIT
HOLDER=""
pass=0; fail=0
ok() { echo "  PASS $1"; pass=$((pass + 1)); }
bad() { echo "  FAIL $1"; fail=$((fail + 1)); }
check() { if eval "$2"; then ok "$1"; else bad "$1"; fi; }

version_of() { /usr/bin/sed -n 's/.*"version": "\([^"]*\)".*/\1/p' "$1/version.json" | head -1; }
VA=$(version_of "$REL"); VB=$(version_of "$NEWER")
STARTUP_REL="Library/Application Support/Microsoft/Office365/User Content.localized/Startup/PowerPoint"
echo "== releases: $VA (manual ZIP) -> $VB (Intune)"

# ---------------------------------------------------------------- 1. package structure (release A)
pkgutil --expand-full "$REL/PPTK-$VA.pkg" "$WORK/a" >/dev/null
COMP="$WORK/a/component.pkg"
echo "-- package structure"
check "component identifier/version" "grep -q 'identifier=\"com.pptk.mac.pkg\"' '$COMP/PackageInfo' && grep -q 'version=\"$VA\"' '$COMP/PackageInfo'"
check "helper app is not relocatable" "grep -q 'relocatable=\"false\"' '$COMP/PackageInfo' && ! grep -q '<relocate>' '$COMP/PackageInfo'"
check "payload has exactly the expected files" "[ \"\$(cd '$COMP/Payload' && find . -type f | sort | tr '\n' '|')\" = \"./Applications/PPTK Keys.app/Contents/Info.plist|./Applications/PPTK Keys.app/Contents/MacOS/PPTKKeys|./Applications/PPTK Keys.app/Contents/_CodeSignature/CodeResources|./Library/Application Support/Microsoft/Office365/User Content.localized/Startup/PowerPoint/PPTK.ppam|./Library/Application Support/PPTK/InstrumentaAppleScriptPlugin.applescript|./Library/Application Support/PPTK/pptk-detect.sh|./Library/Application Support/PPTK/pptk-uninstall.sh|./Library/Application Support/PPTK/version.json|./Library/LaunchAgents/com.pptk.keys.plist|\" ]"
check "add-in is in Office's Startup folder and carries the Mac ribbon" "unzip -p '$COMP/Payload/$STARTUP_REL/PPTK.ppam' customUI/customUI.xml | grep -q PPTKGridMacMenu"
check "scripts are self-contained (common.sh inlined, version filled)" "! grep -l 'DIR/common.sh\|__PPTK_VERSION__' '$COMP/Scripts/preinstall' '$COMP/Scripts/postinstall' '$COMP/Payload/Library/Application Support/PPTK/'*.sh >/dev/null"
check "minimum macOS 13 and both architectures in Distribution" "grep -q 'os-version min=\"13.0\"' '$WORK/a/Distribution' && grep -q 'hostArchitectures=\"arm64,x86_64\"' '$WORK/a/Distribution'"
APPX="$COMP/Payload/Applications/PPTK Keys.app"
check "helper is universal (arm64 + x86_64)" "lipo -info '$APPX/Contents/MacOS/PPTKKeys' | grep -q 'x86_64 arm64\|arm64 x86_64'"
check "helper bundle id and version (Intune detection)" "[ \"\$(defaults read '$APPX/Contents/Info' CFBundleIdentifier)\" = com.pptk.keys ] && [ \"\$(defaults read '$APPX/Contents/Info' CFBundleShortVersionString)\" = '$VA' ]"
check "helper signature verifies (hardened runtime, Apple Events entitlement)" "codesign --verify --strict '$APPX' 2>/dev/null && codesign -d --entitlements - '$APPX' 2>/dev/null | grep -q automation.apple-events && codesign -dv '$APPX' 2>&1 | grep -q 'flags=.*runtime'"
check "helper --version reports $VA" "[ \"\$('$APPX/Contents/MacOS/PPTKKeys' --version)\" = 'PPTKKeys $VA' ]"
check "LaunchAgent: label, program, Aqua only" "plutil -lint '$COMP/Payload/Library/LaunchAgents/com.pptk.keys.plist' >/dev/null && [ \"\$(plutil -extract Label raw '$COMP/Payload/Library/LaunchAgents/com.pptk.keys.plist')\" = com.pptk.keys ] && [ \"\$(plutil -extract ProgramArguments.0 raw '$COMP/Payload/Library/LaunchAgents/com.pptk.keys.plist')\" = '/Applications/PPTK Keys.app/Contents/MacOS/PPTKKeys' ] && [ \"\$(plutil -extract LimitLoadToSessionType raw '$COMP/Payload/Library/LaunchAgents/com.pptk.keys.plist')\" = Aqua ]"
check "PPPC profile is a valid plist" "plutil -lint '$REL/PPTK-PPPC.mobileconfig' >/dev/null"
check "managed login items profile: valid, rule on the LaunchAgent label" "plutil -lint '$REL/PPTK-LoginItems.mobileconfig' >/dev/null && grep -q '<string>com.apple.servicemanagement</string>' '$REL/PPTK-LoginItems.mobileconfig' && grep -q '<string>com.pptk.keys</string>' '$REL/PPTK-LoginItems.mobileconfig' && [ -f '$REL/intune/PPTK-LoginItems.mobileconfig' ]"
check "release SHA256SUMS match" "(cd '$REL' && shasum -a 256 -c SHA256SUMS.txt >/dev/null)"

# ---------------------------------------------------------------- 2. one PKG for both channels
echo "-- ZIP and Intune carry the same PKG"
mkdir "$WORK/zip" && ditto -x -k "$REL/PPTK-Mac-$VA-Installer.zip" "$WORK/zip"
Z="$WORK/zip/PPTK-Mac-$VA-Installer"
check "ZIP contents" "[ -f '$Z/PPTK-$VA.pkg' ] && [ -x '$Z/Install PPTK.command' ] && [ -x '$Z/Uninstall PPTK.command' ] && [ -x '$Z/pptk-uninstall.sh' ] && [ -x '$Z/pptk-detect.sh' ] && [ -f '$Z/INSTALL-MAC.md' ] && [ -f '$Z/INSTALL-MAC-AR.md' ]"
check "ZIP checksums match; ZIP PKG = release PKG = Intune PKG" "(cd '$Z' && shasum -a 256 -c SHA256SUMS.txt >/dev/null) && cmp -s '$Z/PPTK-$VA.pkg' '$REL/PPTK-$VA.pkg' && cmp -s '$REL/intune/PPTK-$VA.pkg' '$REL/PPTK-$VA.pkg'"
check "command scripts parse" "bash -n '$Z/Install PPTK.command' && bash -n '$Z/Uninstall PPTK.command'"
pkgutil --expand-full "$NEWER/intune/PPTK-$VB.pkg" "$WORK/b" >/dev/null
check "newer PKG: same identifier and install paths, higher version" "grep -q 'identifier=\"com.pptk.mac.pkg\"' '$WORK/b/component.pkg/PackageInfo' && grep -q 'version=\"$VB\"' '$WORK/b/component.pkg/PackageInfo' && [ \"\$(cd '$WORK/b/component.pkg/Payload' && find . -type f | sort)\" = \"\$(cd '$COMP/Payload' && find . -type f | sort)\" ]"

# ---------------------------------------------------------------- 3. fake root, stub launchctl, simulated Intune detection
FAKE="$WORK/root"
STUBLOG="$WORK/launchctl.log"
printf '#!/bin/bash\necho "$*" >> "%s"\n[ "$1" = print ] && echo "state = running"\nexit 0\n' "$STUBLOG" > "$WORK/launchctl"
chmod 755 "$WORK/launchctl"
export PPTK_TEST_ROOT="$FAKE" PPTK_LAUNCHCTL="$WORK/launchctl" PPTK_TEST_CONSOLE_USERS="alice:501"
mkdir -p "$FAKE/Users/alice/Library" "$FAKE/Applications" "$FAKE/Library/LaunchAgents"
ADDIN="$FAKE/$STARTUP_REL/PPTK.ppam"
DETECT="$FAKE/Library/Application Support/PPTK/pptk-detect.sh"

place_payload() { # like the installer: each file written beside its target, then renamed into place
    (cd "$1" && find . -type f) | while read -r rel; do
        target="$FAKE/${rel#./}"
        mkdir -p "$(dirname "$target")"
        cp -p "$1/$rel" "$target.pptk-new" && mv -f "$target.pptk-new" "$target"
    done
}
install_pkg() { # install_pkg <expanded component.pkg>   -> exit code of the install
    "$1/Scripts/preinstall" "" "" "/" >"$WORK/pre.out" 2>&1 || return $?
    place_payload "$1/Payload"
    "$1/Scripts/postinstall" "" "" "/" >"$WORK/post.out" 2>&1
}
intune_detect() { # Intune "macOS app (PKG)" rule: app with bundle id com.pptk.keys and this version present
    local plist="$FAKE/Applications/PPTK Keys.app/Contents/Info"
    [ -f "$plist.plist" ] && [ "$(defaults read "$plist" CFBundleIdentifier 2>/dev/null)" = com.pptk.keys ] \
        && [ "$(defaults read "$plist" CFBundleShortVersionString 2>/dev/null)" = "$1" ]
}
counts() { echo "apps=$(find "$FAKE" -name 'PPTK Keys.app' -maxdepth 3 | wc -l | tr -d ' ') agents=$(find "$FAKE" -name com.pptk.keys.plist -path '*LaunchAgents*' | wc -l | tr -d ' ') addins=$(find "$FAKE/$STARTUP_REL" -name '*.ppam' | wc -l | tr -d ' ')"; }

echo "-- a development install exists (it is migrated, so only one helper remains)"
mkdir -p "$FAKE/Users/alice/Library/LaunchAgents" "$FAKE/Users/alice/Library/Application Support/PPTK/PPTKKeys.app"
echo "<plist/>" > "$FAKE/Users/alice/Library/LaunchAgents/com.pptk.keys.plist"
echo "#!/bin/zsh" > "$FAKE/Users/alice/Library/Application Support/PPTK/pptkkeys-ctl"

echo "-- step 1: IT installs $VA manually (ZIP)"
install_pkg "$COMP"; rc=$?
check "manual install exit 0" "[ $rc = 0 ]"
check "development agent moved to backups" "[ ! -e '$FAKE/Users/alice/Library/LaunchAgents/com.pptk.keys.plist' ] && ls '$FAKE/Users/alice/Library/Application Support/PPTK/backups/'legacy-install-*/com.pptk.keys.plist >/dev/null 2>&1"
check "helper stopped then started for the logged-in user" "grep -q 'bootout gui/501/com.pptk.keys' '$STUBLOG' && grep -q 'bootstrap gui/501 $FAKE/Library/LaunchAgents/com.pptk.keys.plist' '$STUBLOG'"
check "add-in in the Startup folder, root-only location" "[ -f '$ADDIN' ] && unzip -p '$ADDIN' customUI/customUI.xml | grep -q PPTKGridMacMenu"
check "machine and user detection pass" "'$DETECT' --machine >/dev/null && '$DETECT' | grep -q 'PPTK $VA installed'"
check "one helper, one agent, one add-in ($(counts))" "[ \"\$(counts)\" = 'apps=1 agents=1 addins=1' ]"
out=$(PPTK_SYSTEM_ROOT="$FAKE/Library/Application Support/PPTK" PPTK_APPLESCRIPTS_DIR="$FAKE/Users/alice/Library/Application Scripts/com.microsoft.Powerpoint" "$FAKE/Applications/PPTK Keys.app/Contents/MacOS/PPTKKeys" --sync-plugin)
check "helper copies the Email Slides plugin for the user ($out)" "cmp -s '$FAKE/Users/alice/Library/Application Scripts/com.microsoft.Powerpoint/InstrumentaAppleScriptPlugin.applescript' '$FAKE/Library/Application Support/PPTK/InstrumentaAppleScriptPlugin.applescript'"
check "a changed file is detected" "cp '$ADDIN' '$WORK/addin.bak' && echo x >> '$ADDIN' && '$DETECT' --machine | grep -q changed; rc2=\$?; cp '$WORK/addin.bak' '$ADDIN'; [ \$rc2 = 0 ]"

echo "-- step 2: the $VA Intune app is assigned to this Mac"
check "Intune detects the manual install as $VA (no reinstall)" "intune_detect '$VA'"
: > "$STUBLOG"
install_pkg "$COMP"; rc=$?   # e.g. IT runs the ZIP again, or Intune reinstalls the same version
check "re-running the same version is harmless (exit 0, still one of each)" "[ $rc = 0 ] && '$DETECT' >/dev/null && [ \"\$(counts)\" = 'apps=1 agents=1 addins=1' ]"

echo "-- step 3: PowerPoint is open with the $VA add-in loaded"
python3 - "$ADDIN" "$WORK/holder.ready" "$WORK/holder.go" "$WORK/holder.out" <<'EOF' &
import hashlib, os, sys, time
path, ready, go, out = sys.argv[1:5]
f = open(path, "rb")                      # PowerPoint keeps the add-in it loaded open
open(ready, "w").close()
while not os.path.exists(go): time.sleep(0.05)
f.seek(0); data = f.read()
open(out, "w").write(hashlib.sha256(data).hexdigest())
EOF
HOLDER=$!
for i in $(seq 1 50); do [ -f "$WORK/holder.ready" ] && break; sleep 0.1; done
OLD_SHA=$(shasum -a 256 "$ADDIN" | awk '{print $1}')

echo "-- step 4: Intune publishes $VB (same Intune app, new PKG)"
check "Intune detection for $VB fails on the $VA Mac, so Intune installs" "! intune_detect '$VB'"
: > "$STUBLOG"
install_pkg "$WORK/b/component.pkg"; rc=$?
touch "$WORK/holder.go"; wait $HOLDER 2>/dev/null; HOLDER=""
NEW_SHA=$(shasum -a 256 "$ADDIN" | awk '{print $1}')
check "Intune upgrade exit 0" "[ $rc = 0 ]"
check "Intune detection for $VB now succeeds" "intune_detect '$VB'"
check "helper restarted on the new version" "grep -q 'bootout gui/501/com.pptk.keys' '$STUBLOG' && grep -q 'bootstrap gui/501' '$STUBLOG' && [ \"\$('$FAKE/Applications/PPTK Keys.app/Contents/MacOS/PPTKKeys' --version)\" = 'PPTKKeys $VB' ]"
check "the open add-in kept reading the old version, undamaged" "[ \"\$(cat '$WORK/holder.out')\" = '$OLD_SHA' ]"
check "the Startup folder now holds the new add-in (used at the next PowerPoint start)" "[ '$NEW_SHA' != '$OLD_SHA' ] && [ \"\$(shasum -a 256 '$WORK/b/component.pkg/Payload/$STARTUP_REL/PPTK.ppam' | awk '{print \$1}')\" = '$NEW_SHA' ]"
check "no duplicates after the update ($(counts))" "[ \"\$(counts)\" = 'apps=1 agents=1 addins=1' ] && [ ! -e '$ADDIN.pptk-new' ]"
check "detection reports $VB" "'$DETECT' | grep -q 'PPTK $VB installed'"

echo "-- step 5: an older PKG is refused"
before=$(find "$FAKE/Applications" "$FAKE/Library" -type f -exec shasum -a 256 {} + | sort | shasum)
install_pkg "$COMP"; rc=$?
after=$(find "$FAKE/Applications" "$FAKE/Library" -type f -exec shasum -a 256 {} + | sort | shasum)
check "downgrade $VB -> $VA refused (exit 1), nothing changed" "[ $rc = 1 ] && [ '$before' = '$after' ] && grep -q refusing '$WORK/pre.out'"

echo "-- step 6: uninstall"
echo "other add-in" > "$FAKE/$STARTUP_REL/Other.ppam"
"$FAKE/Library/Application Support/PPTK/pptk-uninstall.sh" >"$WORK/un.out" 2>&1; rc=$?
check "uninstall exit 0" "[ $rc = 0 ]"
check "helper, agent, add-in and package files removed" "[ ! -e '$FAKE/Applications/PPTK Keys.app' ] && [ ! -e '$FAKE/Library/LaunchAgents/com.pptk.keys.plist' ] && [ ! -e '$ADDIN' ] && [ ! -e '$FAKE/Library/Application Support/PPTK' ]"
check "other add-ins in the Startup folder, the user's plugin and backups are kept" "[ -f '$FAKE/$STARTUP_REL/Other.ppam' ] && [ -f '$FAKE/Users/alice/Library/Application Scripts/com.microsoft.Powerpoint/InstrumentaAppleScriptPlugin.applescript' ] && ls -d '$FAKE/Users/alice/Library/Application Support/PPTK/backups/'legacy-install-* >/dev/null 2>&1"
check "detection fails after uninstall" "! '$Z/pptk-detect.sh' --machine >/dev/null"

echo
echo "== $pass passed, $fail failed"
[ $fail = 0 ]
