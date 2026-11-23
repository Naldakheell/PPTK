#!/bin/bash
# Builds the PPTK for Mac release: one PKG, the IT ZIP that wraps it, and the Intune folder.
#
#   src/Distribution/mac/build-dist.sh [--ppam <file>] [--out <dir>]
#
# Signing (all optional; without them the output is an UNSIGNED TEST build, never for employees):
#   PPTK_APP_SIGN_ID   "Developer ID Application: <Org> (<TEAMID>)"   signs PPTK Keys.app (hardened runtime)
#   PPTK_PKG_SIGN_ID   "Developer ID Installer: <Org> (<TEAMID>)"     signs the PKG
#   PPTK_NOTARY_PROFILE  notarytool keychain profile (xcrun notarytool store-credentials ...)  notarize + staple
# Tests only: PPTK_VERSION_OVERRIDE builds another version number without the add-in version check.
#
# Inputs: the exported add-in (default bin/PPTK DEV.ppam), PPTKKeys sources, bin/InstrumentaAppleScriptPlugin.applescript.
# Needs: Xcode Command Line Tools (swift, lipo, codesign, pkgbuild, productbuild, notarytool), python3 + oletools.
set -euo pipefail

HERE="$(cd "$(dirname "$0")" && pwd)"
REPO="$(cd "$HERE/../../.." && pwd)"
PPAM="$REPO/bin/PPTK DEV.ppam"
OUT="$REPO/dist/mac"
while [ $# -gt 0 ]; do
    case "$1" in
        --ppam) PPAM="$2"; shift 2 ;;
        --out) OUT="$2"; shift 2 ;;
        *) echo "unknown option $1"; exit 2 ;;
    esac
done

LABEL="com.pptk.keys"
PKG_ID="com.pptk.mac.pkg"
MIN_MACOS="13.0"
APP_NAME="PPTK Keys.app"

# ---------------------------------------------------------------- version (single source: ModulePPTKBrand.bas)
SRC_VERSION=$(sed -n 's/^Public Const PPTK_VERSION As String = "\(.*\)".*/\1/p' "$REPO/src/Modules/ModulePPTKBrand.bas")
[ -n "$SRC_VERSION" ] || { echo "PPTK_VERSION not found in ModulePPTKBrand.bas"; exit 1; }
[ -f "$PPAM" ] || { echo "add-in not found: $PPAM"; exit 1; }
PPAM_VERSION=$(python3 - "$PPAM" <<'EOF'
import re, sys
from oletools.olevba import VBA_Parser
for _, _, name, code in VBA_Parser(sys.argv[1]).extract_macros():
    m = re.search(r'Public Const PPTK_VERSION As String = "([^"]+)"', code)
    if m: print(m.group(1)); break
EOF
)
VERSION="${PPTK_VERSION_OVERRIDE:-$SRC_VERSION}"
if [ -z "${PPTK_VERSION_OVERRIDE:-}" ] && [ "$PPAM_VERSION" != "$SRC_VERSION" ]; then
    echo "The add-in contains PPTK_VERSION '$PPAM_VERSION' but the source says '$SRC_VERSION'. Export the add-in again."
    exit 1
fi
# The add-in must carry the Mac ribbon (collapsed groups); a Windows-ribbon export would scroll on Mac.
python3 - "$PPAM" <<'EOF'
import sys, zipfile
x = zipfile.ZipFile(sys.argv[1]).read("customUI/customUI.xml").decode("utf-8")
assert 'id="PPTKGridMacMenu"' in x and 'image="pptk-replace-fonts"' in x, "the add-in does not contain the Mac ribbon (make_mac_ribbon.py)"
EOF

SIGNED=0
[ -n "${PPTK_APP_SIGN_ID:-}" ] && [ -n "${PPTK_PKG_SIGN_ID:-}" ] && SIGNED=1
SUFFIX=""
[ "$SIGNED" = "1" ] || SUFFIX="-UNSIGNED-TEST"
[ -z "${PPTK_VERSION_OVERRIDE:-}" ] || SUFFIX="$SUFFIX-VERSIONTEST"
NAME="PPTK-Mac-$VERSION$SUFFIX"
WORK="$(mktemp -d)"
trap 'rm -rf "$WORK"' EXIT
REL="$OUT/$NAME"
rm -rf "$REL"
mkdir -p "$REL"
echo "== PPTK for Mac $VERSION ($([ "$SIGNED" = 1 ] && echo "Developer ID signed" || echo "UNSIGNED TEST build"))"

# ---------------------------------------------------------------- helper: universal binary (arm64 + x86_64)
HELPER_SRC="$REPO/src/PPTKKeys/mac"
for triple in arm64-apple-macosx$MIN_MACOS x86_64-apple-macosx$MIN_MACOS; do
    (cd "$HELPER_SRC" && swift build -c release --triple "$triple" --build-path "$WORK/build/$triple" >/dev/null)
done
ARM=$(find "$WORK/build/arm64-apple-macosx$MIN_MACOS" -path '*release/PPTKKeys' -type f | head -1)
X86=$(find "$WORK/build/x86_64-apple-macosx$MIN_MACOS" -path '*release/PPTKKeys' -type f | head -1)

ROOT="$WORK/root"
APP="$ROOT/Applications/$APP_NAME"
SYS="$ROOT/Library/Application Support/PPTK"
STARTUP_REL="Library/Application Support/Microsoft/Office365/User Content.localized/Startup/PowerPoint"
mkdir -p "$APP/Contents/MacOS" "$SYS" "$ROOT/Library/LaunchAgents" "$ROOT/$STARTUP_REL"
lipo -create "$ARM" "$X86" -output "$APP/Contents/MacOS/PPTKKeys"
chmod 755 "$APP/Contents/MacOS/PPTKKeys"
cat > "$APP/Contents/Info.plist" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>CFBundleIdentifier</key><string>$LABEL</string>
  <key>CFBundleName</key><string>PPTK Keys</string>
  <key>CFBundleDisplayName</key><string>PPTK Keys</string>
  <key>CFBundleExecutable</key><string>PPTKKeys</string>
  <key>CFBundlePackageType</key><string>APPL</string>
  <key>CFBundleShortVersionString</key><string>$VERSION</string>
  <key>CFBundleVersion</key><string>$VERSION</string>
  <key>LSMinimumSystemVersion</key><string>$MIN_MACOS</string>
  <key>LSUIElement</key><true/>
  <key>NSAppleEventsUsageDescription</key><string>PPTK Keys asks PowerPoint for the current selection and runs the PPTK add-in commands for your keyboard shortcuts.</string>
  <key>NSHumanReadableCopyright</key><string>PPTK</string>
</dict>
</plist>
EOF
cat > "$WORK/entitlements.plist" <<'EOF'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <!-- Hardened runtime: sending Apple Events to PowerPoint needs this entitlement. -->
  <key>com.apple.security.automation.apple-events</key><true/>
</dict>
</plist>
EOF
if [ "$SIGNED" = "1" ]; then
    codesign --force --options runtime --timestamp --entitlements "$WORK/entitlements.plist" \
        --identifier "$LABEL" --sign "$PPTK_APP_SIGN_ID" "$APP"
else
    # Ad-hoc: runs on the build Mac and in tests. Not for distribution (Gatekeeper / PPPC need Developer ID).
    codesign --force --options runtime --entitlements "$WORK/entitlements.plist" --identifier "$LABEL" --sign - "$APP"
fi
codesign --verify --strict "$APP"
APP_REQ=$(codesign -dr - "$APP" 2>&1 | sed -n 's/^designated => //p')

# ---------------------------------------------------------------- package files
# The add-in goes into Office's system-wide Startup folder: PowerPoint loads it for every user at start.
cp "$PPAM" "$ROOT/$STARTUP_REL/PPTK.ppam"
cp "$REPO/bin/InstrumentaAppleScriptPlugin.applescript" "$SYS/InstrumentaAppleScriptPlugin.applescript"
# Scripts are made self-contained (common.sh inlined) so each can be used alone (Intune runs single files).
inline() { # inline <script> <out>
    python3 - "$1" "$HERE/scripts/common.sh" "$2" "$VERSION" <<'EOF'
import sys
src, common, out, version = sys.argv[1:5]
text = open(src).read()
lib = open(common).read().split("\n", 1)[1]           # drop the shebang
text = text.replace('. "$DIR/common.sh"', "# ---- common.sh (inlined at build) ----\n" + lib + "# ---- end common.sh ----")
text = text.replace("__PPTK_VERSION__", version)
open(out, "w").write(text)
EOF
    chmod 755 "$2"
}
SCRIPTS="$WORK/scripts"
mkdir -p "$SCRIPTS"
inline "$HERE/scripts/preinstall" "$SCRIPTS/preinstall"
inline "$HERE/scripts/postinstall" "$SCRIPTS/postinstall"
inline "$HERE/scripts/pptk-uninstall.sh" "$SYS/pptk-uninstall.sh"
inline "$HERE/scripts/pptk-detect.sh" "$SYS/pptk-detect.sh"

cat > "$ROOT/Library/LaunchAgents/$LABEL.plist" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key><string>$LABEL</string>
  <key>ProgramArguments</key>
  <array><string>/Applications/$APP_NAME/Contents/MacOS/PPTKKeys</string></array>
  <key>EnvironmentVariables</key>
  <dict><key>PPTK_LOG_TO_USER_LIBRARY</key><string>1</string></dict>
  <key>RunAtLoad</key><true/>
  <key>KeepAlive</key>
  <dict><key>SuccessfulExit</key><false/></dict>
  <key>ThrottleInterval</key><integer>30</integer>
  <key>ProcessType</key><string>Interactive</string>
  <key>LimitLoadToSessionType</key><string>Aqua</string>
</dict>
</plist>
EOF
plutil -lint "$ROOT/Library/LaunchAgents/$LABEL.plist" >/dev/null

# version.json: one "files" entry per line, paths relative to "/" (pptk-detect.sh reads it with sed)
COMMIT=$(git -C "$REPO" rev-parse --short HEAD 2>/dev/null || echo unknown)
DIRTY=$([ -n "$(git -C "$REPO" status --porcelain 2>/dev/null)" ] && echo true || echo false)
{
    echo "{"
    echo "  \"schemaVersion\": 1,"
    echo "  \"product\": \"PPTK for Mac\","
    echo "  \"version\": \"$VERSION\","
    echo "  \"bundleId\": \"$LABEL\","
    echo "  \"pkgId\": \"$PKG_ID\","
    echo "  \"minimumMacOS\": \"$MIN_MACOS\","
    echo "  \"architectures\": [\"arm64\", \"x86_64\"],"
    echo "  \"signed\": $([ "$SIGNED" = 1 ] && echo true || echo false),"
    echo "  \"build\": {\"commit\": \"$COMMIT\", \"dirty\": $DIRTY, \"date\": \"$(date -u +%Y-%m-%dT%H:%M:%SZ)\"},"
    echo "  \"files\": {"
    first=1
    for rel in "$STARTUP_REL/PPTK.ppam" \
               "Library/Application Support/PPTK/InstrumentaAppleScriptPlugin.applescript" \
               "Library/Application Support/PPTK/pptk-uninstall.sh" "Library/Application Support/PPTK/pptk-detect.sh" \
               "Applications/$APP_NAME/Contents/MacOS/PPTKKeys" "Applications/$APP_NAME/Contents/Info.plist"; do
        f="$ROOT/$rel"
        sha=$(shasum -a 256 "$f" | awk '{print $1}'); size=$(stat -f %z "$f")
        [ $first = 1 ] || echo ","
        printf '    "%s": {"sha256": "%s", "size": %s}' "$rel" "$sha" "$size"
        first=0
    done
    echo ""
    echo "  }"
    echo "}"
} > "$SYS/version.json"
python3 -c "import json,sys; json.load(open(sys.argv[1]))" "$SYS/version.json"

# ---------------------------------------------------------------- PKG
pkgbuild --analyze --root "$ROOT" "$WORK/component.plist" >/dev/null
# The helper must always go to /Applications: never "relocated" onto another copy with the same bundle id
# (for example a development copy in ~/Library/Application Support/PPTK).
python3 - "$WORK/component.plist" <<'EOF'
import plistlib, sys
p = sys.argv[1]; items = plistlib.load(open(p, "rb"))
for item in items:
    item["BundleIsRelocatable"] = False
    item["BundleOverwriteAction"] = "upgrade"
plistlib.dump(items, open(p, "wb"))
EOF
pkgbuild --root "$ROOT" --component-plist "$WORK/component.plist" --identifier "$PKG_ID" --version "$VERSION" \
    --scripts "$SCRIPTS" --install-location / --ownership recommended "$WORK/component.pkg" >/dev/null
cat > "$WORK/distribution.xml" <<EOF
<?xml version="1.0" encoding="utf-8"?>
<installer-gui-script minSpecVersion="2">
  <title>PPTK for Mac $VERSION</title>
  <options customize="never" require-scripts="false" hostArchitectures="arm64,x86_64"/>
  <domains enable_localSystem="true" enable_currentUserHome="false" enable_anywhere="false"/>
  <volume-check><allowed-os-versions><os-version min="$MIN_MACOS"/></allowed-os-versions></volume-check>
  <choices-outline><line choice="pptk"/></choices-outline>
  <choice id="pptk" title="PPTK for Mac" visible="false"><pkg-ref id="$PKG_ID"/></choice>
  <pkg-ref id="$PKG_ID" version="$VERSION" onConclusion="none">component.pkg</pkg-ref>
</installer-gui-script>
EOF
PKG="$REL/PPTK-$VERSION.pkg"
if [ "$SIGNED" = "1" ]; then
    productbuild --distribution "$WORK/distribution.xml" --package-path "$WORK" --sign "$PPTK_PKG_SIGN_ID" --timestamp "$PKG" >/dev/null
else
    productbuild --distribution "$WORK/distribution.xml" --package-path "$WORK" "$PKG" >/dev/null
fi
NOTARIZED=false
if [ "$SIGNED" = "1" ] && [ -n "${PPTK_NOTARY_PROFILE:-}" ]; then
    xcrun notarytool submit "$PKG" --keychain-profile "$PPTK_NOTARY_PROFILE" --wait
    xcrun stapler staple "$PKG"
    NOTARIZED=true
fi

# ---------------------------------------------------------------- PPPC profile (Accessibility + Automation)
PP_REQ=$(codesign -dr - "/Applications/Microsoft PowerPoint.app" 2>&1 | sed -n 's/^designated => //p' | sed 's#/\* exists \*/##g; s/  */ /g')
[ -n "$PP_REQ" ] || PP_REQ='identifier "com.microsoft.Powerpoint" and anchor apple generic and certificate 1[field.1.2.840.113635.100.6.2.6] and certificate leaf[field.1.2.840.113635.100.6.1.13] and certificate leaf[subject.OU] = UBF8T346G9'
xml_escape() { sed 's/&/\&amp;/g; s/</\&lt;/g; s/>/\&gt;/g'; }
APP_REQ_X=$(printf '%s' "$APP_REQ" | xml_escape)
PP_REQ_X=$(printf '%s' "$PP_REQ" | xml_escape)
PPPC="$REL/PPTK-PPPC.mobileconfig"
cat > "$PPPC" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<!-- PPTK Keys privacy permissions (PPPC). Deploy with Intune as a custom macOS configuration profile.
     Grants Accessibility and Automation (PowerPoint). Input Monitoring cannot be granted by a profile (Apple):
     it lets a standard user approve it instead. The code requirement below is the one of THIS build:
     $([ "$SIGNED" = 1 ] && echo "Developer ID signed, stable across PPTK updates signed by the same team." || echo "UNSIGNED TEST build: ad-hoc requirement (cdhash) - valid for this exact binary only. Do not deploy.") -->
<plist version="1.0">
<dict>
  <key>PayloadType</key><string>Configuration</string>
  <key>PayloadVersion</key><integer>1</integer>
  <key>PayloadIdentifier</key><string>com.pptk.keys.pppc</string>
  <key>PayloadUUID</key><string>$(uuidgen)</string>
  <key>PayloadDisplayName</key><string>PPTK Keys permissions</string>
  <key>PayloadScope</key><string>System</string>
  <key>PayloadContent</key>
  <array>
    <dict>
      <key>PayloadType</key><string>com.apple.TCC.configuration-profile-policy</string>
      <key>PayloadVersion</key><integer>1</integer>
      <key>PayloadIdentifier</key><string>com.pptk.keys.pppc.tcc</string>
      <key>PayloadUUID</key><string>$(uuidgen)</string>
      <key>PayloadDisplayName</key><string>PPTK Keys privacy</string>
      <key>Services</key>
      <dict>
        <key>Accessibility</key>
        <array><dict>
          <key>Identifier</key><string>$LABEL</string>
          <key>IdentifierType</key><string>bundleID</string>
          <key>CodeRequirement</key><string>$APP_REQ_X</string>
          <key>StaticCode</key><false/>
          <key>Authorization</key><string>Allow</string>
          <key>Comment</key><string>Keyboard shortcuts: reads the focused PowerPoint element</string>
        </dict></array>
        <key>AppleEvents</key>
        <array><dict>
          <key>Identifier</key><string>$LABEL</string>
          <key>IdentifierType</key><string>bundleID</string>
          <key>CodeRequirement</key><string>$APP_REQ_X</string>
          <key>AEReceiverIdentifier</key><string>com.microsoft.Powerpoint</string>
          <key>AEReceiverIdentifierType</key><string>bundleID</string>
          <key>AEReceiverCodeRequirement</key><string>$PP_REQ_X</string>
          <key>Authorization</key><string>Allow</string>
          <key>Comment</key><string>Runs PPTK add-in commands in PowerPoint for the keyboard shortcuts</string>
        </dict></array>
        <key>ListenEvent</key>
        <array><dict>
          <key>Identifier</key><string>$LABEL</string>
          <key>IdentifierType</key><string>bundleID</string>
          <key>CodeRequirement</key><string>$APP_REQ_X</string>
          <key>Authorization</key><string>AllowStandardUserToSetSystemService</string>
          <key>Comment</key><string>Input Monitoring: the user approves it (cannot be granted by MDM)</string>
        </dict></array>
      </dict>
    </dict>
  </array>
</dict>
</plist>
EOF
plutil -lint "$PPPC" >/dev/null

# ---------------------------------------------------------------- Managed login items (background item stays on)
LOGIN="$REL/PPTK-LoginItems.mobileconfig"
cat > "$LOGIN" <<LOGINEOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<!-- PPTK Keys managed login item (macOS 13+). The background item that starts the helper at login stays on;
     users cannot turn it off in System Settings > General > Login Items. Matches the LaunchAgent label. -->
<plist version="1.0">
<dict>
  <key>PayloadType</key><string>Configuration</string>
  <key>PayloadVersion</key><integer>1</integer>
  <key>PayloadIdentifier</key><string>com.pptk.keys.loginitems</string>
  <key>PayloadUUID</key><string>$(uuidgen)</string>
  <key>PayloadDisplayName</key><string>PPTK Keys background item</string>
  <key>PayloadScope</key><string>System</string>
  <key>PayloadContent</key>
  <array>
    <dict>
      <key>PayloadType</key><string>com.apple.servicemanagement</string>
      <key>PayloadVersion</key><integer>1</integer>
      <key>PayloadIdentifier</key><string>com.pptk.keys.loginitems.rules</string>
      <key>PayloadUUID</key><string>$(uuidgen)</string>
      <key>PayloadDisplayName</key><string>PPTK Keys</string>
      <key>Rules</key>
      <array>
        <dict>
          <key>RuleType</key><string>Label</string>
          <key>RuleValue</key><string>$LABEL</string>
          <key>Comment</key><string>PPTK Keys LaunchAgent</string>
        </dict>
      </array>
    </dict>
  </array>
</dict>
</plist>
LOGINEOF
plutil -lint "$LOGIN" >/dev/null

# ---------------------------------------------------------------- ZIP (IT manual install) and Intune folder
DOCS="$HERE/docs"
ZIPDIR="$WORK/PPTK-Mac-$VERSION-Installer"
mkdir -p "$ZIPDIR"
cp "$PKG" "$ZIPDIR/"
cp "$SYS/pptk-uninstall.sh" "$SYS/pptk-detect.sh" "$ZIPDIR/"
for f in "Install PPTK.command" "Uninstall PPTK.command"; do
    sed "s/__PPTK_VERSION__/$VERSION/g" "$HERE/zip/$f" > "$ZIPDIR/$f"
    chmod 755 "$ZIPDIR/$f"
done
sed "s/__PPTK_VERSION__/$VERSION/g" "$DOCS/INSTALL-MAC.md" > "$ZIPDIR/INSTALL-MAC.md"
sed "s/__PPTK_VERSION__/$VERSION/g" "$DOCS/INSTALL-MAC-AR.md" > "$ZIPDIR/INSTALL-MAC-AR.md"
(cd "$ZIPDIR" && shasum -a 256 "PPTK-$VERSION.pkg" pptk-uninstall.sh pptk-detect.sh "Install PPTK.command" "Uninstall PPTK.command" > SHA256SUMS.txt)
(cd "$WORK" && ditto -c -k --sequesterRsrc --keepParent "PPTK-Mac-$VERSION-Installer" "$REL/PPTK-Mac-$VERSION-Installer.zip")

INTUNE="$REL/intune"
mkdir -p "$INTUNE"
cp "$PKG" "$INTUNE/"
cp "$SYS/pptk-uninstall.sh" "$SYS/pptk-detect.sh" "$PPPC" "$LOGIN" "$INTUNE/"
sed "s/__PPTK_VERSION__/$VERSION/g" "$DOCS/INTUNE-MAC.md" > "$INTUNE/INTUNE-MAC.md"
cp "$SYS/version.json" "$REL/version.json"

(cd "$REL" && shasum -a 256 "PPTK-$VERSION.pkg" "PPTK-Mac-$VERSION-Installer.zip" PPTK-PPPC.mobileconfig PPTK-LoginItems.mobileconfig version.json \
    intune/* > SHA256SUMS.txt)
cat > "$REL/BUILD-INFO.txt" <<EOF
PPTK for Mac $VERSION
built        $(date -u +%Y-%m-%dT%H:%M:%SZ) from commit $COMMIT (working tree dirty: $DIRTY)
add-in       $PPAM (PPTK_VERSION $PPAM_VERSION) -> /$STARTUP_REL/PPTK.ppam
helper       universal (arm64, x86_64), minimum macOS $MIN_MACOS
app signing  $([ "$SIGNED" = 1 ] && echo "$PPTK_APP_SIGN_ID (hardened runtime)" || echo "ad-hoc - UNSIGNED TEST BUILD, not for distribution")
pkg signing  $([ "$SIGNED" = 1 ] && echo "$PPTK_PKG_SIGN_ID" || echo "unsigned - UNSIGNED TEST BUILD, not for distribution")
notarized    $NOTARIZED
designated   $APP_REQ
EOF
echo "== done: $REL"
ls -la "$REL"
