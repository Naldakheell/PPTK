#!/bin/bash
# PPTK for Mac - detection / health check (no changes are made).
#   pptk-detect.sh [--machine] [expected-version]
#   --machine   only the machine-wide install (skips the logged-in user's helper check)
# Exit 0 and prints "PPTK <version> installed" when the installation is complete and matches the expected
# version (default: the version this script was built with). Otherwise exit 1 with one line per problem.
# Checks the machine-wide install; when a user is logged in, also that user's helper.
# Usable by IT, by an Intune custom attribute script, or after a manual install.
set -u
DIR="$(cd "$(dirname "$0")" && pwd)"
. "$DIR/common.sh"

MACHINE_ONLY=0
if [ "${1:-}" = "--machine" ]; then MACHINE_ONLY=1; shift; fi
EXPECTED="${1:-__PPTK_VERSION__}"
problems=()
problem() { problems+=("$1"); }

# 1. Package files and their hashes (version.json)
manifest="$SYS/version.json"
version=$(json_version "$manifest")
[ -n "$version" ] || problem "missing $manifest"
[ -z "$version" ] || [ "$version" = "$EXPECTED" ] || problem "package version $version, expected $EXPECTED"
if [ -f "$manifest" ]; then
    # "files": {"<relative path>": {"sha256": "...", ...}} - one entry per line (written by build-dist.sh)
    # paths are relative to the disk root ("/")
    while IFS='|' read -r rel sha; do
        [ -n "$rel" ] || continue
        file="$ROOT/$rel"
        if [ ! -f "$file" ]; then problem "missing $file"; continue; fi
        actual=$(/usr/bin/shasum -a 256 "$file" | /usr/bin/awk '{print $1}')
        [ "$actual" = "$sha" ] || problem "changed $file"
    done < <(/usr/bin/sed -n 's/^ *"\([^"]*\)": *{"sha256": *"\([0-9a-f]*\)".*/\1|\2/p' "$manifest")
fi

# 2. Helper app: bundle id, version, signature
plist="$APP/Contents/Info.plist"
if [ -f "$plist" ]; then
    bid=$(/usr/bin/defaults read "${plist%.plist}" CFBundleIdentifier 2>/dev/null)
    bver=$(/usr/bin/defaults read "${plist%.plist}" CFBundleShortVersionString 2>/dev/null)
    [ "$bid" = "$LABEL" ] || problem "helper bundle id '$bid'"
    [ "$bver" = "$EXPECTED" ] || problem "helper version '$bver', expected $EXPECTED"
    /usr/bin/codesign --verify --strict "$APP" >/dev/null 2>&1 || problem "helper signature is not valid"
else
    problem "missing $APP"
fi

# 3. LaunchAgent (exactly one: the system-wide agent, no legacy per-user copy)
if [ -f "$AGENT" ]; then
    prog=$(/usr/bin/plutil -extract ProgramArguments.0 raw "$AGENT" 2>/dev/null)
    [ "$prog" = "${EXE#$ROOT}" ] || problem "LaunchAgent runs '$prog'"
else
    problem "missing $AGENT"
fi
while read -r home; do
    [ -f "$home/$USER_AGENT" ] && problem "legacy per-user agent $home/$USER_AGENT"
done < <(user_homes)

# 4. Installer receipt (both install methods leave the same receipt)
if [ -z "$ROOT" ]; then
    receipt=$(/usr/sbin/pkgutil --pkg-info "$PKG_ID" 2>/dev/null | /usr/bin/awk '/^version:/ {print $2}')
    [ "$receipt" = "$EXPECTED" ] || problem "receipt $PKG_ID version '${receipt:-none}'"
fi

# 5. Logged-in users: helper running (it needs Accessibility and Input Monitoring once)
while IFS=: read -r name uid; do
    [ -n "$name" ] || continue
    [ "$MACHINE_ONLY" = "1" ] && break
    if ! "$LAUNCHCTL" print "gui/$uid/$LABEL" 2>/dev/null | /usr/bin/grep -q "state = running"; then
        problem "helper not running for $name (permissions not granted yet, or just installed)"
    fi
done < <(console_users)

if [ ${#problems[@]} -eq 0 ]; then
    echo "PPTK $EXPECTED installed"
    exit 0
fi
printf 'PPTK not detected: %s\n' "${problems[@]}"
exit 1
