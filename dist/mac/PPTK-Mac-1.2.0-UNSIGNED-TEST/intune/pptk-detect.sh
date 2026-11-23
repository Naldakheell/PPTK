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
# ---- common.sh (inlined at build) ----
# Shared paths and helpers for the PPTK for Mac package scripts (preinstall, postinstall, uninstall, detect).
# One definition of every path, label and identifier, so the ZIP and Intune installs are the same install.
#
# Tests only: PPTK_TEST_ROOT prefixes every absolute path (a fake "/"), PPTK_LAUNCHCTL replaces launchctl,
# PPTK_TEST_CONSOLE_USERS="name:uid ..." replaces the logged-in user lookup. Production never sets them.

ROOT="${PPTK_TEST_ROOT:-}"
# preinstall / postinstall set PPTK_TARGET to the installer's target volume ($3). Another volume than "/" is
# installed into, but no helper of the running system is stopped or started.
OTHER_VOLUME=0
if [ -z "$ROOT" ] && [ -n "${PPTK_TARGET:-}" ] && [ "${PPTK_TARGET%/}" != "" ]; then
    ROOT="${PPTK_TARGET%/}"
    OTHER_VOLUME=1
fi
LAUNCHCTL="${PPTK_LAUNCHCTL:-/bin/launchctl}"

LABEL="com.pptk.keys"                       # LaunchAgent label = helper bundle id
PKG_ID="com.pptk.mac.pkg"                   # installer receipt id
APP="$ROOT/Applications/PPTK Keys.app"      # helper (Intune detection: bundle id + version)
EXE="$APP/Contents/MacOS/PPTKKeys"
SYS="$ROOT/Library/Application Support/PPTK" # version.json, Email Slides plugin, uninstall/detect (root-owned)
# The add-in: Office's system-wide Startup folder. PowerPoint loads every add-in in it at start, for every user,
# without registration or a file-access dialog (verified 2026-11-22). Root-only, so only the package changes it.
STARTUP_DIR="$ROOT/Library/Application Support/Microsoft/Office365/User Content.localized/Startup/PowerPoint"
ADDIN="$STARTUP_DIR/PPTK.ppam"
AGENT="$ROOT/Library/LaunchAgents/$LABEL.plist"
USERS_DIR="$ROOT/Users"

# Per-user locations (relative to a home folder)
USER_PPTK="Library/Application Support/PPTK"           # backups of a migrated development install
USER_AGENT="Library/LaunchAgents/$LABEL.plist"          # legacy per-user agent (development installs)
USER_LOGS="Library/Logs/PPTK"
USER_SCRIPTS="Library/Application Scripts/com.microsoft.Powerpoint"

log() { echo "PPTK: $*"; }

# "name:uid" for every user with a graphical login session now (none when installing onto another volume).
console_users() {
    [ "$OTHER_VOLUME" = "1" ] && return
    if [ -n "${PPTK_TEST_CONSOLE_USERS+x}" ]; then
        for entry in $PPTK_TEST_CONSOLE_USERS; do echo "$entry"; done
        return
    fi
    /usr/bin/who | /usr/bin/awk '$2 == "console" {print $1}' | /usr/bin/sort -u | while read -r name; do
        [ "$name" = "root" ] && continue
        uid=$(/usr/bin/id -u "$name" 2>/dev/null) || continue
        echo "$name:$uid"
    done
}

# Prints the version in a version.json, or nothing.
json_version() {
    [ -f "$1" ] || return 0
    /usr/bin/sed -n 's/.*"version"[[:space:]]*:[[:space:]]*"\([^"]*\)".*/\1/p' "$1" | /usr/bin/head -1
}

# version_gt A B: true when A > B (numeric dotted versions).
version_gt() {
    /usr/bin/awk -v a="$1" -v b="$2" 'BEGIN {
        n = split(a, x, "."); m = split(b, y, "."); k = (n > m) ? n : m
        for (i = 1; i <= k; i++) { p = x[i] + 0; q = y[i] + 0; if (p > q) exit 0; if (p < q) exit 1 }
        exit 1 }'
}

# Stops the helper agent in every logged-in user's session (either the package's or a legacy per-user one).
stop_helpers() {
    console_users | while IFS=: read -r name uid; do
        "$LAUNCHCTL" bootout "gui/$uid/$LABEL" >/dev/null 2>&1 && log "stopped helper for $name"
    done
    return 0
}

# Starts the package's agent in every logged-in user's session (other users get it at their next login).
start_helpers() {
    [ -f "$AGENT" ] || return 0
    console_users | while IFS=: read -r name uid; do
        if "$LAUNCHCTL" bootstrap "gui/$uid" "$AGENT" >/dev/null 2>&1; then
            log "started helper for $name"
        else
            "$LAUNCHCTL" kickstart -k "gui/$uid/$LABEL" >/dev/null 2>&1 && log "restarted helper for $name"
        fi
    done
    return 0
}

# Every user home folder (for per-user migration and cleanup).
user_homes() {
    for home in "$USERS_DIR"/*; do
        [ -d "$home/Library" ] || continue
        case "$(basename "$home")" in Shared|Guest|.localized) continue ;; esac
        echo "$home"
    done
}
# ---- end common.sh ----

MACHINE_ONLY=0
if [ "${1:-}" = "--machine" ]; then MACHINE_ONLY=1; shift; fi
EXPECTED="${1:-1.2.0}"
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
