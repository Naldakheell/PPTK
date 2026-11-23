#!/bin/bash
# PPTK for Mac - uninstall (run as root: sudo, Intune shell script, or "Uninstall PPTK.command").
# Removes the helper, the system-wide LaunchAgent, the add-in in Office's Startup folder, the package files, each
# user's PPTK logs, and the installer receipt. Kept: legacy backups (~/Library/Application Support/PPTK/backups) and the shared
# Instrumenta AppleScript plugin (also used by Instrumenta).
# An open PowerPoint keeps the loaded add-in until it quits; after that it is no longer loaded. PowerPoint may keep a
# "PPTK" line under Tools > PowerPoint Add-ins until the user removes it ("-").
set -u
DIR="$(cd "$(dirname "$0")" && pwd)"
. "$DIR/common.sh"

if [ -z "$ROOT" ] && [ "$(/usr/bin/id -u)" != "0" ]; then
    echo "Run as root: sudo \"$0\"" >&2
    exit 1
fi

INSTALLED=$(json_version "$SYS/version.json")

stop_helpers
/bin/rm -f "$AGENT"
/bin/rm -rf "$APP"
/bin/rm -f "$ADDIN"   # PPTK's own file only; other add-ins in the Startup folder are not touched

user_homes | while read -r home; do
    /bin/rm -rf "$home/$USER_LOGS"
    # ~/Library/Application Scripts/com.microsoft.Powerpoint/InstrumentaAppleScriptPlugin.applescript is kept:
    # Instrumenta uses the same file, and Office's own container is not inspected (macOS protects it).
    /bin/rmdir "$home/$USER_PPTK" 2>/dev/null   # only when empty (backups stay)
done

/bin/rm -rf "$SYS"
if [ -z "$ROOT" ]; then
    /usr/sbin/pkgutil --forget "$PKG_ID" >/dev/null 2>&1
fi
log "uninstalled${INSTALLED:+ $INSTALLED}. Remove 'PPTK' from PowerPoint > Tools > PowerPoint Add-ins if it is still listed."
exit 0
