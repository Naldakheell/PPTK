#!/bin/zsh
# PPTKKeys background helper control (per-user LaunchAgent, no Terminal, no Dock icon).
#
#   pptkkeys-ctl install [binary]   first install, or update: back up, copy, sign, (re)start
#   pptkkeys-ctl update  [binary]   same as install (default binary: the repository release build)
#   pptkkeys-ctl start | stop | restart | status | logs
#   pptkkeys-ctl uninstall          stop, remove the agent and the app (a backup copy is kept)
#
# Installed layout (all inside the user's Library, nothing system-wide):
#   ~/Library/Application Support/PPTK/PPTKKeys.app     the helper (LSUIElement, bundle id com.pptk.keys)
#   ~/Library/Application Support/PPTK/pptkkeys-ctl     this script
#   ~/Library/Application Support/PPTK/backups/         previous apps and agent files
#   ~/Library/LaunchAgents/com.pptk.keys.plist          the LaunchAgent
#   ~/Library/Logs/PPTK/PPTKKeys.log                    helper output
#
# Signing: the local identity "PPTK Local Code Signing" (login keychain) keeps the same code identity across
# updates, so the Accessibility / Input Monitoring permissions stay granted. Without it the app is ad-hoc
# signed and the permissions must be toggled again after each update.

set -u
setopt null_glob
LABEL="com.pptk.keys"
ROOT="$HOME/Library/Application Support/PPTK"
APP="$ROOT/PPTKKeys.app"
EXE="$APP/Contents/MacOS/PPTKKeys"
BACKUPS="$ROOT/backups"
PLIST="$HOME/Library/LaunchAgents/$LABEL.plist"
LOGDIR="$HOME/Library/Logs/PPTK"
LOG="$LOGDIR/PPTKKeys.log"
DOMAIN="gui/$(id -u)"
REPO_BUILD="$HOME/Documents/GitHub/PPTK/src/PPTKKeys/mac/.build/release/PPTKKeys"
IDENTITY_NAME="PPTK Local Code Signing"
STAMP=$(date +%Y%m%d-%H%M%S)

loaded() { launchctl print "$DOMAIN/$LABEL" >/dev/null 2>&1; }
helper_pids() { pgrep -f "PPTKKeys.app/Contents/MacOS/PPTKKeys|release/PPTKKeys" 2>/dev/null; }

rotate_log() {
  mkdir -p "$LOGDIR"
  if [ -f "$LOG" ] && [ $(stat -f %z "$LOG") -gt 5000000 ]; then mv -f "$LOG" "$LOG.1"; fi
}

write_plist() {
  mkdir -p "$HOME/Library/LaunchAgents"
  cat > "$PLIST" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key><string>$LABEL</string>
  <key>ProgramArguments</key>
  <array><string>$EXE</string></array>
  <key>RunAtLoad</key><true/>
  <key>KeepAlive</key>
  <dict><key>SuccessfulExit</key><false/></dict>
  <key>ThrottleInterval</key><integer>30</integer>
  <key>ProcessType</key><string>Interactive</string>
  <key>LimitLoadToSessionType</key><string>Aqua</string>
  <key>StandardOutPath</key><string>$LOG</string>
  <key>StandardErrorPath</key><string>$LOG</string>
</dict>
</plist>
EOF
  plutil -lint "$PLIST" >/dev/null || { echo "invalid plist"; exit 1; }
}

write_bundle() {
  local src="$1"
  mkdir -p "$APP/Contents/MacOS"
  cp -f "$src" "$EXE"
  chmod 755 "$EXE"
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
  <key>CFBundleShortVersionString</key><string>2.0</string>
  <key>CFBundleVersion</key><string>$STAMP</string>
  <key>LSUIElement</key><true/>
  <key>NSAppleEventsUsageDescription</key><string>PPTK Keys asks PowerPoint for the current selection and runs the PPTK add-in commands for your keyboard shortcuts.</string>
</dict>
</plist>
EOF
  local id
  id=$(security find-identity -p codesigning 2>/dev/null | grep "\"$IDENTITY_NAME\"" | head -1 | awk '{print $2}')
  if [ -n "$id" ]; then
    codesign --force --sign "$id" --identifier "$LABEL" "$APP" || { echo "codesign failed"; exit 1; }
  else
    echo "warning: '$IDENTITY_NAME' not found - ad-hoc signature (permissions must be re-granted after updates)"
    codesign --force --sign - --identifier "$LABEL" "$APP" || exit 1
  fi
  codesign --verify --strict "$APP" 2>/dev/null || echo "note: signature verifies only as untrusted self-signed (expected)"
}

backup_existing() {
  mkdir -p "$BACKUPS"
  [ -d "$APP" ] && cp -Rp "$APP" "$BACKUPS/PPTKKeys.app.$STAMP" && echo "backed up app -> $BACKUPS/PPTKKeys.app.$STAMP"
  [ -f "$PLIST" ] && cp -p "$PLIST" "$BACKUPS/$LABEL.plist.$STAMP" && echo "backed up agent -> $BACKUPS/$LABEL.plist.$STAMP"
  # keep the 5 newest backups of each kind
  ls -1dt "$BACKUPS"/PPTKKeys.app.* 2>/dev/null | tail -n +6 | while read -r old; do rm -rf "$old"; done
  ls -1t "$BACKUPS"/$LABEL.plist.* 2>/dev/null | tail -n +6 | while read -r old; do rm -f "$old"; done
}

start() {
  rotate_log
  if loaded; then launchctl kickstart "$DOMAIN/$LABEL"; else launchctl bootstrap "$DOMAIN" "$PLIST"; fi
}

stop() {
  if loaded; then launchctl bootout "$DOMAIN/$LABEL" 2>/dev/null; fi
  for i in 1 2 3 4 5 6 7 8 9 10; do pgrep -f "PPTKKeys.app/Contents/MacOS/PPTKKeys" >/dev/null || break; sleep 0.3; done
}

cmd="${1:-status}"
case "$cmd" in
  install|update)
    src="${2:-$REPO_BUILD}"
    [ -x "$src" ] || { echo "binary not found: $src (build with: swift build -c release)"; exit 1; }
    stop
    backup_existing
    write_bundle "$src"
    write_plist
    mkdir -p "$ROOT"
    [ "$0" -ef "$ROOT/pptkkeys-ctl" ] || cp -f "$0" "$ROOT/pptkkeys-ctl"
    chmod 755 "$ROOT/pptkkeys-ctl"
    start
    sleep 2
    "$0" status
    ;;
  start) start; sleep 1; "$0" status ;;
  stop) stop; echo "stopped (it starts again at next login; use uninstall to remove it)" ;;
  restart) rotate_log; if loaded; then launchctl kickstart -k "$DOMAIN/$LABEL"; else start; fi; sleep 1; "$0" status ;;
  status)
    if loaded; then
      launchctl print "$DOMAIN/$LABEL" 2>/dev/null | grep -E '^\s*(state|pid|last exit code|runs) =' | sed 's/^\s*/  /'
    else
      echo "  agent not loaded"
    fi
    echo "  helper processes: $(helper_pids | wc -l | tr -d ' ')"
    [ -f "$LOG" ] && tail -n 4 "$LOG" | sed 's/^/  log: /'
    ;;
  logs) tail -n "${2:-40}" "$LOG" ;;
  uninstall)
    stop
    backup_existing
    rm -f "$PLIST"
    rm -rf "$APP"
    echo "uninstalled. Backups: $BACKUPS. To remove completely: rm -rf \"$ROOT\" \"$LOGDIR\";"
    echo "then remove 'PPTK Keys' from Privacy & Security > Accessibility and Input Monitoring, and optionally"
    echo "delete the '$IDENTITY_NAME' certificate in Keychain Access."
    ;;
  *) echo "usage: pptkkeys-ctl install|update [binary] | start | stop | restart | status | logs [n] | uninstall"; exit 1 ;;
esac
