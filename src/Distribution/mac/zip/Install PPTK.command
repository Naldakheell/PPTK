#!/bin/bash
# PPTK for Mac __PPTK_VERSION__ - manual installation for IT staff.
# Installs the same PKG Intune deploys (same files, same receipt), so Intune can update this Mac later.
# Needs an administrator password. Nothing is downloaded and no security setting is changed.
cd "$(dirname "$0")" || exit 1
PKG="PPTK-__PPTK_VERSION__.pkg"
echo "PPTK for Mac __PPTK_VERSION__"
echo

# 1. Package integrity (SHA256SUMS.txt from the build)
expected=$(awk -v f="$PKG" '$2 == f {print $1}' SHA256SUMS.txt 2>/dev/null)
actual=$(shasum -a 256 "$PKG" 2>/dev/null | awk '{print $1}')
if [ -z "$expected" ] || [ "$expected" != "$actual" ]; then
    echo "The package does not match SHA256SUMS.txt. Use a fresh copy of the ZIP."; read -r -p "Press Return to close."; exit 1
fi

# 2. Requirements
major=$(sw_vers -productVersion | cut -d. -f1)
if [ "$major" -lt 13 ]; then echo "macOS 13 or later is required."; read -r -p "Press Return to close."; exit 1; fi
[ -d "/Applications/Microsoft PowerPoint.app" ] || echo "Note: Microsoft PowerPoint is not in /Applications. PPTK needs Microsoft 365 PowerPoint."

# 3. Signature status (informational; an unsigned test build says so)
echo "Package signature:"
pkgutil --check-signature "$PKG" | sed -n '1,3p'
echo

# 4. Install (the installer runs the package's own preinstall/postinstall)
echo "Enter an administrator password to install."
if ! sudo /usr/sbin/installer -pkg "$PKG" -target /; then
    echo "Installation failed. A newer PPTK may already be installed (downgrades are refused)."
    read -r -p "Press Return to close."; exit 1
fi
echo
sudo "./pptk-detect.sh" --machine || true

# 5. Permissions the user approves once (macOS asks; they cannot be set by this installer)
cat <<'EOF'

Next steps for the employee (once):
  1. System Settings > Privacy & Security > Accessibility: turn on "PPTK Keys".
  2. System Settings > Privacy & Security > Input Monitoring: turn on "PPTK Keys".
  3. Quit PowerPoint (Command-Q) and open it again: the PPTK tab appears.
     When macOS asks whether "PPTK Keys" may control "Microsoft PowerPoint", click OK.
On Intune-managed Macs the PPTK privacy profile can pre-approve steps 1 and 3 (Input Monitoring stays manual).
EOF
open "x-apple.systempreferences:com.apple.preference.security?Privacy_Accessibility" 2>/dev/null
read -r -p "Press Return to close."
