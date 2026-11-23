#!/bin/bash
# PPTK for Mac __PPTK_VERSION__ - uninstall (needs an administrator password).
cd "$(dirname "$0")" || exit 1
echo "This removes PPTK for Mac (helper, add-in, LaunchAgent) for every user of this Mac."
read -r -p "Type YES to continue: " answer
[ "$answer" = "YES" ] || { echo "Cancelled."; exit 0; }
if [ -x "/Library/Application Support/PPTK/pptk-uninstall.sh" ]; then
    sudo "/Library/Application Support/PPTK/pptk-uninstall.sh"
else
    sudo "./pptk-uninstall.sh"
fi
echo "If PowerPoint still lists PPTK under Tools > PowerPoint Add-ins, select it and click the minus button."
read -r -p "Press Return to close."
