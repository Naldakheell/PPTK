#!/usr/bin/env python3
"""Mac version of the PPTK Windows v2 ribbon (customUI.xml).

PowerPoint for Mac never collapses add-in ribbon groups (it scrolls them; tested 2026-11-22, autoScale has no
effect), while Windows Office collapses Rows / TK, Tools and Shape into single drop-down buttons at laptop widths.
This makes those three groups single large drop-down buttons on Mac, so the PPTK tab fits without scrolling, and
gives Replace Fonts a PPTK icon (imageMso "FontsReplaceFonts" does not exist in Office for Mac).

Usage: make_mac_ribbon.py <windows customUI.xml> <out customUI.xml>
Everything else in the XML is unchanged. The new image id is "pptk-replace-fonts" (customUI/images, see ribbon_swap).
"""
import re
import sys
import xml.etree.ElementTree as ET

NS = "http://schemas.microsoft.com/office/2009/07/customui"
ET.register_namespace("", NS)
Q = lambda tag: "{%s}%s" % (NS, tag)

# group id -> (drop-down label, image, element ids left out). A ribbon menu cannot hold the width / height
# spin boxes; on Mac they are always on PowerPoint's own Shape Format tab, so they are left out here.
COLLAPSE = {
    "PPTKGrid": ("Rows / TK", "sb-row-add-keep-large", set()),
    "Tools": ("Tools", "pptk-smart-bar-large", set()),
    "Shape": ("Shape", "pptk-change-shape-large", {"ShapeWidthBox", "ShapeHeightBox"}),
}
NOT_IN_MENU = ("size", "showLabel")   # ribbon-only attributes; inside a menu every item shows its label


def main(src, dst):
    tree = ET.parse(src)
    root = tree.getroot()
    found = set()
    for group in root.iter(Q("group")):
        gid = group.get("id")
        if gid not in COLLAPSE:
            continue
        label, image, drop = COLLAPSE[gid]
        children = list(group)
        menu = ET.Element(Q("menu"), {"id": gid + "MacMenu", "label": label, "image": image,
                                      "size": "large", "itemSize": "normal",
                                      "screentip": label, "supertip": "PPTK " + label + " commands"})
        for child in children:
            group.remove(child)
            if child.get("id") in drop:
                continue
            for attr in NOT_IN_MENU:
                child.attrib.pop(attr, None)
            if child.get("image", "").endswith("-large"):
                child.set("image", child.get("image")[:-len("-large")])   # menu items use the 16 px icon
            menu.append(child)
        group.append(menu)
        found.add(gid)
    missing = set(COLLAPSE) - found
    if missing:
        raise SystemExit("groups not found: %s" % sorted(missing))
    replaced = 0
    for button in root.iter(Q("button")):
        if button.get("imageMso") == "FontsReplaceFonts":
            del button.attrib["imageMso"]
            button.set("image", "pptk-replace-fonts")
            replaced += 1
    if replaced != 1:
        raise SystemExit("Replace Fonts button not found exactly once (%d)" % replaced)
    xml = ET.tostring(root, encoding="unicode")
    xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml
    open(dst, "w", encoding="utf-8").write(xml)
    print("Mac ribbon: collapsed", sorted(found), "+ Replace Fonts icon")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
