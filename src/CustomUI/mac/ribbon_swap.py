#!/usr/bin/env python3
"""Replace the ribbon of a .pptm with the ribbon of another .pptm (PPTK Mac: Windows v2 ribbon).

Usage: ribbon_swap.py <target.pptm> <source.pptm>
Only customUI/* members, the ui/extensibility relationship in _rels/.rels and (if needed) the png
content type change. Every other zip member is checked byte-identical before the target is replaced.
"""
import hashlib, os, re, sys, zipfile

target, source = sys.argv[1], sys.argv[2]
options = set(a for a in sys.argv[3:] if "=" not in a)
kv = dict(a.split("=", 1) for a in sys.argv[3:] if "=" in a)
# macxml=<file>: use this customUI.xml instead of the source's (Mac ribbon from make_mac_ribbon.py)
# image=<id>:<png>: add customUI/images/<id>.png and its relationship (repeatable as image2=, image3=)
extra_images = {}
for key, value in kv.items():
    if key.startswith("image"):
        iid, path = value.split(":", 1)
        extra_images[iid] = open(path, "rb").read()   # nocontext: drop <contextMenus> (Office for Mac has no ribbon context menus); rel2006

tz, sz = zipfile.ZipFile(target), zipfile.ZipFile(source)
src_ui = [i for i in sz.infolist() if i.filename.startswith("customUI/")]
old_ui = {i.filename for i in tz.infolist() if i.filename.startswith("customUI/")}
rel_re = re.compile(r'<Relationship [^>]*Type="http://schemas\.microsoft\.com/office/20\d\d/relationships/ui/extensibility"[^>]*/>')
src_rel = rel_re.findall(sz.read("_rels/.rels").decode("utf-8"))
assert len(src_rel) == 1, "source extensibility relationship"
rels = tz.read("_rels/.rels").decode("utf-8")
assert len(rel_re.findall(rels)) == 1, "target extensibility relationship"
new_rels = rel_re.sub(src_rel[0], rels).encode("utf-8")
ct = tz.read("[Content_Types].xml").decode("utf-8")
if 'Extension="png"' not in ct:
    ct = ct.replace("<Default ", '<Default Extension="png" ContentType="image/png"/><Default ', 1)
new_ct = ct.encode("utf-8")
added_ui = {"customUI/images/%s.png" % iid: png for iid, png in extra_images.items()}
for name in added_ui:
    src_ui.append(zipfile.ZipInfo(name, (2026, 11, 22, 0, 0, 0)))

def ui_data(name):
    if name in added_ui:
        return added_ui[name]
    data = sz.read(name)
    if name == "customUI/customUI.xml" and "macxml" in kv:
        data = open(kv["macxml"], "rb").read()
    if name == "customUI/_rels/customUI.xml.rels" and extra_images:
        text = data.decode("utf-8-sig")
        for iid in extra_images:
            rel = '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="images/%s.png"/>' % (iid, iid)
            assert ('Id="%s"' % iid) not in text
            text = text.replace("</Relationships>", "  " + rel + "\n</Relationships>")
        data = text.encode("utf-8")
    if name == "customUI/customUI.xml" and "nocontext" in options:
        text = data.decode("utf-8")
        s = text.find("<contextMenus>")
        e = text.find("</contextMenus>")
        assert s > 0 and e > s
        text = text[:s] + text[e + len("</contextMenus>"):]
        data = text.encode("utf-8")
    return data

if "rel2006" in options:
    src_rel[0] = src_rel[0].replace("/office/2007/relationships/ui/extensibility", "/office/2006/relationships/ui/extensibility")
    new_rels = rel_re.sub(src_rel[0], rels).encode("utf-8")
changed = {"_rels/.rels", "[Content_Types].xml"} | old_ui
before = {i.filename: hashlib.sha256(tz.read(i.filename)).hexdigest() for i in tz.infolist() if i.filename not in changed}
tmp = target + ".ribbon-tmp"
with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
    for info in tz.infolist():
        if info.filename in old_ui:
            continue
        data = new_rels if info.filename == "_rels/.rels" else new_ct if info.filename == "[Content_Types].xml" else tz.read(info.filename)
        zi = zipfile.ZipInfo(info.filename, info.date_time)
        zi.compress_type, zi.external_attr = info.compress_type, info.external_attr
        zout.writestr(zi, data)
    for info in src_ui:
        zi = zipfile.ZipInfo(info.filename, info.date_time)
        zi.compress_type = zipfile.ZIP_DEFLATED
        zout.writestr(zi, ui_data(info.filename))
nz = zipfile.ZipFile(tmp)
assert nz.testzip() is None
for name, digest in before.items():
    if hashlib.sha256(nz.read(name)).hexdigest() != digest:
        os.remove(tmp); raise SystemExit("member changed: " + name)
for info in src_ui:
    assert nz.read(info.filename) == ui_data(info.filename)
os.replace(tmp, target)
print("OK ribbon replaced:", len(old_ui), "old customUI members ->", len(src_ui), "new")
