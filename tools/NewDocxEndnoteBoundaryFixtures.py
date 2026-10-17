"""Isolate inline image and paragraph ownership in complete endnote blocks."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
import hashlib
import json
import xml.etree.ElementTree as E

root = Path(__file__).resolve().parent.parent
output = root / "tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes"
source = output / "mixed-note-endnote-paragraph.docx"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
E.register_namespace("w", w[1:-1])
with ZipFile(source) as z:
    original = {name: z.read(name) for name in z.namelist()}
rows = []
for variant in ["no-image", "one-paragraph"]:
    entries = dict(original)
    xml = E.fromstring(entries["word/endnotes.xml"])
    note = next(item for item in xml if item.get(w + "id") == "37")
    first, second = note.findall(w + "p")
    if variant == "no-image":
        for run in second.findall(w + "r"):
            if run.find(w + "drawing") is not None:
                second.remove(run)
    else:
        space = E.SubElement(first, w + "r")
        E.SubElement(space, w + "t", {"{http://www.w3.org/XML/1998/namespace}space": "preserve"}).text = " "
        for child in list(second):
            if child.tag != w + "pPr":
                first.append(child)
        note.remove(second)
    entries["word/endnotes.xml"] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
    target = output / ("endnote-boundary-" + variant + ".docx")
    with ZipFile(target, "w", ZIP_DEFLATED) as z:
        for name, payload in sorted(entries.items()):
            info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, payload)
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        HeldOut=False, ExpectedEquationCount=2,
        ExpectedImagePlacements=0 if variant == "no-image" else 1,
        ExpectedText=["Public destination Alpha.", "Public destination Delta.", "Public note calculation"]))
plain = root / "tests/Lokad.OoxPdf.Tests/Cases/note-custom-mark-endnote.docx"
rows.append(dict(Id="endnote-placement-plain-custom", Input=plain.relative_to(root).as_posix(),
    Sha256=hashlib.sha256(plain.read_bytes()).hexdigest().upper(), HeldOut=False,
    ExpectedEquationCount=0, ExpectedImagePlacements=0,
    ExpectedText=["Public destination Alpha.", "Public destination Delta."]))
(output / "endnote-boundary-controls.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Pinned two endnote transformations and the existing plain source control")
