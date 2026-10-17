"""Supported custom note marks beside one or two normal-font fractions."""
from pathlib import Path
from copy import deepcopy
from xml.etree import ElementTree as E
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
import hashlib
import json

root = Path(__file__).resolve().parent.parent
cases = root / "tests/Lokad.OoxPdf.Tests/Cases"
output = cases / "CompositionProbes"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
E.register_namespace("w", w[1:-1])
E.register_namespace("m", m[1:-1])
with ZipFile(output / "math-footnote-before-body-plain.docx") as z:
    equation = next(E.fromstring(z.read("word/document.xml")).iter(m + "oMath"))
rows = []
for kind in ["footnote", "endnote"]:
    source = cases / ("note-custom-mark-" + kind + ".docx")
    for count in [1, 2]:
        with ZipFile(source) as z:
            entries = {name: z.read(name) for name in z.namelist()}
        xml = E.fromstring(entries["word/document.xml"])
        paragraph = next(p for p in xml.iter(w + "p") if any(r.get(w + "customMarkFollows") == "1" for r in p.iter(w + kind + "Reference")))
        marker = next(run for run in paragraph if any(r.get(w + "customMarkFollows") == "1" for r in run.iter(w + kind + "Reference")))
        position = list(paragraph).index(marker)
        paragraph.insert(position, E.Element(w + "bookmarkStart", {w + "id": "81", w + "name": "CustomNoteEquation"}))
        paragraph.insert(position + 1, deepcopy(equation))
        if count == 2:
            separator = E.Element(w + "r")
            E.SubElement(separator, w + "t", {"{http://www.w3.org/XML/1998/namespace}space": "preserve"}).text = " and "
            paragraph.insert(position + 3, separator)
            paragraph.insert(position + 4, deepcopy(equation))
        paragraph.insert(position + (5 if count == 2 else 3), E.Element(w + "bookmarkEnd", {w + "id": "81"}))
        entries["word/document.xml"] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
        target = output / ("math-custom-note-" + kind + "-" + str(count) + ".docx")
        with ZipFile(target, "w", ZIP_DEFLATED) as z:
            for name, data in sorted(entries.items()):
                info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                z.writestr(info, data)
        rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
            Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
            SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
            Kind=kind, HeldOut=False, ExpectedEquationCount=count,
            ExpectedText=["Public destination Alpha.", "Public destination Delta."]))
(output / "math-custom-notes.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Pinned four custom-marker/math source controls")
