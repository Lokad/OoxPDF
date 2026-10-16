"""Public authored-whitespace controls shared by body, cell and header math."""
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib, json

root = Path(__file__).resolve().parent.parent
cases = root / "tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes"
source = cases / "math-printed-frame-fraction-anchor12-multiple.docx"
assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == "EA7FB764C259DEA27489D7BBDA18BE64CC1AC37C6C091656715B6FC7897C0AFC"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
r = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
pkg = "{http://schemas.openxmlformats.org/package/2006/relationships}"
ct = "{http://schemas.openxmlformats.org/package/2006/content-types}"
for prefix, namespace in [("w", w), ("m", m), ("r", r)]:
    E.register_namespace(prefix, namespace[1:-1])
with ZipFile(source) as archive:
    entries = {name: archive.read(name) for name in archive.namelist()}
document = E.fromstring(entries["word/document.xml"])
body = document.find(w + "body")
section = deepcopy(body.find(w + "sectPr"))
source_paragraph = next(p for p in body.findall(w + "p") if len(list(p.iter(m + "oMath"))) == 2)

def paragraphs(container):
    result = []
    for kind in ["Leading", "Hard", "Trailing"]:
        paragraph = deepcopy(source_paragraph)
        first = next(paragraph.iter(w + "t"))
        first.text = ("  " if kind == "Leading" else "") + container + " " + kind + ": " + first.text
        first.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        if kind == "Hard":
            middle = next(run for run in paragraph.findall(w + "r") if run.find(w + "t") is not None and run.find(w + "t").text == " and ")
            middle.insert(0, E.Element(w + "br"))
        if kind == "Trailing":
            run = E.SubElement(paragraph, w + "r")
            E.SubElement(run, w + "br")
            E.SubElement(run, w + "t", {"{http://www.w3.org/XML/1998/namespace}space": "preserve"}).text = "  "
        result.append(paragraph)
    return result

body.clear()
body.extend(paragraphs("Body"))
table = E.SubElement(body, w + "tbl")
props = E.SubElement(table, w + "tblPr")
E.SubElement(props, w + "tblW", {w + "w": "9360", w + "type": "dxa"})
E.SubElement(props, w + "tblLayout", {w + "type": "fixed"})
E.SubElement(E.SubElement(table, w + "tblGrid"), w + "gridCol", {w + "w": "9360"})
cell = E.SubElement(E.SubElement(table, w + "tr"), w + "tc")
E.SubElement(E.SubElement(cell, w + "tcPr"), w + "tcW", {w + "w": "9360", w + "type": "dxa"})
cell.extend(paragraphs("Cell"))
E.SubElement(E.SubElement(body, w + "p"), w + "r")
header = E.Element(w + "hdr")
header.extend(paragraphs("Header"))
E.SubElement(section, w + "headerReference", {w + "type": "default", r + "id": "InlineWhitespaceHeader"})
body.append(section)
entries["word/document.xml"] = E.tostring(document, encoding="utf-8", xml_declaration=True)
entries["word/inline-whitespace-header.xml"] = E.tostring(header, encoding="utf-8", xml_declaration=True)
relationships = E.fromstring(entries["word/_rels/document.xml.rels"])
E.SubElement(relationships, pkg + "Relationship", {"Id": "InlineWhitespaceHeader", "Type": r[1:-1] + "/header", "Target": "inline-whitespace-header.xml"})
entries["word/_rels/document.xml.rels"] = E.tostring(relationships, encoding="utf-8", xml_declaration=True)
types = E.fromstring(entries["[Content_Types].xml"])
E.SubElement(types, ct + "Override", {"PartName": "/word/inline-whitespace-header.xml", "ContentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"})
entries["[Content_Types].xml"] = E.tostring(types, encoding="utf-8", xml_declaration=True)
target = cases / "math-authored-inline-whitespace.docx"
with ZipFile(target, "w", ZIP_DEFLATED) as archive:
    for name, data in sorted(entries.items()):
        info = ZipInfo(name, (2026, 10, 16, 0, 0, 0)); info.compress_type = ZIP_DEFLATED
        archive.writestr(info, data)
manifest = [dict(Id=target.stem, Input=target.relative_to(root).as_posix(), Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
    HeldOut=False, ExpectedEquationCount=18, ExpectedInlineImageCount=0,
    ExpectedText=[container + " " + kind + ":" for container in ["Body", "Cell", "Header"] for kind in ["Leading", "Hard", "Trailing"]])]
target.with_suffix(".json").write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
print("Pinned one shared body/cell/header whitespace control: leading, authored break and trailing-break spaces.")
