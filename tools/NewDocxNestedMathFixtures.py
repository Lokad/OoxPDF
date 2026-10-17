"""Ordinary nested math across body depths, static stories and page notes."""
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
import hashlib
import json
import xml.etree.ElementTree as E

root = Path(__file__).resolve().parent.parent
output = root / "tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
r = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
rel = "{http://schemas.openxmlformats.org/package/2006/relationships}"
ct = "{http://schemas.openxmlformats.org/package/2006/content-types}"

def load(path):
    with ZipFile(path) as z:
        return {name: z.read(name) for name in z.namelist()}

def prose(text):
    return E.fromstring('<w:p xmlns:w="' + w[1:-1] + '"><w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr><w:t>' + text + '</w:t></w:r></w:p>')

def write(id, entries):
    target = output / (id + ".docx")
    with ZipFile(target, "w", ZIP_DEFLATED) as z:
        for name, payload in sorted(entries.items()):
            info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, payload)
    return target

base = output / "mixed-nested-cell-inside.docx"
original = load(base)
rows = []
def record(target, container, equations, images, expected):
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(), Container=container,
        HeldOut=False, ExpectedEquationCount=equations, ExpectedImagePlacements=images,
        ExpectedText=expected))

record(base, "body", 2, 1, ["Public outer head", "Public outer tail"])
entries = dict(original)
xml = E.fromstring(entries["word/document.xml"])
inner = xml.find(w + "body/" + w + "tbl/" + w + "tr/" + w + "tc/" + w + "tbl")
cell = inner.find(w + "tr/" + w + "tc")
mixed = cell.find(w + "p")
nested = deepcopy(inner)
cell.remove(mixed)
cell.append(nested)
cell.append(prose("Public middle tail"))
entries["word/document.xml"] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
record(write("mixed-nested-math-deep-body", entries), "body", 2, 1,
       ["Public outer head", "Public middle tail", "Public outer tail"])

entries = dict(original)
xml = E.fromstring(entries["word/document.xml"])
body = xml.find(w + "body")
table = body.find(w + "tbl")
body.remove(table)
body.insert(0, prose("Public body anchor"))
section = body.find(w + "sectPr")
margins = section.find(w + "pgMar")
for name, value in [("top", "2880"), ("bottom", "2880"), ("header", "360"), ("footer", "360")]:
    margins.set(w + name, value)
relationships = E.fromstring(entries["word/_rels/document.xml.rels"])
types = E.fromstring(entries["[Content_Types].xml"])
for kind in ["header", "footer"]:
    part = "nested-math-" + kind + ".xml"
    story = E.Element(w + ("hdr" if kind == "header" else "ftr"))
    story.append(deepcopy(table))
    for properties in story.iter("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr"):
        properties.set("id", "9001" if kind == "header" else "9002")
    entries["word/" + part] = E.tostring(story, encoding="utf-8", xml_declaration=True)
    image_relationships = E.Element(rel + "Relationships")
    for relationship in relationships:
        if relationship.get("Type", "").endswith("/image"):
            image_relationships.append(deepcopy(relationship))
    entries["word/_rels/" + part + ".rels"] = E.tostring(image_relationships, encoding="utf-8", xml_declaration=True)
    E.SubElement(relationships, rel + "Relationship", {"Id": "nested" + kind,
        "Type": r[1:-1] + "/" + kind, "Target": part})
    E.SubElement(types, ct + "Override", {"PartName": "/word/" + part,
        "ContentType": "application/vnd.openxmlformats-officedocument.wordprocessingml." + kind + "+xml"})
    section.insert(0, E.Element(w + kind + "Reference", {w + "type": "default", r + "id": "nested" + kind}))
entries["word/document.xml"] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
entries["word/_rels/document.xml.rels"] = E.tostring(relationships, encoding="utf-8", xml_declaration=True)
entries["[Content_Types].xml"] = E.tostring(types, encoding="utf-8", xml_declaration=True)
record(write("mixed-nested-math-static", entries), "static", 4, 2,
       ["Public body anchor", "Public outer head", "Public outer tail"])

for kind in ["footnote", "endnote"]:
    entries = load(output / ("mixed-note-" + kind + "-table.docx"))
    part = "word/" + kind + "s.xml"
    xml = E.fromstring(entries[part])
    story = next(item for item in xml if item.get(w + "id") == "37")
    table = story.find(w + "tbl")
    nested = deepcopy(table)
    cell = table.find(w + "tr/" + w + "tc")
    for paragraph in cell.findall(w + "p"):
        cell.remove(paragraph)
    cell.append(nested)
    cell.append(prose("Public note outer tail"))
    entries[part] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
    record(write("mixed-nested-math-" + kind, entries), kind, 2, 1,
           ["Public note calculation", "Public note outer tail", "Public destination Alpha.", "Public destination Delta."])

(output / "mixed-nested-math.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Pinned five nested-math controls across two body depths and four story containers")
