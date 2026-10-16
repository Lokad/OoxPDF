"""One public shared-content control for body and simple header/footer cells."""
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib
import json

root = Path(__file__).resolve().parent.parent
cases = root / "tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
r = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
pkg = "{http://schemas.openxmlformats.org/package/2006/relationships}"
ct = "{http://schemas.openxmlformats.org/package/2006/content-types}"
for prefix, namespace in [("w", w), ("m", m), ("r", r)]:
    E.register_namespace(prefix, namespace[1:-1])
source = cases / "math-authored-inline-whitespace.docx"
assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == "F931E0F03C16724018E93D52F577FD67C6BB250A194280FC9C17768B9EDDBC38"
with ZipFile(source) as archive:
    entries = {name: archive.read(name) for name in archive.namelist()}
image_source = cases / "math-printed-frame-fraction-anchor12-cell-flow-leading-image.docx"
assert hashlib.sha256(image_source.read_bytes()).hexdigest().upper() == "29782EEC02950009C229EAA7E7471B24A141955EA3B90FAFA283D0AABEA10BF3"
with ZipFile(image_source) as archive:
    image_document = E.fromstring(archive.read("word/document.xml"))
    image_run = next(run for run in image_document.iter(w + "r") if run.find(w + "drawing") is not None)
    image_bytes = archive.read("word/media/leading-image.png")

document = E.fromstring(entries["word/document.xml"])
body = document.find(w + "body")
section = deepcopy(body.find(w + "sectPr"))
for reference in list(section):
    if reference.tag in (w + "headerReference", w + "footerReference"):
        section.remove(reference)
paragraph_source = deepcopy(body.find(w + "p"))


def paragraph(owner, static=False):
    result = deepcopy(paragraph_source)
    next(result.iter(w + "t")).text = owner + ": "
    picture = deepcopy(image_run)
    result.insert(list(result).index(result.findall(m + "oMath")[1]), picture)
    text = E.SubElement(E.SubElement(result, w + "r"), w + "t")
    text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text.text = " Page "
    for kind in (["PAGE", "NUMPAGES"] if static else ["PAGE"]):
        if kind == "NUMPAGES":
            text = E.SubElement(E.SubElement(result, w + "r"), w + "t")
            text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            text.text = " of "
        field = E.SubElement(result, w + "fldSimple", {w + "instr": kind})
        E.SubElement(E.SubElement(field, w + "r"), w + "t").text = "1"
    return result


def table(content):
    result = E.Element(w + "tbl")
    props = E.SubElement(result, w + "tblPr")
    E.SubElement(props, w + "tblW", {w + "w": "9360", w + "type": "dxa"})
    E.SubElement(props, w + "tblLayout", {w + "type": "fixed"})
    E.SubElement(E.SubElement(result, w + "tblGrid"), w + "gridCol", {w + "w": "9360"})
    cell = E.SubElement(E.SubElement(result, w + "tr"), w + "tc")
    E.SubElement(E.SubElement(cell, w + "tcPr"), w + "tcW", {w + "w": "9360", w + "type": "dxa"})
    cell.append(content)
    return result


body.clear()
body.append(paragraph("Body"))
body.append(table(paragraph("Cell")))
body.append(E.Element(w + "p"))
relationships = E.fromstring(entries["word/_rels/document.xml.rels"])
types = E.fromstring(entries["[Content_Types].xml"])
for relationship in list(relationships):
    if relationship.get("Type", "").endswith(("/header", "/footer")):
        relationships.remove(relationship)
        part = "word/" + relationship.get("Target")
        entries.pop(part, None)
        for content_type in list(types):
            if content_type.get("PartName") == "/" + part:
                types.remove(content_type)
entries["word/media/leading-image.png"] = image_bytes
image_relationship = {"Id": "leadingImage", "Type": r[1:-1] + "/image", "Target": "media/leading-image.png"}
E.SubElement(relationships, pkg + "Relationship", image_relationship)
if not any(x.get("Extension") == "png" for x in types):
    E.SubElement(types, ct + "Default", {"Extension": "png", "ContentType": "image/png"})
for owner in ["header", "footer"]:
    name = "static-cell-" + owner + ".xml"
    story = E.Element(w + ("hdr" if owner == "header" else "ftr"))
    story.append(table(paragraph(owner.title(), static=True)))
    story.append(E.Element(w + "p"))
    entries["word/" + name] = E.tostring(story, encoding="utf-8", xml_declaration=True)
    rels = E.Element(pkg + "Relationships")
    E.SubElement(rels, pkg + "Relationship", image_relationship)
    entries["word/_rels/" + name + ".rels"] = E.tostring(rels, encoding="utf-8", xml_declaration=True)
    E.SubElement(relationships, pkg + "Relationship", {"Id": "StaticCell" + owner, "Type": r[1:-1] + "/" + owner, "Target": name})
    E.SubElement(section, w + owner + "Reference", {w + "type": "default", r + "id": "StaticCell" + owner})
    E.SubElement(types, ct + "Override", {"PartName": "/word/" + name, "ContentType": "application/vnd.openxmlformats-officedocument.wordprocessingml." + owner + "+xml"})
body.append(section)
entries["word/document.xml"] = E.tostring(document, encoding="utf-8", xml_declaration=True)
entries["word/_rels/document.xml.rels"] = E.tostring(relationships, encoding="utf-8", xml_declaration=True)
entries["[Content_Types].xml"] = E.tostring(types, encoding="utf-8", xml_declaration=True)
target = cases / "math-static-cell-composition.docx"
with ZipFile(target, "w", ZIP_DEFLATED) as archive:
    for name, data in sorted(entries.items()):
        info = ZipInfo(name, (2026, 10, 16, 0, 0, 0))
        info.compress_type = ZIP_DEFLATED
        archive.writestr(info, data)
manifest = [dict(Id=target.stem, Input=target.relative_to(root).as_posix(), Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
    HeldOut=False, ExpectedEquationCount=8, ExpectedInlineImageCount=4, ExpectedText=["Body:", "Cell:", "Header:", "Footer:"])]
for part in ["word/document.xml", "word/static-cell-header.xml", "word/static-cell-footer.xml"]:
    tree = E.fromstring(entries[part])
    for parent in tree.iter():
        for child in list(parent):
            if child.tag in (m + "oMath", w + "drawing"):
                parent.remove(child)
    entries[part] = E.tostring(tree, encoding="utf-8", xml_declaration=True)
plain = cases / "static-cell-plain-frame.docx"
with ZipFile(plain, "w", ZIP_DEFLATED) as archive:
    for name, data in sorted(entries.items()):
        info = ZipInfo(name, (2026, 10, 16, 0, 0, 0))
        info.compress_type = ZIP_DEFLATED
        archive.writestr(info, data)
manifest.append(dict(Id=plain.stem, Input=plain.relative_to(root).as_posix(), Sha256=hashlib.sha256(plain.read_bytes()).hexdigest().upper(),
    HeldOut=False, ExpectedEquationCount=0, ExpectedInlineImageCount=0, ExpectedText=["Body:", "Cell:", "Header:", "Footer:"]))
target.with_suffix(".json").write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
print("Pinned shared body/cell/header/footer equations, image and page-field content.")
