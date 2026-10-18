"""Exercise content-driven minimum slots for inline/display math in containers."""
from argparse import ArgumentParser
from copy import deepcopy
import hashlib
import json
from pathlib import Path
from xml.etree import ElementTree as E
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED

root = Path(__file__).resolve().parent.parent
parser = ArgumentParser()
parser.add_argument("--output-directory", type=Path,
                    default=root / "tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes")
output = parser.parse_args().output_directory.resolve()
output.mkdir(parents=True, exist_ok=True)
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
r = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
rel = "{http://schemas.openxmlformats.org/package/2006/relationships}"
ct = "{http://schemas.openxmlformats.org/package/2006/content-types}"
source = root / "tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/multiple-integral-double.docx"
with ZipFile(source) as archive:
    original = {path: archive.read(path) for path in archive.namelist()}
rows = []
for container, minimum in (("body", 12), ("body", 72), ("cell", 72), ("header", 72)):
    parts = dict(original)
    document = E.fromstring(parts["word/document.xml"])
    body = document.find(w + "body")
    for paragraph in body:
        if paragraph.tag != w + "p" or not list(paragraph.iter(m + "oMath")):
            continue
        properties = paragraph.find(w + "pPr")
        if properties is None:
            properties = E.Element(w + "pPr")
            paragraph.insert(0, properties)
        spacing = properties.find(w + "spacing")
        if spacing is None:
            spacing = E.SubElement(properties, w + "spacing")
        spacing.set(w + "line", str(minimum * 20))
        spacing.set(w + "lineRule", "atLeast")
    part = "word/document.xml"
    if container != "body":
        paragraphs = [deepcopy(p) for p in body if p.tag == w + "p"]
        section = body.find(w + "sectPr")
        body[:] = []
        if container == "cell":
            table = E.SubElement(body, w + "tbl")
            properties = E.SubElement(table, w + "tblPr")
            E.SubElement(properties, w + "tblW", {w + "w": "9360", w + "type": "dxa"})
            E.SubElement(properties, w + "tblLayout", {w + "type": "fixed"})
            E.SubElement(E.SubElement(table, w + "tblGrid"), w + "gridCol", {w + "w": "9360"})
            cell = E.SubElement(E.SubElement(table, w + "tr"), w + "tc")
            E.SubElement(E.SubElement(cell, w + "tcPr"), w + "tcW", {w + "w": "9360", w + "type": "dxa"})
            cell.extend(paragraphs)
            E.SubElement(body, w + "p")
        else:
            part = "word/headerMinimum.xml"
            header = E.Element(w + "hdr")
            header.extend(paragraphs)
            parts[part] = E.tostring(header, encoding="utf-8", xml_declaration=True)
            relationships = E.fromstring(parts["word/_rels/document.xml.rels"])
            E.SubElement(relationships, rel + "Relationship", dict(Id="rIdMinimumHeader",
                Type=r[1:-1] + "/header", Target="headerMinimum.xml"))
            parts["word/_rels/document.xml.rels"] = E.tostring(relationships, encoding="utf-8", xml_declaration=True)
            types = E.fromstring(parts["[Content_Types].xml"])
            E.SubElement(types, ct + "Override", dict(PartName="/word/headerMinimum.xml",
                ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"))
            parts["[Content_Types].xml"] = E.tostring(types, encoding="utf-8", xml_declaration=True)
            section.insert(0, E.Element(w + "headerReference", {w + "type": "default", r + "id": "rIdMinimumHeader"}))
            E.SubElement(E.SubElement(E.SubElement(body, w + "p"), w + "r"), w + "t").text = "Composed story."
        if section is not None:
            body.append(section)
    parts["word/document.xml"] = E.tostring(document, encoding="utf-8", xml_declaration=True)
    name = "math-minimum-line-" + str(minimum) + ("" if container == "body" else "-" + container)
    destination = output / (name + ".docx")
    with ZipFile(destination, "w", ZIP_DEFLATED) as archive:
        for path, data in sorted(parts.items()):
            info = ZipInfo(path, (2026, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    expected = ["Before formula:", "after formula.", "Side limits:", "Stacked limits:", "Following prose."]
    if container == "header":
        expected.append("Composed story.")
    rows.append(dict(Id=name, Input=destination.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(destination.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        HeldOut=False, ExpectedEquationCount=3, ExpectedText=expected,
        Container=container, Part=part, MinimumLineHeight=minimum))
(output / "minimum-line-math.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Four minimum-slot controls generated; integral source unchanged.")
