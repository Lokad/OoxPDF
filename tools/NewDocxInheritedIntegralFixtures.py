"""Place standalone integrals in cells/static stories with document defaults."""
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
ct = "{http://schemas.openxmlformats.org/package/2006/content-types}"
rel = "{http://schemas.openxmlformats.org/package/2006/relationships}"
rows = []
for container, kind, location in (("cell", "side", "subSup"), ("header", "stacked", "undOvr")):
    source = root / ("tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/display-integral-" + kind + ".docx")
    with ZipFile(source) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    document = E.fromstring(parts["word/document.xml"])
    body = document.find(w + "body")
    section = body.find(w + "sectPr")
    paragraphs = [deepcopy(p) for p in body if p.tag == w + "p"]
    assert len(paragraphs) == 4
    for paragraph in paragraphs:
        for properties in paragraph.iter(m + "naryPr"):
            for node in list(properties):
                if node.tag == m + "limLoc":
                    properties.remove(node)
    settings = E.fromstring(parts["word/settings.xml"])
    math = settings.find(m + "mathPr")
    if math is None:
        math = E.SubElement(settings, m + "mathPr")
    for node in list(math):
        if node.tag == m + "intLim":
            math.remove(node)
    E.SubElement(math, m + "intLim", {m + "val": location})
    parts["word/settings.xml"] = E.tostring(settings, encoding="utf-8", xml_declaration=True)
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
        header = E.Element(w + "hdr")
        header.extend(paragraphs)
        parts["word/headerC195.xml"] = E.tostring(header, encoding="utf-8", xml_declaration=True)
        relationships = E.fromstring(parts["word/_rels/document.xml.rels"])
        E.SubElement(relationships, rel + "Relationship", dict(
            Id="rIdC195Header", Type=r[1:-1] + "/header", Target="headerC195.xml"))
        parts["word/_rels/document.xml.rels"] = E.tostring(relationships, encoding="utf-8", xml_declaration=True)
        types = E.fromstring(parts["[Content_Types].xml"])
        E.SubElement(types, ct + "Override", dict(PartName="/word/headerC195.xml",
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"))
        parts["[Content_Types].xml"] = E.tostring(types, encoding="utf-8", xml_declaration=True)
        section.insert(0, E.Element(w + "headerReference", {w + "type": "default", r + "id": "rIdC195Header"}))
        E.SubElement(E.SubElement(E.SubElement(body, w + "p"), w + "r"), w + "t").text = "Composed story."
    if section is not None:
        body.append(section)
    parts["word/document.xml"] = E.tostring(document, encoding="utf-8", xml_declaration=True)
    name = "inherited-integral-" + container
    destination = output / (name + ".docx")
    with ZipFile(destination, "w", ZIP_DEFLATED) as archive:
        for path, data in sorted(parts.items()):
            info = ZipInfo(path, (2026, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    expected = ["Before formula:", "after formula.", "Following prose."]
    if container == "header":
        expected.append("Composed story.")
    rows.append(dict(Id=name, Input=destination.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(destination.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(), Container=container,
        DocumentIntegralLimit=location, HeldOut=False, ExpectedEquationCount=1, ExpectedText=expected))
(output / "inherited-integral.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Two inherited-limit/container controls generated; original sources unchanged.")
