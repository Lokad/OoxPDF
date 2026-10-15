"""Author unchanged-format links and formula bookmark targets in body and cell."""
from copy import deepcopy
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET

repo = Path(__file__).resolve().parents[1]
source = repo / "tests/Lokad.OoxPdf.Tests/Cases/docx-omml-inline-fraction-middle12.docx"
source_sha = hashlib.sha256(source.read_bytes()).hexdigest().upper()
assert source_sha == "142C73F88E81FE11BCB8FFCDB5E465ACD7D0214DA200E4517DCF713B0753F810"
output = repo / "tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/single-math-annotations.docx"
manifest = output.with_suffix(".json")
if output.exists():
    pin = json.loads(manifest.read_text(encoding="utf-8"))[0]
    assert pin["SourceSha256"] == source_sha
    assert pin["Sha256"] == hashlib.sha256(output.read_bytes()).hexdigest().upper()
    raise SystemExit(0)
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
ET.register_namespace("w", w[1:-1])
ET.register_namespace("m", m[1:-1])
with ZipFile(source) as parent, ZipFile(output, "w", ZIP_DEFLATED) as result:
    for entry in parent.infolist():
        data = parent.read(entry)
        if entry.filename == "word/document.xml":
            doc = ET.fromstring(data)
            body = doc.find(w + "body")
            section = body.find(w + "sectPr")
            paragraph = next(p for p in body.findall(w + "p") if p.find(m + "oMath") is not None)

            def target(name, identifier):
                p = deepcopy(paragraph)
                math = p.find(m + "oMath")
                offset = list(p).index(math)
                p.insert(offset, ET.Element(w + "bookmarkStart", {w + "id": str(identifier), w + "name": name}))
                p.insert(offset + 2, ET.Element(w + "bookmarkEnd", {w + "id": str(identifier)}))
                run = next(r for r in p.findall(w + "r") if r.find(w + "t") is not None)
                offset = list(p).index(run)
                p.remove(run)
                link = ET.Element(w + "hyperlink", {w + "anchor": name})
                link.append(run)
                p.insert(offset, link)
                return p

            body.clear()
            body.append(target("BodyFormula", 41))
            table = ET.SubElement(body, w + "tbl")
            ET.SubElement(ET.SubElement(table, w + "tblPr"), w + "tblW", {w + "type": "dxa", w + "w": "9360"})
            ET.SubElement(ET.SubElement(table, w + "tblGrid"), w + "gridCol", {w + "w": "9360"})
            cell = ET.SubElement(ET.SubElement(table, w + "tr"), w + "tc")
            ET.SubElement(ET.SubElement(cell, w + "tcPr"), w + "tcW", {w + "type": "dxa", w + "w": "9360"})
            cell.append(target("CellFormula", 42))
            body.append(ET.Element(w + "p"))
            body.append(section)
            data = ET.tostring(doc, encoding="utf-8", xml_declaration=True)
        result.writestr(entry, data)
manifest.write_text(json.dumps([{
    "Id": "single-math-annotations", "Input": output.relative_to(repo).as_posix(),
    "Sha256": hashlib.sha256(output.read_bytes()).hexdigest().upper(),
    "SourceSha256": source_sha, "HeldOut": False, "ExpectedEquationCount": 2,
    "ExpectedInlineImageCount": 0, "ExpectedText": ["Before formula:", "after formula."]
}], indent=2) + "\n", encoding="utf-8")
