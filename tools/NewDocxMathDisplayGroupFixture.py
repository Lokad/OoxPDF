"""Author display-group alignment controls without changing the fixed portfolio."""
from copy import deepcopy
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET

repo = Path(__file__).resolve().parents[1]
cases = repo / "tests/Lokad.OoxPdf.Tests/Cases"
source = cases / "CompositionProbes/math-navigation.docx"
source_sha = hashlib.sha256(source.read_bytes()).hexdigest().upper()
assert source_sha == "34F2CE8E3FC2612244DD4F114277318B1EB337A7835738D1CB813D10D2F13ACC"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
ET.register_namespace("w", w[1:-1])
ET.register_namespace("m", m[1:-1])
output = cases / "CompositionProbes/math-display-groups.docx"
manifest = output.with_suffix(".json")
if output.exists():
    pin = json.loads(manifest.read_text(encoding="utf-8"))[0]
    assert pin["SourceSha256"] == source_sha
    assert pin["Sha256"] == hashlib.sha256(output.read_bytes()).hexdigest().upper()
    raise SystemExit(0)
with ZipFile(source) as parent, ZipFile(output, "w", ZIP_DEFLATED) as result:
    for entry in parent.infolist():
        data = parent.read(entry)
        if entry.filename == "word/document.xml":
            doc = ET.fromstring(data)
            body = doc.find(w + "body")
            section = body.find(w + "sectPr")
            formula = next(doc.iter(m + "oMath"))
            body.clear()

            def text(value):
                p = ET.SubElement(body, w + "p")
                ET.SubElement(ET.SubElement(p, w + "r"), w + "t").text = value
                return p

            title = text("Display group controls")
            title.insert(0, ET.Element(w + "commentRangeStart", {w + "id": "0"}))
            ET.SubElement(title, w + "commentRangeEnd", {w + "id": "0"})
            ET.SubElement(ET.SubElement(title, w + "r"), w + "commentReference", {w + "id": "0"})
            for alignment, count in [("left", 2), ("center", 3), ("right", 2)]:
                text(alignment.capitalize() + " paragraph group")
                p = ET.SubElement(body, w + "p")
                ET.SubElement(ET.SubElement(p, w + "pPr"), w + "jc", {w + "val": alignment})
                for _ in range(count):
                    p.append(deepcopy(formula))
            p = text("Inline before ")
            for _ in range(2):
                p.append(deepcopy(formula))
            ET.SubElement(ET.SubElement(p, w + "r"), w + "t", {
                "{http://www.w3.org/XML/1998/namespace}space": "preserve"}).text = " after inline."
            body.append(section)
            data = ET.tostring(doc, encoding="utf-8", xml_declaration=True)
        result.writestr(entry, data)
manifest.write_text(json.dumps([{
    "Id": "math-display-groups", "Input": output.relative_to(repo).as_posix(),
    "Sha256": hashlib.sha256(output.read_bytes()).hexdigest().upper(),
    "SourceSha256": source_sha, "HeldOut": False, "ExpectedEquationCount": 11,
    "ExpectedInlineImageCount": 0,
    "ExpectedText": ["Display group controls", "Left paragraph group", "Center paragraph group",
                     "Right paragraph group", "Inline before", "after inline."]
}], indent=2) + "\n", encoding="utf-8")
