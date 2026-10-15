"""Author public page-capacity and continuation controls; never regenerate pinned inputs implicitly."""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import hashlib
import json
import xml.etree.ElementTree as ET

repo = Path(__file__).resolve().parents[1]
cases = repo / "tests/Lokad.OoxPdf.Tests/Cases"
source = cases / "Composition/letter.docx"
source_hash = hashlib.sha256(source.read_bytes()).hexdigest().upper()
assert source_hash == "50FE6A8DDA140D39F98C9F0C4327642E90C0839A858526251060C2BEE329B9A4"
output = cases / "ReviewBodyFrames"
output.mkdir(exist_ok=True)
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
r = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
ET.register_namespace("w", w[1:-1])
ET.register_namespace("r", r[1:-1])
manifest = output / "probes.json"
existing = json.loads(manifest.read_text(encoding="utf-8")) if manifest.exists() else []
rows = []
for story in ["body", "header", "footer", "table-widow-on", "table-widow-off"]:
    path = output / (story + ".docx")
    if path.exists():
        pin = next(row for row in existing if row["Id"] == story)
        assert pin["SourceSha256"] == source_hash
        assert hashlib.sha256(path.read_bytes()).hexdigest().upper() == pin["Sha256"]
        rows.append(pin)
        continue
    with ZipFile(source) as template, ZipFile(path, "w", ZIP_DEFLATED) as target:
        for entry in template.infolist():
            data = template.read(entry)
            if entry.filename == "word/document.xml":
                doc = ET.fromstring(data)
                body = doc.find(w + "body")
                section = body.find(w + "sectPr")
                for child in list(section):
                    if child.tag in [w + "headerReference", w + "footerReference"] and child.tag != w + story + "Reference":
                        section.remove(child)
                body.clear()
                for label, height in [("Public review anchor", 14), ("Public filler", 510), ("Plain spacer", 72), ("Following body", 90)]:
                    paragraph = ET.SubElement(body, w + "p")
                    props = ET.SubElement(paragraph, w + "pPr")
                    ET.SubElement(props, w + "spacing", {w + "before": "0", w + "after": "0", w + "lineRule": "exact", w + "line": str(height * 20)})
                    if height == 14:
                        ET.SubElement(paragraph, w + "commentRangeStart", {w + "id": "0"})
                    run = ET.SubElement(paragraph, w + "r")
                    ET.SubElement(ET.SubElement(run, w + "rPr"), w + "sz", {w + "val": "24"})
                    ET.SubElement(run, w + "t").text = label
                    if height == 14:
                        ET.SubElement(paragraph, w + "commentRangeEnd", {w + "id": "0"})
                        ET.SubElement(ET.SubElement(paragraph, w + "r"), w + "commentReference", {w + "id": "0"})
                if story.startswith("table-widow-"):
                    for paragraph in list(body)[1:]:
                        body.remove(paragraph)
                    table = ET.SubElement(body, w + "tbl")
                    props = ET.SubElement(table, w + "tblPr")
                    ET.SubElement(props, w + "tblW", {w + "type": "dxa", w + "w": "3200"})
                    ET.SubElement(props, w + "tblLayout", {w + "type": "fixed"})
                    ET.SubElement(ET.SubElement(table, w + "tblGrid"), w + "gridCol", {w + "w": "3200"})
                    cell = ET.SubElement(ET.SubElement(table, w + "tr"), w + "tc")
                    ET.SubElement(ET.SubElement(cell, w + "tcPr"), w + "tcW", {w + "type": "dxa", w + "w": "3200"})
                    paragraph = ET.SubElement(cell, w + "p")
                    props = ET.SubElement(paragraph, w + "pPr")
                    ET.SubElement(props, w + "spacing", {w + "before": "0", w + "after": "0", w + "lineRule": "auto", w + "line": "240"})
                    ET.SubElement(props, w + "widowControl", {w + "val": "1" if story.endswith("-on") else "0"})
                    run = ET.SubElement(paragraph, w + "r")
                    ET.SubElement(ET.SubElement(run, w + "rPr"), w + "sz", {w + "val": "48"})
                    ET.SubElement(run, w + "t").text = " ".join(["AAA"] * 200)
                    ET.SubElement(ET.SubElement(ET.SubElement(body, w + "p"), w + "r"), w + "t").text = "After public continuation."
                body.append(section)
                data = ET.tostring(doc, encoding="utf-8", xml_declaration=True)
            elif entry.filename in ["word/header1.xml", "word/footer1.xml"]:
                doc = ET.fromstring(data)
                doc.clear()
                ET.SubElement(ET.SubElement(ET.SubElement(doc, w + "p"), w + "r"), w + "t").text = "Public static capacity control"
                data = ET.tostring(doc, encoding="utf-8", xml_declaration=True)
            elif entry.filename == "word/_rels/document.xml.rels":
                doc = ET.fromstring(data)
                for rel in list(doc):
                    kind = rel.get("Type", "").rsplit("/", 1)[-1]
                    if kind in ["footnotes", "header", "footer"] and kind != story:
                        doc.remove(rel)
                data = ET.tostring(doc, encoding="utf-8", xml_declaration=True)
            target.writestr(entry, data)
    rows.append({"Id": story, "Input": path.relative_to(repo).as_posix(), "Sha256": hashlib.sha256(path.read_bytes()).hexdigest().upper(),
                 "SourceSha256": source_hash, "ExpectedText": ["Public review anchor", "AAA", "After public continuation."] if story.startswith("table-") else ["Public review anchor", "Public filler", "Plain spacer", "Following body"],
                 "ExpectedEquationCount": 0, "HeldOut": False})
manifest.write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
