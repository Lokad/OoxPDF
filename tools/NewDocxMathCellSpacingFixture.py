"""Probe mathematical cell occupancy versus paragraph after-spacing."""
import argparse
from copy import deepcopy
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET

parser = argparse.ArgumentParser()
parser.add_argument("--fixed-grid", action="store_true")
parser.add_argument("--normal-math", action="store_true")
args = parser.parse_args()
suffix = "-fixed" if args.fixed_grid else ""
repo = Path(__file__).resolve().parents[1]
cases = repo / "tests/Lokad.OoxPdf.Tests/Cases"
source = cases / "Composition/report.docx"
source_sha = hashlib.sha256(source.read_bytes()).hexdigest().upper()
assert source_sha == "10F63D162D0D46DBE8CD682BB50452DB6672104706F116EA50CDCAFEE054575F"
output = cases / (("PortableFonts/math-cell-spacing-normal" if args.normal_math else
                   "CompositionProbes/math-cell-spacing") + suffix + ".docx")
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
            section = deepcopy(body.find(w + "sectPr"))
            for child in list(section):
                if child.tag in (w + "headerReference", w + "footerReference"):
                    section.remove(child)
            table = deepcopy(body.find(w + "tbl"))
            fraction = deepcopy(next(item for item in doc.iter(m + "oMath") if item.find(m + "f") is not None))
            for row in list(table.findall(w + "tr")):
                table.remove(row)
            if args.fixed_grid:
                props = table.find(w + "tblPr")
                ET.SubElement(props, w + "tblLayout", {w + "type": "fixed"})
                grid = table.find(w + "tblGrid")
                if grid is None:
                    grid = ET.Element(w + "tblGrid")
                    table.insert(1, grid)
                grid.clear()
                for _ in range(2):
                    ET.SubElement(grid, w + "gridCol", {w + "w": "4500"})
            values = (None, "empty", 0, 6, 12, 24) if args.fixed_grid else (0, 6, 12, 24)
            for after in values:
                row = ET.SubElement(table, w + "tr")
                for math in (False, True):
                    cell = ET.SubElement(row, w + "tc")
                    if args.fixed_grid:
                        ET.SubElement(ET.SubElement(cell, w + "tcPr"), w + "tcW", {w + "type": "dxa", w + "w": "4500"})
                    paragraph = ET.SubElement(cell, w + "p")
                    if after is not None or not math:
                        props = ET.SubElement(paragraph, w + "pPr")
                        if after != "empty" or not math:
                            ET.SubElement(props, w + "spacing", {w + "after": str(after * 20 if math else 0)})
                    if math:
                        paragraph.append(deepcopy(fraction))
                    else:
                        ET.SubElement(ET.SubElement(paragraph, w + "r"), w + "t").text = f"After spacing {after} points"
            body.clear()
            body.append(table)
            body.append(ET.Element(w + "p"))
            body.append(section)
            if args.normal_math:
                for run in doc.iter(m + "r"):
                    props = run.find(m + "rPr")
                    if props is None:
                        props = ET.Element(m + "rPr")
                        run.insert(0, props)
                    ET.SubElement(props, m + "nor", {m + "val": "1"})
            data = ET.tostring(doc, encoding="utf-8", xml_declaration=True)
        result.writestr(entry, data)
manifest.write_text(json.dumps([{
    "Id": ("math-cell-spacing-normal" if args.normal_math else "math-cell-spacing") + suffix, "Input": output.relative_to(repo).as_posix(),
    "Sha256": hashlib.sha256(output.read_bytes()).hexdigest().upper(),
    "SourceSha256": source_sha, "HeldOut": False, "ExpectedEquationCount": 6 if args.fixed_grid else 4,
    "ExpectedInlineImageCount": 0,
    "ExpectedText": [f"After spacing {after} points" for after in ((None, "empty", 0, 6, 12, 24) if args.fixed_grid else (0, 6, 12, 24))]
}], indent=2) + "\n", encoding="utf-8")
