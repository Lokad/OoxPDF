"""Public table PAGE/math control on actual pages nine and ten."""
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib, json

root = Path(__file__).resolve().parent.parent
cases = root / "tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes"
source = cases / "math-fields-simple-cell.docx"
pin = next(row for row in json.loads((cases / "math-fields.json").read_text(encoding="utf-8")) if row["Id"] == source.stem)
assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == pin["Sha256"]
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
E.register_namespace("w", w[1:-1])
with ZipFile(source) as archive:
    entries = {name: archive.read(name) for name in archive.namelist()}
document = E.fromstring(entries["word/document.xml"])
body = document.find(w + "body")
section = deepcopy(body.find(w + "sectPr"))
table = deepcopy(body.find(w + "tbl"))
for field in table.iter(w + "fldSimple"):
    field.set(w + "instr", " PAGE ")
    field.find(".//" + w + "t").text = "1"
body.clear()
for number in range(1, 11):
    paragraph = E.SubElement(body, w + "p")
    props = E.SubElement(paragraph, w + "pPr")
    if number != 1:
        E.SubElement(props, w + "pageBreakBefore")
    run = E.SubElement(paragraph, w + "r")
    style = E.SubElement(run, w + "rPr")
    E.SubElement(style, w + "rFonts", {w + "ascii": "Arial", w + "hAnsi": "Arial"})
    E.SubElement(style, w + "sz", {w + "val": "24"})
    E.SubElement(run, w + "t").text = f"Public table counter page {number}."
    if number in (9, 10):
        body.append(deepcopy(table))
# Word requires a final paragraph after a table.
E.SubElement(body, w + "p")
body.append(section)
entries["word/document.xml"] = E.tostring(document, encoding="utf-8", xml_declaration=True)
target = cases / "math-cell-page-counter-boundary.docx"
with ZipFile(target, "w", ZIP_DEFLATED) as archive:
    for name, data in sorted(entries.items()):
        info = ZipInfo(name, (2026, 10, 16, 0, 0, 0))
        info.compress_type = ZIP_DEFLATED
        archive.writestr(info, data)
manifest = [dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
    Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(), HeldOut=False,
    ExpectedEquationCount=4, ExpectedInlineImageCount=0,
    ExpectedText=["Public table counter page 9.", "Public table counter page 10.", "Start", "middle", "End."])]
(cases / "math-cell-page-counter-boundary.json").write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
print("Pinned one ten-page table counter control with math on pages nine and ten.")
