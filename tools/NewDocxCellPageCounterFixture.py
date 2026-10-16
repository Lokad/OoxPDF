"""Public table PAGE/math control on actual pages nine and ten."""
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import argparse, hashlib, json

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("--body", action="store_true", help="Create one body paragraph crossing pages nine and ten")
args = parser.parse_args()

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
for number in range(1, 10 if args.body else 11):
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
        if args.body:
            spacer = E.SubElement(body, w + "p")
            E.SubElement(E.SubElement(spacer, w + "pPr"), w + "spacing",
                {w + "line": "12000", w + "lineRule": "exact", w + "before": "0", w + "after": "0"})
            E.SubElement(E.SubElement(spacer, w + "r"), w + "t").text = " "
            content = deepcopy(next(p for p in table.iter(w + "p") if p.find(w + "fldSimple") is not None))
            properties = content.find(w + "pPr")
            if properties is None:
                properties = E.Element(w + "pPr"); content.insert(0, properties)
            E.SubElement(properties, w + "widowControl", {w + "val": "0"})
            middle = next(r for r in content.findall(w + "r") if r.find(w + "t") is not None and "middle" in r.find(w + "t").text)
            middle.insert(0, E.Element(w + "br"))
            body.append(content)
            paragraph = E.SubElement(body, w + "p")
            E.SubElement(E.SubElement(paragraph, w + "r"), w + "t").text = "Public body counter page 10."
        else:
            body.append(deepcopy(table))
# Word requires a final paragraph after a table.
E.SubElement(body, w + "p")
body.append(section)
entries["word/document.xml"] = E.tostring(document, encoding="utf-8", xml_declaration=True)
target = cases / ("math-body-page-counter-boundary.docx" if args.body else "math-cell-page-counter-boundary.docx")
with ZipFile(target, "w", ZIP_DEFLATED) as archive:
    for name, data in sorted(entries.items()):
        info = ZipInfo(name, (2026, 10, 16, 0, 0, 0))
        info.compress_type = ZIP_DEFLATED
        archive.writestr(info, data)
manifest = [dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
    Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(), HeldOut=False,
    ExpectedEquationCount=2 if args.body else 4, ExpectedInlineImageCount=0,
    ExpectedText=["Public table counter page 9.", "Public body counter page 10." if args.body else "Public table counter page 10.", "Start", "middle", "End."])]
target.with_suffix(".json").write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
print("Pinned " + target.stem + " with math across the page nine/ten boundary.")
