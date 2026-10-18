"""Bounded matrix/radical composition controls, derived from a public matrix.

Keep existing and held-out inputs unchanged. Office references are exported
independently; this generator never supplies expected layout coordinates.
"""
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib
import json

root = Path(__file__).resolve().parent.parent
cases = root / "tests/Lokad.OoxPdf.Tests/Cases"
source = cases / "docx-omml-matrix-inline-literal.docx"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
for prefix, namespace in [("w", w), ("m", m)]:
    E.register_namespace(prefix, namespace[1:-1])

with ZipFile(source) as archive:
    original = {name: archive.read(name) for name in archive.namelist()}
source_hash = hashlib.sha256(source.read_bytes()).hexdigest().upper()
rows = []

def control(owner, size):
    properties = E.SubElement(E.SubElement(owner, m + "ctrlPr"), w + "rPr")
    E.SubElement(properties, w + "rFonts", {w + "ascii": "Cambria Math", w + "hAnsi": "Cambria Math"})
    E.SubElement(properties, w + "sz", {w + "val": str(size * 2)})

for variant in ["inline", "display-indexed", "inline-control20"]:
    entries = dict(original)
    document = E.fromstring(entries["word/document.xml"])
    paragraph = document.find(w + "body/" + w + "p")
    equation = paragraph.find(m + "oMath")
    matrix = deepcopy(equation.find(m + "m"))
    if variant == "inline-control20":
        properties = E.Element(m + "mPr")
        control(properties, 20)
        matrix.insert(0, properties)
    equation.clear()
    radical = E.SubElement(equation, m + "rad")
    properties = E.SubElement(radical, m + "radPr")
    indexed = variant == "display-indexed"
    E.SubElement(properties, m + "degHide", {m + "val": "0" if indexed else "1"})
    control(properties, 12)
    degree = E.SubElement(radical, m + "deg")
    if indexed:
        run = deepcopy(matrix.find(".//" + m + "r"))
        run.find(m + "t").text = "3"
        degree.append(run)
    E.SubElement(radical, m + "e").append(matrix)
    if indexed:
        for child in list(paragraph):
            if child.tag != w + "pPr":
                paragraph.remove(child)
        E.SubElement(paragraph, m + "oMathPara").append(equation)
    entries["word/document.xml"] = E.tostring(document, encoding="utf-8", xml_declaration=True)
    name = "radical-matrix-" + variant
    destination = cases / "CompositionProbes" / (name + ".docx")
    with ZipFile(destination, "w", ZIP_DEFLATED) as archive:
        for path, data in sorted(entries.items()):
            info = ZipInfo(path, (2026, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    rows.append(dict(Id=name, Input="tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/" + name + ".docx",
        Sha256=hashlib.sha256(destination.read_bytes()).hexdigest().upper(), SourceSha256=source_hash,
        HeldOut=False, ExpectedEquationCount=1, ExpectedText=["Following prose."],
        Features=["radical", "matrix", "indexed" if indexed else "inline"],
        ExpectedMathText="abcd" + ("3" if indexed else "")))
(cases / "CompositionProbes/radical-matrix.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Generated three radical/matrix composition controls; existing sources unchanged.")
