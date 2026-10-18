"""Derive bounded n-ary/matrix probes from the public radical/matrix controls."""
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
for prefix, namespace in [("w", w), ("m", m)]:
    E.register_namespace(prefix, namespace[1:-1])
parents = {row["Id"]: row for row in json.loads((cases / "radical-matrix.json").read_text(encoding="utf-8"))}
rows = []
for variant, parent, operator, placement in [
    ("inline-sum", "inline", "\u2211", "subSup"),
    ("display-product", "display-indexed", "\u220f", "undOvr"),
    ("inline-control20", "inline-control20", "\u2211", "subSup")]:
    source = cases / ("radical-matrix-" + parent + ".docx")
    source_hash = hashlib.sha256(source.read_bytes()).hexdigest().upper()
    assert source_hash == parents[source.stem]["Sha256"]
    with ZipFile(source) as archive:
        entries = {name: archive.read(name) for name in archive.namelist()}
    document = E.fromstring(entries["word/document.xml"])
    equation = document.find(".//" + m + "oMath")
    radical = equation.find(m + "rad")
    matrix = deepcopy(radical.find(m + "e/" + m + "m"))
    nary = E.Element(m + "nary")
    properties = E.SubElement(nary, m + "naryPr")
    E.SubElement(properties, m + "chr", {m + "val": operator})
    E.SubElement(properties, m + "limLoc", {m + "val": placement})
    properties.append(deepcopy(radical.find(m + "radPr/" + m + "ctrlPr")))
    for role, text in [("sub", "i"), ("sup", "n")]:
        run = deepcopy(matrix.find(".//" + m + "r"))
        run.find(m + "t").text = text
        E.SubElement(nary, m + role).append(run)
    E.SubElement(nary, m + "e").append(matrix)
    equation.remove(radical)
    equation.append(nary)
    entries["word/document.xml"] = E.tostring(document, encoding="utf-8", xml_declaration=True)
    name = "nary-matrix-" + variant
    destination = cases / (name + ".docx")
    with ZipFile(destination, "w", ZIP_DEFLATED) as archive:
        for path, data in sorted(entries.items()):
            info = ZipInfo(path, (2026, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    rows.append(dict(Id=name, Input="tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/" + name + ".docx",
        Sha256=hashlib.sha256(destination.read_bytes()).hexdigest().upper(), SourceSha256=source_hash,
        HeldOut=False, ExpectedEquationCount=1, ExpectedText=["Following prose."],
        ExpectedMathText=operator + "abcdni", Features=["nary", "matrix", placement]))
(cases / "nary-matrix.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Generated three n-ary/matrix controls; existing sources unchanged.")
