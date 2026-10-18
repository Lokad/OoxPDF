"""Combine qualified operator/matrix formulas with a development specification.

Existing and held-out documents retain their bytes; tables, prose, notes,
numbering, links, comments and static stories remain from the public source.
"""
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib
import json

root = Path(__file__).resolve().parent.parent
cases = root / "tests/Lokad.OoxPdf.Tests/Cases"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
for prefix, namespace in [("w", w), ("m", m)]:
    E.register_namespace(prefix, namespace[1:-1])
portfolio = json.loads((cases / "Composition/portfolio.json").read_text(encoding="utf-8"))
source_row = next(row for row in portfolio if row["Id"] == "specification")
source = root / source_row["Input"]
assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == source_row["Sha256"]
with ZipFile(source) as archive:
    entries = {name: archive.read(name) for name in archive.namelist()}
document = E.fromstring(entries["word/document.xml"])
body = document.find(w + "body")
direct = [equation for paragraph in body.findall(w + "p") for equation in paragraph.iter(m + "oMath")]
cell = body.find(".//" + w + "tbl//" + m + "oMath")
assert len(direct) == 4 and cell is not None
provenance = []
for target, name in [(direct[0], "radical-matrix-inline"), (direct[1], "nary-matrix-inline-sum"),
                     (cell, "nary-matrix-display-product")]:
    path = cases / "CompositionProbes" / (name + ".docx")
    with ZipFile(path) as archive:
        equation = E.fromstring(archive.read("word/document.xml")).find(".//" + m + "oMath")
    target[:] = [deepcopy(child) for child in equation]
    provenance.append(dict(Input="CompositionProbes/" + path.name,
        Sha256=hashlib.sha256(path.read_bytes()).hexdigest().upper()))
entries["word/document.xml"] = E.tostring(document, encoding="utf-8", xml_declaration=True)
name = "composed-operator-matrix-specification"
destination = cases / "Composition" / (name + ".docx")
with ZipFile(destination, "w", ZIP_DEFLATED) as archive:
    for path, data in sorted(entries.items()):
        info = ZipInfo(path, (2026, 1, 1, 0, 0, 0))
        info.compress_type = ZIP_DEFLATED
        archive.writestr(info, data)
row = dict(source_row, Id=name, Input="tests/Lokad.OoxPdf.Tests/Cases/Composition/" + name + ".docx",
    Sha256=hashlib.sha256(destination.read_bytes()).hexdigest().upper(), HeldOut=False,
    SourceSha256=source_row["Sha256"], FormulaSources=provenance,
    Features=source_row["Features"] + ["matrix", "radical", "nary"])
(cases / "Composition/operator-matrix-specification.json").write_text(json.dumps([row], indent=2) + "\n", encoding="utf-8")
print("Generated one compound development specification; existing sources unchanged.")
