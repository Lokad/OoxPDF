"""Add supported fractions without changing the existing page-numbering sources."""
from pathlib import Path
from xml.etree import ElementTree as E
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from copy import deepcopy
import hashlib
import json

root = Path(__file__).resolve().parent.parent
output = root / "tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
E.register_namespace("w", w[1:-1])
E.register_namespace("m", m[1:-1])
with ZipFile(output / "math-footnote-before-body-plain.docx") as z:
    equation = next(E.fromstring(z.read("word/document.xml")).iter(m + "oMath"))
rows = []
for typography in ["calibri-12-zero", "arial-11-zero"]:
    source = output / ("footnote-page-owners-" + typography + ".docx")
    with ZipFile(source) as z:
        entries = {name: z.read(name) for name in z.namelist()}
    xml = E.fromstring(entries["word/document.xml"])
    count = 0
    for paragraph in xml.iter(w + "p"):
        owner = next((child for child in paragraph if any(True for _ in child.iter(w + "footnoteReference"))), None)
        if owner is not None:
            paragraph.insert(list(paragraph).index(owner), deepcopy(equation))
            count += 1
    assert count == 4
    entries["word/document.xml"] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
    target = output / ("math-page-notes-" + typography + ".docx")
    with ZipFile(target, "w", ZIP_DEFLATED) as z:
        for name, data in sorted(entries.items()):
            info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, data)
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        HeldOut=False, ExpectedEquationCount=4, ExpectedMarkers=["09", "10", "09", "10"],
        ExpectedText=["Public destination " + name + "." for name in ["Alpha", "Bravo", "Charlie", "Delta"]]))
(output / "math-page-notes.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Pinned two mixed-math page-numbering controls")
