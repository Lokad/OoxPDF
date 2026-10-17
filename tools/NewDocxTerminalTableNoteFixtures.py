"""Counterfactuals for the source paragraph following a terminal math table."""
from pathlib import Path
from xml.etree import ElementTree as E
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
import hashlib
import json

root = Path(__file__).resolve().parent.parent
output = root / "tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes"
source = output / "math-section-notes-endnote-cell.docx"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
E.register_namespace("w", w[1:-1])
E.register_namespace("m", "http://schemas.openxmlformats.org/officeDocument/2006/math")
rows = []
for shape in ["original", "empty-default", "empty-zero", "prose-zero"]:
    target = source
    if shape != "original":
        with ZipFile(source) as z:
            entries = {name: z.read(name) for name in z.namelist()}
        xml = E.fromstring(entries["word/document.xml"])
        body = xml.find(w + "body")
        paragraph = E.Element(w + "p")
        if shape.endswith("zero"):
            E.SubElement(E.SubElement(paragraph, w + "pPr"), w + "spacing",
                         {w + "before": "0", w + "after": "0"})
        if shape == "prose-zero":
            E.SubElement(E.SubElement(paragraph, w + "r"), w + "t").text = "Public ending paragraph."
        body.insert(list(body).index(body.find(w + "sectPr")), paragraph)
        entries["word/document.xml"] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
        target = output / ("math-terminal-table-note-" + shape + ".docx")
        with ZipFile(target, "w", ZIP_DEFLATED) as z:
            for name, data in sorted(entries.items()):
                info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                z.writestr(info, data)
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        Shape=shape, HeldOut=False, ExpectedEquationCount=8,
        ExpectedText=["Before formula:", "Public destination 4."] +
                     (["Public ending paragraph."] if shape == "prose-zero" else [])))
(output / "math-terminal-table-notes.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Pinned original and three terminal-paragraph counterfactuals")
