"""Place existing mixed page-note paragraphs inside simple public table cells."""
from copy import deepcopy
from pathlib import Path
from xml.etree import ElementTree as E
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
import hashlib
import json

root = Path(__file__).resolve().parent.parent
output = root / "tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
E.register_namespace("w", w[1:-1])
rows = []
for typography, composition in [(t, c) for c in ["", "mixed-custom-"] for t in ["calibri-12-zero", "arial-11-zero"]]:
    source = output / (("math-page-notes-compound-" if not composition else "math-page-notes-mixed-custom-") + typography + ".docx")
    with ZipFile(source) as z:
        entries = {name: z.read(name) for name in z.namelist()}
    xml = E.fromstring(entries["word/document.xml"])
    body = xml.find(w + "body")
    originals = list(body)
    body.clear()
    cells = 0
    for child in originals:
        if child.tag != w + "p" or child.find(".//" + w + "footnoteReference") is None:
            body.append(deepcopy(child))
            continue
        table = E.SubElement(body, w + "tbl")
        properties = E.SubElement(table, w + "tblPr")
        E.SubElement(properties, w + "tblW", {w + "w": "9360", w + "type": "dxa"})
        borders = E.SubElement(properties, w + "tblBorders")
        for edge in ["top", "left", "bottom", "right", "insideH", "insideV"]:
            E.SubElement(borders, w + edge, {w + "val": "nil"})
        E.SubElement(properties, w + "tblLayout", {w + "type": "fixed"})
        margins = E.SubElement(properties, w + "tblCellMar")
        for edge in ["top", "left", "bottom", "right"]:
            E.SubElement(margins, w + edge, {w + "w": "0", w + "type": "dxa"})
        grid = E.SubElement(table, w + "tblGrid")
        E.SubElement(grid, w + "gridCol", {w + "w": "9360"})
        row = E.SubElement(table, w + "tr")
        cell = E.SubElement(row, w + "tc")
        cell_properties = E.SubElement(cell, w + "tcPr")
        E.SubElement(cell_properties, w + "tcW", {w + "w": "9360", w + "type": "dxa"})
        E.SubElement(cell_properties, w + "vAlign", {w + "val": "top"})
        cell.append(deepcopy(child))
        cells += 1
    assert cells == 4
    entries["word/document.xml"] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
    target = output / ("cell-page-notes-" + composition + typography + ".docx")
    with ZipFile(target, "w", ZIP_DEFLATED) as z:
        for name, data in sorted(entries.items()):
            info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, data)
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        HeldOut=False, ExpectedEquationCount=8, ExpectedImagePlacements=4,
        ExpectedText=["Public destination Alpha.", "Public destination Delta."]))
(output / "cell-page-notes.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Pinned four four-cell controls with automatic/custom note owners and shared math/images")
