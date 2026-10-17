"""Keep one mixed paragraph unchanged while moving a plain nested table around it."""
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
import hashlib
import json
import xml.etree.ElementTree as E

root = Path(__file__).resolve().parent.parent
output = root / "tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
source = output / "math-picture-first-table-single-cell.docx"
with ZipFile(source) as z:
    original = {name: z.read(name) for name in z.namelist()}
xml = E.fromstring(original["word/document.xml"])
body = xml.find(w + "body")
table = body.find(w + "tbl")
mixed = table.find(w + "tr/" + w + "tc/" + w + "p")
assert len(list(mixed.iter(m + "oMath"))) == 2

def prose(text):
    paragraph = E.Element(w + "p")
    paragraph.append(deepcopy(mixed.find(w + "pPr")))
    run = E.SubElement(paragraph, w + "r")
    run.append(deepcopy(paragraph.find(w + "pPr/" + w + "rPr")))
    E.SubElement(run, w + "t").text = text
    return paragraph

def cell_with(table, children):
    cell = table.find(w + "tr/" + w + "tc")
    for child in list(cell):
        if child.tag != w + "tcPr":
            cell.remove(child)
    for child in children:
        cell.append(child)
    return table

rows = []
for placement in ["plain", "before", "after", "inside"]:
    document = deepcopy(xml)
    outer = document.find(w + "body/" + w + "tbl")
    nested = cell_with(deepcopy(table), [prose("Public nested cell")])
    if placement == "plain":
        children = [deepcopy(mixed)]
    elif placement == "before":
        children = [nested, deepcopy(mixed)]
    elif placement == "after":
        children = [deepcopy(mixed), nested, prose("Public outer tail")]
    else:
        nested = cell_with(deepcopy(table), [deepcopy(mixed)])
        children = [prose("Public outer head"), nested, prose("Public outer tail")]
    cell_with(outer, children)
    entries = dict(original)
    entries["word/document.xml"] = E.tostring(document, encoding="utf-8", xml_declaration=True)
    target = output / ("mixed-nested-cell-" + placement + ".docx")
    with ZipFile(target, "w", ZIP_DEFLATED) as z:
        for name, payload in sorted(entries.items()):
            info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, payload)
    expected = []
    if placement in ["before", "after"]:
        expected.append("Public nested cell")
    if placement in ["after", "inside"]:
        expected.append("Public outer tail")
    if placement == "inside":
        expected.append("Public outer head")
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        Placement=placement, HeldOut=False, ExpectedEquationCount=2,
        ExpectedImagePlacements=1, ExpectedText=expected))
(output / "mixed-nested-cells.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Pinned four nested-cell placement controls without changing the mixed paragraph")
