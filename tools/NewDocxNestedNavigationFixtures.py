"""Exercise formula and prose navigation across ordinary nested-cell owners."""
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
source = output / "mixed-nested-cell-inside.docx"
with ZipFile(source) as z:
    original = {name: z.read(name) for name in z.namelist()}
rows = []
for direction in ["within", "incoming", "outgoing"]:
    entries = dict(original)
    xml = E.fromstring(entries["word/document.xml"])
    body = xml.find(w + "body")
    outer = body.find(w + "tbl/" + w + "tr/" + w + "tc")
    mixed = outer.find(w + "tbl/" + w + "tr/" + w + "tc/" + w + "p")
    formulas = mixed.findall(m + "oMathPara")
    assert len(formulas) == 2
    target = formulas[1] if direction != "outgoing" else outer.findall(w + "p")[-1].find(w + "r")
    parent = next(node for node in xml.iter() if target in list(node))
    index = list(parent).index(target)
    parent.insert(index, E.Element(w + "bookmarkStart", {w + "id": "901", w + "name": "PublicNestedTarget"}))
    parent.insert(index + 2, E.Element(w + "bookmarkEnd", {w + "id": "901"}))
    linked = formulas[0] if direction != "incoming" else body.find(w + "p/" + w + "r")
    parent = next(node for node in xml.iter() if linked in list(node))
    index = list(parent).index(linked)
    parent.remove(linked)
    hyperlink = E.Element(w + "hyperlink", {w + "anchor": "PublicNestedTarget"})
    hyperlink.append(linked)
    parent.insert(index, hyperlink)
    entries["word/document.xml"] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
    target = output / ("nested-math-navigation-" + direction + ".docx")
    with ZipFile(target, "w", ZIP_DEFLATED) as z:
        for name, payload in sorted(entries.items()):
            info = ZipInfo(name, (2026, 10, 18, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, payload)
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        Direction=direction, HeldOut=False, ExpectedEquationCount=2, ExpectedImagePlacements=1,
        ExpectedText=["Public outer head", "Public outer tail", "Following prose."]))
(output / "nested-math-navigation.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Pinned three nested navigation directions with unchanged visible source content")
