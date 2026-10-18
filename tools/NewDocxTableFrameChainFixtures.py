"""Two small source-order controls for a shared printed body origin."""
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib
import json

root = Path(__file__).resolve().parent.parent
cases = root / "tests/Lokad.OoxPdf.Tests/Cases"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
r = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
pkg = "{http://schemas.openxmlformats.org/package/2006/relationships}"
ct = "{http://schemas.openxmlformats.org/package/2006/content-types}"
source = cases / "Composition/composed-nested-navigation-report.docx"
source_hash = hashlib.sha256(source.read_bytes()).hexdigest().upper()
assert source_hash == "9B907564FC68DA3439B773366A6D81F8E8F1389DA94C56E819AB08C88E638A73"
with ZipFile(source) as archive:
    original = {name: archive.read(name) for name in archive.namelist()}
rows = []
for variant in ["direct", "paragraph"]:
    entries = dict(original)
    document = E.fromstring(entries["word/document.xml"])
    body = document.find(w + "body")
    paragraphs = body.findall(w + "p")
    tables = body.findall(w + "tbl")
    section = deepcopy(body.find(w + "sectPr"))
    assert len(tables) == 2
    body.clear()
    body.extend([deepcopy(paragraphs[0]), deepcopy(paragraphs[1]), deepcopy(tables[0])])
    if variant == "paragraph":
        empty = E.SubElement(body, w + "p")
        E.SubElement(E.SubElement(empty, w + "pPr"), w + "spacing",
            {w + "before": "0", w + "after": "0", w + "line": "240", w + "lineRule": "auto"})
    body.append(deepcopy(tables[1]))
    tail = E.SubElement(body, w + "p")
    E.SubElement(tail, w + "commentRangeStart", {w + "id": "0"})
    E.SubElement(E.SubElement(tail, w + "r"), w + "t").text = "Following tables."
    E.SubElement(tail, w + "commentRangeEnd", {w + "id": "0"})
    E.SubElement(E.SubElement(tail, w + "r"), w + "commentReference", {w + "id": "0"})
    body.append(section)
    entries["word/document.xml"] = E.tostring(document, encoding="utf-8", xml_declaration=True)
    relationships = E.fromstring(entries["word/_rels/document.xml.rels"])
    for relationship in list(relationships):
        if relationship.get("Type", "").endswith("/footnotes"):
            relationships.remove(relationship)
    entries["word/_rels/document.xml.rels"] = E.tostring(relationships, encoding="utf-8", xml_declaration=True)
    types = E.fromstring(entries["[Content_Types].xml"])
    for item in list(types):
        if item.get("PartName") == "/word/footnotes.xml":
            types.remove(item)
    entries["[Content_Types].xml"] = E.tostring(types, encoding="utf-8", xml_declaration=True)
    entries.pop("word/footnotes.xml", None)
    target = cases / ("CompositionProbes/table-frame-chain-" + variant + ".docx")
    with ZipFile(target, "w", ZIP_DEFLATED) as archive:
        for name, payload in sorted(entries.items()):
            info = ZipInfo(name, (2026, 10, 18, 0, 0, 0)); info.compress_type = ZIP_DEFLATED
            archive.writestr(info, payload)
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(), SourceSha256=source_hash,
        HeldOut=False, ExpectedEquationCount=4, ExpectedInlineImageCount=1,
        ExpectedText=["Planning input", "Observed series", "Calculation summary", "Detail: review the accepted ratio.", "Following tables."],
        Variant=variant))
(cases / "CompositionProbes/table-frame-chain.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == source_hash
print("Pinned adjacent and paragraph-separated table controls with the same mixed body frame")
