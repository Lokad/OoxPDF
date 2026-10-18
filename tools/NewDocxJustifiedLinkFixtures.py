"""Add source links and wrapped prose around existing body/cell math pictures."""
from pathlib import Path
from xml.etree import ElementTree as E
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
import hashlib
import json
import copy

root = Path(__file__).resolve().parent.parent
output = root / "tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
rows = []
for container in ["body", "cell"]:
    source = output / ("math-display-image-" + container + "-combined.docx")
    with ZipFile(source) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    document = E.fromstring(parts["word/document.xml"])
    paragraph = next(p for p in document.iter(w + "p") if list(p.iter(m + "oMath")))
    original = list(paragraph)
    assert paragraph.find(w + "pPr") is None

    def text(value):
        run = E.Element(w + "r")
        E.SubElement(run, w + "t", {"{http://www.w3.org/XML/1998/namespace}space": "preserve"}).text = value
        return run

    paragraph.clear()
    properties = E.SubElement(paragraph, w + "pPr")
    E.SubElement(properties, w + "jc", {w + "val": "both"})
    E.SubElement(paragraph, w + "bookmarkStart", {w + "id": "97", w + "name": "PublicJustification"})
    link = E.SubElement(paragraph, w + "hyperlink", {w + "anchor": "PublicJustification", w + "history": "1"})
    link.append(text("The linked calculation "))
    link.extend(original)
    link.append(text(" remains visible when ordinary public prose wraps around the supported illustration and formulas." * 5))
    E.SubElement(paragraph, w + "bookmarkEnd", {w + "id": "97"})
    changed = E.tostring(document, encoding="utf-8", xml_declaration=True)
    paragraph.clear()
    paragraph.extend(original)
    assert E.tostring(document) == E.tostring(E.fromstring(parts["word/document.xml"]))
    parts["word/document.xml"] = changed
    name = "justified-linked-image-" + container
    destination = output / (name + ".docx")
    with ZipFile(destination, "w", ZIP_DEFLATED) as archive:
        for part, data in sorted(parts.items()):
            info = ZipInfo(part, (2026, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    rows.append(dict(Id=name, Input=destination.relative_to(root).as_posix(),
                     Sha256=hashlib.sha256(destination.read_bytes()).hexdigest().upper(),
                     SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
                     HeldOut=False, ExpectedEquationCount=1, ExpectedInlineImageCount=1,
                     ExpectedText=["The linked calculation", "supported illustration and formulas."]))
(output / "justified-linked-images.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Two linked-picture justification controls generated; source math/image/container retained.")

# Keep the realistic report's content while extending its existing story owners.
report = root / "tests/Lokad.OoxPdf.Tests/Cases/Composition/report.docx"
with ZipFile(report) as archive:
    original = {name: archive.read(name) for name in archive.namelist()}
with ZipFile(output / "justified-linked-image-body.docx") as archive:
    mixed = E.fromstring(archive.read("word/document.xml"))
    image = archive.read("word/media/math-row-trend.png")
paragraph = next(p for p in mixed.iter(w + "p") if list(p.iter(m + "oMath")))
target = next(E.fromstring(original["word/document.xml"]).iter(w + "bookmarkStart")).get(w + "name")
relations = "{http://schemas.openxmlformats.org/package/2006/relationships}"
stories = []
for story, part in [("footer", "word/footer1.xml"), ("footnote", "word/footnotes.xml")]:
    parts = dict(original)
    document = E.fromstring(parts[part])
    owner = document if story == "footer" else next(n for n in document if n.get(w + "id") == "1")
    added = copy.deepcopy(paragraph)
    for child in list(added):
        if child.tag in [w + "bookmarkStart", w + "bookmarkEnd"]:
            added.remove(child)
    link = added.find(w + "hyperlink")
    link.set(w + "anchor", target)
    list(link.iter(w + "t"))[-1].text = " remains visible when ordinary public prose wraps around the supported illustration and formulas."
    owner.append(added)
    parts[part] = E.tostring(document, encoding="utf-8", xml_declaration=True)
    owner.remove(added)
    assert E.tostring(document) == E.tostring(E.fromstring(original[part]))
    relation_part = "word/_rels/" + Path(part).name + ".rels"
    assert relation_part not in parts and "word/media/math-row-trend.png" not in parts
    rels = E.Element(relations + "Relationships")
    E.SubElement(rels, relations + "Relationship", dict(Id="rIdRowImage",
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", Target="media/math-row-trend.png"))
    parts[relation_part] = E.tostring(rels, encoding="utf-8", xml_declaration=True)
    parts["word/media/math-row-trend.png"] = image
    assert [name for name in original if original[name] != parts[name]] == [part]
    name = "linked-inline-" + story + "-report"
    destination = output / (name + ".docx")
    with ZipFile(destination, "w", ZIP_DEFLATED) as archive:
        for path, data in sorted(parts.items()):
            info = ZipInfo(path, (2026, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    stories.append(dict(Id=name, Input=destination.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(destination.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(report.read_bytes()).hexdigest().upper(), HeldOut=False,
        ExpectedEquationCount=7, ExpectedText=["Quarterly operations report", "Inventory availability", "Review section 12.",
            "All quantities in this synthetic example are illustrative.", "The linked calculation"]))
(output / "linked-inline-stories.json").write_text(json.dumps(stories, indent=2) + "\n", encoding="utf-8")
print("Two related/static story controls generated; all original report content retained.")
