"""Compose existing public formulas/pictures with naturally paginated note owners."""
from copy import deepcopy
from pathlib import Path
from xml.etree import ElementTree as E
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
import hashlib
import json

root = Path(__file__).resolve().parent.parent
output = root / "tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
r = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
wp = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
rels = "{http://schemas.openxmlformats.org/package/2006/relationships}"
ct = "{http://schemas.openxmlformats.org/package/2006/content-types}"
for prefix, namespace in [("w", w), ("m", m), ("r", r), ("a", a), ("wp", wp)]:
    E.register_namespace(prefix, namespace[1:-1])

rows = []
for typography in ["calibri-12-zero", "arial-11-zero"]:
    source = output / ("spanning-page-notes-" + typography + ".docx")
    donor = output / ("math-page-notes-compound-" + typography + ".docx")
    with ZipFile(source) as z:
        entries = {name: z.read(name) for name in z.namelist()}
    with ZipFile(donor) as z:
        donor_xml = E.fromstring(z.read("word/document.xml"))
        picture_bytes = z.read("word/media/page-note-compound.png")
    donor_paragraph = next(p for p in donor_xml.iter(w + "p") if p.find(m + "oMath") is not None)
    children = list(donor_paragraph)
    start = next(i for i, child in enumerate(children) if child.tag == w + "bookmarkStart")
    end = next(i for i, child in enumerate(children) if child.tag == w + "bookmarkEnd")
    composition = children[start:end + 1]
    assert sum(child.tag == m + "oMath" for child in composition) == 2
    assert sum(1 for child in composition for _ in child.iter(w + "drawing")) == 1

    xml = E.fromstring(entries["word/document.xml"])
    paragraph = xml.find(w + "body/" + w + "p")
    markers = [run for run in paragraph.findall(w + "r") if run.find(w + "footnoteReference") is not None]
    assert len(markers) == 4
    space = E.Element(w + "r")
    E.SubElement(space, w + "t", {"{http://www.w3.org/XML/1998/namespace}space": "preserve"}).text = " "
    for index, marker in enumerate(markers):
        items = [deepcopy(space), *deepcopy(composition), deepcopy(space)]
        for item in items:
            for bookmark in item.iter(w + "bookmarkStart"):
                bookmark.set(w + "id", str(500 + index))
                bookmark.set(w + "name", "PublicSpanningCalculation" + str(index + 1))
            for bookmark in item.iter(w + "bookmarkEnd"):
                bookmark.set(w + "id", str(500 + index))
            for blip in item.iter(a + "blip"):
                blip.set(r + "embed", "spanningNoteCompositionImage")
            for properties in item.iter(wp + "docPr"):
                properties.set("id", str(1000 + index))
                properties.set("name", "Public spanning picture " + str(index + 1))
        position = list(paragraph).index(marker)
        for offset, item in enumerate(items):
            paragraph.insert(position + offset, item)
    entries["word/document.xml"] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
    relationships = E.fromstring(entries["word/_rels/document.xml.rels"])
    E.SubElement(relationships, rels + "Relationship", {
        "Id": "spanningNoteCompositionImage",
        "Type": r[1:-1] + "/image",
        "Target": "media/spanning-note-composition.png",
    })
    entries["word/_rels/document.xml.rels"] = E.tostring(relationships, encoding="utf-8", xml_declaration=True)
    entries["word/media/spanning-note-composition.png"] = picture_bytes
    types = E.fromstring(entries["[Content_Types].xml"])
    if not any(item.get("Extension") == "png" for item in types):
        E.SubElement(types, ct + "Default", {"Extension": "png", "ContentType": "image/png"})
    entries["[Content_Types].xml"] = E.tostring(types, encoding="utf-8", xml_declaration=True)
    target = output / ("spanning-note-math-" + typography + ".docx")
    with ZipFile(target, "w", ZIP_DEFLATED) as z:
        for name, data in sorted(entries.items()):
            info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, data)
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        DonorSha256=hashlib.sha256(donor.read_bytes()).hexdigest().upper(),
        HeldOut=False, ExpectedEquationCount=8, ExpectedImagePlacements=4,
        ExpectedText=["Public destination Alpha.", "Public destination Delta."]))
(output / "spanning-note-math.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Pinned two natural paragraphs with eight formulas/four images/four bookmarks/four notes")
