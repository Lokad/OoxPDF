"""Mixed mathematics/picture/bookmark ownership in direct and table note content."""
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
import hashlib
import json
import xml.etree.ElementTree as E

root = Path(__file__).resolve().parent.parent
cases = root / "tests/Lokad.OoxPdf.Tests/Cases"
output = cases / "CompositionProbes"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
r = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
wp = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
rel = "{http://schemas.openxmlformats.org/package/2006/relationships}"
ct = "{http://schemas.openxmlformats.org/package/2006/content-types}"
for prefix, namespace in [("w", w), ("m", m), ("r", r), ("a", a), ("wp", wp)]:
    E.register_namespace(prefix, namespace[1:-1])
donor = output / "math-page-notes-compound-calibri-12-zero.docx"
with ZipFile(donor) as z:
    donor_xml = E.fromstring(z.read("word/document.xml"))
    picture = z.read("word/media/page-note-compound.png")
paragraph = next(p for p in donor_xml.iter(w + "p") if p.find(m + "oMath") is not None)
children = list(paragraph)
start = next(i for i, item in enumerate(children) if item.tag == w + "bookmarkStart")
end = next(i for i, item in enumerate(children) if item.tag == w + "bookmarkEnd")
composition = children[start:end + 1]
assert sum(item.tag == m + "oMath" for item in composition) == 2
assert sum(1 for item in composition for _ in item.iter(w + "drawing")) == 1
table_source = output / "cell-page-notes-plain-calibri-12-zero.docx"
with ZipFile(table_source) as z:
    table_template = E.fromstring(z.read("word/document.xml")).find(w + "body/" + w + "tbl")
rows = []
for kind in ["footnote", "endnote"]:
    source = cases / ("note-custom-mark-" + kind + ".docx")
    for container in ["paragraph", "table"]:
        with ZipFile(source) as z:
            entries = {name: z.read(name) for name in z.namelist()}
        part = "word/" + kind + "s.xml"
        xml = E.fromstring(entries[part])
        story = next(item for item in xml if item.get(w + "id") == "37")
        mixed = E.Element(w + "p")
        mixed.append(deepcopy(paragraph.find(w + "pPr")))
        text = deepcopy(next(run for run in paragraph.findall(w + "r") if run.find(w + "t") is not None))
        for i, item in enumerate(text.iter(w + "t")):
            item.text = "Public note calculation " if i == 0 else ""
        mixed.append(text)
        for item in deepcopy(composition):
            for blip in item.iter(a + "blip"):
                blip.set(r + "embed", "noteCompositionImage")
            for bookmark in item.iter(w + "bookmarkStart"):
                bookmark.set(w + "id", "701")
                bookmark.set(w + "name", "PublicNoteCalculation")
            for bookmark in item.iter(w + "bookmarkEnd"):
                bookmark.set(w + "id", "701")
            for properties in item.iter(wp + "docPr"):
                properties.set("id", "1701")
                properties.set("name", "Public note picture")
            mixed.append(item)
        if container == "table":
            table = deepcopy(table_template)
            cell = table.find(w + "tr/" + w + "tc")
            for child in list(cell):
                if child.tag != w + "tcPr":
                    cell.remove(child)
            cell.append(mixed)
            story.append(table)
            tail = E.SubElement(story, w + "p")
            tail.append(deepcopy(paragraph.find(w + "pPr")))
        else:
            story.append(mixed)
        entries[part] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
        relationships = E.Element(rel + "Relationships")
        E.SubElement(relationships, rel + "Relationship", {
            "Id": "noteCompositionImage", "Type": r[1:-1] + "/image",
            "Target": "media/note-composition.png",
        })
        entries["word/_rels/" + kind + "s.xml.rels"] = E.tostring(relationships, encoding="utf-8", xml_declaration=True)
        entries["word/media/note-composition.png"] = picture
        types = E.fromstring(entries["[Content_Types].xml"])
        if not any(item.get("Extension") == "png" for item in types):
            E.SubElement(types, ct + "Default", {"Extension": "png", "ContentType": "image/png"})
        entries["[Content_Types].xml"] = E.tostring(types, encoding="utf-8", xml_declaration=True)
        target = output / ("mixed-note-" + kind + "-" + container + ".docx")
        with ZipFile(target, "w", ZIP_DEFLATED) as z:
            for name, payload in sorted(entries.items()):
                info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                z.writestr(info, payload)
        rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
            Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
            SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
            DonorSha256=hashlib.sha256(donor.read_bytes()).hexdigest().upper(),
            TableSourceSha256=hashlib.sha256(table_source.read_bytes()).hexdigest().upper(),
            Kind=kind, Container=container, HeldOut=False, ExpectedEquationCount=2,
            ExpectedImagePlacements=1, ExpectedText=["Public note calculation", "Public destination Alpha.", "Public destination Delta."]))
(output / "mixed-note-containers.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Pinned four direct/table footnote/endnote composition controls")
