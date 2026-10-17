"""Add supported fractions without changing the existing page-numbering sources."""
from pathlib import Path
from xml.etree import ElementTree as E
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from copy import deepcopy
import hashlib
import posixpath
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
compound_rows = []
explicit_rows = []
mixed_rows = []
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
    # Challenge ownership without adding a production eligibility exception.
    # Existing base outputs above keep their exact bytes and manifests.
    base_sha = hashlib.sha256(target.read_bytes()).hexdigest().upper()
    image_source = output / "math-picture-first-table-single-cell.docx"
    drawing_ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    relation_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    package_ns = "{http://schemas.openxmlformats.org/package/2006/relationships}"
    types_ns = "{http://schemas.openxmlformats.org/package/2006/content-types}"
    with ZipFile(image_source) as image_zip:
        image_xml = E.fromstring(image_zip.read("word/document.xml"))
        picture = next(run for run in image_xml.iter(w + "r") if run.find(w + "drawing") is not None)
        image_id = next(picture.iter(drawing_ns + "blip")).get(relation_ns + "embed")
        image_rel = next(item for item in E.fromstring(image_zip.read("word/_rels/document.xml.rels")) if item.get("Id") == image_id)
        image_bytes = image_zip.read(posixpath.normpath("word/" + image_rel.get("Target")))
    assert image_bytes.startswith(b"\x89PNG")
    entries["word/media/page-note-compound.png"] = image_bytes
    relationships = E.fromstring(entries["word/_rels/document.xml.rels"])
    E.SubElement(relationships, package_ns + "Relationship", dict(Id="pageNoteCompoundImage",
        Type=relation_ns[1:-1] + "/image", Target="media/page-note-compound.png"))
    entries["word/_rels/document.xml.rels"] = E.tostring(relationships, encoding="utf-8", xml_declaration=True)
    types = E.fromstring(entries["[Content_Types].xml"])
    if not any(item.get("Extension") == "png" for item in types):
        E.SubElement(types, types_ns + "Default", dict(Extension="png", ContentType="image/png"))
    entries["[Content_Types].xml"] = E.tostring(types, encoding="utf-8", xml_declaration=True)
    xml = E.fromstring(entries["word/document.xml"])
    index = 0
    for paragraph in xml.iter(w + "p"):
        first = paragraph.find(m + "oMath")
        if first is None:
            continue
        index += 1
        position = list(paragraph).index(first)
        paragraph.insert(position, E.Element(w + "bookmarkStart", {w + "id": str(index), w + "name": "CompoundNote" + str(index)}))
        separator = E.Element(w + "r")
        E.SubElement(separator, w + "t", {"{http://www.w3.org/XML/1998/namespace}space": "preserve"}).text = " and "
        image = deepcopy(picture)
        next(image.iter(drawing_ns + "blip")).set(relation_ns + "embed", "pageNoteCompoundImage")
        for item in image.iter():
            if item.tag.rsplit("}", 1)[-1] in ["docPr", "cNvPr"]:
                item.set("id", str(1000 + index))
        for offset, item in enumerate([separator, image, deepcopy(first), E.Element(w + "bookmarkEnd", {w + "id": str(index)})], 1):
            paragraph.insert(position + 1 + offset, item)
    assert index == 4
    entries["word/document.xml"] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
    compound = output / ("math-page-notes-compound-" + typography + ".docx")
    with ZipFile(compound, "w", ZIP_DEFLATED) as z:
        for name, data in sorted(entries.items()):
            info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, data)
    compound_rows.append(dict(Id=compound.stem, Input=compound.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(compound.read_bytes()).hexdigest().upper(), SourceSha256=base_sha,
        ImageSha256=hashlib.sha256(image_bytes).hexdigest().upper(), HeldOut=False,
        ExpectedEquationCount=8, ImagePlacements=4, ExpectedMarkers=["09", "10", "09", "10"],
        ExpectedText=rows[-1]["ExpectedText"]))
    # Declaring the same page dependency in document settings must not remove
    # otherwise identical supported paragraph content.
    settings = E.fromstring(entries["word/settings.xml"])
    properties = settings.find(w + "footnotePr")
    restart = properties.find(w + "numRestart")
    if restart is None:
        restart = E.SubElement(properties, w + "numRestart")
    restart.set(w + "val", "eachPage")
    entries["word/settings.xml"] = E.tostring(settings, encoding="utf-8", xml_declaration=True)
    explicit = output / ("math-page-notes-explicit-page-" + typography + ".docx")
    with ZipFile(explicit, "w", ZIP_DEFLATED) as z:
        for name, data in sorted(entries.items()):
            info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, data)
    explicit_rows.append(compound_rows[-1] | dict(Id=explicit.stem, Input=explicit.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(explicit.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(compound.read_bytes()).hexdigest().upper()))
    # A custom mark is visible source content, not an automatic counter value.
    xml = E.fromstring(entries["word/document.xml"])
    reference = next(xml.iter(w + "footnoteReference"))
    reference.set(w + "customMarkFollows", "1")
    reference_run = next(run for run in xml.iter(w + "r") if reference in list(run))
    reference_run.insert(list(reference_run).index(reference) + 1, E.Element(w + "t"))
    reference_run[list(reference_run).index(reference) + 1].text = "*"
    notes = E.fromstring(entries["word/footnotes.xml"])
    destination = next(note for note in notes if note.get(w + "id") == reference.get(w + "id"))
    marker = next(destination.iter(w + "footnoteRef"))
    marker_run = next(run for run in destination.iter(w + "r") if marker in list(run))
    position = list(marker_run).index(marker)
    marker_run.remove(marker)
    mark = E.Element(w + "t")
    mark.text = "*"
    marker_run.insert(position, mark)
    entries["word/document.xml"] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
    entries["word/footnotes.xml"] = E.tostring(notes, encoding="utf-8", xml_declaration=True)
    mixed = output / ("math-page-notes-mixed-custom-" + typography + ".docx")
    with ZipFile(mixed, "w", ZIP_DEFLATED) as z:
        for name, data in sorted(entries.items()):
            info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, data)
    mixed_rows.append(explicit_rows[-1] | dict(Id=mixed.stem, Input=mixed.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(mixed.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(explicit.read_bytes()).hexdigest().upper(),
        ExpectedMarkers=["*", "09", "09", "10"]))
(output / "math-page-notes-mixed-custom.json").write_text(json.dumps(mixed_rows, indent=2) + "\n", encoding="utf-8")
(output / "math-page-notes-explicit-page.json").write_text(json.dumps(explicit_rows, indent=2) + "\n", encoding="utf-8")
(output / "math-page-notes-compound.json").write_text(json.dumps(compound_rows, indent=2) + "\n", encoding="utf-8")
(output / "math-page-notes.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Pinned original, compound, explicit and mixed custom page-note controls")
