"""Generate bounded plain-note mutations of the linked footer report."""
from pathlib import Path
from xml.etree import ElementTree as E
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from io import BytesIO
import copy
import hashlib
import json

root = Path(__file__).resolve().parent.parent
output = root / "tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
manifest = json.loads((output / "occupied-footer-continuations.json").read_text("utf-8"))
for row in manifest:
    restore_footer = row.get("RestoreOriginalFooter", False)
    split_owners = row.get("SplitNoteOwners", False)
    table_reference = row.get("MoveReferenceToTable", False)
    remove_separator = row.get("RemoveFootnoteSeparator", False)
    note_picture = row.get("AddInlineNotePicture", False)
    wrap_picture = row.get("WrapNotePictureParagraph", False)
    second_picture = row.get("AddSecondNotePicture", False)
    note_equation = row.get("AddInlineNoteEquation", False)
    note_minimum = row.get("NoteMinimumPoints")
    link_equation = row.get("LinkNoteEquation", False)
    source = output / ("mixed-math-picture-note-report.docx" if note_minimum is not None or link_equation else
                       "long-note-inline-picture-report.docx" if wrap_picture or second_picture or note_equation else
                       "two-plain-notes-report.docx" if note_picture else
                       "table-note-reference-report.docx" if remove_separator else
                       "two-plain-notes-report.docx" if table_reference else
                       "long-note-original-footer-report.docx" if split_owners else
                       "occupied-footer-long-note-report.docx" if restore_footer else "linked-inline-footer-report.docx")
    assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == row["SourceSha256"]
    with ZipFile(source) as archive:
        original = {name: archive.read(name) for name in archive.namelist()}
    parts = dict(original)
    if link_equation:
        part = "word/footnotes.xml"
        story = E.fromstring(parts[part])
        intro = next(n for n in story if n.get(w + "id") == "2").find(w + "p")
        math = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
        equation = next(x for x in intro if x.tag == math + "oMath")
        index = list(intro).index(equation)
        intro.remove(equation)
        link = E.Element(w + "hyperlink", {w + "anchor": "Calculation"})
        link.append(equation)
        intro.insert(index, link)
    elif note_minimum is not None:
        part = "word/footnotes.xml"
        story = E.fromstring(parts[part])
        intro = next(n for n in story if n.get(w + "id") == "2").find(w + "p")
        assert intro.find(w + "pPr") is None
        properties = E.Element(w + "pPr")
        intro.insert(0, properties)
        E.SubElement(properties, w + "spacing", {w + "line": str(int(note_minimum * 20)), w + "lineRule": "atLeast"})
    elif wrap_picture or second_picture or note_equation:
        part = "word/footnotes.xml"
        story = E.fromstring(parts[part])
        note = next(n for n in story if n.get(w + "id") == "2")
        intro = note.find(w + "p")
        image = next(run for run in intro.findall(w + "r") if run.find(".//" + w + "drawing") is not None)
        if wrap_picture:
            run = E.Element(w + "r")
            E.SubElement(run, w + "t", {"{http://www.w3.org/XML/1998/namespace}space": "preserve"}).text = (
                " " + "The continuation picture must stay with its own wrapped source line. " * 8)
            intro.insert(list(intro).index(image), run)
        if second_picture:
            clone = copy.deepcopy(image)
            docpr = next(clone.iter("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr"))
            docpr.set("id", "502")
            docpr.set("name", "Public wrapped note picture")
            next(p for p in note.findall(w + "p") if "Public continuation note line 50:" in "".join(p.itertext())).append(clone)
        if note_equation:
            math = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
            equation = copy.deepcopy(next(E.fromstring(parts["word/document.xml"]).iter(math + "oMath")))
            E.SubElement(E.SubElement(intro, w + "r"), w + "t",
                         {"{http://www.w3.org/XML/1998/namespace}space": "preserve"}).text = " "
            intro.append(equation)
    elif note_picture:
        part = "word/footnotes.xml"
        a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
        r = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
        wp = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
        pr = "{http://schemas.openxmlformats.org/package/2006/relationships}"
        body = E.fromstring(parts["word/document.xml"])
        drawing = copy.deepcopy(next(body.iter(w + "drawing")))
        blip = next(drawing.iter(a + "blip"))
        relationship = next(rel for rel in E.fromstring(parts["word/_rels/document.xml.rels"])
                            if rel.get("Id") == blip.get(r + "embed"))
        blip.set(r + "embed", "rIdPublicNotePicture")
        for docpr in drawing.iter(wp + "docPr"):
            docpr.set("id", "501")
            docpr.set("name", "Public continuation note picture")
        story = E.fromstring(parts[part])
        intro = next(note for note in story if note.get(w + "id") == "2").find(w + "p")
        E.SubElement(intro, w + "r").append(drawing)
        relpart = "word/_rels/footnotes.xml.rels"
        assert relpart not in parts
        rels = E.Element(pr + "Relationships")
        attrs = dict(relationship.attrib)
        attrs["Id"] = "rIdPublicNotePicture"
        E.SubElement(rels, pr + "Relationship", attrs)
        parts[relpart] = E.tostring(rels, encoding="utf-8", xml_declaration=True)
    elif table_reference:
        part = "word/document.xml"
        story = E.fromstring(parts[part])
        body = story.find(w + "body")
        paragraph = next(p for p in body.findall(w + "p") if any(
            ref.get(w + "id") == "2" for ref in p.iter(w + "footnoteReference")))
        marker = next(run for run in paragraph if any(
            ref.get(w + "id") == "2" for ref in run.iter(w + "footnoteReference")))
        label = paragraph[list(paragraph).index(marker) - 1]
        assert "".join(label.itertext()) == " Additional note"
        target = body.find(w + "tbl").findall(w + "tr")[-1].findall(w + "tc")[-1].findall(w + "p")[-1]
        for run in (label, marker):
            paragraph.remove(run)
            target.append(run)
        assert [ref.get(w + "id") for ref in story.iter(w + "footnoteReference")] == ["2", "1"]
    elif split_owners:
        part = "word/footnotes.xml"
        story = E.fromstring(parts[part])
        first = next(n for n in story if n.get(w + "id") == "1")
        paragraphs = list(first)
        assert len(paragraphs) == 65
        second = E.Element(w + "footnote", {w + "id": "2"})
        paragraph = E.SubElement(second, w + "p")
        E.SubElement(E.SubElement(paragraph, w + "r"), w + "t").text = "Second public continuation note."
        for paragraph in paragraphs[33:]:
            first.remove(paragraph)
            second.append(paragraph)
        story.append(second)
        body = E.fromstring(parts["word/document.xml"])
        paragraph = next(p for p in body.iter(w + "p") if list(p.iter(w + "footnoteReference")))
        run = E.SubElement(paragraph, w + "r")
        E.SubElement(run, w + "t", {"{http://www.w3.org/XML/1998/namespace}space": "preserve"}).text = " Additional note"
        E.SubElement(E.SubElement(paragraph, w + "r"), w + "footnoteReference", {w + "id": "2"})
        parts["word/document.xml"] = E.tostring(body, encoding="utf-8", xml_declaration=True)
    elif restore_footer:
        part = "word/footer1.xml"
        story = E.fromstring(parts[part])
        story.remove(story[-1])
        with ZipFile(root / "tests/Lokad.OoxPdf.Tests/Cases/Composition/report.docx") as archive:
            assert E.tostring(story) == E.tostring(E.fromstring(archive.read(part)))
    else:
        part = "word/footnotes.xml"
        story = E.fromstring(parts[part])
        if remove_separator or "no-separator" in row["Id"]:
            story.remove(next(n for n in story if n.get(w + "type") == "separator"))
        else:
            note = next(n for n in story if n.get(w + "id") == "1")
            for number in range(1, 65):
                paragraph = E.SubElement(note, w + "p")
                run = E.SubElement(paragraph, w + "r")
                E.SubElement(run, w + "t", {"{http://www.w3.org/XML/1998/namespace}space": "preserve"}).text = (
                    f"Public continuation note line {number:02}: the footer must keep every source paragraph clear.")
    parts[part] = E.tostring(story, encoding="utf-8", xml_declaration=True)
    assert set(name for name in original if parts[name] != original[name]) == ({part, "word/document.xml"} if split_owners else {part})
    buffer = BytesIO()
    with ZipFile(buffer, "w", ZIP_DEFLATED) as archive:
        for name, data in sorted(parts.items()):
            info = ZipInfo(name, (2026, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    data = buffer.getvalue()
    assert hashlib.sha256(data).hexdigest().upper() == row["Sha256"], "Preserve qualified public sources"
    (root / row["Input"]).write_bytes(data)
print(f"{len(manifest)} note controls regenerated byte-identically; only declared story/reference parts change.")
