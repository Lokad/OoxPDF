"""Two-section body/cell equations with source-resolved footnote/endnote labels."""
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
E.register_namespace("w", w[1:-1])
E.register_namespace("m", m[1:-1])
rows = []
for kind in ["footnote", "endnote"]:
    for container in ["body", "cell"]:
        source = output / ("math-" + kind + "-before-" + container +
                           ("-plain" if kind == "footnote" else "") + ".docx")
        with ZipFile(source) as z:
            entries = {name: z.read(name) for name in z.namelist()}
        xml = E.fromstring(entries["word/document.xml"])
        for parent in xml.iter():
            for child in list(parent):
                if child.tag in [w + "commentRangeStart", w + "commentRangeEnd", w + "commentReference"]:
                    parent.remove(child)
        body = xml.find(w + "body")
        section = deepcopy(body.find(w + "sectPr"))
        content = next(child for child in body if any(True for _ in child.iter(m + "oMath")))
        notes = E.fromstring(entries["word/" + kind + "s.xml"])
        note = deepcopy(next(child for child in notes if child.get(w + "id") == "1"))
        for child in list(notes):
            if child.get(w + "type") is None:
                notes.remove(child)
        body.clear()
        for section_index in range(2):
            settings = deepcopy(section)
            for child in list(settings):
                if child.tag in [w + kind + "Pr", w + "type"]:
                    settings.remove(child)
            properties = E.SubElement(settings, w + kind + "Pr")
            E.SubElement(properties, w + "numRestart", {w + "val": "eachSect"})
            E.SubElement(properties, w + "numStart", {w + "val": "3" if section_index == 0 else "7"})
            E.SubElement(properties, w + "numFmt", {w + "val": "decimal" if section_index == 0 else "lowerRoman"})
            E.SubElement(settings, w + "type", {w + "val": "nextPage"})
            for within in range(2):
                index = section_index * 2 + within + 1
                block = deepcopy(content)
                marker = next(block.iter(w + kind + "Reference"))
                marker.set(w + "id", str(index))
                body.append(block)
                destination = deepcopy(note)
                destination.set(w + "id", str(index))
                next(destination.iter(w + "t")).text = "Public destination " + str(index) + "."
                notes.append(destination)
            if section_index == 0:
                # Section boundaries are direct body paragraphs even when the
                # supported content is in a table.
                paragraph = E.SubElement(body, w + "p")
                E.SubElement(paragraph, w + "pPr").append(settings)
            else:
                body.append(settings)
        entries["word/document.xml"] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
        entries["word/" + kind + "s.xml"] = E.tostring(notes, encoding="utf-8", xml_declaration=True)
        target = output / ("math-section-notes-" + kind + "-" + container + ".docx")
        with ZipFile(target, "w", ZIP_DEFLATED) as z:
            for name, data in sorted(entries.items()):
                info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                z.writestr(info, data)
        rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
            Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
            SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
            Kind=kind, Container=container, HeldOut=False, ExpectedEquationCount=8,
            ExpectedMarkers=["3", "4", "vii", "viii"],
            ExpectedText=["Before formula:"] + ["Public destination " + str(i) + "." for i in range(1, 5)]))
(output / "math-section-notes.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Pinned four two-section note/math composition controls")
