"""Reproduce fitting and continued minimum-slot notes from a public control."""
from argparse import ArgumentParser
import hashlib
import json
from pathlib import Path
from xml.etree import ElementTree as E
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED

root = Path(__file__).resolve().parent.parent
parser = ArgumentParser()
parser.add_argument("--output-directory", type=Path,
                    default=root / "tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes")
output = parser.parse_args().output_directory.resolve()
output.mkdir(parents=True, exist_ok=True)
source = root / "tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/math-note-content-footnote-wrapped.docx"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
with ZipFile(source) as archive:
    original = {name: archive.read(name) for name in archive.namelist()}
notes = E.fromstring(original["word/footnotes.xml"])
paragraph = next(p for p in notes.iter(w + "p") if p.find(".//" + m + "oMath") is not None)
properties = paragraph.find(w + "pPr")
if properties is None:
    properties = E.Element(w + "pPr")
    paragraph.insert(0, properties)
spacing = properties.find(w + "spacing")
if spacing is None:
    spacing = E.SubElement(properties, w + "spacing")
spacing.set(w + "line", "960")
spacing.set(w + "lineRule", "atLeast")
original["word/footnotes.xml"] = E.tostring(notes, encoding="utf-8", xml_declaration=True)
rows = []
for name, before in (("minimum-note-wrapped", 0), ("minimum-note-page-boundary", 9000)):
    parts = dict(original)
    if before:
        document = E.fromstring(parts["word/document.xml"])
        paragraph = next(p for p in document.iter(w + "p") if p.find(".//" + w + "footnoteReference") is not None)
        properties = paragraph.find(w + "pPr")
        if properties is None:
            properties = E.Element(w + "pPr")
            paragraph.insert(0, properties)
        spacing = properties.find(w + "spacing")
        if spacing is None:
            spacing = E.SubElement(properties, w + "spacing")
        spacing.set(w + "before", str(before))
        parts["word/document.xml"] = E.tostring(document, encoding="utf-8", xml_declaration=True)
    path = output / (name + ".docx")
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        for part, data in sorted(parts.items()):
            info = ZipInfo(part, (2026, 1, 1, 0, 0, 0))
            info.create_system = 0
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    rows.append(dict(Id=name, Input=path.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(path.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        HeldOut=False, ExpectedEquationCount=4, ExpectedInlineImageCount=0,
        ExpectedText=["Public calculation note.", "End of supported note."],
        MinimumLineHeight=48, BodyBeforeSpacing=before / 20))
(output / "owned-note-frames.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Two source-owned note frame controls generated.")
