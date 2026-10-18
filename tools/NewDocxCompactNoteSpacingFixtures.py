"""Generate bounded public counterfactuals for body/note paragraph boundaries."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib
import json

root = Path(__file__).resolve().parent.parent
cases = root / "tests/Lokad.OoxPdf.Tests/Cases"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
rows = []
sources = [
    ("full", "Composition/composed-operator-matrix-specification.docx",
     "D7B869C53D801B3C861513B0489A5B45DF646751C4938710D1D8408CBBF4614E"),
    ("no-header", "CompositionProbes/operator-matrix-specification-without-header.docx",
     "4B166608273F40A0D94E066EC3B2F3BA75DE71AC5539448F89329E1196166E2B"),
]
for context, source, source_hash in sources:
    assert hashlib.sha256((cases / source).read_bytes()).hexdigest().upper() == source_hash
    with ZipFile(cases / source) as archive:
        original = {name: archive.read(name) for name in archive.namelist()}
    for target, after in [("note", 0), ("note", 6), ("note", 12), ("body", 0)]:
        entries = dict(original)
        part = "word/footnotes.xml" if target == "note" else "word/document.xml"
        tree = E.fromstring(entries[part])
        paragraph = (next(note for note in tree.findall(w + "footnote")
                          if note.get(w + "id") == "1").find(w + "p") if target == "note"
                     else tree.find(w + "body").findall(w + "p")[-1])
        assert paragraph is not None
        properties = paragraph.find(w + "pPr")
        if properties is None:
            properties = E.Element(w + "pPr")
            paragraph.insert(0, properties)
        spacing = properties.find(w + "spacing")
        if spacing is None:
            spacing = E.SubElement(properties, w + "spacing")
        spacing.set(w + "after", str(after * 20))
        entries[part] = E.tostring(tree, encoding="utf-8", xml_declaration=True)
        name = f"compact-note-{context}-{target}-after{after}"
        relative = f"CompositionProbes/{name}.docx"
        with ZipFile(cases / relative, "w", ZIP_DEFLATED) as archive:
            for path, data in sorted(entries.items()):
                info = ZipInfo(path, (2026, 1, 1, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                archive.writestr(info, data)
        rows.append(dict(Id=name, Input="tests/Lokad.OoxPdf.Tests/Cases/" + relative,
            Sha256=hashlib.sha256((cases / relative).read_bytes()).hexdigest().upper(),
            SourceSha256=source_hash, HeldOut=False, ExpectedEquationCount=6,
            ExpectedText=["Review conclusion:", "All quantities in this synthetic example are illustrative."],
            Context=context, ChangedTarget=target, AfterPoints=after))
(cases / "CompositionProbes/compact-note-spacing.json").write_text(
    json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Generated eight public body/note spacing controls; existing inputs unchanged.")
