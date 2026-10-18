"""Create plain minimum-height static stories, independent of note holdouts."""
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
r = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
rel = "{http://schemas.openxmlformats.org/package/2006/relationships}"
ct = "{http://schemas.openxmlformats.org/package/2006/content-types}"
with ZipFile(source) as archive:
    original = {name: archive.read(name) for name in archive.namelist() if name != "word/footnotes.xml"}
document = E.fromstring(original["word/document.xml"])
body = document.find(w + "body")
body.remove(next(p for p in body if p.find(".//" + w + "footnoteReference") is not None))
original["word/document.xml"] = E.tostring(document, encoding="utf-8", xml_declaration=True)
relationships = E.fromstring(original["word/_rels/document.xml.rels"])
for relationship in list(relationships):
    if relationship.get("Type", "").endswith("/footnotes"):
        relationships.remove(relationship)
original["word/_rels/document.xml.rels"] = E.tostring(relationships, encoding="utf-8", xml_declaration=True)
types = E.fromstring(original["[Content_Types].xml"])
for item in list(types):
    if item.get("PartName") == "/word/footnotes.xml":
        types.remove(item)
original["[Content_Types].xml"] = E.tostring(types, encoding="utf-8", xml_declaration=True)
rows = []
for kind in ("header", "footer"):
    parts = dict(original)
    document = E.fromstring(parts["word/document.xml"])
    section = document.find(".//" + w + "sectPr")
    section.insert(0, E.Element(w + kind + "Reference", {w + "type": "default", r + "id": "rIdMinimumStatic"}))
    section.find(w + "pgMar").set(w + kind, "720")
    parts["word/document.xml"] = E.tostring(document, encoding="utf-8", xml_declaration=True)
    part = "word/" + kind + "MinimumPlain.xml"
    relationships = E.fromstring(parts["word/_rels/document.xml.rels"])
    E.SubElement(relationships, rel + "Relationship", dict(Id="rIdMinimumStatic",
        Type=r[1:-1] + "/" + kind, Target=part.split("/")[-1]))
    parts["word/_rels/document.xml.rels"] = E.tostring(relationships, encoding="utf-8", xml_declaration=True)
    types = E.fromstring(parts["[Content_Types].xml"])
    E.SubElement(types, ct + "Override", dict(PartName="/" + part,
        ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml." + kind + "+xml"))
    parts["[Content_Types].xml"] = E.tostring(types, encoding="utf-8", xml_declaration=True)
    story = E.Element(w + ("hdr" if kind == "header" else "ftr"))
    expected = ["Leading paragraph.", "Following prose."]
    for i in range(1, 4):
        p = E.SubElement(story, w + "p")
        properties = E.SubElement(p, w + "pPr")
        E.SubElement(properties, w + "spacing", {w + "before": "0", w + "after": "0", w + "line": "400", w + "lineRule": "atLeast"})
        run = E.SubElement(p, w + "r")
        properties = E.SubElement(run, w + "rPr")
        E.SubElement(properties, w + "rFonts", {w + "ascii": "Arial", w + "hAnsi": "Arial"})
        E.SubElement(properties, w + "sz", {w + "val": "20"})
        text = "Public review " + kind + " line " + str(i) + "."
        E.SubElement(run, w + "t").text = text
        expected.append(text)
    parts[part] = E.tostring(story, encoding="utf-8", xml_declaration=True)
    name = "minimum-static-" + kind
    destination = output / (name + ".docx")
    with ZipFile(destination, "w", ZIP_DEFLATED) as archive:
        for path, data in sorted(parts.items()):
            info = ZipInfo(path, (2026, 1, 1, 0, 0, 0))
            info.create_system = 0
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    rows.append(dict(Id=name, Input=destination.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(destination.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        HeldOut=False, ExpectedEquationCount=0, ExpectedInlineImageCount=0,
        ExpectedText=expected, Part=part, MinimumLineHeight=20))
(output / "minimum-static-frames.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Two minimum-height plain static-story controls generated.")
