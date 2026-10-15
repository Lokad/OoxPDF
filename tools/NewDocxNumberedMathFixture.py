"""Pin numbered and unnumbered composition controls from the public report."""
from copy import deepcopy
import argparse
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET

parser = argparse.ArgumentParser()
parser.add_argument("--story", choices=("body", "cell", "header", "footer"), default="body")
parser.add_argument("--normal-math", action="store_true")
args = parser.parse_args()

repo = Path(__file__).resolve().parents[1]
cases = repo / "tests/Lokad.OoxPdf.Tests/Cases"
source = cases / "Composition/report.docx"
source_sha = hashlib.sha256(source.read_bytes()).hexdigest().upper()
assert source_sha == "10F63D162D0D46DBE8CD682BB50452DB6672104706F116EA50CDCAFEE054575F"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
ET.register_namespace("w", w[1:-1])
for numbered in (False, True):
    identifier = "math-composition-" + ("numbered" if numbered else "unnumbered")
    identifier += ("-" + args.story if args.story != "body" else "") + ("-normal" if args.normal_math else "")
    output = cases / ("PortableFonts" if args.normal_math else "CompositionProbes") / (identifier + ".docx")
    manifest = output.with_suffix(".json")
    if output.exists():
        pin = json.loads(manifest.read_text(encoding="utf-8"))[0]
        assert pin["SourceSha256"] == source_sha
        assert pin["Sha256"] == hashlib.sha256(output.read_bytes()).hexdigest().upper()
        continue
    with ZipFile(source) as parent, ZipFile(output, "w", ZIP_DEFLATED) as result:
        story_paragraphs = []
        source_relationships = parent.read("word/_rels/document.xml.rels")
        for entry in parent.infolist():
            data = parent.read(entry)
            if entry.filename == "word/document.xml":
                document = ET.fromstring(data)
                body = document.find(w + "body")
                section = deepcopy(body.find(w + "sectPr"))
                for child in list(section):
                    if child.tag in (w + "headerReference", w + "footerReference"):
                        section.remove(child)
                paragraphs = [deepcopy(p) for p in body.findall(w + "p") if p.find(".//" + m + "oMath") is not None]
                assert len(paragraphs) == 3
                if args.normal_math:
                    for paragraph in paragraphs:
                        for run in paragraph.iter(m + "r"):
                            props = run.find(m + "rPr")
                            if props is None:
                                props = ET.Element(m + "rPr")
                                run.insert(0, props)
                            ET.SubElement(props, m + "nor", {m + "val": "1"})
                body.clear()
                for paragraph in paragraphs:
                    if numbered:
                        properties = paragraph.find(w + "pPr")
                        if properties is None:
                            properties = ET.Element(w + "pPr")
                            paragraph.insert(0, properties)
                        number = ET.SubElement(properties, w + "numPr")
                        ET.SubElement(number, w + "ilvl", {w + "val": "0"})
                        ET.SubElement(number, w + "numId", {w + "val": "1"})
                if args.story == "body":
                    body.extend(paragraphs)
                elif args.story == "cell":
                    table = ET.SubElement(body, w + "tbl")
                    grid = ET.SubElement(table, w + "tblGrid")
                    ET.SubElement(grid, w + "gridCol", {w + "w": "9000"})
                    cell = ET.SubElement(ET.SubElement(table, w + "tr"), w + "tc")
                    props = ET.SubElement(cell, w + "tcPr")
                    ET.SubElement(props, w + "tcW", {w + "w": "9000", w + "type": "dxa"})
                    cell.extend(paragraphs)
                else:
                    story_paragraphs = paragraphs
                    ET.SubElement(ET.SubElement(ET.SubElement(body, w + "p"), w + "r"), w + "t").text = "Composed story."
                    relation = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
                    ET.SubElement(section, w + args.story + "Reference", {w + "type": "default", relation + "id": args.story})
                body.append(section)
                data = ET.tostring(document, encoding="utf-8", xml_declaration=True)
            if args.story in ("header", "footer") and entry.filename == "word/" + args.story + "1.xml":
                document = ET.fromstring(data)
                document.clear()
                document.extend(story_paragraphs)
                data = ET.tostring(document, encoding="utf-8", xml_declaration=True)
            result.writestr(entry, data)
        if args.story in ("header", "footer"):
            relations = ET.fromstring(source_relationships)
            for child in list(relations):
                if not any(child.attrib["Type"].endswith("/" + kind) for kind in ("image", "hyperlink")):
                    relations.remove(child)
            result.writestr("word/_rels/" + args.story + "1.xml.rels", ET.tostring(relations, encoding="utf-8", xml_declaration=True))
    manifest.write_text(json.dumps([{
        "Id": identifier, "Input": output.relative_to(repo).as_posix(),
        "SourceSha256": source_sha, "Sha256": hashlib.sha256(output.read_bytes()).hexdigest().upper(),
        "HeldOut": False, "ExpectedEquationCount": 4, "ExpectedInlineImageCount": 1,
        "ExpectedText": ["Summary:", "share the same input series.", "The trend marker"]
    }], indent=2) + "\n", encoding="utf-8")
