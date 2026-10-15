"""Independent inline binary-break setting controls; no Office dependency."""
import hashlib
import json
import argparse
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET

parser = argparse.ArgumentParser()
parser.add_argument("--story", choices=("body", "cell", "header", "footer"), default="body")
args = parser.parse_args()

repo = Path(__file__).resolve().parents[1]
cases = repo / "tests/Lokad.OoxPdf.Tests/Cases"
source = cases / "docx-omml-inline-overfull12.docx"
source_sha = hashlib.sha256(source.read_bytes()).hexdigest().upper()
assert source_sha == "45F27A6285C1A065D5F4BE435AD335FC9F96E8341AB3BDC1D880BBD6DE2D8867"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
ET.register_namespace("w", w[1:-1])
ET.register_namespace("m", m[1:-1])
for setting in (("default", "before", "after", "repeat") if args.story == "body" else ("default",)):
    for upright in (False, True):
        identifier = "math-literal-inline-break-" + setting + ("-" + args.story if args.story != "body" else "") + ("-upright" if upright else "")
        output = cases / ("PortableFonts" if upright else "CompositionProbes") / (identifier + ".docx")
        manifest = output.with_suffix(".json")
        if output.exists():
            pin = json.loads(manifest.read_text(encoding="utf-8"))[0]
            assert pin["SourceSha256"] == source_sha
            assert pin["Sha256"] == hashlib.sha256(output.read_bytes()).hexdigest().upper()
            continue
        with ZipFile(source) as parent:
            entries = {name: parent.read(name) for name in parent.namelist()}
        for name, data in list(entries.items()):
            if name == "word/settings.xml" and setting != "default":
                document = ET.fromstring(data)
                ET.SubElement(document.find(m + "mathPr"), m + "brkBin", {m + "val": setting})
                data = ET.tostring(document, encoding="utf-8", xml_declaration=True)
            if name == "word/document.xml" and upright:
                document = ET.fromstring(data)
                for run in document.iter(m + "r"):
                    ET.SubElement(ET.SubElement(run, m + "rPr"), m + "sty", {m + "val": "p"})
                data = ET.tostring(document, encoding="utf-8", xml_declaration=True)
            entries[name] = data
        if args.story != "body":
            document = ET.fromstring(entries["word/document.xml"])
            body = document.find(w + "body")
            paragraphs = [deepcopy(p) for p in body.findall(w + "p")]
            section = deepcopy(body.find(w + "sectPr"))
            body.clear()
            if args.story == "cell":
                table = ET.SubElement(body, w + "tbl")
                props = ET.SubElement(table, w + "tblPr")
                ET.SubElement(props, w + "tblW", {w + "w": "2200", w + "type": "dxa"})
                ET.SubElement(props, w + "tblLayout", {w + "type": "fixed"})
                ET.SubElement(ET.SubElement(table, w + "tblGrid"), w + "gridCol", {w + "w": "2200"})
                cell = ET.SubElement(ET.SubElement(table, w + "tr"), w + "tc")
                ET.SubElement(ET.SubElement(cell, w + "tcPr"), w + "tcW", {w + "w": "2200", w + "type": "dxa"})
                cell.extend(paragraphs)
                ET.SubElement(body, w + "p")
            else:
                part = args.story + "-math-break.xml"
                story = ET.Element(w + ("hdr" if args.story == "header" else "ftr"))
                story.extend(paragraphs)
                entries["word/" + part] = ET.tostring(story, encoding="utf-8", xml_declaration=True)
                r = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
                ET.SubElement(section, w + args.story + "Reference", {w + "type": "default", r + "id": "mathStory"})
                package = "{http://schemas.openxmlformats.org/package/2006/relationships}"
                relationships = ET.fromstring(entries["word/_rels/document.xml.rels"])
                ET.SubElement(relationships, package + "Relationship", {"Id": "mathStory", "Type": r[1:-1] + "/" + args.story, "Target": part})
                entries["word/_rels/document.xml.rels"] = ET.tostring(relationships, encoding="utf-8", xml_declaration=True)
                types = ET.fromstring(entries["[Content_Types].xml"])
                ET.SubElement(types, "{http://schemas.openxmlformats.org/package/2006/content-types}Override", {
                    "PartName": "/word/" + part, "ContentType": "application/vnd.openxmlformats-officedocument.wordprocessingml." + args.story + "+xml"})
                entries["[Content_Types].xml"] = ET.tostring(types, encoding="utf-8", xml_declaration=True)
                ET.SubElement(ET.SubElement(ET.SubElement(body, w + "p"), w + "r"), w + "t").text = "Body after static calculation."
            body.append(section)
            entries["word/document.xml"] = ET.tostring(document, encoding="utf-8", xml_declaration=True)
        with ZipFile(output, "w", ZIP_DEFLATED) as result:
            for name, data in entries.items():
                result.writestr(name, data)
        manifest.write_text(json.dumps([{
            "Id": identifier, "Input": output.relative_to(repo).as_posix(), "SourceSha256": source_sha,
            "Sha256": hashlib.sha256(output.read_bytes()).hexdigest().upper(),
            "HeldOut": False, "ExpectedEquationCount": 1, "ExpectedInlineImageCount": 0,
            "ExpectedText": ["Before formula:", "after formula.", "Following prose."]
        }], indent=2) + "\n", encoding="utf-8")
