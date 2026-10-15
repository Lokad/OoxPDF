"""Pin multiple-equation bookmark controls, including an image and a line break."""
import argparse
from copy import deepcopy
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET

parser = argparse.ArgumentParser()
parser.add_argument("--mixed-story", choices=["body", "cell", "header", "footer"], default="body")
args = parser.parse_args()
suffix = "" if args.mixed_story == "body" else "-" + args.mixed_story
repo = Path(__file__).resolve().parents[1]
cases = repo / "tests/Lokad.OoxPdf.Tests/Cases"
source = cases / "CompositionProbes/math-navigation.docx"
source_sha = hashlib.sha256(source.read_bytes()).hexdigest().upper()
assert source_sha == "34F2CE8E3FC2612244DD4F114277318B1EB337A7835738D1CB813D10D2F13ACC"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
ET.register_namespace("w", w[1:-1])
ET.register_namespace("m", m[1:-1])
with ZipFile(cases / "Composition/report.docx") as report:
    picture = next(ET.fromstring(report.read("word/document.xml")).iter(w + "drawing"))
    r = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    embed = picture.find(".//*[@" + r + "embed]").get(r + "embed")
    image_target = next(item.get("Target") for item in ET.fromstring(report.read("word/_rels/document.xml.rels"))
                        if item.get("Id") == embed)

for portable in (False, True):
    output = cases / (f"PortableFonts/inline-math-navigation-normal{suffix}.docx" if portable else
                      f"CompositionProbes/inline-math-navigation{suffix}.docx")
    manifest = output.with_suffix(".json")
    if output.exists():
        pin = json.loads(manifest.read_text(encoding="utf-8"))[0]
        assert pin["SourceSha256"] == source_sha
        assert pin["Sha256"] == hashlib.sha256(output.read_bytes()).hexdigest().upper()
        continue
    with ZipFile(source) as parent, ZipFile(output, "w", ZIP_DEFLATED) as result:
        for entry in parent.infolist():
            data = parent.read(entry)
            if entry.filename in ("word/document.xml", "word/header1.xml", "word/footer1.xml"):
                doc = ET.fromstring(data)
                for paragraph in list(doc.iter(w + "p")):
                    anchor = paragraph.find(w + "bookmarkStart")
                    formula = paragraph.find(m + "oMath")
                    if anchor is None or formula is None:
                        continue
                    name, identifier = anchor.get(w + "name"), int(anchor.get(w + "id"))
                    paragraph.clear()
                    for index in (1, 2):
                        ET.SubElement(paragraph, w + "bookmarkStart", {
                            w + "id": str(identifier * 10 + index), w + "name": name + str(index)})
                        paragraph.append(deepcopy(formula))
                        ET.SubElement(paragraph, w + "bookmarkEnd", {w + "id": str(identifier * 10 + index)})
                        if index == 1 and name == "Math" + args.mixed_story.capitalize():
                            paragraph.append(ET.Element(w + "r"))
                            drawing = deepcopy(picture)
                            if args.mixed_story in ("header", "footer"):
                                r = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
                                drawing.find(".//*[@" + r + "embed]").set(r + "embed", "rIdCompositionImage")
                            paragraph[-1].append(drawing)
                            ET.SubElement(ET.SubElement(paragraph, w + "r"), w + "br")
                for paragraph in list(doc.iter(w + "p")):
                    link = paragraph.find(w + "hyperlink")
                    if link is None:
                        continue
                    name = link.get(w + "anchor")
                    if not name or not name.startswith("Math"):
                        continue
                    for index in (1, 2):
                        copy = deepcopy(link)
                        copy.set(w + "anchor", name + str(index))
                        next(copy.iter(w + "t")).text += " " + str(index)
                        paragraph.insert(list(paragraph).index(link), copy)
                    paragraph.remove(link)
                if portable:
                    for run in doc.iter(m + "r"):
                        props = run.find(m + "rPr")
                        if props is None:
                            props = ET.Element(m + "rPr")
                            run.insert(0, props)
                        ET.SubElement(props, m + "nor", {m + "val": "1"})
                data = ET.tostring(doc, encoding="utf-8", xml_declaration=True)
            result.writestr(entry, data)
        if args.mixed_story in ("header", "footer"):
            rel = "{http://schemas.openxmlformats.org/package/2006/relationships}"
            root = ET.Element(rel + "Relationships")
            ET.SubElement(root, rel + "Relationship", {
                "Id": "rIdCompositionImage", "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                "Target": image_target})
            name = "word/_rels/" + args.mixed_story + "1.xml.rels"
            assert name not in parent.namelist(), "Preserve existing part relationships before extending this authoring control"
            result.writestr(name, ET.tostring(root, encoding="utf-8", xml_declaration=True))
    manifest.write_text(json.dumps([{
        "Id": ("inline-math-navigation-normal" if portable else "inline-math-navigation") + suffix,
        "Input": output.relative_to(repo).as_posix(),
        "Sha256": hashlib.sha256(output.read_bytes()).hexdigest().upper(),
        "SourceSha256": source_sha, "HeldOut": False,
        "ExpectedEquationCount": 8, "ExpectedInlineImageCount": 1,
        "ExpectedText": ["Go to " + story + " calculation " + str(index)
                         for story in ("body", "cell", "header", "footer") for index in (1, 2)]
    }], indent=2) + "\n", encoding="utf-8")
