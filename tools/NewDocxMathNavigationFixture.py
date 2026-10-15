"""Author a public equation/bookmark/link composition control from a pinned source."""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import hashlib
import json
import xml.etree.ElementTree as ET

repo = Path(__file__).resolve().parents[1]
cases = repo / "tests/Lokad.OoxPdf.Tests/Cases"
source = cases / "Composition/report.docx"
source_sha = hashlib.sha256(source.read_bytes()).hexdigest().upper()
assert source_sha == "10F63D162D0D46DBE8CD682BB50452DB6672104706F116EA50CDCAFEE054575F"
output = cases / "CompositionProbes/math-navigation.docx"
manifest = output.with_suffix(".json")
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
ET.register_namespace("w", w[1:-1])
ET.register_namespace("m", m[1:-1])


def normal_math_copy():
    normal = cases / "PortableFonts/math-navigation-normal.docx"
    normal_manifest = normal.with_suffix(".json")
    parent_sha = hashlib.sha256(output.read_bytes()).hexdigest().upper()
    if normal.exists():
        pin = json.loads(normal_manifest.read_text(encoding="utf-8"))
        assert pin["SourceSha256"] == parent_sha
        assert pin["Sha256"] == hashlib.sha256(normal.read_bytes()).hexdigest().upper()
        return
    with ZipFile(output) as parent, ZipFile(normal, "w", ZIP_DEFLATED) as result:
        for entry in parent.infolist():
            data = parent.read(entry)
            if entry.filename.startswith("word/") and entry.filename.endswith(".xml"):
                doc = ET.fromstring(data)
                changed = False
                for run in doc.iter(m + "r"):
                    props = run.find(m + "rPr")
                    if props is None:
                        props = ET.Element(m + "rPr")
                        run.insert(0, props)
                    nor = props.find(m + "nor")
                    if nor is None:
                        nor = ET.SubElement(props, m + "nor")
                    nor.set(m + "val", "1")
                    changed = True
                if changed:
                    data = ET.tostring(doc, encoding="utf-8", xml_declaration=True)
            result.writestr(entry, data)
    normal_manifest.write_text(json.dumps({"Input": normal.relative_to(repo).as_posix(),
        "Sha256": hashlib.sha256(normal.read_bytes()).hexdigest().upper(), "SourceSha256": parent_sha,
        "Transformation": "Add m:rPr/m:nor m:val=1 to every mathematical run for the pinned portable face."}, indent=2) + "\n", encoding="utf-8")


if output.exists():
    pin = json.loads(manifest.read_text(encoding="utf-8"))[0]
    assert pin["SourceSha256"] == source_sha
    assert hashlib.sha256(output.read_bytes()).hexdigest().upper() == pin["Sha256"]
    normal_math_copy()
    raise SystemExit(0)
with ZipFile(source) as template:
    main = ET.fromstring(template.read("word/document.xml"))
    formula = next(element for element in main.iter(m + "oMath") if element.find(m + "f") is not None)

    def target(name, identifier):
        paragraph = ET.Element(w + "p")
        ET.SubElement(paragraph, w + "bookmarkStart", {w + "id": str(identifier), w + "name": name})
        paragraph.append(ET.fromstring(ET.tostring(formula)))
        ET.SubElement(paragraph, w + "bookmarkEnd", {w + "id": str(identifier)})
        return paragraph

    with ZipFile(output, "w", ZIP_DEFLATED) as result:
        for entry in template.infolist():
            data = template.read(entry)
            if entry.filename == "word/document.xml":
                body = main.find(w + "body")
                section = body.find(w + "sectPr")
                table = body.find(w + "tbl")
                cell_paragraph = list(table.iter(w + "tc"))[-1].find(w + "p")
                cell_paragraph.clear()
                for child in target("MathCell", 42):
                    cell_paragraph.append(child)
                body.clear()
                for index, name in enumerate(["Body", "Cell", "Header", "Footer"]):
                    paragraph = ET.SubElement(body, w + "p")
                    if index == 0:
                        ET.SubElement(paragraph, w + "commentRangeStart", {w + "id": "0"})
                    link = ET.SubElement(paragraph, w + "hyperlink", {w + "anchor": "Math" + name})
                    run = ET.SubElement(link, w + "r")
                    ET.SubElement(run, w + "t").text = "Go to " + name.lower() + " calculation"
                    if index == 0:
                        ET.SubElement(paragraph, w + "commentRangeEnd", {w + "id": "0"})
                        ET.SubElement(ET.SubElement(paragraph, w + "r"), w + "commentReference", {w + "id": "0"})
                body.append(target("MathBody", 41))
                body.append(table)
                body.append(section)
                data = ET.tostring(main, encoding="utf-8", xml_declaration=True)
            elif entry.filename in ["word/header1.xml", "word/footer1.xml"]:
                doc = ET.fromstring(data)
                doc.clear()
                header = entry.filename == "word/header1.xml"
                doc.append(target("MathHeader" if header else "MathFooter", 43 if header else 44))
                data = ET.tostring(doc, encoding="utf-8", xml_declaration=True)
            result.writestr(entry, data)
manifest.write_text(json.dumps([{"Id": "math-navigation", "Input": output.relative_to(repo).as_posix(),
    "Sha256": hashlib.sha256(output.read_bytes()).hexdigest().upper(), "SourceSha256": source_sha,
    "ExpectedEquationCount": 4, "ExpectedInlineImageCount": 0, "HeldOut": False,
    "ExpectedText": ["Go to " + name + " calculation" for name in ["body", "cell", "header", "footer"]]}], indent=2) + "\n", encoding="utf-8")
normal_math_copy()
