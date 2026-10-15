"""Pin a small source-inheritance/paragraph-mark matrix before layout changes."""
from copy import deepcopy
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET

repo = Path(__file__).resolve().parents[1]
cases = repo / "tests/Lokad.OoxPdf.Tests/Cases"
source = cases / "docx-omml-inline-overfull12.docx"
source_sha = hashlib.sha256(source.read_bytes()).hexdigest().upper()
assert source_sha == "45F27A6285C1A065D5F4BE435AD335FC9F96E8341AB3BDC1D880BBD6DE2D8867"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
ET.register_namespace("w", w[1:-1])

def run_properties(owner, family, size):
    props = ET.SubElement(owner, w + "rPr")
    ET.SubElement(props, w + "rFonts", {w + "ascii": family, w + "hAnsi": family})
    ET.SubElement(props, w + "sz", {w + "val": str(size * 2)})

for scope in ("bare", "defaults", "normal-unset", "normal-278", "paragraph-defaults", "paragraph-defaults-empty", "normal-278-paragraph-defaults"):
    identifier = "line-spacing-owner-" + scope
    output = cases / "CompositionProbes" / (identifier + ".docx")
    manifest = output.with_suffix(".json")
    if output.exists():
        pin = json.loads(manifest.read_text(encoding="utf-8"))[0]
        assert pin["SourceSha256"] == source_sha
        assert pin["Sha256"] == hashlib.sha256(output.read_bytes()).hexdigest().upper()
        continue
    with ZipFile(source) as parent:
        entries = {name: parent.read(name) for name in parent.namelist()}
    styles = ET.Element(w + "styles")
    if scope != "bare":
        defaults = ET.SubElement(styles, w + "docDefaults")
        run_properties(ET.SubElement(defaults, w + "rPrDefault"), "Calibri", 12)
        if "paragraph-defaults" in scope:
            properties = ET.SubElement(ET.SubElement(defaults, w + "pPrDefault"), w + "pPr")
            if scope != "paragraph-defaults-empty":
                ET.SubElement(properties, w + "spacing", {w + "before": "0", w + "after": "0"})
    if scope.startswith("normal"):
        normal = ET.SubElement(styles, w + "style", {w + "type": "paragraph", w + "default": "1", w + "styleId": "Normal"})
        ET.SubElement(normal, w + "name", {w + "val": "Normal"})
        if scope.startswith("normal-278"):
            props = ET.SubElement(normal, w + "pPr")
            ET.SubElement(props, w + "spacing", {w + "line": "278", w + "lineRule": "auto"})
        run_properties(normal, "Calibri", 12)
    entries["word/styles.xml"] = ET.tostring(styles, encoding="utf-8", xml_declaration=True)
    document = ET.fromstring(entries["word/document.xml"])
    body = document.find(w + "body")
    section = deepcopy(body.find(w + "sectPr"))
    section.find(w + "pgMar").set(w + "right", "1440")
    body.clear()
    for index, (family, size, mark_family, mark_size) in enumerate([
        ("Calibri", 12, None, None), ("Arial", 11, None, None),
        ("Arial", 11, "Arial", 11), ("Arial", 11, "Calibri", 12),
        ("Calibri", 12, "Calibri", 10), ("Calibri", 12, "Calibri", 20)
    ], 1):
        paragraph = ET.SubElement(body, w + "p")
        props = ET.SubElement(paragraph, w + "pPr")
        ET.SubElement(props, w + "spacing", {w + "before": "0", w + "after": "0"})
        if mark_family is not None:
            run_properties(props, mark_family, mark_size)
        run = ET.SubElement(paragraph, w + "r")
        run_properties(run, family, size)
        ET.SubElement(run, w + "t").text = (f"Case {index}. The source mark and inherited paragraph spacing determine the occupied line. "
            "Text and annotations should survive ordinary placement changes, while independently measured line starts remain explicit.")
    body.append(section)
    entries["word/document.xml"] = ET.tostring(document, encoding="utf-8", xml_declaration=True)
    with ZipFile(output, "w", ZIP_DEFLATED) as result:
        for name, data in entries.items():
            result.writestr(name, data)
    manifest.write_text(json.dumps([{
        "Id": identifier, "Input": output.relative_to(repo).as_posix(),
        "SourceSha256": source_sha, "Sha256": hashlib.sha256(output.read_bytes()).hexdigest().upper(),
        "HeldOut": False, "ExpectedEquationCount": 0, "ExpectedInlineImageCount": 0,
        "ExpectedText": ["Case 1.", "Case 6.", "independently measured line starts remain explicit."]
    }], indent=2) + "\n", encoding="utf-8")
