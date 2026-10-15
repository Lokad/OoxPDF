"""Independent inline binary-break setting controls; no Office dependency."""
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
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
ET.register_namespace("w", w[1:-1])
ET.register_namespace("m", m[1:-1])
for setting in ("default", "before", "after", "repeat"):
    for upright in (False, True):
        identifier = "math-literal-inline-break-" + setting + ("-upright" if upright else "")
        output = cases / ("PortableFonts" if upright else "CompositionProbes") / (identifier + ".docx")
        manifest = output.with_suffix(".json")
        if output.exists():
            pin = json.loads(manifest.read_text(encoding="utf-8"))[0]
            assert pin["SourceSha256"] == source_sha
            assert pin["Sha256"] == hashlib.sha256(output.read_bytes()).hexdigest().upper()
            continue
        with ZipFile(source) as parent, ZipFile(output, "w", ZIP_DEFLATED) as result:
            for entry in parent.infolist():
                data = parent.read(entry)
                if entry.filename == "word/settings.xml" and setting != "default":
                    document = ET.fromstring(data)
                    ET.SubElement(document.find(m + "mathPr"), m + "brkBin", {m + "val": setting})
                    data = ET.tostring(document, encoding="utf-8", xml_declaration=True)
                if entry.filename == "word/document.xml" and upright:
                    document = ET.fromstring(data)
                    for run in document.iter(m + "r"):
                        ET.SubElement(ET.SubElement(run, m + "rPr"), m + "sty", {m + "val": "p"})
                    data = ET.tostring(document, encoding="utf-8", xml_declaration=True)
                result.writestr(entry, data)
        manifest.write_text(json.dumps([{
            "Id": identifier, "Input": output.relative_to(repo).as_posix(), "SourceSha256": source_sha,
            "Sha256": hashlib.sha256(output.read_bytes()).hexdigest().upper(),
            "HeldOut": False, "ExpectedEquationCount": 1, "ExpectedInlineImageCount": 0,
            "ExpectedText": ["Before formula:", "after formula.", "Following prose."]
        }], indent=2) + "\n", encoding="utf-8")
