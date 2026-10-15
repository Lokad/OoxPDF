"""Independent page-boundary controls for unpainted mathematical row tails."""
import argparse
from copy import deepcopy
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET

parser = argparse.ArgumentParser()
parser.add_argument("--after", type=int, choices=(0, 6, 24), required=True)
parser.add_argument("--cant-split", action="store_true")
parser.add_argument("--normal-math", action="store_true")
parser.add_argument("--spacer", type=int, choices=(600, 620), default=600)
args = parser.parse_args()
repo = Path(__file__).resolve().parents[1]
cases = repo / "tests/Lokad.OoxPdf.Tests/Cases"
source = cases / "Composition/report.docx"
source_sha = hashlib.sha256(source.read_bytes()).hexdigest().upper()
assert source_sha == "10F63D162D0D46DBE8CD682BB50452DB6672104706F116EA50CDCAFEE054575F"
identifier = f"math-table-boundary-after{args.after}" + ("-cant-split" if args.cant_split else "") + ("-tight" if args.spacer == 620 else "") + ("-normal" if args.normal_math else "")
output = cases / ("PortableFonts" if args.normal_math else "CompositionProbes") / (identifier + ".docx")
manifest = output.with_suffix(".json")
if output.exists():
    pin = json.loads(manifest.read_text(encoding="utf-8"))[0]
    assert pin["SourceSha256"] == source_sha
    assert pin["Sha256"] == hashlib.sha256(output.read_bytes()).hexdigest().upper()
    raise SystemExit(0)
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
ET.register_namespace("w", w[1:-1])
with ZipFile(source) as parent, ZipFile(output, "w", ZIP_DEFLATED) as result:
    for entry in parent.infolist():
        data = parent.read(entry)
        if entry.filename == "word/document.xml":
            doc = ET.fromstring(data)
            body = doc.find(w + "body")
            section = deepcopy(body.find(w + "sectPr"))
            for child in list(section):
                if child.tag in (w + "headerReference", w + "footerReference"):
                    section.remove(child)
            table = deepcopy(body.find(w + "tbl"))
            row = table.findall(w + "tr")[-1]
            if args.cant_split:
                props = row.find(w + "trPr")
                if props is None:
                    props = ET.Element(w + "trPr")
                    row.insert(0, props)
                ET.SubElement(props, w + "cantSplit")
            paragraph = row.findall(w + "tc")[-1].find(w + "p")
            props = ET.Element(w + "pPr")
            ET.SubElement(props, w + "spacing", {w + "after": str(args.after * 20)})
            paragraph.insert(0, props)
            body.clear()
            spacer = ET.SubElement(body, w + "p")
            ET.SubElement(ET.SubElement(spacer, w + "pPr"), w + "spacing", {
                w + "before": "0", w + "after": "0", w + "line": str(args.spacer * 20), w + "lineRule": "exact"})
            body.append(table)
            paragraph = ET.SubElement(body, w + "p")
            ET.SubElement(ET.SubElement(paragraph, w + "r"), w + "t").text = "Following boundary table."
            body.append(section)
            if args.normal_math:
                for run in doc.iter(m + "r"):
                    props = run.find(m + "rPr")
                    if props is None:
                        props = ET.Element(m + "rPr")
                        run.insert(0, props)
                    ET.SubElement(props, m + "nor", {m + "val": "1"})
            data = ET.tostring(doc, encoding="utf-8", xml_declaration=True)
        result.writestr(entry, data)
manifest.write_text(json.dumps([{
    "Id": identifier, "Input": output.relative_to(repo).as_posix(),
    "Sha256": hashlib.sha256(output.read_bytes()).hexdigest().upper(),
    "SourceSha256": source_sha, "HeldOut": False, "ExpectedEquationCount": 1,
    "ExpectedInlineImageCount": 0,
    "ExpectedText": ["Planning input", "Observed series", "Following boundary table."]
}], indent=2) + "\n", encoding="utf-8")
