"""Combine paragraph justification and wrapping with unchanged report formulas."""
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
source = root / "tests/Lokad.OoxPdf.Tests/Cases/Composition/report.docx"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
rows = []
for name, wrapped, justified in [("justified-math-single-report", False, True),
                                 ("left-math-wrapped-report", True, False),
                                 ("justified-math-wrapped-report", True, True)]:
    with ZipFile(source) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    document = E.fromstring(parts["word/document.xml"])
    paragraph = next(p for p in document.iter(w + "p") if len(list(p.iter(m + "oMath"))) == 2)
    assert paragraph.find(w + "pPr") is None
    properties = E.Element(w + "pPr")
    E.SubElement(properties, w + "jc", {w + "val": "both" if justified else "left"})
    paragraph.insert(0, properties)
    if wrapped:
        extra = E.SubElement(paragraph, w + "r")
        E.SubElement(extra, w + "t", {"{http://www.w3.org/XML/1998/namespace}space": "preserve"}).text = (
            " Additional public calculation text preserves the existing source formulas and provides ordinary prose for line wrapping." * 5)
    changed = E.tostring(document, encoding="utf-8", xml_declaration=True)
    paragraph.remove(properties)
    if wrapped:
        paragraph.remove(extra)
    assert E.tostring(document) == E.tostring(E.fromstring(parts["word/document.xml"]))
    parts["word/document.xml"] = changed
    destination = output / (name + ".docx")
    with ZipFile(destination, "w", ZIP_DEFLATED) as archive:
        for path, data in sorted(parts.items()):
            info = ZipInfo(path, (2026, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    expected = ["Quarterly operations report", "Inventory availability", "Review section 12.",
                "All quantities in this synthetic example are illustrative."]
    if wrapped:
        expected.append("Additional public calculation text preserves the existing source formulas")
    rows.append(dict(Id=name, Input=destination.relative_to(root).as_posix(),
                     Sha256=hashlib.sha256(destination.read_bytes()).hexdigest().upper(),
                     SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
                     HeldOut=False, ExpectedEquationCount=6, ExpectedText=expected,
                     Wrapped=wrapped, Justified=justified))
(output / "justified-math.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Three justified/wrapped report controls generated; source formulas retained.")
