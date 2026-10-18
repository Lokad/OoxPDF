"""Change line ownership without changing supported equation content."""
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
cases = root / "tests/Lokad.OoxPdf.Tests/Cases"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
rows = []
for size, container in [(12, "body"), (72, "body"), (72, "cell"), (72, "header"), (48, "report")]:
    suffix = "" if container == "body" else "-" + container
    source = cases / ("Composition/report.docx" if container == "report" else
                      f"CompositionProbes/math-minimum-line-{size}{suffix}.docx")
    part = "word/headerMinimum.xml" if container == "header" else "word/document.xml"
    with ZipFile(source) as archive:
        parts = {path: archive.read(path) for path in archive.namelist()}
    if container == "report":
        document = E.fromstring(parts[part])
        paragraph = next(p for p in document.iter(w + "p") if list(p.iter(m + "oMath")))
        assert paragraph.find(w + "pPr") is None
        properties = E.Element(w + "pPr")
        E.SubElement(properties, w + "spacing", {w + "line": str(size * 20), w + "lineRule": "exact"})
        paragraph.insert(0, properties)
        parts[part] = E.tostring(document, encoding="utf-8", xml_declaration=True)
        name = "exact-line-math-composed-report"
        expected = ["Quarterly operations report", "Inventory availability", "Review section 12.",
                    "All quantities in this synthetic example are illustrative."]
    else:
        assert parts[part].count(b'lineRule="atLeast"') == 3
        parts[part] = parts[part].replace(b'lineRule="atLeast"', b'lineRule="exact"')
        name = f"exact-line-math-{size}{suffix}"
        expected = ["Before formula:", "after formula.", "Side limits:", "Stacked limits:", "Following prose."]
        if container == "header":
            expected.append("Composed story.")
    destination = output / (name + ".docx")
    with ZipFile(destination, "w", ZIP_DEFLATED) as archive:
        for path, data in sorted(parts.items()):
            info = ZipInfo(path, (2026, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    rows.append(dict(Id=name, Input=destination.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(destination.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        HeldOut=container == "report", ExpectedEquationCount=6 if container == "report" else 3,
        ExpectedText=expected, ExactLineHeight=size, Container=container, Part=part))
(output / "exact-line-math.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Five exact-line transformations generated; supported source equations retained.")
