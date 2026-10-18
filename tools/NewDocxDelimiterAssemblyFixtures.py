"""Generate bounded growing/fixed delimiter controls from the retained root probe."""
from argparse import ArgumentParser
from copy import deepcopy
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
source = root / "tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/radical-assembly-child200.docx"
source_pin = json.loads((source.parent / "radical-assembly.json").read_text(encoding="utf-8"))[0]["Sha256"]
assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == source_pin
with ZipFile(source) as archive:
    original = {name: archive.read(name) for name in archive.namelist()}
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
rows = []
for child_size, grow in ((60, True), (200, True), (60, False)):
    document = E.fromstring(original["word/document.xml"])
    equation = document.find(".//" + m + "oMath")
    child = deepcopy(equation.find(".//" + m + "rad/" + m + "e/" + m + "r"))
    child.find(w + "rPr/" + w + "sz").set(w + "val", str(child_size * 2))
    equation[:] = []
    delimiter = E.SubElement(equation, m + "d")
    properties = E.SubElement(delimiter, m + "dPr")
    E.SubElement(properties, m + "begChr", {m + "val": "("})
    E.SubElement(properties, m + "endChr", {m + "val": ")"})
    E.SubElement(properties, m + "grow", {m + "val": "1" if grow else "0"})
    control = E.SubElement(E.SubElement(properties, m + "ctrlPr"), w + "rPr")
    E.SubElement(control, w + "rFonts", {w + "ascii": "Cambria Math", w + "hAnsi": "Cambria Math"})
    E.SubElement(control, w + "sz", {w + "val": "24"})
    E.SubElement(delimiter, m + "e").append(child)
    entries = dict(original)
    entries["word/document.xml"] = E.tostring(document, encoding="utf-8", xml_declaration=True)
    name = "delimiter-assembly-child" + str(child_size) + ("" if grow else "-fixed")
    destination = output / (name + ".docx")
    with ZipFile(destination, "w", ZIP_DEFLATED) as archive:
        for path, data in sorted(entries.items()):
            info = ZipInfo(path, (2026, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    rows.append(dict(Id=name, Input=destination.relative_to(root).as_posix(),
                     Sha256=hashlib.sha256(destination.read_bytes()).hexdigest().upper(),
                     SourceSha256=source_pin, HeldOut=False, ExpectedEquationCount=1,
                     ExpectedText=["Following prose."], ControlSize=12,
                     ChildSize=child_size, Grow=grow))
(output / "delimiter-assembly.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Three authored-control delimiter probes generated; source inputs unchanged.")
