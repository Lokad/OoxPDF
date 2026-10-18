"""Turn retained inline integrals into standalone display controls."""
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
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
rows = []
for limit in ("subSup", "undOvr"):
    source = root / "tests/Lokad.OoxPdf.Tests/Cases/docx-omml-integral-inline-stacked-request.docx"
    with ZipFile(source) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    document = E.fromstring(parts["word/document.xml"])
    body = document.find(w + "body")
    paragraph = next(p for p in body if p.find(m + "oMath") is not None)
    equation = paragraph.find(m + "oMath")
    equation.find(m + "nary/" + m + "naryPr/" + m + "limLoc").set(m + "val", limit)
    children = list(paragraph)
    equation_index = children.index(equation)
    properties = paragraph.find(w + "pPr")

    def new_paragraph():
        p = E.Element(w + "p")
        if properties is not None:
            p.append(deepcopy(properties))
        return p

    before = new_paragraph()
    before.extend(deepcopy(child) for child in children[:equation_index] if child.tag != w + "pPr")
    display = new_paragraph()
    math_paragraph = E.SubElement(display, m + "oMathPara")
    math_properties = E.SubElement(math_paragraph, m + "oMathParaPr")
    E.SubElement(math_properties, m + "jc", {m + "val": "left"})
    math_paragraph.append(deepcopy(equation))
    after = new_paragraph()
    after.extend(deepcopy(child) for child in children[equation_index + 1:])
    index = list(body).index(paragraph)
    body.remove(paragraph)
    for offset, replacement in enumerate((before, display, after)):
        body.insert(index + offset, replacement)
    parts["word/document.xml"] = E.tostring(document, encoding="utf-8", xml_declaration=True)
    name = "display-integral-" + ("side" if limit == "subSup" else "stacked")
    destination = output / (name + ".docx")
    with ZipFile(destination, "w", ZIP_DEFLATED) as archive:
        for path, data in sorted(parts.items()):
            info = ZipInfo(path, (2026, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    rows.append(dict(Id=name, Input=destination.relative_to(root).as_posix(),
                     Sha256=hashlib.sha256(destination.read_bytes()).hexdigest().upper(),
                     SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
                     HeldOut=False, ExpectedEquationCount=1,
                     ExpectedText=["Before formula:", "after formula.", "Following prose."],
                     LimitLocation=limit))
(output / "display-integral.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Two standalone display-integral controls generated; inline sources unchanged.")
