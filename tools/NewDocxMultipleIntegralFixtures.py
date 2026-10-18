"""Compose inline, explicit-side and inherited-stacked double/triple integrals."""
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
source = root / "tests/Lokad.OoxPdf.Tests/Cases/docx-omml-integral-inline-stacked-request.docx"
with ZipFile(source) as archive:
    original = {name: archive.read(name) for name in archive.namelist()}
displays = []
for kind in ("side", "stacked"):
    with ZipFile(root / ("tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/display-integral-" + kind + ".docx")) as archive:
        displays.append(next(p for p in E.fromstring(archive.read("word/document.xml")).find(w + "body")
                             if p.find(m + "oMathPara") is not None))
rows = []
for kind, character in (("double", "\u222c"), ("triple", "\u222d")):
    parts = dict(original)
    document = E.fromstring(parts["word/document.xml"])
    body = document.find(w + "body")
    following = next(p for p in body if "Following prose." in "".join(t.text or "" for t in p.iter(w + "t")))
    index = list(body).index(following)
    for offset, (label, display) in enumerate(zip(("Side limits:", "Stacked limits:"), displays)):
        paragraph = E.Element(w + "p")
        E.SubElement(E.SubElement(paragraph, w + "r"), w + "t").text = label
        body.insert(index + 2 * offset, paragraph)
        body.insert(index + 2 * offset + 1, deepcopy(display))
    equations = list(body.iter(m + "oMath"))
    assert len(equations) == 3
    for i, equation in enumerate(equations):
        properties = equation.find(m + "nary/" + m + "naryPr")
        properties.find(m + "chr").set(m + "val", character)
        if i != 1:  # Inline and stacked display use the document integral default.
            location = properties.find(m + "limLoc")
            if location is not None:
                properties.remove(location)
    parts["word/document.xml"] = E.tostring(document, encoding="utf-8", xml_declaration=True)
    settings = E.fromstring(parts["word/settings.xml"])
    math = settings.find(m + "mathPr")
    if math is None:
        math = E.SubElement(settings, m + "mathPr")
    for node in list(math):
        if node.tag in (m + "intLim", m + "naryLim"):
            math.remove(node)
    E.SubElement(math, m + "intLim", {m + "val": "undOvr"})
    E.SubElement(math, m + "naryLim", {m + "val": "subSup"})
    parts["word/settings.xml"] = E.tostring(settings, encoding="utf-8", xml_declaration=True)
    name = "multiple-integral-" + kind
    destination = output / (name + ".docx")
    with ZipFile(destination, "w", ZIP_DEFLATED) as archive:
        for path, data in sorted(parts.items()):
            info = ZipInfo(path, (2026, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    rows.append(dict(Id=name, Input=destination.relative_to(root).as_posix(),
                     Sha256=hashlib.sha256(destination.read_bytes()).hexdigest().upper(),
                     SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
                     HeldOut=False, ExpectedEquationCount=3,
                     ExpectedText=["Before formula:", "after formula.", "Side limits:", "Stacked limits:", "Following prose."],
                     IntegralCharacter=character))
(output / "multiple-integral.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Two integral-family controls generated; three authored placements in each.")
