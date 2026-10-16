"""Pin adjacent explicit display equations without rewriting accepted controls."""
from copy import deepcopy
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E

root = Path(__file__).resolve().parent.parent
cases = root / "tests/Lokad.OoxPdf.Tests/Cases"
source = cases / "docx-omml-literal-after.docx"
source_sha = hashlib.sha256(source.read_bytes()).hexdigest().upper()
assert source_sha == "FBE2C3D11FFAFD6BE8EE9CBE4A7A4CDA26BA0E938B6FA6C2B7AD5214E36873CD"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
rows = []
for grouping in ["same-group", "separate-groups"]:
    with ZipFile(source) as archive:
        entries = {name: archive.read(name) for name in archive.namelist()}
    document = E.fromstring(entries["word/document.xml"])
    group = next(document.iter(m + "oMathPara"))
    second = deepcopy(group.find(m + "oMath"))
    for node, text in zip(second.iter(m + "t"), ["δ", "+", "ε", "=", "ζ"]):
        node.text = text
    if grouping == "same-group":
        group.append(second)
    else:
        paragraph = next(parent for parent in document.iter() if group in list(parent))
        another = E.Element(m + "oMathPara")
        another.append(second)
        paragraph.insert(list(paragraph).index(group) + 1, another)
    entries["word/document.xml"] = E.tostring(document, encoding="utf-8", xml_declaration=True)
    target = cases / "CompositionProbes" / ("math-explicit-display-" + grouping + ".docx")
    with ZipFile(target, "w", ZIP_DEFLATED) as archive:
        for name, data in entries.items():
            info = ZipInfo(name, (2026, 10, 16, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
        SourceSha256=source_sha, HeldOut=grouping == "separate-groups",
        ExpectedEquationCount=2, ExpectedText=["Before formula:", "Following prose."]))
(cases / "CompositionProbes/math-explicit-display.json").write_text(
    json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Pinned", len(rows), "explicit display-group controls")
