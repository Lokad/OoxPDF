"""Pin authored display alignment across body, table and static-story owners."""
from copy import deepcopy
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E

root = Path(__file__).resolve().parent.parent
cases = root / "tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
sources = {
    "math-explicit-display-same-group": "D4B2C8C1365A081E098C9DD472DB61CFB4270A2B258A9AF1C1E61477DEE0042B",
    "math-tabs-right108-cell": "51C5845DE8E93D176C4ED4054D7BADEA28D568BE32796720D671CEAEA94A4708",
    "math-tabs-center108-footer": "C22BE01FF50DBEFAFFB6CE8C39DD96F12E64508F9151FC4FCDAB9E353D65E5F5",
}
for name, digest in sources.items():
    assert hashlib.sha256((cases / (name + ".docx")).read_bytes()).hexdigest().upper() == digest
with ZipFile(cases / "math-explicit-display-same-group.docx") as archive:
    doc = E.fromstring(archive.read("word/document.xml"))
    display = next(p for p in doc.iter(w + "p") if len(list(p.iter(m + "oMath"))) == 2)

rows = []
for location, alignment, source in [("body", "left", "math-explicit-display-same-group"),
                                    ("body", "right", "math-explicit-display-same-group"),
                                    ("cell", "left", "math-tabs-right108-cell"),
                                    ("footer", "right", "math-tabs-center108-footer")]:
    with ZipFile(cases / (source + ".docx")) as archive:
        entries = {name: archive.read(name) for name in archive.namelist()}
    part = "word/math-tabs-footer.xml" if location == "footer" else "word/document.xml"
    doc = E.fromstring(entries[part])
    paragraph = next(p for p in doc.iter(w + "p") if len(list(p.iter(m + "oMath"))) == 2)
    if location in ["cell", "footer"]:
        parent = next(p for p in doc.iter() if paragraph in list(p))
        index = list(parent).index(paragraph)
        parent.remove(paragraph)
        paragraph = deepcopy(display)
        parent.insert(index, paragraph)
    for group in paragraph.iter(m + "oMathPara"):
        properties = E.Element(m + "oMathParaPr")
        E.SubElement(properties, m + "jc", {m + "val": alignment})
        group.insert(0, properties)
    if location == "body" and alignment == "right":
        properties = E.Element(w + "pPr")
        E.SubElement(properties, w + "jc", {w + "val": "both"})
        paragraph.insert(0, properties)
    entries[part] = E.tostring(doc, encoding="utf-8", xml_declaration=True)
    target = cases / ("math-explicit-display-" + alignment + "-" + location + ".docx")
    with ZipFile(target, "w", ZIP_DEFLATED) as archive:
        for name, data in entries.items():
            info = ZipInfo(name, (2026, 10, 16, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(), SourceSha256=sources[source],
        HeldOut=location != "body", ExpectedEquationCount=2, ExpectedText=["Following prose."]))
(cases / "math-explicit-display-alignment.json").write_text(
    json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Pinned", len(rows), "authored display-alignment controls")
