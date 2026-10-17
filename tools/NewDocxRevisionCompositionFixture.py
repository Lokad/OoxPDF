"""One complete public report combines revision syntax across body and stories."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib
import json

root = Path(__file__).resolve().parent.parent
cases = root / "tests/Lokad.OoxPdf.Tests/Cases/Composition"
source = cases / "composed-static-cell-report.docx"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
E.register_namespace("m", m[1:-1])
E.register_namespace("w", w[1:-1])
with ZipFile(source) as z:
    entries = {name: z.read(name) for name in z.namelist()}
changes = [("word/document.xml", 0, "ins", False, "1201"),
           ("word/footer1.xml", 0, "del", True, "1202"),
           ("word/endnotes.xml", 0, "del", False, "1203"),
           ("word/endnotes.xml", 1, "ins", True, "1204"),
           ("word/header1.xml", 0, "ins", True, "1205")]
documents = {}
for part, index, kind, whole, revision_id in changes:
    xml = documents.setdefault(part, E.fromstring(entries[part]))
    equation = list(xml.iter(m + "oMath"))[index]
    node = equation if whole else list(equation)[0]
    parent = next(parent for parent in xml.iter() if node in list(parent))
    position = list(parent).index(node)
    parent.remove(node)
    wrapper = E.Element(w + kind, {w + "id": revision_id, w + "author": "Composite Math Reviewer",
                                  w + "date": "2026-10-17T00:00:00Z"})
    wrapper.append(node)
    parent.insert(position, wrapper)
for part, xml in documents.items():
    entries[part] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
target = cases / "revision-composition-report.docx"
with ZipFile(target, "w", ZIP_DEFLATED) as z:
    for name, data in sorted(entries.items()):
        info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
        info.compress_type = ZIP_DEFLATED
        z.writestr(info, data)
rows = [dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
             Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
             SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(), HeldOut=False,
             ExpectedEquationCount=10, ExpectedText=["All quantities in this synthetic example are illustrative.",
                                                   "Review section 12."])]
(cases / "revision-composition-report.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Pinned complete revision-composition report; original portfolio unchanged")
