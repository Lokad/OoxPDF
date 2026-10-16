"""Generate plain-word overflow controls from a pinned public narrow-column input."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib
import json

root = Path(__file__).resolve().parent.parent
cases = root / "tests/Lokad.OoxPdf.Tests/Cases"
source = cases / "docx-omml-inline-script-overwide30.docx"
source_hash = hashlib.sha256(source.read_bytes()).hexdigest().upper()
assert source_hash == "1ABBC8AA37727FB1E57894F8F0EFF715BF05FD25FEB3DCACCC0472F76F15E29C"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
output = cases / "CompositionProbes"
rows = []
for label, value in [("default", None), ("on", "1"), ("off", "0")]:
    with ZipFile(source) as archive:
        entries = {name: archive.read(name) for name in archive.namelist()}
    document = E.fromstring(entries["word/document.xml"])
    for paragraph in document.iter(w + "p"):
        for equation in list(paragraph.findall(m + "oMath")):
            paragraph.remove(equation)
        if value is not None:
            properties = E.Element(w + "pPr")
            E.SubElement(properties, w + "wordWrap", {w + "val": value})
            paragraph.insert(0, properties)
    entries["word/document.xml"] = E.tostring(document, encoding="utf-8", xml_declaration=True)
    target = output / ("overwide-word-" + label + ".docx")
    with ZipFile(target, "w", ZIP_DEFLATED) as archive:
        for name, data in entries.items():
            info = ZipInfo(name, (2026, 10, 16, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
        SourceSha256=source_hash, HeldOut=False,
        ExpectedText=["Before formula:", "after formula.", "Following prose."]))
(output / "overwide-word.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
