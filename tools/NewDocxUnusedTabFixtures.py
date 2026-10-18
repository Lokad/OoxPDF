"""Add unused tab metadata without changing the consumed source fields."""
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
cases = root / "tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes"
sources = json.loads((cases / "math-tabs.json").read_text(encoding="utf-8-sig"))
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
rows = []
for container, kind in [("body", "decimal"), ("cell", "leader")]:
    original = next(row for row in sources if row["Id"] == "math-tabs-left72-" + container)
    with ZipFile(root / original["Input"]) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    document = E.fromstring(parts["word/document.xml"])
    paragraph = next(p for p in document.iter(w + "p") if list(p.iter(m + "oMath")))
    tabs = paragraph.find(w + "pPr/" + w + "tabs")
    assert tabs is not None
    attributes = {w + "val": "decimal" if kind == "decimal" else "left", w + "pos": "14400"}
    if kind == "leader":
        attributes[w + "leader"] = "dot"
    added = E.SubElement(tabs, w + "tab", attributes)
    changed = E.tostring(document, encoding="utf-8", xml_declaration=True)
    tabs.remove(added)
    assert E.tostring(document) == E.tostring(E.fromstring(parts["word/document.xml"]))
    parts["word/document.xml"] = changed
    name = "math-tabs-unused-" + kind + "-" + container
    destination = output / (name + ".docx")
    with ZipFile(destination, "w", ZIP_DEFLATED) as archive:
        for path, data in sorted(parts.items()):
            info = ZipInfo(path, (2026, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    row = dict(original)
    row.update(Id=name, Input=destination.relative_to(root).as_posix(),
               Sha256=hashlib.sha256(destination.read_bytes()).hexdigest().upper(),
               SourceSha256=original["Sha256"], HeldOut=False,
               Change="One unused720pt tab stop; active72/144pt stops and all content retained.")
    rows.append(row)
(output / "unused-tab-math.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Two unused-tab metadata controls generated; source fields retained.")
