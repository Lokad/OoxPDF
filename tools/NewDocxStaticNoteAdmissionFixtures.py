"""Public, byte-identical versions of the independently exported C170 controls."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib
import json

root = Path(__file__).resolve().parent.parent
cases = root / "tests/Lokad.OoxPdf.Tests/Cases"
source = cases / "Composition/composed-operator-matrix-specification.docx"
source_hash = hashlib.sha256(source.read_bytes()).hexdigest().upper()
assert source_hash == "D7B869C53D801B3C861513B0489A5B45DF646751C4938710D1D8408CBBF4614E"
with ZipFile(source) as archive:
    original = {name: archive.read(name) for name in archive.namelist()}
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
rows = []
for variant, tags in [("without-footnote", ["footnoteReference"]),
                      ("without-static", ["headerReference", "footerReference"]),
                      ("without-header", ["headerReference"]), ("without-footer", ["footerReference"])]:
    entries = dict(original)
    document = E.fromstring(entries["word/document.xml"])
    for parent in document.iter():
        for child in list(parent):
            if child.tag in [w + tag for tag in tags]:
                parent.remove(child)
    entries["word/document.xml"] = E.tostring(document, encoding="utf-8", xml_declaration=True)
    name = "operator-matrix-specification-" + variant
    destination = cases / "CompositionProbes" / (name + ".docx")
    with ZipFile(destination, "w", ZIP_DEFLATED) as archive:
        for path, data in sorted(entries.items()):
            info = ZipInfo(path, (2026, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    rows.append(dict(Id=name, Input="tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/" + name + ".docx",
        Sha256=hashlib.sha256(destination.read_bytes()).hexdigest().upper(), SourceSha256=source_hash,
        HeldOut=False, ExpectedEquationCount=6, ExpectedText=["Replenishment calculation specification"],
        Features=[variant]))
(cases / "CompositionProbes/static-note-admission.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Generated four public static/note admission controls; existing sources unchanged.")
