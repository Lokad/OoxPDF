"""Create source-pinned balanced literal groups for independent operator-wrap controls."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib
import json

root = Path(__file__).resolve().parent.parent
cases = root / "tests/Lokad.OoxPdf.Tests/Cases"
source = cases / "docx-omml-inline-overfull12.docx"
source_hash = hashlib.sha256(source.read_bytes()).hexdigest().upper()
assert source_hash == "45F27A6285C1A065D5F4BE435AD335FC9F96E8341AB3BDC1D880BBD6DE2D8867"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
rows = []
for name, literal, held_out in [("atomic-math-parenthesized", "(α+β=γ)", False),
                                ("math-grouped-nested", "(α+[β=γ])", True)]:
    with ZipFile(source) as archive:
        entries = {name: archive.read(name) for name in archive.namelist()}
    document = E.fromstring(entries["word/document.xml"])
    texts = list(document.iter(m + "t"))
    texts[0].text = literal
    for text in texts[1:]:
        text.text = ""
    entries["word/document.xml"] = E.tostring(document, encoding="utf-8", xml_declaration=True)
    target = cases / "CompositionProbes" / (name + ".docx")
    with ZipFile(target, "w", ZIP_DEFLATED) as archive:
        for entry_name, data in entries.items():
            info = ZipInfo(entry_name, (2026, 10, 16, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    rows.append(dict(Id=name, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
        SourceSha256=source_hash, HeldOut=held_out, ExpectedEquationCount=1,
        ExpectedText=["Before formula:", "after formula.", "Following prose."]))
(cases / "CompositionProbes/math-grouped-literal.json").write_text(
    json.dumps(rows, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
