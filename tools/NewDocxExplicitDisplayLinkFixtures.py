"""Pin linked explicit-display controls across body, cell and footer owners."""
from copy import deepcopy
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E

root = Path(__file__).resolve().parent.parent
cases = root / "tests/Lokad.OoxPdf.Tests/Cases"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
sources = {
    "CompositionProbes/math-explicit-display-same-group.docx": "D4B2C8C1365A081E098C9DD472DB61CFB4270A2B258A9AF1C1E61477DEE0042B",
    "CompositionProbes/math-tabs-right108-cell.docx": "51C5845DE8E93D176C4ED4054D7BADEA28D568BE32796720D671CEAEA94A4708",
    "CompositionProbes/math-tabs-center108-footer.docx": "C22BE01FF50DBEFAFFB6CE8C39DD96F12E64508F9151FC4FCDAB9E353D65E5F5",
    "docx-omml-fraction-after.docx": "07CC42DAFD24532B756FB75B808102F62AAD8D5E2CDE5B591E5615B433149C0E",
}
for name, digest in sources.items():
    assert hashlib.sha256((cases / name).read_bytes()).hexdigest().upper() == digest
with ZipFile(cases / "CompositionProbes/math-explicit-display-same-group.docx") as archive:
    doc = E.fromstring(archive.read("word/document.xml"))
    display = next(p for p in doc.iter(w + "p") if len(list(p.iter(m + "oMath"))) == 2)

rows = []
for location, source in [("body", "CompositionProbes/math-explicit-display-same-group.docx"),
                         ("cell", "CompositionProbes/math-tabs-right108-cell.docx"),
                         ("footer", "CompositionProbes/math-tabs-center108-footer.docx"),
                         ("fraction", "docx-omml-fraction-after.docx")]:
    with ZipFile(cases / source) as archive:
        entries = {name: archive.read(name) for name in archive.namelist()}
    part = "word/math-tabs-footer.xml" if location == "footer" else "word/document.xml"
    doc = E.fromstring(entries[part])
    paragraph = next(p for p in doc.iter(w + "p") if list(p.iter(m + "oMath")))
    if location in ["cell", "footer"]:
        parent = next(p for p in doc.iter() if paragraph in list(p))
        index = list(parent).index(paragraph)
        parent.remove(paragraph)
        paragraph = deepcopy(display)
        parent.insert(index, paragraph)
    group = next(paragraph.iter(m + "oMathPara"))
    index = list(paragraph).index(group)
    paragraph.remove(group)
    link = E.Element(w + "hyperlink", {w + "anchor": "PublicCalculation"})
    link.append(group)
    paragraph.insert(index, E.Element(w + "bookmarkStart", {w + "id": "41", w + "name": "PublicCalculation"}))
    paragraph.insert(index + 1, link)
    paragraph.insert(index + 2, E.Element(w + "bookmarkEnd", {w + "id": "41"}))
    entries[part] = E.tostring(doc, encoding="utf-8", xml_declaration=True)
    target = cases / "CompositionProbes" / ("math-explicit-display-linked-" + location + ".docx")
    with ZipFile(target, "w", ZIP_DEFLATED) as archive:
        for name, data in entries.items():
            info = ZipInfo(name, (2026, 10, 16, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
        SourceSha256=sources[source], HeldOut=location in ["cell", "footer"],
        ExpectedEquationCount=1 if location == "fraction" else 2,
        ExpectedText=[], ExpectedMathText="".join(node.text or "" for node in paragraph.iter(m + "t"))))
(cases / "CompositionProbes/math-explicit-display-linked.json").write_text(
    json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Pinned", len(rows), "linked explicit-display controls")
