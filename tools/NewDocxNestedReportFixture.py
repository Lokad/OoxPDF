"""A development report combining qualified nested content with ordinary report features.

Derive only from the non-held-out report and the public navigation control.
Existing checkpoint documents and held-out inputs remain byte-identical.
"""
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib
import json
import posixpath

root = Path(__file__).resolve().parent.parent
cases = root / "tests/Lokad.OoxPdf.Tests/Cases"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
r = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
package = "{http://schemas.openxmlformats.org/package/2006/relationships}"
wp = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
for prefix, namespace in [("w", w), ("m", m), ("r", r), ("wp", wp)]:
    E.register_namespace(prefix, namespace[1:-1])

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()

def load(path):
    with ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}

base = cases / "Composition/report.docx"
assert sha(base) == "10F63D162D0D46DBE8CD682BB50452DB6672104706F116EA50CDCAFEE054575F"
portfolio = json.loads((cases / "Composition/portfolio-checkpoint.json").read_text(encoding="utf-8-sig"))
assert not next(item for item in portfolio if item["Id"] == "report")["HeldOut"]
nav = next(item for item in json.loads((cases / "CompositionProbes/nested-math-navigation.json").read_text(encoding="utf-8-sig"))
           if item["Id"] == "nested-math-navigation-outgoing")
source = root / nav["Input"]
assert sha(source) == nav["Sha256"]
entries = load(base)
nested_entries = load(source)
document = E.fromstring(entries["word/document.xml"])
body = document.find(w + "body")
relationships = E.fromstring(entries["word/_rels/document.xml.rels"])
template = E.fromstring(nested_entries["word/document.xml"])
table = deepcopy(template.find(w + "body/" + w + "tbl"))

# Resolve image ownership explicitly while merging the two independent packages.
image = next(item for item in E.fromstring(nested_entries["word/_rels/document.xml.rels"])
             if item.get("Type", "").endswith("/image"))
image_id = "nestedCalculationImage"
assert all(item.get("Id") != image_id for item in relationships)
old_target = posixpath.normpath("word/" + image.get("Target"))
entries["word/media/nested-calculation.png"] = nested_entries[old_target]
E.SubElement(relationships, package + "Relationship", {"Id": image_id,
    "Type": r[1:-1] + "/image", "Target": "media/nested-calculation.png"})
for item in table.iter():
    for name, value in list(item.attrib.items()):
        if name in (r + "embed", r + "link"):
            assert value == image.get("Id")
            item.set(name, image_id)

bookmark_id = str(max([int(item.get(w + "id")) for item in document.iter(w + "bookmarkStart")] + [0]) + 1)
for item in table.iter(w + "bookmarkStart"):
    item.set(w + "id", bookmark_id)
    item.set(w + "name", "NestedCalculationDetails")
for item in table.iter(w + "bookmarkEnd"):
    item.set(w + "id", bookmark_id)
for item in table.iter(w + "hyperlink"):
    assert item.get(w + "anchor") == "PublicNestedTarget"
    item.set(w + "anchor", "NestedCalculationDetails")
for item in table.iter(w + "t"):
    item.text = {"Public outer head": "Calculation summary",
                 "Public outer tail": "Detail: review the accepted ratio."}.get(item.text, item.text)
next_id = max([int(item.get("id")) for item in document.iter(wp + "docPr")] + [0]) + 1
for item in table.iter(wp + "docPr"):
    item.set("id", str(next_id)); next_id += 1
    item.set("name", "Calculation ratio trend")

# Keep the report's paragraphs, table, numbering, comment, footnote and static
# parts intact; place one nested drill-down alongside its existing calculation.
position = list(body).index(body.find(w + "tbl"))
body.insert(position + 1, table)
entries["word/document.xml"] = E.tostring(document, encoding="utf-8", xml_declaration=True)
entries["word/_rels/document.xml.rels"] = E.tostring(relationships, encoding="utf-8", xml_declaration=True)
target = cases / "Composition/composed-nested-navigation-report.docx"
with ZipFile(target, "w", ZIP_DEFLATED) as archive:
    for name, payload in sorted(entries.items()):
        info = ZipInfo(name, (2026, 10, 18, 0, 0, 0))
        info.compress_type = ZIP_DEFLATED
        archive.writestr(info, payload)
row = dict(Id=target.stem, Input=target.relative_to(root).as_posix(), Sha256=sha(target), HeldOut=False,
    BaseInput=base.relative_to(root).as_posix(), BaseInputSha256=sha(base),
    NestedInput=nav["Input"], NestedInputSha256=nav["Sha256"],
    ExpectedEquationCount=8, ExpectedInlineImageCount=2,
    ExpectedText=["Quarterly operations report", "Inventory availability", "Calculation summary",
        "Detail: review the accepted ratio.", "Review section 12.", "All quantities in this synthetic example are illustrative."],
    Features=["original report tables/numbering/footnote/comment/header/footer", "nested fractions and illustration",
        "internal link from a fraction to an enclosing-cell bookmark", "mixed Calibri/Arial/Cambria Math source metrics"],
    SourceFacts=dict(BodyEquations=7, HeaderEquationsPerPage=1, SourceFootnotes=1, NestedFractionLinks=1,
        NestedBookmark="NestedCalculationDetails"))
(cases / "Composition/nested-report.json").write_text(json.dumps([row], indent=2) + "\n", encoding="utf-8")
assert sha(base) == row["BaseInputSha256"] and sha(source) == row["NestedInputSha256"]
print("Pinned one development report; original and held-out portfolio inputs remain unchanged")
