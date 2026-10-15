"""Pin authored whitespace around comment references; no generated spacer is present."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib, json
root = Path(__file__).resolve().parent.parent
cases = root / "tests/Lokad.OoxPdf.Tests/Cases"
source = cases / "Composition/letter.docx"
assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == "50FE6A8DDA140D39F98C9F0C4327642E90C0839A858526251060C2BEE329B9A4"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
E.register_namespace("w", w[1:-1])
with ZipFile(source) as parent:
    entries = {name: parent.read(name) for name in parent.namelist()}
document = E.fromstring(entries["word/document.xml"])
body = document.find(w + "body")
section = body.find(w + "sectPr")
body.clear()
comments = E.Element(w + "comments")
def text(parent, value):
    E.SubElement(parent, w + "t", {"{http://www.w3.org/XML/1998/namespace}space": "preserve"}).text = value
def run(parent, value):
    result = E.SubElement(parent, w + "r")
    text(result, value)
    return result
for index, (left, right, kind) in enumerate([
    ("Reviewed check.", " Supporting note", "note"),
    ("Earlier annotation.", " Separate words", "separate"),
    ("Same run.", " Following text", "same"),
    ("Multiple spaces.", "  Two spaces", "separate"),
    ("Nonbreaking gap.", "\u00a0Following word", "separate"),
    ("Line before break.", "After break. Continued words", "break"),
]):
    paragraph = E.SubElement(body, w + "p")
    E.SubElement(paragraph, w + "commentRangeStart", {w + "id": str(index)})
    first = run(paragraph, left)
    E.SubElement(paragraph, w + "commentRangeEnd", {w + "id": str(index)})
    reference = first if kind == "same" else E.SubElement(paragraph, w + "r")
    E.SubElement(reference, w + "commentReference", {w + "id": str(index)})
    if kind == "same":
        text(reference, right)
    elif kind == "separate":
        run(paragraph, right[:len(right)-len(right.lstrip())])
        run(paragraph, right.lstrip())
    elif kind == "break":
        E.SubElement(reference, w + "br")
        run(paragraph, right)
    else:
        run(paragraph, right)
        E.SubElement(E.SubElement(paragraph, w + "r"), w + "footnoteReference", {w + "id": "1"})
    comment = E.SubElement(comments, w + "comment", {w + "id": str(index), w + "author": "Public reviewer", w + "initials": "RV"})
    run(E.SubElement(comment, w + "p"), "Check authored spacing.")
body.append(section)
entries["word/document.xml"] = E.tostring(document, encoding="utf-8", xml_declaration=True)
entries["word/comments.xml"] = E.tostring(comments, encoding="utf-8", xml_declaration=True)
output = cases / "CompositionProbes/comment-authored-whitespace.docx"
with ZipFile(output, "w", ZIP_DEFLATED) as archive:
    for name, data in sorted(entries.items()):
        info = ZipInfo(name, (2026,10,15,0,0,0)); info.compress_type = ZIP_DEFLATED
        archive.writestr(info, data)
manifest = [{"Id":"comment-authored-whitespace", "Input":output.relative_to(root).as_posix(),
             "Sha256":hashlib.sha256(output.read_bytes()).hexdigest().upper(), "HeldOut":False,
             "ExpectedText":["Reviewed check.", "Supporting note", "Continued words"]}]
output.with_suffix(".json").write_text(json.dumps(manifest, indent=2)+"\n", encoding="utf-8")
print(manifest[0]["Sha256"])
