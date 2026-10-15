"""Pin an Office control for absent, empty and explicit comment initials."""
from copy import deepcopy
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET

repo = Path(__file__).resolve().parents[1]
cases = repo / "tests/Lokad.OoxPdf.Tests/Cases"
source = cases / "Composition/report.docx"
source_sha = hashlib.sha256(source.read_bytes()).hexdigest().upper()
assert source_sha == "10F63D162D0D46DBE8CD682BB50452DB6672104706F116EA50CDCAFEE054575F"
output = cases / "CompositionProbes/comment-labels.docx"
manifest = output.with_suffix(".json")
if output.exists():
    pin = json.loads(manifest.read_text(encoding="utf-8"))[0]
    assert pin["SourceSha256"] == source_sha
    assert pin["Sha256"] == hashlib.sha256(output.read_bytes()).hexdigest().upper()
    raise SystemExit(0)
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
ET.register_namespace("w", w[1:-1])
metadata = [(50, "Public reviewer", None), (10, "Public reviewer", ""),
            (90, "", "RV"), (7, "", None), (42, "Alternate author", "QA")]
comments = ET.Element(w + "comments")
for index, (identifier, author, initials) in enumerate(metadata, 1):
    attrs = {w + "id": str(identifier), w + "author": author}
    if initials is not None:
        attrs[w + "initials"] = initials
    comment = ET.SubElement(comments, w + "comment", attrs)
    ET.SubElement(ET.SubElement(ET.SubElement(comment, w + "p"), w + "r"), w + "t").text = f"Synthetic review note {index}."
with ZipFile(source) as parent, ZipFile(output, "w", ZIP_DEFLATED) as result:
    for entry in parent.infolist():
        data = parent.read(entry)
        if entry.filename == "word/document.xml":
            doc = ET.fromstring(data)
            body = doc.find(w + "body")
            section = deepcopy(body.find(w + "sectPr"))
            for child in list(section):
                if child.tag in (w + "headerReference", w + "footerReference"):
                    section.remove(child)
            body.clear()
            for index, (identifier, _, _) in enumerate(metadata, 1):
                paragraph = ET.SubElement(body, w + "p")
                ET.SubElement(paragraph, w + "commentRangeStart", {w + "id": str(identifier)})
                ET.SubElement(ET.SubElement(paragraph, w + "r"), w + "t").text = f"Public comment anchor {index}."
                ET.SubElement(paragraph, w + "commentRangeEnd", {w + "id": str(identifier)})
                ET.SubElement(ET.SubElement(paragraph, w + "r"), w + "commentReference", {w + "id": str(identifier)})
            body.append(section)
            data = ET.tostring(doc, encoding="utf-8", xml_declaration=True)
        elif entry.filename == "word/comments.xml":
            data = ET.tostring(comments, encoding="utf-8", xml_declaration=True)
        result.writestr(entry, data)
manifest.write_text(json.dumps([{
    "Id": "comment-labels", "Input": output.relative_to(repo).as_posix(),
    "Sha256": hashlib.sha256(output.read_bytes()).hexdigest().upper(),
    "SourceSha256": source_sha, "HeldOut": False, "ExpectedEquationCount": 0,
    "ExpectedInlineImageCount": 0, "ExpectedText": [f"Public comment anchor {i}." for i in range(1, 6)]
}], indent=2) + "\n", encoding="utf-8")
