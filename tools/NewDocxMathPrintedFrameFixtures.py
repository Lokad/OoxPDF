"""Create scaled/unscaled mixed-math controls with independent first-body anchors."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib, json, argparse
from copy import deepcopy
parser = argparse.ArgumentParser()
parser.add_argument("--multiple", action="store_true")
args = parser.parse_args()
root = Path(__file__).resolve().parent.parent
cases = root / "tests/Lokad.OoxPdf.Tests/Cases"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
package = "{http://schemas.openxmlformats.org/package/2006/relationships}"
types = "{http://schemas.openxmlformats.org/package/2006/content-types}"
E.register_namespace("w", w[1:-1])
manifest = []
sources = [('literal', 'docx-omml-literal-inline.docx', '16CFF81A67E2EBC2DDE05873CE789C8FB79901563F8EB0A51D09F9846894F18D'),
           ('fraction', 'docx-omml-normal-fraction-inline-arial-normal.docx', 'DD4907536AF7C2EFCF99A0B0EEB6AB877A453229AB4E57313E17270790F97871'),
           ('root-fraction', 'docx-omml-root-fraction-inline-bar-square.docx', '54DF93237B68A46DA4AE66455848EC9B01E729CFEC63BDA380102628923DDBF3')]
for kind, filename, expected_sha in sources:
    if args.multiple and kind == "root-fraction":
        continue
    source = cases / filename
    assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == expected_sha
    for heading_size in (12, 18):
        with ZipFile(source) as archive:
            entries = {name: archive.read(name) for name in archive.namelist()}
        document = E.fromstring(entries["word/document.xml"])
        body = document.find(w + "body")
        if args.multiple:
            math = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
            paragraph = body.find(w + "p")
            equation = paragraph.find(math + "oMath")
            insert = list(paragraph).index(equation) + 1
            separator = E.Element(w + "r")
            E.SubElement(separator, w + "t", {"{http://www.w3.org/XML/1998/namespace}space":"preserve"}).text = " and "
            paragraph.insert(insert, separator)
            paragraph.insert(insert + 1, deepcopy(equation))
        leading = E.Element(w + "p")
        run = E.SubElement(leading, w + "r")
        E.SubElement(E.SubElement(run, w + "rPr"), w + "sz", {w + "val": str(heading_size * 2)})
        E.SubElement(run, w + "t").text = "Leading paragraph."
        body.insert(0, leading)
        last = body.findall(w + "p")[-1]
        E.SubElement(last, w + "commentRangeStart", {w + "id": "0"})
        E.SubElement(E.SubElement(last, w + "r"), w + "t").text = " Review anchor."
        E.SubElement(last, w + "commentRangeEnd", {w + "id": "0"})
        E.SubElement(E.SubElement(last, w + "r"), w + "commentReference", {w + "id": "0"})
        comments = E.Element(w + "comments")
        comment = E.SubElement(comments, w + "comment", {w + "id": "0", w + "author": "Public reviewer", w + "initials": "RV"})
        E.SubElement(E.SubElement(E.SubElement(comment, w + "p"), w + "r"), w + "t").text = "Check printed geometry."
        rels_name = "word/_rels/document.xml.rels"
        relationships = E.fromstring(entries[rels_name]) if rels_name in entries else E.Element(package + "Relationships")
        E.SubElement(relationships, package + "Relationship", {"Id":"printedComment", "Type":"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments", "Target":"comments.xml"})
        content_types = E.fromstring(entries["[Content_Types].xml"])
        E.SubElement(content_types, types + "Override", {"PartName":"/word/comments.xml", "ContentType":"application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"})
        for name, value in [("word/document.xml",document),("word/comments.xml",comments),(rels_name,relationships),("[Content_Types].xml",content_types)]:
            entries[name] = E.tostring(value, encoding="utf-8", xml_declaration=True)
        identifier = "math-printed-frame-" + kind + "-anchor" + str(heading_size) + ("-multiple" if args.multiple else "")
        output = cases / "CompositionProbes" / (identifier + ".docx")
        with ZipFile(output, "w", ZIP_DEFLATED) as archive:
            for name, data in sorted(entries.items()):
                info = ZipInfo(name, (2026,10,15,0,0,0)); info.compress_type = ZIP_DEFLATED
                archive.writestr(info, data)
        manifest.append({"Id":identifier, "Input":output.relative_to(root).as_posix(), "Sha256":hashlib.sha256(output.read_bytes()).hexdigest().upper(), "SourceSha256":expected_sha, "HeldOut":False, "ExpectedText":["Leading paragraph.", "Before formula:", "Following prose."]})
(cases / ("CompositionProbes/math-printed-frame" + ("-multiple" if args.multiple else "") + ".json")).write_text(json.dumps(manifest,indent=2)+"\n",encoding="utf-8")
print(str(len(manifest)) + " source controls pinned")
