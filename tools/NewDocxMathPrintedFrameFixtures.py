"""Create scaled/unscaled mixed-math controls with independent first-body anchors."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib, json, argparse
from copy import deepcopy
parser = argparse.ArgumentParser()
parser.add_argument("--multiple", action="store_true")
parser.add_argument("--script-only", action="store_true", help="Generate script-frame controls without rewriting other fixture families")
parser.add_argument("--cell-context", choices=["sole", "flow"])
parser.add_argument("--leading-content", choices=["math", "image"])
args = parser.parse_args()
if args.leading_content and not args.cell_context:
    parser.error("Leading-content probes require a cell context")
if args.multiple and args.cell_context:
    parser.error("Use separate source controls for repeated formulas and cell contexts")
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
if args.script_only:
    sources = [("script", "docx-omml-inline-script-middle12.docx", "9EA3F9FE29911876A39983D2222C824A0B27E75883B19074B7EA654F20EB8769")]
for kind, filename, expected_sha in sources:
    if args.leading_content and kind != "fraction":
        continue
    if args.multiple and kind == "root-fraction":
        continue
    source = cases / filename
    assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == expected_sha
    for heading_size in ((12,) if args.cell_context else (12, 18)):
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
        if args.cell_context:
            math = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
            paragraph = body.find(w + "p")
            equation = paragraph.find(math + "oMath")
            properties = paragraph.find(w + "pPr")
            paragraph.clear()
            if properties is not None:
                paragraph.append(properties)
            paragraph.append(equation)
            body.remove(paragraph)
            table = E.Element(w + "tbl")
            table_properties = E.SubElement(table, w + "tblPr")
            E.SubElement(table_properties, w + "tblW", {w + "w":"9360", w + "type":"dxa"})
            E.SubElement(table_properties, w + "tblLayout", {w + "type":"fixed"})
            borders = E.SubElement(table_properties, w + "tblBorders")
            for edge in ["top", "left", "bottom", "right", "insideH", "insideV"]:
                E.SubElement(borders, w + edge, {w + "val":"single", w + "sz":"4", w + "color":"000000"})
            grid = E.SubElement(table, w + "tblGrid")
            E.SubElement(grid, w + "gridCol", {w + "w":"4680"})
            E.SubElement(grid, w + "gridCol", {w + "w":"4680"})
            row = E.SubElement(table, w + "tr")
            for index in range(2):
                cell = E.SubElement(row, w + "tc")
                cell_properties = E.SubElement(cell, w + "tcPr")
                E.SubElement(cell_properties, w + "tcW", {w + "w":"4680", w + "type":"dxa"})
                margins = E.SubElement(cell_properties, w + "tcMar")
                for edge in ["top", "left", "bottom", "right"]:
                    E.SubElement(margins, w + edge, {w + "w":"80", w + "type":"dxa"})
                def cell_text(value):
                    E.SubElement(E.SubElement(E.SubElement(cell, w + "p"), w + "r"), w + "t").text = value
                if index == 0:
                    cell_text("Observed series.")
                else:
                    if args.cell_context == "flow":
                        cell_text("Cell preface.")
                    cell.append(paragraph)
                    if args.cell_context == "flow":
                        cell_text("Cell afterword.")
            body.insert(0, table)
        leading = E.Element(w + "p")
        run = E.SubElement(leading, w + "r")
        E.SubElement(E.SubElement(run, w + "rPr"), w + "sz", {w + "val": str(heading_size * 2)})
        E.SubElement(run, w + "t").text = "Leading paragraph."
        if args.leading_content == "math":
            math = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
            for index, value in enumerate(["a+b", "c"]):
                separator = E.SubElement(leading, w + "r")
                E.SubElement(separator, w + "t", {"{http://www.w3.org/XML/1998/namespace}space":"preserve"}).text = " and " if index else " "
                equation = E.SubElement(leading, math + "oMath")
                math_run = E.SubElement(equation, math + "r")
                E.SubElement(E.SubElement(math_run, w + "rPr"), w + "rFonts", {w + "ascii":"Cambria Math", w + "hAnsi":"Cambria Math"})
                E.SubElement(math_run, math + "t").text = value
        elif args.leading_content == "image":
            image_source = cases / "Composition/report.docx"
            assert hashlib.sha256(image_source.read_bytes()).hexdigest().upper() == "10F63D162D0D46DBE8CD682BB50452DB6672104706F116EA50CDCAFEE054575F"
            rel = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
            with ZipFile(image_source) as archive:
                image_document = E.fromstring(archive.read("word/document.xml"))
                image_run = deepcopy(next(run for run in image_document.iter(w + "r") if run.find(w + "drawing") is not None))
                for node in image_run.iter():
                    if rel + "embed" in node.attrib:
                        source_id = node.attrib[rel + "embed"]
                        node.set(rel + "embed", "leadingImage")
                image_rels = E.fromstring(archive.read("word/_rels/document.xml.rels"))
                image_target = next(node.attrib["Target"] for node in image_rels if node.attrib["Id"] == source_id)
                image_bytes = archive.read("word/" + image_target)
            entries["word/media/leading-image.png"] = image_bytes
            relationships = E.fromstring(entries["word/_rels/document.xml.rels"])
            E.SubElement(relationships, package + "Relationship", {"Id":"leadingImage", "Type":rel[1:-1] + "/image", "Target":"media/leading-image.png"})
            entries["word/_rels/document.xml.rels"] = E.tostring(relationships, encoding="utf-8", xml_declaration=True)
            content_types = E.fromstring(entries["[Content_Types].xml"])
            if not any(node.attrib.get("Extension") == "png" for node in content_types):
                E.SubElement(content_types, types + "Default", {"Extension":"png", "ContentType":"image/png"})
            entries["[Content_Types].xml"] = E.tostring(content_types, encoding="utf-8", xml_declaration=True)
            leading.append(image_run)
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
        identifier = "math-printed-frame-" + kind + "-anchor" + str(heading_size) + ("-multiple" if args.multiple else "") + ("-cell-" + args.cell_context if args.cell_context else "") + ("-leading-" + args.leading_content if args.leading_content else "")
        output = cases / "CompositionProbes" / (identifier + ".docx")
        with ZipFile(output, "w", ZIP_DEFLATED) as archive:
            for name, data in sorted(entries.items()):
                info = ZipInfo(name, (2026,10,15,0,0,0)); info.compress_type = ZIP_DEFLATED
                archive.writestr(info, data)
        manifest.append({"Id":identifier, "Input":output.relative_to(root).as_posix(), "Sha256":hashlib.sha256(output.read_bytes()).hexdigest().upper(), "SourceSha256":expected_sha, "HeldOut":False, "ExpectedText":["Leading paragraph.", "Observed series." if args.cell_context else "Before formula:", "Following prose."]})
(cases / ("CompositionProbes/math-printed-frame" + ("-script" if args.script_only else "") + ("-multiple" if args.multiple else "") + ("-cell-" + args.cell_context if args.cell_context else "") + ("-leading-" + args.leading_content if args.leading_content else "") + ".json")).write_text(json.dumps(manifest,indent=2)+"\n",encoding="utf-8")
print(str(len(manifest)) + " source controls pinned")
