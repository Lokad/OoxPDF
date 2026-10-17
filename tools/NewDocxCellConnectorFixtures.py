"""Vary cell width and author ownership while retaining the original mixed content."""
from copy import deepcopy
import hashlib
import json
from pathlib import Path
from xml.etree import ElementTree as E
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

root = Path(__file__).resolve().parent.parent
output = root / "tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
E.register_namespace("w", w[1:-1])
E.register_namespace("m", "http://schemas.openxmlformats.org/officeDocument/2006/math")
rows = []
for kind in ["ins", "del"]:
    source = output / ("table-math-revision-cell-" + kind + ".docx")
    for shape in ["original", "narrow", "same-author", "second-author"]:
        target = source
        if shape != "original":
            with ZipFile(source) as z:
                entries = {name: z.read(name) for name in z.namelist()}
            xml = E.fromstring(entries["word/document.xml"])
            table = next(xml.iter(w + "tbl"))
            first = table.find(w + "tr")
            if shape == "narrow":
                for width in table.iter():
                    if width.tag in [w + "tblW", w + "tcW", w + "gridCol"]:
                        width.set(w + "w", "6000")
            else:
                properties = first.find(w + "trPr")
                if properties is None:
                    properties = E.Element(w + "trPr")
                    first.insert(0, properties)
                E.SubElement(properties, w + "trHeight", {w + "val": "2880", w + "hRule": "atLeast"})
                second = E.SubElement(table, w + "tr")
                properties = E.SubElement(second, w + "trPr")
                E.SubElement(properties, w + "trHeight", {w + "val": "800", w + "hRule": "atLeast"})
                cell = E.SubElement(second, w + "tc")
                properties = deepcopy(first.find(w + "tc").find(w + "tcPr"))
                marker = properties.find(w + ("cellIns" if kind == "ins" else "cellDel"))
                marker.set(w + "id", "713")
                marker.set(w + "author", "Second Editor" if shape == "second-author" else "Table Editor")
                cell.append(properties)
                paragraph = E.SubElement(cell, w + "p")
                run = E.SubElement(paragraph, w + "r")
                rpr = E.SubElement(run, w + "rPr")
                E.SubElement(rpr, w + "rFonts", {w + "ascii": "Arial", w + "hAnsi": "Arial"})
                E.SubElement(rpr, w + "sz", {w + "val": "22"})
                E.SubElement(run, w + "t").text = "Additional cell."
            entries["word/document.xml"] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
            target = output / ("cell-connector-" + shape + "-" + kind + ".docx")
            with ZipFile(target, "w", ZIP_DEFLATED) as z:
                for name, data in sorted(entries.items()):
                    info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
                    info.compress_type = ZIP_DEFLATED
                    z.writestr(info, data)
        rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
            Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
            SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(), Kind=kind, Shape=shape,
            HeldOut=False, ExpectedEquationCount=4, ExpectedText=["share the same input series."] +
                (["Additional cell."] if shape in ["same-author", "second-author"] else []),
            ProsePrefixes=["Summary:", "governs"], RuleRegion=[120, 0, 210, 792], RuleFillColor="g:0"))
(output / "cell-connectors.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Pinned eight structural cell width/author controls")
