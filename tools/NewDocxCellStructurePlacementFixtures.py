"""Vary cell extent and preceding content without changing structural owners."""
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
    for shape in ["original", "tall", "following", "default-minimum", "exact", "auto"]:
        target = source
        if shape != "original":
            with ZipFile(source) as z:
                entries = {name: z.read(name) for name in z.namelist()}
            xml = E.fromstring(entries["word/document.xml"])
            table = next(xml.iter(w + "tbl"))
            row = table.find(w + "tr")
            if shape in ["tall", "default-minimum", "exact", "auto"]:
                properties = row.find(w + "trPr")
                if properties is None:
                    properties = E.Element(w + "trPr")
                    row.insert(0, properties)
                for old in properties.findall(w + "trHeight"):
                    properties.remove(old)
                attributes = {w + "val": "2880"}
                if shape != "default-minimum":
                    attributes[w + "hRule"] = {"tall": "atLeast", "exact": "exact", "auto": "auto"}[shape]
                E.SubElement(properties, w + "trHeight", attributes)
            else:
                before = E.Element(w + "tr")
                properties = E.SubElement(before, w + "trPr")
                E.SubElement(properties, w + "trHeight", {w + "val": "800", w + "hRule": "atLeast"})
                cell = E.SubElement(before, w + "tc")
                copied = deepcopy(row.find(w + "tc").find(w + "tcPr"))
                for marker in list(copied):
                    if marker.tag in [w + "cellIns", w + "cellDel"]:
                        copied.remove(marker)
                cell.append(copied)
                paragraph = E.SubElement(cell, w + "p")
                E.SubElement(E.SubElement(paragraph, w + "r"), w + "t").text = "Retained preceding row."
                table.insert(list(table).index(row), before)
            entries["word/document.xml"] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
            target = output / ("cell-structure-placement-" + shape + "-" + kind + ".docx")
            with ZipFile(target, "w", ZIP_DEFLATED) as z:
                for name, data in sorted(entries.items()):
                    info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
                    info.compress_type = ZIP_DEFLATED
                    z.writestr(info, data)
        rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
            Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
            SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
            Kind=kind, Shape=shape, HeldOut=False, ExpectedEquationCount=4,
            ExpectedText=["share the same input series."], ProsePrefixes=["Summary:", "governs"],
            RuleRegion=[120, 0, 210, 792], RuleFillColor="g:0"))
placement = [row for row in rows if row["Shape"] in ["original", "tall", "following"]]
(output / "cell-structure-placement.json").write_text(json.dumps(placement, indent=2) + "\n", encoding="utf-8")
(output / "minimum-row-placement.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Pinned six cell placement and six neighboring row-height rule controls")
