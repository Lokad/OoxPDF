"""Challenge structural cell range coalescing with semantic source boundaries."""
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
    for shape in ["same-author", "second-author", "gap", "later-date", "mixed-kind", "columns", "columns-second-author", "vertical-column", "diagonal-column"]:
        source = output / ("cell-connector-" + (shape if shape in ["same-author", "second-author"] else "same-author") + "-" + kind + ".docx")
        target = source
        if shape not in ["same-author", "second-author"]:
            with ZipFile(source) as z:
                entries = {name: z.read(name) for name in z.namelist()}
            xml = E.fromstring(entries["word/document.xml"])
            table = next(xml.iter(w + "tbl"))
            first, second = table.findall(w + "tr")
            cell = second.find(w + "tc")
            marker = cell.find(w + "tcPr").find(w + ("cellIns" if kind == "ins" else "cellDel"))
            if shape == "gap":
                gap = deepcopy(second)
                properties = gap.find(w + "tc").find(w + "tcPr")
                for child in list(properties):
                    if child.tag in [w + "cellIns", w + "cellDel"]:
                        properties.remove(child)
                next(gap.iter(w + "t")).text = "Intervening unchanged row."
                table.insert(list(table).index(second), gap)
            elif shape == "later-date":
                marker.set(w + "date", "2026-10-18T00:00:00Z")
            elif shape == "mixed-kind":
                marker.tag = w + ("cellDel" if kind == "ins" else "cellIns")
            else:
                if shape == "columns-second-author":
                    marker.set(w + "author", "Second Editor")
                if shape in ["vertical-column", "diagonal-column"]:
                    first_right = deepcopy(cell)
                    first_properties = first_right.find(w + "tcPr")
                    first_properties.remove(first_properties.find(w + ("cellIns" if kind == "ins" else "cellDel")))
                    next(first_right.iter(w + "t")).text = "Neighboring cell."
                    first.append(first_right)
                    second_right = deepcopy(first_right)
                    if shape == "diagonal-column":
                        cell.find(w + "tcPr").remove(marker)
                        second_right.find(w + "tcPr").append(marker)
                    second.append(second_right)
                else:
                    table.remove(second)
                    first.append(cell)
                grid = table.find(w + "tblGrid")
                grid.clear()
                for _ in range(2):
                    E.SubElement(grid, w + "gridCol", {w + "w": "4500"})
                for width in table.iter(w + "tcW"):
                    width.set(w + "w", "4500")
            entries["word/document.xml"] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
            target = output / ("cell-range-" + shape + "-" + kind + ".docx")
            with ZipFile(target, "w", ZIP_DEFLATED) as z:
                for name, data in sorted(entries.items()):
                    info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
                    info.compress_type = ZIP_DEFLATED
                    z.writestr(info, data)
        rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
            Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
            SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
            Kind=kind, Shape=shape, HeldOut=False, ExpectedEquationCount=4,
            BalloonRegion=[400, 350, 612, 792],
            ExpectedText=["share the same input series.", "Additional cell."] +
                (["Intervening unchanged row."] if shape == "gap" else []) +
                (["Neighboring cell."] if shape in ["vertical-column", "diagonal-column"] else [])))
(output / "cell-ranges.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Pinned eighteen source-adjacency/author/date/type/column controls")
