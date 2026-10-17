"""Probe exact-row occupied-frame fit, deliberate clipping, borders and following flow."""
from pathlib import Path
from xml.etree import ElementTree as E
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo
import hashlib
import json

root = Path(__file__).resolve().parent.parent
output = root / "tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
E.register_namespace("w", w[1:-1])
E.register_namespace("m", "http://schemas.openxmlformats.org/officeDocument/2006/math")
rows = []
for kind in ["ins", "del"]:
    source = output / ("cell-structure-placement-exact-" + kind + ".docx")
    for shape in ["original", "96", "72", "24", "border", "flow"]:
        target = source
        if shape != "original":
            with ZipFile(source) as z:
                entries = {name: z.read(name) for name in z.namelist()}
            xml = E.fromstring(entries["word/document.xml"])
            body = xml.find(w + "body")
            row = next(xml.iter(w + "tr"))
            height = row.find(w + "trPr").find(w + "trHeight")
            height.set(w + "val", str(int(shape) * 20) if shape.isdigit() else "2880")
            if shape == "border":
                properties = row.find(w + "tc").find(w + "tcPr")
                borders = properties.find(w + "tcBorders")
                if borders is not None:
                    properties.remove(borders)
                borders = E.SubElement(properties, w + "tcBorders")
                for edge in ["top", "bottom", "left", "right"]:
                    E.SubElement(borders, w + edge, {w + "val": "single", w + "sz": "8", w + "color": "000000"})
            elif shape == "flow":
                paragraph = E.Element(w + "p")
                properties = E.SubElement(paragraph, w + "pPr")
                E.SubElement(properties, w + "spacing", {w + "before": "0", w + "after": "0"})
                run = E.SubElement(paragraph, w + "r")
                properties = E.SubElement(run, w + "rPr")
                E.SubElement(properties, w + "rFonts", {w + "ascii": "Arial", w + "hAnsi": "Arial"})
                E.SubElement(properties, w + "sz", {w + "val": "22"})
                E.SubElement(run, w + "t").text = "Following exact row."
                body.insert(list(body).index(body.find(w + "sectPr")), paragraph)
            entries["word/document.xml"] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
            target = output / ("exact-row-frame-" + shape + "-" + kind + ".docx")
            with ZipFile(target, "w", ZIP_DEFLATED) as z:
                for name, data in sorted(entries.items()):
                    info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
                    info.compress_type = ZIP_DEFLATED
                    z.writestr(info, data)
        clipped = shape in ["72", "24"]
        rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
            Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
            SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
            Kind=kind, Shape=shape, HeldOut=False, ExpectedEquationCount=4,
            ContentMustFit=not clipped,
            ImageCount=0 if shape == "24" else 1,
            VisibilityContract="Office deliberately clips source content to the exact cell; raw text is not visible-content proof" if clipped else "Complete contained content",
            ExpectedText=[] if clipped else ["share the same input series."] + (["Following exact row."] if shape == "flow" else []),
            BalloonRegion=[400, 350, 612, 792],
            ProsePrefixes=["Following exact row."] if shape == "flow" else []))
(output / "exact-row-frames.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Pinned twelve exact-height fit/clip/border/flow controls")
