"""Pin deleted and authored-empty matrices with the same actual slot structure."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib
import json

root = Path(__file__).resolve().parent.parent
cases = root / "tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
E.register_namespace("w", w[1:-1])
E.register_namespace("m", m[1:-1])
source = cases / "math-control-revision-matrix.docx"
rows = []
for size in [11, 18]:
    for kind in (["deleted", "authored-empty", "implicit-empty"] if size == 11 else ["deleted", "authored-empty"]):
        with ZipFile(source) as z:
            entries = {name: z.read(name) for name in z.namelist()}
        xml = E.fromstring(entries["word/document.xml"])
        equation = next(xml.iter(m + "oMath"))
        wrapper = equation.find(w + "ins")
        matrix = wrapper.find(m + "m")
        properties = matrix.find(m + "mPr")
        if properties is None:
            properties = E.Element(m + "mPr")
            matrix.insert(0, properties)
        control = properties.find(m + "ctrlPr")
        if control is None:
            control = E.SubElement(properties, m + "ctrlPr")
        run_properties = control.find(w + "rPr")
        if run_properties is None:
            run_properties = E.SubElement(control, w + "rPr")
        for old in list(run_properties.findall(w + "sz")):
            run_properties.remove(old)
        E.SubElement(run_properties, w + "sz", {w + "val": str(size * 2)})
        if kind == "implicit-empty":
            properties.remove(control)
        if kind == "deleted":
            wrapper.tag = w + "del"
        else:
            for parent in list(matrix.iter()):
                for run in list(parent.findall(m + "r")):
                    parent.remove(run)
            wrapper.remove(matrix)
            equation.remove(wrapper)
            equation.append(matrix)
        entries["word/document.xml"] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
        target = cases / ("math-matrix-" + kind + "-" + str(size) + ".docx")
        with ZipFile(target, "w", ZIP_DEFLATED) as z:
            for name, data in sorted(entries.items()):
                info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                z.writestr(info, data)
        rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
                         Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
                         SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
                         ControlSize=size, Kind=kind, HeldOut=False, ExpectedEquationCount=4,
                         ExpectedText=["share the same input series."], ProsePrefixes=["Summary:", "governs"]))
(cases / "empty-matrix.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Pinned five matrix controls at two authored sizes, including implicit metrics")
