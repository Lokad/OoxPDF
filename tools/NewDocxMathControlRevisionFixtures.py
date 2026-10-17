"""Wrap existing supported controls without changing their mixed paragraph owners."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
from copy import deepcopy
import hashlib
import json

root = Path(__file__).resolve().parent.parent
cases = root / "tests/Lokad.OoxPdf.Tests/Cases"
output = cases / "CompositionProbes"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
E.register_namespace("w", w[1:-1])
E.register_namespace("m", m[1:-1])
rows = []
source = output / "structural-placement-root-body.docx"
for role, filename, node in [
    ("script", "docx-omml-isolated-inline-body-script", "sSubSup"),
    ("function", "docx-omml-function-display-sin-literal", "func"),
    ("matrix", "docx-omml-matrix-display-literal", "m"),
]:
    control_source = cases / (filename + ".docx")
    with ZipFile(source) as z:
        entries = {name: z.read(name) for name in z.namelist()}
    with ZipFile(control_source) as z:
        control = next(E.fromstring(z.read("word/document.xml")).iter(m + node))
    xml = E.fromstring(entries["word/document.xml"])
    equation = next(xml.iter(m + "oMath"))
    equation.clear()
    E.SubElement(equation, w + "ins", {
        w + "id": "981", w + "author": "Math Control Reviewer",
        w + "date": "2026-10-17T00:00:00Z",
    }).append(control)
    entries["word/document.xml"] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
    target = output / ("math-control-revision-" + role + ".docx")
    with ZipFile(target, "w", ZIP_DEFLATED) as z:
        for name, data in sorted(entries.items()):
            info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, data)
    rows.append(dict(
        Id=target.stem, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        ControlSourceSha256=hashlib.sha256(control_source.read_bytes()).hexdigest().upper(),
        Role=role, HeldOut=False, ExpectedEquationCount=4,
        ExpectedText=["share the same input series."],
    ))
(output / "math-control-revision.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Pinned three supported-control revision compositions")

# Delete actual cell text while retaining the matrix's unrevised metric owner.
# The same source transformation applies in each supported paragraph container.
cell_rows = []
for story in ["body", "cell", "header", "footer"]:
    source = output / ("structural-placement-root-" + story + ".docx")
    with ZipFile(source) as z:
        entries = {name: z.read(name) for name in z.namelist()}
    control_source = cases / "docx-omml-matrix-display-literal.docx"
    with ZipFile(control_source) as z:
        matrix = next(E.fromstring(z.read("word/document.xml")).iter(m + "m"))
    properties = matrix.find(m + "mPr")
    if properties is None:
        properties = E.Element(m + "mPr")
        matrix.insert(0, properties)
    control = E.SubElement(properties, m + "ctrlPr")
    E.SubElement(E.SubElement(control, w + "rPr"), w + "sz", {w + "val": "22"})
    for index, cell in enumerate(matrix.iter(m + "e")):
        for run in list(cell.findall(m + "r")):
            cell.remove(run)
            E.SubElement(cell, w + "del", {
                w + "id": str(991 + index), w + "author": "Math Control Reviewer",
                w + "date": "2026-10-17T00:00:00Z",
            }).append(run)
    part = "word/" + {"header": "header1", "footer": "footer1"}.get(story, "document") + ".xml"
    xml = E.fromstring(entries[part])
    equation = next(xml.iter(m + "oMath"))
    equation.clear()
    equation.append(matrix)
    entries[part] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
    target = output / ("matrix-cell-revision-" + story + ".docx")
    with ZipFile(target, "w", ZIP_DEFLATED) as z:
        for name, data in sorted(entries.items()):
            info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, data)
    cell_rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
                          Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
                          SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
                          Story=story, HeldOut=False, ExpectedEquationCount=4,
                          ExpectedText=["share the same input series."], ProsePrefixes=["Summary:", "governs"]))
(output / "matrix-cell-revisions.json").write_text(json.dumps(cell_rows, indent=2) + "\n", encoding="utf-8")
print("Pinned four literal cell revisions across body/cell/header/footer")

source = output / "math-control-revision-script.docx"
with ZipFile(source) as z:
    compound_entries = {name: z.read(name) for name in z.namelist()}
with ZipFile(output / "equation-revision-body-ins.docx") as z:
    fraction = next(E.fromstring(z.read("word/document.xml")).iter(m + "f"))
compound_rows = []
for order in ["script-first", "fraction-first"]:
    for kind in ["ordinary", "inserted"]:
        entries = dict(compound_entries)
        xml = E.fromstring(entries["word/document.xml"])
        equation = next(xml.iter(m + "oMath"))
        script = equation.find(w + "ins").find(m + "sSubSup")
        equation.clear()
        plus = E.Element(m + "r")
        E.SubElement(plus, m + "t").text = "+"
        nodes = [script, plus, deepcopy(fraction)] if order == "script-first" else [deepcopy(fraction), plus, script]
        if kind == "ordinary":
            equation.extend(nodes)
        else:
            E.SubElement(equation, w + "ins", {
                w + "id": "998", w + "author": "Compound Math Reviewer",
                w + "date": "2026-10-17T00:00:00Z",
            }).extend(nodes)
        entries["word/document.xml"] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
        target = output / ("compound-revision-" + order + "-" + kind + ".docx")
        with ZipFile(target, "w", ZIP_DEFLATED) as z:
            for name, data in sorted(entries.items()):
                info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                z.writestr(info, data)
        compound_rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
                                  Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
                                  SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
                                  Order=order, Kind=kind, HeldOut=False, ExpectedEquationCount=4,
                                  ExpectedText=["share the same input series."], ProsePrefixes=["Summary:", "governs"]))
(output / "compound-revisions.json").write_text(json.dumps(compound_rows, indent=2) + "\n", encoding="utf-8")
print("Pinned four compound control rows with and without insertion wrappers")

glyphless_rows = []
for order in ["script-first", "fraction-first"]:
    source = output / ("compound-revision-" + order + "-inserted.docx")
    with ZipFile(source) as z:
        original = {name: z.read(name) for name in z.namelist()}
    for kind in ["deleted", "authored-empty", "empty-ins"]:
        entries = dict(original)
        xml = E.fromstring(entries["word/document.xml"])
        equation = next(xml.iter(m + "oMath"))
        wrapper = equation.find(w + "ins")
        if kind == "deleted":
            wrapper.tag = w + "del"
        else:
            for parent in list(wrapper.iter()):
                for run in list(parent.findall(m + "r")):
                    parent.remove(run)
            if kind == "authored-empty":
                for child in list(wrapper):
                    wrapper.remove(child)
                    equation.append(child)
                equation.remove(wrapper)
        entries["word/document.xml"] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
        target = output / ("glyphless-compound-" + order + "-" + kind + ".docx")
        with ZipFile(target, "w", ZIP_DEFLATED) as z:
            for name, data in sorted(entries.items()):
                info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                z.writestr(info, data)
        glyphless_rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
                                   Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
                                   SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
                                   Order=order, Kind=kind, HeldOut=False, ExpectedEquationCount=4,
                                   ExpectedText=["share the same input series."], ProsePrefixes=["Summary:", "governs"],
                                   RuleRegion=[120, 680, 210, 740], RuleFillColor="g:0"))
(output / "glyphless-compounds.json").write_text(json.dumps(glyphless_rows, indent=2) + "\n", encoding="utf-8")
print("Pinned six ordered glyphless-control and deletion compositions")

# Keep ordinary painted neighbors while only the fraction is empty or revised.
mixed_rows = []
for order in ["script-first", "fraction-first"]:
    source = output / ("compound-revision-" + order + "-inserted.docx")
    with ZipFile(source) as z:
        original = {name: z.read(name) for name in z.namelist()}
    for kind in ["authored-empty", "fraction-ins", "fraction-del"]:
        entries = dict(original)
        xml = E.fromstring(entries["word/document.xml"])
        equation = next(xml.iter(m + "oMath"))
        wrapper = equation.find(w + "ins")
        for child in list(wrapper):
            wrapper.remove(child)
            equation.append(child)
        equation.remove(wrapper)
        fraction = equation.find(m + "f")
        if kind != "fraction-del":
            for parent in list(fraction.iter()):
                for run in list(parent.findall(m + "r")):
                    parent.remove(run)
        if kind != "authored-empty":
            index = list(equation).index(fraction)
            equation.remove(fraction)
            revision = E.Element(w + ("ins" if kind == "fraction-ins" else "del"), {
                w + "id": "1201", w + "author": "Mixed Control Reviewer",
                w + "date": "2026-10-17T00:00:00Z",
            })
            revision.append(fraction)
            equation.insert(index, revision)
        entries["word/document.xml"] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
        target = output / ("mixed-control-" + order + "-" + kind + ".docx")
        with ZipFile(target, "w", ZIP_DEFLATED) as z:
            for name, data in sorted(entries.items()):
                info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                z.writestr(info, data)
        mixed_rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
                               Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
                               SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
                               Order=order, Kind=kind, HeldOut=False, ExpectedEquationCount=4,
                               ExpectedText=["share the same input series."], ProsePrefixes=["Summary:", "governs"],
                               RuleRegion=[120, 680, 210, 740], RuleFillColor="g:0"))
(output / "mixed-controls.json").write_text(json.dumps(mixed_rows, indent=2) + "\n", encoding="utf-8")
print("Pinned six mixed painted/metric-control compositions")
