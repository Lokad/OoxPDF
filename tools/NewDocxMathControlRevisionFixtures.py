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

# Move the same mixed contents into existing supported paragraph containers.
placement_rows = []
for story in ["cell", "header", "footer"]:
    source = output / ("structural-placement-root-" + story + ".docx")
    with ZipFile(source) as z:
        original = {name: z.read(name) for name in z.namelist()}
    for order in ["script-first", "fraction-first"]:
        content_source = output / ("mixed-control-" + order + "-fraction-ins.docx")
        with ZipFile(content_source) as z:
            content = next(E.fromstring(z.read("word/document.xml")).iter(m + "oMath"))
        entries = dict(original)
        part = "word/" + {"header": "header1", "footer": "footer1"}.get(story, "document") + ".xml"
        xml = E.fromstring(entries[part])
        equation = next(xml.iter(m + "oMath"))
        equation.clear()
        equation.extend(deepcopy(list(content)))
        entries[part] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
        target = output / ("mixed-placement-" + story + "-" + order + ".docx")
        with ZipFile(target, "w", ZIP_DEFLATED) as z:
            for name, data in sorted(entries.items()):
                info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                z.writestr(info, data)
        placement_rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
                                   Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
                                   SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
                                   ContentSourceSha256=hashlib.sha256(content_source.read_bytes()).hexdigest().upper(),
                                   Story=story, Order=order, HeldOut=False, ExpectedEquationCount=4,
                                   ExpectedText=["share the same input series."], ProsePrefixes=["Summary:", "governs"],
                                   RuleRegion=[120, 0, 210, 792], RuleFillColor="g:0"))
(output / "mixed-control-placement.json").write_text(json.dumps(placement_rows, indent=2) + "\n", encoding="utf-8")
print("Pinned six mixed-control cell/static-story placements")

# One literal script box in either fraction argument; deeper nesting stays explicit.
source = output / "compound-revision-script-first-ordinary.docx"
with ZipFile(source) as z:
    original = {name: z.read(name) for name in z.namelist()}
fraction_script_rows = []
for display in [False, True]:
    for role in ["numerator", "denominator"]:
        entries = dict(original)
        xml = E.fromstring(entries["word/document.xml"])
        equation = next(xml.iter(m + "oMath"))
        script = deepcopy(next(equation.iter(m + "sSubSup")))
        fraction = deepcopy(next(equation.iter(m + "f")))
        argument = fraction.find(m + ("num" if role == "numerator" else "den"))
        argument.clear()
        argument.append(script)
        equation.clear()
        equation.append(fraction)
        if display:
            paragraph = next(xml.iter(w + "p"))
            paragraph.clear()
            E.SubElement(paragraph, m + "oMathPara").append(equation)
        entries["word/document.xml"] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
        target = output / ("fraction-script-" + ("display" if display else "inline") + "-" + role + ".docx")
        with ZipFile(target, "w", ZIP_DEFLATED) as z:
            for name, data in sorted(entries.items()):
                info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                z.writestr(info, data)
        fraction_script_rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
                                         Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
                                         SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
                                         Style="display" if display else "inline", Role=role, HeldOut=False,
                                         ExpectedEquationCount=4, ExpectedText=["share the same input series."],
                                         ProsePrefixes=[] if display else ["Summary:", "governs"],
                                         RuleRegion=[0, 695, 612, 760], RuleFillColor="g:0"))
(output / "fraction-script-arguments.json").write_text(json.dumps(fraction_script_rows, indent=2) + "\n", encoding="utf-8")
print("Pinned four bounded fraction/script argument compositions")

# Keep the visible sibling while an admitted script argument has only metrics.
nested_metric_rows = []
for row in fraction_script_rows:
    source = root / row["Input"]
    with ZipFile(source) as z:
        entries = {name: z.read(name) for name in z.namelist()}
    xml = E.fromstring(entries["word/document.xml"])
    script = next(next(xml.iter(m + "oMath")).iter(m + "sSubSup"))
    for parent in list(script.iter()):
        for run in list(parent.findall(m + "r")):
            parent.remove(run)
    entries["word/document.xml"] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
    target = output / ("nested-metric-" + row["Style"] + "-" + row["Role"] + ".docx")
    with ZipFile(target, "w", ZIP_DEFLATED) as z:
        for name, data in sorted(entries.items()):
            info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, data)
    nested_metric_rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
                                   Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
                                   SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
                                   Style=row["Style"], Role=row["Role"], HeldOut=False,
                                   ExpectedEquationCount=4, ExpectedText=row["ExpectedText"],
                                   ProsePrefixes=row["ProsePrefixes"], RuleRegion=row["RuleRegion"], RuleFillColor="g:0"))
(output / "nested-metric-controls.json").write_text(json.dumps(nested_metric_rows, indent=2) + "\n", encoding="utf-8")
print("Pinned four nested metric-only argument compositions")

# Retain both occupied argument frames even when the complete fraction has no inked text.
glyphless_fraction_rows = []
for row in nested_metric_rows:
    source = root / row["Input"]
    with ZipFile(source) as z:
        entries = {name: z.read(name) for name in z.namelist()}
    xml = E.fromstring(entries["word/document.xml"])
    equation = next(xml.iter(m + "oMath"))
    for parent in list(equation.iter()):
        for run in list(parent.findall(m + "r")):
            parent.remove(run)
    entries["word/document.xml"] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
    target = output / ("glyphless-fraction-" + row["Style"] + "-" + row["Role"] + ".docx")
    with ZipFile(target, "w", ZIP_DEFLATED) as z:
        for name, data in sorted(entries.items()):
            info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, data)
    glyphless_fraction_rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
                                        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
                                        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
                                        Style=row["Style"], Role=row["Role"], HeldOut=False,
                                        ExpectedEquationCount=4, ExpectedText=row["ExpectedText"], ProsePrefixes=row["ProsePrefixes"],
                                        RuleRegion=[285, 0, 330, 792] if row["Style"] == "display" else [120, 0, 210, 792],
                                        RuleFillColor="g:0"))
(output / "glyphless-fraction-arguments.json").write_text(json.dumps(glyphless_fraction_rows, indent=2) + "\n", encoding="utf-8")
print("Pinned four glyphless fraction argument compositions")

# Existing script boxes compose inside the same delimiter argument contract.
source = output / "compound-revision-script-first-ordinary.docx"
with ZipFile(source) as z:
    original = {name: z.read(name) for name in z.namelist()}
delimiter_script_rows = []
for role in ["sub", "sup", "combined", "fraction-base"]:
    entries = dict(original)
    xml = E.fromstring(entries["word/document.xml"])
    equation = next(xml.iter(m + "oMath"))
    script = deepcopy(next(equation.iter(m + "sSubSup")))
    fraction = deepcopy(next(equation.iter(m + "f")))
    if role in ["sub", "sup"]:
        script.tag = m + ("sSub" if role == "sub" else "sSup")
        script.remove(script.find(m + ("sup" if role == "sub" else "sub")))
    elif role == "fraction-base":
        basis = script.find(m + "e")
        basis.clear()
        basis.append(fraction)
    delimiter = E.Element(m + "d")
    properties = E.SubElement(delimiter, m + "dPr")
    E.SubElement(properties, m + "begChr", {m + "val": "("})
    E.SubElement(properties, m + "endChr", {m + "val": ")"})
    E.SubElement(delimiter, m + "e").append(script)
    equation.clear()
    equation.append(delimiter)
    entries["word/document.xml"] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
    target = output / ("delimiter-script-" + role + ".docx")
    with ZipFile(target, "w", ZIP_DEFLATED) as z:
        for name, data in sorted(entries.items()):
            info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, data)
    delimiter_script_rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
                                      Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
                                      SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
                                      Role=role, HeldOut=False, ExpectedEquationCount=4,
                                      ExpectedText=["share the same input series."], ProsePrefixes=["Summary:", "governs"]))
(output / "delimiter-scripts.json").write_text(json.dumps(delimiter_script_rows, indent=2) + "\n", encoding="utf-8")
print("Pinned four delimiter/script compositions")

# Supported child boxes retain their own metrics in every existing argument container.
source = output / "compound-revision-script-first-ordinary.docx"
with ZipFile(source) as z:
    original = {name: z.read(name) for name in z.namelist()}
source_equation = next(E.fromstring(original["word/document.xml"]).iter(m + "oMath"))
script = deepcopy(next(source_equation.iter(m + "sSubSup")))
fraction = deepcopy(next(source_equation.iter(m + "f")))
empty_script, empty_fraction = deepcopy(script), deepcopy(fraction)
for node in [empty_script, empty_fraction]:
    for parent in list(node.iter()):
        for run in list(parent.findall(m + "r")):
            parent.remove(run)
with ZipFile(cases / "docx-omml-function-display-sin-literal.docx") as z:
    child_function = next(E.fromstring(z.read("word/document.xml")).iter(m + "func"))
with ZipFile(cases / "docx-omml-matrix-display-literal.docx") as z:
    child_matrix = next(E.fromstring(z.read("word/document.xml")).iter(m + "m"))
def child_delimiter(child):
    node = E.Element(m + "d")
    properties = E.SubElement(node, m + "dPr")
    E.SubElement(properties, m + "begChr", {m + "val": "("})
    E.SubElement(properties, m + "endChr", {m + "val": ")"})
    E.SubElement(node, m + "e").append(deepcopy(child))
    return node
def child_application(child):
    node = deepcopy(child_function)
    argument = node.find(m + "e")
    argument.clear()
    argument.append(deepcopy(child))
    return node
def child_root(child):
    node = E.Element(m + "rad")
    E.SubElement(E.SubElement(node, m + "radPr"), m + "degHide", {m + "val": "1"})
    E.SubElement(node, m + "deg")
    E.SubElement(node, m + "e").append(deepcopy(child))
    return node
child_nary = E.Element(m + "nary")
properties = E.SubElement(child_nary, m + "naryPr")
for name, value in [("chr", "\u2211"), ("limLoc", "subSup"), ("subHide", "1"), ("supHide", "1")]:
    E.SubElement(properties, m + name, {m + "val": value})
E.SubElement(child_nary, m + "sub")
E.SubElement(child_nary, m + "sup")
E.SubElement(child_nary, m + "e").append(deepcopy(empty_script))
cell = next(child_matrix.iter(m + "e"))
cell.clear()
cell.append(deepcopy(empty_script))
script_base = deepcopy(script)
argument = script_base.find(m + "e")
argument.clear()
argument.append(deepcopy(empty_fraction))
with ZipFile(cases / "docx-omml-matrix-display-literal.docx") as z:
    fraction_matrix = next(E.fromstring(z.read("word/document.xml")).iter(m + "m"))
cell = next(fraction_matrix.iter(m + "e"))
cell.clear()
cell.append(deepcopy(empty_fraction))
child_metric_rows = []
for role, control, counter in [
    ("delimiter-script", child_delimiter(empty_script), False),
    ("function-script", child_application(empty_script), False),
    ("root-script", child_root(empty_script), False),
    ("operator-script", child_nary, False),
    ("matrix-script", child_matrix, False),
    ("script-fraction", script_base, False),
    ("delimiter-fraction", child_delimiter(empty_fraction), True),
    ("function-fraction", child_application(empty_fraction), True),
    ("root-fraction", child_root(empty_fraction), True),
    ("matrix-fraction", fraction_matrix, False),
]:
    entries = dict(original)
    xml = E.fromstring(entries["word/document.xml"])
    equation = next(xml.iter(m + "oMath"))
    equation.clear()
    equation.append(control)
    entries["word/document.xml"] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
    target = output / ("child-metric-" + role + ".docx")
    with ZipFile(target, "w", ZIP_DEFLATED) as z:
        for name, data in sorted(entries.items()):
            info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, data)
    child_metric_rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
                                  Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
                                  SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
                                  Role=role, Counterfactual=counter, HeldOut=False, ExpectedEquationCount=4,
                                  ExpectedText=["share the same input series."], ProsePrefixes=["Summary:", "governs"]))
(output / "child-metric-controls.json").write_text(json.dumps(child_metric_rows, indent=2) + "\n", encoding="utf-8")
print("Pinned ten shared child-metric compositions")

# Ordinary authored control properties use the same owner grammar in fractions
# and scripts. Keep both hidden argument frames and painted child sources.
explicit_control_rows = []
namespace_map = dict(E._namespace_map)
E._namespace_map.pop(w[1:-1], None)
E._namespace_map.pop(m[1:-1], None)
try:
    for role in ["script", "fraction"]:
        source = output / ("child-metric-matrix-" + role + ".docx")
        with ZipFile(source) as z:
            original = {name: z.read(name) for name in z.namelist()}
        for painted in [False, True]:
            for size in [11, 12]:
                entries = dict(original)
                xml = E.fromstring(entries["word/document.xml"])
                matrix = next(xml.iter(m + "m"))
                cell = matrix.find(m + "mr").find(m + "e")
                if painted:
                    cell.clear()
                    cell.append(deepcopy(script if role == "script" else fraction))
                control = list(cell)[0]
                properties = E.Element(control.tag + "Pr")
                owner = E.SubElement(properties, m + "ctrlPr")
                run = E.SubElement(owner, w + "rPr")
                E.SubElement(run, w + "rFonts", {w + "ascii": "Cambria Math", w + "hAnsi": "Cambria Math"})
                E.SubElement(run, w + "sz", {w + "val": str(size * 2)})
                control.insert(0, properties)
                entries["word/document.xml"] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
                target = output / ("explicit-control-" + role + ("-painted-" if painted else "-empty-") + str(size) + ".docx")
                with ZipFile(target, "w", ZIP_DEFLATED) as z:
                    for name, data in sorted(entries.items()):
                        info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
                        info.compress_type = ZIP_DEFLATED
                        z.writestr(info, data)
                explicit_control_rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
                    Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
                    SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
                    Role=role, Painted=painted, ControlSize=size, HeldOut=False, ExpectedEquationCount=4,
                    ExpectedText=["share the same input series."], ProsePrefixes=["Summary:", "governs"]))
finally:
    E._namespace_map.clear()
    E._namespace_map.update(namespace_map)
(output / "explicit-math-controls.json").write_text(json.dumps(explicit_control_rows, indent=2) + "\n", encoding="utf-8")
print("Pinned eight explicit fraction/script control compositions")

# Reuse the same authored controls under cell and direct static-story constraints.
explicit_placement_rows = []
for story in ["cell", "header", "footer"]:
    source = output / ("structural-placement-root-" + story + ".docx")
    with ZipFile(source) as z:
        original = {name: z.read(name) for name in z.namelist()}
    for role, paint, size in [("script", "empty", 11), ("fraction", "painted", 12)]:
        content_source = output / ("explicit-control-" + role + "-" + paint + "-" + str(size) + ".docx")
        with ZipFile(content_source) as z:
            content = next(E.fromstring(z.read("word/document.xml")).iter(m + "oMath"))
        entries = dict(original)
        part = "word/" + {"header": "header1", "footer": "footer1"}.get(story, "document") + ".xml"
        xml = E.fromstring(entries[part])
        equation = next(xml.iter(m + "oMath"))
        equation.clear()
        equation.extend(deepcopy(list(content)))
        entries[part] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
        target = output / ("explicit-placement-" + story + "-" + role + ".docx")
        with ZipFile(target, "w", ZIP_DEFLATED) as z:
            for name, data in sorted(entries.items()):
                info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                z.writestr(info, data)
        explicit_placement_rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
            Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
            SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
            ContentSourceSha256=hashlib.sha256(content_source.read_bytes()).hexdigest().upper(),
            Story=story, Role=role, HeldOut=False, ExpectedEquationCount=4,
            ExpectedText=["share the same input series."], ProsePrefixes=["Summary:", "governs"]))
        if role == "fraction":
            explicit_placement_rows[-1].update(RuleRegion=[120, 0, 210, 792], RuleFillColor="g:0")
(output / "explicit-control-placement.json").write_text(json.dumps(explicit_placement_rows, indent=2) + "\n", encoding="utf-8")
print("Pinned six authored-control cell/static-story placements")
