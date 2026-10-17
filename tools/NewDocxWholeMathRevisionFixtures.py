"""Pin whole-equation revisions without changing their existing mixed content."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib
import json

root = Path(__file__).resolve().parent.parent
cases = root / "tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes"
m = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
E.register_namespace("m", m[1:-1])
E.register_namespace("w", w[1:-1])


def fixture(role, kind, empty=False, story="body"):
    source_role = "root" if role == "root-degree" else role
    source = cases / ("structural-placement-" + source_role + "-" + story + ".docx")
    with ZipFile(source) as z:
        entries = {name: z.read(name) for name in z.namelist()}
    part = "word/" + ({"header": "header1", "footer": "footer1"}.get(story, "document")) + ".xml"
    xml = E.fromstring(entries[part])
    equation = next(xml.iter(m + "oMath"))
    if role == "root-degree":
        with ZipFile(cases.parent / "docx-omml-root-degree-display-cube.docx") as z:
            fragment = next(E.fromstring(z.read("word/document.xml")).iter(m + "rad"))
        equation.clear()
        equation.append(fragment)
    while True:
        pair = next(((parent, child) for parent in equation.iter() for child in parent
                     if child.tag in [w + "ins", w + "del"]), None)
        if pair is None:
            break
        parent, child = pair
        index = list(parent).index(child)
        parent.remove(child)
        for offset, node in enumerate(list(child)):
            parent.insert(index + offset, node)
    if empty:
        for parent in list(equation.iter()):
            for run in list(parent.findall(m + "r")):
                parent.remove(run)
    parent = next(parent for parent in xml.iter() if equation in list(parent))
    index = list(parent).index(equation)
    parent.remove(equation)
    wrapper = E.Element(w + kind, {
        w + "id": "973" if empty else "971" if kind == "ins" else "972",
        w + "author": "Whole Math Reviewer", w + "date": "2026-10-17T00:00:00Z"})
    wrapper.append(equation)
    parent.insert(index, wrapper)
    entries[part] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
    name = ("whole-math-" + ("empty-" if empty else "") + role + "-" + kind
            if story == "body" else "whole-math-placement-" + role + "-" + story + "-" + kind)
    target = cases / (name + ".docx")
    with ZipFile(target, "w", ZIP_DEFLATED) as z:
        for name, data in sorted(entries.items()):
            info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, data)
    row = dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
               Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
               SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
               Role=role, Kind=kind, Empty=empty, HeldOut=False, ExpectedEquationCount=4,
               ExpectedText=["share the same input series."])
    if story != "body":
        row["Story"] = story
    return row


def fragment_fixture(role, kind):
    source = cases / ("whole-math-" + ("root" if role == "fraction" else role) + "-" + kind + ".docx")
    with ZipFile(source) as z:
        entries = {name: z.read(name) for name in z.namelist()}
    xml = E.fromstring(entries["word/document.xml"])
    equation = next(xml.iter(m + "oMath"))
    if role == "fraction":
        with ZipFile(cases / "equation-revision-body-ins.docx") as z:
            fraction = next(E.fromstring(z.read("word/document.xml")).iter(m + "f"))
        equation.clear()
        equation.append(fraction)
    wrapper = next(parent for parent in xml.iter() if equation in list(parent))
    parent = next(parent for parent in xml.iter() if wrapper in list(parent))
    index = list(parent).index(wrapper)
    parent.remove(wrapper)
    wrapper.remove(equation)
    contents = list(equation)
    equation.clear()
    for child in contents:
        wrapper.append(child)
    equation.append(wrapper)
    parent.insert(index, equation)
    entries["word/document.xml"] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
    target = cases / ("math-fragment-" + role + "-" + kind + ".docx")
    with ZipFile(target, "w", ZIP_DEFLATED) as z:
        for name, data in sorted(entries.items()):
            info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, data)
    return dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
                Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
                SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
                Role=role, Kind=kind, HeldOut=False, ExpectedEquationCount=4,
                ExpectedText=["share the same input series."])


profiles = [(role, kind, False) for role in ["root", "delimiter", "operator", "root-degree"]
            for kind in ["ins", "del"]] + [(role, kind, True) for role in ["root", "delimiter", "operator"]
                                         for kind in ["ins", "del"]]
body = [fixture(*profile) for profile in profiles]
placement = [fixture(role, kind, story=story) for role, kind in [("root", "ins"), ("operator", "del")]
             for story in ["cell", "header", "footer"]]
fragments = [fragment_fixture(role, kind) for role in ["root", "delimiter", "operator", "fraction"]
             for kind in ["ins", "del"]]
for name, rows in [("whole-math", body), ("whole-math-placement", placement), ("math-fragment", fragments)]:
    (cases / (name + ".json")).write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
    print("Pinned", len(rows), name, "controls")

empty_fractions = []
for shape in ["whole", "fragment"]:
    for kind in ["ins", "del"]:
        source = cases / (("equation-revision-body-" if shape == "whole" else "math-fragment-fraction-") + kind + ".docx")
        with ZipFile(source) as z:
            entries = {name: z.read(name) for name in z.namelist()}
        xml = E.fromstring(entries["word/document.xml"])
        fraction = next(xml.iter(m + "f"))
        for parent in list(fraction.iter()):
            for run in list(parent.findall(m + "r")):
                parent.remove(run)
        entries["word/document.xml"] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
        target = cases / ("empty-fraction-revision-" + shape + "-" + kind + ".docx")
        with ZipFile(target, "w", ZIP_DEFLATED) as z:
            for name, data in sorted(entries.items()):
                info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                z.writestr(info, data)
        empty_fractions.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
                                    Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
                                    SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
                                    Shape=shape, Kind=kind, HeldOut=False, ExpectedEquationCount=4,
                                    ExpectedText=["share the same input series."], ProsePrefixes=["Summary:", "governs"],
                                    RuleRegion=[120, 680, 180, 740]))
(cases / "empty-fraction-revisions.json").write_text(json.dumps(empty_fractions, indent=2) + "\n", encoding="utf-8")
print("Pinned four glyphless fraction revisions with actual control ownership")
