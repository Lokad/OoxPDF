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


profiles = [(role, kind, False) for role in ["root", "delimiter", "operator", "root-degree"]
            for kind in ["ins", "del"]] + [(role, kind, True) for role in ["root", "delimiter", "operator"]
                                         for kind in ["ins", "del"]]
body = [fixture(*profile) for profile in profiles]
placement = [fixture(role, kind, story=story) for role, kind in [("root", "ins"), ("operator", "del")]
             for story in ["cell", "header", "footer"]]
for name, rows in [("whole-math", body), ("whole-math-placement", placement)]:
    (cases / (name + ".json")).write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
    print("Pinned", len(rows), name, "controls")
