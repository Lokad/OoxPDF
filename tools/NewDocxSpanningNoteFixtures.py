"""Four source-owned notes in one naturally paginated paragraph; no forced breaks."""
from pathlib import Path
from copy import deepcopy
from xml.etree import ElementTree as E
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
import hashlib
import json

root = Path(__file__).resolve().parent.parent
output = root / "tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes"
w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
E.register_namespace("w", w[1:-1])
rows = []
for typography in ["calibri-12-zero", "arial-11-zero"]:
    source = output / ("footnote-page-owners-" + typography + ".docx")
    with ZipFile(source) as z:
        entries = {name: z.read(name) for name in z.namelist()}
    xml = E.fromstring(entries["word/document.xml"])
    body = xml.find(w + "body")
    paragraphs = [p for p in body.findall(w + "p") if p.find(".//" + w + "footnoteReference") is not None]
    assert len(paragraphs) == 4
    section = deepcopy(body.find(w + "sectPr"))
    paragraph = E.Element(w + "p")
    properties = paragraphs[0].find(w + "pPr")
    if properties is not None:
        paragraph.append(deepcopy(properties))
    text_template = next(run for run in paragraphs[0].findall(w + "r") if run.find(w + "t") is not None)
    for original, repeats in zip(paragraphs, [25, 25, 75, 25]):
        run = deepcopy(text_template)
        texts = list(run.iter(w + "t"))
        texts[0].text = "Public item 105. " + "Public wrapped words retain authored digits 108. " * repeats
        for text in texts[1:]:
            text.text = ""
        paragraph.append(run)
        marker = next(run for run in original.findall(w + "r") if run.find(w + "footnoteReference") is not None)
        paragraph.append(deepcopy(marker))
        separator = deepcopy(text_template)
        for index, text in enumerate(separator.iter(w + "t")):
            text.text = " " if index == 0 else ""
        paragraph.append(separator)
    body.clear()
    body.append(paragraph)
    body.append(section)
    entries["word/document.xml"] = E.tostring(xml, encoding="utf-8", xml_declaration=True)
    target = output / ("spanning-page-notes-" + typography + ".docx")
    with ZipFile(target, "w", ZIP_DEFLATED) as z:
        for name, data in sorted(entries.items()):
            info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, data)
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        HeldOut=False, ExpectedEquationCount=0, ExpectedText=["Public destination Alpha.", "Public destination Delta."]))
(output / "spanning-page-notes.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Pinned two natural paragraph/page ownership controls")
