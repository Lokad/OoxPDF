"""Isolate style-default owner presence without changing visible run families."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
import hashlib
import json

root = Path(__file__).resolve().parent.parent
output = root / "tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes"
w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
rel = "http://schemas.openxmlformats.org/package/2006/relationships"
ct = "http://schemas.openxmlformats.org/package/2006/content-types"
owners = ["no-styles", "empty-styles", "empty-doc-defaults", "empty-run-defaults", "run-default-container",
          "font-run-defaults", "sized-run-defaults", "paragraph-defaults", "normal-default-style"]
rows = []
for family in ["Calibri", "Arial"]:
    for owner in owners if family == "Calibri" else ["no-styles", "font-run-defaults", "sized-run-defaults"]:
        fonts = f'<w:rFonts w:ascii="{family}" w:hAnsi="{family}" w:eastAsia="{family}" w:cs="{family}"/>'
        def run(text, properties=""):
            return f'<w:r><w:rPr>{fonts}{properties}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'
        def paragraph(text, properties=""):
            return "<w:p>" + run(text, properties) + "</w:p>"
        body = "<w:p>" + run("Public body default") + f'<w:r><w:rPr>{fonts}<w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2" w:customMarkFollows="1"/><w:t>*</w:t></w:r></w:p>'
        body += paragraph("Public body explicit", '<w:sz w:val="30"/><w:szCs w:val="30"/>')
        body += '<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>' + paragraph("Public cell default") + '</w:tc></w:tr></w:tbl>'
        body += '<w:p/><w:sectPr><w:headerReference w:type="default" r:id="header"/><w:footerReference w:type="default" r:id="footer"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>'
        entries = {
            "_rels/.rels": f'<Relationships xmlns="{rel}"><Relationship Id="document" Type="{r}/officeDocument" Target="word/document.xml"/></Relationships>',
            "word/document.xml": f'<w:document xmlns:w="{w}" xmlns:r="{r}"><w:body>{body}</w:body></w:document>',
            "word/header1.xml": f'<w:hdr xmlns:w="{w}">{paragraph("Public header default")}</w:hdr>',
            "word/footer1.xml": f'<w:ftr xmlns:w="{w}">{paragraph("Public footer default")}</w:ftr>',
            "word/footnotes.xml": f'<w:footnotes xmlns:w="{w}"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p>{run("*", "<w:vertAlign w:val=\"superscript\"/>")}{run(" Public footnote default")}</w:p></w:footnote></w:footnotes>',
        }
        parts = [("document", "document.xml", "document.main"), ("header", "header1.xml", "header"),
                 ("footer", "footer1.xml", "footer"), ("footnotes", "footnotes.xml", "footnotes")]
        if owner != "no-styles":
            defaults = {"empty-styles": "", "empty-doc-defaults": "<w:docDefaults/>",
                "empty-run-defaults": "<w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault></w:docDefaults>",
                "run-default-container": "<w:docDefaults><w:rPrDefault/></w:docDefaults>",
                "font-run-defaults": f'<w:docDefaults><w:rPrDefault><w:rPr>{fonts}</w:rPr></w:rPrDefault></w:docDefaults>',
                "sized-run-defaults": f'<w:docDefaults><w:rPrDefault><w:rPr>{fonts}<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>',
                "paragraph-defaults": "<w:docDefaults><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>",
                "normal-default-style": f'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr>{fonts}</w:rPr></w:style>',
            }[owner]
            entries["word/styles.xml"] = f'<w:styles xmlns:w="{w}">{defaults}</w:styles>'
            parts.append(("styles", "styles.xml", "styles"))
        entries["word/_rels/document.xml.rels"] = f'<Relationships xmlns="{rel}">' + ''.join(
            f'<Relationship Id="{id}" Type="{r}/{id}" Target="{name}"/>' for id, name, _ in parts if id != "document") + '</Relationships>'
        entries["[Content_Types].xml"] = f'<Types xmlns="{ct}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' + ''.join(
            f'<Override PartName="/word/{name}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{suffix}+xml"/>' for _, name, suffix in parts) + '</Types>'
        target = output / ("run-default-owner-" + family.lower() + "-" + owner + ".docx")
        with ZipFile(target, "w", ZIP_DEFLATED) as z:
            for name, text in sorted(entries.items()):
                info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                z.writestr(info, '<?xml version="1.0" encoding="UTF-8"?>' + text)
        rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
            Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(), Family=family,
            Owner=owner, HeldOut=False, ExpectedEquationCount=0, ExpectedImagePlacements=0,
            ExpectedText=["Public body default", "Public body explicit", "Public cell default",
                          "Public header default", "Public footer default", "Public footnote default"]))
(output / "run-default-owners.json").write_text(json.dumps(rows, indent=2) + "\n", encoding="utf-8")
print("Pinned twelve default-owner controls across body, cells, header, footer and footnote")
