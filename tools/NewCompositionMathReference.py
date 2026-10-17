"""Pin math-run anchors from existing independently exported and inspected PDFs.

Usage: python tools/NewCompositionMathReference.py MANIFEST REFERENCES OUTPUT
Uses Cambria Math runs; mixed normal-font formulas need their own source audit.
Does not export Office, inspect PDFs or build the library.
"""
from pathlib import Path
import hashlib
import json
import sys
import unicodedata


def read(path):
    return json.loads(path.read_text(encoding="utf-8-sig"))


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def main():
    manifest, references, output = map(Path, sys.argv[1:])
    root = Path(__file__).resolve().parent.parent
    rows = []
    for source in read(manifest):
        assert sha(root / source["Input"]) == source["Sha256"], "Source identity differs"
        for mode in ["final-preserve", "all-word-compatible"]:
            run = references / (source["Id"] + "-" + mode)
            pin = read(run / "source-identity.json")
            assert pin["InputSha256"] == source["Sha256"] and pin["Mode"] == mode
            assert sha(run / "reference.pdf") == pin["PdfSha256"], "Reference identity differs"
            assert sha(run / "export.log") == pin["ExportLogSha256"], "Office view log differs"
            review = mode == "all-word-compatible"
            expected = dict(ShowRevisions=review, RevisionsView=0, MarkupMode=2 if review else 0,
                            RevisionsMode=2 if review else 0, ShowComments=review,
                            ShowFormatChanges=review, ShowInsertionsAndDeletions=review, RejectAll=False)
            assert pin["OfficeView"] == expected, "Office export view differs"
            fonts = {(row["PageNumber"], row["ResourceName"])
                     for row in read(run / "inspect/font-resources.json")
                     if "CambriaMath" in row["BaseFont"]}
            math = [row for row in read(run / "inspect/text-operations.json")
                    if (row["PageNumber"], row["Font"]) in fonts]
            assert math and {row["PageNumber"] for row in math} == {1}, "Expected single-page math control"
            anchors = []
            offset = 0
            for row in math:
                text = unicodedata.normalize("NFKC", row["DecodedText"])
                anchors.append(dict(Offset=offset, Text=text, X=row["EffectiveX"], Y=row["EffectiveY"],
                                    FontSize=row["FontSize"], FillColor=row["FillColor"]))
                offset += len(text)
            rows.append(dict(Id=source["Id"], Mode=mode, InputSha256=source["Sha256"],
                             ReferencePdfSha256=pin["PdfSha256"], OfficeView=expected,
                             MathText="".join(row["Text"] for row in anchors), MathGlyphCount=offset,
                             ImageCount=source.get("ImageCount", 1), ExpectedText=source["ExpectedText"],
                             Anchors=anchors))
    output.write_text(json.dumps(dict(Rows=rows), indent=2) + "\n", encoding="utf-8")
    print("Pinned", len(rows), "independent math outcomes without exporting or rebuilding")


if __name__ == "__main__":
    main()
