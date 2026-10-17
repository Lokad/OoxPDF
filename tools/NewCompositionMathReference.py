"""Pin math-run anchors from existing independently exported and inspected PDFs.

Usage: python tools/NewCompositionMathReference.py MANIFEST REFERENCES OUTPUT
Uses Cambria Math runs; mixed normal-font formulas need their own source audit.
Does not export Office, inspect PDFs or build the library.
ExpectedByMode can overlay expected metadata; ExpectedMathGlyphCount=0 pins a
view where tracked content is intentionally absent.
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
            outcome = source | source.get("ExpectedByMode", {}).get(mode, {})
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
            text_operations = read(run / "inspect/text-operations.json")
            math = [row for row in text_operations
                    if (row["PageNumber"], row["Font"]) in fonts]
            if outcome.get("ExpectedMathGlyphCount") == 0:
                assert not math, "Expected a view without mathematical content"
            else:
                assert math and {row["PageNumber"] for row in math} == {1}, "Expected single-page math control"
            anchors = []
            offset = 0
            for row in math:
                text = unicodedata.normalize("NFKC", row["DecodedText"])
                anchors.append(dict(Offset=offset, Text=text, X=row["EffectiveX"], Y=row["EffectiveY"],
                                    FontSize=row["FontSize"], FillColor=row["FillColor"]))
                offset += len(text)
            if "ExpectedMathGlyphCount" in outcome:
                assert offset == outcome["ExpectedMathGlyphCount"], "Independent math count differs"
            rows.append(dict(Id=source["Id"], Mode=mode, InputSha256=source["Sha256"],
                             ReferencePdfSha256=pin["PdfSha256"], OfficeView=expected,
                             MathText="".join(row["Text"] for row in anchors), MathGlyphCount=offset,
                             ImageCount=outcome.get("ImageCount", 1), ExpectedText=outcome["ExpectedText"],
                             Anchors=anchors))
            if outcome.get("ProsePrefixes"):
                prose = []
                for prefix in outcome["ProsePrefixes"]:
                    matches = [row for row in text_operations if row["DecodedText"].lstrip().startswith(prefix)]
                    assert len(matches) == 1, "Prose anchor must identify one independent run"
                    row = matches[0]
                    prose.append(dict(Prefix=prefix, X=row["EffectiveX"], Y=row["EffectiveY"]))
                rows[-1]["ProseAnchors"] = prose
            if outcome.get("RuleRegion"):
                left, bottom, right, top = outcome["RuleRegion"]
                graphics = read(run / "inspect/graphics-operations.json")
                rules = [row for row in graphics if row["Kind"] == "Fill" and
                         left < row["MinX"] < right and bottom < row["MinY"] < top and
                         1 < row["MaxX"] - row["MinX"] < 50 and 0 < row["MaxY"] - row["MinY"] < 2 and
                         (not outcome.get("RuleFillColor") or row["FillColor"] == outcome["RuleFillColor"])]
                assert len(rules) == 1, "Expected one source-region rule"
                rows[-1]["RuleBoxes"] = [dict(X=row["MinX"], Y=row["MinY"], Width=row["MaxX"] - row["MinX"],
                                              Height=row["MaxY"] - row["MinY"], FillColor=row["FillColor"]) for row in rules]
                rows[-1]["TotalThinRules"] = sum(row["Kind"] == "Fill" and 1 < row["MaxX"] - row["MinX"] < 50 and
                                                0 < row["MaxY"] - row["MinY"] < 2 for row in graphics)
                if outcome.get("RuleFillColor"):
                    rows[-1]["BlackThinRules"] = sum(row["Kind"] == "Fill" and row["FillColor"] == "g:0" and
                                                    1 < row["MaxX"] - row["MinX"] < 50 and 0 < row["MaxY"] - row["MinY"] < 2
                                                    for row in graphics)
                rows[-1]["RevisionBarCount"] = sum(row["Kind"] == "Fill" and 20 < row["MinX"] < 40 and
                                                  0 < row["MaxX"] - row["MinX"] < 2 and row["MaxY"] - row["MinY"] > 5
                                                  for row in graphics)
    output.write_text(json.dumps(dict(Rows=rows), indent=2) + "\n", encoding="utf-8")
    print("Pinned", len(rows), "independent math outcomes without exporting or rebuilding")


if __name__ == "__main__":
    main()
