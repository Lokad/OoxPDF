"""Compare pinned Office math/prose anchors with an already inspected candidate.

Usage: python tools/SummarizeCompositionMathGeometry.py GOLDEN CANDIDATES REPORT
Retains every failed one-point gate; does not build, convert or rasterize PDFs.
Prose origins may follow a standalone measured space; arbitrary in-run anchors
and transformed whitespace require a richer independent inspection report.
"""
from pathlib import Path
import json
import sys
import unicodedata


def read(path):
    return json.loads(path.read_text(encoding="utf-8-sig"))


def main():
    golden, candidates, output = map(Path, sys.argv[1:])
    outcomes = read(candidates / "outcomes.json")
    identities = {(row["Id"], row["Mode"]): row for row in outcomes["Rows"]}
    rows = []
    for reference in read(golden)["Rows"]:
        key = reference["Id"], reference["Mode"]
        assert identities[key]["InputSha256"] == reference["InputSha256"], "Input identity differs"
        text_path = candidates / (key[0] + "-" + key[1]) / "inspect/text-operations.json"
        text = read(text_path) if text_path.is_file() else []
        math = [row for row in text if row["Font"].startswith("Eq")]
        assert len(math) == reference["MathGlyphCount"], "Math content count differs"
        actual = unicodedata.normalize("NFKC", "".join(row["DecodedText"] for row in math))
        assert actual == reference["MathText"], "Math content differs"
        dx = max((abs(math[anchor["Offset"]]["EffectiveX"] - anchor["X"]) for anchor in reference["Anchors"]), default=0)
        dy = max((abs(math[anchor["Offset"]]["EffectiveY"] - anchor["Y"]) for anchor in reference["Anchors"]), default=0)
        prose = []
        for anchor in reference.get("ProseAnchors", []):
            matches = [row for row in text if anchor["Prefix"] in row["DecodedText"]]
            assert len(matches) == 1, "Expected one inspected prose source"
            row = matches[0]
            leading = row["DecodedText"][:row["DecodedText"].index(anchor["Prefix"])]
            x = row["EffectiveX"]
            if leading:
                assert set(leading) == {" "} and row["EffectiveA"] == 1 and row["CharacterSpacing"] == 0, "Unsupported partial prose origin"
                spaces = [space for space in text if space["DecodedText"] == " " and
                          space["Font"] == row["Font"] and space["FontSize"] == row["FontSize"]]
                assert spaces, "Missing independently inspected space width"
                x += len(leading) * spaces[0]["NaturalWidthPoints"]
            prose.append(dict(Prefix=anchor["Prefix"], Dx=x - anchor["X"], Dy=row["EffectiveY"] - anchor["Y"]))
        passed = dx < 1 and dy < 1 and all(abs(row["Dx"]) < 1 and abs(row["Dy"]) < 1 for row in prose)
        rows.append(dict(Id=key[0], Mode=key[1], MathMaxX=dx, MathMaxY=dy, Prose=prose,
                         OnePointGatePassed=passed))
        if reference.get("RuleBoxes"):
            graphics = read(candidates / (key[0] + "-" + key[1]) / "inspect/graphics-operations.json")
            rules = [row for row in graphics if row["Kind"] == "Fill" and
                     1 < row["MaxX"] - row["MinX"] < 50 and 0 < row["MaxY"] - row["MinY"] < 2]
            bars = sum(row["Kind"] == "Fill" and 20 < row["MinX"] < 40 and
                       0 < row["MaxX"] - row["MinX"] < 2 and row["MaxY"] - row["MinY"] > 5 for row in graphics)
            boxes = []
            for expected in reference["RuleBoxes"]:
                matching = [row for row in rules if row["FillColor"] == expected["FillColor"]]
                if not matching:
                    continue
                rule = min(matching, key=lambda row: abs(row["MinX"] - expected["X"]) + abs(row["MinY"] - expected["Y"]))
                boxes.append(dict(Dx=rule["MinX"] - expected["X"], Dy=rule["MinY"] - expected["Y"],
                                  DWidth=rule["MaxX"] - rule["MinX"] - expected["Width"],
                                  DHeight=rule["MaxY"] - rule["MinY"] - expected["Height"]))
            rule_gate = len(rules) == reference["TotalThinRules"] and bars == reference["RevisionBarCount"] and \
                sum(row["FillColor"] == "g:0" for row in rules) == reference.get("BlackThinRules", reference["TotalThinRules"]) and \
                len(boxes) == len(reference["RuleBoxes"]) and \
                all(abs(row["Dx"]) < 1 and abs(row["Dy"]) < 1 and abs(row["DWidth"]) < .1 and abs(row["DHeight"]) < .03 for row in boxes)
            rows[-1].update(RuleBoxes=boxes, RevisionBarCount=bars, RuleGatePassed=rule_gate,
                            OnePointGatePassed=passed and rule_gate)
    output.write_text(json.dumps(dict(LibrarySha256=outcomes["LibrarySha256"],
                                     Scope="Independent one-point anchors; strict raster gates remain separate", Rows=rows),
                                 indent=2) + "\n", encoding="utf-8")
    print(sum(row["OnePointGatePassed"] for row in rows), "of", len(rows), "one-point anchor gates pass; all failures retained")


if __name__ == "__main__":
    main()
