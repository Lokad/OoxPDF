"""Screen the pinned public portfolio for missing prose and simple math content.

Uses existing PDF inspections, never converts documents or builds the library.
Token presence is not a test of reading order, glyph ink, clipping or placement.
"""
import argparse
from collections import Counter
import hashlib
import json
from pathlib import Path
import re
import posixpath
import subprocess
import unicodedata
from zipfile import ZipFile
import xml.etree.ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def read(path):
    return json.loads(path.read_text(encoding="utf-8-sig"))


def normalize(value):
    return unicodedata.normalize("NFKC", value).replace("\u2019", "'")


def words(value):
    return Counter(re.findall(r"[^\W_]+(?:'[^\W_]+)*", normalize(value).casefold()))


def deficits(expected, actual):
    return dict(sorted((expected - actual).items()))


def unresolved_fields(root):
    """Recognize stored text results, without evaluating fields or page values."""
    def cached_instruction(value):
        opcode = re.match(r"\s*([A-Za-z]+)", value)
        return bool(value.strip()) and (not opcode or opcode[1].upper() not in {"PAGE", "NUMPAGES"})

    def paragraph_items(paragraph):
        pending = list(reversed(list(paragraph)))
        while pending:
            item = pending.pop()
            if item.tag == W + "p":
                continue  # A nested paragraph owns an independent field stack.
            yield item
            pending.extend(reversed(list(item)))

    unresolved = set()
    for field in root.iter(W + "fldSimple"):
        if not cached_instruction(field.get(W + "instr", "")) or not any(item.text for item in field.iter(W + "t")):
            unresolved.add("fldSimple")
    visited = set()
    for paragraph in root.iter(W + "p"):
        stack = []
        for item in paragraph_items(paragraph):
            if item.tag not in {W + "fldChar", W + "instrText", W + "t"}:
                continue
            visited.add(item)
            if item.tag == W + "fldChar":
                kind = item.get(W + "fldCharType")
                if kind == "begin":
                    stack.append({"instruction": "", "separate": False, "text": False})
                elif kind == "separate" and stack and not stack[-1]["separate"]:
                    stack[-1]["separate"] = True
                elif kind == "end" and stack:
                    field = stack.pop()
                    if not field["separate"] or not field["text"] or not cached_instruction(field["instruction"]):
                        unresolved.add("fldChar")
                else:
                    unresolved.add("fldChar")
            elif item.tag == W + "instrText":
                if not stack or stack[-1]["separate"]:
                    unresolved.add("instrText")
                else:
                    stack[-1]["instruction"] += item.text or ""
            elif stack and item.text:
                for field in stack:
                    if field["separate"]:
                        field["text"] = True
                    else:
                        unresolved.add("fldChar")
        if stack:
            unresolved.add("fldChar")
    for item in root.iter():
        if item.tag in {W + "fldChar", W + "instrText"} and item not in visited:
            unresolved.add(item.tag.removeprefix(W))
    return sorted(unresolved)


def source_content(path, mode, pages):
    prose, math = Counter(), Counter()
    parts, unresolved = [], []
    with ZipFile(path) as package:
        static_parts, note_ids = set(), {}
        if path.suffix == ".docx":
            main = ET.fromstring(package.read("word/document.xml"))
            relationships = ET.fromstring(package.read("word/_rels/document.xml.rels"))
            targets = {item.get("Id"): posixpath.normpath(posixpath.join("word", item.get("Target", "")))
                       for item in relationships if item.get("TargetMode") != "External"}
            for kind in ("headerReference", "footerReference"):
                for item in main.iter(W + kind):
                    target = targets.get(item.get(R + "id"))
                    if target:
                        static_parts.add(target)
            if len(list(main.iter(W + "sectPr"))) != 1:
                unresolved.append({"Part": "word/document.xml", "Feature": "multiple-or-missing-sections"})
            for part, reference, kind in (("footnotes", "footnoteReference", "footnote"),
                                          ("endnotes", "endnoteReference", "endnote"),
                                          ("comments", "commentReference", "comment")):
                note_ids["word/" + part + ".xml"] = (kind, {item.get(W + "id") for item in main.iter(W + reference)})
        for name in package.namelist():
            if path.suffix == ".pptx":
                selected = re.fullmatch(r"ppt/slides/slide\d+\.xml", name)
                tag, weight = A + "t", 1
            else:
                static = name in static_parts
                selected_note = name in note_ids and bool(note_ids[name][1]) and (
                    name != "word/comments.xml" or mode == "all-word-compatible")
                selected = static or name == "word/document.xml" or selected_note
                tag, weight = W + "t", pages if static else 1
            if not selected:
                continue
            root = ET.fromstring(package.read(name))
            if name in note_ids:
                kind, ids = note_ids[name]
                selected = ET.Element("selected-notes")
                selected.extend(item for item in root.findall(W + kind) if item.get(W + "id") in ids)
                root = selected
            # This bounded audit does not reimplement OOXML display eligibility.
            for feature in ("del", "ins", "moveFrom", "moveTo", "vanish"):
                if root.find(".//" + W + feature) is not None:
                    unresolved.append({"Part": name, "Feature": feature})
            unresolved.extend({"Part": name, "Feature": feature} for feature in unresolved_fields(root))
            counts = words(" ".join(item.text or "" for item in root.iter(tag)))
            prose.update({token: count * weight for token, count in counts.items()})
            for equation in root.iter(M + "oMath"):
                # The portfolio has literal and simple fraction expressions.
                # More complex math needs a structural, source-backed check.
                kinds = {item.tag for item in equation.iter()}
                if any(kind not in {M + item for item in (
                    "oMath", "r", "t", "rPr", "ctrlPr", "f", "fPr", "type", "num", "den", "argPr", "nor", "sty")}
                       and not kind.startswith(W) for kind in kinds):
                    unresolved.append({"Part": name, "Feature": "complex-math"})
                    continue
                value = re.sub(r"\s+", "", normalize("".join(item.text or "" for item in equation.iter(M + "t"))))
                if len(value) >= 3:
                    math[value] += weight
                else:
                    unresolved.append({"Part": name, "Feature": "ambiguous-short-math"})
            parts.append({"Part": name, "Occurrences": weight, "ProseTokens": sum(counts.values()) * weight})
    return prose, math, parts, unresolved


def inspect_reference(pdf, inspector, cache):
    pdf_hash, inspector_hash = sha(pdf), sha(inspector)
    folder = cache / (pdf_hash + "-" + inspector_hash[:12])
    pin = folder / "identity.json"
    operations = folder / "text-operations.json"
    if not folder.exists():
        folder.mkdir(parents=True)
        with (folder / "inspect.log").open("w", encoding="utf-8") as log:
            subprocess.run(["dotnet", str(inspector), str(pdf), str(folder), "--text-only"],
                           stdout=log, stderr=subprocess.STDOUT, check=True)
        pin.write_text(json.dumps({"PdfSha256": pdf_hash, "InspectorSha256": inspector_hash,
                                  "OperationsSha256": sha(operations)}, indent=2) + "\n", encoding="utf-8")
    identity = read(pin)
    if identity != {"PdfSha256": pdf_hash, "InspectorSha256": inspector_hash, "OperationsSha256": sha(operations)}:
        raise ValueError("Reference text cache identity differs: " + str(folder))
    return read(operations), sha(operations)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    for name in ("artifact", "candidate", "references", "inspector", "cache", "output"):
        parser.add_argument("--" + name, required=True, type=Path)
    parser.add_argument("--manifest", default="tests/Lokad.OoxPdf.Tests/Cases/Composition/portfolio.json",
                        help="Public fixture manifest relative to the input root")
    parser.add_argument("--input-root", type=Path,
                        help="Optional clean tracked public inputs; reuse the frozen runtime without rebuilding")
    args = parser.parse_args()
    artifact, candidate, references = args.artifact.resolve(), args.candidate.resolve(), args.references.resolve()
    inspector = args.inspector.resolve()
    pin = read(artifact / "identity.json")
    library = artifact / "src/Lokad.OoxPdf/bin/Release/net10.0/Lokad.OoxPdf.dll"
    inputs = args.input_root.resolve() if args.input_root else artifact
    public_cases = (inputs / "tests/Lokad.OoxPdf.Tests/Cases").resolve()
    manifest_name = Path(args.manifest).as_posix()
    if Path(manifest_name).is_absolute() or ".." in Path(manifest_name).parts:
        raise ValueError("Manifest must be relative to the public input root")
    manifest = inputs / manifest_name
    if not manifest.resolve().is_relative_to(public_cases):
        raise ValueError("Only public fixture manifests are supported")
    input_revision = None
    def check_tracked_input(path):
        relative = path.relative_to(inputs).as_posix()
        git = ["git", "--literal-pathspecs", "-C", str(inputs)]
        subprocess.run(git + ["ls-files", "--error-unmatch", "--", relative],
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        if subprocess.check_output(git + ["status", "--porcelain", "--untracked-files=all", "--", relative]):
            raise ValueError("Supplemental public inputs must be committed: " + relative)
    if args.input_root:
        check_tracked_input(manifest)
        input_revision = subprocess.check_output(["git", "-C", str(inputs), "rev-parse", "HEAD"], text=True).strip()
    outcomes = read(candidate / "outcomes.json")
    comparisons = read(candidate / "office-comparisons.json")
    if sha(library) != pin["LibrarySha256"] or outcomes["LibrarySha256"] != pin["LibrarySha256"]:
        raise ValueError("Candidate/frozen library identity differs")
    if (not args.input_root and sha(manifest) != pin["Inputs"][manifest_name]) or outcomes["ManifestSha256"] != sha(manifest):
        raise ValueError("Pinned public manifest identity differs")
    inspector_hash = sha(inspector)
    if outcomes["InspectorSha256"] != inspector_hash:
        raise ValueError("Use the same qualified PDF inspector")
    cases = {case["Id"]: case for case in read(manifest)}
    compared = {(row["Id"], row["Mode"]): row for row in comparisons["Rows"]}
    expected_modes = {(case["Id"], mode) for case in cases.values() for mode in (
        ("final-preserve", "all-word-compatible") if case["Input"].endswith(".docx") else ("presentation",))}
    actual_modes = [(row["Id"], row["Mode"]) for row in outcomes["Rows"]]
    if set(actual_modes) != expected_modes or len(actual_modes) != len(expected_modes) or set(compared) != expected_modes:
        raise ValueError("Portfolio outcomes/comparisons are missing or duplicated")
    if comparisons["CandidateLibrarySha256"] != pin["LibrarySha256"]:
        raise ValueError("Comparison library identity differs")
    rows = []
    for outcome in outcomes["Rows"]:
        identifier, mode = outcome["Id"], outcome["Mode"]
        case, comparison = cases[identifier], compared[identifier, mode]
        source = inputs / case["Input"]
        if not source.resolve().is_relative_to(public_cases):
            raise ValueError("Only public fixture inputs are supported")
        if args.input_root:
            check_tracked_input(source)
        if sha(source) != case["Sha256"] or outcome["InputSha256"] != case["Sha256"] or (not args.input_root and pin["Inputs"][case["Input"]] != case["Sha256"]):
            raise ValueError("Public source identity differs")
        folder = identifier + "-" + mode
        run = candidate / folder
        pdf, reference = run / "candidate.pdf", references / folder / "reference.pdf"
        if sha(pdf) != outcome["PdfSha256"] or comparison["CandidatePdfSha256"] != sha(pdf) or comparison["ReferencePdfSha256"] != sha(reference):
            raise ValueError("Comparison PDF identities differ")
        operations_path = run / "inspect/text-operations.json"
        operations = read(operations_path)
        office, office_ops_hash = inspect_reference(reference, inspector, args.cache.resolve())
        candidate_words = words(" ".join(item.get("DecodedText") or "" for item in operations))
        office_words = words(" ".join(item.get("DecodedText") or "" for item in office))
        page_deficits = []
        for page in sorted({item["PageNumber"] for item in office + operations if item["PageNumber"] is not None}):
            expected_page = words(" ".join(item.get("DecodedText") or "" for item in office if item["PageNumber"] == page))
            observed_page = words(" ".join(item.get("DecodedText") or "" for item in operations if item["PageNumber"] == page))
            missing = deficits(expected_page, observed_page)
            if missing:
                page_deficits.append({"Page": page, "Tokens": missing})
        pages = max(item["PageNumber"] for item in office if item["PageNumber"] is not None)
        prose, expressions, parts, unresolved = source_content(source, mode, pages)
        math_stream = re.sub(r"\s+", "", normalize("".join(item.get("DecodedText") or "" for item in operations
                                                           if item["Font"].startswith("Eq"))))
        observed_math = Counter({expression: math_stream.count(expression) for expression in expressions})
        rows.append({"Id": identifier, "Mode": mode, "HeldOut": case["HeldOut"], "InputSha256": sha(source),
                     "CandidatePdfSha256": sha(pdf), "ReferencePdfSha256": sha(reference),
                     "CandidateOperationsSha256": sha(operations_path), "ReferenceOperationsSha256": office_ops_hash,
                     "ReferencePages": pages, "CandidatePages": outcome["InspectedPages"],
                     "SourceProseTokens": sum(prose.values()), "SourceParts": parts,
                     "SourceTokensMissingOffice": deficits(prose, office_words),
                     "SourceTokensMissingCandidate": deficits(prose, candidate_words),
                     "OfficeTokensMissingCandidate": deficits(office_words, candidate_words),
                     "OfficeTokensMissingCandidateByPage": page_deficits,
                     "ExpectedSimpleMath": dict(expressions), "ObservedSimpleMath": dict(observed_math),
                     "MissingSimpleMath": deficits(expressions, observed_math),
                     "IgnoredEquationDiagnostics": outcome["IgnoredEquationDiagnostics"],
                     "UnresolvedSourceFeatures": unresolved, "VisualGateFailures": comparison["Failures"]})
    if sha(library) != pin["LibrarySha256"] or sha(inspector) != inspector_hash:
        raise ValueError("Audit runtime changed")
    report = {"Scope": "Public portfolio prose-token and simple-math presence screen; not reading order, ink, images, clipping or geometry proof",
              "ToolSha256": sha(Path(__file__)),
              "FrozenRevision": pin["Revision"], "LibrarySha256": pin["LibrarySha256"], "InspectorSha256": inspector_hash,
              "ManifestSha256": sha(manifest), "Conversions": 0, "LibraryBuilds": 0,
              "InputRoot": str(inputs), "InputRevision": input_revision,
              "Rows": rows}
    if args.output.exists():
        raise ValueError("Preserve immutable reports; select a fresh output file")
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, indent=2, ensure_ascii=True) + "\n", encoding="utf-8")
    failed = sum(bool(row["SourceTokensMissingOffice"] or row["SourceTokensMissingCandidate"] or
                      row["OfficeTokensMissingCandidate"] or row["OfficeTokensMissingCandidateByPage"] or row["MissingSimpleMath"] or row["UnresolvedSourceFeatures"] or
                      row["IgnoredEquationDiagnostics"] or row["ReferencePages"] != row["CandidatePages"]) for row in rows)
    print(f"{len(rows)} mode outcomes; {failed} deficit/unresolved screens. No conversions or library builds.")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
