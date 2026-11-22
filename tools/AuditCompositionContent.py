"""Screen the pinned public portfolio for missing prose, simple math and optional picture ink.

Uses existing PDF inspections, never converts documents or builds the library.
Token presence is not a test of reading order, glyph ink, clipping or placement.
"""
import argparse
from collections import Counter
import hashlib
import json
import math
from pathlib import Path
import re
import posixpath
import subprocess
from tempfile import TemporaryDirectory
import unicodedata
from zipfile import ZipFile
import xml.etree.ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def sha(path):
    return sha_bytes(path.read_bytes())


def sha_bytes(data):
    return hashlib.sha256(data).hexdigest().upper()


def read(path):
    return json.loads(path.read_text(encoding="utf-8-sig"))


def resolve_reference(directories, identifier, mode, expected_hash):
    """Every present explicit cache entry must match the compared PDF identity."""
    folder = identifier + "-" + mode
    matches = []
    for directory in directories:
        source = directory / folder
        if not source.exists():
            continue
        pdf = source / "reference.pdf"
        if not pdf.is_file():
            raise ValueError("Reference PDF missing: " + str(pdf))
        if sha(pdf) != expected_hash:
            raise ValueError("Reference PDF identity differs: " + str(pdf))
        matches.append(pdf)
    if not matches:
        raise ValueError("Reference PDF missing: " + folder)
    return matches[0]


def normalize(value):
    return unicodedata.normalize("NFKC", value).replace("\u2019", "'")


def words(value):
    return Counter(re.findall(r"[^\W_]+(?:'[^\W_]+)*", normalize(value).casefold()))


def operation_text(operations):
    """Keep paint order; join only text whose measured origins are contiguous.

    A text-show boundary is not a word boundary: equations commonly paint each
    glyph separately. Missing metrics, empty replacement text and differing
    resources/state remain boundaries. The 0.003pt tolerance covers rounding of
    three-decimal PDF coordinates, not a font-size-dependent word-space guess.
    """
    state = ("PageNumber", "ObjectNumber", "Generation", "Font", "FillColor",
             "FontSize", "CharacterSpacing", "EffectiveA", "EffectiveB", "EffectiveC", "EffectiveD")
    metrics = ("FontSize", "CharacterSpacing", "EffectiveA", "EffectiveB", "EffectiveC", "EffectiveD",
               "EffectiveX", "EffectiveY", "EmittedAdvancePoints")

    def contiguous(left, right):
        if left is None:
            return False
        before, after = left.get("DecodedText") or "", right.get("DecodedText") or ""
        if not before or not after or before[-1].isspace() or after[0].isspace():
            return False
        if any(left.get(key) is None or left[key] != right.get(key) for key in state):
            return False
        for item in (left, right):
            if any(not isinstance(item.get(key), (int, float)) or isinstance(item[key], bool) or
                   not math.isfinite(item[key]) for key in metrics):
                return False
            if item["FontSize"] <= 0 or item["EmittedAdvancePoints"] <= 0:
                return False
            if item["EffectiveA"] * item["EffectiveD"] == item["EffectiveB"] * item["EffectiveC"]:
                return False
        expected_x = left["EffectiveX"] + left["EmittedAdvancePoints"] * left["EffectiveA"]
        expected_y = left["EffectiveY"] + left["EmittedAdvancePoints"] * left["EffectiveB"]
        return math.hypot(right["EffectiveX"] - expected_x, right["EffectiveY"] - expected_y) <= 0.003

    chunks, previous = [], None
    for item in operations:
        if chunks and not contiguous(previous, item):
            chunks.append(" ")
        chunks.append(item.get("DecodedText") or "")
        previous = item
    return "".join(chunks)


def deficits(expected, actual):
    return dict(sorted((expected - actual).items()))


def screen_picture_ink(source, candidate, reference, case, fitz):
    """Screen the public portfolio's opaque two-color PNGs and named SVG fill."""
    source_count = case.get("ExpectedInlineImageCount")
    if source_count is None:
        return {"Status": "No picture expectation"}
    svg_color = tuple(case.get("SvgWitnessColor", ()))
    expected = {"Png": source_count - bool(svg_color), "Svg": int(bool(svg_color))}
    with ZipFile(source) as package:
        media = [Path(name) for name in package.namelist() if name.startswith("word/media/") and not name.endswith("/")]
        svg_stems = {part.stem for part in media if part.suffix == ".svg"}
        # The paired raster preview is a fallback for the same authored SVG.
        rasters = {package.read(part.as_posix()) for part in media
                   if part.suffix == ".png" and part.stem not in svg_stems}
    if (expected["Png"] < 0 or any(part.suffix not in {".png", ".svg"} for part in media)
            or len(svg_stems) != expected["Svg"] or len(rasters) != int(expected["Png"] > 0)):
        return {"Status": "Unverified source pictures", "ExpectedSourcePictures": source_count}

    palette = set()
    if rasters:
        raster = fitz.Pixmap(next(iter(rasters)))
        if raster.colorspace is None or raster.colorspace.n != 3:
            return {"Status": "Unverified PNG color space", "ExpectedSourcePictures": source_count}
        samples = raster.samples
        palette = set(zip(samples[0::raster.n], samples[1::raster.n], samples[2::raster.n]))
        if len(palette) != 2 or (raster.alpha and any(value != 255 for value in samples[3::4])):
            return {"Status": "Unverified PNG palette", "ExpectedSourcePictures": source_count}

    scale = fitz.Matrix(3, 3)

    def render_region(page, rect):
        return page.get_pixmap(matrix=scale, clip=rect, colorspace=fitz.csRGB, annots=False)

    def has_colors(bitmap, colors):
        samples = iter(bitmap.samples)
        return colors.issubset(zip(samples, samples, samples))

    def count_regions(page, bounds, colors):
        count = 0
        for box in {tuple(box) for box in bounds}:
            rect = fitz.Rect(box) & page.rect
            if rect.is_empty or not rect.is_valid:
                continue
            count += has_colors(render_region(page, rect), colors)
        return count

    def count_pngs(page):
        # The paint log supplies the order of the image and any foreground paint.
        # Bounds are conservative. An obscured witness may remain unverified.
        rotation = page.rotation_matrix
        paint = [(kind, fitz.Rect(box) * rotation) for kind, box in page.get_bboxlog()
                 if kind != "ignore-text"]
        images = {tuple(box): index for index, (kind, box) in enumerate(paint) if kind == "fill-image"}
        neutral = next((value, value, value) for value in range(3) if (value, value, value) not in palette)
        visible = uncertain = 0
        for box, index in images.items():
            rect = fitz.Rect(box) & page.rect
            if rect.is_empty or not rect.is_valid:
                continue
            bitmap = render_region(page, rect)
            if not has_colors(bitmap, palette):
                continue
            covered = False
            for _, later in paint[index + 1:]:
                overlap = rect & later
                if not overlap.is_empty:
                    covered |= bitmap.set_rect((overlap * scale).irect, neutral)
            if covered and not has_colors(bitmap, palette):
                uncertain += 1
            else:
                visible += 1
        return visible, uncertain

    def scan(path):
        counts, uncertain_pngs = {}, {}
        with fitz.open(path) as pdf:
            for page in pdf:
                pngs, uncertain = count_pngs(page) if palette else (0, 0)
                paths = (drawing["rect"] for drawing in page.get_drawings() if drawing["fill"] and
                         tuple(round(channel * 255) for channel in drawing["fill"]) == svg_color) if svg_color else ()
                svgs = count_regions(page, paths, {svg_color}) if svg_color else 0
                if pngs or svgs:
                    counts[page.number + 1] = {"Png": pngs, "Svg": svgs}
                if uncertain:
                    uncertain_pngs[page.number + 1] = uncertain
        return counts, uncertain_pngs

    office, uncertain_office = scan(reference)
    actual, uncertain_candidate = scan(candidate)
    office_totals = {kind: sum(page[kind] for page in office.values()) for kind in expected}
    actual_totals = {kind: sum(page[kind] for page in actual.values()) for kind in expected}
    if uncertain_office or any(office_totals[kind] < expected[kind] for kind in expected):
        status = "Unverified Office picture witnesses"
    elif uncertain_candidate:
        status = "Unverified candidate picture witnesses"
    elif any(actual_totals[kind] < office_totals[kind] for kind in expected):
        status = "Missing picture ink"
    elif actual != office:
        status = "Picture page counts differ"
    else:
        status = "Picture witnesses present"
    return {"Status": status, "ExpectedSourcePictures": source_count, "OfficeByPage": office,
            "CandidateByPage": actual, "UnverifiedOfficePngsByPage": uncertain_office,
            "UnverifiedCandidatePngsByPage": uncertain_candidate, "PngColors": sorted(palette),
            "SvgColor": svg_color}


def unresolved_fields(root):
    """Recognize stored text results, without evaluating fields or page values."""
    def cached_instruction(value):
        opcode = re.match(r"\s*([^\s\\]+)", value)
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
    # Legacy entries lack a pin for the log that establishes the page count.
    folder = cache / ("v2-" + pdf_hash + "-" + inspector_hash[:12])
    if not folder.exists():
        cache.mkdir(parents=True, exist_ok=True)
        with TemporaryDirectory(prefix="reference-inspection-", dir=cache) as temporary:
            staged = Path(temporary) / "inspection"
            staged.mkdir()
            with (staged / "inspect.log").open("w", encoding="utf-8") as log:
                subprocess.run(["dotnet", str(inspector), str(pdf), str(staged), "--text-only"],
                               stdout=log, stderr=subprocess.STDOUT, check=True)
            if (sha(pdf), sha(inspector)) != (pdf_hash, inspector_hash):
                raise ValueError("Reference PDF or inspector changed during inspection")
            result = read_reference_inspection(staged, pdf_hash, inspector_hash, creating=True)
            try:
                staged.rename(folder)
            except OSError:
                # A concurrent inspection can publish the same entry first.
                if not folder.is_dir():
                    raise
            else:
                return result
    return read_reference_inspection(folder, pdf_hash, inspector_hash)


def read_reference_inspection(folder, pdf_hash, inspector_hash, *, creating=False):
    pin = folder / "identity.json"
    operations = (folder / "text-operations.json").read_bytes()
    log = (folder / "inspect.log").read_bytes()
    identity = {"PdfSha256": pdf_hash, "InspectorSha256": inspector_hash,
                "OperationsSha256": sha_bytes(operations), "InspectionLogSha256": sha_bytes(log)}
    if not creating and read(pin) != identity:
        raise ValueError("Reference text cache identity differs: " + str(folder))
    # Page objects include empty and trailing blank pages. Text-operation page
    # numbers cannot establish the size of an empty reference document.
    log_text = log.decode("utf-8-sig").replace("\r\n", "\n")
    pages = len(re.findall(r"^\d+ \d+ obj: page$", log_text, re.MULTILINE))
    if pages < 1:
        raise ValueError("No inspected reference pages: " + str(folder))
    records = json.loads(operations.decode("utf-8-sig"))
    if not isinstance(records, list):
        raise ValueError("Invalid reference text operations: " + str(folder))
    if creating:
        pin.write_text(json.dumps(identity, indent=2) + "\n", encoding="utf-8")
    return records, identity["OperationsSha256"], pages, identity["InspectionLogSha256"]


def validate_outcome_modes(cases, outcomes, comparisons):
    modes = {"Both": ("final-preserve", "all-word-compatible"),
             "Final": ("final-preserve",), "Review": ("all-word-compatible",)}
    preset = outcomes.get("DocxPreset", "Both")
    if preset not in modes:
        raise ValueError("Unknown measured DOCX preset")
    expected = {(case["Id"], mode) for case in cases.values() for mode in (
        modes[preset] if case["Input"].endswith(".docx") else ("presentation",))}
    actual = [(row["Id"], row["Mode"]) for row in outcomes["Rows"]]
    compared = [(row["Id"], row["Mode"]) for row in comparisons["Rows"]]
    if (set(actual) != expected or len(actual) != len(expected) or
            set(compared) != expected or len(compared) != len(expected)):
        raise ValueError("Portfolio outcomes/comparisons are missing or duplicated")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    for name in ("artifact", "candidate", "inspector", "cache", "output"):
        parser.add_argument("--" + name, required=True, type=Path)
    parser.add_argument("--references", required=True, nargs="+", type=Path,
                        help="Explicit Office cache roots; every present entry must match the compared PDF")
    parser.add_argument("--manifest", default="tests/Lokad.OoxPdf.Tests/Cases/Composition/portfolio.json",
                        help="Public fixture manifest relative to the input root")
    parser.add_argument("--input-root", type=Path,
                        help="Optional clean tracked public inputs; reuse the frozen runtime without rebuilding")
    parser.add_argument("--screen-pictures", action="store_true",
                        help="Use PyMuPDF to screen the curated public PNG/SVG picture witnesses")
    args = parser.parse_args()
    if args.screen_pictures:
        import pymupdf
    artifact, candidate = args.artifact.resolve(), args.candidate.resolve()
    references = [directory.resolve(strict=True) for directory in args.references]
    if any(not directory.is_dir() for directory in references):
        raise ValueError("Reference roots must be directories")
    inspector = args.inspector.resolve()
    pin = read(artifact / "identity.json")
    framework = pin.get("Framework", "net10.0")
    if framework not in {"net10.0", "net11.0"}:
        raise ValueError("Unknown frozen artifact framework")
    library = artifact / "src/Lokad.OoxPdf/bin/Release" / framework / "Lokad.OoxPdf.dll"
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
    validate_outcome_modes(cases, outcomes, comparisons)
    if comparisons["CandidateLibrarySha256"] != pin["LibrarySha256"]:
        raise ValueError("Comparison library identity differs")
    reference_pdfs = {(row["Id"], row["Mode"]): resolve_reference(references, row["Id"], row["Mode"],
        compared[row["Id"], row["Mode"]]["ReferencePdfSha256"]) for row in outcomes["Rows"]}
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
        pdf, reference = run / "candidate.pdf", reference_pdfs[identifier, mode]
        pdf_hash, reference_hash = sha(pdf), sha(reference)
        if pdf_hash != outcome["PdfSha256"] or comparison["CandidatePdfSha256"] != pdf_hash or comparison["ReferencePdfSha256"] != reference_hash:
            raise ValueError("Comparison PDF identities differ")
        operations_path = run / "inspect/text-operations.json"
        operations_pin = outcome.get("TextOperationsSha256")
        if not operations_pin:
            raise ValueError("Text inspection identity is missing; reinspect the candidate PDF: " + folder)
        if not operations_path.is_file():
            raise ValueError("Candidate text inspection is missing: " + folder)
        operations_bytes = operations_path.read_bytes()
        operations_hash = sha_bytes(operations_bytes)
        if operations_hash != operations_pin:
            raise ValueError("Candidate text inspection differs: " + folder)
        operations = json.loads(operations_bytes.decode("utf-8-sig"))
        if not isinstance(operations, list):
            raise ValueError("Invalid candidate text operations: " + folder)
        office, office_ops_hash, pages, office_log_hash = inspect_reference(reference, inspector, args.cache.resolve())
        candidate_words = words(operation_text(operations))
        office_words = words(operation_text(office))
        page_deficits = []
        for page in sorted({item["PageNumber"] for item in office + operations if item["PageNumber"] is not None}):
            expected_page = words(operation_text(item for item in office if item["PageNumber"] == page))
            observed_page = words(operation_text(item for item in operations if item["PageNumber"] == page))
            missing = deficits(expected_page, observed_page)
            if missing:
                page_deficits.append({"Page": page, "Tokens": missing})
        prose, expressions, parts, unresolved = source_content(source, mode, pages)
        math_stream = re.sub(r"\s+", "", normalize("".join(item.get("DecodedText") or "" for item in operations
                                                           if item["Font"].startswith("Eq"))))
        observed_math = Counter({expression: math_stream.count(expression) for expression in expressions})
        pictures = screen_picture_ink(source, pdf, reference, case, pymupdf) if args.screen_pictures else {
            "Status": "Unverified" if "ExpectedInlineImageCount" in case else "No picture expectation"}
        if args.screen_pictures and (sha(pdf) != pdf_hash or sha(reference) != reference_hash):
            raise ValueError("PDF identities changed during picture screening: " + folder)
        rows.append({"Id": identifier, "Mode": mode, "HeldOut": case["HeldOut"], "InputSha256": sha(source),
                     "CandidatePdfSha256": pdf_hash, "ReferencePdfSha256": reference_hash,
                     "ReferenceSourceDirectory": str(reference.parent),
                     "CandidateOperationsSha256": operations_hash, "ReferenceOperationsSha256": office_ops_hash,
                     "ReferenceInspectionLogSha256": office_log_hash,
                     "ReferencePages": pages, "CandidatePages": outcome["InspectedPages"],
                     "SourceProseTokens": sum(prose.values()), "SourceParts": parts,
                     "SourceTokensMissingOffice": deficits(prose, office_words),
                     "SourceTokensMissingCandidate": deficits(prose, candidate_words),
                     "OfficeTokensMissingCandidate": deficits(office_words, candidate_words),
                     "OfficeTokensMissingCandidateByPage": page_deficits,
                     "ExpectedSimpleMath": dict(expressions), "ObservedSimpleMath": dict(observed_math),
                     "MissingSimpleMath": deficits(expressions, observed_math), "PictureInk": pictures,
                     "IgnoredEquationDiagnostics": outcome["IgnoredEquationDiagnostics"],
                     "UnresolvedSourceFeatures": unresolved, "VisualGateFailures": comparison["Failures"]})
    if sha(library) != pin["LibrarySha256"] or sha(inspector) != inspector_hash:
        raise ValueError("Audit runtime changed")
    report = {"Scope": "Public portfolio prose-token and simple-math screen with optional curated picture ink. Reading order, complete clipping and exact geometry require separate evidence.",
              "ToolSha256": sha(Path(__file__)),
              "OutcomesSha256": sha(candidate / "outcomes.json"), "ComparisonsSha256": sha(candidate / "office-comparisons.json"),
              "FrozenRevision": pin["Revision"], "Framework": framework,
              "LibrarySha256": pin["LibrarySha256"], "InspectorSha256": inspector_hash,
              "ManifestSha256": sha(manifest), "Conversions": 0, "LibraryBuilds": 0,
              "ReferenceDirectories": [str(directory) for directory in references],
              "PictureScreenPyMuPDF": pymupdf.VersionBind if args.screen_pictures else None,
              "InputRoot": str(inputs), "InputRevision": input_revision,
              "Rows": rows}
    if args.output.exists():
        raise ValueError("Preserve immutable reports; select a fresh output file")
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, indent=2, ensure_ascii=True) + "\n", encoding="utf-8")
    failed = sum(bool(row["SourceTokensMissingOffice"] or row["SourceTokensMissingCandidate"] or
                      row["OfficeTokensMissingCandidate"] or row["OfficeTokensMissingCandidateByPage"] or row["MissingSimpleMath"] or row["UnresolvedSourceFeatures"] or
                      row["IgnoredEquationDiagnostics"] or row["ReferencePages"] != row["CandidatePages"] or
                      (args.screen_pictures and row["PictureInk"]["Status"] not in
                       {"Picture witnesses present", "No picture expectation"})) for row in rows)
    print(f"{len(rows)} mode outcomes; {failed} deficit/unresolved screens. No conversions or library builds.")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
