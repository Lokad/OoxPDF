"""Validation-only TrueType ink screen on pinned PDFs; no builds, conversions or Office.

MuPDF supplies paint operations and the exact font program for each text span.
Display names are deliberately never used to select a font. Rectangle clips are
qualified; other clip shapes, unsupported fonts and stroked text remain unknown.
This is not an image, overpaint, reading-order or Office-parity gate.
"""

import argparse
import hashlib
import io
import json
from pathlib import Path
import xml.etree.ElementTree as ET

import fontTools
from fontTools.ttLib import TTFont
import pymupdf as fitz


def sha256(data):
    return hashlib.sha256(data).hexdigest().upper()


class FontPaints(fitz.mupdf.FzDevice2):
    """Keep native paint order and bind spans to their actual font buffers."""

    def __init__(self):
        super().__init__()
        self.events = []
        self.programs = {}
        self.use_virtual_fill_text()
        self.use_virtual_stroke_text()

    def collect(self, kind, text):
        spans = []
        raw = text.head
        while raw:
            span = fitz.mupdf.FzTextSpan(raw)
            font = span.font()
            digest = None
            if font.m_internal.buffer:
                # Same reference-counted buffer access as PyMuPDF Font.buffer.
                buffer = fitz.mupdf.FzBuffer(
                    fitz.mupdf.ll_fz_keep_buffer(font.m_internal.buffer))
                data = fitz.mupdf.fz_buffer_extract_copy(buffer)
                digest = sha256(data)
                self.programs.setdefault(digest, data)
            spans.append((digest, [span.items(i).gid for i in range(span.m_internal.len)]))
            raw = raw.next
        self.events.append((kind, spans))

    def fill_text(self, ctx, text, ctm, colorspace, color, alpha, color_params):
        self.collect("fill_text", text)

    def stroke_text(self, ctx, text, stroke, ctm, colorspace, color, alpha, color_params):
        self.collect("stroke_text", text)


def transform_box(bounds, matrix):
    points = [fitz.Point(x, y) * matrix
              for x in (bounds[0], bounds[2]) for y in (bounds[1], bounds[3])]
    return fitz.Rect(min(p.x for p in points), min(p.y for p in points),
                     max(p.x for p in points), max(p.y for p in points))


def rectangle_clip(node):
    if node.tag != "clip_path":
        return None
    commands = list(node)
    if any(c.tag not in ("moveto", "lineto", "closepath") for c in commands):
        return None
    edges = commands[:-1] if commands and commands[-1].tag == "closepath" else commands
    if not edges or edges[0].tag != "moveto" or any(c.tag != "lineto" for c in edges[1:]):
        return None
    points = [(float(c.attrib["x"]), float(c.attrib["y"]))
              for c in commands if c.tag in ("moveto", "lineto")]
    if len(points) == 5 and points[0] == points[-1]:
        points.pop()
    if len(points) != 4 or len(set(points)) != 4:
        return None
    matrix = fitz.Matrix([float(x) for x in node.attrib["transform"].split()])
    points = [fitz.Point(x, y) * matrix for x, y in points]
    # Diagonal and quarter-turn transforms keep rectangle edges axis-aligned.
    for first, last in zip(points, points[1:] + points[:1]):
        horizontal = abs(first.y - last.y) < .001
        vertical = abs(first.x - last.x) < .001
        if horizontal == vertical:
            return None
    return fitz.Rect(min(p.x for p in points), min(p.y for p in points),
                     max(p.x for p in points), max(p.y for p in points))


def analyze_pdf(path, expected_sha256=None):
    path = Path(path)
    digest = sha256(path.read_bytes())
    if expected_sha256 and digest != expected_sha256.upper():
        raise ValueError("PDF hash differs: " + str(path))
    result = {"PdfSha256": digest, "GlyphsMeasured": 0, "ClippedTextTraces": [],
              "UnmeasuredTextTraces": [], "InklessGlyphs": 0, "Pages": []}
    document = fitz.open(path)
    try:
        for page in document:
            device = FontPaints()
            fitz.mupdf.fz_run_page(page.this, device, fitz.mupdf.FzMatrix(1, 0, 0, 1, 0, 0),
                                  fitz.mupdf.FzCookie())
            fitz.mupdf.fz_close_device(device)
            embedded = {}
            for entry in page.get_fonts(full=True):
                data = document.extract_font(entry[0])[3]
                if data:
                    embedded.setdefault(sha256(data), []).append(entry[4])
            fonts = {}
            for identity, data in device.programs.items():
                if identity in embedded:
                    font = TTFont(io.BytesIO(data))
                    if "glyf" in font:
                        fonts[identity] = font
            buffer = fitz.mupdf.fz_new_buffer(1024)
            output = fitz.mupdf.FzOutput(buffer)
            trace = fitz.mupdf.fz_new_trace_device(output)
            fitz.mupdf.fz_run_page(page.this, trace, fitz.mupdf.FzMatrix(1, 0, 0, 1, 0, 0),
                                  fitz.mupdf.FzCookie())
            fitz.mupdf.fz_close_device(trace)
            output.fz_close_output()
            tree = ET.fromstring(b"<trace>" + fitz.mupdf.fz_buffer_extract_copy(buffer) + b"</trace>")
            stack = [(page.rect, False)]
            event_index = 0
            names = {}

            def visit(parent):
                nonlocal event_index
                for node in parent:
                    if node.tag in ("group", "layer", "trace", "structure", "metatext"):
                        visit(node)
                        continue
                    if node.tag.startswith("clip_"):
                        rectangle = rectangle_clip(node)
                        clip, unknown = stack[-1]
                        stack.append((clip & rectangle if rectangle is not None else clip,
                                      unknown or rectangle is None))
                        continue
                    if node.tag == "pop_clip":
                        if len(stack) <= 1:
                            raise ValueError("Unbalanced native clip trace")
                        stack.pop()
                        continue
                    if node.tag not in ("fill_text", "stroke_text"):
                        continue
                    if event_index >= len(device.events):
                        raise ValueError("Font paint trace is incomplete")
                    kind, bindings = device.events[event_index]
                    event_index += 1
                    if kind != node.tag or len(node) != len(bindings):
                        raise ValueError("Font paint order differs from geometry trace")
                    clip, unknown_clip = stack[-1]
                    matrix = fitz.Matrix([float(x) for x in node.attrib["transform"].split()])
                    for span, (identity, glyph_ids) in zip(node, bindings):
                        name = span.attrib["font"]
                        prefix, plus, suffix = name.partition("+")
                        display_name = suffix if plus and len(prefix) == 6 and all("A" <= c <= "Z" for c in prefix) else name
                        names.setdefault(display_name, set()).add(identity)
                        font = fonts.get(identity)
                        if len(span) != len(glyph_ids):
                            raise ValueError("Font span lengths differ from geometry trace")
                        if font is not None:
                            order = font.getGlyphOrder()
                            for node_glyph, glyph_id in zip(span, glyph_ids):
                                if glyph_id < 0 or glyph_id >= len(order):
                                    raise ValueError("GID outside its actual embedded font program")
                                if node_glyph.attrib["glyph"] not in (str(glyph_id), order[glyph_id]):
                                    raise ValueError("Font span glyphs differ from geometry trace")
                        reason = ("UnknownClip" if unknown_clip else
                                  "StrokedText" if kind == "stroke_text" else
                                  "TransparentPaint" if float(node.attrib.get("alpha", "1")) <= 0 else
                                  "UnavailableEmbeddedTrueTypeProgram" if font is None else None)
                        if reason:
                            if any(g.attrib.get("unicode", "").strip() for g in span):
                                result["UnmeasuredTextTraces"].append(
                                    {"Page": page.number + 1, "Paint": event_index,
                                     "FontProgramSha256": identity, "Reason": reason})
                            continue
                        order = font.getGlyphOrder()
                        units = font["head"].unitsPerEm
                        a, b, c, d = [float(x) for x in span.attrib["trm"].split()]
                        cut = []
                        for index, glyph_node in enumerate(span):
                            character = glyph_node.attrib.get("unicode", "")
                            if not character or character.isspace():
                                continue
                            glyph_id = glyph_ids[index]
                            if glyph_id < 0 or glyph_id >= len(order):
                                raise ValueError("GID outside its actual embedded font program")
                            glyph = font["glyf"][order[glyph_id]]
                            if not hasattr(glyph, "xMin"):
                                result["InklessGlyphs"] += 1
                                continue
                            result["GlyphsMeasured"] += 1
                            glyph_matrix = fitz.Matrix(a / units, b / units, c / units, d / units,
                                                       float(glyph_node.attrib["x"]),
                                                       float(glyph_node.attrib["y"])) * matrix
                            ink = transform_box((glyph.xMin, glyph.yMin, glyph.xMax, glyph.yMax),
                                                glyph_matrix)
                            if (ink.x0 < clip.x0 - .01 or ink.y0 < clip.y0 - .01 or
                                    ink.x1 > clip.x1 + .01 or ink.y1 > clip.y1 + .01):
                                cut.append({"Index": index, "Character": character, "GlyphId": glyph_id,
                                            "Ink": list(ink), "FullyHidden": (ink & clip).is_empty})
                        if cut:
                            result["ClippedTextTraces"].append(
                                {"Page": page.number + 1, "Paint": event_index,
                                 "FontProgramSha256": identity, "Clip": list(clip), "CutGlyphs": cut})

            visit(tree)
            if len(stack) != 1 or event_index != len(device.events):
                raise ValueError("Native paint/clip traces are incomplete")
            result["Pages"].append(
                {"Page": page.number + 1, "PaintsBound": event_index,
                 "FontPrograms": sorted(fonts),
                 "FontResourceBindings": {identity: embedded[identity] for identity in sorted(fonts)},
                 "DisplayNameCollisions": {name: sorted(x for x in values if x is not None)
                                           for name, values in names.items() if len(values) > 1}})
            for font in fonts.values():
                font.close()
    finally:
        document.close()
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--pdf", type=Path)
    source.add_argument("--outcomes", type=Path)
    parser.add_argument("--expected-pdf-sha256")
    parser.add_argument("--expected-library-sha256")
    parser.add_argument("--manifest", type=Path, help="Select the IDs of an existing document manifest")
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    if args.output.exists():
        parser.error("Preserve existing evidence; use a fresh output file")
    if args.pdf:
        if args.manifest or args.expected_library_sha256:
            parser.error("Manifest/library identities require --outcomes")
        rows = [analyze_pdf(args.pdf, args.expected_pdf_sha256)]
        inputs = {"Pdf": str(args.pdf)}
    else:
        if args.expected_pdf_sha256:
            parser.error("Use the per-outcome PDF identities with --outcomes")
        raw = args.outcomes.read_bytes()
        outcomes = json.loads(raw.decode("utf-8-sig"))
        if (args.expected_library_sha256 and
                outcomes["LibrarySha256"] != args.expected_library_sha256.upper()):
            parser.error("Frozen library identity differs")
        cases = None
        inputs = {"Outcomes": str(args.outcomes), "OutcomesSha256": sha256(raw),
                  "LibrarySha256": outcomes["LibrarySha256"]}
        if args.manifest:
            raw_manifest = args.manifest.read_bytes()
            cases = {x["Id"] for x in json.loads(raw_manifest.decode("utf-8-sig"))}
            inputs.update(Manifest=str(args.manifest), ManifestSha256=sha256(raw_manifest))
        rows = []
        seen = set()
        for outcome in outcomes["Rows"]:
            if cases is not None and outcome["Id"] not in cases:
                continue
            path = args.outcomes.parent / (outcome["Id"] + "-" + outcome["Mode"]) / "candidate.pdf"
            row = analyze_pdf(path, outcome["PdfSha256"])
            row.update(Id=outcome["Id"], Mode=outcome["Mode"])
            rows.append(row)
            seen.add(outcome["Id"])
        if not rows or cases is not None and seen != cases:
            parser.error("The selected outcome set is incomplete")
    report = {"Scope": __doc__, "Inputs": inputs, "Rows": rows,
              "ToolSha256": sha256(Path(__file__).read_bytes()),
              "PyMuPdfVersion": fitz.VersionBind, "FontToolsVersion": fontTools.__version__,
              "NewConversions": 0, "NewBuilds": 0, "NewOfficeActions": 0}
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with args.output.open("x", encoding="utf-8") as output:
        json.dump(report, output, indent=2)
        output.write("\n")
    print(f"{len(rows)} PDFs screened; {sum(x['GlyphsMeasured'] for x in rows)} glyphs measured; "
          f"{sum(len(x['ClippedTextTraces']) for x in rows)} clipped traces; "
          f"{sum(len(x['UnmeasuredTextTraces']) for x in rows)} unknown traces.")


if __name__ == "__main__":
    main()
