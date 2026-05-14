# Lokad.OoxPdf Current Execution Plan

This ExecPlan is the current working plan for `Lokad.OoxPdf`. It intentionally omits the historical bootstrap checklist that has already been completed. Keep `Progress`, `Private Evidence`, `Backlog`, `Decisions`, and `Validation` current as work proceeds.

## Goal

`Lokad.OoxPdf` is a dependency-free .NET library that converts `.pptx` and `.docx` OOXML documents to static PDF. The library must not call Office, PDFium, PowerShell, external executables, or third-party packages. Office and PDFium are allowed only in `tools/` for validation.

The project is now past the initial vertical slice. The next phase is fidelity: use public visual cases and private local-only documents to identify missing Office features, implement them as small scoped commits, and keep diagnostics honest when a feature is still missing.

## Repository Map

- `src/Lokad.OoxPdf`: production library, OOXML parsers, renderers, PDF writer, font/image code.
- `src/Lokad.OoxPdf.Cli`: local conversion CLI.
- `tests/Lokad.OoxPdf.Tests`: dependency-free console test runner and synthetic fixtures.
- `tools/CheckVisualCase.ps1`: public visual validation harness.
- `tools/CheckPrivateCase.ps1`: private local-only visual validation harness.
- `tools/RenderReference.ps1`: Office COM reference renderer.
- `tools/RasterizePdf.ps1`: PDFium rasterization wrapper.
- `tools/Lokad.OoxPdf.VisualDiff`: PNG comparison tool.
- `tools/Lokad.OoxPdf.PdfiumRasterizer`: local PDFium P/Invoke rasterizer.
- `visual-cases/`: public visual case manifests.
- `private-cases/`: ignored private manifests and inputs.
- `artifacts/`: ignored validation output.
- `docs/`: user-facing public docs.

## Progress

- [x] Dependency-free `.slnx` solution, library, CLI, tests, visual tools, docs, public fixtures, and private validation lane exist.
- [x] NuGet package version is set to `0.1.0` for the first package.
- [x] OOXML package layer handles ZIP parts, content types, relationships, safe part normalization, XML hardening, and package size limits.
- [x] PDF writer emits deterministic static PDFs with pages, drawing operators, embedded TrueType/CID fonts, ToUnicode maps, JPEG passthrough, PNG image XObjects, and alpha soft masks.
- [x] CLI supports `convert input output`, `--diagnostics`, `--strict`, and exit codes `0`, `1`, `2`, and `3`.
- [x] Visual validation can render Office references, rasterize candidate PDFs with PDFium, compute PNG metrics, and write comparison artifacts.
- [x] Private validation keeps inputs/manifests under ignored `private-cases/`, rejects tracked/private-unsafe paths, and writes ignored artifacts under `artifacts/private-visual/`.
- [x] PPTX parser/renderer supports slide order/size, solid backgrounds, basic rectangles/ellipses/lines, rotation/flip, common theme colors/fonts, common master/layout inheritance, text boxes, basic styled text, JPEG/PNG pictures, basic crop clipping, grouped shape transforms, fixed-grid tables, and unsupported-feature diagnostics.
- [x] DOCX parser/renderer supports page setup, margins, document defaults, paragraph styles, character styles, paragraphs/runs, basic styled text, greedy wrapping, simple page breaking, bullets/decimal numbering, inline JPEG/PNG images, fixed-width tables, default headers/footers, page number approximation, and unsupported-feature diagnostics.
- [x] PNG support covers non-interlaced RGB/RGBA, 8-bit grayscale, 8-bit indexed color, and packed low-bit-depth indexed color.
- [x] Unsupported PPTX image formats now emit `IMAGE_UNSUPPORTED_FORMAT` diagnostics instead of aborting the entire conversion.
- [x] PowerPoint reference export is sorted numerically so decks with more than 9 slides compare against the correct candidate pages.
- [x] Private PPTX assessment completed on a large 84-slide deck without exposing private contents.
- [x] Private DOCX assessment completed on an 18-candidate-page document without exposing private contents.

## Private Evidence

Private evidence is intentionally anonymized. Do not copy private text, screenshots, filenames, or document-specific business content into public notes.

- Private PPTX run `artifacts/private-visual/lokad-value-based/20260514-154018`:
  - 84 candidate pages, all dimensions matched reference pages.
  - Mean absolute error: `17.34937080899104`.
  - Max mean absolute error: `76.6187287808642`.
  - Mean changed-pixel ratio at threshold 16: `0.1864059055335096`.
  - Diagnostics: 9 unsupported charts, 2 unsupported interlaced PNG image occurrences.
  - Manual private inspection identified gaps in chart fallback, interlaced embedded images, dense image/group placement, text spacing, text-frame anchoring/clipping, and transparent overlays.
- Private DOCX run `artifacts/private-visual/user-requirements-spec/20260514-164847`:
  - Reference output had 16 pages; candidate output had 18 pages.
  - Candidate page height differed by 1 raster pixel from reference at 144 DPI, preventing pixel metrics.
  - Diagnostics were empty.
  - This identifies DOCX pagination/page geometry fidelity as a high-priority gap, especially because no diagnostic currently explains the mismatch.

## Backlog

### Release-Blocking Fidelity

- [ ] Implement Adam7 interlaced PNG decoding so embedded interlaced images render instead of being skipped.
- [ ] Make omitted embedded image content release-blocking: either render it, use a safe fallback, or emit an explicit high-severity diagnostic.
- [ ] Improve PPTX chart fallback rendering, starting with cached chart images where present, then static bar/line chart drawing for common chart XML.
- [ ] Fix DOCX page geometry and pagination fidelity: page height rounding, section page size/margins, line heights, paragraph spacing, and page-break decisions.
- [ ] Add diagnostics when DOCX reference-like pagination risks are detected: multi-section layout, unsupported page break variants, unsupported paragraph keep rules, or unsupported line-height semantics.

### PPTX Feature Survey

- [ ] Text layout: preserve spaces, tabs, line breaks, soft line breaks, kerning-like advances, font fallback, mixed run spacing, character spacing, superscript/subscript, and baseline offsets.
- [ ] Text frames: internal margins, vertical anchoring, clipping, overflow, autofit, shrink-to-fit, multi-column text, text rotation, and text inside arbitrary shapes.
- [ ] Fonts: select bold/italic faces instead of drawing approximations; support fallback fonts, embedded fonts, complex scripts, and bidirectional text.
- [ ] Shapes: more preset geometries, freeform paths, connectors, arrows, callouts, rounded rectangles, custom geometry, compound paths, and accurate line joins/caps/dashes.
- [ ] Fills/effects: transparency, gradients, pattern fills, picture fills, shadows, glows, reflections, blur, soft edges, and 3D effects.
- [ ] Images: grouped/placeholder-bound image placement, crop modes, rotation/flip interactions, recolor/duotone, transparency, SVG/EMF/WMF, TIFF/GIF/BMP, and image compression variants.
- [ ] Tables: merged cells, vertical alignment, per-edge borders, table styles, cell margins, rich text inside cells, and precise row/column sizing.
- [ ] Charts: cached chart images, chart XML rendering, axes, labels, legends, series styling, stacked/grouped bars, line charts, combo charts, and embedded chart data.
- [ ] SmartArt/diagrams: use fallback drawings when present; otherwise emit precise diagnostics.
- [ ] Slide inheritance: deeper master/layout placeholder resolution, theme variants, background styles, footer/date/slide-number placeholders, and hidden placeholder semantics.
- [ ] Media and dynamic features: videos, audio, animations, transitions, and OLE/ActiveX should remain static/diagnostic-only unless a reliable fallback is available.
- [ ] Comments/notes: speaker notes and comments should be ignored with diagnostics or exposed through an optional mode, not silently dropped.

### DOCX Feature Survey

- [ ] Pagination: Word-compatible line height, paragraph spacing collapse, keep-with-next, keep-lines-together, widow/orphan control, manual page/column breaks, section breaks, and page size rounding.
- [ ] Text layout: tabs, tab stops, indents, hanging indents, justification, hyphenation, nonbreaking spaces, soft hyphens, field result handling, superscript/subscript, and baseline shifts.
- [ ] Fonts: bold/italic face selection, fallback fonts, embedded fonts, complex scripts, bidirectional text, OpenType features, and symbol fonts.
- [ ] Numbering: exact level text expansion, hanging indents, bullets from symbol fonts, restart rules, multi-level lists, custom number formats, and style-linked numbering.
- [ ] Tables: auto-fit, preferred widths, nested tables, merged cells, vertical merges, cell margins, borders, shading, row height rules, repeating header rows, and page breaks inside tables.
- [ ] Images/drawings: anchored/floating drawings, wrap modes, relative positioning, cropping, rotation, text boxes, shapes, SmartArt, charts, and drawing canvases.
- [ ] Headers/footers: first/odd/even variants, section-specific variants, distance from edge, fields beyond `PAGE`, total page count, and dynamic date/doc properties.
- [ ] Fields: `PAGE`, `NUMPAGES`, `DATE`, `REF`, `HYPERLINK`, `TOC`, `SEQ`, form fields, and cached-field fallback semantics.
- [ ] Footnotes/endnotes/comments: render bodies or emit precise diagnostics with usable fallback behavior.
- [ ] Tracked changes: choose final, original, or marked-up view explicitly and document the behavior.
- [ ] Multi-column layout, text boxes, sidebars, bookmarks, hyperlinks, outlines, and document properties.

### PDF/Infrastructure

- [ ] Add PDF hyperlinks, outlines/bookmarks, metadata, and optional tagged-PDF structure if needed by consumers.
- [ ] Add font subsetting to reduce output size while keeping deterministic output.
- [ ] Add image deduplication and compression choices for large decks.
- [ ] Improve diagnostics severity model so release-blocking omissions are distinguishable from harmless approximations.
- [ ] Add visual comparison support for dimension-near-matches, so a 1-pixel raster rounding mismatch can still produce pixel metrics.
- [ ] Add private-case summary tooling that reports page count, dimension mismatches, diagnostics grouped by feature, and worst visual pages without exposing content.

## Next Implementation Targets

1. Adam7 interlaced PNG support in `src/Lokad.OoxPdf/Imaging/PngImage.cs`, with synthetic tests and a private PPTX rerun.
2. DOCX page geometry/pagination investigation using the private DOCX evidence, starting with page height rounding and section/page-size conversion.
3. PPTX chart fallback detection: locate cached chart images or alternate content before attempting chart XML drawing.
4. PPTX text spacing and text-frame layout fixes, validated against public cases and the private PPTX metrics.
5. Dense PPTX image/group placement fidelity, especially for image-heavy slides.

## Decisions

- The library remains dependency-free. Third-party packages are not allowed in `src/Lokad.OoxPdf`.
- Office and PDFium remain validation-only under `tools/`.
- Private documents remain under ignored `private-cases/`; generated private artifacts remain under ignored `artifacts/private-visual/`.
- Public notes from private documents must be anonymized to feature gaps and metrics only.
- Diagnostics must prefer continued conversion over crashing, but omitted visible content must not be treated as acceptable final behavior.
- Pixel metrics are evidence, not the final truth. Human/agent visual inspection remains necessary for representative cases.

## Validation

Latest public validation:

```powershell
dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal
dotnet run --project tests/Lokad.OoxPdf.Tests --tl:off
dotnet pack src/Lokad.OoxPdf/Lokad.OoxPdf.csproj --tl:off --nologo -v minimal --no-restore
```

Current expected test result:

```text
46 passed, 0 failed
```

Representative public visual cases already exist for PPTX blank/shapes/text/images/tables/corporate-theme and DOCX blank/basic paragraphs/numbering/images/tables/headers-footers.

Private validation commands:

```powershell
pwsh tools/CheckPrivateCase.ps1 -Case private-cases/lokad-value-based.json
pwsh tools/CheckPrivateCase.ps1 -Case private-cases/user-requirements-spec.json
```

Do not commit private inputs, private manifests, private rendered pages, private diagnostics, private comparison HTML, or private assessments.

## Idempotence And Recovery

All build/test/pack commands are safe to rerun. Visual validation writes timestamped directories. If Office COM automation leaves an Office process running after a failure, close it only after confirming no unrelated user document is open.

If a private case reveals a missing feature, record only public-safe feature gaps here, then create synthetic public tests for the implementation. Do not derive public fixtures from private documents.
