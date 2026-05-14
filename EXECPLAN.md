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
- [x] NuGet package output is configured under ignored `artifacts/nuget/`.
- [x] OOXML package layer handles ZIP parts, content types, relationships, safe part normalization, XML hardening, and package size limits.
- [x] PDF writer emits deterministic static PDFs with pages, drawing operators, embedded TrueType/CID fonts, ToUnicode maps, JPEG passthrough, PNG image XObjects, and alpha soft masks.
- [x] CLI supports `convert input output`, `--diagnostics`, `--strict`, and exit codes `0`, `1`, `2`, and `3`.
- [x] Visual validation can render Office references, rasterize candidate PDFs with PDFium, compute PNG metrics, and write comparison artifacts.
- [x] Private validation keeps inputs/manifests under ignored `private-cases/`, rejects tracked/private-unsafe paths, and writes ignored artifacts under `artifacts/private-visual/`.
- [x] PPTX parser/renderer supports slide order/size, solid backgrounds, basic rectangles/ellipses/lines, rotation/flip, common theme colors/fonts, theme discovery through presentation or slide master relationships, common scheme color aliases, common master/layout inheritance, text boxes with body insets, line breaks, basic tab advances, direct bullet characters, paragraph spacing, vertical anchoring, clipping, basic styled text, JPEG/PNG pictures, basic crop clipping, grouped shape and picture transforms, fixed-grid tables, static bar-chart fallback, and unsupported-feature diagnostics.
- [x] DOCX parser/renderer supports page setup, margins, document defaults, paragraph styles, character styles, paragraphs/runs, basic styled text, greedy wrapping, simple page breaking, page-break-before, exact/at-least line heights, bullets/decimal numbering, inline JPEG/PNG images, fixed-width tables in body order with explicit row heights and row-level page breaks, default headers/footers, page number approximation, and unsupported-feature diagnostics.
- [x] DOCX diagnostics flag pagination-risk features that are still approximated or ignored: manual page/column breaks, direct and style-level keep/widow rules, style spacing variants, numbering indents, table styles/header rows, and paragraph section breaks.
- [x] PNG support covers non-interlaced RGB/RGBA, 8-bit grayscale, 8-bit indexed color, and packed low-bit-depth indexed color.
- [x] PNG support covers Adam7 interlaced RGBA images.
- [x] Unsupported PPTX and DOCX image formats now emit release-blocking `IMAGE_UNSUPPORTED_FORMAT` diagnostics while continuing the conversion.
- [x] PowerPoint reference export is sorted numerically so decks with more than 9 slides compare against the correct candidate pages.
- [x] Private PPTX assessment completed on a large 84-slide deck without exposing private contents.
- [x] Private-safe PPTX slide inventory tooling reports per-slide feature counts without exposing slide text or images.
- [x] Private DOCX assessment completed on an 18-candidate-page document without exposing private contents.
- [x] VisualDiff computes overlap metrics even when reference/candidate raster dimensions differ by a small page-rounding mismatch.

## Private Evidence

Private evidence is intentionally anonymized. Do not copy private text, screenshots, filenames, or document-specific business content into public notes.

- Private PPTX run `artifacts/private-visual/lokad-value-based/20260514-154018`:
  - 84 candidate pages, all dimensions matched reference pages.
  - Mean absolute error: `17.34937080899104`.
  - Max mean absolute error: `76.6187287808642`.
  - Mean changed-pixel ratio at threshold 16: `0.1864059055335096`.
  - Diagnostics: 9 unsupported charts, 2 unsupported interlaced PNG image occurrences.
  - Manual private inspection identified gaps in chart fallback, interlaced embedded images, dense image/group placement, text spacing, text-frame anchoring/clipping, and transparent overlays.
- Private PPTX conversion-only run `artifacts/private-visual/lokad-value-based/manual-20260514-165850`:
  - Conversion completed.
  - Interlaced PNG image diagnostics dropped to zero after Adam7 decoding support.
  - Remaining diagnostics: 9 unsupported charts.
- Private PPTX rerun `artifacts/private-visual/lokad-value-based/20260514-171646`:
  - 84 candidate pages, all dimensions matched reference pages.
  - Unsupported chart diagnostics dropped to zero after static bar-chart fallback.
  - Diagnostics: 9 chart static fallback informational diagnostics.
  - Mean absolute error: `17.430568`; max mean absolute error: `76.618729`; mean changed-pixel ratio at threshold 16: `0.187419`.
- Private PPTX rerun `artifacts/private-visual/lokad-value-based/20260514-172019`:
  - 84 candidate pages, all dimensions matched reference pages.
  - Diagnostics: 9 chart static fallback informational diagnostics.
  - Text-frame body insets are now honored.
  - Mean absolute error: `17.525889`; max mean absolute error: `76.657147`; mean changed-pixel ratio at threshold 16: `0.188083`.
- Private PPTX rerun `artifacts/private-visual/lokad-value-based/20260514-175356`:
  - 84 candidate pages, all dimensions matched reference pages.
  - Diagnostics: 9 chart static fallback informational diagnostics.
  - Grouped picture transforms and text clipping are now honored.
  - Mean absolute error: `16.223412`; max mean absolute error: `75.220066`; mean changed-pixel ratio at threshold 16: `0.176527`.
  - Private visual inspection found the generated `output.pdf` is not acceptable on any slide. Aggregate pixel metrics are useful for regression tracking but are not evidence of visual correctness for this deck.
  - Working hypothesis: broad PPTX failure is caused by interacting slide-level issues, especially z-order, master/layout/placeholder inheritance, theme resolution, text/image/table/chart placement, and missing effects rather than one isolated unsupported primitive.
- Private PPTX inventory `artifacts/private-visual/lokad-value-based/inventory/20260514-182157.json`:
  - 84 slides inventoried without exposing text or images.
  - Feature counts: 72 slides with pictures, 61 with grouped content, 11 with tables, 9 with charts, 23 with effects, 10 with transparency, and 84 with inherited layout/master content.
  - Slide-level totals: 1846 shapes, 440 pictures, 12 table nodes, and 9 chart nodes.
  - Initial simplest-slide strategy should start with slide 1, which is structurally small but still exercises slide content plus inherited layout/master content.
- Private PPTX slide-1 rerun `artifacts/private-visual/lokad-value-based/20260514-182802`:
  - Theme colors now resolve for slide text after loading themes through slide-master relationships and handling scheme aliases.
- Private PPTX slide-1 rerun `artifacts/private-visual/lokad-value-based/20260514-185419`:
  - Direct bullet characters now render on slide text.
  - TTC font discovery and large-text baseline handling make the title/body text structurally close enough to expose the next slide-1 gap.
- Private PPTX slide-1 rerun `artifacts/private-visual/lokad-value-based/20260514-185955`:
  - Bullet hanging indents from `marL`/`indent` now separate bullet glyphs from bullet text.
- Private PPTX slide-1 rerun `artifacts/private-visual/lokad-value-based/20260514-190411`:
  - PPTX text now uses per-run font resources, including bullet font selection from `buFont`.
- Private PPTX slide-1 rerun `artifacts/private-visual/lokad-value-based/20260514-191256`:
  - Level-1 list-style defaults now supply missing run size/style/color/typeface and line spacing.
  - Remaining slide-1 generic gaps are fine text metrics, exact run advance/underline placement, and spacing precision.
- Private PPTX slide-2 rerun `artifacts/private-visual/lokad-value-based/20260514-192433`:
  - Slide background pictures now render before slide shapes, exposing later outline shapes.
  - Shape `fontRef` colors now supply fallback text color, fixing placeholder text color.
- Private PPTX slide-2 rerun `artifacts/private-visual/lokad-value-based/20260514-193029`:
  - Run-level text highlights now render as background rectangles.
  - Remaining slide-2 generic gaps include finer mixed-run text advance, title positioning, and exact highlight bounds.
- Private DOCX run `artifacts/private-visual/user-requirements-spec/20260514-164847`:
  - Reference output had 16 pages; candidate output had 18 pages.
  - Candidate page height differed by 1 raster pixel from reference at 144 DPI, preventing pixel metrics.
  - Diagnostics were empty.
  - This identifies DOCX pagination/page geometry fidelity as a high-priority gap, especially because no diagnostic currently explains the mismatch.
- Private DOCX rerun `artifacts/private-visual/user-requirements-spec/20260514-171244`:
  - Reference output had 16 pages; candidate output had 16 pages after preserving DOCX table body order and reducing default table row height.
  - Candidate page height still differs by 1 raster pixel at 144 DPI.
  - Overlap metrics are now available despite the dimension mismatch; paired-page mean absolute error was `19.89926`, and mean changed-pixel ratio at threshold 16 was `0.16322`.
  - Diagnostics were empty, so remaining pagination gaps still need explicit diagnostics or rendering fixes.
- Private DOCX rerun `artifacts/private-visual/user-requirements-spec/20260514-173117`:
  - Reference output had 16 pages; candidate output had 16 pages.
  - All 16 rasterized page dimensions matched after A4 media-box normalization.
  - Mean absolute error: `19.88376`; mean changed-pixel ratio at threshold 16: `0.163965`.
  - Diagnostics were empty.
- Private DOCX rerun `artifacts/private-visual/user-requirements-spec/20260514-175910`:
  - Reference output had 16 pages; candidate output had 17 pages.
  - Candidate has one extra page; paired-page mean absolute error was `19.226760`, and mean changed-pixel ratio at threshold 16 was `0.158228`.
  - Diagnostics were empty.
  - This keeps DOCX pagination fidelity as the top active risk; attempted manual-break and keep-with-next heuristics were reverted because they did not resolve the private page-count mismatch.
  - Anonymized structure survey found no direct manual page/column breaks, no direct paragraph keep rules, no direct paragraph spacing, no inline/anchored drawings, and one section.
  - The same survey found 198 body paragraphs, 13 body tables, 129 table rows, 422 table cells, 45 numbered paragraphs, 24 style-level spacing definitions, 30 style-level keep-rule definitions, 36 numbering levels with indents, 13 table preferred widths, 422 cell widths, 18 cell vertical-alignment declarations, and 1 repeating table-header row.
  - Working hypothesis: the 17th candidate page is driven by accumulated small layout errors, especially style-derived paragraph spacing/keep rules, numbering indents/hanging indents, and table sizing/header behavior rather than an explicit page-break feature.
- Private DOCX rerun `artifacts/private-visual/user-requirements-spec/20260514-180723`:
  - Reference output had 16 pages; candidate output had 17 pages.
  - Candidate still has one extra page; paired-page mean absolute error was `19.226760`, and mean changed-pixel ratio at threshold 16 was `0.158228`.
  - Diagnostics now identify the public-safe pagination risk categories: `DOCX_NUMBERING_INDENT`, `DOCX_STYLE_PARAGRAPH_KEEP_RULE`, `DOCX_STYLE_PARAGRAPH_SPACING`, `DOCX_STYLE_TABLE_STYLE`, `DOCX_UNSUPPORTED_TABLE_HEADER_ROW`, and `DOCX_UNSUPPORTED_TABLE_STYLE`.
  - Private visual inspection found DOCX tables in `output.pdf` are visibly wrong enough to require their own recovery track, not just incremental pagination tuning.
  - Next implementation should start with layout tracing or one of those diagnosed categories; avoid broad paragraph parser rewrites until drift location is known.

## Backlog

### Release-Blocking Fidelity

- [x] Implement Adam7 interlaced PNG decoding so embedded interlaced images render instead of being skipped.
- [x] Make omitted embedded image content release-blocking: render supported JPEG/PNG, otherwise emit explicit high-severity diagnostics.
- [x] Improve PPTX chart fallback rendering for cached numeric bar-chart XML with an approximate static grouped-bar fallback.
- [ ] Extend PPTX chart rendering beyond basic bar fallbacks: cached image fallbacks when present, labels, legends, axes, line charts, pie charts, stacked/grouped variants, and style fidelity.
- [ ] Fix DOCX page geometry and pagination fidelity: section page size/margins, paragraph spacing, manual page/column breaks, and keep/widow page-break decisions.
- [x] Add diagnostics when DOCX reference-like pagination risks are detected: multi-section layout, unsupported page break variants, or unsupported paragraph keep rules.

### PPTX Feature Survey

- [ ] Text layout: preserve spaces, tabs, line breaks, soft line breaks, kerning-like advances, font fallback, mixed run spacing, character spacing, superscript/subscript, and baseline offsets.
- [ ] Text frames: overflow behavior beyond hard clipping, autofit, shrink-to-fit, multi-column text, text rotation, and text inside arbitrary shapes.
- [ ] Fonts: select bold/italic faces instead of drawing approximations; support fallback fonts, embedded fonts, complex scripts, and bidirectional text.
- [ ] Shapes: more preset geometries, freeform paths, connectors, arrows, callouts, rounded rectangles, custom geometry, compound paths, and accurate line joins/caps/dashes.
- [ ] Fills/effects: transparency, gradients, pattern fills, picture fills, shadows, glows, reflections, blur, soft edges, and 3D effects.
- [ ] Images: placeholder-bound image placement, crop modes, rotation/flip interactions, recolor/duotone, transparency, SVG/EMF/WMF, TIFF/GIF/BMP, and image compression variants.
- [ ] Tables: merged cells, vertical alignment, per-edge borders, table styles, cell margins, rich text inside cells, and precise row/column sizing.
- [ ] Charts: cached chart images, chart XML rendering, axes, labels, legends, series styling, stacked/grouped bars, line charts, combo charts, and embedded chart data.
- [ ] SmartArt/diagrams: use fallback drawings when present; otherwise emit precise diagnostics.
- [ ] Slide inheritance: deeper master/layout placeholder resolution, theme variants, background styles, footer/date/slide-number placeholders, and hidden placeholder semantics.
- [ ] Media and dynamic features: videos, audio, animations, transitions, and OLE/ActiveX should remain static/diagnostic-only unless a reliable fallback is available.
- [ ] Comments/notes: speaker notes and comments should be ignored with diagnostics or exposed through an optional mode, not silently dropped.

### PPTX Private Deck Recovery Plan

- [x] Add a private-safe PPTX slide inventory tool that reports counts and feature flags per slide: shapes, grouped shapes, pictures, charts, tables, text boxes, placeholders, inherited master/layout content, theme references, fills/effects, transforms, clips, relationships, and diagnostics without exposing slide text or images.
- [ ] Select the first private slide and exhaust it before moving on: inspect reference vs candidate, list every visible failure, implement or diagnose each generic gap, rerender, and repeat until the slide is acceptable or every remaining issue is explicitly planned.
- [ ] After the first slide is exhausted, continue one private slide at a time. Prefer the next simplest failing slide before typical and worst slides so foundational ordering/inheritance/text/image problems are isolated early.
- [ ] Maintain a private per-slide checklist with only public-safe fields: slide number, private rating, missing content, wrong order, wrong placement, wrong sizing, wrong text layout, wrong styling, and unsupported features lacking diagnostics.
- [ ] Establish private visual gates for those representative slides: page count and dimensions must match, diagnostics must explain omissions, and a private human/agent rating must improve before optimizing aggregate deck metrics.
- [ ] Fix slide composition order first: master, layout, slide background, inherited placeholders, slide shapes, groups, pictures, tables, charts, and overlays must render in PowerPoint z-order.
- [ ] Audit and fix master/layout/placeholder inheritance, including placeholder matching, hidden placeholders, text/body placeholders, footer/date/slide-number placeholders, theme variants, and background styles.
- [ ] Build slide-level diagnostics for unsupported or approximated visible content: effects, transparent fills, complex shapes, chart fallbacks, SmartArt, media/OLE, unsupported images, and placeholder fallbacks.
- [ ] Address dominant primitives after ordering/inheritance are under control: text autofit/shrink, bullets, font fallback; image placeholder crop/fit and rotation/flip; table styles and merged cells; chart cached-image fallbacks and labels.
- [ ] For every generic capability fixed from a private slide, add a small public synthetic test. Do not derive public fixtures from private slide content.
- [ ] Run `pwsh tools/CheckPrivateCase.ps1 -Case private-cases/lokad-value-based.json` after each scoped PPTX fix and summarize only counts, diagnostics, worst-page numbers, and private ratings.

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

### DOCX 16-vs-17 Page Mismatch Plan

- [ ] Select the first private page and exhaust it before moving on: inspect reference vs candidate, list every visible failure, implement or diagnose each generic gap, rerender, and repeat until the page is acceptable or every remaining issue is explicitly planned.
- [ ] Continue one private page at a time after the first page. Prefer pages with obvious table/pagination failures before using document-wide aggregate metrics.
- [ ] Maintain a private per-page checklist with only public-safe fields: page number, private rating, missing content, wrong order, wrong placement, wrong sizing, wrong text layout, table defects, pagination drift, and unsupported features lacking diagnostics.
- [ ] Add an internal DOCX layout trace mode that records public-safe per-page counts and consumed vertical space by block kind, so private runs can locate where candidate pagination drifts without exposing text.
- [x] Extend DOCX diagnostics to inspect styles and numbering parts, not just direct `word/document.xml` elements, so style-level spacing, keep rules, indents, table styles, and numbering layout risks are visible.
- [ ] Implement style-derived paragraph spacing accurately, including before/after values, `contextualSpacing`, `beforeAutospacing`/`afterAutospacing`, and Word-like adjacent paragraph spacing collapse.
- [ ] Implement paragraph and numbering indents: left/right/first-line/hanging indents from paragraph styles and numbering levels, with corresponding wrapping-width changes.
- [ ] Improve numbering layout: render labels in their own hanging-indent area, support level text expansion beyond the current simple label prefix, and honor restart/start rules.
- [ ] Improve table layout accumulation: preferred table widths, cell widths, row minimum height from content, cell vertical alignment, cell margins, and repeating header rows.
- [ ] Revisit keep rules only after layout tracing exists: support style-derived `keepNext`, `keepLines`, and widow/orphan control with synthetic tests and private page-count checks.
- [ ] Reattempt manual page/column break support with a parser change that does not alter paragraphs when no matching break exists; previous paragraph-splitting attempts changed the private page count and were reverted.
- [ ] For every generic capability fixed from a private page, add a small public synthetic test. Do not derive public fixtures from private page content.
- [ ] After each scoped fix, run `pwsh tools/CheckPrivateCase.ps1 -Case private-cases/user-requirements-spec.json` and record only page counts, aggregate metrics, diagnostics, and worst-page numbers.

### DOCX Table Recovery Plan

- [ ] Add a DOCX table inventory/trace mode that reports public-safe table metrics per table: row count, column count, grid columns, preferred table width, cell width declarations, row height declarations, header rows, vertical alignment, margins, borders, shading, grid spans, vertical merges, and page index.
- [ ] Select representative private tables for repeated inspection: one simple table, one typical dense table, and one worst table. Record only table ordinal/page, private rating, and public-safe feature gaps.
- [ ] Fix the table layout model before cosmetic styling: resolve `tblGrid`, `tblW`, `tcW`, page content width, percentage/auto widths, and grid scaling consistently.
- [ ] Compute row heights from actual cell content: wrap text within cell width, include cell margins, respect explicit `trHeight` rules, and avoid the current fixed/default row-height behavior for content-heavy rows.
- [ ] Render cell text as paragraphs instead of flattened cell text: preserve paragraph breaks, basic run styling, numbering/bullets inside cells, alignment, and line spacing.
- [ ] Implement table and cell styling: table styles, conditional first/header row formatting, cell shading, per-edge borders, border widths/colors, and vertical alignment.
- [ ] Implement structural table features: horizontal merges (`gridSpan`), vertical merges (`vMerge`), repeating header rows across page breaks, and page-break behavior inside rows.
- [ ] Add synthetic public tests for each table capability before using the private document as evidence; never derive fixtures from private table content.

### PDF/Infrastructure

- [ ] Add PDF hyperlinks, outlines/bookmarks, metadata, and optional tagged-PDF structure if needed by consumers.
- [ ] Add font subsetting to reduce output size while keeping deterministic output.
- [ ] Add image deduplication and compression choices for large decks.
- [x] Improve diagnostics severity model so embedded-image omissions are distinguishable from harmless approximations.
- [x] Add visual comparison support for dimension-near-matches, so a 1-pixel raster rounding mismatch can still produce pixel metrics.
- [x] Add private-case summary tooling that reports page count, dimension mismatches, diagnostics grouped by feature, and worst visual pages without exposing content.

## Next Implementation Targets

1. Select and exhaust the first private PPTX slide, starting with slide 1 from the inventory, before adding more isolated PPTX primitives.
2. Fix PPTX slide composition order and master/layout/placeholder inheritance based on the selected slide evidence.
3. Start DOCX table recovery for `user-requirements-spec`: table inventory/trace first, then width resolution, row-height/content wrapping, and styling.
4. Continue DOCX page geometry/pagination work using the 16-vs-17 page mismatch plan: trace drift first, then address style spacing, numbering indents, table sizing, and keep rules piecewise.
5. Improve diagnostics coverage for other visible-content omissions that still lack feature-specific diagnostics.

## Decisions

- The library remains dependency-free. Third-party packages are not allowed in `src/Lokad.OoxPdf`.
- Office and PDFium remain validation-only under `tools/`.
- Private documents remain under ignored `private-cases/`; generated private artifacts remain under ignored `artifacts/private-visual/`.
- Public notes from private documents must be anonymized to feature gaps and metrics only.
- Diagnostics must prefer continued conversion over crashing, but omitted visible content must not be treated as acceptable final behavior.
- Pixel metrics are late-stage regression evidence only. Until selected private slides/pages are mostly visually correct, do not use MAE or changed-pixel ratios to prioritize work or judge acceptability.
- Current fidelity work proceeds one private PPTX slide or one private DOCX page at a time. Exhaust the selected slide/page before moving on, and convert each generic private failure into a public synthetic test where feasible.

## Validation

Latest public validation:

```powershell
dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal
dotnet run --project tests/Lokad.OoxPdf.Tests --tl:off
dotnet pack src/Lokad.OoxPdf/Lokad.OoxPdf.csproj --tl:off --nologo -v minimal --no-restore
```

Current expected test result:

```text
75 passed, 0 failed
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
