# Lokad.OoxPdf Current Execution Plan

This ExecPlan is the current working plan for `Lokad.OoxPdf`. It intentionally omits the historical bootstrap checklist that has already been completed. Keep `Progress`, `Private Evidence`, `Backlog`, `Decisions`, and `Validation` current as work proceeds.

## Goal

`Lokad.OoxPdf` is a dependency-free .NET library that converts `.pptx` and `.docx` OOXML documents to static PDF. The library must not call Office, PDFium, PowerShell, external executables, or third-party packages. Office and PDFium are allowed only in `tools/` for validation.

The project is now past the initial vertical slice. The next phase is fidelity: use Office-exported PDFs from public synthetic OOXML as the implementation oracle, inspect those PDFs and their raster output before changing renderer behavior, use private local-only documents only to discover missing Office features, and keep diagnostics honest when a feature is still missing.

## Repository Map

- `src/Lokad.OoxPdf`: production library, OOXML parsers, renderers, PDF writer, font/image code.
- `src/Lokad.OoxPdf.Cli`: local conversion CLI.
- `tests/Lokad.OoxPdf.Tests`: dependency-free console test runner and synthetic fixtures.
- `tools/CheckVisualCase.ps1`: public visual validation harness.
- `tools/CheckPrivateCase.ps1`: private local-only visual validation harness.
- `tools/RenderReference.ps1`: Office COM reference renderer.
- `tools/RasterizePdf.ps1`: PDFium rasterization wrapper.
- `tools/InspectPdf.ps1`: Office/candidate PDF object and stream inspection wrapper.
- `tools/Lokad.OoxPdf.VisualDiff`: PNG comparison tool.
- `tools/Lokad.OoxPdf.PdfiumRasterizer`: local PDFium P/Invoke rasterizer.
- `tools/Lokad.OoxPdf.PdfInspect`: dependency-free PDF object/stream inspection tool.
- `visual-cases/`: public visual case manifests.
- `private-cases/`: ignored private manifests and inputs.
- `artifacts/`: ignored validation output.
- `docs/`: user-facing public docs.
- `docs/unit-test-audit.md`: Office-PDF-first unit-test audit and rewrite candidates.

## Reference Workflow

Every PPTX and DOCX fidelity task should start from what Office actually emits, not from OOXML interpretation alone.

1. Create or select the smallest public synthetic `.pptx` or `.docx` that isolates one feature.
2. Export that file with Office to the reference PDF and inspect the PDF/raster output: page boxes, draw order, text positions, fills/strokes, images, and pagination.
3. Inspect Office's observable PDF composition strategy for the feature: text objects and matrices, path/fill/stroke operators, clipping, image masks, transparency groups, resource reuse, and draw order.
4. Render the same OOXML with `ooxpdf`, inspect the candidate PDF/raster output, and identify the smallest visible or structural difference.
5. Prefer renderer/PDF-writer changes that converge toward Office-like PDF structure when practical, not arbitrary PDF output that only happens to match a narrow raster case.
6. Implement the smallest renderer change that closes that difference without using private document content.
7. Lock the case with a public visual manifest once it is pixel-perfect or as close as realistically possible.
8. Revisit unit tests touched by the feature: keep tests that protect public API, diagnostics, parsing, safety, and deterministic PDF structure; rewrite brittle operator-position tests when Office-PDF inspection shows a better behavioral assertion.

Private PPTX/DOCX documents remain acceptance and feature-discovery corpora. Their Office PDFs may be inspected locally to identify generic gaps, but renderer changes should be driven by public synthetic fixtures unless a safety or diagnostics issue is involved.

## Progress

- [x] Dependency-free `.slnx` solution, library, CLI, tests, visual tools, docs, public fixtures, and private validation lane exist.
- [x] NuGet package version is set to `0.1.0` for the first package.
- [x] NuGet package output is configured under ignored `artifacts/nuget/`.
- [x] OOXML package layer handles ZIP parts, content types, relationships, safe part normalization, XML hardening, and package size limits.
- [x] PDF writer emits deterministic static PDFs with pages, drawing operators, embedded TrueType/CID fonts, ToUnicode maps, JPEG passthrough, PNG image XObjects, and alpha soft masks.
- [x] CLI supports `convert input output`, `--diagnostics`, `--strict`, and exit codes `0`, `1`, `2`, and `3`.
- [x] Visual validation can export Office reference PDFs for PPTX and DOCX, rasterize reference/candidate PDFs with PDFium, compute PNG metrics, and write comparison artifacts.
- [x] Public visual validation uses Office-exported PDFs/renders as the reference oracle; fidelity work should inspect the Office PDF/raster output before treating a metric as meaningful.
- [x] Initial unit-test audit started: `pptx-ladder-02-plain-text` proves the visual gate should own exact text fidelity while unit tests avoid freezing candidate-specific `Tm`/`Tj` structure.
- [x] DOCX public ladder started with `docx-ladder-00-blank`: Office PDF reference and candidate dimensions match, diagnostics are empty, MAE `0`, changed-pixel ratio threshold 16 `0`.
- [x] DOCX `docx-ladder-01-plain-paragraph` baseline is Office-aligned: reference baseline `697.42`, candidate baseline `697.44`, diagnostics empty, MAE `0.004605`, changed-pixel ratio threshold 16 `0.000126`.
- [x] DOCX `docx-ladder-01-line-height` exact line-height baseline stack is Office-aligned: candidate baselines `691.176` and `655.176`, diagnostics empty, MAE `0.014739`, changed-pixel ratio threshold 16 `0.000394`.
- [x] DOCX `docx-ladder-01-paragraph-spacing` locks explicit auto line-height plus adjacent paragraph spacing collapse: candidate baselines `697.44` and `657.84`, diagnostics empty, MAE `0.014739`, changed-pixel ratio threshold 16 `0.000394`.
- [x] Private validation keeps inputs/manifests under ignored `private-cases/`, rejects tracked/private-unsafe paths, and writes ignored artifacts under `artifacts/private-visual/`.
- [x] PPTX parser/renderer supports slide order/size, solid backgrounds, basic rectangles/ellipses/lines/rounded rectangles, connector lines with triangle arrowheads, down-arrow preset shapes, rotation/flip, common theme colors/fonts, basic scheme luminance transforms, theme discovery through presentation or slide master relationships, common scheme color aliases, common master/layout inheritance, placeholder text bounds/styles, text boxes with body insets, line breaks, direct bullet characters, paragraph spacing, 100% default line spacing, vertical anchoring, clipping, mixed-run paragraph wrapping, basic styled text, JPEG/PNG pictures, basic crop clipping, grouped shape and picture transforms, fixed-grid tables with fills and explicit borders, static bar-chart fallback, and unsupported-feature diagnostics.
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
- Private PPTX slide-2 rerun `artifacts/private-visual/lokad-value-based/20260514-193918`:
  - Centered mixed-run paragraphs now align as one paragraph instead of independently centering each run.
  - Remaining slide-2 generic gaps include exact text metrics, small footer text fit, and exact highlight bounds.
- Private PPTX slide-2 rerun `artifacts/private-visual/lokad-value-based/20260514-195155`:
  - Text frames now allow vertical overflow unless `bodyPr vertOverflow="clip"` is set, exposing previously clipped small footer text.
  - Remaining slide-2 generic gaps are fine typography/metrics: placeholder word fit, mixed-run advance, exact highlight bounds, and line placement.
- Private PPTX slide-2 rerun `artifacts/private-visual/lokad-value-based/20260514-195905`:
  - Mixed-run cursor advance now uses resolved font metrics instead of character-count heuristics.
  - Remaining slide-2 gaps are mostly font availability/substitution, exact text fit, and highlight/underline bounds.
- Private PPTX slide-3 rerun `artifacts/private-visual/lokad-value-based/20260514-200756`:
  - Connector line shapes and the down-arrow preset now render instead of being omitted or approximated as a rectangle.
  - 84 candidate pages, all dimensions matched reference pages.
  - Mean absolute error: `15.755963`; max mean absolute error: `75.403370`; mean changed-pixel ratio at threshold 16: `0.174249`.
  - Remaining slide-3 generic gaps include inherited banner fidelity, text-frame wrapping/overlap, fine text metrics, line/fill color transforms, and icon/image placement precision.
- Private PPTX slide-3 rerun `artifacts/private-visual/lokad-value-based/20260514-202523`:
  - Scheme luminance transforms now make inherited gray banner/ribbon fills visible.
  - Mixed-run paragraph wrapping now flows runs onto shared lines instead of wrapping each run independently.
  - 84 candidate pages, all dimensions matched reference pages.
  - Slide 3 mean absolute error: `12.545817`; changed-pixel ratio at threshold 16: `0.145434`.
  - Deck mean absolute error: `15.770431`; max mean absolute error: `75.640299`; mean changed-pixel ratio at threshold 16: `0.175994`.
  - Remaining slide-3 generic gaps include text autofit/fit-to-box, fine font metrics/substitution, exact highlight bounds, and icon/image placement precision.
- Private PPTX slide-3 rerun `artifacts/private-visual/lokad-value-based/20260514-203115`:
  - Default text line spacing now uses 100% when no `lnSpc` is specified, reducing vertical drift in dense text boxes.
  - 84 candidate pages, all dimensions matched reference pages.
  - Slide 3 mean absolute error: `12.359507`; changed-pixel ratio at threshold 16: `0.144392`.
  - Deck mean absolute error: `15.804445`; max mean absolute error: `75.736192`; mean changed-pixel ratio at threshold 16: `0.176924`.
  - Remaining slide-3 generic gaps include fine font metrics/substitution, exact highlight bounds, run-boundary spacing for heavily segmented text, and icon/image placement precision.
- Private PPTX slide-4 rerun `artifacts/private-visual/lokad-value-based/20260514-204901`:
  - Slide title placeholder text now inherits bounds and master title text style instead of being silently omitted.
  - 84 candidate pages, all dimensions matched reference pages.
  - Slide 4 mean absolute error: `5.739969`; changed-pixel ratio at threshold 16: `0.062321`.
  - Deck mean absolute error: `15.834825`; max mean absolute error: `75.834314`; mean changed-pixel ratio at threshold 16: `0.177074`.
  - Remaining slide-4 generic gaps include large-title font substitution/advance, exact title size, icon/text placement precision, and fine text wrapping.
- Private PPTX slide-5 inspection from `artifacts/private-visual/lokad-value-based/20260514-204901`:
  - Slide 5 mean absolute error: `14.77`; changed-pixel ratio at threshold 16: `0.15`.
  - The dominant generic gap is chart fidelity: the current static bar fallback lacks category/value axes, labels, stacked/overlay series styling, reference lines, annotations, and Office chart style colors.
  - Text blocks and separators are present but still affected by fine text metrics and run-boundary spacing.
- Private PPTX slide-6 rerun `artifacts/private-visual/lokad-value-based/20260514-210051`:
  - PPTX table rendering now strokes only explicit visible cell borders instead of inventing black borders for every cell.
  - 84 candidate pages, all dimensions matched reference pages.
  - Slide 6 mean absolute error: `12.067719`; changed-pixel ratio at threshold 16: `0.137847`.
  - Deck mean absolute error: `15.503991`; max mean absolute error: `75.834314`; mean changed-pixel ratio at threshold 16: `0.174401`.
  - Remaining slide-6 generic gaps include table text wrapping/positioning, text color/style inheritance, title font metrics, and image/text placement precision.
- Private PPTX slide-7 rerun `artifacts/private-visual/lokad-value-based/20260514-211029`:
  - Connector line triangle arrowheads now render on the expected endpoints.
  - 84 candidate pages, all dimensions matched reference pages.
  - Slide 7 mean absolute error: `9.230441`; changed-pixel ratio at threshold 16: `0.102554`.
  - Deck mean absolute error: `15.502189`; max mean absolute error: `75.834314`; mean changed-pixel ratio at threshold 16: `0.174386`.
  - Remaining slide-7 generic gaps include custom geometry/freeform curve paths, exact connector styling, large-title text metrics, and dense paragraph spacing.
- Private PPTX slide-8 inspection from `artifacts/private-visual/lokad-value-based/20260514-211029`:
  - The dominant generic gap is grouped table-like layout fidelity: row labels, icons, red separators, and text blocks are vertically compressed or overlapped.
  - Left-side callout arrows/text are misplaced, and fine text wrapping/spacing remains weak.
  - This should be handled after a focused audit of nested group transforms, placeholder-derived text styles, and table-like grouped shape ordering.
- Private PPTX slide-9 rerun `artifacts/private-visual/lokad-value-based/20260514-211655`:
  - Rounded rectangle preset shapes now render with rounded corners instead of rectangular outlines.
  - 84 candidate pages, all dimensions matched reference pages.
  - Slide 9 mean absolute error: `18.658330`; changed-pixel ratio at threshold 16: `0.209670`.
  - Deck mean absolute error: `15.495442`; max mean absolute error: `75.801014`; mean changed-pixel ratio at threshold 16: `0.174317`.
  - Remaining slide-9 generic gaps include rotated text labels, curved connectors, exact line/shape placement, and fine text metrics/wrapping.
- Public PPTX ladder reruns:
  - `pptx-blank` at `artifacts/visual/pptx-blank/20260514-213058`: page count and dimensions matched, diagnostics were empty, MAE `0`, changed-pixel ratio threshold 16 `0`.
  - `pptx-ladder-01-solid-background` at `artifacts/visual/pptx-ladder-01-solid-background/20260514-213058`: page count and dimensions matched, diagnostics were empty, MAE `0`, changed-pixel ratio threshold 16 `0`.
  - `pptx-ladder-01-master-background` at `artifacts/visual/pptx-ladder-01-master-background/20260514-213144`: page count and dimensions matched, diagnostics were empty, MAE `0`, changed-pixel ratio threshold 16 `0`.
  - The visual harness now enforces optional manifest gates for page count, dimensions, maximum MAE, maximum changed-pixel ratio, and empty diagnostics.
- Public PPTX ladder text rerun:
  - `pptx-ladder-02-plain-text` at `artifacts/visual/pptx-ladder-02-plain-text/20260514-213736`: page count and dimensions matched, diagnostics were empty, MAE `0.043046`, changed-pixel ratio threshold 16 `0.000511`.
  - First-line PPTX text baselines now use a lower Office-aligned offset for plain top-anchored text boxes.
- Public PPTX ladder text-flow rerun:
  - `pptx-ladder-03-text-flow` at `artifacts/visual/pptx-ladder-03-text-flow/20260514-214108`: page count and dimensions matched, diagnostics were empty, MAE `0.414685`, changed-pixel ratio threshold 16 `0.003763`.
  - Single-run centered paragraphs now receive the same alignment offset as mixed-run centered paragraphs.
  - `pptx-ladder-03-text-anchor-overflow` at `artifacts/visual/pptx-ladder-03-text-anchor-overflow/20260514-214416`: page count and dimensions matched, diagnostics were empty, MAE `0.413942`, changed-pixel ratio threshold 16 `0.003552`.
  - Clipped PPTX text frames now suppress lines that cannot fit inside the clip box.
- Private PPTX acceptance rerun `artifacts/private-visual/lokad-value-based/20260514-223939`:
  - 84 candidate pages, all dimensions matched reference pages.
  - Diagnostics: 9 chart static fallback informational diagnostics.
  - Slide 1 mean absolute error: `18.001507`; changed-pixel ratio at threshold 16: `0.147195`.
  - Slide-1 public-safe gaps map to styled text/list fixtures: large centered serif title placement, white text over dark image background, underlined run bounds, and multi-paragraph bullet list wrapping/line spacing.
- Public PPTX ladder styled-text rerun:
  - `pptx-ladder-04-bullet-wrap` at `artifacts/visual/pptx-ladder-04-bullet-wrap/20260514-224326`: page count and dimensions matched, diagnostics were empty, MAE `1.253179`, changed-pixel ratio threshold 16 `0.009518`.
  - Bullet glyph placement, hanging indents, and continuation-line alignment are now covered by a public synthetic fixture before using private slide-1 evidence.
  - `pptx-ladder-04-serif-title-underline` at `artifacts/visual/pptx-ladder-04-serif-title-underline/20260514-224715`: page count and dimensions matched, diagnostics were empty, MAE `0.481747`, changed-pixel ratio threshold 16 `0.005513`.
  - Large centered serif text over a dark background and underlined mixed-run text are now covered by a public synthetic fixture before using private slide-1 evidence.
  - `pptx-ladder-04-mixed-paragraph-stack` at `artifacts/visual/pptx-ladder-04-mixed-paragraph-stack/20260514-225350`: page count and dimensions matched, diagnostics were empty, MAE `3.041919`, changed-pixel ratio threshold 16 `0.027087`.
  - This fixture is intentionally not gated yet. It exposes remaining Ladder 4 gaps in multi-paragraph vertical rhythm, mixed large/small text stacks, and exact serif text metrics; a broad default line-spacing change was rejected because it regressed existing text tests and worsened this fixture.
  - `pptx-ladder-04-paragraph-advance` at `artifacts/visual/pptx-ladder-04-paragraph-advance/20260514-230129`: page count and dimensions matched, diagnostics were empty, MAE `0.340723`, changed-pixel ratio threshold 16 `0.003961`.
  - Consecutive paragraph advance now uses an Office-like default without changing intra-paragraph line breaks/wraps or explicit `lnSpc`; the new fixture is gated, and `pptx-ladder-04-bullet-wrap` was tightened after improving to MAE `0.740986`, changed-pixel ratio threshold 16 `0.006748`.
  - `pptx-ladder-04-empty-paragraph-gap` at `artifacts/visual/pptx-ladder-04-empty-paragraph-gap/20260514-231355`: page count and dimensions matched, diagnostics were empty, MAE `2.109550`, changed-pixel ratio threshold 16 `0.017209`.
  - Formatting-only empty paragraphs now consume vertical advance using their own paragraph/run formatting or the paragraph default, without borrowing a preceding large-title font.
  - `pptx-ladder-04-bold-italic-face` at `artifacts/visual/pptx-ladder-04-bold-italic-face/20260514-233419`: page count and dimensions matched, diagnostics were empty, MAE `2.512482`, changed-pixel ratio threshold 16 `0.018524`.
  - Font resolution now selects bold/italic faces when available, and synthetic bold/italic is applied only when the requested face cannot be resolved.
  - Split font-face anchors:
    - `pptx-ladder-04-bold-face-single` at `artifacts/visual/pptx-ladder-04-bold-face-single/20260514-234212`: MAE `0.229329`, changed-pixel ratio threshold 16 `0.002122`.
    - `pptx-ladder-04-italic-face-single` at `artifacts/visual/pptx-ladder-04-italic-face-single/20260514-234221`: MAE `0.259048`, changed-pixel ratio threshold 16 `0.002407`.
    - `pptx-ladder-04-bold-italic-face-single` at `artifacts/visual/pptx-ladder-04-bold-italic-face-single/20260514-234229`: MAE `0.253966`, changed-pixel ratio threshold 16 `0.002414`.
  - `pptx-ladder-04-character-spacing` at `artifacts/visual/pptx-ladder-04-character-spacing/20260514-235026`: page count and dimensions matched, diagnostics were empty, MAE `0.920131`, changed-pixel ratio threshold 16 `0.006839`.
  - Run-level `spc` character spacing now affects text advance, wrapping, PDF text state, underline/highlight extents, and fallback measurement.
  - `pptx-ladder-04-baseline-shift` at `artifacts/visual/pptx-ladder-04-baseline-shift/20260514-235445`: page count and dimensions matched, diagnostics were empty, MAE `0.269687`, changed-pixel ratio threshold 16 `0.002153`.
  - Run-level `baseline` now shifts superscript/subscript text, highlights, and underlines relative to the paragraph baseline.
  - `pptx-ladder-04-highlight-single` at `artifacts/visual/pptx-ladder-04-highlight-single/20260514-235728`: page count and dimensions matched, diagnostics were empty, MAE `0.261774`, changed-pixel ratio threshold 16 `0.004202`.
  - Run-level highlight rendering is now locked by a public visual gate for a single highlighted text run.
  - `pptx-ladder-04-line-spacing-points` at `artifacts/visual/pptx-ladder-04-line-spacing-points/20260515-000244`: page count and dimensions matched, diagnostics were empty, MAE `0.494735`, changed-pixel ratio threshold 16 `0.004269`.
  - Absolute PPTX line spacing (`a:lnSpc/a:spcPts`) now drives intra-paragraph line breaks and explicit paragraph advance in the public ladder.
  - `pptx-ladder-04-bullet-style` at `artifacts/visual/pptx-ladder-04-bullet-style/20260515-181539`: page count and dimensions matched, diagnostics were empty, MAE `0.029019`, changed-pixel ratio threshold 16 `0.000670`.
  - Bullet-specific color and point-size formatting (`a:buClr`, `a:buSzPts`) is order-aware: valid properties before the bullet marker are honored, while invalid late properties after `a:buChar` match Office's ignored behavior in the current public visual fixture.
  - `pptx-ladder-04-tab-stop` at `artifacts/visual/pptx-ladder-04-tab-stop/20260515-174515`: page count and dimensions matched, diagnostics were empty, MAE `0.011007`, changed-pixel ratio threshold 16 `0.000341`.
  - Office ignores the standalone synthetic `<a:tab/>` child used by this fixture and emits continuous text, so the renderer now matches that behavior. A separate Office-authored fixture is still needed for real tab characters/tab-stop semantics.
  - `pptx-ladder-04-strikethrough-single` at `artifacts/visual/pptx-ladder-04-strikethrough-single/20260515-001236`: page count and dimensions matched, diagnostics were empty, MAE `0.268133`, changed-pixel ratio threshold 16 `0.002249`.
  - Single-run strikethrough (`a:rPr @strike`) now renders and is locked by a public visual gate.
  - `pptx-ladder-04-all-caps` at `artifacts/visual/pptx-ladder-04-all-caps/20260515-001508`: page count and dimensions matched, diagnostics were empty, MAE `0.182496`, changed-pixel ratio threshold 16 `0.001935`.
  - Run-level all-caps text (`a:rPr @cap="all"`) now transforms text before measurement and drawing in the public ladder.
  - `pptx-ladder-03-preserved-spaces` at `artifacts/visual/pptx-ladder-03-preserved-spaces/20260515-075831`: page count and dimensions matched, diagnostics were empty, MAE `0.139198`, changed-pixel ratio threshold 16 `0.001193`.
  - Preserved spaces in PPTX runs are visually locked by a public Office-PDF-backed gate. Office emits this case as one `TJ` text object, while the candidate still emits multiple `Tj` objects split around space groups; track this as a PDF-structure improvement, not a blocker for the visual rung.
  - `pptx-ladder-04-underline-single` at `artifacts/visual/pptx-ladder-04-underline-single/20260515-080227`: page count and dimensions matched, diagnostics were empty, MAE `0.151413`, changed-pixel ratio threshold 16 `0.001580`.
  - PPTX underlines now use an Office-like filled rectangle instead of a stroked line; the older broad serif-title underline gate also improved to MAE `0.429467`, changed-pixel ratio threshold 16 `0.005498`.
  - `pptx-ladder-04-mixed-font-size-line` at `artifacts/visual/pptx-ladder-04-mixed-font-size-line/20260515-080529`: page count and dimensions matched, diagnostics were empty, MAE `0.195173`, changed-pixel ratio threshold 16 `0.001995`.
  - Same-line mixed font sizes now have their own public Office-PDF-backed gate. Office and candidate both place 36pt and 18pt runs on a shared baseline with closely matching x advances.
  - `pptx-ladder-04-mixed-font-size-stack` now isolates large/small/large paragraph vertical rhythm. Office advances line tops and recomputes each paragraph baseline from the current paragraph font size; the renderer now follows that model. The isolated fixture is gated at MAE `0.065659`, changed-pixel ratio threshold 16 `0.001726`.
  - The line-top text layout change tightened `pptx-ladder-04-mixed-font-size-line` to MAE `0.007254`, changed-pixel ratio threshold 16 `0.000173`, and `pptx-ladder-04-paragraph-advance` to MAE `0.216017`, changed-pixel ratio threshold 16 `0.003088`.
  - The same change improved the broader ungated `pptx-ladder-04-mixed-paragraph-stack` from MAE `3.041919` to MAE `0.873692`, changed-pixel ratio threshold 16 `0.011333`. Remaining gaps are now concentrated in serif font metrics, underline bounds in combined text, and wrapped bullet continuation rather than gross paragraph vertical drift.
  - Default intra-paragraph wrap/line-break advance now uses the same Office-like `1.2 * fontSize` line-top advance as paragraph advance, while explicit line spacing remains explicit. This tightened `pptx-ladder-04-bullet-wrap` to MAE `0.149908`, changed-pixel ratio threshold 16 `0.003331`.
  - Absolute PPTX line-spacing baselines now place text lower inside the fixed line box, matching the Office PDF stream for `a:spcPts`. This tightened `pptx-ladder-04-line-spacing-points` to MAE `0.305809`, changed-pixel ratio threshold 16 `0.002707`.
  - Adjacent underlined PPTX flow segments now coalesce before PDF emission so multi-word underlined runs draw one continuous underline rectangle, matching Office's continuous underline shape. `pptx-ladder-04-underline-single` tightened to MAE `0.118922`, changed-pixel ratio threshold 16 `0.001273`; `pptx-ladder-04-serif-title-underline` is now gated at MAE `0.353182`, changed-pixel ratio threshold 16 `0.004945`.
  - Classic OpenType `kern` pairs now affect PPTX text measurement and PDF emission through `TJ` arrays when available, moving the renderer closer to Office's pair-adjusted text streams. This improved `pptx-ladder-04-mixed-paragraph-stack` to MAE `0.520333`, changed-pixel ratio threshold 16 `0.009028`.
  - Adjacent same-style PPTX text runs on the same baseline now coalesce before font resolution and PDF emission, moving line text closer to Office's grouped `TJ` strategy. This tightened `pptx-ladder-03-preserved-spaces` from MAE `0.139198`, changed-pixel ratio threshold 16 `0.001193` to MAE `0.001662`, changed-pixel ratio threshold 16 `0`, and tightened `pptx-ladder-04-bullet-wrap` from MAE `0.149908`, changed-pixel ratio threshold 16 `0.003331` to MAE `0.137307`, changed-pixel ratio threshold 16 `0.003059`.
  - `pptx-ladder-07-image-crop` now locks a minimal left/right cropped PNG picture against Office at exact raster parity: page count and dimensions matched, diagnostics were empty, MAE `0`, changed-pixel ratio threshold 16 `0`.
  - `pptx-ladder-07-image-alpha` now locks a minimal transparent PNG picture over a solid slide background. Office and candidate raster output match exactly at MAE `0`, changed-pixel ratio threshold 16 `0`; the candidate PDF carries the alpha channel as a soft mask image.
  - `pptx-ladder-07-jpeg-image` now locks minimal JPEG picture placement with no diagnostics at MAE `0.134097`, changed-pixel ratio threshold 16 `0.005486`. The remaining delta is JPEG decode/re-encode edge color variance between Office and direct PDF embedding, not placement.
  - `pptx-ladder-07-image-rotation` now locks a rotated rectangular PNG picture. Pictures now reuse the Office-aligned shape transform path for rotation/flips before image emission, tightening this case from MAE `5.737500`, changed-pixel ratio threshold 16 `0.056250` to exact raster parity.
  - `pptx-ladder-07-image-flip` now locks a horizontally flipped PNG picture with exact Office raster parity, covering the non-rotated flip branch of picture transforms.
  - `pptx-ladder-07-image-rotate-flip` now locks the combined rotated plus horizontally flipped picture matrix order with exact Office raster parity.
  - `pptx-ladder-08-grouped-picture` now locks grouped picture child coordinate scaling against Office with exact raster parity, complementing the existing grouped-shape visual rung.
  - `pptx-ladder-08-grouped-text` now locks text boxes inside grouped content. Text run layout now applies ancestor group transforms before computing text bounds and clip regions, improving this fixture from MAE `0.377745`, changed-pixel ratio threshold 16 `0.002546` to MAE `0.002687`, changed-pixel ratio threshold 16 `0.000079`.
  - `pptx-ladder-08-nested-grouped-text` now locks nested group transform composition for text boxes at MAE `0.002687`, changed-pixel ratio threshold 16 `0.000079`.
  - `pptx-ladder-08-text-shape-zorder` now locks simple text/shape sibling order: a later opaque shape covers an earlier text box at exact raster parity. Simple slides without pictures or graphic frames now render shapes and their text in shape-tree order instead of in separate all-shapes/all-text layers.
  - `pptx-ladder-08-shape-picture-zorder` now locks simple shape/picture sibling order at exact raster parity. The ordered slide path now includes pictures while preserving the exact image crop rung.
  - `pptx-ladder-06-dashed-connector` now locks the Office `dash` preset for connectors. Office emits a `[4w 3w]` dash array for this case; the renderer now maps `a:prstDash val="dash"` accordingly and resets the dash state after stroking, improving the case from MAE `0.180208`, changed-pixel ratio threshold 16 `0.001736` to exact raster parity.
  - `pptx-ladder-06-dash-dot-connector` now locks the Office `dashDot` preset for connectors. Office emits a `[4w 3w 1w 3w]` dash array; the PDF graphics builder now supports arbitrary dash arrays, improving the case from MAE `0.229766`, changed-pixel ratio threshold 16 `0.002214` to exact raster parity while preserving the plain dashed connector gate.
  - `pptx-ladder-06-round-cap-connector` now locks `a:ln cap="rnd"` against Office. Office emits PDF round cap/join operators (`1 J 1 j`), and the renderer now does the same for round-capped lines, improving the case from MAE `0.069703`, changed-pixel ratio threshold 16 `0.000505` to exact raster parity.
  - `pptx-ladder-06-square-cap-connector` now locks `a:ln cap="sq"` against Office. Office emits `2 J 1 j`, and the renderer now maps square caps to the PDF projecting-square line cap, improving the case from MAE `0.089531`, changed-pixel ratio threshold 16 `0.000625` to exact raster parity.
  - `pptx-ladder-06-connector-arrow` now locks a straight connector with triangle tail arrowhead and a down-arrow preset at exact raster parity. Arrowed connectors now use Office-like filled line/triangle geometry, and the down-arrow preset now uses Office's default shoulder proportion.
  - `pptx-ladder-06-more-polygons` now locks right triangle, pentagon, hexagon, and octagon preset geometry with exact raster parity against Office-exported PDF paths.
  - `pptx-ladder-06-double-arrows` now locks left-right and up-down block arrow presets with exact raster parity against Office-exported PDF paths.
  - `pptx-ladder-06-symbol-polygons` now locks the plus/cross preset polygon with exact raster parity against Office-exported PDF paths.
  - `pptx-ladder-06-rect-callout` now locks the `wedgeRectCallout` preset polygon with exact raster parity and no unsupported-callout diagnostic; other callout presets still emit the diagnostic until individually supported.
  - Unsupported PPTX custom geometry (`a:custGeom`) and preset callout shapes now emit slide-scoped diagnostics instead of silently falling back to ordinary rectangles.
  - `pptx-ladder-10-explicit-borders` now locks a minimal 2x2 table with per-edge borders. Explicit table borders are collected, coalesced, and stroked after cell fills to match Office's PDF order and avoid double-stroking shared edges, improving the case from MAE `0.445408`, changed-pixel ratio threshold 16 `0.002479` to exact raster parity.
  - `pptx-ladder-10-vertical-align` now locks top, center, and bottom table-cell anchors through `a:tcPr anchor`. Center and bottom anchored cells shift by Office's table text-area height, improving the case from MAE `0.509194`, changed-pixel ratio threshold 16 `0.005176` to MAE `0.013784`, changed-pixel ratio threshold 16 `0.000302`.
  - `pptx-ladder-10-unstyled-grid` now locks the default grid style for tables without a table style id. The renderer now uses black unstyled grid lines on row boundaries while preserving the existing styled-table white grid behavior, improving this fixture from MAE `0.467088`, changed-pixel ratio threshold 16 `0.002770` to exact raster parity.
  - `pptx-ladder-10-horizontal-merge` now locks a minimal horizontally merged table cell with exact raster parity. Table layout now honors `a:tc @gridSpan` and skips `hMerge`/`vMerge` continuations, and default unstyled grids suppress the internal vertical segment covered by a horizontal merge.
  - `pptx-ladder-10-vertical-merge` now locks a minimal vertically merged table cell with exact raster parity. Table layout now honors `a:tc @rowSpan`, and default unstyled grids suppress the internal horizontal segment covered by a vertical merge.
- Private PPTX rerun `artifacts/private-visual/lokad-value-based/20260514-232256`:
  - 84 candidate pages, all dimensions matched reference pages.
  - Diagnostics: 9 chart static fallback informational diagnostics.
  - Slide 1 mean absolute error: `15.366243`; changed-pixel ratio at threshold 16: `0.130552`.
  - The formatted-empty-paragraph fix materially improved slide-1 title/body separation. Remaining slide-1 generic gaps are fine font metrics, exact title/body baseline placement, underline bounds, and dense bullet/list wrapping precision.
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
- [ ] Shapes: more preset geometries, freeform paths, callouts, custom geometry, compound paths, and accurate line joins/caps/dashes.
- [ ] Fills/effects: transparency, gradients, pattern fills, picture fills, shadows, glows, reflections, blur, soft edges, and 3D effects.
- [ ] Images: placeholder-bound image placement, crop modes, rotation/flip interactions, recolor/duotone, transparency, SVG/EMF/WMF, TIFF/GIF/BMP, and image compression variants.
- [ ] Tables: merged cells, vertical alignment, per-edge borders, table styles, cell margins, rich text inside cells, and precise row/column sizing.
- [ ] Charts: cached chart images, chart XML rendering, axes, labels, legends, series styling, stacked/grouped bars, line charts, combo charts, and embedded chart data.
- [ ] SmartArt/diagrams: use fallback drawings when present; otherwise emit precise diagnostics.
- [ ] Slide inheritance: deeper master/layout placeholder resolution, theme variants, background styles, footer/date/slide-number placeholders, and hidden placeholder semantics.
- [ ] Media and dynamic features: videos, audio, animations, transitions, and OLE/ActiveX should remain static/diagnostic-only unless a reliable fallback is available.
- [ ] Comments/notes: speaker notes and comments should be ignored with diagnostics or exposed through an optional mode, not silently dropped.

### PPTX Synthetic Fidelity Ladder

Build these as public, minimal, one-slide fixtures. Each rung must start with Office PDF/raster inspection, then compare the candidate PDF/raster output, then receive strict page/dimension checks, expected diagnostics, and a visual gate once the primitive is close. It is acceptable for private deck pages to regress while early rungs are rebuilt; the goal is a strict bottom-up progression.

- [x] Ladder 0: blank slide, page size, white/default background, deterministic PDF bytes, and no diagnostics.
- [x] Ladder 1: solid slide backgrounds and master/layout background inheritance in isolation.
- [x] Ladder 2: one plain text box with fixed bounds, one font size, one font family, no wrapping, and baseline locked against reference.
- [x] Ladder 3: text wrapping with preserved spaces, explicit line breaks, tabs, paragraph alignment, body insets, vertical anchor, and overflow behavior. Current public gates: `pptx-ladder-03-preserved-spaces` at exact raster parity, `pptx-ladder-03-text-flow` at MAE `0.065508` and changed16 `0.001155`, and `pptx-ladder-03-text-anchor-overflow` at MAE `0.141281` and changed16 `0.001517`.
- [ ] Ladder 4: styled text runs: bold, italic, underline, color, highlight, mixed fonts, bullet glyphs, bullet hanging indents, paragraph spacing, and line spacing. Subcase `pptx-ladder-04-character-spacing` now performs expanded-spacing line breaking during layout and emits Office-like `TJ` per-glyph tracking arrays instead of draw-time `Tc`; it is gated at MAE `0.165546`, changed-pixel ratio threshold 16 `0.001458`. Font-style subcases now keep regular, bold, italic, and bold-italic face resources distinct: `pptx-ladder-04-bold-face-single` is gated at MAE `0.017018`, changed16 `0.000461`; `pptx-ladder-04-italic-face-single` at MAE `0.010782`, changed16 `0.000346`; `pptx-ladder-04-bold-italic-face-single` at MAE `0.022819`, changed16 `0.000641`; and the combined `pptx-ladder-04-bold-italic-face` at MAE `0.111322`, changed16 `0.003214`. Baseline-shifted runs now render at Office's two-thirds glyph size while keeping the declared-size baseline offset, locking `pptx-ladder-04-baseline-shift` at MAE `0.017311`, changed16 `0.000467`. Explicit-break baselines now use the next line's run size instead of a fixed default, tightening `pptx-ladder-04-line-spacing-points` to MAE `0.014381`, changed16 `0.000410`. Percentage line spacing now follows Office's percent-of-normal-line-advance behavior: `pptx-ladder-04-empty-paragraph-gap` is gated at MAE `0.306072`, changed16 `0.005082`, and `pptx-ladder-04-paragraph-advance` at MAE `0.163712`, changed16 `0.002857`. `pptx-ladder-04-tab-stop` now matches Office's continuous-text handling for the standalone synthetic `<a:tab/>` fixture at MAE `0.011007`, changed16 `0.000341`; real Office-authored tab character semantics remain to be isolated. Bullet styling is now order-aware: `pptx-ladder-04-bullet-style` is gated at MAE `0.029019`, changed16 `0.000670`, while valid `buClr`/`buSzPts` before the marker remain unit-tested. Underline geometry now uses Office-like baseline offset and thickness: `pptx-ladder-04-underline-single` is gated at MAE `0.065736`, changed16 `0.000936`, and `pptx-ladder-04-serif-title-underline` at MAE `0.255534`, changed16 `0.004376`. Strikethrough now uses an Office-like filled rectangle at the strike baseline, locking `pptx-ladder-04-strikethrough-single` at MAE `0.008026`, changed16 `0.000115`. Subcase `pptx-ladder-04-mixed-font-size-stack` is locked at MAE `0.071192`, changed-pixel ratio threshold 16 `0.001852`, and `pptx-ladder-04-mixed-paragraph-stack` is now gated at MAE `0.452275`, changed-pixel ratio threshold 16 `0.008806`; remaining combined-stack gaps are glyph/font details, especially bullet glyph font selection.
- [x] Ladder 5: basic shapes: rectangle, rounded rectangle, ellipse, line, fills, strokes, stroke widths, rotation, flips, and clipping-free z-order. `pptx-ladder-05-basic-shapes` reuses the public `pptx-shapes` fixture and is gated at MAE `0.009024`, changed-pixel ratio threshold 16 `0.000356`, with no diagnostics.
- [ ] Ladder 6: preset and connector shapes: arrows, connector endpoints, arrowheads, dashes, line caps/joins, callouts, and common freeform/custom path fallbacks. Subcase `pptx-ladder-06-connector-arrow` locks a straight connector with triangle tail arrowhead and a down-arrow preset with exact raster parity; subcase `pptx-ladder-06-line-arrowheads` locks Office `type="arrow"` head and both-end connector arrowheads at near-exact raster parity (MAE <= `0.002`, changed-pixel ratio threshold 16 <= `0.00003`); subcase `pptx-ladder-06-block-arrows` locks up/left/right block-arrow preset geometry with exact raster parity; subcase `pptx-ladder-06-double-arrows` locks left-right and up-down block arrow presets with exact raster parity; subcase `pptx-ladder-06-basic-polygons` locks triangle and diamond preset geometry with exact raster parity; subcase `pptx-ladder-06-quadrilateral-presets` locks parallelogram and trapezoid preset geometry with exact raster parity; subcase `pptx-ladder-06-more-polygons` locks right triangle, pentagon, hexagon, and octagon preset geometry with exact raster parity; subcase `pptx-ladder-06-chevron-home-plate` locks chevron and home-plate preset geometry with exact raster parity; subcase `pptx-ladder-06-star-seal` locks five-point and six-point star preset geometry with exact raster parity; subcase `pptx-ladder-06-more-stars` locks four-point and eight-point star preset geometry with exact raster parity; subcase `pptx-ladder-06-symbol-polygons` locks the plus/cross preset polygon with exact raster parity; subcase `pptx-ladder-06-rect-callout` locks the `wedgeRectCallout` preset with exact raster parity; subcases `pptx-ladder-06-dashed-connector` and `pptx-ladder-06-dash-dot-connector` lock Office dash presets with exact raster parity; subcases `pptx-ladder-06-round-cap-connector` and `pptx-ladder-06-square-cap-connector` lock round and square caps/joins with exact raster parity; subcases `pptx-ladder-06-bevel-join-rect` and `pptx-ladder-06-round-join-rect` lock explicit bevel/round line joins on stroked rectangles with exact raster parity; custom geometry and unsupported callout shapes emit explicit unsupported diagnostics; remaining subcases should isolate additional visual callout rendering and other preset geometries.
- [ ] Ladder 7: images: JPEG/PNG placement, alpha masks, crop rectangles, aspect-fit/fill behavior, rotation/flip interactions, and unsupported image diagnostics. Subcase `pptx-ladder-07-basic-image` locks the existing public stretched-image fixture with exact raster parity (MAE `0`, changed-pixel ratio threshold 16 `0`), `pptx-ladder-07-image-crop` locks a minimal left/right cropped PNG picture with exact raster parity, `pptx-ladder-07-image-fill-rect` locks destination `a:fillRect` placement with exact raster parity, `pptx-ladder-07-image-crop-fill-rect` locks source crop combined with destination fill-rect placement with exact raster parity, `pptx-ladder-07-image-alpha` locks a transparent PNG soft-mask case with exact raster parity, `pptx-ladder-07-jpeg-image` locks minimal JPEG placement at MAE `0.134097`, changed-pixel ratio threshold 16 `0.005486`, `pptx-ladder-07-image-rotation` locks rotated picture transforms with exact raster parity, `pptx-ladder-07-image-flip` locks horizontal picture flips with exact raster parity, and `pptx-ladder-07-image-rotate-flip` locks combined rotation/flip with exact raster parity; remaining subcases should isolate aspect-fit/fill variants and unsupported image diagnostics.
- [ ] Ladder 8: grouped content: nested group transforms, grouped pictures, grouped text, grouped shapes, child coordinate scaling, z-order, and clips. Subcase `pptx-ladder-08-grouped-shape` locks grouped shape child coordinate scaling at MAE `0.000142`, changed-pixel ratio threshold 16 `0`; subcase `pptx-ladder-08-grouped-picture` locks grouped picture scaling with exact raster parity; subcase `pptx-ladder-08-grouped-text` locks text boxes inside grouped content at MAE `0.002687`, changed-pixel ratio threshold 16 `0.000079`; subcase `pptx-ladder-08-nested-grouped-text` locks nested group transform composition for text boxes at the same thresholds; subcases `pptx-ladder-08-text-shape-zorder` and `pptx-ladder-08-shape-picture-zorder` lock simple sibling order with exact raster parity; subcase `pptx-ladder-08-table-shape-zorder` locks table graphic frames participating in slide sibling order at MAE `0.062251`, changed-pixel ratio threshold 16 `0.000748`; remaining subcases should isolate z-order with charts and clipping.
- [ ] Ladder 9: slide inheritance: placeholders, master/layout text styles, hidden placeholders, footer/date/slide number placeholders, theme fonts, and theme color transforms. Subcase `pptx-ladder-09-title-placeholder` isolates a slide title placeholder inheriting layout bounds, 60pt centered style, bottom anchor, and master title run defaults at MAE `1.001682`, changed-pixel ratio threshold 16 `0.006872`, using Calibri-family fallback when Aptos theme fonts are unavailable; subcase `pptx-ladder-09-title-placeholder-arial-theme` removes that font variable and locks the same inherited placeholder geometry at MAE `0.215164`, changed-pixel ratio threshold 16 `0.002463`.
- [ ] Ladder 10: tables: fixed grid, per-edge borders, fills, cell margins, vertical alignment, merged cells, rich text inside cells, and table styles. Subcase `pptx-ladder-10-basic-table` now locks a fixed-grid table with explicit cell fills, default Office grid lines, and Office-like cell text insets/baselines at MAE `0.049042`, changed-pixel ratio threshold 16 `0.000916`; subcase `pptx-ladder-10-explicit-borders` locks per-edge borders with exact raster parity; subcase `pptx-ladder-10-vertical-align` locks top/center/bottom table-cell anchors at MAE `0.013784`, changed-pixel ratio threshold 16 `0.000302`; subcase `pptx-ladder-10-unstyled-grid` locks unstyled default grid lines with exact raster parity; subcases `pptx-ladder-10-horizontal-merge` and `pptx-ladder-10-vertical-merge` lock merged cells with exact raster parity; subcase `pptx-ladder-10-rich-text-cell` locks styled table-cell run sequencing and whitespace-preserving run gaps at MAE `0.092446`, changed-pixel ratio threshold 16 `0.001351`; subcase `pptx-ladder-10-cell-margins` locks table-cell body insets at MAE `0.259091`, changed-pixel ratio threshold 16 `0.002151`; subcase `pptx-ladder-10-cell-fill-alpha` locks table-cell fill transparency at MAE `0.035698`, changed-pixel ratio threshold 16 `0.000926`; subcase `pptx-ladder-10-border-alpha` locks explicit table-border transparency at MAE `0.223404`, changed-pixel ratio threshold 16 `0.002494`; remaining subcases should isolate broader table style variants.
- [ ] Ladder 11: charts: cached image fallback, basic bar/line/pie rendering, axes, labels, legends, series styles, stacked/grouped variants, and chart diagnostics.
- [ ] Ladder 12: effects and advanced fills: transparency, gradients, pattern fills, shadows, glows, soft edges, picture fills, and explicit diagnostics for unsupported effects. Unsupported gradient fills, pattern fills, shape picture fills, and effect lists now emit slide-scoped diagnostics instead of being silently dropped; subcases `pptx-ladder-12-solid-alpha` and `pptx-ladder-12-line-alpha` lock solid shape fill transparency and line transparency with exact raster parity; subcase `pptx-ladder-12-arrow-alpha` locks transparent Office `type="arrow"` connector arrowheads at MAE `0.006754`, changed-pixel ratio threshold 16 `0.000174` after Office-style endpoint geometry; subcase `pptx-ladder-12-picture-alpha` locks whole-picture `alphaModFix` transparency at MAE `0.109499` with no pixels changed above threshold 16; subcase `pptx-ladder-12-text-alpha` locks run text transparency at MAE `0.006230`, changed-pixel ratio threshold 16 `0.000170`; subcase `pptx-ladder-12-shadow-diagnostic` locks explicit shadow/effect diagnostics while allowing the expected omitted-shadow visual delta at MAE `0.408920`, changed-pixel ratio threshold 16 `0.006530`; subcases `pptx-ladder-12-gradient-diagnostic`, `pptx-ladder-12-pattern-diagnostic`, and `pptx-ladder-12-picture-fill-diagnostic` lock explicit advanced-fill diagnostics while allowing their expected omitted-fill visual deltas; remaining work is visual rendering for each effect/fill family.
- [ ] For every ladder rung, keep public synthetic fixture content artificial and minimal. Do not derive fixture text, images, layout, or styling from private documents.
- [ ] Run the relevant public visual case after each rung change; run private PPTX only as feature-discovery smoke evidence until the public ladder is much more complete.
- [ ] Revisit PPTX unit tests under the Office-PDF-first workflow: keep parser/safety/API tests, keep deterministic low-level PDF writer tests, but replace brittle renderer operator-position assertions with public visual gates or assertions derived from inspected Office reference geometry.

### PPTX Private Deck Recovery Plan

- [x] Add a private-safe PPTX slide inventory tool that reports counts and feature flags per slide: shapes, grouped shapes, pictures, charts, tables, text boxes, placeholders, inherited master/layout content, theme references, fills/effects, transforms, clips, relationships, and diagnostics without exposing slide text or images.
- [ ] Revisit `lokad-value-based` one slide at a time as an acceptance corpus. For each slide, inspect reference vs candidate, list only generic public-safe gaps, map each gap to existing ladder rungs, and rerender after every relevant gated public fixture change.
- [ ] If a private-slide gap is not already covered by a passing public ladder fixture, create or tighten a minimal public synthetic fixture first. Do not implement private-slide-driven renderer changes until the corresponding public fixture is close to pixel-perfect and gated.
- [ ] After a slide is close to pixel-perfect or every remaining gap is covered by explicit planned public fixtures, continue to the next slide. Private slides should test combinations and acceptance, not replace the bottom-up ladder.
- [ ] Maintain a private per-slide checklist with only public-safe fields: slide number, private rating, missing content, wrong order, wrong placement, wrong sizing, wrong text layout, wrong styling, and unsupported features lacking diagnostics.
- [ ] Establish private visual gates for accepted slides: page count and dimensions must match, diagnostics must explain omissions, and private human/agent rating must be close to pixel-perfect before optimizing aggregate deck metrics.
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

### DOCX Synthetic Fidelity Ladder

Build a DOCX ladder comparable to the PPTX ladder. Each rung must be public, synthetic, minimal, Office-PDF-inspected, visually gated when close, and free of private content.

- [x] Ladder 0: blank document, page size, margins, deterministic PDF structure, and no diagnostics.
- [ ] Ladder 1: plain paragraphs with Word reference baselines, line height, paragraph before/after spacing, and page margins. Subcases `docx-ladder-01-plain-paragraph`, `docx-ladder-01-line-height`, and `docx-ladder-01-paragraph-spacing` are gated; remaining subcases should isolate multi-line flow and wrapping.
- [ ] Ladder 2: run styling: bold/italic face selection, underline, strikethrough, color, highlight, superscript/subscript, and font fallback.
- [ ] Ladder 3: paragraph layout: tabs/tab stops, indents, hanging indents, alignment, spacing collapse, nonbreaking spaces, soft line breaks, and manual page breaks.
- [ ] Ladder 4: numbering and bullets: label area, hanging indents, level text expansion, restart/start rules, symbol bullets, and multi-level lists.
- [ ] Ladder 5: tables: grid widths, preferred widths, cell margins, row height from content, borders, shading, vertical alignment, merged cells, and page breaks.
- [ ] Ladder 6: images and drawings: inline images, anchored/floating drawings, wrap modes, cropping, rotation, shapes, and unsupported drawing diagnostics.
- [ ] Ladder 7: headers/footers and fields: first/odd/even variants, section variants, PAGE/NUMPAGES/date/property fields, and distance from edge.
- [ ] Revisit DOCX unit tests under the Office-PDF-first workflow: prefer assertions about parsed model, diagnostics, page counts, public API, and visual gates over fragile text-coordinate/operator expectations where Word PDF inspection gives the real behavior.

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

- [x] Add a PDF inspection tool for Office/candidate PDFs that lists objects and extracts decodable streams for content-operator inspection.
- [ ] Audit current PDF generation patterns against Office reference PDFs: text grouping, text matrices, clipping regions, image masks, transparency state, path construction, stroke/fill order, resource naming/reuse, and page content stream organization.
- [ ] Improve PPTX text-line emission toward Office-like text objects: preserve visual behavior while avoiding unnecessary run splitting around spaces when the line can be emitted as one positioned text object.
- [ ] Refactor PDF rendering primitives where Office-like structure is more robust for fidelity, while preserving deterministic output and keeping `src/Lokad.OoxPdf` dependency-free.
- [ ] Add PDF hyperlinks, outlines/bookmarks, metadata, and optional tagged-PDF structure if needed by consumers.
- [ ] Add font subsetting to reduce output size while keeping deterministic output.
- [ ] Add image deduplication and compression choices for large decks.
- [x] Improve diagnostics severity model so embedded-image omissions are distinguishable from harmless approximations.
- [x] Add visual comparison support for dimension-near-matches, so a 1-pixel raster rounding mismatch can still produce pixel metrics.
- [x] Add private-case summary tooling that reports page count, dimension mismatches, diagnostics grouped by feature, and worst visual pages without exposing content.

## Next Implementation Targets

1. Audit the existing unit tests against the Office-PDF-first workflow. Keep strong low-level tests, and list renderer tests that should be replaced or complemented by public visual fixtures.
2. Continue PPTX Ladder 4 bottom-up: inspect the Office reference PDF/raster for each minimal styled-text fixture, tighten each fixture toward near pixel-perfect output, and only then revisit larger combinations.
3. Start a public DOCX synthetic ladder before optimizing private DOCX pages: inspect Office PDFs for blank/plain paragraph/table primitives, then add visual gates.
4. Use `lokad-value-based` and `user-requirements-spec` only to discover public-safe feature gaps. Do not optimize for private MAE or changed-pixel ratios while public ladders are incomplete.
5. For every private-discovered gap not covered by a passing public fixture, create or tighten the smallest public synthetic fixture first.

## Decisions

- The library remains dependency-free. Third-party packages are not allowed in `src/Lokad.OoxPdf`.
- Office and PDFium remain validation-only under `tools/`.
- Private documents remain under ignored `private-cases/`; generated private artifacts remain under ignored `artifacts/private-visual/`.
- Public notes from private documents must be anonymized to feature gaps and metrics only.
- Diagnostics must prefer continued conversion over crashing, but omitted visible content must not be treated as acceptable final behavior.
- Pixel metrics are late-stage regression evidence only. Until selected private slides/pages are mostly visually correct, do not use MAE or changed-pixel ratios to prioritize work or judge acceptability.
- Office-exported PDFs are the primary fidelity reference. Raster metrics are useful gates after manual/agent inspection confirms the candidate is targeting the same Office behavior.
- Emulate Office's observable rendering strategy where it matters for fidelity: PDF operator structure, resource usage, clipping, transparency, image placement, and text positioning should move toward Office-like patterns when practical. Do not depend on Office at runtime or claim byte-for-byte PDF equivalence.
- PPTX fidelity is bottom-up: minimal public synthetic fixtures are made close to pixel-perfect and gated before larger public combinations or private documents matter.
- Private PPTX pages may regress while lower public rungs are rebuilt. Until the public ladder is feature-complete enough, private MAE and changed-pixel ratios are smoke evidence only, not implementation targets.
- DOCX fidelity should move to the same Office-PDF-first public ladder as PPTX. Private pages remain acceptance evidence and gap discovery, not the main implementation driver.

## Validation

Latest public validation:

```powershell
dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal
dotnet run --project tests/Lokad.OoxPdf.Tests --tl:off
dotnet pack src/Lokad.OoxPdf/Lokad.OoxPdf.csproj --tl:off --nologo -v minimal --no-restore
```

Current expected test result:

```text
90 passed, 0 failed
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
