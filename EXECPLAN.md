# Lokad.OoxPdf Current Execution Plan

This ExecPlan is the current working plan for `Lokad.OoxPdf` and is maintained under `PLANS.md` in the
repository root. It intentionally omits the historical bootstrap checklist that has already been completed.
Keep `Progress`, `Private Evidence`, `Backlog`, `Decision Log`, `Validation`, `Surprises & Discoveries`, and
`Outcomes & Retrospective` current as work proceeds.

## Purpose / Big Picture

`Lokad.OoxPdf` is a dependency-free .NET library that converts `.pptx` and `.docx` OOXML documents to static
PDF. The library must not call Office, PDFium, PowerShell, external executables, or third-party packages.
Office and PDFium are allowed only in `tools/` for validation.

The project is now past the initial vertical slice. The next phase is fidelity: use Office-exported PDFs from
public synthetic OOXML as the implementation oracle, inspect those PDFs and their raster output before
changing renderer behavior, use private local-only documents only to discover missing Office features, and
keep diagnostics honest when a feature is still missing.

## Context and Orientation

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

The external reference renderer used for PPTX prior art lives outside this repository at
`C:\Users\JoannesVermorel\code\pptx-renderer`. It is a local TypeScript renderer and must not become a
runtime dependency of `Lokad.OoxPdf`.

## Reference Workflow

Every PPTX and DOCX fidelity task should start from what Office actually emits, not from OOXML interpretation
alone.

1. Create or select the smallest public synthetic `.pptx` or `.docx` that isolates one feature.
2. Export that file with Office to the reference PDF and inspect the PDF/raster output: page boxes, draw
   order, text positions, fills/strokes, images, and pagination.
3. Inspect Office's observable PDF composition strategy for the feature: text objects and matrices,
   path/fill/stroke operators, clipping, image masks, transparency groups, resource reuse, and draw order.
4. Render the same OOXML with `ooxpdf`, inspect the candidate PDF/raster output, and identify the smallest
   visible or structural difference.
5. Prefer renderer/PDF-writer changes that converge toward Office-like PDF structure when practical, not
   arbitrary PDF output that only happens to match a narrow raster case.
6. Implement the smallest renderer change that closes that difference without using private document content.
7. Lock the case with a public visual manifest once it is pixel-perfect or as close as realistically possible.
8. Revisit unit tests touched by the feature: keep tests that protect public API, diagnostics, parsing,
   safety, and deterministic PDF structure; rewrite brittle operator-position tests when Office-PDF
   inspection shows a better behavioral assertion.

Private PPTX/DOCX documents remain acceptance and feature-discovery corpora. Their Office PDFs may be
inspected locally to identify generic gaps, but renderer changes should be driven by public synthetic
fixtures unless a safety or diagnostics issue is involved.

## Long-Term Architecture Track

The long-term goal is pixel-perfect visual outcome through structural alignment with Office's observable
PDF output. "Structural alignment" means that when Office represents a feature as text matrices, clipping
paths, line strokes, filled rectangles, transparency groups, form XObjects, image masks, or reusable PDF
resources, `ooxpdf` should move toward the same kind of PDF structure instead of compensating with
feature-specific coordinates or raster-only nudges. The renderer still does not need byte-for-byte Office
PDFs, and it must remain dependency-free at runtime, but Office-exported PDFs should define the behavior.

The next architecture push is seven linked work tracks:

1. Build chart structural oracle tooling. The current PDF inspection tools are strong for text; charts need
   comparable public tooling that extracts and compares plot-area rectangles, axis/gridline line coordinates,
   legend/title/data-label text positions, marker geometry, and clipping/resource structure from Office and
   candidate PDFs. This should start with public synthetic chart fixtures and should not depend on private
   deck content.
2. Complete the chart scene model. Chart XML reads should move into typed scene/layout records for chart
   kinds, plot areas, axes, series, data labels, markers, title, legend, fills, strokes, and text styles.
   Raw XML may remain attached as source evidence while an OOXML surface is incomplete, but production
   rendering should increasingly consume typed chart data.
3. Replace chart fallback geometry. Named constants under `PptxChartMetricRules` are useful inventory, not
   the destination. Each chart ratio, offset, and text metric should either be replaced by an Office-PDF
   observed rule or explicitly classified as a temporary approximation with a public case that demonstrates
   the remaining gap.
4. Continue the slide-17 typography path through public evidence. The current private evidence suggests the
   remaining issue is broader text placement, not connector geometry or chart manual layout. Use private
   slide evidence only to identify generic missing text behavior, then lock that behavior with public
   synthetic Office-PDF-backed fixtures.
5. Migrate text frames to the model-first path. PPTX text should keep its four observable stages explicit:
   style cascade, line layout, glyph positioning, and PDF emission. Shared text-frame layout should become
   the source for shape text, chart text, and table text instead of each surface estimating height and
   baselines independently.
6. Converge table text with the common text model. Table-cell text should stop using table-local height and
   vertical-alignment estimates where the shared PPTX text-frame model can provide the same observable
   Office behavior. The long-term target is one text layout engine with surface-specific frame geometry.
7. Replace one-off OOXML enum handling with explicit ladders. Whenever an incomplete enum family is touched,
   survey the full family, record unsupported values, and add public fixtures or diagnostics in a ladder.
   Priority families remain text orientation and autofit, line dash/cap/join and arrows, preset geometry and
   adjustments, fills/color transforms, picture crop/tile/recolor, table borders/anchors, chart types and
   markers, and DOCX layout enums.

## External Renderer Survey

`C:\Users\JoannesVermorel\code\pptx-renderer` is a local TypeScript PPTX renderer that is reportedly much
closer to Office output on `lokad-value-based`. Treat it as high-priority prior art for architecture and
feature behavior, not as code to copy mechanically.

Initial survey findings:

- It uses a three-stage pipeline: parse ZIP/XML/relationships, build a normalized presentation model, then
  render slides from that model.
- Its model layer has typed slide nodes for shapes, pictures, tables, groups, and charts, with raw XML kept
  as opaque source when needed.
- Its render context resolves slide -> layout -> master -> theme once, then passes that context through
  dedicated background, shape, text, table, image, group, and chart renderers.
- Text rendering is driven by a merged inheritance cascade: master defaults, master text style, master
  placeholder, layout placeholder, shape list style, paragraph properties, and run properties.
- Its text style cascade keeps master `defaultTextStyle`, master `txStyles`, master placeholder `lstStyle`,
  layout placeholder `lstStyle`, shape `lstStyle`, paragraph `pPr`, and run `rPr` as explicit layers.
- Its placeholder resolver separately inherits geometry and text `bodyPr` from layout/master placeholders,
  with matching by `idx` before type and special handling for default/body placeholders.
- Its color resolver applies master/layout color-map remapping before theme color lookup and supports more
  color forms/modifiers than `ooxpdf` currently does, including `phClr`, preset/HSL/scrgb colors, gradients,
  and broader modifier handling.
- Its shape renderer resolves theme format-scheme `fillRef` and `lnRef` through `fillStyleLst` and
  `lnStyleLst`, replacing `phClr` with the reference color. `ooxpdf` now has the same basic solid
  fill/line path for default Office shapes.
- Its generated typography oracle ladder includes public cases for fonts, sizes, styles, alignments, colors,
  mixed formatting, bullets, vertical text, anchoring, and line spacing.
- Its visual loop is explicitly Office-oracle based, with generated public cases, PowerPoint PDF ground
  truth, SSIM/color-histogram pass gates, and MAE kept diagnostic.
- Its architecture keeps ownership clear: package parsing owns safe files and relationship targets; the model
  owns typed nodes and inheritance resolution; the render layer owns dispatch, draw order, lifecycle, and
  per-node error handling.
- Its render context is the key dependency boundary. It resolves slide, layout, master, and theme once, then
  exposes relationship sets, media/cache state, color cache, group-fill context, and navigation hooks to
  specialized renderers.
- Its slide renderer treats master/layout non-placeholder shapes as template content rendered behind slide
  nodes, while placeholders are inheritance templates rather than directly rendered content.
- Its model parsers keep raw XML attached to typed nodes. This is a useful compromise for `ooxpdf`: parse
  stable fields into records while keeping source XML available until every OOXML edge case has a typed home.

High-priority actions:

- [ ] Survey `pptx-renderer` feature by feature against `ooxpdf`: model objects, inheritance cascade,
  text layout, group transforms, shape geometry, fills/strokes, images, tables, charts, SmartArt, and oracle
  tooling.
- [ ] Inspect `pptx-renderer` architecturally, not just feature-by-feature: package parsing boundaries,
  normalized model ownership, render context lifetime, per-node renderers, style/color resolvers,
  diagnostics/error isolation, asset lifetime, and test/oracle pipeline.
- [ ] Convert the architectural survey into an `ooxpdf` migration design: what belongs in a presentation
  scene/model, what remains direct PDF rendering, and which abstractions should replace ad hoc XML traversal.
- [ ] Survey OOXML enumeration handling across PPTX and DOCX readers/renderers, then create explicit
  progress ladders for incomplete enum families instead of implementing one-off values. Priority families:
  PPTX text orientation (`a:bodyPr @vert`), paragraph alignment/anchor/overflow/autofit, line dash/cap/join
  and arrowhead presets, preset geometries and adjustments, fills/color transforms, picture crop/tile/recolor
  modes, table borders/anchors, chart types/marker styles, and DOCX alignment, breaks, table layout,
  numbering, underline/border/shading, drawing wrap, and section/page settings.
- [ ] Prioritize the `pptx-renderer` typography architecture before broad deck work: explicit text body,
  paragraph, run, line, and glyph-position models must replace ad hoc layout/emission decisions.
- [ ] Split PPTX text into four observable stages: style cascade, line layout, glyph positioning, and PDF
  emission. Each stage needs synthetic Office-PDF-backed cases before private-deck tuning.
- [ ] Port `pptx-renderer` text renderer unit coverage as clean `ooxpdf` tests for line spacing, paragraph
  spacing, character spacing, kerning thresholds, font fallback, EA/CS font fallback, bullets, baseline
  shifts, tabs, highlights, and no-fill/outline text where PDF support exists.
- [x] Respect the DrawingML text-body wrap mode in line layout:
  `a:bodyPr @wrap="none"` now disables automatic word wrapping in both text placement and the text-height
  estimate used by vertical anchoring, while manual line breaks remain explicit content. A synthetic PPTX
  typography test locks the no-wrap behavior.
- [x] Resolve inherited DrawingML paragraph indentation before line layout:
  direct text layout now falls back from local `a:pPr` to the default/cascaded paragraph properties for
  `marL` and `indent`, and text-frame snapshots expose resolved margin and hanging-indent values. A synthetic
  PPTX test locks emitted span placement from inherited `a:lvl1pPr` indentation.
- [x] Resolve ellipse auto-shape text origins structurally:
  PPTX text frames now apply the DrawingML ellipse preset horizontal text-rectangle inset before bodyPr
  insets for non-text-box shapes, and clipped glyph emission no longer drops glyphs merely because the
  baseline is just outside the clip. Public
  `pptx-ladder-04-typography-small-label-origin-probe` locks the Office/candidate text-op X delta at
  `0.03 pt`; vertical baseline parity remains open at about `1.27 pt`.
- [x] Resolve the first small-label middle-anchor vertical delta structurally:
  the vertical anchor estimator now uses the resolved font's OpenType typographic line box for default
  line spacing instead of assuming a CSS `1.2x` line box for text-height centering. The public small-label
  origin probe now locks Office/candidate text-op position within `0.35 pt` and decoded text content, with Y
  delta reduced from `1.27 pt` to `0.27 pt`.
- [x] Port the first `pptx-renderer` no-fill text rule:
  `a:rPr/a:noFill` now makes the run transparent while preserving its layout advance, instead of falling
  through to inherited or black text color. A synthetic PPTX typography case locks the behavior.
- [x] Port the `pptx-renderer` hyperlink color precedence rule:
  hyperlink runs without an explicit fill now use the theme `hlink` color before inherited shape/default
  colors. A synthetic theme-backed PPTX case locks the behavior.
- [x] Port the first EA/CS font-family fallback rule:
  run/default typeface resolution now considers `a:latin`, then `a:ea`, then `a:cs`, so runs that only
  specify East Asian or complex-script fonts do not fall back to the generic Latin default.
- [x] Port `pptx-renderer` theme font slots beyond Latin:
  theme parsing now stores major/minor Latin, East Asian, and complex-script typefaces, and resolves
  `+mj-ea`, `+mj-cs`, `+mn-ea`, and `+mn-cs` aliases before font measurement/emission.
- [x] Port the first broader `pptx-renderer` color-resolution slice:
  PPTX text/shape color paths now resolve `a:sysClr @lastClr`, common `a:prstClr` preset colors,
  `a:scrgbClr`, and `a:hslClr`, including alpha/modifier handling through the shared color resolver.
- [x] Port the first `pptx-renderer` text-outline rule:
  `a:rPr/a:ln` is now a first-class run style that flows through text layout into PDF text rendering mode,
  so no-fill outline text is stroked while preserving its layout advance. A synthetic PPTX typography case
  locks the solid-outline path.
- [x] Add a PDF-inspection typography harness that compares Office and candidate text matrices, TJ arrays,
  baseline positions, highlight rectangles, and clipping boxes before relying on raster metrics.
- [x] Extend PDF inspection for large private decks with page-aware, text-only extraction:
  `tools/InspectPdf.ps1 -TextOnly` skips image stream decoding and emits `PageNumber` on text operations, so
  slide/page-level PDF text structure can be compared without dumping large private image streams.
- [x] Extend PDF text-operation inspection with effective text matrices:
  `PdfInspect` now composes text matrices with the active graphics-state CTM, and
  `ComparePdfTextOperations.ps1 -UseEffectiveMatrix` can compare Office/candidate text placement after
  `cm`-driven rotation or translation. This is required for vertical PPTX parity because Office often emits
  rotation directly in `Tm`, while `ooxpdf` commonly rotates the frame with `cm` and emits upright local text.
- [x] Decode inspected PDF text-operation payloads through `/ToUnicode` maps:
  `PdfInspect` now emits `DecodedText` beside the raw `Tj`/`TJ` payload, so Office literal strings and
  candidate embedded-font hex strings can be compared structurally before changing glyph grouping or
  vertical layout.
- [x] Add an optional decoded-text gate to PDF text-operation comparison:
  visual manifests can set `expected.compareDecodedTextOperations` to require matching decoded text content
  in addition to matrix, position, font-size, and tracking tolerances.
- [x] Add opt-in decoded text summaries to PDF text-line comparison:
  `ComparePdfTextLineStarts.ps1 -ShowText` now prints per-line reference/candidate decoded operation text
  after the coordinate table. This keeps default output compact while making operation grouping mismatches
  diagnosable without opening `text-operations.json` by hand.
- [ ] Classify typography visual cases as `approximate`, `needs-review`, or `locked`. Only `locked` cases
  should enforce near-pixel-perfect thresholds; approximate gates should not mask text readability bugs.
- [x] Lock the first exact typography cases with PDF text-operation gates:
  `pptx-ladder-04-all-caps`, `pptx-ladder-04-field-text`,
  `pptx-ladder-04-nonbreaking-space`, `pptx-ladder-04-soft-hyphen`, and
  `pptx-ladder-04-tab-character`.
- [x] Lock additional simple typography cases with text-operation gates:
  `pptx-ladder-04-bold-face-single`, `pptx-ladder-04-italic-face-single`,
  `pptx-ladder-04-underline-single`, `pptx-ladder-04-strikethrough-single`, and
  `pptx-ladder-04-line-spacing-points`.
- [x] Add decoded-text gates to simple locked typography cases:
  `pptx-ladder-04-all-caps`, `pptx-ladder-04-bold-face-single`,
  `pptx-ladder-04-italic-face-single`, `pptx-ladder-04-underline-single`, and
  `pptx-ladder-04-strikethrough-single` now require matching decoded PDF text content as well as position
  and font-size tolerances.
- [x] Extend decoded-text gates to special text-operation cases that remain structurally locked:
  `pptx-ladder-04-field-text`, `pptx-ladder-04-soft-hyphen`, `pptx-ladder-04-tab-character`,
  and `pptx-ladder-04-nonbreaking-space`.
- [x] Extend decoded-text gates to the remaining simple locked text-operation cases:
  `pptx-ladder-04-line-spacing-points` and `pptx-ladder-04-mixed-font-size-line` now require decoded
  Office/candidate text parity in addition to their existing position and font-size gates.
- [x] Extend decoded-text gates to locked text-operation probe cases:
  `pptx-ladder-04-typography-capital-spacing-probe` and
  `pptx-ladder-04-typography-boundary-invariance-probe` now lock decoded PDF text content.
- [x] Revisit `pptx-ladder-04-nonbreaking-space` before adding decoded-text gating:
  the stale failure was a real structural split-flow issue. Hidden NBSP advances now preserve font kerning
  and tracking from the preceding logical glyph, restoring the visual gate to MAE `0.003519` and locking
  decoded `A`/`B` PDF text operations with the remaining font-table X delta bounded at `0.1 pt`.
- [x] Tighten near-miss simple typography cases before locking:
  `pptx-ladder-04-highlight-single` now passes its tight visual gate after the display-size baseline fix,
  and `pptx-ladder-04-mixed-font-size-line` now has a PDF text-operation gate with the remaining second-run
  x delta bounded at `0.07 pt`.
- [x] Introduce the first behavior-neutral intermediate text layout object: `TextLayoutLine` owns positioned
  runs and line end advance before alignment/PDF emission.
- [x] Extend the intermediate typography model from line ownership to explicit paragraph/run style records
  in the direct PPTX renderer. `ResolvedParagraphTextStyle` and `ResolvedRunTextStyle` now gather cascade
  decisions before line layout and PDF emission.
- [x] Promote PPTX direct-renderer text input to a first-class text-frame model:
  `PptxTextFrameModel`, `PptxTextParagraphModel`, and `PptxTextRunModel` now preserve raw OOXML while
  carrying resolved style, inherited frame geometry, clipping, and autofit inputs into layout.
- [x] Promote PPTX direct-renderer text layout to a first-class layout model:
  `PptxTextLayoutModel`, frame layouts, paragraph layouts, line layouts, and span layouts now exist before
  flattening to PDF text runs. Inspection tests lock model and layout ownership separately from emission.
- [x] Add first-class PPTX text line boxes:
  each line layout now owns top Y, baseline Y, baseline offset, line advance, max font size, and line-spacing
  kind before spans are flattened to PDF `TextRun`s. This gives baseline, highlight, wrapping, and line-height
  fixes a model home instead of another PDF-emission tweak.
- [x] Add first-class text layout atoms under positioned spans:
  each span now exposes word, space, tab, and hidden-advance atom kinds through the layout inspector, so
  wrapping and justification can reason about structured text instead of opaque strings.
- [x] Continue the intermediate typography model from resolved style records to explicit text body,
  paragraph, line, positioned run, glyph span, and hidden-control advance records. The direct PPTX renderer
  now has a first-class `PptxTextFlowModel` between resolved frame models and measured line layout:
  flow frames own text boxes, flow paragraphs own resolved paragraph styles, and flow runs own tabs,
  hidden advances, boundary punctuation, breaks, and caps scaling before measuring or PDF emission.
- [ ] Reverse-engineer the remaining PPTX text formulas from Office/PDF evidence and `pptx-renderer`
  semantics, replacing local constants with named general rules for hidden advances, baseline offsets,
  line advances, highlight rectangles, and text-operation boundaries.
- [x] Replace the first hidden-advance approximation with a general font-metric rule:
  non-drawn NBSP and narrow NBSP flow segments now advance by measuring their Unicode space glyphs in the
  resolved run font, instead of using a font-size multiplier for U+202F.
- [x] Preserve logical glyph-boundary metrics across hidden text-flow controls:
  split flow segments now carry kerning/tracking from the previous logical code point into the next advance,
  so hidden NBSP does not fall back to an isolated space width while remaining absent from decoded PDF text.
  The broader `pptx-ladder-04-typography-nbsp-narrow-space` probe was rebaselined to MAE `0.258598`,
  changed16 `0.003142`; its PDF line-start gate still passes with starts within `0.5 pt`, so the remaining
  raster drift is tracked as glyph-shape/font-table parity rather than a hidden-control placement failure.
- [x] Replace the default-tab approximation with the OOXML default tab interval:
  tabs without explicit stops now advance to the next 914400 EMU stop from the paragraph text origin,
  matching the `pptx-renderer` rule instead of scaling by font size.
- [x] Port the `pptx-renderer` formula for `normAutofit @lnSpcReduction`:
  resolved paragraph line spacing now scales explicit `spcPct`/`spcPts` by
  `1 - lnSpcReduction / 100000` while preserving default line spacing behavior.
- [x] Centralize remaining PPTX text constants behind named metric rules:
  baseline, default line-height fallback, manual-break offsets, small caps, superscript/subscript scaling,
  fallback advances, synthetic bold, highlight, strike, and coalescing tolerances now have one explicit
  formula owner instead of scattered literals in flow, layout, and drawing code.
- [x] Remove the hidden text-flow hard-coded advance escape hatch:
  `PptxTextFlowSegment` no longer carries an arbitrary font-size factor. Hidden advances now have to be
  measurable text in the resolved font, or they need a separately named formula in the metric-rule layer.
- [x] Port the `pptx-renderer`/CSS superscript and subscript scale:
  baseline-shifted runs now use `0.65x` font size instead of the previous `2/3` approximation, with a
  synthetic PPTX unit lock.
- [ ] Continue replacing text constants with formula-owned measurements, starting with baseline/line-box
  offsets and highlight/strike geometry, and lock each rule with Office-PDF text-operation or rectangle
  probes before broad visual MAE gates.
- [ ] Replace the current `TextRun`-backed layout spans with glyph-position spans that own decoded Unicode,
  font resource, glyph ids, glyph advances, kerning adjustments, and hidden-control advances before PDF
  `TJ` array construction.
- [x] Introduce the first explicit glyph-run emission object:
  `TextGlyphRun` now owns glyph hex, `TJ` positioning array, baseline, line width, and synthetic style
  flags before PDF text operators are written. This is behavior-neutral but separates emission from layout.
- [x] Extend `TextGlyphRun` with glyph atoms: decoded code point, glyph id, glyph advance, and adjustment
  before each glyph are now observable before PDF text operators are written.
- [x] Move the first glyph-positioning slice upstream into PPTX text layout:
  `PptxTextSpanLayout` now owns a `PptxTextGlyphSpanLayout` with code points, glyph ids, advances,
  kerning/tracking adjustments, natural width, and layout width before PDF emission. The justified text
  model test locks parity between the layout-owned glyph span and downstream glyph-run inspection.
- [x] Preserve hidden/control boundary adjustments in layout-owned glyph spans:
  visible glyph spans now expose a `LeadingAdjustment` carrying pending kerning/tracking adjustments from
  preceding hidden advances such as no-break spaces. PDF output remains behavior-compatible, but the
  intermediate model no longer loses the structural adjustment needed before future `TJ` construction work.
- [x] Expose resolved typeface at the layout glyph level:
  `PptxTextGlyphLayout` now carries the resolved typeface per glyph, not only on the containing span. This is
  still behavior-neutral, but it removes a one-font-per-span assumption from the inspection model and prepares
  the next accented-Latin/font-fallback slice where a visible run may need multiple Office-like font groups.
- [x] Preserve current font ownership through glyph-run inspection:
  downstream `TextGlyphRun` atoms and `PptxTextGlyphRunSnapshot` now carry each glyph's current resolved
  typeface plus the PDF font resource that emits it. This is behavior-neutral and intentionally still reports
  one resource for each current glyph run; the point is to make the future Office-like font fallback split
  observable at the PDF-emission boundary instead of hiding it behind raster drift.
- [x] Add the first structural per-glyph font fallback path:
  PPTX text measurement and glyph layout now resolve missing code points against discovered non-math text
  fonts, keep primary requested typefaces for exact hits, and split positioned spans by glyph typeface before
  PDF emission so fallback glyph ids are encoded against their own font resource. A synthetic Arial/CJK test
  locks decoded glyph order plus multiple PDF font resources.
- [x] Repair and structurally lock the accented-Latin public probe:
  the checked-in `pptx-ladder-04-typography-accent-spacing-probe.pptx` had mojibake in `slide1.xml` even
  though the generator source was correct UTF-8. Regenerating the fixture restored true accented text; the
  case now gates decoded text and line starts while leaving raster metrics approximate.
- [x] Make PPTX shape-text emission consume layout-owned glyph spans directly:
  `PptxPositionedTextSpan` now carries the legacy `TextRun`, line box, atoms, and glyph span through
  flattening. Shape-text emission and glyph-run inspection build `TextGlyphRun` from the carried glyph span
  instead of remapping glyphs from `TextRun`; table text remains on the legacy bridge for a later slice.
- [x] Preserve glyph metadata through the first coalescing bridge:
  positioned-span coalescing now keeps atoms and line-box identity, and rebuilds merged glyph spans when
  same-style spans are coalesced so `TJ` construction still has pre-emission glyph data.
- [x] Preserve measured inter-span glyph gaps during coalescing:
  same-style positioned spans now merge by carrying the original glyph layouts plus the measured gap between
  spans, instead of recomputing the merged `TJ` array from the concatenated text. This keeps the intermediate
  typography model authoritative and avoids another source of subtle inter-letter drift.
- [x] Route highlighted PPTX text through layout-owned glyph spans:
  highlighted runs no longer fall back to legacy `TextRun` glyph remapping after the highlight background is
  drawn. This keeps text with yellow fills on the same emission path as ordinary shape text.
- [x] Route PPTX table-cell text through positioned glyph spans:
  table text now reuses the same text-frame and glyph-span layout path as shape text before PDF emission,
  matching the `pptx-renderer` pattern where tables delegate cell text to the common text-body renderer.
- [ ] Move highlight, underline, and strike geometry to layout-owned line boxes and glyph spans:
  highlight drawing still receives legacy `TextRun`s today, while underline/strike now use glyph-run width
  from positioned spans for PPTX shape text.
- [x] Re-run the justified typography probe after glyph-span emission:
  `pptx-ladder-04-typography-justify-port` stayed behavior-compatible at MAE `4.210743`; remaining drift is
  Office word-position and line-box parity, not a glyph-span bridge regression.
- [x] Add a model-level justified-word diagnostic:
  the justified text layout test now asserts monotonic word starts and exposes which spans absorb distributed
  spacing through `GlyphSpan.LayoutWidth > NaturalWidth`, so Office text-op x drift can be compared against
  layout data rather than inferred from raster output.
- [x] Survey the current typography visual ladder after positioned glyph-span emission:
  `justify-port` MAE `4.210743`, `highlight-single` MAE `0.574984` and still failing its locked gate,
  `boundary-invariance` MAE `2.835927` with a text-op gate failure, `inventory-opti` MAE `1.148489`, and
  `accent-spacing` MAE `2.149007`. The failures point to baseline/line-box parity and Office word-position
  strategy before broader private-deck tuning.
- [ ] Keep highlighted PPTX text on the legacy emission path until highlight geometry is ported cleanly:
  an attempted line-box/glyph-span highlight migration regressed the locked `highlight-single` visual gate,
  so highlight needs an Office-PDF text and rectangle geometry pass before switching.
- [ ] Port `pptx-renderer`'s text-cascade shape more explicitly: a seven-level paragraph cascade
  (`defaultTextStyle`, master text style, master placeholder, layout placeholder, shape `lstStyle`,
  paragraph `pPr`, run `rPr`) should produce resolved paragraph/run style records before layout.
- [x] Make paragraph cascade inputs explicit in the direct PPTX text model:
  `PptxParagraphStyleCascade` now records the paragraph level and source layers before merged paragraph
  defaults and resolved styles are produced.
- [x] Name the paragraph cascade layers in the direct text model:
  `shape.lstStyle`, inherited placeholder list styles, inherited text styles, and `defaultTextStyle` are now
  observable before merge. This mirrors the useful `pptx-renderer` pattern of making cascade layers testable
  instead of keeping anonymous XML fallback lists.
- [ ] Extend the cascade model from paragraph defaults to a full named seven-level resolver with separate
  paragraph, run, bodyPr, placeholder geometry, and theme font/color fallback stages.
- [ ] Port `pptx-renderer` text edge-case tests as .NET unit/visual cases for hyperlink color, shape
  `fontRef` color precedence, table text overrides, gradient/no-fill/outline text, `kern` thresholds,
  tabs, repeated spaces, absolute line-height wrappers, `normAutofit`, and end-paragraph run sizing.
- [x] Next PPTX typography sequence: add diagnostics for Office-vs-candidate word starts per visual line,
  then use those diagnostics to tighten justified text spacing without relying on late-game MAE.
  `tools/ComparePdfTextLineStarts.ps1` now groups inspected PDF text operations by visual line and compares
  text-operation start positions. `CheckVisualCase.ps1` can enforce this with `maxTextLineStartDelta`; the
  boundary-invariance probe locks line starts at `0.1pt` while retaining the existing text-op gate.
- [x] Next PPTX typography sequence: move highlight rectangles from the legacy `TextRun` drawing path to
  layout-owned line boxes and glyph spans, keeping `highlight-single` locked throughout the migration.
  Shape-text highlights now consume positioned spans and line-box baselines while highlighted text emission
  remains on the legacy text path. `highlight-single` stays locked at MAE `0.032134`, and the boundary
  invariance probe passes its text-operation gate with MAE `0.921847`.
- [x] Next PPTX typography sequence: add and lock repeated-space, tab/space, non-breaking-space, and
  narrow-space interaction cases because these are direct bottom-up probes for phantom inter-letter gaps.
  Split public cases now exist for repeated spaces, NBSP/narrow spaces, punctuation boundaries, and
  tab-plus-space handling. Repeated spaces, tab-plus-space, NBSP/narrow-space, and punctuation boundaries
  are locked on raster metrics; all except the broad mixed probe also use text line-start gates. NBSP now
  advances as a hidden regular space, narrow NBSP as a narrow hidden advance, and the broad whitespace probe
  improved to MAE `0.583445`, changed16 `0.007705`.
- [x] Add the first whitespace-control interaction probe:
  `pptx-ladder-04-typography-whitespace-controls-probe` covers repeated spaces, NBSP, narrow NBSP,
  punctuation-adjacent words, tab plus spaces, and accented Latin with spacing controls. `U+202F` narrow
  NBSP is now modeled as a hidden advance instead of a visible glyph, with a unit test locking the split
  word spans. After the NBSP/narrow-space tightening, the probe passes approximate gates at MAE `0.583445`,
  changed16 `0.007705`; Office still emits more granular accent text operations than the candidate.
- [x] Next PPTX typography sequence: attack accented Latin and punctuation-adjacent word cases after the
  spacing-control cases, using Office PDF text operations to separate fallback-font splits from positioning.
  Punctuation-adjacent hyphen boundaries are now represented as separate layout spans and locked by
  `pptx-ladder-04-typography-punctuation-boundaries` with line-start parity within `0.25pt`. The accented
  Latin probe was regenerated from the correct UTF-8 generator source and now locks decoded text plus
  line-start parity; remaining drift is glyph raster/font parity, not a fallback/positioning blocker.
- [ ] Next PPTX typography sequence: continue porting `pptx-renderer` typography tests for end-paragraph run
  sizing, superscript/subscript baseline, no-fill/outline text, hyperlink color, and table text overrides.
- [x] Port the `pptx-renderer` justified-alignment text oracle as a public ooxpdf probe:
  `pptx-ladder-04-typography-justify-port` now exposes the missing PowerPoint word-spacing behavior for
  wrapped justified paragraphs.
- [x] Implement the first PowerPoint-like justified paragraph layout slice: fully wrapped lines distribute
  extra width across word spaces, final visual lines and manual-break lines remain left aligned, and
  justified layout spans are protected from text-run coalescing so the expanded positions reach PDF emission.
- [ ] Tighten justified paragraph parity against Office: current public probe still has large raster drift,
  so inspect Office text operations and refine baseline, line advance, and per-word spacing strategy.
- [x] Move justified PPTX spacing ownership from glyph width to positioned word starts:
  justified lines now split drawable word spans from stretchable spaces before PDF emission, so individual
  word glyph spans keep natural widths while line layout owns distributed x positions. The public
  `justify-port` raster metric stayed at MAE `4.210743`; Office's inspected text-operation count is still
  higher because Office chunks words inside `TJ` arrays, so the remaining work is word-position formula and
  PDF text-array parity rather than layout ownership.
- [x] Inspect Office vs candidate text operations for `pptx-ladder-04-typography-justify-port`: word-level
  positions now reach candidate PDF emission, but baselines remain about 0.7 pt low and raster drift remains
  high, so the next pass should target baseline/line-advance parity before private-deck tuning.
- [x] Reject font-name-specific baseline constants after comparing `pptx-renderer`: its text layout uses
  OOXML line-height semantics and browser/font metrics rather than Calibri/Aptos special cases. `ooxpdf`
  baseline work must stay metric-driven and generic across all resolved fonts.
- [x] Reverse-engineer `pptx-renderer`'s baseline and line-height strategy for the next typography slice:
  `spcPct` becomes unitless CSS line-height, `spcPts` becomes absolute point line-height, manual line breaks
  under absolute spacing use block wrappers with explicit height, and normal text baselines are delegated to
  browser/font layout. Only OOXML run `baseline` is an explicit superscript/subscript shift. `ooxpdf` should
  therefore expose deterministic font metrics and line-box diagnostics, not hard-code family-specific
  offsets.
- [x] Replace the rejected font-name baseline tweak with a generic OpenType-metric rule:
  first-line PPTX baselines now use the resolved font's ascender metrics when available, falling back to the
  previous heuristic only when font resolution fails. The justified text probe improved from MAE ~`4.63` to
  `4.210743`, with the first-line baseline delta reduced from about `0.71 pt` to about `0.32 pt`.
- [ ] Align `ooxpdf` line-height handling with that strategy: keep percent line spacing as a multiplier,
  keep point line spacing as absolute line box height, model manual-break line boxes explicitly, and compare
  line top, baseline, and highlight rectangle geometry against Office PDF text operations before changing
  broader deck rendering.
- [x] Port the `pptx-renderer`/CSS rule for explicit percent line spacing:
  `spcPct=100000` now resolves to one font size, not `1.2x` font size. Existing paragraph/list-style tests
  were updated to lock the new text baselines, `pptx-ladder-04-line-spacing-points` remains locked at MAE
  `0.014381`, and `pptx-ladder-04-line-spacing-port` still passes its approximate gate with line-to-line
  deltas matching Office while retaining the known generic baseline offset.
- [x] Match Office's first-paragraph spacing behavior for PPTX text frames:
  the first laid-out paragraph no longer applies `spcBef`, while later paragraphs still include
  `spcAft + spcBef`. This improved `pptx-ladder-04-line-spacing-port` from MAE `4.347075` to `1.443223`
  and brought all inspected text baselines within about `0.32 pt` of Office, so the case now has a PDF
  text-operation position gate.
- [x] Port the first `pptx-renderer` manual-break line-spacing lock:
  a synthetic layout inspection test now proves `spcPts` line spacing stays as an absolute line-box advance
  across `<a:br/>`, matching the renderer pattern that uses explicit line wrappers for absolute spacing.
- [ ] Continue the same metric-driven track without font-family exceptions: inspect Office vs candidate line
  boxes for `pptx-ladder-04-typography-justify-port` and decide whether PowerPoint is using adjusted
  ascender, internal leading, or another generic font metric for top-to-baseline placement.
- [x] Refine the generic baseline metric rule without font-family exceptions:
  small/body text keeps the resolved font ascender ratio, while display-size text uses a `0.974` lower
  bound before applying larger font ascenders. This restores the locked Arial all-caps baseline
  (`pptx-ladder-04-all-caps` MAE `0.003613`) and brings `pptx-ladder-04-highlight-single` under its tight
  gate (MAE `0.032134`) while preserving the Calibri line-spacing text-operation gate.
- [x] Add the first `pptx-renderer`-style cascade lock for shape `fontRef` color precedence:
  explicit run `solidFill` now has a focused unit test proving it overrides the shape-level fallback color.
- [x] Port `pptx-renderer`'s `a:kern` threshold behavior into PPTX text layout and PDF emission:
  kerning is disabled when the run font size is below the OOXML threshold, and a unit test locks the
  no-adjustment `TJ` array for Times New Roman `To`.
- [x] Port `pptx-renderer`'s basic `normAutofit fontScale` handling into PPTX text-frame layout:
  run and paragraph font sizes are scaled before layout/PDF emission, the unsupported autofit diagnostic
  was removed for this supported subset, and a unit test locks `30pt * 80% = 24pt`.
- [x] Add explicit public probes for the private-deck spacing failures: capital-letter kerning,
  accented-letter pairs, Cambria/Calibri/Arial differences, highlighted runs, and mixed run boundaries.
  The first probes are `pptx-ladder-04-typography-capital-spacing-probe`,
  `pptx-ladder-04-typography-accent-spacing-probe`, and
  `pptx-ladder-04-typography-boundary-invariance-probe`.
- [x] Add public probes for the private slide 2/3 typography symptoms:
  `pptx-ladder-04-typography-inventory-opti-probe` targets the `Inventory Optimization` phantom-spacing
  failure, and `pptx-ladder-04-typography-dense-column-probe` targets small dense-column legibility with
  accented letters and capital-letter pairs.
- [x] Split highlighted-run decoration from text emission for PPTX text drawing. This matches Office's
  pattern where highlight rectangles are painted separately while unchanged font metrics can still emit
  a single continuous text operation across run boundaries.
- [x] Tighten `pptx-ladder-04-typography-accent-spacing-probe` structurally:
  the apparent Office/candidate operation split was caused by a mojibake fixture, not true accented Latin.
  The corrected fixture now emits one decoded text operation per line on both sides and passes line-start
  parity at `0.1pt`; raster metrics remain approximate for later font-rendering parity work.
- [x] Align justified paragraph punctuation/final-line operation boundaries with Office:
  sentence periods now split as Office text-operation boundary punctuation, and non-expanded final lines in
  justified paragraphs are still split into word/punctuation glyph spans for PDF emission. The justify probe
  improved structurally: lines 0-4 now have matching operation text/counts, while paragraph 3 still wraps
  `like` one line later than Office and remains the next line-break metric target. Validation: typography
  `66 passed, 0 failed, 2 skipped`; public justify visual `20260524-214947` passed with MAE `4.187326`
  and PDF text-operation parity for the first five lines; skip-slow/full suites and packing passed.
- [x] Expose natural pre-expansion text-line width in PPTX layout inspection:
  justified line snapshots now keep both `EndX` and `NaturalEndX`, and the private-safe layout diagnostic
  reports natural width plus alignment expansion. This confirms the remaining public justify paragraph-3
  wrap mismatch is a narrow metric/tolerance gap: adding `like` would exceed the frame by about `1.1pt`
  under current Calibri measurements, while Office keeps it on the previous justified line.
- [x] Re-run public spacing probes after the atom/glyph architecture slice:
  `inventory-opti`, `accent-spacing`, and `boundary-invariance` still render without diagnostics; boundary
  invariance remains structurally close while inventory/accent remain acceptance probes for later glyph
  parity work.
- [ ] Introduce a PPTX render context in `ooxpdf` analogous to `pptx-renderer`: slide model, layout/master
  model, theme, relationships, media lookup, diagnostics sink, font/color caches, and group context.
- [x] Introduce the first behavior-neutral PPTX render-context boundary for package, document, theme,
  slide XML, inherited XML, slide identity, and diagnostics; extend it toward relationships/media/cache
  ownership in later rendering splits.
- [x] Split slide inheritance ownership in `PptxRenderContext`:
  master and layout XML are now carried as a `PptxSlideInheritance` object while preserving the legacy
  ordered inherited-source list. This is a bridge toward `pptx-renderer`-style typed slide/layout/master
  context without forcing a full renderer rewrite.
- [x] Route PPTX text inspection/model construction through the render-context loader:
  `InspectTextRuns`, `InspectTextLayout`, and `InspectTextFrameModels` now share slide/theme/inheritance
  context setup with production rendering instead of rebuilding loose XML/document/theme tuples. This keeps
  model-stage tests aligned with the renderer boundary.
- [x] Route background and shape traversal through the PPTX render context and move that dispatch into a
  dedicated renderer partial, keeping the shape drawing primitive behavior-neutral for later typed-node work.
- [x] Move picture traversal, image decoding, crop/fill, and alpha handling into a dedicated renderer partial
  that consumes slide relationships from the PPTX render context.
- [x] Route table frame rendering through the PPTX render context so table layout/fill/border/text code no
  longer takes document/theme as separate ad hoc dependencies.
- [x] Add context-based text-run readers for inherited, slide, and ordered-shape text so typography rendering
  no longer spreads document/theme/slide-number arguments through top-level PPTX dispatch.
- [x] Move ordered slide z-order dispatch into a dedicated partial, keeping top-level `RenderPages` focused on
  orchestration and leaving node dispatch in one context-driven place.
- [x] Cache the presentation theme once per PPTX render pass instead of reloading it per slide, matching the
  context-lifetime model and reducing repeated parsing for large decks.
- [x] Move PPTX shape/preset rendering, arrowheads, dash/cap/join handling, picture-fill clipping, group
  transforms, and shape theme fill/line lookup into the shape renderer partial.
- [x] Add private-safe PPTX scene inspection snapshots:
  `PptxRenderer.InspectScene` now exposes slide/master/layout node counts, node kinds, transform flags,
  text-body paragraph/run counts, picture/table/chart ownership, and group children without exposing text
  content. The private-safe layout diagnostic now includes these scene counts so future slide-level schema
  work can distinguish typed-model coverage gaps from text-layout/PDF-emission gaps.
- [x] Split ordered scene shape dispatch into shape and text node phases:
  ordered rendering now draws shape geometry through `RenderShapeNode` and emits text through `RenderTextNode`,
  which first checks the typed scene node's `TextBody` ownership before invoking the existing text layout
  pipeline. This is behavior-neutral, but it makes text a distinct scene-node rendering phase instead of an
  inline raw-XML side effect of shape drawing.
- [x] Add a typed unknown-graphic-frame fallback:
  unsupported non-chart, non-table, non-SmartArt `graphicFrame` nodes now emit
  `PPTX_UNSUPPORTED_GRAPHIC_FRAME` instead of being indistinguishable from intentionally ignored content.
  Ordered scene rendering also has an explicit `UnknownGraphicFrame` branch for inherited scene nodes.
- [x] Keep ordered scene rendering active when unknown graphic frames are present:
  slide-level unknown `graphicFrame` nodes no longer force the legacy fallback path, so known sibling
  shapes/text/tables/charts still follow scene order while the unknown node is ignored with diagnostics.
  `PptxSyntheticUnknownGraphicFrameDoesNotDisableSiblingOrder` locks the text-before-covering-shape order
  across an intervening unknown frame and avoids duplicate slide-level diagnostics from ordered dispatch.
  The old ordered-render eligibility branch was removed from `RenderPages`; all slides now flow through the
  typed ordered scene dispatcher.
- [x] Remove the obsolete PPTX fallback traversal entry points:
  the private XML-wide `RenderPictures`, `RenderShapes`, and `RenderCharts` passes, plus their container
  walkers, were deleted after all slides moved to typed ordered scene dispatch. The grouped-table unknown
  frame test was renamed to `PptxSyntheticGroupedTableUsesGroupTransformWithUnknownGraphicFrame` so the test
  catalog no longer describes a fallback renderer path that no longer exists.
- [x] Move PPTX table font prepass onto typed scene nodes:
  ordered rendering no longer calls the raw XML `RenderTables(..., new PdfGraphicsBuilder())` traversal just
  to discover table text fonts. `ReadSceneTableTextSpans` now walks master, layout, and slide scene nodes with
  the same group transforms as ordered dispatch, and obsolete XML table-frame overloads were removed.
- [x] Move PPTX shape text font prepass onto typed scene nodes:
  ordered rendering now gathers shape text fonts from master, layout, and slide `PptxSceneNode` lists with the
  same placeholder rules as ordered dispatch, instead of running whole-part inherited/slide XML text scans
  before rendering.
- [x] 2026-05-24: Re-ran private PPTX acceptance after adding the generic unknown graphic-frame diagnostic.
  Private run `artifacts/private-visual/lokad-value-based/20260524-221657` compared 84/84 pages with zero
  dimension mismatches, deck MAE `9.005915`, changed16 `0.116052`, and still only
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`; the new `PPTX_UNSUPPORTED_GRAPHIC_FRAME` diagnostic did not appear in
  this private deck. Slide 17 measured MAE `2.880739`, changed16 `0.044888`, SSIM `0.920083`.
- [ ] Split PPTX rendering dispatch by typed scene node: background, shape, text, picture, table, chart,
  group, and unknown/diagnostic fallback should be separate renderers consuming the same context.
- [x] Move master/layout shape/text rendering into the ordered scene pipeline: non-placeholder master and
  layout scene nodes now render through `RenderOrderedSceneNodes` before slide nodes instead of the old XML
  shape-container pass.
- [x] Finish relationship-aware inherited scene rendering for master/layout pictures and charts: scene slides
  now retain master/layout part names, and ordered dispatch passes the source part's relationship map to
  shapes, pictures, and chart frames instead of reusing slide relationships.
- [x] 2026-05-24: Re-ran package and private PPTX acceptance after source-part relationship dispatch for
  inherited scene nodes. `dotnet pack` succeeded and private run
  `artifacts/private-visual/lokad-value-based/20260524-115144` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, only `PPTX_UNSUPPORTED_IMAGE_RECOLOR`,
  and slide 17 MAE `2.945717`, changed16 `0.045530`, SSIM `0.917662`.
- [x] Move solid slide/master/layout background fill into `PptxSceneSlide`: background color and alpha are now
  parsed once in the scene model and rendering consumes typed background fields before falling back to raw XML.
- [x] Move ordered-render eligibility onto typed scene nodes: dispatch now checks `PptxSceneNodeKind` for
  unknown graphic frames instead of rescanning slide XML and re-running node classification.
- [x] 2026-05-24: Re-ran the full test suite, package, and private PPTX acceptance after moving
  ordered-render eligibility onto scene nodes. The test suite passed 183/183, `dotnet pack` succeeded, and
  private run `artifacts/private-visual/lokad-value-based/20260524-120726` stayed stable: 84/84 compared
  pages, zero dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, only
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`, and slide 17 MAE `2.945717`, changed16 `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Re-ran the image-focused tests, full test suite, package, and private PPTX acceptance after
  centralizing picture crop/fill/alpha/recolor parsing on the scene builder. The test runner passed 184/184,
  `dotnet pack` succeeded, and private run `artifacts/private-visual/lokad-value-based/20260524-121521`
  stayed stable: 84/84 compared pages, zero dimension mismatches, deck MAE `9.043369`, changed16
  `0.116418`, and only one `PPTX_UNSUPPORTED_IMAGE_RECOLOR`.
- [x] 2026-05-24: Re-ran model/composition-focused tests, the full test suite, package, and private PPTX
  acceptance after moving group transforms into `PptxSceneNode`. The test runner passed 184/184,
  `dotnet pack` succeeded, and private run `artifacts/private-visual/lokad-value-based/20260524-121950`
  stayed stable: 84/84 compared pages, zero dimension mismatches, deck MAE `9.043369`, changed16
  `0.116418`, and only one `PPTX_UNSUPPORTED_IMAGE_RECOLOR`.
- [x] 2026-05-24: Re-ran chart/model-focused tests, the full test suite, package, and private PPTX acceptance
  after adding `PptxSceneChart` relationship ids. The test runner passed 184/184, `dotnet pack` succeeded,
  and private run `artifacts/private-visual/lokad-value-based/20260524-122341` stayed stable: 84/84
  compared pages, zero dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`.
- [x] 2026-05-24: Re-ran composition/table-focused tests, the full test suite, package, and private PPTX
  acceptance after routing ordered table frame bounds through scene nodes plus group transforms. The test
  runner passed 185/185, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-122826` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`.
- [x] 2026-05-24: Re-ran composition/table-focused tests, the full test suite, package, and private PPTX
  acceptance after carrying group transforms through the legacy table fallback path. The test runner passed
  186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-123212` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`.
- [x] 2026-05-24: Re-ran model/table-focused tests, the full test suite, package, and private PPTX acceptance
  after moving table grid widths and row heights into `PptxSceneTable`. The test runner passed 186/186,
  `dotnet pack` succeeded, and private run `artifacts/private-visual/lokad-value-based/20260524-130135`
  stayed stable: 84/84 compared pages, zero dimension mismatches, deck MAE `9.043369`, changed16
  `0.116418`, and only one `PPTX_UNSUPPORTED_IMAGE_RECOLOR`.
- [x] 2026-05-24: Re-ran model/table-focused tests, the full test suite, package, and private PPTX acceptance
  after moving table cell span/continuation metadata into `PptxSceneTable`. The test runner passed 186/186,
  `dotnet pack` succeeded, and private run `artifacts/private-visual/lokad-value-based/20260524-130519`
  stayed stable: 84/84 compared pages, zero dimension mismatches, deck MAE `9.043369`, changed16
  `0.116418`, and only one `PPTX_UNSUPPORTED_IMAGE_RECOLOR`.
- [x] 2026-05-24: Re-ran model/table-focused tests, the full test suite, package, and private PPTX acceptance
  after moving table-cell text margins and vertical anchors into `PptxSceneTable`. The test runner passed
  186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-131039` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`.
- [x] 2026-05-24: Re-ran model/table-focused tests, the full test suite, package, and private PPTX acceptance
  after moving direct table-cell solid fills into `PptxSceneTableCell`. Focused model/table tests passed, the
  full runner passed 186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-131555` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Re-ran model/table-focused tests, the full test suite, package, and private PPTX acceptance
  after moving explicit table-cell borders into `PptxSceneTableCell`. Focused tests passed, the full runner
  passed 186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-132138` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Re-ran model/table-focused tests, the full test suite, package, and private PPTX acceptance
  after moving built-in table-style descriptors into `PptxSceneTable`. Focused tests passed, the full runner
  passed 186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-132803` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Re-ran model/table-focused tests, the full test suite, package, and private PPTX acceptance
  after materializing per-cell built-in table-style fills and text defaults in `PptxSceneTableCell`. Focused
  tests passed, the full runner passed 186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-133517` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Re-ran model/table-focused tests, the full test suite, package, and private PPTX acceptance
  after moving raw table-cell `txBody` ownership into `PptxSceneTableCell`. Focused tests passed, the full
  runner passed 186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-133919` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Re-ran model/table-focused tests, the full test suite, package, and private PPTX acceptance
  after moving the retained raw table element into `PptxSceneTable.Source`. Focused tests passed, the full
  runner passed 186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-134331` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Re-ran model/chart-focused tests, the full test suite, package, and private PPTX acceptance
  after resolving chart relationship targets into `PptxSceneChart.TargetPartName`. Focused tests passed, the
  full runner passed 186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-134950` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Re-ran model/chart-focused tests, the full test suite, package, and private PPTX acceptance
  after moving chart XML and color-style palette ownership into `PptxSceneChart`. Focused tests passed, the
  full runner passed 186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-135452` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Re-ran model/chart-focused tests, the full test suite, package, and private PPTX acceptance
  after adding typed chart plot summaries to `PptxSceneChart`. Focused tests passed, the full runner passed
  186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-135835` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Re-ran model/chart-focused tests, the full test suite, package, and private PPTX acceptance
  after adding typed chart axis catalogs to `PptxSceneChart`. Focused tests passed, the full runner passed
  186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-140154` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Re-ran model/chart-focused tests, the full test suite, package, and private PPTX acceptance
  after adding typed explicit chart title and legend metadata to `PptxSceneChart`. Focused tests passed, the
  full runner passed 186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-140522` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Re-ran model/chart-focused tests, the full test suite, package, and private PPTX acceptance
  after adding typed chart series summaries to `PptxSceneChart.Plots`. Focused tests passed, the full runner
  passed 186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-143158` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Re-ran model/chart-focused tests, the full test suite, package, and private PPTX acceptance
  after preserving scatter/bubble series channels (`c:xVal`, `c:yVal`, and `c:bubbleSize`) in
  `PptxSceneChartSeries`. Focused tests passed after a transient parallel build file lock was rerun
  serially, the full runner passed 186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-143705` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Re-ran model/chart-focused tests, the full test suite, package, and private PPTX acceptance
  after adding typed chart plot attributes to `PptxSceneChartPlot`: grouping, bar direction, scatter style,
  vary-colors, gap width, overlap, and hole size. Focused tests passed after a transient parallel build file
  lock was rerun serially, the full runner passed 186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-144151` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Re-ran chart/model-focused tests, the full test suite, package, and private PPTX acceptance
  after wiring ordered chart rendering to prefer scene-owned plot attributes before falling back to raw plot
  XML. Focused tests passed, the full runner passed 186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-144548` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Re-ran chart/model-focused tests, the full test suite, package, and private PPTX acceptance
  after making ordered chart rendering prefer scene-owned cached series data for bar, line, area, radar, pie,
  doughnut, scatter, and bubble charts. Focused tests passed, the full runner passed 186/186, `dotnet pack`
  succeeded, and private run `artifacts/private-visual/lokad-value-based/20260524-144936` stayed stable:
  84/84 compared pages, zero dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Re-ran chart/model-focused tests, the full test suite, package, and private PPTX acceptance
  after making ordered bar and line chart category labels prefer scene-owned cached category labels before
  raw `c:cat` XML scans. Focused tests passed, the full runner passed 186/186, `dotnet pack` succeeded, and
  private run `artifacts/private-visual/lokad-value-based/20260524-145224` stayed stable: 84/84 compared
  pages, zero dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Re-ran chart/model-focused tests, the full test suite, package, and private PPTX acceptance
  after adding typed chart-series style metadata to `PptxSceneChartSeries`: direct solid fill, direct line,
  marker symbol/size/fill/line, and smooth flag. Focused tests passed after correcting the expected OOXML
  line-width conversion, the full runner passed 186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-145742` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Make ordered chart rendering consume scene-owned series styles first for series fills,
  strokes, markers, and smooth flags. Pattern fills are modeled explicitly in `PptxSceneChartSeries`, so
  scene-first style consumption does not narrow the existing XML style support. Focused model/chart tests
  passed after a transient parallel build file lock was rerun serially, the full runner passed 186/186,
  `dotnet pack` succeeded, and private run `artifacts/private-visual/lokad-value-based/20260524-150545`
  stayed stable: 84/84 compared pages, zero dimension mismatches, deck MAE `9.043369`, changed16
  `0.116418`, and only one `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE
  `2.945717`, changed16 `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Move chart per-point style overrides into `PptxSceneChartPointStyle`: `c:dPt` index,
  direct fill, pattern fill, line, and raw explosion values are now owned by the scene, while renderer
  helpers normalize percentages at the PDF boundary and retain XML fallback when scene point styles are
  absent. Focused model/chart tests passed after a transient parallel build file lock was rerun serially,
  the full runner passed 186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-151352` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Make chart title/legend layout and title rendering consume scene-owned metadata first.
  `PptxSceneChartTitle` now also captures `c:v` title text, the renderer uses scene title/legend records
  for supported chart layout and fallback title drawing, and the XML title fallback remains for the existing
  single-series bar automatic title rule. Focused model/chart tests passed, the full runner passed 186/186,
  `dotnet pack` succeeded, and private run `artifacts/private-visual/lokad-value-based/20260524-151957`
  stayed stable: 84/84 compared pages, zero dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`,
  and only one `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`,
  changed16 `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Extend `PptxSceneChartAxis` beyond the initial id/kind catalog to preserve value-axis
  scaling, major/minor units, and major/minor gridline visibility. Bar and line chart rendering now consumes
  those scene axis records first for value extents, tick/gridline units, and gridline visibility while
  preserving XML fallback for legacy/raw chart paths and richer label formatting. Focused model/chart tests
  passed, the full runner passed 186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-152728` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Extend scene-owned chart-axis metadata to tick-label position and number format.
  Bar/line category and value-axis label visibility, value-axis side placement, and value-axis tick label
  formatting now prefer `PptxSceneChartAxis` while retaining XML fallback for secondary/raw paths and
  unmodeled rich text/style state. Focused model/chart tests passed, the full runner passed 186/186,
  `dotnet pack` succeeded, and private run `artifacts/private-visual/lokad-value-based/20260524-153319`
  stayed stable: 84/84 compared pages, zero dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`,
  and only one `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`,
  changed16 `0.045530`, SSIM `0.917662`.
- [x] 2026-05-25: Preserve chart-axis boolean source presence for deletion and multi-level labels.
  `PptxSceneChartAxis` now keeps `c:delete` and `c:noMultiLvlLbl` as nullable source booleans, so a missing
  element remains distinct from an explicit `val="0"` disable. The chart renderer still coalesces only
  explicit deletion to hidden axes, preserving existing output while preventing the scene model from erasing
  OOXML schema intent before later Office-alignment work. Focused scene tests passed, the full runner passed
  204/204, and `dotnet pack` succeeded.
- [x] 2026-05-25: Preserve chart title and legend boolean source presence.
  `PptxSceneChartTitle.IsAutoDeleted` now keeps missing `c:autoTitleDeleted` distinct from explicit false,
  and `PptxSceneChartLegend` now separates legend element presence from nullable `c:delete` metadata. The
  renderer still treats only explicit title auto-delete and legend delete as hidden, preserving current output
  while retaining the OOXML state needed for later Office-aligned title/legend layout decisions. Focused
  scene tests passed, the full runner passed 204/204, and `dotnet pack` succeeded.
- [x] 2026-05-24: Move chart data-label value/percent flags into `PptxSceneChartPlot.DataLabels` and make
  bar, line, pie, and doughnut data-label rendering consume the typed scene options before raw XML fallback.
  This preserves the existing plot-level then first-series `c:dLbls` precedence while removing another
  renderer-local chart XML scan from the supported scene path. Focused model/chart tests passed, the full
  runner passed 186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-153953` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [ ] Extend the chart data-label scene model beyond the current metadata subset: rich text runs inside
  custom label text, leader lines, exact Office label-box geometry/auto-fit, and richer position semantics
  still need typed ownership before data-label layout can be aligned structurally with Office instead of
  renderer heuristics.
- [x] 2026-05-24: Make chart legend-entry name construction scene-first. Bar/combo and line legend entries
  now consume `PptxSceneChartPlot.Series[].Name` before falling back to raw `c:ser` XML, with the existing
  `Series N` default preserved for unnamed series. Focused model/chart tests passed after a transient
  parallel build lock was rerun serially, the full runner passed 186/186, `dotnet pack` succeeded, and
  private run `artifacts/private-visual/lokad-value-based/20260524-154412` stayed stable: 84/84 compared
  pages, zero dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Move chart area and plot area solid fill/line style ownership into `PptxSceneChart`.
  Supported chart renderers now consume scene-owned `ChartAreaStyle` and `PlotAreaStyle` before raw XML
  fallback, while the renderer boundary converts the typed scene styles into PDF chart drawing styles.
  Focused model/chart tests passed after a transient parallel build lock was rerun serially, the full runner
  passed 186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-154927` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Preserve explicit chart-area/plot-area `a:noFill` state in the scene model.
  `PptxSceneChartShapeStyle.NoFill` now distinguishes an Office-authored no-fill choice from an absent or
  unsupported fill, and the renderer boundary maps it to the existing no-background output without
  re-scanning chart `c:spPr`. Focused model/chart tests passed after a transient parallel build lock was
  rerun serially, the full runner passed 186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-170005` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Move chart-area/plot-area pattern fills into `PptxSceneChartShapeStyle`.
  Chart and plot shape styles now preserve `a:pattFill` preset, foreground, and background colors through
  the same typed `PptxScenePatternFill` contract used by chart series and points. The chart renderer consumes
  those pattern fills through the existing chart pattern-fill drawing path instead of treating chart/plot-area
  backgrounds as solid-only. Focused model/chart tests passed after a transient parallel build lock was
  rerun serially, the full runner passed 186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-170856` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043200`, changed16 `0.116408`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [ ] Extend chart area/plot area style modeling beyond direct solid fill/no-fill, pattern fill, and simple
  line: gradient fill, theme style references, transparency groups, and effect inheritance need typed
  ownership before chart background and plot-box styling can be considered structurally Office-aligned.
- [x] 2026-05-24: Move explicit chart-axis line style ownership into `PptxSceneChartAxis`. Supported bar
  and line chart rendering now consumes scene-owned value/category axis strokes first, including the existing
  explicit `a:ln/a:noFill` hidden-axis-line case represented as a transparent zero-width stroke, while raw
  XML axis style fallback remains for legacy paths. Focused model/chart tests passed after a transient
  parallel build lock was rerun serially, the full runner passed 186/186, `dotnet pack` succeeded, and
  private run `artifacts/private-visual/lokad-value-based/20260524-155508` stayed stable: 84/84 compared
  pages, zero dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Move direct major/minor chart-gridline stroke ownership into `PptxSceneChartAxis`.
  Supported bar and line chart renderers now consume scene-owned gridline color, width, alpha, and explicit
  `a:ln/a:noFill` for the main value-axis gridline path, preserving the previous gray defaults when no
  explicit gridline style is present. Focused model/chart tests passed after a transient parallel build lock
  was rerun serially, the full runner passed 186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-160202` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Preserve and consume chart stroke dash/cap/join metadata on scene-owned chart lines.
  `ReadChartLine` now keeps `a:prstDash`, line caps, and joins in `PptxSceneLineStyle`; the chart stroke
  renderer value object carries those fields through gridlines, axes, series, markers, points, and chart
  area/plot-area strokes instead of truncating them to color/alpha/width. Focused model/chart tests passed
  after a transient parallel build lock was rerun serially, the full runner passed 186/186, `dotnet pack`
  succeeded, and private run `artifacts/private-visual/lokad-value-based/20260524-170506` stayed stable:
  84/84 compared pages, zero dimension mismatches, deck MAE `9.043200`, changed16 `0.116408`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [ ] Extend chart gridline styling beyond direct solid/noFill/dash/cap/join lines: compound lines,
  theme style references, and chart-style inherited defaults still need typed ownership before gridlines can
  be considered fully structurally aligned with Office.
- [x] 2026-05-24: Move simple chart text-style overrides into the scene model. `PptxSceneChart.TextStyle`
  and `PptxSceneChartAxis.TextStyle` now preserve chart-level and axis-level `c:txPr/a:defRPr` font family,
  font size, and solid text color, and supported category/value axis labels consume those scene styles before
  XML fallback. Focused model/chart tests passed after a transient parallel build lock was rerun serially,
  the full runner passed 186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-160802` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Preserve and consume simple chart title/legend `txPr/a:defRPr` overrides through the
  typed chart scene model. `PptxSceneChartTitle` and `PptxSceneChartLegend` now carry nullable text-style
  overrides, and supported title/legend rendering merges chart-level defaults with title/legend font family,
  font size, solid color, bold, and italic instead of using fixed renderer-local text styles. Public scene and
  PDF tests lock title and legend font/color consumption, the full runner passed 186/186, `dotnet pack`
  succeeded, and private run `artifacts/private-visual/lokad-value-based/20260524-181800` stayed stable:
  84/84 compared pages, zero dimension mismatches, deck MAE `9.042022`, changed16 `0.116405`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [ ] Extend chart text-style ownership beyond simple `defRPr` font/color/size and default-run bold/italic:
  rich text runs, rotation, full non-default tick-label offset ladders, multi-level category labels, and
  chart-style inherited text defaults still need structural modeling before axis and data-label text can
  match Office without renderer heuristics.
- [x] 2026-05-24: Move simple plot-area manual-layout ownership into `PptxSceneChart.PlotAreaLayout`.
  Supported bar and line chart layouts now consume scene-owned `c:plotArea/c:layout/c:manualLayout`
  factors before XML fallback, preserving the existing candidate geometry while moving another chart layout
  decision out of renderer-local XML scans. Focused model/chart tests passed after a transient parallel build
  lock was rerun serially, the full runner passed 186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-161346` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Consume manual plot-area edge modes for right/bottom coordinates. The renderer now carries
  preserved `wMode="edge"` and `hMode="edge"` through scene and XML fallback paths so right/bottom layout
  edges are not mistaken for width/height factors. A public PDF test locks the edge-mode rectangle, the full
  runner passed 187/187, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-182312` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.042022`, changed16 `0.116405`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [ ] Extend plot-area layout ownership beyond current `x/y/w/h` factor and right/bottom edge support:
  `layoutTarget`, x/y edge semantics, inner-vs-outer plot area semantics, title/legend overlay interactions,
  and reuse across area/scatter/radar/pie/doughnut chart families still need structural modeling before plot
  bounds can be treated as Office-aligned instead of approximate geometry.
- [x] 2026-05-24: Extend scene-owned chart data-label metadata beyond value/percent visibility. The scene
  model now preserves category-name, series-name, and leader-line flags plus label position, separator, and
  number-format metadata from plot/series `c:dLbls`; current renderers still consume only the already-rendered
  value/percent subset, but the richer Office label contract is no longer discarded at scene build time.
  Later on 2026-05-24, supported bar/line renderers also began consuming the preserved category-name and
  series-name flags for label text composition; leader-line geometry, position semantics, rich text, per-label
  overrides, and data-label shape styles remain open.
  Focused model/chart tests passed after a transient parallel build lock was rerun serially, the full runner
  passed 186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-161747` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [ ] Extend data-label rendering to consume the remaining richer scene metadata: leader-line geometry, rich text
  runs inside custom labels, exact label-box geometry/auto-fit, and richer position semantics still need
  renderer support and visual cases.
- [x] 2026-05-24: Make secondary value-axis label rendering consume scene-owned axis metadata when available.
  Combo and secondary-axis fallback paths now carry the matching right-side `PptxSceneChartAxis` into
  visibility, scaling, unit, number-format, and text-style decisions instead of dropping back to raw axis XML
  on those label paths. Focused model/chart tests passed after a transient parallel build lock was rerun
  serially, the full runner passed 186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-162216` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Preserve chart axis crossing and orientation metadata in the scene model. `PptxSceneChartAxis`
  now owns `crossAx`, `crosses`, `crossesAt`, `crossBetween`, and reversed scaling orientation so future
  secondary-axis and crossing layout can bind axes structurally instead of inferring sides from ad hoc XML
  searches. The scene fixture locks the metadata, the full runner passed 187/187, `dotnet pack` succeeded,
  and private run `artifacts/private-visual/lokad-value-based/20260524-182625` stayed stable: 84/84 compared
  pages, zero dimension mismatches, deck MAE `9.042022`, changed16 `0.116405`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-25: Consume value-axis crossing metadata for value-gridline endpoint filtering.
  Supported bar and line chart renderers now derive the excluded gridline tick from scene/XML `c:crossesAt`
  or `c:crosses`, falling back to Office-like `autoZero` behavior: zero when the value range contains zero,
  otherwise the nearest visible endpoint. This removes the previous hard-coded "exclude minimum" rule for
  non-default crossings while preserving the grouped gridline PDF structure and XML fallback paths. A
  synthetic chart with `c:crosses val="max"` locks the behavior by requiring the maximum crossing tick to be
  omitted while the minimum tick remains a gridline. Focused chart tests passed with 13 passed, 0 failed,
  0 skipped; the clustered-column public visual gate passed at
  `artifacts/visual/pptx-ladder-11-chart-column-clustered-port/20260525-020737`; the full suite passed with
  209 passed, 0 failed, 0 skipped; and `dotnet pack` succeeded. Remaining gap: the axis line and bar/line
  coordinate crossing geometry still do not fully consume `crosses`, `crossesAt`, reversed axes, or
  secondary-axis bindings.
- [x] 2026-05-25: Consume value-axis crossing metadata for vertical chart category-axis strokes.
  Supported vertical bar and line chart renderers now place the category-axis line at the scene/XML
  value-axis crossing value instead of always drawing it at the old zero/min baseline. The same crossing
  resolver used by gridlines feeds the axis stroke, so `crosses=max`, `crosses=min`, explicit `crossesAt`,
  and `autoZero` follow one axis-owned rule. A synthetic `c:crosses val="max"` chart locks the horizontal
  category-axis stroke at the value-axis maximum. Focused chart tests passed with 14 passed, 0 failed,
  0 skipped; the clustered-column public visual gate passed at
  `artifacts/visual/pptx-ladder-11-chart-column-clustered-port/20260525-021112`; the full suite passed with
  210 passed, 0 failed, 0 skipped; and `dotnet pack` succeeded. Remaining gap: horizontal bar axes,
  value-axis side/cross-axis placement, reversed scales, series baselines, and secondary axes still need
  separate Office-PDF-backed slices.
- [x] 2026-05-25: Consume value-axis reversed orientation on the supported vertical bar/line value-position
  path. `c:scaling/c:orientation val="maxMin"` now flows from scene/XML axis metadata into value-axis
  labels, gridline coordinates, vertical category-axis crossing placement, line-series point positions, and
  clustered vertical bar endpoints instead of preserving the metadata without using it. A synthetic two-point
  line chart locks the inverted line geometry. Focused chart tests passed with 15 passed, 0 failed,
  0 skipped; the clustered-column public visual gate passed at
  `artifacts/visual/pptx-ladder-11-chart-column-clustered-port/20260525-021718`; the full suite passed with
  211 passed, 0 failed, 0 skipped; and `dotnet pack` succeeded. Remaining gap: horizontal bars, stacked
  bar/column accumulation, chart data-label positions, secondary-axis orientation, and Office-backed public
  visual coverage for reversed axes still need separate slices.
- [x] 2026-05-24: Make same-side secondary value-axis slotting scene-aware on the supported bar/combo path.
  The side-slot resolver now consumes scene-owned tick-label position when available instead of re-reading
  raw axis XML, keeping raw XML only as fallback. The full runner passed 187/187, `dotnet pack` succeeded,
  and private run `artifacts/private-visual/lokad-value-based/20260524-183014` stayed stable: 84/84 compared
  pages, zero dimension mismatches, deck MAE `9.042022`, changed16 `0.116405`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Preserve chart axis tick and label-placement metadata in the scene model. `PptxSceneChartAxis`
  now owns major/minor tick mark choices, label offset, tick-label skip, tick-mark skip, and the multi-level
  category-label suppression flag so future axis layout can use Office-authored structure instead of fixed
  label offsets. The scene fixture locks the metadata, the full runner passed 187/187, `dotnet pack`
  succeeded, and private run `artifacts/private-visual/lokad-value-based/20260524-183351` stayed stable:
  84/84 compared pages, zero dimension mismatches, deck MAE `9.042022`, changed16 `0.116405`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Preserve chart style IDs in the scene model. `PptxSceneChart.StyleId` now carries
  `c:chartSpace/c:style @val`, giving future chart-style inherited defaults a structural owner instead of
  forcing renderers to rediscover the raw chart XML. The scene fixture locks the metadata, the full runner
  passed 187/187, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-183716` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.042022`, changed16 `0.116405`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Carry data-label leader-line flags through the renderer option model instead of dropping
  them after scene/XML parsing. `ChartDataLabelOptions` and point overrides now preserve
  `showLeaderLines` through plot, series, XML fallback, and per-label resolution, so the remaining
  leader-line work is geometry/output rather than metadata recovery. The full runner passed 187/187,
  `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-185426` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.042022`, changed16 `0.116405`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Honor chart legend overlay semantics in supported bar-chart plot-box sizing. A visible
  legend with `overlay=true` now draws as an overlay instead of reserving plot space, while non-overlay
  legends keep the existing reserved layout. A public synthetic bar-chart test locks the plot-area rectangle
  for an overlaid bottom legend, the full runner passed 188/188, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-190019` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.042022`, changed16 `0.116405`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Normalize chart data-label boolean parsing to the shared OOXML boolean-element rule.
  Renderer-side data-label flags now treat element-only booleans as enabled and `val="false"` as disabled,
  matching the scene model and avoiding a raw-XML fallback discrepancy. The scene fixture now includes both
  boolean forms for chart labels, the full runner passed 188/188, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-190422` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.042022`, changed16 `0.116405`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Normalize raw chart legend overlay parsing to the shared OOXML boolean-element rule.
  The raw fallback now treats `<c:overlay/>` like `val="1"`, matching the scene-owned legend metadata and
  removing another chart-path discrepancy between typed parsing and direct XML rendering. The public overlay
  legend layout test now uses the element-only form, the full runner passed 188/188, `dotnet pack` succeeded,
  and private run `artifacts/private-visual/lokad-value-based/20260524-190742` stayed stable: 84/84 compared
  pages, zero dimension mismatches, deck MAE `9.042022`, changed16 `0.116405`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Normalize chart `varyColors` parsing to the shared OOXML boolean-element rule while
  preserving the existing default-on behavior when the element is absent. The scene model and raw chart
  fallback now both treat `val="false"` like `val="0"` instead of accidentally enabling per-point colors;
  public fixtures cover `val="false"` in scene parsing and element-only true in rendering. Full runner passed
  188/188, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-191141` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.042022`, changed16 `0.116405`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Normalize custom-geometry path `stroke` parsing to OOXML boolean attribute semantics.
  Scene-owned custom paths and the raw custom-geometry fallback now treat `stroke="0"` like `stroke="false"`
  while preserving the default-on behavior when the attribute is absent. The public scene fixture asserts
  `AllowsStroke=false` for `stroke="0"`, full runner passed 188/188, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-191516` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.042022`, changed16 `0.116405`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Normalize DOCX `w:onOff` run-property parsing for WordprocessingML `on`/`off` values.
  The DOCX reader now treats `w:val="on"` as enabled and `w:val="off"` as disabled for run properties such
  as bold and italic, matching the existing element-only, `1`/`0`, and `true`/`false` handling. A public
  model-level unit test covers both forms; full runner passed 189/189 and `dotnet pack` succeeded. No private
  PPTX rerun was taken for this DOCX-only parser slice.
- [x] 2026-05-24: Centralize OOXML boolean parsing behind `OoxBoolean` in the shared OOXML layer.
  PPTX chart flags, PPTX custom geometry/table-style booleans, PPTX transform flags, and DOCX `w:onOff`
  parsing now flow through one helper for attribute, optional-attribute, and element-with-`val` forms. This
  reduces format-local boolean drift after the recent normalization slices; full runner passed 189/189 and
  `dotnet pack` succeeded.
- [x] 2026-05-24: Pair secondary value-axis scene metadata to raw chart XML by `axId` before consuming
  right-axis style, scale, units, and labels. Combo bar-chart secondary detection now accepts a scene-owned
  right value axis when the raw axis is absent or less complete, and the fallback secondary-label/stroke paths
  no longer independently pick unrelated first-right XML and scene axes. Focused chart tests passed 5/5, the
  full runner passed 189/189, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-192509` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.042022`, changed16 `0.116405`, and only one diagnostic
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`. This is an identity-alignment slice only: exact secondary-axis crossing
  geometry and Office spacing remain open.
- [ ] Finish secondary-axis structural alignment for chart families beyond the current supported bar/line
  paths: crossing geometry still needs to consume the preserved scene metadata, and exact Office spacing still
  needs explicit model-to-renderer plumbing.
- [x] 2026-05-24: Consume scene-owned data-label separator and number-format metadata in supported label
  rendering. Bar and line value labels now format through `ChartDataLabelOptions`, and pie/doughnut
  value-plus-percent labels use the typed separator instead of a fixed comma when the OOXML provides one;
  raw XML fallback still parses the same metadata when no scene plot is available. Focused model/chart tests
  passed after a transient parallel build lock was rerun serially, the full runner passed 186/186,
  `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-162638` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Consume scene/XML data-label category-name and series-name flags in supported bar/line
  rendering. Data labels now compose series name, category name, and value text from typed scene series and
  category metadata with the OOXML separator. Public bar/line chart tests lock a custom separator glyph in
  emitted PDF text, the full runner passed 186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-171811` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043200`, changed16 `0.116408`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Consume scene/XML data-label position metadata in supported bar/line rendering.
  `dLblPos` now flows into `ChartDataLabelOptions`, and explicit OOXML positions drive bar and line label
  anchors instead of being discarded after scene parsing. Public line-chart tests lock a below-positioned
  data-label text matrix, the full runner passed 186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-172435` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043200`, changed16 `0.116408`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Preserve and consume plot/series data-label `txPr` text style in supported chart labels.
  `PptxSceneChartDataLabels` now carries font family, font size, and solid text color from `c:dLbls/c:txPr`,
  and bar/line/pie/doughnut label rendering merges that typed style into the label text runs instead of using
  a fixed 8.5pt theme-text default. Public scene tests lock the typed model fields, and the line-chart PDF
  test locks the resulting label font size, RGB fill operator, and shifted text matrix. Focused model/chart
  tests passed, the full runner passed 186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-173244` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043200`, changed16 `0.116408`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Preserve and consume plot/series data-label `spPr` shape style in supported chart labels.
  `PptxSceneChartDataLabels` now carries label fill/line shape style from `c:dLbls/c:spPr`, and supported
  bar/line/pie/doughnut labels draw the styled label rectangle before emitting text. Public scene tests lock
  fill/line style ownership, and the line-chart PDF test locks the label rectangle fill/stroke operators
  before the label text operators. Focused model/chart tests passed after the known transient parallel build
  lock was rerun serially, the full runner passed 186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-173836` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043200`, changed16 `0.116408`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Preserve and consume indexed `c:dLbl` per-label overrides for supported chart labels.
  `PptxSceneChartDataLabels` now owns an indexed override list carrying visibility flags, position,
  separator, number format, text style, and shape style. Supported bar/line/pie/doughnut label rendering
  merges those overrides by point index before formatting, positioning, drawing label boxes, and emitting
  text. Public scene tests lock the override model, and the line-chart PDF test locks a point-specific label
  box fill, font size, text color, and centered text matrix. Focused model/chart tests passed after the known
  transient parallel build lock was rerun serially, the full runner passed 186/186, `dotnet pack` succeeded,
  and private run `artifacts/private-visual/lokad-value-based/20260524-174557` stayed stable: 84/84 compared
  pages, zero dimension mismatches, deck MAE `9.043200`, changed16 `0.116408`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Preserve custom chart data-label text from `c:dLbl/c:tx` in the scene model and consume it
  before composed series/category/value text. The first pass flattens `c:tx/c:rich` and cached string
  references into one label string; run-level rich formatting remains intentionally open under chart text
  style work. Focused scene/chart tests passed, the full runner passed 186/186, `dotnet pack` succeeded, and
  private run `artifacts/private-visual/lokad-value-based/20260524-175407` stayed stable: 84/84 compared
  pages, zero dimension mismatches, deck MAE `9.043200`, changed16 `0.116408`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Preserve series-level chart data-label definitions separately from plot-level labels.
  `PptxSceneChartSeries.DataLabels` now owns `c:ser/c:dLbls` with an explicit `IsDefined` bit, and supported
  bar/line renderers prefer series-owned label options only for series that actually define them. This
  removes the previous first-series collapse for multi-series labels while preserving plot-level fallback for
  ordinary charts. Focused scene/chart tests passed, the full runner passed 186/186, `dotnet pack` succeeded,
  and private run `artifacts/private-visual/lokad-value-based/20260524-175927` stayed stable: 84/84 compared
  pages, zero dimension mismatches, deck MAE `9.043246`, changed16 `0.116409`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [ ] Continue data-label rendering alignment with Office: leader-line geometry, rich text runs inside labels,
  label layout/auto-fit, and exact Office label-box geometry remain approximate.
- [x] 2026-05-24: Preserve plot-area manual-layout target and mode fields in the scene model. `PptxSceneChart`
  now carries `layoutTarget`, `xMode`, `yMode`, `wMode`, and `hMode` alongside the existing manual
  `x/y/w/h` factors, so later plot-box work can distinguish Office's inner/outer target and factor/edge
  semantics instead of treating all manual layouts as anonymous factors. Rendering still uses the previous
  geometry until those semantics are implemented. Focused model/chart tests passed after a transient parallel
  build lock was rerun serially, the full runner passed 186/186, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-163057` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one
  `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained dimension-matched at MAE `2.945717`, changed16
  `0.045530`, SSIM `0.917662`.
- [ ] Consume plot-area manual-layout target and mode semantics in geometry: `layoutTarget=inner`,
  edge/factor modes, title/legend overlay interactions, and non-bar/line chart-family plot boxes still need
  Office-evidenced rendering rules.
- [x] 2026-05-24: Re-ran the full test suite, package, and private PPTX acceptance after scene-owned
  backgrounds. The test runner executed 183/183 passing tests, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-120402` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, only `PPTX_UNSUPPORTED_IMAGE_RECOLOR`,
  and slide 17 MAE `2.945717`, changed16 `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Re-ran package and private PPTX acceptance after moving master/layout shape/text rendering
  to ordered scene dispatch. `dotnet pack` succeeded and private run
  `artifacts/private-visual/lokad-value-based/20260524-114540` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, only `PPTX_UNSUPPORTED_IMAGE_RECOLOR`,
  and slide 17 MAE `2.945717`, changed16 `0.045530`, SSIM `0.917662`.
- [ ] Keep raw XML on scene nodes until typed coverage is complete, but make new renderer code prefer typed
  fields and resolver outputs instead of repeated ad hoc descendant queries.
- [ ] Decide whether `ooxpdf` needs an intermediate presentation scene/model between OOXML parsing and PDF
  generation before more large changes to `PptxRenderer`.
- [ ] Port `pptx-renderer` package/model boundary patterns: relationship target resolution, media lookup,
  slide/layout/master/theme ownership, and raw XML retention on typed nodes.
- [ ] Port `pptx-renderer` render-context ownership: one context object should expose slide identity,
  inheritance, relationships, media caches, theme/color resolvers, diagnostics, and group-fill state.
- [ ] Port `pptx-renderer` specialized renderer split: background, shape, text, table, image, group, chart,
  and fallback renderers should consume the same context instead of ad hoc XML/document parameters.
- [ ] Port `pptx-renderer` slide inheritance behavior: render master/layout non-placeholder template content
  behind slide content, and treat placeholders as inheritance templates for geometry/body/text styles.
- [ ] Port `pptx-renderer` placeholder matching rules: match by `idx` before type, handle default/body/title
  fallback consistently, and keep placeholder geometry separate from text-body inheritance.
- [ ] Port `pptx-renderer` text cascade layers as explicit records:
  `defaultTextStyle`, master text styles, master placeholder, layout placeholder, shape list style,
  paragraph properties, default run properties, and run properties.
- [ ] Port `pptx-renderer` theme font resolution: major/minor Latin, East Asian, complex script, and symbol
  font fallback must be explicit diagnostics-bearing stages before glyph mapping.
- [ ] Port `pptx-renderer` color resolver coverage: color maps, theme colors, `phClr`, scheme colors,
  preset colors, HSL/scrgb colors, alpha/lum/tint/shade modifiers, and fallback colors.
- [ ] Port `pptx-renderer` format-scheme fill/line resolution: `fillRef`, `lnRef`, style lists, `phClr`
  replacement, and default shape style resolution should be model-visible.
- [ ] Port `pptx-renderer` text body handling: insets, anchors, vertical overflow, fit modes, wrapping,
  text direction, vertical text, multi-column text, and unsupported diagnostics where rendering is absent.
- [ ] Port `pptx-renderer` line-height behavior completely: `spcPct`, `spcPts`, paragraph before/after
  spacing, `lnSpcReduction`, manual-break line wrappers, and trailing `endParaRPr` line-height effects.
- [ ] Port `pptx-renderer` run style behavior: font family, size, bold/italic, underline, strike, color,
  highlight, character spacing, kerning thresholds, caps, superscript/subscript, and hyperlink style.
- [ ] Port `pptx-renderer` whitespace behavior: regular spaces, repeated spaces, non-breaking spaces,
  tabs, soft hyphens, explicit line breaks, fields, and end-paragraph runs must remain observable.
- [ ] Port `pptx-renderer` bullet and numbering behavior: bullet suppression for metadata placeholders,
  `buChar`, `buAutoNum`, `buFont`, `buClr`, `buSz`, hanging indents, and inherited bullet defaults.
- [ ] Port `pptx-renderer` table text behavior: cell text style inheritance, table style overrides,
  vertical alignment, margins, merged cells, and per-cell text diagnostics.
  - [x] Add the first table scene record: `PptxSceneTable` now carries raw grid column widths and row heights,
    and ordered table rendering consumes those typed layout primitives instead of parsing grid/row dimensions
    inside the renderer.
  - [x] Extend `PptxSceneTable` to row/cell layout metadata: column spans, row spans, and horizontal/vertical
    merge continuations are now parsed into scene records and consumed by ordered table rendering.
  - [x] Move table-cell text margins and vertical anchors into scene records. Ordered table text rendering now
    consumes scene-owned insets/anchors at the text-shape bridge instead of re-reading `tcPr` at draw time.
  - [x] Move direct table-cell solid fills into `PptxSceneTableCell`. Ordered table rendering now consumes
    scene-owned `tcPr` fills; built-in table-style fallback fills remain renderer-side until table-style
    conditional formatting is represented as a typed style-resolution model.
  - [x] Move explicit table-cell borders into `PptxSceneTableCell`. The scene now distinguishes explicit
    border presence from drawable border lines, preserving the current behavior where explicit `noFill`
    borders suppress the default grid while rendered border strokes are emitted from typed line data.
  - [x] Move built-in table-style descriptors into `PptxSceneTable`: style id, supported style family/accent,
    and conditional-format flags are now scene-owned inputs for table fills, text defaults, and default grid
    behavior. The Office table-style cascade itself remains a future typed resolver.
  - [x] Materialize per-cell built-in table-style outputs in `PptxSceneTableCell`: the renderer now consumes
    scene-owned style fills and text defaults after direct cell formatting. The current limited built-in
    formulas are isolated in `PptxTableStyleResolver`, pending replacement by a full Office table-style
    cascade model.
  - [x] Move raw table-cell `txBody` ownership into `PptxSceneTableCell` as an interim bridge. Table text is
    still rendered through the shared text-shape XML bridge, but draw-time cell text no longer re-reads
    `a:tc/a:txBody` from the table cell XML.
  - [x] Move the raw table element retained for table row/text traversal into `PptxSceneTable.Source`.
    Ordered table rendering now starts from the scene table record instead of re-reading
    `node.Source` to find `a:tbl`.
  - [ ] Replace `PptxTableStyleResolver`'s supported-style formulas with a real Office table-style cascade:
    parse table style parts/theme style matrices, conditional formatting priority, `phClr` replacement, and
    unsupported-style diagnostics instead of expanding GUID-specific logic.
  - [x] Route ordered table frame bounds through `PptxSceneNode.Bounds` and the active group transform instead
    of re-reading untransformed graphic-frame bounds. Public unit
    `PptxSyntheticGroupedTableUsesGroupTransform` locks grouped table placement at the PDF-operator level.
  - [x] Apply the same group-transform traversal to the legacy table fallback path used when unknown graphic
    frames disable ordered scene rendering. Public unit
    `PptxSyntheticGroupedTableUsesGroupTransformInFallbackPath` locks that fallback behavior.
- [ ] Port `pptx-renderer` shape geometry coverage: preset geometries, custom geometry, rotations, flips,
  group transforms, connectors, arrows, dash/cap/join, and picture-fill clipping.
- [ ] Port `pptx-renderer` image behavior: relationship resolution, crop/fill/stretch, alpha/soft masks,
  SVG or unsupported-image diagnostics, media caching, and reuse across slides.
- [ ] Port `pptx-renderer` chart behavior as a first-class native renderer: parse a typed chart model for
  series, axes, legends, labels, styles, and layouts; emit diagnostics only for unsupported chart features.
  - [x] Add the first chart scene record: `PptxSceneChart` now carries the chart relationship id, and ordered
    scene dispatch resolves chart parts from typed scene data instead of re-reading the graphic frame source
    XML. Chart rendering itself is still the existing XML-driven native renderer until a typed chart model is
    introduced.
  - [x] Resolve chart relationship targets during scene construction: `PptxSceneChart.TargetPartName` now carries
    the package-resolved chart part for slide/layout/master relationship scopes, and ordered scene chart rendering
    consumes that target directly while keeping relationship-map fallback for raw XML rendering.
  - [x] Move chart-part XML and chart color-style palette ownership into `PptxSceneChart`: ordered scene chart
    rendering now consumes scene-owned `ChartXml` and `PaletteColors`, while the raw XML path still opens chart
    parts directly as a compatibility fallback.
  - [x] Add the first typed chart-plot summary: `PptxSceneChart.Plots` records each plot family's OOXML kind,
    series count, and referenced axis ids so future combo/multi-axis rendering can bind chart groups to axes
    without repeated loose descendant scans.
  - [x] Add the first typed chart-axis catalog: `PptxSceneChart.Axes` records OOXML axis kind, id, position,
    and delete state so plot-to-axis binding can become structural instead of repeatedly searching raw XML.
  - [x] Extend the chart-axis scene model with value-axis scaling, major/minor units, and gridline
    visibility, and consume those records first in bar/line chart rendering.
  - [x] Extend the chart-axis scene model with tick-label position and number format, and consume those
    records first for bar/line axis-label visibility, side placement, and numeric formatting.
  - [x] Add typed explicit chart title and legend metadata to `PptxSceneChart`; this keeps OOXML title/legend
    state distinct from renderer fallback heuristics and prepares title/legend layout to become scene-driven.
  - [x] Add the first typed chart-series summaries to `PptxSceneChart.Plots`: series names, cached numeric
    values, and category labels are now scene data, preparing chart rendering to consume structural model
    records instead of repeatedly scanning `c:ser` XML.
  - [x] Preserve scatter/bubble chart data channels in `PptxSceneChartSeries`: `c:xVal`, `c:yVal`, and
    `c:bubbleSize` now stay distinct from category/value series instead of being renderer-only XML scans.
  - [x] Add typed chart plot attributes to `PptxSceneChartPlot`: grouping, bar direction, scatter style,
    vary-colors, gap width, overlap, and doughnut hole size now have a scene-model owner before renderer
    consumption is migrated away from raw plot XML.
  - [x] Make ordered chart rendering consume scene-owned plot attributes first for bar grouping/direction,
    vary-colors, gap/overlap, area grouping, scatter line style, and doughnut hole size, retaining raw XML
    fallback only for legacy chart paths.
  - [x] Make ordered chart rendering consume scene-owned cached chart series first for bar, line, area, radar,
    pie, doughnut, scatter, and bubble charts; raw `c:ser` scans are now fallbacks for legacy chart paths and
    unmodeled data.
  - [x] Make ordered bar/line chart category-label rendering consume scene-owned cached category labels first,
    leaving raw `c:cat` scans as fallback for legacy chart paths.
  - [x] Add typed chart-series style metadata to `PptxSceneChartSeries`: direct solid fill, direct line,
    marker symbol/size/fill/line, and smooth flag now have scene-model ownership before renderer consumption
    is migrated away from raw `c:ser` style scans.
  - [x] Make ordered chart rendering consume scene-owned series style metadata first for fills, strokes,
    markers, and smooth flags, retaining raw `c:ser` style scans only as fallback.
  - [x] Add and consume typed chart point-style metadata: `PptxSceneChartPointStyle` preserves `c:dPt`
    index, direct fill, pattern fill, line, and raw explosion values, with percentage normalization only
    at the PDF-rendering boundary.
  - [x] Make chart title/legend layout scene-first: supported chart layouts and fallback title drawing now
    consume `PptxSceneChartTitle`/`PptxSceneChartLegend`, while XML fallback remains for legacy paths and
    the single-series automatic bar-title rule.
  - [x] Add and consume typed chart data-label visibility options: `PptxSceneChartPlot.DataLabels` preserves
    plot/series fallback `showVal` and `showPercent` flags, and bar, line, pie, and doughnut rendering now
    consume those options before raw `c:dLbls` fallback.
  - [x] Extend chart data-label scene metadata to preserve category-name, series-name, leader-line,
    position, separator, and number-format options.
  - [x] Consume scene/XML data-label separator and number-format metadata for supported value-label text.
  - [x] Consume scene/XML data-label category-name and series-name flags in supported bar/line rendering.
  - [x] Consume scene/XML data-label position metadata in supported bar/line rendering.
  - [ ] Extend chart data-label rendering to cover leader-line geometry, rich text runs inside custom labels,
    Office label-box geometry/auto-fit, and richer position semantics before attempting finer
    Office-aligned data-label layout.
  - [x] Make legend-entry names scene-first for supported bar/combo and line chart paths: legend builders
    now consume scene series names and reserve raw `c:ser` name scans for fallback paths.
  - [x] Add and consume scene-owned chart area and plot area solid fill/line styles, so supported chart
    rendering no longer re-scans chart-level and plot-area `c:spPr` for the simple style path.
  - [x] Preserve explicit chart-area/plot-area `a:noFill` state in scene shape-style records and consume it
    at the renderer boundary without changing the visible no-background output.
  - [x] Preserve and consume chart-area/plot-area pattern fills in scene shape-style records using the same
    chart pattern-fill renderer path as series and point fills.
  - [ ] Extend chart area and plot area style records to cover the remaining shape-style family instead of
    only direct solid fill/no-fill, pattern fill, and simple line.
  - [x] Add and consume scene-owned chart-axis line styles for supported bar and line charts, including
    explicit `a:ln/a:noFill`, so axis stroke handling follows the typed axis catalog instead of renderer
    XML scans on the main path.
  - [x] Add scene-owned major/minor gridline style records and consume direct solid/noFill strokes for
    color, width, and alpha; hard-coded gray strokes are now defaults only when no explicit gridline style
    is present.
  - [x] Preserve and consume scene-owned chart line dash/cap/join fields through chart strokes, including
    gridlines, axes, series, markers, points, and chart/plot-area borders.
  - [ ] Extend scene-owned gridline style records to cover compound lines, theme style references, and
    chart-style inherited defaults.
  - [x] Add scene-owned chart-level, axis-level, title, and legend text-style overrides for the supported
    `txPr/defRPr` subset: font family, font size, solid color, bold, and italic.
  - [ ] Extend chart text-style records to cover rich text runs, rotation, full non-default tick-label offset
    ladders, multi-level category labels, per-run data-label styles, and chart-style inherited defaults.
  - [x] Carry scene-owned secondary value-axis metadata into label visibility, formatting, scale, unit, and
    text-style decisions for supported bar/line chart paths.
  - [x] Preserve cross-axis IDs, crossing behavior, cross-between, and reversed orientation metadata in
    `PptxSceneChartAxis`.
  - [x] Consume scene-owned tick-label side metadata for same-side secondary value-axis label slotting on
    supported bar/combo paths.
  - [x] Preserve chart-axis tick mark, label offset, tick skip, and multi-level label metadata in
    `PptxSceneChartAxis`.
  - [x] Preserve chart style IDs in `PptxSceneChart` as a prerequisite for chart-style inherited defaults.
  - [x] Consume value-axis crossing metadata for bar/line chart gridline endpoint filtering.
  - [x] Consume value-axis crossing metadata for vertical bar/line category-axis stroke placement.
  - [x] Consume value-axis reversed orientation for vertical bar/line gridlines, value labels, crossing-line
    placement, line-series geometry, and clustered vertical bar endpoints.
  - [ ] Consume axis crossing/orientation metadata for horizontal bars, stacked bars/columns, data labels,
    value-axis side/cross-axis placement, series coordinate baselines, and secondary axes instead of relying
    on right-side XML/layout assumptions.
  - [x] Add and consume scene-owned plot-area manual-layout factors for supported bar and line charts.
  - [x] Preserve scene-owned plot-area manual-layout target and mode fields.
  - [x] Consume scene/XML `wMode="edge"` and `hMode="edge"` manual-layout semantics for right/bottom plot-area
    edges.
  - [ ] Extend chart plot-area layout records to cover `layoutTarget`, x/y edge semantics, inner/outer plot
    semantics, title/legend overlay effects, and non-bar/line chart-family consumers in rendered geometry.
- [ ] Keep SmartArt as a separate diagnostics-first feature until a real SmartArt renderer exists.
- [ ] Port `pptx-renderer` error isolation: one unsupported or malformed node should emit a diagnostic with
  slide/node context instead of aborting the whole render pass when recovery is possible.
- [ ] Port `pptx-renderer` generated public visual-suite organization: normalized case names, grouped
  fixtures, Office reference caching, and separate approximate, needs-review, and locked thresholds.
- [ ] Port `pptx-renderer` oracle tooling ideas that fit `.NET`: compact text-op diffs, visual metrics,
  cached Office references, deterministic artifact paths, and fast focused case selection.
- [ ] Port `pptx-renderer` performance lessons: avoid repeated ZIP/XML/theme/font parsing, cache immutable
  resources per render pass, and measure large-deck hot spots before private-deck tuning.
- [x] Prototype the smallest `ooxpdf` PPTX intermediate scene slice for slide/master/layout node lists,
  node kind classification, placeholder metadata, and bounds extraction.
- [x] Extend the `ooxpdf` scene slice to text bodies: body properties, list style, paragraphs, levels,
  ordered text/break/field runs, run properties, and end paragraph properties.
- [x] Extend the `ooxpdf` scene slice to resolved text styling: paragraph style cascade and run style
  cascade from master/layout/shape/paragraph/run context.
- [x] Compare the current direct `ReadTextRuns` path with the model-slice output on Ladder 4 typography cases
  before replacing behavior.
- [x] Fix a direct-renderer text cascade gap found by that comparison: paragraph defaults now use the
  paragraph's actual `lvlNpPr` level, not only `lvl1pPr`, and the direct inspection hook exposes resolved
  text-run snapshots for future scene-vs-renderer checks.
- [ ] Consider porting the testing strategy, especially generated Office-oracle case families and SSIM plus
  color-histogram metrics, while keeping `src/Lokad.OoxPdf` dependency-free.
- [x] Add dependency-free VisualDiff metrics for global luminance SSIM and foreground RGB histogram
  correlation, so MAE remains diagnostic rather than the only quality signal.
- [x] Fix `tools/CheckVisualCase.ps1` stale-build detection for CLI-based validation: the harness now
  considers `src/Lokad.OoxPdf` source timestamps, not only `src/Lokad.OoxPdf.Cli`.

## PPTX Renderer Test Port Plan

`pptx-renderer` has a much broader and better organized PPTX test corpus. Treat it as a test-design asset:
clean-port the cases and expected behaviors into `ooxpdf`; do not vendor TypeScript/Python code or private
assets.

Porting priorities:

- [x] Inventory `pptx-renderer` unit tests by capability: model parsing, theme/style/color resolution,
  text rendering, shapes/presets, groups, images, tables, charts, SmartArt, security, and public API.
- [x] Inventory `pptx-renderer` Office-oracle/e2e cases by generated fixture family and map them to the
  `ooxpdf` public visual ladder naming scheme.
- [x] Inventory `pptx-renderer` generated typography oracle family and map it to current `ooxpdf` Ladder 4
  coverage.
- [x] Port the typography oracle family first: font families, sizes, bold/italic/underline, colors,
  mixed formatting, bullets, vertical text, anchoring, and line spacing.
- [x] Port the generated shape-adjustment oracle family, preserving Office-authored public fixtures and
  source-aligned variants from `oracle-pypptx-shape-adj-0001..0031`.
- [ ] Port broader shape/preset oracle families next, preserving Office-authored/generated public fixtures
  and adding PDF/raster inspection notes for every accepted gate.
- [ ] Port layout-composition cases: master/layout placeholders, grouped transforms, z-order, image crop,
  table placement, native chart rendering, SmartArt fallback/diagnostics, and mixed slide content.
- [x] Add focused unit locks for PPTX shape stroke dash/cap/join, RGBA PNG soft masks, merged-table interior
  grid suppression, and area/scatter/radar/doughnut native chart recognition.
- [x] Port the first test organization pattern: visual cases now have capability-family manifests and a
  family runner, matching the `pptx-renderer` idea of generated/oracle cases grouped by feature area.
- [x] Port the first unit-runner organization pattern: unit tests now come from a capability catalog and can
  be listed or filtered by group, so PPTX typography/shapes/images/composition checks can run independently.
- [x] Continue porting test organization patterns: visual family runs now emit ignored JSON/CSV reports,
  support `-OnlyUnsupported`, `-OnlyChanged`, and `-UpdateCatalog`, and can be compared for regressions.
- [x] Make strict visual support explicit: case manifests can gate on SSIM/color histogram, and family
  manifests define a stricter `supported` threshold so loose legacy gates remain runnable while bottom-up
  pixel-perfect rungs are tracked separately as `supported` versus `needs-review`.
- [ ] Continue porting generated-case manifests and richer parity reporting for the remaining
  `pptx-renderer` oracle families.
- [x] Add first parity tracking in `EXECPLAN.md`: ported families now state source coverage, `ooxpdf`
  fixture names, current status, and gate type; keep extending this table as more families are ported.
- [ ] Keep all ported tests public and synthetic. If a `pptx-renderer` case uses generated assets, recreate
  the minimal OOXML/PPTX or generator logic locally under `tests/` or `tools/` with dependency-free runtime
  constraints for `src/Lokad.OoxPdf`.

Unit-test capability inventory from `pptx-renderer`:

- Public API/viewer lifecycle tests mostly do not map to `ooxpdf`, except input-type validation,
  deterministic conversion, batch/list handling, and safe disposal of temporary resources.
- Parser/model tests map to `OoxPackage`, relationships, units, `PptxDocument`, `PptxTheme`,
  `PptxSlide`, `PptxScene`, and typed scene nodes.
- Renderer boundary tests map to future `PptxRenderContext`, background/slide/group/image/shape/text/table
  renderer partials, style resolver, color resolver, and chart renderer.
- Shape utility tests map to preset path generation, adjustment handling, custom geometry, and arc geometry.
- Utility tests map to visual metrics, media path safety, EMF/WMF diagnostics or fallback, PDF inspection,
  and preview/raster scaling.

Generated Office-oracle family inventory from `pptx-renderer`:

- Text `oracle-pypptx-text-0001..0038`: ported as public Ladder 4 typography visual cases.
- Shape adjustments `oracle-pypptx-shape-adj-0001..0031`: ported as public Ladder 6 shape-adjustment cases.
- Composite `oracle-pypptx-composite-0001..0010`: next visual-port target after shape adjustments.
- Charts `oracle-pypptx-chart-0001..0021`: later visual-port target, after native chart architecture is
  separated from the main PPTX renderer.

Current `pptx-renderer` parity tracking:

- Source `oracle-pypptx-text-0001..0038`:
  - Coverage: font families, sizes, styles, alignments, colors, bullets, vertical text, anchoring,
    and line spacing.
  - OOXPDF fixtures: `pptx-ladder-04-*` typography cases.
  - Gate type: public visual family `pptx-typography` plus unit group `pptx-typography`.
  - Status: ported; several typography rungs remain `needs-review` quality targets rather than final
    pixel-perfect gates.
- Source `oracle-pypptx-shape-adj-0001..0031`:
  - Coverage: adjustment handles for common preset shapes.
  - OOXPDF fixtures: `pptx-ladder-06-*` shape-adjustment cases.
  - Gate type: public visual family `pptx-shapes`.
  - Status: ported as Office-oracle public fixtures; remaining work is broader preset coverage and
    effect/fill families.
- Source `oracle-pypptx-composite-0001..0010`:
  - Coverage: mixed shapes, text, tables, charts, and dashboard-like slides.
  - OOXPDF fixtures: `pptx-ladder-08-composite-port-a`, `pptx-ladder-10-composite-table-port`,
    `pptx-ladder-11-composite-chart-port`, `pptx-ladder-11-composite-two-charts-port`, and
    `pptx-ladder-11-dashboard-table-chart-port`.
  - Gate type: public visual families `pptx-composition`, `pptx-tables`, and `pptx-charts`.
  - Status: partially ported; remaining composite ports should stay Office-backed and public.
- Source `oracle-pypptx-chart-0001..0021`:
  - Coverage: column, bar, line, pie, doughnut, area, scatter, radar, and bubble variants.
  - OOXPDF fixtures: `pptx-ladder-11-composite-chart-port`,
    `pptx-ladder-11-chart-column-clustered-port`, `pptx-ladder-11-chart-bar-clustered-port`,
    `pptx-ladder-11-chart-line-3series-port`, `pptx-ladder-11-chart-pie-5-categories-port`,
    `pptx-ladder-11-chart-column-negative-port`, `pptx-ladder-11-chart-column-stacked-port`,
    `pptx-ladder-11-chart-column-100-stacked-port`, `pptx-ladder-11-chart-bar-stacked-port`,
    `pptx-ladder-11-chart-line-markers-port`, `pptx-ladder-11-chart-line-stacked-port`,
    `pptx-ladder-11-chart-pie-exploded-port`, `pptx-ladder-11-chart-doughnut-port`,
    `pptx-ladder-11-chart-doughnut-exploded-port`, `pptx-ladder-11-chart-area-2series-port`,
    `pptx-ladder-11-chart-area-stacked-port`, `pptx-ladder-11-chart-line-trend-port`,
    `pptx-ladder-11-chart-scatter-clusters-port`, `pptx-ladder-11-chart-scatter-smooth-port`,
    `pptx-ladder-11-chart-radar-2series-port`, `pptx-ladder-11-chart-radar-filled-port`, and
    `pptx-ladder-11-chart-bubble-port`.
  - Gate type: public visual family `pptx-charts`.
  - Status: first bottom-up chart ports are gated with a loose native chart renderer; remaining chart
    families should be ported incrementally while chart rendering is promoted to a first-class model.

Composite oracle family map:

- Non-chart composite cases `0001..0004`, `0007`, and `0009` are ported in
  `pptx-ladder-08-composite-port-a`.
- Default shape styling/theme formatting in `pptx-ladder-08-composite-port-a` is now resolved for solid
  `fillRef` and `lnRef` references. Gradient/pattern/effect format-scheme styles remain future work.
- Table composite `0005` is ported in `pptx-ladder-10-composite-table-port` as a baseline.
  It now recognizes the common Medium Style 2 Accent 1 built-in style for first-row/accent/banded fills
  and header text color, but remains visually poor because full Office table styles are not resolved from
  all table style ids, banding flags, border layers, and cell text layout details.
- Chart composite `0006` is represented by `pptx-ladder-11-composite-chart-port` as an Office-authored
  default clustered-column chart baseline. It should be tightened after the chart renderer is driven by a
  typed model instead of ad hoc chart-family branches.
- Chart composites `0008` and `0010` are now ported as public fixtures
  `pptx-ladder-11-composite-two-charts-port` and `pptx-ladder-11-dashboard-table-chart-port`.
  They are intentionally loose chart/composition gates until native chart rendering is separated.
- Cached numeric line and pie chart rendering is now covered by a synthetic unit test and by
  `pptx-ladder-11-composite-two-charts-port`. The port improved from MAE `21.305318`, changed16
  `0.223965` with unsupported-chart diagnostics to MAE `16.063279`, changed16 `0.202201`. The dashboard
  table/chart port currently gates at MAE `18.630753`, changed16 `0.246106`.
- First standalone chart-family ports are now gated from `pptx-renderer`: clustered column MAE
  `16.371853`, changed16 `0.218999`; clustered horizontal bar MAE `15.860384`, changed16 `0.185180`;
  3-series line MAE `3.515034`, changed16 `0.032532`; and 5-category pie MAE `10.866346`,
  changed16 `0.158026`. The horizontal bar port confirms that `barDir="bar"` needs different geometry,
  but all four still need Office-like axes, gridlines, labels, legends, data labels, and chart-area layout.
- Negative and stacked chart ports are now gated: negative columns MAE `8.843941`, changed16 `0.103373`;
  stacked columns MAE `13.334796`, changed16 `0.179621`; 100% stacked columns MAE `15.340911`,
  changed16 `0.225668`; and stacked horizontal bars MAE `16.026032`, changed16 `0.189521`.
  The renderer now separates clustered, stacked, and 100% stacked bar geometry and draws bars around a
  zero axis for negative values. Remaining gaps are axis scaling/ticks, labels, legends, overlap/gap width,
  Office chart templates, and cached chart-image fallback.
- Additional chart ports are gated: line with markers MAE `2.864938`, changed16 `0.026264`;
  stacked line MAE `3.087048`, changed16 `0.026508`; exploded pie MAE `13.159580`, changed16
  `0.165800`; and doughnut MAE `12.345606`, changed16 `0.139337`. Doughnut charts now render natively
  instead of emitting an unsupported-chart diagnostic. Remaining gaps are marker shapes, stacked-line
  scaling semantics, exploded-slice offsets, doughnut hole size, labels, and legends.
- Area/long-line chart ports are now gated: exploded doughnut MAE `10.484783`, changed16 `0.123386`;
  2-series area MAE `14.977591`, changed16 `0.304524`; stacked area MAE `17.576610`, changed16
  `0.341778`; and 24-month line trend MAE `3.299024`, changed16 `0.028131`. Area charts now render from
  cached numeric values with translucent filled polygons, but still need Office series draw order, alpha,
  smoothing, labels, axes, legends, and exact plot-area bounds.
- Scatter/radar/bubble ports complete the current `pptx-renderer` chart-family import. Gated baselines:
  scatter clusters MAE `2.183515`, changed16 `0.019723`; smooth scatter MAE `2.748172`, changed16
  `0.024889`; 2-series radar MAE `4.098692`, changed16 `0.091844`; filled radar MAE `4.689211`,
  changed16 `0.106274`; and bubble MAE `5.392324`, changed16 `0.051635`. Remaining work is full Office
  chart styling/layout rather than chart-type recognition: axes/ticks, titles, legends, labels, marker
  shapes, bubble scaling, radar grid rings, smoothing, cached chart images, and theme/style colors.
- Cached bar/column chart rendering now honors simple per-series `c:spPr/a:solidFill/a:srgbClr` fills for
  clustered, stacked, 100% stacked, horizontal, and vertical bars. The next chart styling slice should
  generalize series color resolution through theme/chart style parts instead of only direct RGB fills.
- Chart series fill resolution now uses the shared PPTX solid-color resolver, so direct `srgbClr`,
  `schemeClr`, color transforms, and fill alpha flow through the same theme-aware path as shapes/tables.
  Remaining chart color work is chart style/color-style parts and defaults inherited from Office templates.
- A broad document-theme palette substitution for all chart families was rejected by the public
  `pptx-charts` gates because area and bubble ports moved in the wrong direction. The next color slice
  should be a first-class chart palette resolver that reads chart style/color parts, then falls back to
  document theme accents or legacy Office palette defaults only when those parts are absent.
- Native bar/column chart rendering now reads chart color-style relationship parts (`chartColorStyle`) and
  uses their resolved DrawingML colors before falling back to document theme accents or legacy palette
  defaults. This is locked by a synthetic chart fixture with an explicit `colorsN.xml` palette.
- Line chart renderer now reads explicit series `c:spPr/a:ln` stroke color, alpha, and width through the
  shared line resolver. The same stroke-style model should next be applied to scatter, radar, and area
  outlines before chart style/color-style parts are tackled.
- Area, scatter, bubble, and radar chart renderer now share the same explicit series fill/stroke style
  plumbing as bars and lines. Remaining chart styling work is mostly inherited chart style/color-style
  defaults, point-level overrides, marker shapes, axes/ticks, labels, legends, and exact plot-area layout.
- Line and scatter chart renderer now read basic series marker symbols and sizes, with first-pass support
  for `circle`, `square`, `diamond`, `triangle`, and `none`. Remaining marker work is Office's full marker
  preset set, marker fill/line overrides, and exact marker sizing from the chart style parts.
- Line and scatter chart markers now honor marker-level `c:marker/c:spPr` solid fills and line strokes.
  Remaining marker work is Office's full marker preset set, automatic marker inheritance from chart
  style/color-style parts, and exact Office marker sizing.
- Line and scatter chart markers now render line-only `plus`, `x`, and `dash` marker presets and filled
  `dot`/`star` presets with default stroke fallback when no marker stroke override is present. Remaining
  marker work is automatic marker inheritance from chart style/color-style parts and exact Office marker
  sizing.
- Line and scatter chart renderer now honor per-series `c:smooth` by emitting cubic Bezier paths from the
  point sequence instead of straight segments. Remaining smoothing work is exact Office spline tension and
  interaction with missing/blank points.
- Bar and line chart renderer now render simple `c:majorGridlines` inside the plot area. Remaining axis
  work is tick labels, axis scaling, axis crossing, minor gridlines, axis line styling, and exact Office
  plot-area bounds.
- Bar and line chart renderer now honor `c:valAx/c:spPr/a:ln` and `c:catAx/c:spPr/a:ln` axis line
  styling through the shared line resolver. Remaining axis work is tick labels, axis scaling, axis crossing,
  minor gridlines, and exact Office plot-area bounds.
- Bar and line chart renderer now render simple `c:minorGridlines` separately from major gridlines with
  lighter intermediate lines. Remaining gridline work is Office style inheritance, non-default intervals,
  and exact plot-area bounds.
- Bar and line chart legends now read `c:legend/c:legendPos` for basic top, bottom, left, and right
  placement and honor deleted legends instead of always emitting a right-side legend. Remaining legend work
  is exact overlay placement beyond bar plot-box reservation, manual layout, text styling, entry
  order/filtering, and exact Office spacing.
- Supported chart renderer now render simple `c:chartSpace/c:spPr` chart-area fills and borders before
  the plot content. Remaining chart-area work is rounded corners, effects, plot-area fills/borders, and
  exact Office chart layout.
- Pie and doughnut chart renderer now honor point-level `c:dPt/c:spPr` solid fills for individual slices
  through the shared theme-aware color resolver. Remaining point-level chart work is exploded offsets,
  slice borders, data labels, and inherited chart style/color-style defaults.
- Doughnut chart renderer now reads `c:holeSize` and uses it for the inner cutout instead of a fixed ratio.
  Remaining doughnut work is exploded slices, slice border styling, data labels, and Office chart layout.
- Pie and doughnut chart renderer now honor point-level `c:explosion` by offsetting slices along their
  midpoint angle. Remaining slice work is border styling, exact Office explosion scaling, and data labels.
- Pie and doughnut chart renderer now honor point-level `c:dPt/c:spPr/a:ln` slice borders through the
  shared line resolver. Remaining slice work is data labels, exact Office explosion scaling, and chart
  style/color-style inherited defaults.
- Supported chart renderer now render simple chart titles through the shared PPTX text/font pipeline.
  `PptxSceneChartTitle` preserves title overlay/manual layout and shape styling, but rendering still needs
  Office-PDF evidence before consuming those fields for exact placement. Remaining title work is exact Office
  title layout, rich text styling, and inherited chart style defaults.
- Bar and line chart renderer now honor `c:plotArea/c:spPr` fill and border styling through the shared
  chart shape-style helper. Remaining plot-area work is full manual-layout mode semantics, rounded
  corners/effects, and extending exact plot bounds to area/scatter/radar/pie/doughnut families.
- Bar and line chart renderer now render first-series cached category labels through the shared PPTX text
  pipeline. Remaining axis-label work is tick values, rich text, label rotation, multi-level categories,
  manual positioning, and Office chart font/style inheritance.
- Bar and line chart renderer now render basic value-axis tick labels from the same numeric extents used
  by the plotted series. Remaining tick-label work is explicit axis scaling/units, number formats, hidden
  axes, label positions, and exact Office chart text styling.
- Bar and line chart renderer now honor `c:valAx/c:delete` and `c:catAx/c:delete` by suppressing the
  deleted axis line and labels while keeping plot geometry and gridlines intact. Remaining hidden-axis work
  is interaction with crossings, tick marks, and chart-style inherited axis visibility.
- Chart boolean delete flags now treat both `val="1"` and element-only forms such as `<c:delete/>` as
  enabled for axis and legend visibility handling.
- Bar and line chart renderer now render simple legends from cached series names, using the same fill or
  stroke styles as the plotted series. Remaining legend work is exact Office layout positions, overlay
  placement beyond bar plot-box reservation, rich text, hidden/deleted entries, and chart style inheritance.
- Pie and doughnut chart renderer now render basic value data labels when `c:dLbls/c:showVal` is enabled.
  Remaining data-label work is category/percentage labels, rich text, leader lines, custom positions,
  number formats, and exact Office label collision behavior.
- Pie and doughnut chart renderer now render percentage data labels when `c:dLbls/c:showPercent` is enabled.
  Remaining pie-like label work is combined value/percentage/category labels, rich text, leader lines,
  separator handling, custom positions, and number formats.
- Pie and doughnut chart renderer now combine value and percentage labels when both `showVal` and
  `showPercent` are enabled, using a comma separator until explicit `c:separator` handling is added.
- Chart data-label options now resolve both chart-level `c:dLbls` and first-series `c:ser/c:dLbls`, covering
  another common Office emission shape before richer per-point/per-series label overrides are implemented.
- Bar and line chart plot boxes now honor simple `c:plotArea/c:layout/c:manualLayout` fractional `x/y/w/h`
  values instead of always using heuristic margins. Remaining manual-layout work is `xMode/yMode/wMode/hMode`,
  target variants, title/legend interaction, and applying the same box model to other chart families.
- Bar/column chart plot boxes now use a larger Office-backed default when the chart has no visible title
  and no visible legend. This is locked by the secondary-axis overlay probe and updates brittle unit
  assertions that encoded the older low/short plot area. Remaining chart-layout work is replacing heuristic
  title/legend margins with a typed chart-layout model.
- Bar and line chart rendering now route through a first `ChartLayout` model that owns the chart frame,
  plot box, title text, and legend layout before drawing. This preserves the current Office-backed output
  while moving chart work toward the `pptx-renderer` pattern of computing layout as data before rendering.
- Bar and line chart renderer now honor explicit `c:valAx/c:scaling` min/max values for both plotted
  geometry and value-axis tick labels. Remaining axis-scaling work is orientation, logarithmic scales,
  major/minor units, crossing rules, and date/category axis semantics.
- Bar and line chart renderer now honor explicit `c:valAx/c:majorUnit` and `c:minorUnit` for value-axis
  labels and gridline placement. Remaining unit work is automatic Office unit selection, logarithmic/date
  units, label skip rules, and exact major/minor tick mark rendering.
- Bar/column chart renderer now honor `c:varyColors` for single-series charts without explicit series
  fills, assigning palette colors by category. Remaining palette work is Office chart style/color-style
  parts, point-level overrides for bars, and theme-derived automatic colors.
- Bar/column chart renderer now honor point-level `c:dPt/c:spPr` fills, sharing the same point-style reader
  used by pie and doughnut slices. Remaining point-level bar work is point borders, inherited style defaults,
  data labels, and deleted/hidden points.
- Bar/column chart renderer now honor point-level `c:dPt/c:spPr/a:ln` borders through the shared line
  resolver. Remaining point-level bar work is inherited style defaults, data labels, deleted/hidden points,
  and exact Office border draw order for stacked/clustered variants.
- Bar/column chart renderer now render basic value data labels when `c:dLbls/c:showVal` is enabled.
  Remaining bar data-label work is label positions, stacked-series totals, rich text, number formats,
  deleted/hidden points, and collision/overflow behavior.
- Line chart renderer now render basic value data labels when `c:dLbls/c:showVal` is enabled.
  Remaining line data-label work is label positions, rich text, number formats, per-point overrides,
  deleted/hidden points, leader lines, and collision/overflow behavior.
- Combo bar charts now treat additional `c:barChart` nodes as part of the same chart text model for
  legends and value data labels, instead of drawing only the first chart's labels. This mirrors the
  `pptx-renderer` merged-option approach more closely and is locked by the public native chart synthetic
  case with a secondary-axis stacked bar chart. Remaining combo work is exact Office layout and axis
  interaction across mixed chart families.
- Combo bar charts now resolve axis visibility and strokes through each chart group's own `axId` bindings,
  following the `pptx-renderer` axis-object pattern more closely. Secondary right-axis strokes and labels are
  no longer suppressed just because the primary left value axis is deleted.

PPTX table style targets from the table composite port:

- Resolve built-in Office table style ids into first-row, whole-table, banded-row, border, and text
  formatting layers.
- Medium Style 2 built-in table styles now resolve all accent variants from the `pptx-renderer` predefined
  style map, not only Accent 1. A public synthetic lock covers Accent 6 header and banded-row fills.
- Medium Style 2 table text formatting now carries conditional first-row and first-column bold through
  the table text-run path, alongside the existing first-row light text color.
- Medium Style 2 table fills now honor first-column, last-column, and last-row conditional accent fills
  using row/column context, while explicit cell fills still override style fills.
- Light Style 1 built-in table styles now resolve all accent variants from the `pptx-renderer` predefined
  style map. First-row accent fills, first-row/first-column bold, light first-row text, and translucent
  banded-row fills are covered by the public table-style synthetic lock.
- Dark Style 1 built-in table styles now resolve all accent variants from the `pptx-renderer` predefined
  style map. Dark first-row fills, shaded accent body/band/column fills, and conditional bold/text color
  are covered by the same public table-style synthetic lock.
- Apply table style fills/borders before explicit cell overrides, matching Office draw order.
- Add a focused synthetic table style ladder before tightening `pptx-ladder-10-composite-table-port`.

Typography oracle family map:

- Fonts `oracle-pypptx-text-0001..0008`: partially covered by
  `pptx-ladder-04-typography-font-families` and the clean-port gate
  `pptx-ladder-04-font-family-port` for Times New Roman, Courier New, Georgia, Verdana, Impact,
  and Comic Sans MS.
- Sizes `0009..0015`: covered by `pptx-ladder-04-font-size-port` for isolated
  `10/14/18/24/36/48/72 pt` Office-generated text, plus older mixed-size and large-text rungs.
- Styles `0016..0021`: mostly covered by bold/italic/underline/combined public rungs; still needs a
  clean-port manifest tying each source case to the corresponding public fixture and Office metrics.
- Alignment `0022..0025`: covered by `pptx-ladder-04-color-align-port` for left, center, right, and justify.
- Colors `0026..0030`: covered by `pptx-ladder-04-color-align-port` for direct red, green, blue, orange,
  and purple RGB text.
- Mixed formatting `0031`: covered by `pptx-ladder-04-mixed-bullet-port`, alongside older mixed-run rungs.
- Bullet list `0032`: covered by `pptx-ladder-04-mixed-bullet-port` for a multi-level PowerPoint-authored
  bullet list.
- Vertical text `0033..0034`: ported as `pptx-ladder-04-vertical-text-port` and
  `pptx-ladder-04-vertical-text-270`. First-class orientation routing is present for known `a:bodyPr @vert`
  values, but stacked glyph order, clipping, column positions, and exact baseline placement remain open.
- Anchor `0035..0037`: covered by `pptx-ladder-04-anchor-port` for top, middle, and bottom shape text
  anchors.
- Line spacing `0038`: ported as `pptx-ladder-04-line-spacing-port`. It is gated as a baseline but still
  has poor structural similarity, so paragraph spacing and baseline placement remain implementation targets.

Shape adjustment oracle family map:

- Round rect `shape-adj-0001..0002`: source-aligned ports now exist as
  `pptx-ladder-06-shape-roundrect-small-port` and `pptx-ladder-06-shape-roundrect-large-port`.
  Gated baselines after the theme-fill fallback are MAE `5.954323`, changed16 `0.202635`, and MAE
  `7.136207`, changed16 `0.197371`.
- Chevron and arrow `0003..0006`: source-aligned ports now exist as
  `pptx-ladder-06-shape-chevron-shallow-port`, `pptx-ladder-06-shape-chevron-deep-port`,
  `pptx-ladder-06-shape-arrow-thin-port`, and `pptx-ladder-06-shape-arrow-wide-head-port`.
  Gated baselines after the theme-fill fallback are MAE `7.186928`, `4.102828`, `7.001761`, and
  `6.315083`. Remaining work is actual preset `avLst` adjustment support, gradient fills, shadows, and
  exact Office path geometry.
- Star, donut, cross, trapezoid, and triangle `0007..0014` plus `0021..0022`: ported in
  `pptx-ladder-06-shape-adjust-port-b`.
- Block arc, folded corner, bevel, pentagon, can, and heart `0015..0020` plus `0023..0027`: ported in
  `pptx-ladder-06-shape-adjust-port-c`.
- Moon and left brace `0028..0031`: ported in `pptx-ladder-06-shape-adjust-port-d`.

## Test Suite Performance

The public test loop is now slow enough to reduce iteration speed. The full custom console runner can take
minutes, despite `pptx-renderer` sustaining a much larger Office-oracle suite. Treat test performance as a
quality requirement, not as incidental tooling work.

High-priority actions:

- [x] Profile the dependency-free console runner and identify the slowest tests by name and elapsed time.
- [x] Add runner-level timing output so regressions in test duration are visible in normal local runs.
- [x] Cache Windows font resolver discovery across resolver instances. This reduced the full custom console
  runner from roughly 180 seconds to roughly 13 seconds on the local machine, with `108 passed, 0 failed`.
- [x] Cache immutable Office reference PDFs and rasterized oracle pages for public/private visual cases under
  ignored `artifacts/reference-cache/`, keyed by input hash, DPI, and reference/raster tool hashes.
- [x] Cache decoded PPTX image XObjects by package part and cache PDF image resource keys. The private
  `lokad-value-based` conversion-only Release run dropped from `52.474s` to `25.622s` with identical PDF
  bytes, confirming repeated image decoding and repeated full-image SHA-256 hashing were major large-deck
  anti-patterns.
- [x] Cache additional expensive test fixtures where useful, especially parsed fonts and repeated synthetic
  packages.
- [x] Keep oracle caches under ignored `artifacts/` or another ignored cache directory; never modify checked-in
  reference inputs as part of caching.
- [x] Harden Office automation scripts so chart fixture generation closes embedded chart workbooks, releases
  COM objects, and quits only automation-owned Excel instances with no remaining workbooks. This keeps Office
  reference generation usable without leaving blocking Excel/PowerPoint processes behind.
- [x] Split fast unit tests from slower visual/oracle tests so routine `dotnet run --project
  tests/Lokad.OoxPdf.Tests` stays fast and visual gates can be run explicitly.
  The custom runner now supports `--skip-slow`, `--only-slow`, and `--list`.
- [ ] Review `pptx-renderer`'s generated oracle and report reuse strategy for ideas that fit the .NET
  dependency-free constraint.

## Progress

- [x] 2026-05-25: Added the seven-part long-term architecture track to this ExecPlan without removing
  existing open progress or private-safe evidence. The track makes structural Office-PDF alignment the
  mechanism for pixel-perfect output and sets chart structural oracle tooling as the next concrete slice.
- [x] 2026-05-25: Build the first chart structural oracle tooling slice. `PdfInspect` now emits
  `graphics-operations.json` with path, clip, stroke, fill, bounds, line width, color, dash, cap, and join
  summaries; `ComparePdfGraphicsOperations.ps1` compares those structures; and `CheckVisualCase.ps1` can
  opt visual manifests into graphics-operation gates through `expected.maxGraphicsOperationBoundsDelta`.
- [x] 2026-05-25: Extend chart structural oracle tooling from generic PDF path operations to semantic chart
  structures. `ClassifyPdfChartGraphics.ps1` now classifies inspected graphics into `ClipBox`,
  `HorizontalLine`, `VerticalLine`, `FilledRegion`, `MarkerCandidate`, and derived `PlotBoxCandidate`
  records, and `CheckVisualCase.ps1` can opt manifests into chart-structure gates through
  `expected.maxChartGraphicsStructureBoundsDelta`.
- [x] 2026-05-25: Add explicit plot-order metadata to the typed chart scene model. `PptxSceneChartPlot`
  now preserves both the direct `plotArea` order and the per-kind index, and the chart renderer now
  resolves chart plot XML through direct `plotArea` child elements instead of broad `Descendants()` scans.
  This does not complete the chart model track, but it removes one structural ambiguity that matters for
  Office-aligned combo charts and future scene-owned rendering.
- [x] 2026-05-25: Improve chart-structure classification with Office-PDF evidence from several public chart
  families. The first classifier intentionally exposes raw semantic buckets; it still needs per-family
  review for pie/doughnut/radar polar plots, legend swatches, and Office's common use of clipping boxes and
  Bezier paths where naive rectangle/line classification is incomplete.
- [x] 2026-05-25: Extend chart-structure classification across multiple public chart families. Public
  Office/candidate PDFs for pie, doughnut, radar, scatter, and line-marker cases show that Office often
  exposes charts through repeated clip boxes, filled Bezier regions, and sparse axis lines rather than only
  gridline rectangles. `ClassifyPdfChartGraphics.ps1` now derives `PlotAreaClipBoxCandidate`,
  `AxisPairPlotBoxCandidate`, and `PolarPlotBoxCandidate` structures in addition to the earlier raw buckets.
- [ ] 2026-05-25: Continue chart-structure classification toward legend swatches, data-label text positions,
  and polar/radar shape semantics. The new derived candidates improve the structural oracle surface, but they
  still do not classify chart text matrices, legend entries, leader lines, or Office radar polygon strokes as
  first-class chart structures.
- [x] 2026-05-25: Add opt-in semantic chart gridline candidates to the PDF chart graphics classifier.
  `ClassifyPdfChartGraphics.ps1` now emits `HorizontalGridlineCandidate` and `VerticalGridlineCandidate`
  records for line strokes that span the derived plot box while excluding the plot-box axis edges. Existing
  raw line and plot-box structures are unchanged, so current visual gates remain stable unless a manifest
  explicitly compares the new semantic kinds; `CheckVisualCase.ps1` also forwards optional gridline
  classification tolerances for future manifests. Validation on `pptx-ladder-11-chart-column-clustered-port`
  showed Office with zero horizontal gridline candidates and the candidate with six, exposing a real
  Office/candidate structural mismatch for future renderer work instead of hiding it behind raster metrics.
- [ ] 2026-05-25: Decide the Office-aligned gridline rendering rule before gating gridline candidates. The
  public clustered-column case shows that Office does not expose candidate-like horizontal gridline strokes
  for that default chart even though the candidate currently does; the next chart renderer slice should
  inspect Office's path/clip/fill strategy and determine whether gridlines should be suppressed, restyled,
  reordered, clipped differently, or represented through another Office-like PDF structure.
- [x] 2026-05-25: Add a chart text-structure oracle slice. `ClassifyPdfChartText.ps1` now classifies inspected
  PDF text operations relative to derived chart plot boxes as `AbovePlotText`, `BelowPlotText`,
  `InsidePlotText`, `LeftAxisText`, `RightSideText`, `OuterChartText`, or generic `ChartText`; it records
  stable text hashes and effective text matrices without exposing private text content. `CheckVisualCase.ps1`
  can opt public manifests into this structural gate through `expected.maxChartTextStructurePositionDelta`.
- [x] 2026-05-25: Put the first public chart-structure gate on an existing visual case rather than leaving
  the classifier as a detached probe. `pptx-ladder-11-chart-column-clustered-port` now requires the derived
  `AxisPairPlotBoxCandidate` to stay within a bounded Office-PDF structural delta. The actual public manifest
  passed at `artifacts/visual/pptx-ladder-11-chart-column-clustered-port/20260525-003145`, comparing one
  reference and one candidate axis-pair plot box with zero reported structure deltas. This is deliberately
  narrower than the remaining fallback-geometry work: it guards one observable plot-box structure while
  leaving the broader `PptxChartMetricRules` replacement item open.
- [x] 2026-05-25: Promote chart legend labels from a generic right-side text bucket to a structural
  `LegendText` role when PDF text is supported by an adjacent marker/short-line swatch or by a right-side
  legend clipping container. The ignored line-marker probe at
  `artifacts/visual/tmp-chart-legend-gate-line-markers/20260525-003459` exercises the harness path with two
  reference and two candidate `LegendText` records. The probe intentionally uses a loose position tolerance:
  it proves semantic presence/count, while also recording that the current candidate legend placement is still
  materially lower than Office's placement.
- [x] 2026-05-25: Preserve chart series identity/order in the typed scene model. `PptxSceneChartSeries` now
  carries nullable OOXML `c:idx` and `c:order` values separately from XML list position, and the scene-builder
  fixture locks both values on bar and bubble plots. Rendering order is intentionally unchanged in this slice;
  the point is to make Office's structural series identity available before any future legend, data-label, or
  combo-chart ordering work consumes it. Validation: full console suite `204 passed, 0 failed, 0 skipped`;
  `dotnet pack src/Lokad.OoxPdf/Lokad.OoxPdf.csproj --tl:off --nologo -v minimal --no-restore` succeeded.
- [x] 2026-05-25: Preserve chart title overlay, manual layout, fill, and stroke in the typed scene model.
  `PptxSceneChartTitle` now carries `c:overlay`, `c:layout/c:manualLayout`, and `c:spPr` shape styling beside
  text and `txPr` style. The scene-builder fixture locks these fields without changing rendering behavior; the
  remaining title work is to consume the title layout structurally after public Office-PDF evidence defines the
  exact placement rule. Validation: the focused scene test passed `1 passed, 0 failed, 0 skipped`; the full
  console suite passed `204 passed, 0 failed, 0 skipped`; and `dotnet pack` succeeded.
- [x] 2026-05-25: Preserve axis titles as typed chart scene data instead of leaving them as raw axis XML.
  Each `PptxSceneChartAxis` now carries a `PptxSceneChartTitle`, reusing the same title text, overlay,
  manual-layout, shape-style, and text-style record as the chart title. The scene-builder fixture locks both
  category-axis and value-axis title metadata. Validation: the focused scene test passed
  `1 passed, 0 failed, 0 skipped`; the full console suite passed `204 passed, 0 failed, 0 skipped`; and
  `dotnet pack` succeeded.
- [x] 2026-05-25: Preserve chart legend manual layout and shape styling in the typed scene model.
  `PptxSceneChartLegend` now carries `c:layout/c:manualLayout` and `c:spPr` fill/stroke data beside position,
  overlay, visibility, and text style. The scene-builder fixture locks legend layout factors, fill, stroke
  width, and stroke color without changing rendering behavior; remaining legend placement work should consume
  this structure only after public Office-PDF evidence defines the placement rule. Validation: the focused
  scene test passed `1 passed, 0 failed, 0 skipped`; the full console suite passed
  `204 passed, 0 failed, 0 skipped`; and `dotnet pack` succeeded.
- [x] 2026-05-25: Promote chart text classification from broad position buckets to first semantic roles.
  `ClassifyPdfChartText.ps1` now emits `CategoryAxisTickLabel`, `ValueAxisTickLabel`, `DataLabelText`,
  `LegendText`, and title-candidate roles where geometry supports them, while preserving broad fallback roles
  for ambiguous text. The clustered-column public visual case now gates the four `CategoryAxisTickLabel`
  records structurally against Office PDF text positions. Validation:
  `pwsh tools/CheckVisualCase.ps1 -Case visual-cases/cases/pptx-ladder-11-chart-column-clustered-port/case.json`
  passed at `artifacts/visual/pptx-ladder-11-chart-column-clustered-port/20260525-005334`.
- [ ] 2026-05-25: Continue chart text classification beyond the first semantic roles. Remaining gaps include
  robust chart-title disambiguation, top/bottom legend containers, data labels outside the plot box,
  annotations, and multi-chart pages. Value-axis origin labels are now classified structurally, but value-axis
  gates should wait until candidate tick generation no longer emits extra tick labels against Office.
- [x] 2026-05-25: Tighten chart text classification for value-axis origin labels near plot-box edges.
  `ClassifyPdfChartText.ps1` now uses a slightly wider axis-label vertical band for text left/right of the
  plot box, so labels near the plot origin are classified as `ValueAxisTickLabel` instead of generic
  `OuterChartText`. The public clustered-column visual case still passes its existing category-axis text gate,
  and the richer classification exposes the remaining extra candidate value-axis tick as renderer work rather
  than oracle ambiguity.
- [x] 2026-05-25: Preserve per-data-label manual layout in the typed chart scene model.
  `PptxSceneChartDataLabelOverride` now carries `c:layout/c:manualLayout` alongside per-label visibility flags,
  custom text, position, separator, number format, text style, and shape style. The scene-builder fixture locks
  per-label `x/y/w/h` factors without changing renderer behavior; the remaining work is to consume label layout
  only after public Office-PDF text/label-box evidence defines the placement rule. Validation: the focused
  scene test passed `1 passed, 0 failed, 0 skipped`; the full console suite passed
  `204 passed, 0 failed, 0 skipped`; and `dotnet pack` succeeded.
- [ ] 2026-05-25: Complete the chart scene model so chart kinds, plot areas, axes, series, data labels,
  markers, title, legend, fills, strokes, and text styles are represented as typed data before PDF emission.
- [ ] 2026-05-25: Replace chart fallback geometry by turning each named `PptxChartMetricRules`
  approximation into an Office-PDF-observed rule or an explicitly classified temporary gap with a public
  visual case.
- [ ] 2026-05-25: Continue the private slide-17 typography investigation through public fixtures: use private
  evidence only to identify generic missing text behavior, then lock the behavior with synthetic
  Office-backed cases.
- [ ] 2026-05-25: Migrate PPTX text frames fully to the model-first path where style cascade, line layout,
  glyph positioning, and PDF emission are separate observable stages.
- [ ] 2026-05-25: Converge table-cell text on the common PPTX text-frame layout model so table-local height
  and vertical-alignment estimates are retired when shared layout can express the same Office behavior.
- [ ] 2026-05-25: Replace one-off OOXML enum handling with explicit ladders for touched enum families,
  including unsupported-value inventory, public fixtures where visible, and diagnostics where rendering is
  intentionally incomplete.
- [x] 2026-05-24: Preserved chart `txPr/a:defRPr` bold/italic as structural chart text-style data instead
  of a render-time heuristic. `PptxSceneChartTextStyleOverride` and the chart renderer `ChartTextStyle`
  now carry nullable bold/italic through chart defaults, axes, plot/series data labels, and per-label
  overrides; bar/line/pie data labels and axis tick labels now pass those flags into `TextRun`. Scene tests
  lock explicit true and false values, and the line/pie chart render test locks styled data-label font
  emission with a PDF font-descriptor check.
- [x] 2026-05-24: Restored the long-form ExecPlan after an over-aggressive compaction lost useful open
  planning context. Future trimming must preserve open progress items unless a duplicate or obsolete item is
  proven from checked-in code, tests, tools, or validation artifacts.
- [x] 2026-05-24: Verified the external reference renderer path exists at
  `C:\Users\JoannesVermorel\code\pptx-renderer`, and verified the current working tree still contains the
  public `PptxSyntheticCurvedConnector2LoopUsesQuarterTurnTangents` test for the slide-17 connector fix.
- [x] 2026-05-24: Audited the current PPTX render boundary. `PptxScene` and `PptxRenderContext` exist, but
  production rendering still mostly walks raw slide XML through the shape/text/table/chart partial renderers.
  This confirms that the next architecture slice should make typed scene nodes authoritative incrementally,
  not by discarding the existing render path.
- [x] 2026-05-24: Extended the scene-builder coverage so connector and chart nodes are locked alongside
  shape, picture, and table nodes. This protects the intermediate-model migration path before render dispatch
  starts consuming `PptxSceneNode` directly.
- [x] 2026-05-24: Moved ordered slide dispatch onto the scene model's node-kind classifier
  (`PptxSceneBuilder.ReadNodeKind`) so render routing no longer owns a separate XML-name/type decision tree.
  This is intentionally behavior-preserving; the renderers still consume source XML while the model boundary is
  hardened.
- [x] Add hierarchical children to `PptxSceneNode` for grouped shapes. The scene builder now records direct
  children for group nodes and the scene-builder test locks the group -> child shape hierarchy.
- [x] Move group coordinate mapping into the scene model: `PptxSceneNode` now carries a typed group transform
  with child coordinate origin/extents, scale, rotation, and flips. Ordered scene dispatch consumes this
  transform directly, and legacy XML group paths route through the same scene-builder parser.
- [x] Replace `RenderOrderedShapeTextContainer` with scene-node iteration for the ordered slide path. Normal
  PPTX rendering now builds a `PptxScene`, carries the current `PptxSceneSlide` in `PptxRenderContext`, and
  ordered slide rendering iterates `PptxSceneNode`/`Children` while leaf renderers still consume source XML.
- [x] Move the first leaf PPTX renderer inputs toward typed scene data: ordered picture rendering now takes
  relationship id and EMU bounds from `PptxSceneNode`/`PptxScenePicture`, while keeping source XML only for
  crop/recolor/svg details that still need typed models.
- [x] Resolve simple picture properties into `PptxScenePicture`: crop rectangle, stretch fill rectangle, and
  alpha are now parsed by the scene model and consumed by ordered picture rendering.
- [x] Move picture recolor intent into `PptxScenePicture`: grayscale, bi-level, luminance, and duotone OOXML
  recolor instructions are now scene-level data and ordered picture rendering converts that intent at the
  image-render boundary.
- [x] Move SVG picture placement/crop handoff onto `PptxScenePicture`: ordered SVG picture rendering now
  receives the scene-parsed crop and stretch fill rectangles instead of rereading them from source XML.
- [x] Centralize the legacy XML picture fallback on the same scene-builder readers for crop, stretch fill,
  alpha, and recolor intent. This keeps old and scene-backed picture paths structurally aligned while the
  leaf PDF/image emission boundary remains shared.
- [ ] Continue typed picture migration by moving any remaining SVG-specific paint/path decisions that belong
  to OOXML interpretation into `PptxScenePicture` or a dedicated SVG picture model.
- [x] Move the first ordered shape/connector leaf inputs toward typed scene data: `PptxSceneShape` now owns
  preset geometry names, and ordered shape/connector rendering takes preset and EMU bounds from scene nodes
  before falling back to XML for fills, strokes, effects, custom geometry, and picture fills.
- [x] Move connector line-end intent into `PptxSceneShape`: head/tail marker kind and width/length scale are
  now scene-level data and ordered straight/curved connector rendering consumes them through a renderer
  conversion boundary.
- [x] Move shape/connector line style into `PptxSceneShape`: line color, width, and alpha are now resolved in
  the scene model and consumed by ordered shape/connector rendering through a renderer conversion boundary.
- [x] Move shape/connector dash/cap/join into `PptxSceneShape`: preset dash arrays, PDF line cap, and PDF line
  join are now resolved in the scene model and consumed by ordered shape/connector rendering.
- [x] Move custom geometry presence into `PptxSceneShape`; ordered shape rendering now uses the scene model to
  decide whether the node is custom geometry before handing source XML to the existing path interpreter.
- [x] Move solid shape fill into `PptxSceneShape`: resolved fill color and alpha are now scene-level data and
  ordered shape rendering consumes them through a renderer conversion boundary.
- [x] Move supported shape pattern-fill intent into `PptxSceneShape`: diagonal pattern preset, foreground,
  background, and alpha are now scene-level data and ordered shape rendering consumes them through the existing
  pattern stroke algorithm.
- [x] Move supported shape effect intent into `PptxSceneShape`: glow and outer-shadow color, alpha, radius,
  and offset vectors are now scene-level data while rendering still uses the existing preset duplicate
  approximation.
- [x] Move shape picture-fill structure into `PptxSceneShape`: relationship id, source crop, and stretch fill
  rectangle are now scene-level data while the renderer still resolves image resources from the package at
  draw time.
- [x] Move renderable custom geometry into `PptxSceneShape`: guide formulas, path dimensions, fill/stroke
  gates, and move/line/cubic/quadratic/arc/close commands are now typed scene records consumed directly by
  ordered shape rendering.
- [x] Move preset geometry adjustment values into `PptxSceneShape`: supported `avLst` `val` guides now travel
  through the scene model and feed arc and curved-connector geometry instead of being re-read at draw time.
- [x] Add a PDF-level axial shading primitive and route supported two-stop linear shape gradients through
  `PptxSceneShape` into `/Shading` resources instead of sampled rectangle approximations.
- [x] Expand linear shape gradient coverage beyond two-stop fills: `PptxSceneShape` now keeps arbitrary
  non-alpha linear gradient stops and the PDF writer emits a Type 3 stitched function over the axial shading
  instead of flattening intermediate stops.
- [ ] Continue reducing renderer XML fallbacks by adding non-linear/path gradient support, gradient alpha
  handling, richer scene-owned effect families, and a JPEG recolor strategy without format-specific shortcuts.
- [x] 2026-05-24: Re-ran the full test suite, package, and private PPTX acceptance after multi-stop linear
  gradient support. `dotnet run --project tests/Lokad.OoxPdf.Tests --tl:off --nologo -v minimal` passed
  183/183, `dotnet pack` succeeded, and private run
  `artifacts/private-visual/lokad-value-based/20260524-115834` stayed stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, only `PPTX_UNSUPPORTED_IMAGE_RECOLOR`,
  and slide 17 MAE `2.945717`, changed16 `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Re-ran package and private PPTX acceptance after adding PDF axial shading and scene-owned
  two-stop linear shape gradients. `dotnet pack` succeeded and private run
  `artifacts/private-visual/lokad-value-based/20260524-114141` remained stable: 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, only `PPTX_UNSUPPORTED_IMAGE_RECOLOR`,
  and slide 17 MAE `2.945717`, changed16 `0.045530`, SSIM `0.917662`.
- [x] 2026-05-24: Re-ran package and private PPTX acceptance after completing the scene migration for
  shape fills, pattern fills, picture fills, effects, line styles, line ends, and renderable custom geometry.
  `dotnet pack` succeeded and private run `artifacts/private-visual/lokad-value-based/20260524-110413`
  produced 84/84 compared pages, zero dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`,
  and only one remaining diagnostic (`PPTX_UNSUPPORTED_IMAGE_RECOLOR`). Slide 17 stayed stable at MAE
  `2.945717`, changed16 `0.045530`, SSIM `0.917662`.
- [ ] JPEG recolor strategy: the remaining private warning is a three-component JPEG with duotone plus alpha.
  Do not special-case the private slide. Choose between a dependency-free JPEG pixel decoder, a principled PDF
  color-transform approach that matches Office output, or an explicit documented limitation with a public
  Office-authored rung that keeps the diagnostic stable.
- [x] 2026-05-24: Re-ran package and private PPTX acceptance after moving solid shape fill into the scene
  model. `dotnet pack` succeeded and private run
  `artifacts/private-visual/lokad-value-based/20260524-104526` produced 84/84 compared pages, zero dimension
  mismatches, deck MAE `9.043369`, changed16 `0.116418`, and one `PPTX_UNSUPPORTED_IMAGE_RECOLOR`
  diagnostic. Page 17 remained stable at MAE `2.945717`, changed16 `0.045530`.
- [x] 2026-05-24: Re-ran package and private PPTX acceptance after moving shape line style, dash/cap/join, and
  custom geometry presence into the scene model. `dotnet pack` succeeded and private run
  `artifacts/private-visual/lokad-value-based/20260524-104147` produced 84/84 compared pages, zero dimension
  mismatches, deck MAE `9.043369`, changed16 `0.116418`, and one `PPTX_UNSUPPORTED_IMAGE_RECOLOR`
  diagnostic. Page 17 remained stable at MAE `2.945717`, changed16 `0.045530`.
- [x] 2026-05-24: Re-ran package and private PPTX acceptance after scene-backed ordered rendering and typed
  picture/shape leaf inputs. `dotnet pack` succeeded and private run
  `artifacts/private-visual/lokad-value-based/20260524-101919` produced 84/84 compared pages, zero dimension
  mismatches, deck MAE `9.043369`, changed16 `0.116418`, and only one `PPTX_UNSUPPORTED_IMAGE_RECOLOR`
  diagnostic. Page 17 stayed at MAE `2.945717`, changed16 `0.045530`, confirming the architecture migration
  did not disturb the slide-17 connector fix.
- [x] 2026-05-24: Re-ran package and private PPTX acceptance after moving picture recolor intent and connector
  line ends into the scene model. `dotnet pack` succeeded and private run
  `artifacts/private-visual/lokad-value-based/20260524-102715` again produced 84/84 compared pages, zero
  dimension mismatches, deck MAE `9.043369`, changed16 `0.116418`, and one `PPTX_UNSUPPORTED_IMAGE_RECOLOR`
  diagnostic. Page 17 remained stable at MAE `2.945717`, changed16 `0.045530`.
- [ ] Trim this ExecPlan conservatively: first add missing `PLANS.md`-required sections and current evidence,
  then consolidate only completed historical detail that is already represented by checked-in fixtures,
  tests, or tool support. Do not remove open checkboxes during this cleanup unless a direct duplicate is
  found and noted.
- [x] Dependency-free `.slnx` solution, library, and CLI exist.
- [x] Dependency-free tests, visual tools, docs, public fixtures, and private validation lane exist.
- [x] NuGet package version is set to `0.1.0` for the first package.
- [x] NuGet package output is configured under ignored `artifacts/nuget/`.
- [x] OOXML package layer handles ZIP parts, content types, relationships, safe part normalization, XML
  hardening, and package size limits.
- [x] PDF writer emits deterministic static PDFs with pages and drawing operators.
- [x] PDF writer embeds TrueType/CID fonts and ToUnicode maps.
- [x] PDF writer supports JPEG passthrough, PNG image XObjects, and alpha soft masks.
- [x] CLI supports `convert input output`, `--diagnostics`, `--strict`, and exit codes `0`, `1`, `2`, and `3`.
- [x] Visual validation can export Office reference PDFs for PPTX and DOCX.
- [x] Visual validation can rasterize reference/candidate PDFs with PDFium.
- [x] Visual validation can compute PNG metrics and write comparison artifacts.
- [x] Public visual validation uses Office-exported PDFs/renders as the reference oracle; fidelity work
  should inspect the Office PDF/raster output before treating a metric as meaningful.
- [x] Initial unit-test audit started: `pptx-ladder-02-plain-text` proves the visual gate should own exact
  text fidelity while unit tests avoid freezing candidate-specific `Tm`/`Tj` structure.
- [x] DOCX public ladder started with `docx-ladder-00-blank`: Office PDF reference and candidate dimensions
  match, diagnostics are empty, MAE `0`, changed-pixel ratio threshold 16 `0`.
- [x] DOCX `docx-ladder-01-plain-paragraph` baseline is Office-aligned: reference baseline `697.42`,
  candidate baseline `697.44`, diagnostics empty, MAE `0.004605`, changed-pixel ratio threshold 16 `0.000126`.
- [x] DOCX `docx-ladder-01-line-height` exact line-height baseline stack is Office-aligned: candidate
  baselines `691.176` and `655.176`, diagnostics empty, MAE `0.014739`, changed-pixel ratio threshold 16
  `0.000394`.
- [x] DOCX `docx-ladder-01-paragraph-spacing` locks explicit auto line-height plus adjacent paragraph spacing
  collapse: candidate baselines `697.44` and `657.84`, diagnostics empty, MAE `0.014739`, changed-pixel ratio
  threshold 16 `0.000394`.
- [x] Private validation keeps inputs/manifests under ignored `private-cases/`, rejects
  tracked/private-unsafe paths, and writes ignored artifacts under `artifacts/private-visual/`.
- [x] PPTX parser/renderer supports slide order/size and solid backgrounds.
- [x] PPTX parser/renderer supports common theme colors/fonts, scheme aliases, basic scheme luminance
  transforms, and theme discovery through presentation or slide master relationships.
- [x] PPTX parser/renderer supports common master/layout inheritance and placeholder text bounds/styles.
- [x] PPTX parser/renderer supports basic rectangles, ellipses, lines, rounded rectangles, connector lines
  with triangle arrowheads, down-arrow preset shapes, and rotation/flip.
- [x] PPTX parser/renderer supports text boxes with body insets, line breaks, direct bullet characters,
  paragraph spacing, 100% default line spacing, vertical anchoring, and clipping.
- [x] PPTX parser/renderer supports mixed-run paragraph wrapping and basic styled text.
- [x] PPTX parser/renderer supports JPEG/PNG pictures and basic crop clipping.
- [x] PPTX parser/renderer supports grouped shape and picture transforms.
- [x] PPTX parser/renderer supports fixed-grid tables with fills and explicit borders.
- [x] PPTX parser/renderer supports native bar-chart rendering and unsupported-feature diagnostics.
- [x] DOCX parser/renderer supports page setup, margins, document defaults, paragraph styles, and character
  styles.
- [x] DOCX parser/renderer supports paragraphs/runs, basic styled text, greedy wrapping, simple page
  breaking, page-break-before, and exact/at-least line heights.
- [x] DOCX run-property parser honors WordprocessingML `w:onOff` element-only, `1`/`0`, `true`/`false`, and
  `on`/`off` forms for supported run styles.
- [x] DOCX parser/renderer supports bullets/decimal numbering.
- [x] DOCX parser/renderer supports inline JPEG/PNG images.
- [x] DOCX parser/renderer supports fixed-width tables in body order with explicit row heights and row-level
  page breaks.
- [x] DOCX parser/renderer supports default headers/footers and page number approximation.
- [x] DOCX parser/renderer supports unsupported-feature diagnostics.
- [x] DOCX diagnostics flag pagination-risk features that are still approximated or ignored: manual
  page/column breaks, direct and style-level keep/widow rules, style spacing variants, numbering indents,
  table styles/header rows, and paragraph section breaks.
- [x] PNG support covers non-interlaced RGB/RGBA, 8-bit grayscale, 8-bit indexed color, and packed
  low-bit-depth indexed color.
- [x] PNG support covers Adam7 interlaced RGBA images.
- [x] Unsupported PPTX and DOCX image formats now emit release-blocking `IMAGE_UNSUPPORTED_FORMAT`
  diagnostics while continuing the conversion.
- [x] PowerPoint reference export is sorted numerically so decks with more than 9 slides compare against the
  correct candidate pages.
- [x] Private PPTX assessment completed on a large 84-slide deck without exposing private contents.
- [x] Private-safe PPTX slide inventory tooling reports per-slide feature counts without exposing slide text
  or images.
- [x] Private DOCX assessment completed on an 18-candidate-page document without exposing private contents.
- [x] VisualDiff computes overlap metrics even when reference/candidate raster dimensions differ by a small
  page-rounding mismatch.

## Private Evidence

Private evidence is intentionally anonymized. Do not copy private text, screenshots, filenames, or
document-specific business content into public notes.

- Private PPTX rerun `artifacts/private-visual/lokad-value-based/20260524-191516` after the custom-geometry
  path `stroke` boolean normalization slice:
  - 84/84 pages compared with zero dimension mismatches.
  - Mean absolute error: `9.042022`; max mean absolute error: `19.097502`; mean changed-pixel ratio at
    threshold 16: `0.116405`.
  - Diagnostics remain limited to one `PPTX_UNSUPPORTED_IMAGE_RECOLOR`.
  - Page 17 remained dimension-matched at MAE `2.945717`, changed16 `0.045530`, SSIM `0.917662`.
- Private PPTX run `artifacts/private-visual/lokad-value-based/20260514-154018`:
  - 84 candidate pages, all dimensions matched reference pages.
  - Mean absolute error: `17.34937080899104`.
  - Max mean absolute error: `76.6187287808642`.
  - Mean changed-pixel ratio at threshold 16: `0.1864059055335096`.
  - Diagnostics: 9 unsupported charts, 2 unsupported interlaced PNG image occurrences.
  - Manual private inspection identified gaps in chart rendering, interlaced embedded images, dense
    image/group placement, text spacing, text-frame anchoring/clipping, and transparent overlays.
- Private PPTX conversion-only run `artifacts/private-visual/lokad-value-based/manual-20260514-165850`:
  - Conversion completed.
  - Interlaced PNG image diagnostics dropped to zero after Adam7 decoding support.
  - Remaining diagnostics: 9 unsupported charts.
- Private PPTX rerun `artifacts/private-visual/lokad-value-based/20260514-171646`:
  - 84 candidate pages, all dimensions matched reference pages.
  - Unsupported chart diagnostics dropped to zero after native bar-chart rendering.
  - Diagnostics: 9 legacy chart-rendering informational diagnostics.
  - Mean absolute error: `17.430568`; max mean absolute error: `76.618729`; mean changed-pixel ratio at
    threshold 16: `0.187419`.
- Private PPTX rerun `artifacts/private-visual/lokad-value-based/20260514-172019`:
  - 84 candidate pages, all dimensions matched reference pages.
  - Diagnostics: 9 legacy chart-rendering informational diagnostics.
  - Text-frame body insets are now honored.
  - Mean absolute error: `17.525889`; max mean absolute error: `76.657147`; mean changed-pixel ratio at
    threshold 16: `0.188083`.
- Private PPTX rerun `artifacts/private-visual/lokad-value-based/20260514-175356`:
  - 84 candidate pages, all dimensions matched reference pages.
  - Diagnostics: 9 legacy chart-rendering informational diagnostics.
  - Grouped picture transforms and text clipping are now honored.
  - Mean absolute error: `16.223412`; max mean absolute error: `75.220066`; mean changed-pixel ratio at
    threshold 16: `0.176527`.
  - Private visual inspection found the generated `output.pdf` is not acceptable on any slide. Aggregate
    pixel metrics are useful for regression tracking but are not evidence of visual correctness for this deck.
  - Working hypothesis: broad PPTX failure is caused by interacting slide-level issues, especially z-order,
    master/layout/placeholder inheritance, theme resolution, text/image/table/chart placement, and missing
    effects rather than one isolated unsupported primitive.
- Private PPTX inventory `artifacts/private-visual/lokad-value-based/inventory/20260514-182157.json`:
  - 84 slides inventoried without exposing text or images.
  - Feature counts: 72 slides with pictures, 61 with grouped content, 11 with tables, 9 with charts, 23 with
    effects, 10 with transparency, and 84 with inherited layout/master content.
  - Slide-level totals: 1846 shapes, 440 pictures, 12 table nodes, and 9 chart nodes.
  - Initial simplest-slide strategy should start with slide 1, which is structurally small but still
    exercises slide content plus inherited layout/master content.
- Private PPTX slide-1 rerun `artifacts/private-visual/lokad-value-based/20260514-182802`:
  - Theme colors now resolve for slide text after loading themes through slide-master relationships and
    handling scheme aliases.
- Private PPTX slide-1 rerun `artifacts/private-visual/lokad-value-based/20260514-185419`:
  - Direct bullet characters now render on slide text.
  - TTC font discovery and large-text baseline handling make the title/body text structurally close enough to
    expose the next slide-1 gap.
- Private PPTX slide-1 rerun `artifacts/private-visual/lokad-value-based/20260514-185955`:
  - Bullet hanging indents from `marL`/`indent` now separate bullet glyphs from bullet text.
- Private PPTX slide-1 rerun `artifacts/private-visual/lokad-value-based/20260514-190411`:
  - PPTX text now uses per-run font resources, including bullet font selection from `buFont`.
- Private PPTX slide-1 rerun `artifacts/private-visual/lokad-value-based/20260514-191256`:
  - Level-1 list-style defaults now supply missing run size/style/color/typeface and line spacing.
  - Remaining slide-1 generic gaps are fine text metrics, exact run advance/underline placement, and spacing
    precision.
- Private PPTX slide-2 rerun `artifacts/private-visual/lokad-value-based/20260514-192433`:
  - Slide background pictures now render before slide shapes, exposing later outline shapes.
  - Shape `fontRef` colors now supply fallback text color, fixing placeholder text color.
- Private PPTX slide-2 rerun `artifacts/private-visual/lokad-value-based/20260514-193029`:
  - Run-level text highlights now render as background rectangles.
  - Remaining slide-2 generic gaps include finer mixed-run text advance, title positioning, and exact
    highlight bounds.
- Private PPTX slide-2 rerun `artifacts/private-visual/lokad-value-based/20260514-193918`:
  - Centered mixed-run paragraphs now align as one paragraph instead of independently centering each run.
  - Remaining slide-2 generic gaps include exact text metrics, small footer text fit, and exact highlight
    bounds.
- Private PPTX slide-2 rerun `artifacts/private-visual/lokad-value-based/20260514-195155`:
  - Text frames now allow vertical overflow unless `bodyPr vertOverflow="clip"` is set, exposing previously
    clipped small footer text.
  - Remaining slide-2 generic gaps are fine typography/metrics: placeholder word fit, mixed-run advance,
    exact highlight bounds, and line placement.
- Private PPTX slide-2 rerun `artifacts/private-visual/lokad-value-based/20260514-195905`:
  - Mixed-run cursor advance now uses resolved font metrics instead of character-count heuristics.
  - Remaining slide-2 gaps are mostly font availability/substitution, exact text fit, and highlight/underline
    bounds.
- Private PPTX slide-3 rerun `artifacts/private-visual/lokad-value-based/20260514-200756`:
  - Connector line shapes and the down-arrow preset now render instead of being omitted or approximated as a
    rectangle.
  - 84 candidate pages, all dimensions matched reference pages.
  - Mean absolute error: `15.755963`; max mean absolute error: `75.403370`; mean changed-pixel ratio at
    threshold 16: `0.174249`.
  - Remaining slide-3 generic gaps include inherited banner fidelity, text-frame wrapping/overlap, fine text
    metrics, line/fill color transforms, and icon/image placement precision.
- Private PPTX slide-3 rerun `artifacts/private-visual/lokad-value-based/20260514-202523`:
  - Scheme luminance transforms now make inherited gray banner/ribbon fills visible.
  - Mixed-run paragraph wrapping now flows runs onto shared lines instead of wrapping each run independently.
  - 84 candidate pages, all dimensions matched reference pages.
  - Slide 3 mean absolute error: `12.545817`; changed-pixel ratio at threshold 16: `0.145434`.
  - Deck mean absolute error: `15.770431`; max mean absolute error: `75.640299`; mean changed-pixel ratio at
    threshold 16: `0.175994`.
  - Remaining slide-3 generic gaps include text autofit/fit-to-box, fine font metrics/substitution, exact
    highlight bounds, and icon/image placement precision.
- Private PPTX slide-3 rerun `artifacts/private-visual/lokad-value-based/20260514-203115`:
  - Default text line spacing now uses 100% when no `lnSpc` is specified, reducing vertical drift in dense
    text boxes.
  - 84 candidate pages, all dimensions matched reference pages.
  - Slide 3 mean absolute error: `12.359507`; changed-pixel ratio at threshold 16: `0.144392`.
  - Deck mean absolute error: `15.804445`; max mean absolute error: `75.736192`; mean changed-pixel ratio at
    threshold 16: `0.176924`.
  - Remaining slide-3 generic gaps include fine font metrics/substitution, exact highlight bounds,
    run-boundary spacing for heavily segmented text, and icon/image placement precision.
- Private PPTX slide-4 rerun `artifacts/private-visual/lokad-value-based/20260514-204901`:
  - Slide title placeholder text now inherits bounds and master title text style instead of being silently
    omitted.
  - 84 candidate pages, all dimensions matched reference pages.
  - Slide 4 mean absolute error: `5.739969`; changed-pixel ratio at threshold 16: `0.062321`.
  - Deck mean absolute error: `15.834825`; max mean absolute error: `75.834314`; mean changed-pixel ratio at
    threshold 16: `0.177074`.
  - Remaining slide-4 generic gaps include large-title font substitution/advance, exact title size, icon/text
    placement precision, and fine text wrapping.
- Private PPTX slide-5 inspection from `artifacts/private-visual/lokad-value-based/20260514-204901`:
  - Slide 5 mean absolute error: `14.77`; changed-pixel ratio at threshold 16: `0.15`.
  - The dominant generic gap is chart fidelity: the current static bar fallback lacks category/value axes,
    labels, stacked/overlay series styling, reference lines, annotations, and Office chart style colors.
  - Text blocks and separators are present but still affected by fine text metrics and run-boundary spacing.
- Private PPTX slide-6 rerun `artifacts/private-visual/lokad-value-based/20260514-210051`:
  - PPTX table rendering now strokes only explicit visible cell borders instead of inventing black borders
    for every cell.
  - 84 candidate pages, all dimensions matched reference pages.
  - Slide 6 mean absolute error: `12.067719`; changed-pixel ratio at threshold 16: `0.137847`.
  - Deck mean absolute error: `15.503991`; max mean absolute error: `75.834314`; mean changed-pixel ratio at
    threshold 16: `0.174401`.
  - Remaining slide-6 generic gaps include table text wrapping/positioning, text color/style inheritance,
    title font metrics, and image/text placement precision.
- Private PPTX slide-7 rerun `artifacts/private-visual/lokad-value-based/20260514-211029`:
  - Connector line triangle arrowheads now render on the expected endpoints.
  - 84 candidate pages, all dimensions matched reference pages.
  - Slide 7 mean absolute error: `9.230441`; changed-pixel ratio at threshold 16: `0.102554`.
  - Deck mean absolute error: `15.502189`; max mean absolute error: `75.834314`; mean changed-pixel ratio at
    threshold 16: `0.174386`.
  - Remaining slide-7 generic gaps include custom geometry/freeform curve paths, exact connector styling,
    large-title text metrics, and dense paragraph spacing.
- Private PPTX slide-8 inspection from `artifacts/private-visual/lokad-value-based/20260514-211029`:
  - The dominant generic gap is grouped table-like layout fidelity: row labels, icons, red separators, and
    text blocks are vertically compressed or overlapped.
  - Left-side callout arrows/text are misplaced, and fine text wrapping/spacing remains weak.
  - This should be handled after a focused audit of nested group transforms, placeholder-derived text styles,
    and table-like grouped shape ordering.
- Private PPTX slide-9 rerun `artifacts/private-visual/lokad-value-based/20260514-211655`:
  - Rounded rectangle preset shapes now render with rounded corners instead of rectangular outlines.
  - 84 candidate pages, all dimensions matched reference pages.
  - Slide 9 mean absolute error: `18.658330`; changed-pixel ratio at threshold 16: `0.209670`.
  - Deck mean absolute error: `15.495442`; max mean absolute error: `75.801014`; mean changed-pixel ratio at
    threshold 16: `0.174317`.
  - Remaining slide-9 generic gaps include rotated text labels, curved connectors, exact line/shape
    placement, and fine text metrics/wrapping.
- Public PPTX ladder reruns:
  - `pptx-blank` at `artifacts/visual/pptx-blank/20260514-213058`: page count and dimensions matched,
    diagnostics were empty, MAE `0`, changed-pixel ratio threshold 16 `0`.
  - `pptx-ladder-01-solid-background` at `artifacts/visual/pptx-ladder-01-solid-background/20260514-213058`:
    page count and dimensions matched, diagnostics were empty, MAE `0`, changed-pixel ratio threshold 16 `0`.
  - `pptx-ladder-01-master-background` at
    `artifacts/visual/pptx-ladder-01-master-background/20260514-213144`: page count and dimensions matched,
    diagnostics were empty, MAE `0`, changed-pixel ratio threshold 16 `0`.
  - The visual harness now enforces optional manifest gates for page count, dimensions, maximum MAE, maximum
    changed-pixel ratio, and empty diagnostics.
- Public PPTX ladder text rerun:
  - `pptx-ladder-02-plain-text` at `artifacts/visual/pptx-ladder-02-plain-text/20260514-213736`: page count
    and dimensions matched, diagnostics were empty, MAE `0.043046`, changed-pixel ratio threshold 16
    `0.000511`.
  - First-line PPTX text baselines now use a lower Office-aligned offset for plain top-anchored text boxes.
- Public PPTX ladder text-flow rerun:
  - `pptx-ladder-03-text-flow` at `artifacts/visual/pptx-ladder-03-text-flow/20260514-214108`: page count and
    dimensions matched, diagnostics were empty, MAE `0.414685`, changed-pixel ratio threshold 16 `0.003763`.
  - Single-run centered paragraphs now receive the same alignment offset as mixed-run centered paragraphs.
  - `pptx-ladder-03-text-anchor-overflow` at
    `artifacts/visual/pptx-ladder-03-text-anchor-overflow/20260514-214416`: page count and dimensions
    matched, diagnostics were empty, MAE `0.413942`, changed-pixel ratio threshold 16 `0.003552`.
  - Clipped PPTX text frames now suppress lines that cannot fit inside the clip box.
- Private PPTX acceptance rerun `artifacts/private-visual/lokad-value-based/20260514-223939`:
  - 84 candidate pages, all dimensions matched reference pages.
  - Diagnostics: 9 legacy chart-rendering informational diagnostics.
  - Slide 1 mean absolute error: `18.001507`; changed-pixel ratio at threshold 16: `0.147195`.
  - Slide-1 public-safe gaps map to styled text/list fixtures: large centered serif title placement, white
    text over dark image background, underlined run bounds, and multi-paragraph bullet list wrapping/line
    spacing.
- Public PPTX ladder styled-text rerun:
  - `pptx-ladder-04-bullet-wrap` at `artifacts/visual/pptx-ladder-04-bullet-wrap/20260514-224326`: page count
    and dimensions matched, diagnostics were empty, MAE `1.253179`, changed-pixel ratio threshold 16
    `0.009518`.
  - Bullet glyph placement, hanging indents, and continuation-line alignment are now covered by a public
    synthetic fixture before using private slide-1 evidence.
  - `pptx-ladder-04-serif-title-underline` at
    `artifacts/visual/pptx-ladder-04-serif-title-underline/20260514-224715`: page count and dimensions
    matched, diagnostics were empty, MAE `0.481747`, changed-pixel ratio threshold 16 `0.005513`.
  - Large centered serif text over a dark background and underlined mixed-run text are now covered by a
    public synthetic fixture before using private slide-1 evidence.
  - `pptx-ladder-04-mixed-paragraph-stack` at
    `artifacts/visual/pptx-ladder-04-mixed-paragraph-stack/20260514-225350`: page count and dimensions
    matched, diagnostics were empty, MAE `3.041919`, changed-pixel ratio threshold 16 `0.027087`.
  - This fixture is intentionally not gated yet. It exposes remaining Ladder 4 gaps in multi-paragraph
    vertical rhythm, mixed large/small text stacks, and exact serif text metrics; a broad default
    line-spacing change was rejected because it regressed existing text tests and worsened this fixture.
  - `pptx-ladder-04-paragraph-advance` at
    `artifacts/visual/pptx-ladder-04-paragraph-advance/20260514-230129`: page count and dimensions matched,
    diagnostics were empty, MAE `0.340723`, changed-pixel ratio threshold 16 `0.003961`.
  - Consecutive paragraph advance now uses an Office-like default without changing intra-paragraph line
    breaks/wraps or explicit `lnSpc`; the new fixture is gated, and `pptx-ladder-04-bullet-wrap` was
    tightened after improving to MAE `0.740986`, changed-pixel ratio threshold 16 `0.006748`.
  - `pptx-ladder-04-empty-paragraph-gap` at
    `artifacts/visual/pptx-ladder-04-empty-paragraph-gap/20260514-231355`: page count and dimensions matched,
    diagnostics were empty, MAE `2.109550`, changed-pixel ratio threshold 16 `0.017209`.
  - Formatting-only empty paragraphs now consume vertical advance using their own paragraph/run formatting or
    the paragraph default, without borrowing a preceding large-title font.
  - `pptx-ladder-04-bold-italic-face` at `artifacts/visual/pptx-ladder-04-bold-italic-face/20260514-233419`:
    page count and dimensions matched, diagnostics were empty, MAE `2.512482`, changed-pixel ratio threshold
    16 `0.018524`.
  - Font resolution now selects bold/italic faces when available, and synthetic bold/italic is applied only
    when the requested face cannot be resolved.
  - Split font-face anchors:
    - `pptx-ladder-04-bold-face-single` at
      `artifacts/visual/pptx-ladder-04-bold-face-single/20260514-234212`: MAE `0.229329`, changed-pixel ratio
      threshold 16 `0.002122`.
    - `pptx-ladder-04-italic-face-single` at
      `artifacts/visual/pptx-ladder-04-italic-face-single/20260514-234221`: MAE `0.259048`, changed-pixel
      ratio threshold 16 `0.002407`.
    - `pptx-ladder-04-bold-italic-face-single` at
      `artifacts/visual/pptx-ladder-04-bold-italic-face-single/20260514-234229`: MAE `0.253966`,
      changed-pixel ratio threshold 16 `0.002414`.
  - `pptx-ladder-04-character-spacing` at
    `artifacts/visual/pptx-ladder-04-character-spacing/20260514-235026`: page count and dimensions matched,
    diagnostics were empty, MAE `0.920131`, changed-pixel ratio threshold 16 `0.006839`.
  - Run-level `spc` character spacing now affects text advance, wrapping, PDF text state, underline/highlight
    extents, and fallback measurement.
  - `pptx-ladder-04-baseline-shift` at `artifacts/visual/pptx-ladder-04-baseline-shift/20260514-235445`: page
    count and dimensions matched, diagnostics were empty, MAE `0.269687`, changed-pixel ratio threshold 16
    `0.002153`.
  - Run-level `baseline` now shifts superscript/subscript text, highlights, and underlines relative to the
    paragraph baseline.
  - `pptx-ladder-04-highlight-single` at `artifacts/visual/pptx-ladder-04-highlight-single/20260514-235728`:
    page count and dimensions matched, diagnostics were empty, MAE `0.261774`, changed-pixel ratio threshold
    16 `0.004202`.
  - Run-level highlight rendering is now locked by a public visual gate for a single highlighted text run.
  - `pptx-ladder-04-line-spacing-points` at
    `artifacts/visual/pptx-ladder-04-line-spacing-points/20260515-000244`: page count and dimensions matched,
    diagnostics were empty, MAE `0.494735`, changed-pixel ratio threshold 16 `0.004269`.
  - Absolute PPTX line spacing (`a:lnSpc/a:spcPts`) now drives intra-paragraph line breaks and explicit
    paragraph advance in the public ladder.
  - `pptx-ladder-04-bullet-style` at `artifacts/visual/pptx-ladder-04-bullet-style/20260515-181539`: page
    count and dimensions matched, diagnostics were empty, MAE `0.029019`, changed-pixel ratio threshold 16
    `0.000670`.
  - Bullet-specific color and point-size formatting (`a:buClr`, `a:buSzPts`) is order-aware: valid properties
    before the bullet marker are honored, while invalid late properties after `a:buChar` match Office's
    ignored behavior in the current public visual fixture.
  - `pptx-ladder-04-tab-stop` at `artifacts/visual/pptx-ladder-04-tab-stop/20260515-174515`: page count and
    dimensions matched, diagnostics were empty, MAE `0.011007`, changed-pixel ratio threshold 16 `0.000341`.
  - Office ignores the standalone synthetic `<a:tab/>` child used by this fixture and emits continuous text,
    so the renderer now matches that behavior. A separate Office-authored fixture is still needed for real
    tab characters/tab-stop semantics.
  - `pptx-ladder-04-strikethrough-single` at
    `artifacts/visual/pptx-ladder-04-strikethrough-single/20260515-001236`: page count and dimensions
    matched, diagnostics were empty, MAE `0.268133`, changed-pixel ratio threshold 16 `0.002249`.
  - Single-run strikethrough (`a:rPr @strike`) now renders and is locked by a public visual gate.
  - `pptx-ladder-04-all-caps` at `artifacts/visual/pptx-ladder-04-all-caps/20260515-001508`: page count and
    dimensions matched, diagnostics were empty, MAE `0.182496`, changed-pixel ratio threshold 16 `0.001935`.
  - Run-level all-caps text (`a:rPr @cap="all"`) now transforms text before measurement and drawing in the
    public ladder.
  - `pptx-ladder-03-preserved-spaces` at `artifacts/visual/pptx-ladder-03-preserved-spaces/20260515-075831`:
    page count and dimensions matched, diagnostics were empty, MAE `0.139198`, changed-pixel ratio threshold
    16 `0.001193`.
  - Preserved spaces in PPTX runs are visually locked by a public Office-PDF-backed gate. Office emits this
    case as one `TJ` text object, while the candidate still emits multiple `Tj` objects split around space
    groups; track this as a PDF-structure improvement, not a blocker for the visual rung.
  - `pptx-ladder-04-underline-single` at `artifacts/visual/pptx-ladder-04-underline-single/20260515-080227`:
    page count and dimensions matched, diagnostics were empty, MAE `0.151413`, changed-pixel ratio threshold
    16 `0.001580`.
  - PPTX underlines now use an Office-like filled rectangle instead of a stroked line; the older broad
    serif-title underline gate also improved to MAE `0.429467`, changed-pixel ratio threshold 16 `0.005498`.
  - `pptx-ladder-04-mixed-font-size-line` at
    `artifacts/visual/pptx-ladder-04-mixed-font-size-line/20260515-080529`: page count and dimensions
    matched, diagnostics were empty, MAE `0.195173`, changed-pixel ratio threshold 16 `0.001995`.
  - Same-line mixed font sizes now have their own public Office-PDF-backed gate. Office and candidate both
    place 36pt and 18pt runs on a shared baseline with closely matching x advances.
  - `pptx-ladder-04-mixed-font-size-stack` now isolates large/small/large paragraph vertical rhythm. Office
    advances line tops and recomputes each paragraph baseline from the current paragraph font size; the
    renderer now follows that model. The isolated fixture is gated at MAE `0.065659`, changed-pixel ratio
    threshold 16 `0.001726`.
  - The line-top text layout change tightened `pptx-ladder-04-mixed-font-size-line` to MAE `0.007254`,
    changed-pixel ratio threshold 16 `0.000173`, and `pptx-ladder-04-paragraph-advance` to MAE `0.216017`,
    changed-pixel ratio threshold 16 `0.003088`.
  - The same change improved the broader ungated `pptx-ladder-04-mixed-paragraph-stack` from MAE `3.041919`
    to MAE `0.873692`, changed-pixel ratio threshold 16 `0.011333`. Remaining gaps are now concentrated in
    serif font metrics, underline bounds in combined text, and wrapped bullet continuation rather than gross
    paragraph vertical drift.
  - Default intra-paragraph wrap/line-break advance now uses the same Office-like `1.2 * fontSize` line-top
    advance as paragraph advance, while explicit line spacing remains explicit. This tightened
    `pptx-ladder-04-bullet-wrap` to MAE `0.149908`, changed-pixel ratio threshold 16 `0.003331`.
  - Absolute PPTX line-spacing baselines now place text lower inside the fixed line box, matching the Office
    PDF stream for `a:spcPts`. This tightened `pptx-ladder-04-line-spacing-points` to MAE `0.305809`,
    changed-pixel ratio threshold 16 `0.002707`.
  - Adjacent underlined PPTX flow segments now coalesce before PDF emission so multi-word underlined runs
    draw one continuous underline rectangle, matching Office's continuous underline shape.
    `pptx-ladder-04-underline-single` tightened to MAE `0.118922`, changed-pixel ratio threshold 16
    `0.001273`; `pptx-ladder-04-serif-title-underline` is now gated at MAE `0.353182`, changed-pixel ratio
    threshold 16 `0.004945`.
  - Classic OpenType `kern` pairs now affect PPTX text measurement and PDF emission through `TJ` arrays when
    available, moving the renderer closer to Office's pair-adjusted text streams. This improved
    `pptx-ladder-04-mixed-paragraph-stack` to MAE `0.520333`, changed-pixel ratio threshold 16 `0.009028`.
  - Adjacent same-style PPTX text runs on the same baseline now coalesce before font resolution and PDF
    emission, moving line text closer to Office's grouped `TJ` strategy. This tightened
    `pptx-ladder-03-preserved-spaces` from MAE `0.139198`, changed-pixel ratio threshold 16 `0.001193` to MAE
    `0.001662`, changed-pixel ratio threshold 16 `0`, and tightened `pptx-ladder-04-bullet-wrap` from MAE
    `0.149908`, changed-pixel ratio threshold 16 `0.003331` to MAE `0.137307`, changed-pixel ratio threshold
    16 `0.003059`.
  - `pptx-ladder-07-image-crop` now locks a minimal left/right cropped PNG picture against Office at exact
    raster parity: page count and dimensions matched, diagnostics were empty, MAE `0`, changed-pixel ratio
    threshold 16 `0`.
  - `pptx-ladder-07-image-alpha` now locks a minimal transparent PNG picture over a solid slide background.
    Office and candidate raster output match exactly at MAE `0`, changed-pixel ratio threshold 16 `0`; the
    candidate PDF carries the alpha channel as a soft mask image.
  - `pptx-ladder-07-jpeg-image` now locks minimal JPEG picture placement with no diagnostics at MAE
    `0.134097`, changed-pixel ratio threshold 16 `0.005486`. The remaining delta is JPEG decode/re-encode
    edge color variance between Office and direct PDF embedding, not placement.
  - `pptx-ladder-07-image-rotation` now locks a rotated rectangular PNG picture. Pictures now reuse the
    Office-aligned shape transform path for rotation/flips before image emission, tightening this case from
    MAE `5.737500`, changed-pixel ratio threshold 16 `0.056250` to exact raster parity.
  - `pptx-ladder-07-image-flip` now locks a horizontally flipped PNG picture with exact Office raster parity,
    covering the non-rotated flip branch of picture transforms.
  - `pptx-ladder-07-image-rotate-flip` now locks the combined rotated plus horizontally flipped picture
    matrix order with exact Office raster parity.
  - `pptx-ladder-08-grouped-picture` now locks grouped picture child coordinate scaling against Office with
    exact raster parity, complementing the existing grouped-shape visual rung.
  - `pptx-ladder-08-grouped-text` now locks text boxes inside grouped content. Text run layout now applies
    ancestor group transforms before computing text bounds and clip regions, improving this fixture from MAE
    `0.377745`, changed-pixel ratio threshold 16 `0.002546` to MAE `0.002687`, changed-pixel ratio threshold
    16 `0.000079`.
  - `pptx-ladder-08-nested-grouped-text` now locks nested group transform composition for text boxes at MAE
    `0.002687`, changed-pixel ratio threshold 16 `0.000079`.
  - `pptx-ladder-08-text-shape-zorder` now locks simple text/shape sibling order: a later opaque shape covers
    an earlier text box at exact raster parity. Simple slides without pictures or graphic frames now render
    shapes and their text in shape-tree order instead of in separate all-shapes/all-text layers.
  - `pptx-ladder-08-shape-picture-zorder` now locks simple shape/picture sibling order at exact raster
    parity. The ordered slide path now includes pictures while preserving the exact image crop rung.
  - `pptx-ladder-06-dashed-connector` now locks the Office `dash` preset for connectors. Office emits a `[4w
    3w]` dash array for this case; the renderer now maps `a:prstDash val="dash"` accordingly and resets the
    dash state after stroking, improving the case from MAE `0.180208`, changed-pixel ratio threshold 16
    `0.001736` to exact raster parity.
  - `pptx-ladder-06-dash-dot-connector` now locks the Office `dashDot` preset for connectors. Office emits a
    `[4w 3w 1w 3w]` dash array; the PDF graphics builder now supports arbitrary dash arrays, improving the
    case from MAE `0.229766`, changed-pixel ratio threshold 16 `0.002214` to exact raster parity while
    preserving the plain dashed connector gate.
  - `pptx-ladder-06-round-cap-connector` now locks `a:ln cap="rnd"` against Office. Office emits PDF round
    cap/join operators (`1 J 1 j`), and the renderer now does the same for round-capped lines, improving the
    case from MAE `0.069703`, changed-pixel ratio threshold 16 `0.000505` to exact raster parity.
  - `pptx-ladder-06-square-cap-connector` now locks `a:ln cap="sq"` against Office. Office emits `2 J 1 j`,
    and the renderer now maps square caps to the PDF projecting-square line cap, improving the case from MAE
    `0.089531`, changed-pixel ratio threshold 16 `0.000625` to exact raster parity.
  - `pptx-ladder-06-connector-arrow` now locks a straight connector with triangle tail arrowhead and a
    down-arrow preset at exact raster parity. Arrowed connectors now use Office-like filled line/triangle
    geometry, and the down-arrow preset now uses Office's default shoulder proportion.
  - `pptx-ladder-06-more-polygons` now locks right triangle, pentagon, hexagon, and octagon preset geometry
    with exact raster parity against Office-exported PDF paths.
  - `pptx-ladder-06-double-arrows` now locks left-right and up-down block arrow presets with exact raster
    parity against Office-exported PDF paths.
  - `pptx-ladder-06-symbol-polygons` now locks the plus/cross preset polygon with exact raster parity against
    Office-exported PDF paths.
  - `pptx-ladder-06-rect-callout` now locks the `wedgeRectCallout` preset polygon with exact raster parity
    and no unsupported-callout diagnostic; other callout presets still emit the diagnostic until individually
    supported.
  - Unsupported PPTX custom geometry (`a:custGeom`) and preset callout shapes now emit slide-scoped
    diagnostics instead of silently falling back to ordinary rectangles.
  - PPTX custom geometry now supports `a:arcTo` paths with common guide formulas (`val`, `+-`, `*/`,
    `abs`, `min`, `max`, `?:`, `sin`, `cos`). Public synthetic unit
    `PptxSyntheticCustomGeometryArcPathRendersCurve` locks the behavior, and the private deck no longer
    emits `PPTX_UNSUPPORTED_CUSTOM_GEOMETRY`.
  - `pptx-ladder-10-explicit-borders` now locks a minimal 2x2 table with per-edge borders. Explicit table
    borders are collected, coalesced, and stroked after cell fills to match Office's PDF order and avoid
    double-stroking shared edges, improving the case from MAE `0.445408`, changed-pixel ratio threshold 16
    `0.002479` to exact raster parity.
  - `pptx-ladder-10-vertical-align` now locks top, center, and bottom table-cell anchors through `a:tcPr
    anchor`. Center and bottom anchored cells shift by Office's table text-area height, improving the case
    from MAE `0.509194`, changed-pixel ratio threshold 16 `0.005176` to MAE `0.013784`, changed-pixel ratio
    threshold 16 `0.000302`.
  - `pptx-ladder-10-unstyled-grid` now locks the default grid style for tables without a table style id. The
    renderer now uses black unstyled grid lines on row boundaries while preserving the existing styled-table
    white grid behavior, improving this fixture from MAE `0.467088`, changed-pixel ratio threshold 16
    `0.002770` to exact raster parity.
  - `pptx-ladder-10-horizontal-merge` now locks a minimal horizontally merged table cell with exact raster
    parity. Table layout now honors `a:tc @gridSpan` and skips `hMerge`/`vMerge` continuations, and default
    unstyled grids suppress the internal vertical segment covered by a horizontal merge.
  - `pptx-ladder-10-vertical-merge` now locks a minimal vertically merged table cell with exact raster
    parity. Table layout now honors `a:tc @rowSpan`, and default unstyled grids suppress the internal
    horizontal segment covered by a vertical merge.
- Private PPTX rerun `artifacts/private-visual/lokad-value-based/20260514-232256`:
  - 84 candidate pages, all dimensions matched reference pages.
  - Diagnostics: 9 legacy chart-rendering informational diagnostics.
  - Slide 1 mean absolute error: `15.366243`; changed-pixel ratio at threshold 16: `0.130552`.
  - The formatted-empty-paragraph fix materially improved slide-1 title/body separation. Remaining slide-1
    generic gaps are fine font metrics, exact title/body baseline placement, underline bounds, and dense
    bullet/list wrapping precision.
- Private PPTX rerun `artifacts/private-visual/lokad-value-based/20260516-090605`:
  - 84 visual comparison entries were produced for local assessment.
  - Mean absolute error average: `13.278538`; max mean absolute error: `30.864359`.
  - Mean changed-pixel ratio at threshold 16: `0.155895`; max changed-pixel ratio: `0.468661`.
  - Diagnostics remain public-safe feature categories: chart rendering gaps plus unsupported effects,
    custom geometry, and transparency.
  - This run is assessment evidence only; the implementation track remains public bottom-up PPTX typography
    and feature fixtures before private-slide tuning.
- Private PPTX slide 2/3 feature survey from
  `artifacts/private-visual/lokad-value-based/20260516-110649`:
  - 84 visual comparison entries were produced; all paired page dimensions matched.
  - Slide 2 metrics: mean absolute error `6.866764`; RMSE `21.429890`; changed-pixel ratio threshold 16
    `0.109705`; changed-pixel ratio threshold 32 `0.034132`.
  - Slide 2 has no unsupported-feature diagnostics. Generic gaps are text style inheritance for
    `fontRef`/theme colors when runs lack direct fills, text frame vertical anchoring/insets, highlighted
    mixed-run geometry, and Cambria/Cambria Math advance and baseline fidelity.
  - Slide 3 metrics: mean absolute error `11.521155`; RMSE `43.489984`; changed-pixel ratio threshold 16
    `0.135031`; changed-pixel ratio threshold 32 `0.117801`.
  - Slide 3 emits `PPTX_UNSUPPORTED_TEXT_AUTOFIT` for text autofit. Generic gaps are `spAutoFit`
    shrink-to-fit behavior, square-wrapped overflow text frames, grouped picture-plus-caption positioning,
    centered/italic label metrics, highlighted headline geometry, and dense wrapped text layout.
  - Bottom-up response: add or tighten public fixtures for these capabilities before attempting private
    slide-specific tuning.
- Public PPTX slide-2 bottom-up rungs:
  - `pptx-ladder-04-fontref-centered-text` locks a no-fill text shape with visible line, centered
    Cambria/Cambria Math text, and `fontRef`/theme color inheritance when the run has no direct fill.
    Latest gated run: `artifacts/visual/pptx-ladder-04-fontref-centered-text/20260516-113529`, MAE
    `0.085627`, changed-pixel ratio threshold 16 `0.001779`.
  - `pptx-ladder-04-cambria-highlight-center` locks centered mixed-run Cambria/Cambria Math text with
    a highlighted middle run and Office-authored highlight geometry. Latest gated run:
    `artifacts/visual/pptx-ladder-04-cambria-highlight-center/20260516-113529`, MAE `0.291465`,
    changed-pixel ratio threshold 16 `0.004212`.
  - `pptx-ladder-04-cambria-highlight-footer` locks footer-sized centered mixed-run Cambria/Cambria Math
    text with a highlighted middle run over a dark background. Latest gated run:
    `artifacts/visual/pptx-ladder-04-cambria-highlight-footer/20260516-114746`, MAE `0.021215`,
    changed-pixel ratio threshold 16 `0.000521`.
  - `pptx-ladder-04-small-centered-highlight-box` locks a small vertically centered text frame with
    explicit Office-style insets, a visible line, and highlighted mixed-run text. Latest gated run:
    `artifacts/visual/pptx-ladder-04-small-centered-highlight-box/20260516-115708`, MAE `0.202937`,
    changed-pixel ratio threshold 16 `0.002488`.
- Public PPTX slide-3 bottom-up rungs:
  - `pptx-ladder-04-spautofit-overflow` captures a minimal `spAutoFit` overflow case. Latest run:
    `artifacts/visual/pptx-ladder-04-spautofit-overflow/20260516-114453`, MAE `0.224563`, changed-pixel
    ratio threshold 16 `0.003963`. Office keeps the text inside the box on three lines; the candidate now
    matches the three-line wrap after suppressing draw-time double wrapping. Plain `spAutoFit` is no longer
    diagnosed as unsupported because Office persists the grown text-frame geometry in the package; shrink
    and scale autofit through `normAutofit` remains diagnostic-only.
  - `pptx-ladder-08-grouped-picture-caption` locks a grouped picture plus centered italic caption cell.
    Latest gated run: `artifacts/visual/pptx-ladder-08-grouped-picture-caption/20260516-114125`, MAE
    `0.124819`, changed-pixel ratio threshold 16 `0.002018`.
  - `pptx-ladder-04-square-wrap-mixed-small` captures a 10.5/12 pt square-wrapped mixed-run text frame.
    Latest gated run: `artifacts/visual/pptx-ladder-04-square-wrap-mixed-small/20260516-120526`, MAE
    `0.771472`, changed-pixel ratio threshold 16 `0.008302`. Hyphen-aware flow segmentation now matches
    Office's break after the hyphenated prefix, and wrapped-line advance now uses the paragraph's actual
    font size instead of an 18 pt floor. Remaining drift is mixed bold/italic metrics and exact baseline
    placement.
  - `pptx-ladder-04-square-wrap-small-plain` locks the same compact square-wrapped small-text behavior
    without mixed styles. Latest gated run:
    `artifacts/visual/pptx-ladder-04-square-wrap-small-plain/20260516-120434`, MAE `0.202193`,
    changed-pixel ratio threshold 16 `0.003082`.
  - `pptx-ladder-04-highlighted-headline-runs` locks a headline-sized mixed-run line with two highlighted
    spans and a wrapped continuation line. Latest gated run:
    `artifacts/visual/pptx-ladder-04-highlighted-headline-runs/20260516-115932`, MAE `0.292685`,
    changed-pixel ratio threshold 16 `0.006998`.
  - `pptx-ladder-04-vertical-text-270` captures a minimal vertical Latin text frame. Latest run:
    `artifacts/visual/pptx-ladder-04-vertical-text-270/20260516-124317`, MAE `0.368511`, changed-pixel
    ratio threshold 16 `0.003190`, with `PPTX_UNSUPPORTED_TEXT_ORIENTATION` still expected. Office stacks
    and rotates glyphs inside the vertical frame; the current renderer still lays the text horizontally.
- Private PPTX rerun `artifacts/private-visual/lokad-value-based/20260516-121040` after the slide-2/3
  public ladder work:
  - 84 visual comparison entries were produced; all paired page dimensions matched.
  - Deck mean absolute error average: `12.734628`; max mean absolute error: `31.067204`.
  - Deck mean changed-pixel ratio at threshold 16: `0.151359`; max changed-pixel ratio: `0.467519`.
  - Slide 2 metrics improved to MAE `6.560317`, RMSE `20.121548`, changed16 `0.106846`,
    changed32 `0.031196`.
  - Slide 3 metrics improved to MAE `10.224075`, RMSE `40.503589`, changed16 `0.121273`,
    changed32 `0.104931`.
  - The previous slide-3 text-autofit diagnostic is gone. Remaining diagnostics are public-safe categories:
    chart rendering gaps plus unsupported effects, custom geometry, vertical text, transparency,
    multi-column text, and image recolor.
- Private PPTX rerun `artifacts/private-visual/lokad-value-based/20260516-123248` after restricting GPOS
  pair positioning to the active `kern` feature:
  - 84 visual comparison entries were produced; all paired page dimensions matched.
  - Deck mean absolute error average: `12.734628`; max mean absolute error: `31.067204`.
  - Deck mean changed-pixel ratio at threshold 16: `0.151359`; max changed-pixel ratio: `0.467519`.
  - Aggregate metrics are effectively unchanged; this run is typography evidence for manual inspection of
    the reported parasite inter-letter gaps, not a deck-level visual-correctness milestone.
- Private DOCX run `artifacts/private-visual/user-requirements-spec/20260514-164847`:
  - Reference output had 16 pages; candidate output had 18 pages.
  - Candidate page height differed by 1 raster pixel from reference at 144 DPI, preventing pixel metrics.
  - Diagnostics were empty.
  - This identifies DOCX pagination/page geometry fidelity as a high-priority gap, especially because no
    diagnostic currently explains the mismatch.
- Private DOCX rerun `artifacts/private-visual/user-requirements-spec/20260514-171244`:
  - Reference output had 16 pages; candidate output had 16 pages after preserving DOCX table body order and
    reducing default table row height.
  - Candidate page height still differs by 1 raster pixel at 144 DPI.
  - Overlap metrics are now available despite the dimension mismatch; paired-page mean absolute error was
    `19.89926`, and mean changed-pixel ratio at threshold 16 was `0.16322`.
  - Diagnostics were empty, so remaining pagination gaps still need explicit diagnostics or rendering fixes.
- Private DOCX rerun `artifacts/private-visual/user-requirements-spec/20260514-173117`:
  - Reference output had 16 pages; candidate output had 16 pages.
  - All 16 rasterized page dimensions matched after A4 media-box normalization.
  - Mean absolute error: `19.88376`; mean changed-pixel ratio at threshold 16: `0.163965`.
  - Diagnostics were empty.
- Private DOCX rerun `artifacts/private-visual/user-requirements-spec/20260514-175910`:
  - Reference output had 16 pages; candidate output had 17 pages.
  - Candidate has one extra page; paired-page mean absolute error was `19.226760`, and mean changed-pixel
    ratio at threshold 16 was `0.158228`.
  - Diagnostics were empty.
  - This keeps DOCX pagination fidelity as the top active risk; attempted manual-break and keep-with-next
    heuristics were reverted because they did not resolve the private page-count mismatch.
  - Anonymized structure survey found no direct manual page/column breaks, no direct paragraph keep rules, no
    direct paragraph spacing, no inline/anchored drawings, and one section.
  - The same survey found 198 body paragraphs, 13 body tables, 129 table rows, 422 table cells, 45 numbered
    paragraphs, 24 style-level spacing definitions, 30 style-level keep-rule definitions, 36 numbering levels
    with indents, 13 table preferred widths, 422 cell widths, 18 cell vertical-alignment declarations, and 1
    repeating table-header row.
  - Working hypothesis: the 17th candidate page is driven by accumulated small layout errors, especially
    style-derived paragraph spacing/keep rules, numbering indents/hanging indents, and table sizing/header
    behavior rather than an explicit page-break feature.
- Private DOCX rerun `artifacts/private-visual/user-requirements-spec/20260514-180723`:
  - Reference output had 16 pages; candidate output had 17 pages.
  - Candidate still has one extra page; paired-page mean absolute error was `19.226760`, and mean
    changed-pixel ratio at threshold 16 was `0.158228`.
  - Diagnostics now identify the public-safe pagination risk categories: `DOCX_NUMBERING_INDENT`,
    `DOCX_STYLE_PARAGRAPH_KEEP_RULE`, `DOCX_STYLE_PARAGRAPH_SPACING`, `DOCX_STYLE_TABLE_STYLE`,
    `DOCX_UNSUPPORTED_TABLE_HEADER_ROW`, and `DOCX_UNSUPPORTED_TABLE_STYLE`.
  - Private visual inspection found DOCX tables in `output.pdf` are visibly wrong enough to require their own
    recovery track, not just incremental pagination tuning.
  - Next implementation should start with layout tracing or one of those diagnosed categories; avoid broad
    paragraph parser rewrites until drift location is known.

## Backlog

### Release-Blocking Fidelity

- [x] Implement Adam7 interlaced PNG decoding so embedded interlaced images render instead of being skipped.
- [x] Make omitted embedded image content release-blocking: render supported JPEG/PNG, otherwise emit
  explicit high-severity diagnostics.
- [x] Improve PPTX chart rendering for cached numeric bar-chart XML with an approximate grouped-bar path.
- [ ] Extend PPTX chart rendering beyond the current loose native renderer: labels, legends, axes,
  marker styles, theme/chart style colors, and tighter plot-area layout fidelity.
- [ ] Fix DOCX page geometry and pagination fidelity: section page size/margins, paragraph spacing, manual
  page/column breaks, and keep/widow page-break decisions.
- [x] Add diagnostics when DOCX reference-like pagination risks are detected: multi-section layout,
  unsupported page break variants, or unsupported paragraph keep rules.

### PPTX Feature Survey

- [ ] Text layout: preserve spaces, tabs, line breaks, soft line breaks, kerning-like advances, font
  fallback, mixed run spacing, character spacing, superscript/subscript, and baseline offsets.
- [ ] Text frames: overflow behavior beyond hard clipping, autofit, shrink-to-fit, multi-column text, text
  rotation, and text inside arbitrary shapes.
- [ ] Fonts: select bold/italic faces instead of drawing approximations; support fallback fonts, embedded
  fonts, complex scripts, and bidirectional text.
- [ ] Shapes: more preset geometries, freeform paths, callouts, custom geometry, compound paths, and accurate
  line joins/caps/dashes.
- [ ] Fills/effects: transparency, gradients, pattern fills, picture fills, shadows, glows, reflections,
  blur, soft edges, and 3D effects.
- [ ] Images: placeholder-bound image placement, crop modes, rotation/flip interactions, recolor/duotone,
  transparency, SVG/EMF/WMF, TIFF/GIF/BMP, and image compression variants.
- [ ] Tables: merged cells, vertical alignment, per-edge borders, table styles, cell margins, rich text
  inside cells, and precise row/column sizing.
- [ ] Charts: cached chart images, chart XML rendering, axes, labels, legends, series styling,
  stacked/grouped bars, line charts, combo charts, and embedded chart data.
- [ ] SmartArt/diagrams: use fallback drawings when present; otherwise emit precise diagnostics.
- [ ] Slide inheritance: deeper master/layout placeholder resolution, theme variants, background styles,
  footer/date/slide-number placeholders, and hidden placeholder semantics.
- [ ] Media and dynamic features: videos, audio, animations, transitions, and OLE/ActiveX should remain
  static/diagnostic-only unless a reliable fallback is available.
- [ ] Comments/notes: speaker notes and comments should be ignored with diagnostics or exposed through an
  optional mode, not silently dropped.

### PPTX Synthetic Fidelity Ladder

Build these as public, minimal, one-slide fixtures. Each rung must start with Office PDF/raster inspection,
then compare the candidate PDF/raster output, then receive strict page/dimension checks, expected
diagnostics, and a visual gate once the primitive is close. It is acceptable for private deck pages to
regress while early rungs are rebuilt; the goal is a strict bottom-up progression.

Private deck runs are intentionally expensive and should not be routine. Use `lokad-value-based` only for
feature discovery or occasional acceptance checks after several public rungs have been locked. Typography
work must be driven by public synthetic fixtures and Office PDF inspection because small advance, kerning,
baseline, and highlight errors cascade into unreadable text on larger decks.

Visual-case names should converge to `pptx-ladder-NN-topic-capability[-variant]` and
`docx-ladder-NN-topic-capability[-variant]`, where `NN` is the ladder number, `topic` is a stable domain
bucket such as `typography`, `text-flow`, `shape`, `image`, `group`, or `table`, and `capability` is the
isolated Office behavior. Existing names may remain while active work is in flight, but new cases should use
the normalized scheme and old cases should be renamed in mechanical commits that update manifests, fixture
paths, and ExecPlan references together.

- [x] Ladder 0: blank slide, page size, white/default background, deterministic PDF bytes, and no diagnostics.
- [x] Ladder 1: solid slide backgrounds and master/layout background inheritance in isolation.
- [x] Ladder 2: one plain text box with fixed bounds, one font size, one font family, no wrapping, and
  baseline locked against reference.
- [x] Ladder 3: text wrapping with preserved spaces, explicit line breaks, tabs, paragraph alignment, body
  insets, vertical anchor, and overflow behavior.
- [x] Ladder 3 gate `pptx-ladder-03-preserved-spaces` is at exact raster parity.
- [x] Ladder 3 gate `pptx-ladder-03-text-flow` is at MAE `0.065508` and changed16 `0.001155`.
- [x] Ladder 3 gate `pptx-ladder-03-text-anchor-overflow` is at MAE `0.141281` and changed16 `0.001517`.
- [ ] Ladder 4: styled text runs: bold, italic, underline, color, highlight, mixed fonts, bullet glyphs,
  bullet hanging indents, paragraph spacing, and line spacing.
- [x] Ladder 4 subcase `pptx-ladder-04-all-caps` is near-exact at MAE `0.003613`, changed16 `0.000113`.
- [x] Ladder 4 small-caps runs split lowercase text into scaled uppercase fragments on the same baseline.
- [x] Ladder 4 gate `pptx-ladder-04-small-caps` is at MAE `0.177228`, changed16 `0.001463`.
- [x] Ladder 4 subcase `pptx-ladder-04-character-spacing` performs expanded-spacing line breaking during
  layout and emits Office-like `TJ` per-glyph tracking arrays instead of draw-time `Tc`.
- [x] Ladder 4 gate `pptx-ladder-04-character-spacing` is at MAE `0.165546`, changed16 `0.001458`.
- [x] Ladder 4 font-style subcases keep regular, bold, italic, and bold-italic face resources distinct.
- [x] Ladder 4 gate `pptx-ladder-04-bold-face-single` is at MAE `0.017018`, changed16 `0.000461`.
- [x] Ladder 4 gate `pptx-ladder-04-italic-face-single` is at MAE `0.010782`, changed16 `0.000346`.
- [x] Ladder 4 gate `pptx-ladder-04-bold-italic-face-single` is at MAE `0.022819`, changed16 `0.000641`.
- [x] Ladder 4 gate `pptx-ladder-04-bold-italic-face` is at MAE `0.111322`, changed16 `0.003214`.
- [x] Ladder 4 baseline-shifted runs render at Office's two-thirds glyph size while keeping the declared-size
  baseline offset.
- [x] Ladder 4 gate `pptx-ladder-04-baseline-shift` is at MAE `0.017311`, changed16 `0.000467`.
- [x] Ladder 4 explicit-break baselines use the next line's run size instead of a fixed default.
- [x] Ladder 4 gate `pptx-ladder-04-line-spacing-points` is at MAE `0.014381`, changed16 `0.000410`.
- [x] Ladder 4 mixed-size runs on one line are locked at MAE `0.007254`, changed16 `0.000173`.
- [x] Ladder 4 percentage line spacing follows Office's percent-of-normal-line-advance behavior.
- [x] Ladder 4 percent paragraph before/after spacing is resolved from the paragraph font size.
- [x] Ladder 4 gate `pptx-ladder-04-paragraph-spacing-percent` is at MAE `0.415738`, changed16 `0.004076`.
- [x] Ladder 4 gate `pptx-ladder-04-empty-paragraph-gap` is at MAE `0.306072`, changed16 `0.005082`.
- [x] Ladder 4 gate `pptx-ladder-04-paragraph-advance` is at MAE `0.163712`, changed16 `0.002857`.
- [x] Ladder 4 subcase `pptx-ladder-04-tab-stop` matches Office's continuous-text handling for the standalone
  synthetic `<a:tab/>` fixture.
- [x] Ladder 4 gate `pptx-ladder-04-tab-stop` is at MAE `0.011007`, changed16 `0.000341`.
- [x] Ladder 4 gate `pptx-ladder-04-tab-character` locks explicit tab-character positioning at exact raster
  parity.
- [x] Ladder 4 run highlight rectangles follow Office's taller marker bounds.
- [x] Ladder 4 gate `pptx-ladder-04-highlight-single` is at MAE `0.032134`, changed16 `0.001214`.
- [x] Ladder 4 gate `pptx-ladder-04-transparent-text` locks supported run-level alpha text at MAE
  `0.008701`, changed16 `0.000261`.
- [x] Ladder 4 gate `pptx-ladder-04-field-text` locks cached field text and Office-style slide-number field
  resolution with exact raster parity.
- [x] Ladder 4 gate `pptx-ladder-04-soft-hyphen` locks non-breaking soft hyphen suppression with exact raster
  parity.
- [x] Ladder 4 gate `pptx-ladder-04-nonbreaking-space` locks NBSP text spacing at MAE `0.003519`, changed16
  `0.000095`.
- [x] Ladder 4 bullet styling and hanging wrap are order-aware and Office-positioned.
- [x] Ladder 4 gate `pptx-ladder-04-bullet-style` is at MAE `0.029019`, changed16 `0.000670`.
- [x] Ladder 4 gate `pptx-ladder-04-bullet-wrap` is at MAE `0.137307`, changed16 `0.003059`.
- [x] Ladder 4 gate `pptx-ladder-04-bullet-size-percent` locks `buSzPct` marker sizing at MAE
  `0.040695`, changed16 `0.000765`.
- [x] Windows Symbol cmap subtables are parsed so common private-use PPTX bullet glyphs map through
  `buFont typeface="Symbol"`.
- [x] Ladder 4 gate `pptx-ladder-04-symbol-bullet` is at MAE `0.001766`, changed16 `0.000043`.
- [x] Ladder 4 auto-numbered PPTX paragraphs render basic sequential `buAutoNum` labels.
- [x] Ladder 4 gate `pptx-ladder-04-auto-number-bullets` is at MAE `0.968827`, changed16 `0.008053`.
- [x] Ladder 4 gate `pptx-ladder-04-alpha-number-bullets` locks alphabetic `buAutoNum` expansion at MAE
  `0.580815`, changed16 `0.005619`.
- [x] Ladder 4 Roman `buAutoNum` labels expand through the same numbered-list path.
- [x] Ladder 4 gate `pptx-ladder-04-roman-number-bullets` is at MAE `0.494357`, changed16 `0.004535`.
- [x] Ladder 4 rotated text boxes apply the shape rotation around the Office-like text-frame center.
- [x] Ladder 4 gate `pptx-ladder-04-rotated-text-box` is at MAE `0.032549`, changed16 `0.000842`.
- [x] PPTX `normAutofit` and `spAutoFit` text frames emit `PPTX_UNSUPPORTED_TEXT_AUTOFIT` until shrink/grow
  fitting is implemented.
- [x] Office baseline check: a synthetic non-overflowing `normAutofit fontScale` does not render scaled text in
  Office, so ooxpdf must not apply `fontScale` blindly without an Office-authored overflow/autofit fixture.
- [x] PPTX multi-column and vertical text frames emit explicit unsupported diagnostics instead of silently
  flattening to horizontal single-column text.
- [x] Ladder 4 valid `buClr`/`buSzPts` before the marker remain unit-tested.
- [x] Ladder 4 underline geometry uses scaled OpenType underline metrics.
- [x] Ladder 4 gate `pptx-ladder-04-underline-single` is at MAE `0.052158`, changed16 `0.000768`.
- [x] Ladder 4 gate `pptx-ladder-04-serif-title-underline` is stable at MAE `0.255534`, changed16 `0.004376`.
- [x] Ladder 4 strikethrough uses an Office-like filled rectangle at the strike baseline.
- [x] Ladder 4 gate `pptx-ladder-04-strikethrough-single` is at MAE `0.008026`, changed16 `0.000115`.
- [x] OpenType GPOS pair-positioning x-advance records feed the existing kerning/TJ path for fonts that do
  not expose legacy `kern` pairs.
- [x] OpenType GPOS pair positioning is unit-tested with Arial `To` as typography infrastructure for later
  visual tightening.
- [x] OpenType GPOS pair positioning is restricted to lookups referenced by the `kern` feature, avoiding
  inactive pair-positioning lookups that created parasite inter-letter gaps in Cambria/Cambria Math text.
- [x] Ladder 4 gate `pptx-ladder-04-kerning-accent-highlight` covers normal words with uppercase starts,
  accented letters, and reported French/English word shapes. Latest gated run:
  `artifacts/visual/pptx-ladder-04-kerning-accent-highlight/20260516-130917`, MAE `0.197946`,
  changed16 `0.005130`.
- [x] PPTX highlighted-run rectangles use font Windows ascender/descender metrics instead of a fixed
  height multiplier, keeping Arial highlight parity while tightening Cambria/Cambria Math marker height.
- [ ] Add normalized typography rungs for Office-authored kerning words by font family: Arial, Aptos/Calibri,
  Cambria/Cambria Math, and Segoe UI.
- [x] Add normalized typography rung `pptx-ladder-04-typography-font-families` for Office-authored
  Cambria/Cambria Math, Arial, and Calibri word-spacing probes. Initial gated run:
  `artifacts/visual/pptx-ladder-04-typography-font-families/20260516-132813`, MAE `0.715017`,
  changed16 `0.011127`.
  Office emits distinct font resources and positive/negative `TJ` adjustments per family; the candidate is
  structurally close for Cambria but still has larger Calibri/Arial drift and small baseline deltas.
- [x] Add normalized typography rung `pptx-ladder-04-typography-run-boundaries` for Office-authored
  Cambria/Cambria Math words, accented Latin, headline words, and highlighted run boundaries. Latest gated
  run: `artifacts/visual/pptx-ladder-04-typography-run-boundaries/20260516-130900`, MAE `0.203680`,
  changed16 `0.005800`.
- [x] OpenType GPOS extension lookups are parsed for active `kern` feature pair positioning, recovering
  Office-like small Cambria/Cambria Math intra-word adjustments without reintroducing inactive feature
  lookups that caused large parasite gaps.
- [x] Typography experiment: exact font resolution still keeps math-table TTC faces exact, while PPTX
  presentation text now uses a generic same-collection text face when a math-table face has one. This avoids
  font-name special cases while matching Office's slide-text behavior more closely.
- [x] Slide-3 typography gap: fix CID font width emission so every embedded glyph has a `/W` entry.
  Missing glyph widths caused PDF viewers to use default advances, producing the parasite inter-letter
  gaps seen with exact math-table fonts. `PdfEmbeddedFontWidthsCoverEncodedGlyphs` now locks this.
- [x] Slide-3 typography gap: include Microsoft cloud-font cache discovery in the Windows resolver.
  Office-authored decks may use cloud fonts that are absent from `C:\Windows\Fonts`; missing those fonts
  sent text through unrelated installed fallbacks.
- [x] Slide-3 typography gap: clamp highlight rectangles to line-box proportions when font OS/2 metrics
  are extreme, instead of letting raw ascender/descender values create multi-line yellow bands.
- [x] Revisit locked public typography visual thresholds after exact-font rendering. Reject out-of-range
  ascender metrics for baseline placement instead of clamping them, which keeps exact math-table fonts
  while restoring the public typography gates: `pptx-ladder-04-typography-run-boundaries`
  MAE `0.186656`, changed16 `0.005533`; `pptx-ladder-04-typography-font-families`
  MAE `0.715017`, changed16 `0.011127`.
- [ ] Remove remaining PPTX text-flow approximations that pick specific families as default behavior.
  Defaults must come from OOXML theme resolution, font metadata, or a documented generic fallback stack,
  not font-by-font aliases.
- [x] Add `pptx-ladder-04-cambria-math-dense-wrap-probe` for the private slide-11 class of issues:
  dense Cambria Math paragraphs, mixed bold spans, an empty paragraph, and a narrow public text frame. OOXPDF
  now keeps the short final heading word on the first line through a width-relative final-word wrap rule
  instead of a broad font-size-only tolerance, and the visual gate is at MAE `2.769287`, changed16
  `0.025235`.
- [x] Add a public Aptos-theme centered logo-box probe from the private slide-2 pattern: a 40pt default
  minor-font run in a narrow centered frame must wrap into two centered lines. This prevents the
  width-relative final-word tolerance from letting large overflows remain clipped on one line.
- [ ] Continue tightening `pptx-ladder-04-typography-run-boundaries` toward near-pixel parity by comparing
  Office and candidate `TJ` arrays, baseline `Tm` values, and highlight geometry line by line.
- [ ] Add normalized typography rungs for accented Latin and punctuation-adjacent words: French accents,
  apostrophes, non-breaking spaces, narrow spaces, currency symbols, and hyphen/dash variants.
- [ ] Add normalized typography rungs for run-boundary invariance: same text as one run, per-word runs,
  per-style runs, highlighted spans, and mixed fill spans must keep the same visible advances when styles
  do not change metrics.
- [x] Add diagnostic public rungs for capital spacing, accented Latin spacing, and run-boundary invariance.
  `pptx-ladder-04-typography-capital-spacing-probe` and
  `pptx-ladder-04-typography-boundary-invariance-probe` now enforce Office PDF text-operation parity at
  `0.05pt`; the corrected accented probe now enforces decoded text and line-start parity while remaining
  a raster/font-shape probe.
- [x] Port the `pptx-renderer` baseline-shift rule: small PPTX `baseline` nudges keep the nominal font size,
  while larger super/subscript shifts use the reduced font scale. The public unit coverage now locks both
  `baseline="10000"` and `baseline="30000"` behavior without font-family-specific logic.
- [x] Port the `pptx-renderer` trailing-break/end-paragraph sizing rule: a final `<a:br/>` before
  `endParaRPr` must consume line height using the end-paragraph run size, so following paragraphs keep
  Office-compatible vertical placement.
- [x] Isolate the private slide-3 XML locally and keep only public-safe structural findings in the work
  log. The source text does not contain the visible parasite capital/accent spaces, so the defect is in
  renderer layout/emission rather than private document content.
- [x] Add public slide-3-inspired typography probes:
  `pptx-ladder-04-typography-alignment-values-probe` covers `just`, `dist`, `justLow`, and `thaiDist`;
  `pptx-ladder-04-typography-cambria-math-run-boundaries-probe` covers mixed highlighted Cambria Math runs;
  `pptx-ladder-04-typography-slide3-narrow-cambria-probe` covers narrow wrapped Cambria Math columns.
- [x] Split PPTX alignment handling so `dist`, `justLow`, and `thaiDist` remain observable text-flow modes
  instead of collapsing into `Justify`. Office-style distributed Latin text now uses positioned per-glyph
  spans, while `justLow` and `thaiDist` use word-spacing justification for Latin text.
- [x] Remove the small-text natural-line-height exception from PPTX wrapping. The slide-3-inspired narrow
  Cambria probe showed Office advancing 12pt wrapped lines by about `14.4pt`; candidate line Y deltas dropped
  from multi-point drift to about `0.3pt` after using the general `1.2` default line-height rule.
- [x] Refine the public slide-3 narrow Cambria probe to mirror private-safe structure: `spAutoFit`, explicit
  `spcAft=600`, non-bold body runs, and a short highlighted run. Latest visual run:
  `artifacts/visual/pptx-ladder-04-typography-slide3-narrow-cambria-probe/20260517-092613`, MAE `1.332658`.
- [ ] Reverse-engineer Office's `spAutoFit` text fitting for narrow PPTX columns. The refined probe shows
  Office applying about `-0.036pt` character spacing to 12pt body lines, which changes wraps around short
  highlighted runs; candidate currently emits `Tc=0`.
- [x] Fix horizontal `spAutoFit` height-overflow semantics against Office PDF output:
  `pptx-ladder-04-typography-spautofit-headline-wrap-probe` reproduces the slide-9/66 dense-heading
  wrapping pattern with public text. Office keeps the run at `18pt` and wraps into four text operations;
  OOXPDF no longer shrinks horizontal `spAutoFit` merely because the text is taller than the original
  box. The case is gated with PDF text-operation parity at `0.02pt` position tolerance.
- [x] Add stress probes for highlight tracking across no-autofit, wide `spAutoFit`, and narrow `spAutoFit`
  variants. Office evidence shows plain `spAutoFit` text keeps `Tc=0`; compact tracking starts on the
  highlighted run and carries through following runs in the same paragraph.
- [x] Apply Office's highlighted-run tracking rule in the PPTX text model:
  `CharacterSpacing += -0.003 * fontSize` on the highlighted run and following text runs in the paragraph.
  The candidate still encodes this via `TJ` positioning arrays rather than Office-style `Tc`, but layout
  measurement and rendered glyph positions now share the same rule.
- [ ] Tighten the slide-3-inspired Cambria probes next: compare Office and candidate line breaks, `TJ`
  arrays, highlight rectangles, and font metrics before rerunning the private deck.
- [ ] Add normalized typography rungs for Office PDF text-object structure: compare candidate `TJ` arrays
  and text matrices against Office for simple lines before accepting near-pixel raster gates.
- [x] Add a public bold-wrap typography rung for the slide-11 class of issues:
  `pptx-ladder-04-typography-bold-wrap-probe` locks two-line Cambria Math bold wrapping, synthetic-bold
  fill-and-stroke emission, and PDF text-operation positions against Office at `0.08pt` tolerance.
- [x] Extend `PdfInspect` with `text-operations.json` output so public typography cases can compare Office
  and candidate font size, `Tc`, `Tm`, and raw `TJ/Tj` structure before raster tuning. The first
  `pptx-ladder-03-text-flow` inspection shows sub-point baseline drift plus missing Office micro-`TJ`
  adjustments rather than a gross layout miss.
- [x] Add `tools/ComparePdfTextOperations.ps1` for compact Office-vs-candidate `text-operations.json`
  deltas. On `pptx-ladder-03-text-flow`, it reports equal text operation counts, first-line Y delta
  `-0.712pt`, missing Office `Tc=-0.006`, and small X/Y deltas on centered/wrapped lines.
- [x] Wire optional PDF text-operation gates into `CheckVisualCase.ps1` and lock the supported
  `pptx-ladder-02-plain-text` rung at `0.05pt` position tolerance. This creates a stricter bottom-up
  typography gate before broader raster metrics are meaningful.
- [x] Refine explicit `<a:br/>` line metrics separately from wrapped line metrics. The public
  `pptx-ladder-03-text-flow` rung improved from MAE `0.065508`, changed16 `0.001155`,
  SSIM `0.969415` to MAE `0.030416`, changed16 `0.000838`, SSIM `0.994694`, and now has
  a PDF text-operation gate at `0.12pt` position tolerance.
- [x] Tighten the early manual-break/overflow visual gates after the same line-metric fix:
  `pptx-ladder-03-text-anchor-overflow` is now at MAE `0.027775`, changed16 `0.000711`,
  SSIM `0.993028`.
- [x] Typography shaping note: expanding GPOS class-pair parsing to include class-0 right glyphs did not
  change `pptx-ladder-03-text-flow` text operations or metrics, so the remaining Arial micro-`TJ`
  differences are not explained by that parser gap.
- [x] Fix the Ladder 3 clipped-overflow text path so visible clipped lines are still rendered while fully
  out-of-clip following lines are suppressed by a baseline-aware precheck. Latest gated run:
  `artifacts/visual/pptx-ladder-03-text-anchor-overflow/20260516-160344`, MAE `0.080874`,
  changed16 `0.001259`, SSIM `0.958285`; it remains `needs-review` under the strict support gate.
- [x] Ladder 4 gate `pptx-ladder-04-mixed-font-size-stack` is at MAE `0.071192`, changed16 `0.001852`.
- [x] Ladder 4 gate `pptx-ladder-04-mixed-paragraph-stack` is at MAE `0.452275`, changed16 `0.008806`.
- [ ] Ladder 4 remaining combined-stack gaps are finer glyph/font details after basic bullet font selection.
- [x] Ladder 5: basic shapes cover rectangle, rounded rectangle, ellipse, line, fills, strokes, stroke widths,
  rotation, flips, and clipping-free z-order.
- [x] Ladder 5 gate `pptx-ladder-05-basic-shapes` reuses the public `pptx-shapes` fixture and is at MAE
  `0.009024`, changed16 `0.000356`, with no diagnostics.
- [ ] Ladder 6: preset and connector shapes cover arrows, connector endpoints, arrowheads, dashes,
  line caps/joins, callouts, and common freeform/custom path fallbacks.
- [x] Ladder 6 gate `pptx-ladder-06-connector-arrow` locks a straight connector with triangle tail arrowhead
  and a down-arrow preset with exact raster parity.
- [x] Ladder 6 gate `pptx-ladder-06-line-arrowheads` locks Office `type="arrow"` head and both-end connector
  arrowheads at near-exact raster parity.
- [x] Ladder 6 gate `pptx-ladder-06-block-arrows` locks up/left/right block-arrow preset geometry with exact
  raster parity.
- [x] Ladder 6 gate `pptx-ladder-06-double-arrows` locks left-right and up-down block arrow presets with exact
  raster parity.
- [x] Ladder 6 gate `pptx-ladder-06-basic-polygons` locks triangle and diamond preset geometry with exact
  raster parity.
- [x] Ladder 6 gate `pptx-ladder-06-quadrilateral-presets` locks parallelogram and trapezoid preset geometry
  with exact raster parity.
- [x] Ladder 6 gate `pptx-ladder-06-more-polygons` locks right triangle, pentagon, hexagon, and octagon preset
  geometry with exact raster parity.
- [x] Ladder 6 gate `pptx-ladder-06-chevron-home-plate` locks chevron and home-plate preset geometry with exact
  raster parity.
- [x] Ladder 6 gate `pptx-ladder-06-star-seal` locks five-point and six-point star preset geometry with exact
  raster parity.
- [x] Ladder 6 gate `pptx-ladder-06-more-stars` locks four-point and eight-point star preset geometry with exact
  raster parity.
- [x] Ladder 6 gate `pptx-ladder-06-symbol-polygons` locks the plus/cross preset polygon with exact raster
  parity.
- [x] Ladder 6 gate `pptx-ladder-06-rect-callout` locks the `wedgeRectCallout` preset with exact raster parity.
- [x] Ladder 6 gates `pptx-ladder-06-dashed-connector` and `pptx-ladder-06-dash-dot-connector` lock Office dash
  presets with exact raster parity.
- [x] PPTX stroke enum unit coverage now covers all known `a:prstDash` values, non-triangle line-end
  marker variants (`stealth`, `diamond`, `oval`), and `w`/`len` marker size variants with first-pass
  PDF geometry.
- [x] Ladder 6 gates `pptx-ladder-06-round-cap-connector` and `pptx-ladder-06-square-cap-connector` lock round
  and square caps/joins with exact raster parity.
- [x] Ladder 6 gates `pptx-ladder-06-bevel-join-rect` and `pptx-ladder-06-round-join-rect` lock explicit
  bevel/round line joins on stroked rectangles with exact raster parity.
- [x] Ladder 6 custom geometry and unsupported callout shapes emit explicit unsupported diagnostics.
- [ ] Ladder 6 remaining subcases should isolate additional visual callout rendering and other preset
  geometries.
- [x] Ladder 7: images cover JPEG/PNG placement, alpha masks, crop rectangles, aspect-fit/fill behavior,
  rotation/flip interactions, and unsupported image diagnostics.
- [x] Ladder 7 gate `pptx-ladder-07-basic-image` locks the existing public stretched-image fixture with exact
  raster parity.
- [x] Ladder 7 gate `pptx-ladder-07-image-crop` locks a minimal left/right cropped PNG picture with exact raster
  parity.
- [x] Ladder 7 gate `pptx-ladder-07-image-fill-rect` locks destination `a:fillRect` placement with exact raster
  parity.
- [x] Ladder 7 gate `pptx-ladder-07-image-crop-fill-rect` locks source crop combined with destination fill-rect
  placement with exact raster parity.
- [x] Ladder 7 gate `pptx-ladder-07-shape-picture-fill` locks rectangular shape `a:blipFill` rendering through
  the image pipeline with exact raster parity.
- [x] Ladder 7 gate `pptx-ladder-07-ellipse-picture-fill` locks clipped shape picture fills at MAE
  `0.000142`, changed16 `0`.
- [x] Ladder 7 gate `pptx-ladder-07-image-alpha` locks a transparent PNG soft-mask case with exact raster parity.
- [x] Ladder 7 gate `pptx-ladder-07-jpeg-image` locks minimal JPEG placement at MAE `0.134097`, changed16
  `0.005486`.
- [x] Ladder 7 gate `pptx-ladder-07-bmp-image` locks uncompressed 24-bit BMP placement with exact raster
  parity.
- [x] Ladder 7 gate `pptx-ladder-07-bmp-alpha` locks Office-compatible 32-bit BI_RGB BMP alpha-byte handling
  with exact raster parity; Office treats the byte as unused in this case.
- [x] Ladder 7 gate `pptx-ladder-07-image-rotation` locks rotated picture transforms with exact raster parity.
- [x] Ladder 7 gate `pptx-ladder-07-image-flip` locks horizontal picture flips with exact raster parity.
- [x] Ladder 7 gate `pptx-ladder-07-image-rotate-flip` locks combined rotation/flip with exact raster parity.
- [x] PPTX tiled image fills and blip recolor modes emit explicit unsupported diagnostics instead of being
  silently flattened.
- [x] Ladder 7 remaining subcases isolate aspect/fill variants and unsupported image diagnostics, including
  shape picture fills, clipped picture fills, tiled fills, recolor modes, and unsupported advanced fills.
- [ ] Ladder 8: grouped content covers nested group transforms, grouped pictures, grouped text, grouped shapes,
  child coordinate scaling, z-order, and clips.
- [x] Ladder 8 gate `pptx-ladder-08-grouped-shape` locks grouped shape child coordinate scaling at MAE
  `0.000142`, changed16 `0`.
- [x] Ladder 8 gate `pptx-ladder-08-grouped-picture` locks grouped picture scaling with exact raster parity.
- [x] Ladder 8 gate `pptx-ladder-08-grouped-text` locks text boxes inside grouped content at MAE `0.002687`,
  changed16 `0.000079`.
- [x] Ladder 8 gate `pptx-ladder-08-nested-grouped-text` locks nested group transform composition for text
  boxes at the same thresholds.
- [x] Ladder 8 gates `pptx-ladder-08-text-shape-zorder` and `pptx-ladder-08-shape-picture-zorder` lock simple
  sibling order with exact raster parity.
- [x] Ladder 8 gate `pptx-ladder-08-table-shape-zorder` locks table graphic frames participating in slide
  sibling order at MAE `0.062251`, changed16 `0.000748`.
- [ ] Ladder 8 remaining subcases should isolate z-order with charts and clipping.
- [ ] Ladder 9: slide inheritance covers placeholders, master/layout text styles, hidden placeholders,
  footer/date/slide number placeholders, theme fonts, and theme color transforms.
- [x] Ladder 9 gate `pptx-ladder-09-title-placeholder` isolates a slide title placeholder inheriting layout
  bounds, 60pt centered style, bottom anchor, and master title run defaults.
- [x] Ladder 9 gate `pptx-ladder-09-title-placeholder` is at MAE `1.001682`, changed16 `0.006872`, using
  Calibri-family fallback when Aptos theme fonts are unavailable.
- [x] Ladder 9 gate `pptx-ladder-09-title-placeholder-arial-theme` removes that font variable and locks the
  same inherited placeholder geometry at MAE `0.215164`, changed16 `0.002463`.
- [ ] Ladder 10: tables cover fixed grid, per-edge borders, fills, cell margins, vertical alignment, merged
  cells, rich text inside cells, and table styles.
- [x] Ladder 10 gate `pptx-ladder-10-basic-table` locks a fixed-grid table with explicit cell fills, default
  Office grid lines, and Office-like cell text insets/baselines.
- [x] Ladder 10 gate `pptx-ladder-10-basic-table` is at MAE `0.049042`, changed16 `0.000916`.
- [x] Ladder 10 gate `pptx-ladder-10-explicit-borders` locks per-edge borders with exact raster parity.
- [x] Ladder 10 gate `pptx-ladder-10-vertical-align` locks top/center/bottom table-cell anchors at MAE
  `0.013784`, changed16 `0.000302`.
- [x] Ladder 10 gate `pptx-ladder-10-unstyled-grid` locks unstyled default grid lines with exact raster parity.
- [x] Ladder 10 gates `pptx-ladder-10-horizontal-merge` and `pptx-ladder-10-vertical-merge` lock merged cells
  with exact raster parity.
- [x] Ladder 10 gate `pptx-ladder-10-rich-text-cell` locks styled table-cell run sequencing and
  whitespace-preserving run gaps at MAE `0.092446`, changed16 `0.001351`.
- [x] Ladder 10 gate `pptx-ladder-10-cell-margins` locks table-cell body insets at MAE `0.259091`, changed16
  `0.002151`.
- [x] Ladder 10 gate `pptx-ladder-10-cell-fill-alpha` locks table-cell fill transparency at MAE `0.035698`,
  changed16 `0.000926`.
- [x] Ladder 10 gate `pptx-ladder-10-border-alpha` locks explicit table-border transparency at MAE `0.223404`,
  changed16 `0.002494`.
- [x] Supported built-in PPTX table styles now honor `bandCol` conditional fills in addition to `bandRow`
  for Light Style 1, Medium Style 2, and Dark Style 1 synthetic coverage.
- [x] Supported built-in PPTX table styles now apply last-row and last-column conditional text styling
  instead of limiting conditional bold/color handling to first rows and first columns.
- [x] PPTX table conditional flags now resolve both attribute form (`firstRow="1"`) and child element form
  (`<a:bandCol/>`), matching the broader flag handling observed in `pptx-renderer`.
- [ ] Ladder 10 remaining subcases should isolate broader table style variants.
- [ ] Ladder 11: charts: cached image fallback, basic bar/line/pie rendering, axes, labels, legends, series
  styles, stacked/grouped variants, and chart diagnostics.
- [x] Preserve chart data-label category-name and series-name flags through native bar/line rendering:
  `showCatName` and `showSerName` now compose labels from the typed scene series/category metadata instead
  of being parsed and then dropped before PDF emission. Public chart tests lock a custom separator glyph in
  emitted chart labels.
- [ ] Ladder 12: effects and advanced fills cover transparency, gradients, pattern fills, shadows, glows, soft
  edges, picture fills, and explicit diagnostics for unsupported effects.
- [x] Ladder 12 unsupported gradient fills, pattern fills, advanced picture fills, and effect lists emit
  slide-scoped diagnostics instead of being silently dropped.
- [x] Ladder 12 gates `pptx-ladder-12-solid-alpha` and `pptx-ladder-12-line-alpha` lock solid shape fill
  transparency and line transparency with exact raster parity.
- [x] Ladder 12 gate `pptx-ladder-12-arrow-alpha` locks transparent Office `type="arrow"` connector arrowheads
  at MAE `0.006754`, changed16 `0.000174` after Office-style endpoint geometry.
- [x] Ladder 12 gate `pptx-ladder-12-picture-alpha` locks whole-picture `alphaModFix` transparency at MAE
  `0.109499` with no pixels changed above threshold 16.
- [x] Ladder 12 gate `pptx-ladder-12-text-alpha` locks run text transparency at MAE `0.006230`, changed16
  `0.000170`.
- [x] Ladder 12 gate `pptx-ladder-12-shadow-diagnostic` locks explicit shadow/effect diagnostics while
  allowing the expected omitted-shadow visual delta at MAE `0.408920`, changed16 `0.006530`.
- [x] Ladder 12 gates `pptx-ladder-12-gradient-diagnostic`, `pptx-ladder-12-pattern-diagnostic`, and
  `pptx-ladder-12-picture-fill-diagnostic` lock explicit advanced-fill diagnostics while allowing expected
  omitted-fill visual deltas; the picture-fill fixture now covers an unsupported Office-authored heart fill at
  MAE `3.368542`, changed16 `0.038278`.
- [ ] Ladder 12 remaining work is visual rendering for each effect/fill family.
- [ ] For every ladder rung, keep public synthetic fixture content artificial and minimal. Do not derive
  fixture text, images, layout, or styling from private documents.
- [ ] Run the relevant public visual case after each rung change; run private PPTX only as feature-discovery
  smoke evidence until the public ladder is much more complete.
- [ ] Revisit PPTX unit tests under the Office-PDF-first workflow: keep parser/safety/API tests, keep
  deterministic low-level PDF writer tests, but replace brittle renderer operator-position assertions with
  public visual gates or assertions derived from inspected Office reference geometry.

### PPTX Private Deck Recovery Plan

- [x] Add a private-safe PPTX slide inventory tool that reports counts and feature flags per slide: shapes,
  grouped shapes, pictures, charts, tables, text boxes, placeholders, inherited master/layout content, theme
  references, fills/effects, transforms, clips, relationships, and diagnostics without exposing slide text or
  images.
- [ ] Revisit `lokad-value-based` one slide at a time as an acceptance corpus. For each slide, inspect
  reference vs candidate, list only generic public-safe gaps, map each gap to existing ladder rungs, and
  rerender after every relevant gated public fixture change.
- [ ] If a private-slide gap is not already covered by a passing public ladder fixture, create or tighten a
  minimal public synthetic fixture first. Do not implement private-slide-driven renderer changes until the
  corresponding public fixture is close to pixel-perfect and gated.
- [ ] After a slide is close to pixel-perfect or every remaining gap is covered by explicit planned public
  fixtures, continue to the next slide. Private slides should test combinations and acceptance, not replace
  the bottom-up ladder.
- [ ] Maintain a private per-slide checklist with only public-safe fields: slide number, private rating,
  missing content, wrong order, wrong placement, wrong sizing, wrong text layout, wrong styling, and
  unsupported features lacking diagnostics.
- [ ] Establish private visual gates for accepted slides: page count and dimensions must match, diagnostics
  must explain omissions, and private human/agent rating must be close to pixel-perfect before optimizing
  aggregate deck metrics.
- [ ] Fix slide composition order first: master, layout, slide background, inherited placeholders, slide
  shapes, groups, pictures, tables, charts, and overlays must render in PowerPoint z-order.
- [ ] Audit and fix master/layout/placeholder inheritance, including placeholder matching, hidden
  placeholders, text/body placeholders, footer/date/slide-number placeholders, theme variants, and background
  styles.
- [ ] Build slide-level diagnostics for unsupported or approximated visible content: effects, transparent
  fills, complex shapes, chart renderings, SmartArt, media/OLE, unsupported images, and placeholder fallbacks.
- [ ] Address dominant primitives after ordering/inheritance are under control: text autofit/shrink, bullets,
  font fallback; image placeholder crop/fit and rotation/flip; table styles and merged cells; chart
  cached-image fallbacks and labels.
- [x] Render supported DrawingML diagonal pattern fills for ordinary PPTX shapes as background plus clipped
  hatches, and keep diagnostics only for pattern presets that still fall outside the implemented subset.
- [x] Honor `bodyPr@compatLnSpc` as Office-compatible tight default line spacing in PPTX text frames, with a
  synthetic line-break fixture locking the reduced baseline step for multi-column/text-flow cases.
- [ ] Private slide 5 visible remaining problem: the right-side chart has an incorrectly placed secondary
  value axis and an incorrectly placed upward green arrow.
  Inspect whether this is a value-axis title, rotated axis label text, tick-label formatting, or chart-style
  inheritance, then reproduce with a minimal public chart-axis fixture before changing renderer logic.
  - [x] Inspect the chart XML shape: this case uses a multi-axis bar-chart pattern with a right-side value
    axis rather than a standalone rotated axis title.
  - [x] Add a public synthetic bar-chart fixture with a secondary right value axis and lock the right-axis
    stroke/label path in the native chart renderer.
  - [x] Render additional `barChart` groups in the same `plotArea` against their referenced value axes,
    covering combo stacked-bar charts with separate left/right scales.
  - [x] Replace the fixed four-interval auto tick fallback with an Office-like 1/2/5/10 major-unit rule and
    expand value-axis label clipping so zero/max labels remain visible.
  - [x] Render chart `a:pattFill` series/point fills as background plus clipped diagonal hatches instead of
    collapsing patterned Office fills to a solid palette color. The public chart fixture now locks this path.
  - [x] Route chart graphic frames through the same group-transform and z-order dispatcher as shapes, so
    grouped charts and their overlay arrows/labels keep Office-equivalent placement.
  - [x] Isolate the private slide-5 upward green arrow as a flipped vertical `straightConnector1` overlay,
    not as chart XML. A public synthetic connector case now locks the zero-width, `flipV`, stealth-tail
    geometry pattern for future arrow-placement refinements.
  - [x] Port the `pptx-renderer`/OOXML `gapWidth` bar-sizing rule for clustered and stacked bar charts,
    replacing fixed bar-width constants with the general category-band formula.
  - [x] Honor explicit chart-axis `a:ln/a:noFill` so value/category axes do not fall back to a dark default
    stroke when Office suppresses the axis line.
  - [x] Expand category-axis label layout beyond the exact category slot so labels such as `Inventory` are
    not visually clipped by an over-tight synthetic text box.
  - [x] Expand the category-axis label clipping rectangle vertically so descenders are not cut from bottom
    axis labels.
  - [x] Fix horizontal bar chart category/value ordering against the public Office-backed chart rung:
    category index zero now maps to the bottom bar for `orientation="minMax"`.
  - [x] Add Office-like automatic chart titles from the first series name only for single-series bar charts
    when `autoTitleDeleted` is false and no explicit title node is present.
  - [x] Port the `pptx-renderer` nice-axis maximum rule for axes without explicit `c:max`, so max data values
    such as `45` expand to an Office-like `50` axis cap.
  - [x] Use document theme accent colors for default/vary-colors bar chart fills instead of a hard-coded
    Excel palette; the public horizontal-bar rung now matches Office's category colors.
  - [x] Increase the default native bar-chart plot height so Office-like value axes use the full plot region
    instead of compressing labels downward. This directly improves the private slide-5 right axis and keeps
    the public bar/column chart gates passing.
  - [x] Restrict implicit chart-title fallback to single-series bar charts. Applying the fallback to other
    chart families creates false titles in the current Office-backed chart ladder; multi-series bar charts no
    longer get false titles from their first series name.
  - [x] Bind secondary right-axis label rendering to the extra bar-chart group that owns the right value axis,
    so combo charts do not use the primary series as the fallback range for secondary-axis ticks.
  - [x] Port the `pptx-renderer` group-transform flip rule into the shared PPTX transform model:
    group `flipH`/`flipV` now mirrors child bounds and composes child flips before shape, image, text,
    and chart overlay rendering. A public synthetic connector case locks a flipped grouped vertical
    stealth-arrow pattern.
  - [x] Port the first `pptx-renderer` axis-label option slice for native charts:
    `c:tickLblPos val="none"` hides tick labels without hiding the axis line, and value-axis `c:numFmt`
    drives basic currency, percent, thousands, and decimal tick-label formatting. A public synthetic
    chart case locks hidden category labels and formatted value-axis labels.
  - [x] Extend the native chart axis-label slice to `tickLblPos high/low` for value axes:
    labels can now be placed on the high or low side independently of the physical axis line.
  - [x] Add public Office-backed `pptx-ladder-11-secondary-axis-overlay-probe`, generated from PowerPoint,
    to lock the slide-5/slide-25 chart pattern bottom-up: stacked column groups, independent primary and
    secondary value axes, patterned secondary bars, and an overlay connector.
  - [x] Add public Office-backed `pptx-ladder-11-compact-stacked-secondary-axis-probe` to mirror the compact
    private slide-5/slide-25 chart shape more closely: no legend/title, suppressed gridlines, stacked
    primary and secondary column groups, patterned secondary series, compact tick labels, and overlay
    connectors. The fixture generator now closes embedded Excel workbooks after editing chart data so it
    does not leave orphaned Excel instances.
  - [x] Honor hidden chart gridline styles: `c:majorGridlines/c:spPr/a:ln/a:noFill` and the corresponding
    minor gridline form no longer draw gridlines. This tightens the compact stacked secondary-axis probe
    to MAE `0.905004` and changed-pixel ratio threshold 16 `0.012336`.
  - [x] Render chart percentage pattern fills (`pct*`) as tiled dot fields instead of diagonal hatches,
    following the `pptx-renderer` distinction between percentage and diagonal presets. This tightens the
    compact stacked secondary-axis probe to MAE `0.885312` and changed-pixel ratio threshold 16 `0.011982`.
  - [x] Place multiple value-axis label columns on the same side instead of overlaying them, and size the
    label boxes from tick text instead of a fixed fraction of the plot width.
  - [ ] Extend combo/multi-axis chart support beyond the first bottom-up slice: bind each chart group to its
    referenced axes, honor axis crossing/orientation, keep primary/secondary scales independent, and place
    non-axis overlays such as the private slide 5 upward green arrow with Office-equivalent transforms.
  - [ ] Private slide 25 repeats the remaining multi-axis chart gap: the left schema's right-side value-axis
    labels and the overlaid green upward arrow are still vertically off. Treat this as the same generic
    secondary-axis/overlay-transform problem as slide 5, not as a static chart fallback.
- [ ] Private slide 6 visible remaining problem: a centered line of text is vertically misaligned inside its
  grey box. Reproduce with a public text-box fixture covering vertical anchor, body insets, line height, and
  shape fill/stroke context.
  - [x] Check `pptx-renderer` handling: table cells apply `a:tcPr` margins as padding and map
    `a:tcPr@anchor` to content-box vertical alignment with top as the Office default.
  - [x] Route PPTX table-cell text through the shared first-class text layout pipeline by synthesizing a
    bounded text shape per cell, preserving `tcPr` margins, vertical anchor, clipping, and table-style text
    defaults. This removes the table-only text estimator before pursuing tighter Office wrap/metric parity.
  - [x] Tighten PPTX table text wrapping against Office output. Public synthetic case
    `PptxSyntheticTableKeepsSlide6HeaderOnOneLine` mirrors the private slide 6 header cell and locks the
    Office behavior where fixed table-cell layout gives wrapping more room than the clipped content box.
    OOXPDF now carries a separate text wrap width for table-cell text while preserving tcPr drawing insets
    and clipping.
  - [x] Apply the same model to OOXPDF table-cell text placement: tcPr margins override bodyPr defaults,
    middle/bottom anchors account for estimated wrapped content height, and the behavior is locked by
    `PptxSyntheticTableCentersTextByContentHeight`.
  - [x] Align table-cell vertical anchoring with the actual text-flow wrap width instead of the clipped text
    width. `PptxSyntheticTableKeepsSlide6HeaderOnOneLine` now locks both one-line wrapping and the centered
    baseline for the slide-6-derived header, and private regen `20260523-210035` places the three header
    baselines at the Office reference y-position.
- [ ] Private slide 10 visible remaining problem: one headline line is positioned too high. Map it to public
  typography baselines, paragraph spacing, inherited bodyPr insets, or placeholder geometry before fixing.
  - [x] Latest private regen `20260523-121550` no longer shows the headline-position issue at manual
    inspection scale; keep open only if manual review finds a remaining delta.
- [ ] Private slide 11 visible remaining problem: a lower-left image overlaps preceding text, and adjacent
  left-side title text formatting mismatches Office. Reproduce as public picture/text z-order and inherited
  text-style fixtures.
  - [ ] Latest inspection suggests the overlap is caused by preceding text wrapping/flowing too tall, not by
    the timeline graphic being placed too high. Isolate with public typography fixtures for requested-font
    metrics, mixed bold/regular run widths, and default line advance before moving the graphic.
  - [ ] Current layout diagnostics after the break-space wrapping fix show the text frame ending above the
    timeline graphic, while the graphic matrix itself already matches the Office PDF. Regenerate the private
    visual case and keep this open for any remaining title/font/style mismatch.
  - [x] Private regen `20260523-230344` no longer shows the lower-left timeline overlapping the preceding
    text. Keep the slide item open for the remaining left text wrap/style differences, now covered first by
    public bold-wrap typography rungs.
- [ ] Private slide 19 visible remaining problem: lower-left logo rendering, paired-arrow geometry, and
  center/right font selection differ from Office. Split into public logo/image recolor or crop diagnostics,
  arrow geometry fixtures, and font fallback/style inheritance fixtures.
  - [x] Add SVG linear-gradient support for compound path pictures by clipping to the SVG path and painting
    sampled gradient strips; this fixes the flat-color logo fallback seen in the private deck.
- [ ] Private slide 23 visible remaining problem: top text block formatting is off and an emphasized fragment
  overlaps the horizontal separator. Reproduce with public mixed-run typography, paragraph spacing, and
  separator/z-order fixtures.
- [ ] Private slide 30 visible remaining problem: the lower center/right date schema is badly off, likely a
  geometry or grouped-transform issue. Inventory shapes/connectors/transforms and isolate public geometry
  fixtures before changing renderer logic.
  - [x] Inspect the broken date schema: the timeline cluster includes a rotated group whose child coordinate
    space was scaled but not rotated by OOXPDF.
  - [x] Add public coverage for rotated grouped shapes and propagate group rotation into child bounds before
    rendering.
  - [x] Render `prstGeom arc` strokes from preset guide angles instead of falling back to rectangle stroking;
    private page 30 now shows dashed arcs around the milestones rather than dashed bounding boxes.
- [x] Slide 2 public ladder: lock a minimal shape-text fixture for `fontRef`/theme color inheritance when
  text runs have no direct fill, including a no-fill shape with a visible line and centered text.
- [x] Slide 2 public ladder: lock highlighted centered mixed-run title-sized text using Cambria/Cambria Math
  without explicit character spacing.
- [x] Slide 2 public ladder: lock highlighted centered mixed-run footer-sized text using Cambria/Cambria Math
  without explicit character spacing.
- [x] Slide 2 public ladder: lock Office-compatible text frame insets, vertical anchoring, and highlight
  rectangles for small centered text boxes.
- [x] Port the Office/PDF synthetic-bold rule observed in private deck baselines and `pptx-renderer` browser
  output: when no bold face is available, render text as fill-and-stroke with a font-size-proportional stroke
  instead of drawing duplicate offset glyphs. `PptxSyntheticTextBoxUsesOfficeSyntheticBoldStroke` locks the
  fill-and-stroke path for Cambria Math.
- [x] Slide 3 public ladder: lock `spAutoFit` overflow/autosize text boxes with Office-authored PDFs before
  enabling renderer-side autofit behavior.
- [x] Slide 3 public ladder: lock square-wrapped overflow text frames with 10.5 pt and 12 pt text,
  including mixed regular, bold, and italic runs.
- [x] Slide 3 public ladder: lock grouped picture-plus-caption cells, including group transform precision,
  picture positioning, centered italic caption text, and z-order.
- [x] Slide 3 public ladder: lock highlighted headline text with multiple runs and Office-matched highlight
  bounds.
- [x] Slide 2/3 public ladder: lock a Cambria/Cambria Math kerning/accent typography probe for parasite
  gaps such as uppercase-start words, accented words, and dense French/English word pairs.
- [x] PPTX typography ladder: refine highlighted-run rectangle origin and height for mixed highlighted text
  after kerning fixes, using Office PDF/raster inspection rather than private-deck MAE.
- [ ] Private slide 9 visible remaining problem: left-side schema geometry is visibly broken. Survey the
  involved shapes/connectors/group transforms on public-safe diagnostics, then reproduce with minimal public
  geometry fixtures before changing renderer logic.
  - [ ] New review note: right-side text expands enough to overlap the lower boxing line, and a one-line
    label wraps into two lines. Map this to text metrics, autofit, body insets, and shape bounds with public
    text-fit fixtures.
    - [x] Inspect the wrapping label: the source uses `bodyPr vert="vert270"` with `spAutoFit`.
      OOXPDF was wrapping before fitting, while Office keeps the rotated label as one fitted line.
    - [x] Add public coverage for vertical `spAutoFit` text and update shape autofit to evaluate the
      unwrapped extent before accepting word wrapping.
  - [x] Add minimal public coverage and renderer support for `curvedConnector3`; this removes the most
    obvious rectangular/straight connector artifacts in the slide-9 schema.
  - [x] Replace the approximate one-cubic curved connector path with the OOXML preset formulas used by
    `pptx-renderer`: `curvedConnector2` is one cubic with distinct controls, while `curvedConnector3`
    is two cubic segments joined at the vertical midpoint and honoring `adj1`.
  - [x] Add public Office-backed `pptx-ladder-06-curved-connector-transform-probe` for rotated/flipped
    curved connectors with `tailEnd type="arrow"`, and render curved connector arrows through the same
    Office-style arrow marker path used by straight connectors. Latest gate: MAE `0.206527`, changed16
    `0.002595`.
  - [x] Add first-class `PptxTextOrientation` coverage for known `a:bodyPr @vert` variants (`vert`,
    `vert270`, `eaVert`, `mongolianVert`, `wordArtVert`, `wordArtVertRtl`) through a public synthetic unit
    fixture and route them through orientation-aware text-frame flow instead of silently treating them as
    horizontal text.
  - [x] Fix the lower input boxes being too high by using visible run font sizes for line advance and
    middle-anchor height estimation instead of seeding every visible line with the unrelated 18 pt fallback.
  - [x] Fix the center/right text overlap with the lowermost separator. Office PDF inspection shows the three
    right-side dense columns should end on a baseline around `41.52pt`, not wrap to an extra line around
    `27.17pt`. Breakable spaces are now owned by the following word and line fitting has an Office-like
    tolerance, so slide 9 and the duplicate slide 66 diagnostics end at `41.57pt`.
  - [ ] Continue with public fixtures for vertical label text and curved connector arrowhead/control-point
    fidelity before marking the slide-9 schema as resolved. Latest private rerun
    `artifacts/private-visual/lokad-value-based/20260524-001402` shows the straight-line connector fallback
    is gone, but top-left schema text overlap and connector endpoint/arrow parity remain visible.
- [ ] Private slide 12 visible remaining problem: overlapping image on the left. Also inspect miscellaneous
  issues on the right, especially around the bottom-right content, and map them to public image/layout
  fixtures.
  - [ ] New review note: the lower-right column text has increasing vertical-position drift from top to
    bottom. Isolate public fixtures for repeated column items with inherited paragraph spacing, vertical
    anchor, and cumulative line-height rounding.
    - [x] Latest inspection showed the bottom-right cell had a leading empty paragraph containing only
      `endParaRPr`, followed by visible centered text. OOXPDF now prunes leading empty table-cell
      paragraphs when visible content follows, so the phantom line no longer pushes the real text below the
      clipped cell. Public synthetic case `PptxSyntheticTableIgnoresLeadingEmptyCellParagraph` locks this
      slide-12 pattern.
  - [x] Fix the dominant right-side table text overflow: PPTX table-cell text now wraps words to the cell
    text width instead of emitting one horizontal run per paragraph. Public synthetic unit
    `PptxSyntheticTableWrapsCellTextToColumnWidth` locks the generic behavior, and private page 12 now
    wraps the rightmost table column.
  - [ ] Continue slide-12 parity with public fixtures for exact table vertical alignment, cell text anchoring,
    image placement, and any remaining bottom-right clipping/overlap after table wrapping.
- [ ] Private slide 13 visible remaining problem: text overflows on the bottom-right content. Reproduce via
  public autofit/overflow fixtures before adjusting text layout.
  - [x] Add first-class PPTX text-column flow for `a:bodyPr numCol`/`spcCol`, remove the unsupported
    diagnostic for handled columns, and lock a public synthetic multi-column text fixture.
  - [x] Header text issue: the top-left section label font color is incorrect. Inspect whether the color
    should come from inherited placeholder/theme style, direct run properties, or master/layout shape style,
    then lock with a public placeholder/theme fixture.
    - Fixed by giving shape `fontRef` color precedence over generic default text style while keeping direct
      run color as the strongest source. Public fixture
      `PptxSyntheticShapeFontRefColorOverridesDefaultTextStyle` locks the general rule.
  - [x] Header placement issue: the top-left title line is a few pixels too high. Inspect inherited
    placeholder bounds, body insets, baseline metrics, and paragraph spacing against Office PDF output, then
    reproduce with a public title/header fixture.
    - Fixed by treating PPTX `spcPct` line spacing as a percentage over the normal PowerPoint line height
      and placing explicit-percentage baselines at the resolved line advance. Private page-13 header bounds
      now match the Office raster in the top text band.
- [ ] Private slide 7 visible remaining problem: curves on the left render as straight horizontal lines.
  Survey the shape presets/path data structurally, then reproduce with public curve/connector fixtures before
  changing renderer logic.
  - [x] Inspect slide-7 geometry structurally: the broken curves are `a:custGeom` freeform paths using
    `moveTo` plus `cubicBezTo`, not preset connectors.
  - [x] Add public synthetic coverage for cubic custom geometry paths and render DrawingML custom paths with
    `moveTo`, `lnTo`, `cubicBezTo`, `quadBezTo`, and `close` instead of falling back to rectangles.
  - [x] Re-run the private case and inspect page 7. The curve geometry now follows the Office-style smooth
    paths; remaining differences are dominated by text metrics, line/axis styling, and small placement gaps.
- [ ] Private slide 17 visible remaining problem: left-side schema geometry has issues. Inventory the involved
  shapes, groups, connectors, and transforms with public-safe diagnostics, then isolate public geometry
  fixtures.
  - [x] Inspect slide-17 geometry structurally: the broken circular loop uses four `curvedConnector2`
    quarter-curve connectors with triangle tail arrows.
  - [x] Add public synthetic `curvedConnector2` coverage and render it through the curved connector path
    instead of falling back to rectangular geometry. Private page 17 now shows the loop as curved geometry.
  - [x] Fix missing circular-flow arrowheads by resolving degenerate Bezier endpoint tangents for
    `curvedConnector2` tail markers; public synthetic connector coverage now locks the tail arrowhead.
  - [x] Fix the remaining circular-flow schema geometry by rendering `curvedConnector2` as a quarter-turn
    connector with Office-like endpoint tangents instead of an S-shaped diagonal Bezier. Public synthetic
    coverage now locks both the standalone connector and a rotated/flipped loop slice. Private rerun
    `artifacts/private-visual/lokad-value-based/20260524-091728` improved slide 17 from MAE `3.031757`,
    changed16 `0.049716` to MAE `2.945717`, changed16 `0.045530`; remaining visible drift is now dominated
    by text placement/typography around the schema, not connector geometry.
  - [x] Isolate the remaining node-label X offset as preset-shape text-rectangle behavior, not paragraph
    indentation or glyph origin. A public ellipse small-label probe reproduced the candidate-left X delta,
    while an equivalent text-box probe did not. Applying the ellipse horizontal text-rectangle inset improved
    private slide 17 from MAE `2.945717`, changed16 `0.045530`, SSIM `0.917662` to MAE `2.880535`,
    changed16 `0.044913`, SSIM `0.920176` on run
    `artifacts/private-visual/lokad-value-based/20260524-201042`.
  - [x] Continue slide-17 typography with baseline/vertical anchoring: default-line middle-anchor text-height
    estimation now uses resolved OpenType typographic line boxes instead of a CSS fallback line box. Public
    small-label text-op Y delta fell from `1.27 pt` to `0.27 pt`, and private run
    `artifacts/private-visual/lokad-value-based/20260524-202138` improved slide 17 to MAE `2.876335`,
    changed16 `0.044819`, SSIM `0.920246`.
  - [x] Add private-safe scene-schema diagnostics for slide 17. The 2026-05-24 diagnostic reports 5 master
    nodes, 4 layout nodes, and 14 top-level slide nodes; flattened scene nodes include 4 slide connectors,
    1 slide group, 3 slide pictures, 8 slide shapes, 17 text-body nodes, 31 text paragraphs, 76 text runs,
    4 flipped nodes, 1 rotated node, no tables, and no charts. This corrects the next investigation target:
    slide 17 is not only a typography case; it still includes connector/group/picture transform structure
    that must remain observable while residual text parity is improved.
  - [ ] Continue residual slide-17 text parity from public PDF evidence: the small-label probe is now tightly
    bounded, but remaining page drift still includes broader text metrics, non-label typography, and the
    typed scene structure now exposed by the private-safe diagnostic.
- [ ] Private slide 15 visible remaining problem: weird mirror artifact in rendering. Inspect transforms,
  flips, and group/image drawing order, then create public transform fixtures if coverage is missing.
  - [x] Add a public synthetic `rot=180deg` plus `flipV` text-box fixture and normalize single-flip shape
    transforms as rotation adjustments for text, keeping shape text readable instead of mirrored. This covers
    the first structural transform gap found on the slide without depending on private content.
  - [x] Add a public synthetic `a:bodyPr rot="0"` fixture for rotated/flipped shapes and let explicit
    text-body rotation override inherited shape text rotation/flips. The private slide uses this pattern to
    keep text readable inside a transformed shape.
  - [x] Re-run the private case and inspect page 15 again. The mirror artifact is gone in
    `artifacts/private-visual/lokad-value-based/20260517-115012`; remaining differences are now dominated
    by overlay placement/coverage and ordinary text-flow drift, not mirrored text.
  - [ ] Private slide 15 visible remaining problem: left-side images and their matching text items are
    vertically misaligned. Inspect picture bounds, text-frame bounds, z-order, and any shared grouping
    assumptions, then reproduce with a public image-plus-text alignment fixture.
    - [x] Fix the dominant text-frame side of the misalignment: center-anchored PPTX text frames now estimate
      wrapped line height before computing vertical anchor offset, instead of centering as if each paragraph
      were a single unwrapped line. Public synthetic unit
      `PptxSyntheticTextBoxVerticalAnchorUsesWrappedHeight` locks the generic rule, and private page 15 now
      places the text rows much closer to the matching images.
  - [ ] If slide-15 issues remain after text flow improves, isolate public fixtures for connector flips,
    picture flips, and grouped transform edge cases.
- [ ] Private slide 56 visible remaining problem: text is incorrectly boxed. Inspect whether the issue comes
  from shape fill/stroke, text highlight, clipping, or placeholder/text-frame bounds, then lock the generic
  behavior with public synthetic fixtures.
  - [x] Fix the first generic slide-56 text-list gap: symbol-font `buChar` values with OOXML charset `2`
    now map legacy byte bullets into the font private-use range before glyph lookup. Public synthetic unit
    `PptxSyntheticTextBoxMapsSymbolFontBulletCharacters` locks the rule, and private page 56 now renders
    the right-side square bullets.
  - [ ] Continue slide-56 text-list parity: bold list emphasis and red arrow callouts still differ from the
    Office PDF, so isolate those as public typography and line/arrow fixtures rather than treating the slide
    as resolved.
- [ ] Private-deck sweep loop: iterate over all `lokad-value-based` slides, keep a public-safe issue inventory,
  and for each visible problem add a minimal synthetic public case before implementing the generic fix.
  - [x] After custom arc geometry support, the private deck no longer has unsupported custom-geometry
    diagnostics. Remaining private diagnostics are now concentrated in broader public-feature families:
    approximate chart renderings, `outerShdw`/`glow` effects, image recolor (`lum`, `duotone`), and alpha
    values inside effect definitions.
  - [ ] Add public synthetic effect rungs for `outerShdw` and `glow` before attempting private-slide shadow
    parity. Start with no-blur/low-blur cases, then add blur/alpha/direction/distance variants only when
    the simpler Office PDF paths are understood.
    - [x] Add first-pass `outerShdw` rendering for ordinary preset shapes as a translucent offset duplicate
      behind the foreground shape. Public synthetic unit `PptxSyntheticOuterShadowRendersOffsetShape` locks
      the rendering and diagnostic behavior. Private effect diagnostics dropped from 17 to 1 and private
      transparency diagnostics dropped from 8 to 1; the remaining pair is the unsupported `glow` family.
    - [x] Add first-pass `glow` rendering for ordinary preset shapes as an expanded translucent duplicate
      behind the foreground shape. Public synthetic unit `PptxSyntheticGlowRendersExpandedShape` locks the
      rendering and diagnostic behavior. The private deck now has no unsupported effect or transparency
      diagnostics.
  - [ ] Add public synthetic image-recolor rungs for `a:lum` and `a:duotone`; do not remove diagnostics until
    recolored raster output is actually generated, and make the image cache key include recolor parameters.
    - [x] Add PNG/BMP raster recolor support for `a:lum` brightness/contrast and `a:duotone`, with image-cache
      keys including recolor parameters. Public synthetic units
      `PptxSyntheticPngPictureAppliesLuminanceRecolor` and
      `PptxSyntheticPngPictureAppliesDuotoneRecolor` lock the path. The private deck image-recolor
      diagnostics dropped from 3 to 1; the remaining case is a JPEG duotone image and still emits a
      content-aware unsupported diagnostic.
    - [x] Complete decoded-raster recolor coverage for `a:grayscl` and `a:biLevel` on PNG/BMP images.
      Public synthetic unit `PptxSyntheticPngPictureAppliesGrayAndBilevelRecolor` locks both modes and
      verifies recolor-specific image cache entries.
    - [x] Add a public synthetic JPEG duotone-plus-alpha diagnostic rung:
      `PptxSyntheticJpegPictureDuotoneRecolorEmitsDiagnostic` locks the remaining unsupported path without
      depending on the private deck. Keep the architectural item open: the long-term fix must be a
      dependency-free JPEG pixel path or a principled PDF color-transform/mask strategy that matches Office
      output, not a private-slide special case.
- [ ] Architecture initiative: whenever a fix touches shared PPTX behavior, improve class composition and
  first-class intermediate models rather than piling more ad hoc logic into rendering code.
- [ ] Implementation-gap initiative: when an incomplete OOXML enum, preset, transform, or layout rule is
  discovered, add it to the survey/backlog and prefer filling the general gap over patching a single deck.
- [ ] PPTX typography ladder: add Office-PDF-backed visual gates for all known `a:bodyPr @vert` variants.
  Unit coverage now routes `vert`, `vert270`, `eaVert`, `mongolianVert`, `wordArtVert`, and
  `wordArtVertRtl` through first-class orientation handling, but the ladder must still lock glyph stacking,
  anchoring, clipping, and exact baseline placement before vertical labels are considered pixel-close.
  - [x] Add vertical-orientation character-boundary wrapping for overlong Latin words. Public visual reruns:
    `pptx-ladder-04-vertical-text-270` at MAE `0.292444`, changed16 `0.002682`; and
    `pptx-ladder-04-vertical-text-port` improved to MAE `0.796678`, changed16 `0.005765`.
  - [x] Route `mongolianVert` through the same first-class vertical orientation model. The ported visual case
    now preserves declared font size instead of shrinking stacked text against a narrow rotated width. Latest
    public run `20260524-203630`: MAE `0.791842`, changed16 `0.005630`; remaining differences are still
    dominated by Office's exact stacked-glyph positioning.
  - [x] Use effective PDF text matrices for vertical text investigation. In
    `pptx-ladder-04-vertical-text-port`, the second vertical frame is position-close after CTM composition
    (candidate effective Y `480.00` vs Office `479.98`, most X deltas under `0.1 pt`), while the stacked
    frame still differs in text-operation grouping and missing/clipped columns. This keeps the next work item
    structural: glyph/column layout, not raw matrix normalization.
  - [x] Add decoded text payloads to PDF inspection, so the vertical port can be compared as text content
    plus matrices instead of raw glyph hex. In the stacked case, candidate chunks such as decoded `L S`
    expose grouping/column differences directly against Office literal payloads such as `L ` and `S`.
  - [x] Stop dropping transformed text before PDF clipping:
    baseline-vs-clip prechecks now defer rotated/flipped text to the actual transformed PDF clip instead of
    testing an unrotated baseline. The public vertical port now emits 25 candidate text operations instead
    of 20 and preserves the leading decoded `VERTIC...` content; one Office operation and column ordering
    remain open.
  - [ ] Continue vertical text parity with Office text-operation inspection: stacked-letter orientation,
    column order, per-column x positions, and baseline placement remain visibly approximate.
- [ ] For every generic capability fixed from a private slide, add a small public synthetic test. Do not
  derive public fixtures from private slide content.
- [ ] Run `pwsh tools/CheckPrivateCase.ps1 -Case private-cases/lokad-value-based.json` after each scoped PPTX
  fix and summarize only counts, diagnostics, worst-page numbers, and private ratings.

### DOCX Feature Survey

- [ ] Pagination: Word-compatible line height, paragraph spacing collapse, keep-with-next,
  keep-lines-together, widow/orphan control, manual page/column breaks, section breaks, and page size
  rounding.
- [ ] Text layout: tabs, tab stops, indents, hanging indents, justification, hyphenation, nonbreaking spaces,
  soft hyphens, field result handling, superscript/subscript, and baseline shifts.
- [ ] Fonts: bold/italic face selection, fallback fonts, embedded fonts, complex scripts, bidirectional text,
  OpenType features, and symbol fonts.
- [ ] Numbering: exact level text expansion, hanging indents, bullets from symbol fonts, restart rules,
  multi-level lists, custom number formats, and style-linked numbering.
- [ ] Tables: auto-fit, preferred widths, nested tables, merged cells, vertical merges, cell margins,
  borders, shading, row height rules, repeating header rows, and page breaks inside tables.
- [ ] Images/drawings: anchored/floating drawings, wrap modes, relative positioning, cropping, rotation, text
  boxes, shapes, SmartArt, charts, and drawing canvases.
- [ ] Headers/footers: first/odd/even variants, section-specific variants, distance from edge, fields beyond
  `PAGE`, total page count, and dynamic date/doc properties.
- [ ] Fields: `PAGE`, `NUMPAGES`, `DATE`, `REF`, `HYPERLINK`, `TOC`, `SEQ`, form fields, and cached-field
  fallback semantics.
- [ ] Footnotes/endnotes/comments: render bodies or emit precise diagnostics with usable fallback behavior.
- [ ] Tracked changes: choose final, original, or marked-up view explicitly and document the behavior.
- [ ] Multi-column layout, text boxes, sidebars, bookmarks, hyperlinks, outlines, and document properties.

### DOCX Synthetic Fidelity Ladder

Build a DOCX ladder comparable to the PPTX ladder. Each rung must be public, synthetic, minimal,
Office-PDF-inspected, visually gated when close, and free of private content.

- [x] Ladder 0: blank document, page size, margins, deterministic PDF structure, and no diagnostics.
- [ ] Ladder 1: plain paragraphs with Word reference baselines, line height, paragraph before/after spacing,
  and page margins. Subcases `docx-ladder-01-plain-paragraph`, `docx-ladder-01-line-height`, and
  `docx-ladder-01-paragraph-spacing` are gated; remaining subcases should isolate multi-line flow and
  wrapping.
- [ ] Ladder 2: run styling: bold/italic face selection, underline, strikethrough, color, highlight,
  superscript/subscript, and font fallback.
- [ ] Ladder 3: paragraph layout: tabs/tab stops, indents, hanging indents, alignment, spacing collapse,
  nonbreaking spaces, soft line breaks, and manual page breaks.
- [ ] Ladder 4: numbering and bullets: label area, hanging indents, level text expansion, restart/start
  rules, symbol bullets, and multi-level lists.
- [ ] Ladder 5: tables: grid widths, preferred widths, cell margins, row height from content, borders,
  shading, vertical alignment, merged cells, and page breaks.
- [ ] Ladder 6: images and drawings: inline images, anchored/floating drawings, wrap modes, cropping,
  rotation, shapes, and unsupported drawing diagnostics.
- [ ] Ladder 7: headers/footers and fields: first/odd/even variants, section variants,
  PAGE/NUMPAGES/date/property fields, and distance from edge.
- [ ] Revisit DOCX unit tests under the Office-PDF-first workflow: prefer assertions about parsed model,
  diagnostics, page counts, public API, and visual gates over fragile text-coordinate/operator expectations
  where Word PDF inspection gives the real behavior.

### DOCX 16-vs-17 Page Mismatch Plan

- [ ] Select the first private page and exhaust it before moving on: inspect reference vs candidate, list
  every visible failure, implement or diagnose each generic gap, rerender, and repeat until the page is
  acceptable or every remaining issue is explicitly planned.
- [ ] Continue one private page at a time after the first page. Prefer pages with obvious table/pagination
  failures before using document-wide aggregate metrics.
- [ ] Maintain a private per-page checklist with only public-safe fields: page number, private rating,
  missing content, wrong order, wrong placement, wrong sizing, wrong text layout, table defects, pagination
  drift, and unsupported features lacking diagnostics.
- [ ] Add an internal DOCX layout trace mode that records public-safe per-page counts and consumed vertical
  space by block kind, so private runs can locate where candidate pagination drifts without exposing text.
- [x] Extend DOCX diagnostics to inspect styles and numbering parts, not just direct `word/document.xml`
  elements, so style-level spacing, keep rules, indents, table styles, and numbering layout risks are visible.
- [ ] Implement style-derived paragraph spacing accurately, including before/after values,
  `contextualSpacing`, `beforeAutospacing`/`afterAutospacing`, and Word-like adjacent paragraph spacing
  collapse.
- [ ] Implement paragraph and numbering indents: left/right/first-line/hanging indents from paragraph styles
  and numbering levels, with corresponding wrapping-width changes.
- [ ] Improve numbering layout: render labels in their own hanging-indent area, support level text expansion
  beyond the current simple label prefix, and honor restart/start rules.
- [ ] Improve table layout accumulation: preferred table widths, cell widths, row minimum height from
  content, cell vertical alignment, cell margins, and repeating header rows.
- [ ] Revisit keep rules only after layout tracing exists: support style-derived `keepNext`, `keepLines`, and
  widow/orphan control with synthetic tests and private page-count checks.
- [ ] Reattempt manual page/column break support with a parser change that does not alter paragraphs when no
  matching break exists; previous paragraph-splitting attempts changed the private page count and were
  reverted.
- [ ] For every generic capability fixed from a private page, add a small public synthetic test. Do not
  derive public fixtures from private page content.
- [ ] After each scoped fix, run `pwsh tools/CheckPrivateCase.ps1 -Case
  private-cases/user-requirements-spec.json` and record only page counts, aggregate metrics, diagnostics, and
  worst-page numbers.

### DOCX Table Recovery Plan

- [ ] Add a DOCX table inventory/trace mode that reports public-safe table metrics per table: row count,
  column count, grid columns, preferred table width, cell width declarations, row height declarations, header
  rows, vertical alignment, margins, borders, shading, grid spans, vertical merges, and page index.
- [ ] Select representative private tables for repeated inspection: one simple table, one typical dense
  table, and one worst table. Record only table ordinal/page, private rating, and public-safe feature gaps.
- [ ] Fix the table layout model before cosmetic styling: resolve `tblGrid`, `tblW`, `tcW`, page content
  width, percentage/auto widths, and grid scaling consistently.
- [ ] Compute row heights from actual cell content: wrap text within cell width, include cell margins,
  respect explicit `trHeight` rules, and avoid the current fixed/default row-height behavior for
  content-heavy rows.
- [ ] Render cell text as paragraphs instead of flattened cell text: preserve paragraph breaks, basic run
  styling, numbering/bullets inside cells, alignment, and line spacing.
- [ ] Implement table and cell styling: table styles, conditional first/header row formatting, cell shading,
  per-edge borders, border widths/colors, and vertical alignment.
- [ ] Implement structural table features: horizontal merges (`gridSpan`), vertical merges (`vMerge`),
  repeating header rows across page breaks, and page-break behavior inside rows.
- [ ] Add synthetic public tests for each table capability before using the private document as evidence;
  never derive fixtures from private table content.

### PDF/Infrastructure

- [x] Add a PDF inspection tool for Office/candidate PDFs that lists objects and extracts decodable streams
  for content-operator inspection.
- [ ] Audit current PDF generation patterns against Office reference PDFs: text grouping, text matrices,
  clipping regions, image masks, transparency state, path construction, stroke/fill order, resource
  naming/reuse, and page content stream organization.
- [x] PDF font resource emission now merges ToUnicode coverage for repeated embedded-font resources, so
  independently rendered chart/table/slide text cannot lose glyph mappings when they share the same base
  font resource key.
- [ ] Improve PPTX text-line emission toward Office-like text objects: PPTX now emits positioned `TJ` arrays
  even when a line has no explicit kerning/tracking adjustment, matching Office's common text-object pattern
  without changing raster output. Continue by reducing unnecessary run splitting around spaces when a line
  can be emitted as one positioned text object.
- [x] Split PPTX text layout and text drawing into `PptxRenderer` partial files so typography work no longer
  lands in the monolithic renderer body.
- [ ] Continue splitting `PptxRenderer` by responsibility: style resolvers. Text, shapes/presets, pictures,
  table layout, diagnostics, color helpers, shared geometry/types, table-style helpers, chart-fallback
  helpers, and z-order dispatch now live in dedicated partials. Keep splits mechanical and behavior-neutral
  unless a scoped fidelity fix is part of the same area.
- [ ] Refactor PDF rendering primitives where Office-like structure is more robust for fidelity, while
  preserving deterministic output and keeping `src/Lokad.OoxPdf` dependency-free.
- [ ] Add PDF hyperlinks, outlines/bookmarks, metadata, and optional tagged-PDF structure if needed by
  consumers.
- [ ] Add font subsetting to reduce output size while keeping deterministic output.
- [ ] Add image deduplication and compression choices for large decks.
- [x] Improve diagnostics severity model so embedded-image omissions are distinguishable from harmless
  approximations.
- [x] Add visual comparison support for dimension-near-matches, so a 1-pixel raster rounding mismatch can
  still produce pixel metrics.
- [x] Add private-case summary tooling that reports page count, dimension mismatches, diagnostics grouped by
  feature, and worst visual pages without exposing content.

## Next Implementation Targets

1. Make Office-PDF structure the primary fidelity oracle. Pixel metrics stay useful, but every serious fix
   should first ask what Office emitted: text matrices, glyph advances, clipping, image XObjects, paths,
   transparency groups, resources, and drawing order. Extend `PdfInspect`/comparison tooling when a mismatch
   cannot be explained structurally.
2. Make the PPTX scene/render-context architecture authoritative in small slices. `PptxScene` now models
   slides, backgrounds, nodes, bounds, text bodies, picture intent, shape styles/geometry, group transforms,
   chart relationship ids, resolved chart part targets, chart XML, chart palettes, chart plot summaries,
   chart plot attributes, chart series summaries including scatter/bubble data channels, chart series and
   point styles, chart data-label metadata/text/shape/per-label overrides including default-run bold/italic
   style flags, chart axis catalogs with scaling/units/gridlines/label options, titles, and legends, while
   `PptxRenderContext` owns package, theme, inheritance, relationships, image cache, and diagnostics. Keep
   retiring XML fallbacks family by family: next slices should promote
   table layout/style records, chart series/axis/layout records, and remaining text/layout inputs into typed
   models while preserving source XML only as an escape hatch for unported features.
3. Systematically retire heuristics. When a magic constant, special-case branch, or private-slide coordinate
   rule is touched, either replace it with an OOXML/Office-derived rule or write down why the approximation is
   temporary. Prefer named geometry/style/text models over narrow `if this preset/slide shape` logic.
4. Continue the `pptx-renderer` track as architecture input and public test source. The local reference lives
   at `C:\Users\JoannesVermorel\code\pptx-renderer`; port lessons and minimal public Office-oracle cases into
   `ooxpdf`, starting with typography, then shape/preset, layout/composition, images, tables, charts, and
   SmartArt. Do not vendor code or private assets.
5. Push typography through structural alignment, not pixel nudging. PPTX Ladder 4 should tighten style
   cascade, font resolution, paragraph metrics, baselines, advances, tracking, highlights, underlines,
   bullets, and paragraph flow against Office PDF operators and public visual cases.
6. Promote shapes, connectors, tables, and charts to resolved models before broad fidelity patches. Shape work
   should converge on Office-like preset paths, connector tangent/marker rules, fills/strokes/effects, and
   group transforms. Table and chart work should resolve layout and styling into typed models before PDF
   emission so fixes are reusable across cases.
7. Keep the validation ladder scalable and honest. Cached oracles, fast/slow/oracle lanes, richer visual
   metrics, and public fixture families should make hundreds of public cases practical. Private decks remain
   gap discovery and acceptance evidence only. DOCX should stay deferred until the PPTX public ladder and
   `pptx-renderer`-derived cases are strong enough to avoid spreading architectural debt across formats.

## Surprises & Discoveries

- Observation: The chart title scene model preserved title text and `txPr` style, but not the sibling title
  structure that will matter for Office-like placement: `c:overlay`, `c:layout/c:manualLayout`, and `c:spPr`.
  Evidence: `PptxSceneChartTitle` had only `Text`, `IsAutoDeleted`, and `TextStyle`; the scene-builder test now
  locks title overlay, manual-layout factors, fill, and stroke as model data before any renderer consumption.
- Observation: Axis titles are structurally the same chart-title sub-tree under each chart axis, but were not
  available to rendering or inspection as axis-owned scene data.
  Evidence: `PptxSceneChartAxis` preserved tick label style and number format but had no title record; the
  scene-builder fixture now locks category-axis and value-axis titles through `PptxSceneChartAxis.Title`.
- Observation: Chart legends already had typed position, overlay, visibility, and text style, but not the
  sibling layout and shape properties Office can use to place and frame a legend.
  Evidence: `PptxSceneChartLegend` now preserves `c:layout/c:manualLayout` and `c:spPr` fill/stroke data, and
  the scene-builder fixture locks those values before any renderer placement change consumes them.
- Observation: Generic chart text position buckets were too coarse for long-term structural alignment, but
  overly broad legend-container detection also misclassifies category-axis text when charts use clipping boxes
  above the plot.
  Evidence: `ClassifyPdfChartText.ps1` now emits explicit category/value/data/legend/title roles where
  geometry supports them, limits container-based legend detection to right-side containers, and gates
  `CategoryAxisTickLabel` on the public clustered-column case. A line-marker visual probe still fails its
  pre-existing pixel MAE threshold (`3.640602` actual vs `2.9` limit) before a text gate can run, so it remains
  a separate chart-fidelity cleanup target rather than a tolerance change hidden inside this slice.
- Observation: Per-point chart data labels preserved visibility/text/style overrides, but still dropped the
  label-local `c:layout/c:manualLayout` subtree that Office can use for explicit label placement.
  Evidence: `PptxSceneChartDataLabelOverride` now carries `PptxSceneChartManualLayout`, and the scene-builder
  fixture locks per-label `x/y/w/h` factors before the renderer attempts any Office-aligned label placement.
- Observation: The slide-17 schema issue was not only "curvedConnector2 is unsupported"; after initial
  support and arrowhead tangent fixes, the remaining visible problem came from using an S-shaped one-cubic
  connector where Office behaves like a quarter-turn loop segment.
  Evidence: the current implementation now uses the standard quarter-arc control constant
  `4 / 3 * (sqrt(2) - 1)` in `PptxRenderer.Shapes.cs`, and public tests include a rotated/flipped
  `curvedConnector2` loop slice.
- Observation: Full PDF inspection of the large private `lokad-value-based` candidate/reference PDFs can time
  out and produce excessive intermediate output.
  Evidence: a full `tools/InspectPdf.ps1` run against those PDFs exceeded the local command timeout. Prefer
  focused public synthetic PDFs or already extracted page/operator evidence for hot loops.
- Observation: The remaining slide-17 schema delta is typography/text placement, not chart or connector
  geometry.
  Evidence: private-safe slide inventory shows no charts, tables, or groups on the page, but many text bodies
  and a few rotated/flipped transforms. Page-aware PDF text inspection shows Office and candidate differ in
  text-operation count on the page, with the candidate still emitting 41 text operations after the no-wrap
  fix. The four small no-wrap node labels are left-aligned in OOXML with default insets, so the residual
  horizontal offset should be investigated as Office font advance/glyph origin/paragraph metric behavior, not
  as a hard-coded center-alignment correction.
- Observation: Slide 17's remaining small-label horizontal offset is not caused by local or inherited
  paragraph indentation.
  Evidence: private-safe slide XML inspection found only `algn="l"` on the four small node-label paragraphs,
  no local `marL` or `indent`, and no shape-level list-style defaults. A 2026-05-24 layout diagnostic after
  inherited-indent support shows the same four labels resolving `MarginLeft=0` and `HangingIndent=0`, with
  emitted span X at shape X plus the default text inset. Continue with glyph-origin/font-metric/PDF text
  operation alignment instead of another paragraph-margin patch.
- Observation: The slide-17 small-label X delta comes from auto-shape preset text rectangles.
  Evidence: a public text-box small-label probe did not reproduce the X delta, but the same label in an
  ellipse auto-shape reproduced the candidate-left offset. Applying the ellipse horizontal text-rectangle
  inset reduced the public text-op X delta to `0.03 pt` and improved private slide 17. Symmetric vertical
  ellipse insets worsened the public baseline, so vertical preset text-rectangle/baseline behavior remains
  open and must be solved from Office PDF evidence rather than a paired inset assumption.
- Observation: The same slide-17 public small-label probe showed that vertical centering was overestimating
  a one-line default-spaced label as a CSS `1.2x` line box.
  Evidence: switching only the middle-anchor text-height estimator to the resolved font's OpenType
  typographic line box reduced the public text-op Y delta from `1.27 pt` to `0.27 pt` and improved the
  private deck/page metrics. Line layout advances remain unchanged; this is an anchoring estimate fix, not
  a global line-height rewrite.
- Observation: The historical private evidence block should not be aggressively trimmed without first
  preserving its facts elsewhere.
  Evidence: a local check on 2026-05-24 found that most older `artifacts/private-visual/lokad-value-based`
  run directories referenced by the plan are no longer present, so the public-safe metrics and findings in
  this file may be the only convenient record of those runs.
- Observation: OOXPDF already has the start of the desired PPTX intermediate architecture, but it is not yet
  the source of truth for rendering.
  Evidence: `PptxScene.cs` defines scene, slide, node, bounds, paragraph, and run records, while
  `PptxRenderer.cs` and the shape/text/table/chart partials still dispatch mostly from raw XML plus
  `PptxRenderContext`.
- Observation: Scene-backed ordered rendering can be introduced without changing leaf PDF emission.
  Evidence: `RenderPages` now builds a `PptxScene`, carries `PptxSceneSlide` through `PptxRenderContext`, and
  ordered slide rendering iterates scene nodes and group children while shape/picture/table/chart renderers
  still receive the same source XML elements as before.
- Observation: The scene model needs to preserve source-unit values, not only convenient point conversions, if
  it is going to feed Office-like renderer transforms.
  Evidence: `PptxSceneBounds` now keeps EMU coordinates and extents plus point conversion properties; ordered
  picture rendering uses those EMU bounds directly through the existing group transform math.
- Observation: `pptx-ladder-04-typography-accent-spacing-probe` used to be a false structural gap because
  the checked-in PPTX contained mojibake despite a correct UTF-8 generator source.
  Evidence: after regenerating only that fixture, `slide1.xml` contains true accented French text. The
  2026-05-24 run `20260524-214440` has MAE `0.516824`, changed16 `0.005876`, SSIM `0.919226`, and the new
  text gates pass: four Office text lines and four candidate text lines, matching decoded text, and
  max line-start delta `0.00 pt`. The remaining raster drift is glyph-shape/font-rendering parity, not text
  grouping or placement.
- Observation: The glyph-run emission boundary can now split one OOXML run across multiple PDF font resources
  when glyph fallback is required, but fallback selection is still a simple discovered-font search.
  Evidence: `PptxSyntheticTextBoxSplitsMissingGlyphsToFallbackFont` locks an Arial run containing a CJK code
  point into multiple emitted resources while preserving decoded glyph order. The next refinement should make
  script-aware fallback preference explicit and compare it against Office PDFs before treating accented/CJK
  grouping as locked.
- Observation: The previous slide-17 private inventory was too coarse for schema work.
  Evidence: the new private-safe scene snapshot for slide 17 reports top-level slide connectors, a group,
  pictures, and flattened shape/text counts that were not explicit in the older inventory summary. Future
  slide-17 work should use scene snapshots plus PDF text-operation evidence before concluding that a remaining
  difference is purely typography.
- Observation: Public chart PDFs already show large structural differences before any raster analysis.
  Evidence: on `pptx-ladder-11-chart-column-clustered-port`, generic PDF graphics comparison reports Office
  reference and candidate path-operation counts of `32` and `34`, with deltas in operation kind, bounds, and
  stroke/fill sequence. This confirms that chart convergence needs PDF-structure evidence, not just tighter
  raster thresholds.
- Observation: The first semantic chart classifier shows that raw PDF structure differs by producer enough
  that classification must remain evidence-driven.
  Evidence: on the same public clustered-column chart, the candidate has eight obvious horizontal gridlines
  and a derived `PlotBoxCandidate` from their bounds, while the Office PDF exposes fewer line-like strokes and
  more repeated clipping boxes/Bezier-filled regions. The next classifier work should survey multiple chart
  families before turning any bucket into a strict gate.
- Observation: Explicit semantic gridline classification exposes a chart fidelity gap that the current
  clustered-column manifest deliberately does not gate.
  Evidence: rerunning `ClassifyPdfChartGraphics.ps1` on the public clustered-column inspected PDFs showed
  zero Office `HorizontalGridlineCandidate` records and six candidate records inside the Office-aligned
  `AxisPairPlotBoxCandidate`. The existing public visual gate still passes because it only compares
  `AxisPairPlotBoxCandidate`, which confirms the new gridline kind is available for diagnosis without
  destabilizing unrelated gates.

## Decision Log

- The library remains dependency-free. Third-party packages are not allowed in `src/Lokad.OoxPdf`.
- Office and PDFium remain validation-only under `tools/`.
- Private documents remain under ignored `private-cases/`; generated private artifacts remain under ignored
  `artifacts/private-visual/`.
- Public notes from private documents must be anonymized to feature gaps and metrics only.
- Diagnostics must prefer continued conversion over crashing, but omitted visible content must not be treated
  as acceptable final behavior.
- Newly discovered OOXML chart structure should be preserved in `PptxScene` before rendering is changed. This
  keeps Office semantics inspectable and prevents missing title, legend, axis, or data-label behavior from
  being replaced with narrow placement heuristics.
- Pixel metrics are late-stage regression evidence only. Until selected private slides/pages are mostly
  visually correct, do not use MAE or changed-pixel ratios to prioritize work or judge acceptability.
- Office-exported PDFs are the primary fidelity reference. Raster metrics are useful gates after manual/agent
  inspection confirms the candidate is targeting the same Office behavior.
- Emulate Office's observable rendering strategy where it matters for fidelity: PDF operator structure,
  resource usage, clipping, transparency, image placement, and text positioning should move toward
  Office-like patterns when practical. Do not depend on Office at runtime or claim byte-for-byte PDF
  equivalence.
- PPTX fidelity is bottom-up: minimal public synthetic fixtures are made close to pixel-perfect and gated
  before larger public combinations or private documents matter.
- Private PPTX pages may regress while lower public rungs are rebuilt. Until the public ladder is
  feature-complete enough, private MAE and changed-pixel ratios are smoke evidence only, not implementation
  targets.
- DOCX fidelity should move to the same Office-PDF-first public ladder as PPTX. Private pages remain
  acceptance evidence and gap discovery, not the main implementation driver.
- Long-term PPTX work should optimize for structural PDF alignment with Office, not isolated raster nudges.
  Pixel-perfect output remains the outcome target, but the mechanism should be typed OOXML resolution and
  Office-like PDF operators wherever the structure is observable.
- Decision: Execute the next long-term push as seven linked architecture tracks, starting with chart
  structural oracle tooling before deeper chart layout changes.
  Rationale: The chart renderer now has many named fallback metrics, but naming heuristics is only inventory.
  A chart-specific Office/PDF structural comparison path is the evidence layer needed to delete those
  heuristics safely and move chart output toward Office-like plot, axis, label, and resource structure.
  Date/Author: 2026-05-25 / Codex.
- Decision: Add generic graphics-operation inspection before adding chart-semantic classification.
  Rationale: Plot areas, axes, gridlines, markers, and legend swatches all appear as ordinary PDF path,
  clip, stroke, and fill operations. Capturing those primitives once in `PdfInspect` makes the tooling useful
  beyond charts and avoids baking chart assumptions into the PDF parser.
  Date/Author: 2026-05-25 / Codex.
- Decision: Keep chart semantic classification as a separate script over inspected PDF graphics JSON.
  Rationale: The same inspected primitives can support multiple chart-family classifiers, and keeping the
  semantic layer out of `PdfInspect` lets the plan add, revise, or remove chart-specific candidates without
  changing the lower-level PDF parser.
  Date/Author: 2026-05-25 / Codex.
- Decision: Make chart gridline candidates opt-in semantic records rather than default visual gates.
  Rationale: The clustered-column evidence shows a real Office/candidate mismatch: the candidate has six
  plot-spanning horizontal strokes where Office exposes no matching gridline strokes. That should drive a
  renderer investigation, but turning it into a gate before understanding Office's PDF strategy would freeze
  the current mismatch instead of eliminating it structurally.
  Date/Author: 2026-05-25 / Codex.

## Outcomes & Retrospective

- The project is past the initial vertical slice: the library, CLI, console tests, public visual validation,
  private-case validation, PDF inspection, and package output are all in place. Current work is fidelity and
  architecture, not bootstrapping.
- The active PPTX direction is bottom-up and `pptx-renderer`-informed: preserve public Office-backed fixtures,
  move renderer behavior toward explicit intermediate models, and use private decks only to discover generic
  gaps.
- The latest slide-17 schema work resolved the connector-geometry portion of that private issue. Current
  page-aware PDF evidence keeps the next target on typography/text placement: text wrapping, inherited
  paragraph indentation, ellipse auto-shape horizontal text rectangles, and font-metric middle anchoring are
  now structurally modeled. Slide 17 improved on runs `20260524-201042` and `20260524-202138`, then stayed
  stable on `20260524-204001`; remaining evidence points to broader text metrics rather than connector
  geometry, paragraph margins, or small-label preset-shape origins.

## Concrete Steps

Work from `C:\Users\JoannesVermorel\code\ooxpdf` in PowerShell. Use `--tl:off` on .NET commands to avoid
dynamic terminal logger output.

For normal validation after code changes, run:

```powershell
dotnet restore Lokad.OoxPdf.slnx --tl:off -v minimal
dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal
dotnet run --project tests/Lokad.OoxPdf.Tests --tl:off --nologo -v minimal
dotnet pack src/Lokad.OoxPdf/Lokad.OoxPdf.csproj --tl:off --nologo -v minimal --no-restore
```

For a public visual case or family, run:

```powershell
pwsh tools/CheckVisualCase.ps1 -Case visual-cases/cases/<case>/case.json
pwsh tools/CheckVisualFamily.ps1 -Family <family>
```

For PDF structural comparison, inspect both PDFs and compare the emitted JSON:

```powershell
pwsh tools/InspectPdf.ps1 -InputPdf <reference.pdf> -OutputDirectory artifacts/<case>-ref-inspect
pwsh tools/InspectPdf.ps1 -InputPdf <candidate.pdf> -OutputDirectory artifacts/<case>-cand-inspect
pwsh tools/ComparePdfTextOperations.ps1 -Reference artifacts/<case>-ref-inspect/text-operations.json -Candidate artifacts/<case>-cand-inspect/text-operations.json
pwsh tools/ComparePdfGraphicsOperations.ps1 -Reference artifacts/<case>-ref-inspect/graphics-operations.json -Candidate artifacts/<case>-cand-inspect/graphics-operations.json -MatchByBounds
pwsh tools/ClassifyPdfChartGraphics.ps1 -InputPath artifacts/<case>-ref-inspect/graphics-operations.json -Output artifacts/<case>-ref-inspect/chart-structures.json
pwsh tools/ClassifyPdfChartGraphics.ps1 -InputPath artifacts/<case>-cand-inspect/graphics-operations.json -Output artifacts/<case>-cand-inspect/chart-structures.json
```

For private PPTX acceptance evidence after a scoped public fix, run:

```powershell
pwsh tools/CheckPrivateCase.ps1 -Case private-cases/lokad-value-based.json
```

## Validation

Latest public validation:

```powershell
dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal
dotnet run --project tests/Lokad.OoxPdf.Tests --tl:off --nologo -v minimal
dotnet pack src/Lokad.OoxPdf/Lokad.OoxPdf.csproj --tl:off --nologo -v minimal --no-restore
```

Current expected test result:

```text
207 passed, 0 failed, 0 skipped
```

Latest private PPTX acceptance baseline:

```text
lokad-value-based / 20260524-235547: 84/84 compared pages, 0 dimension mismatches,
deck MAE 9.005915, changed16 0.116052, only PPTX_UNSUPPORTED_IMAGE_RECOLOR.
Page 17: MAE 2.880739, RMSE 19.298084, changed16 0.044888, changed32 0.035257,
SSIM 0.920083, foreground histogram correlation 0.999858.
```

Latest public small-label origin probe:

```text
pptx-ladder-04-typography-small-label-origin-probe / 20260524-204739:
MAE 0.005514, changed16 0.000073; PDF text-op X delta 0.03 pt, Y delta 0.27 pt;
decoded text gate enabled.
```

Latest public chart structural-tooling validation:

```text
PdfInspect graphics extraction build: succeeded, 0 warnings, 0 errors.
ComparePdfGraphicsOperations reference-vs-reference on
pptx-ladder-11-chart-column-clustered-port / 20260523-224316: 32 reference operations,
32 candidate operations, 0 deltas.
ComparePdfGraphicsOperations Office-vs-ooxpdf on the same public chart: 32 reference operations,
34 candidate operations, 34 deltas, proving the tool exposes chart PDF-structure differences.
ClassifyPdfChartGraphics on the same inspected PDFs: 28 Office chart-structure records and
35 candidate chart-structure records, including candidate `HorizontalLine`, `VerticalLine`, and
derived `PlotBoxCandidate` structures.
ComparePdfGraphicsOperations reference-vs-reference over classified chart structures: 28 reference
structures, 28 candidate structures, 0 deltas.
Full console suite: 204 passed, 0 failed, 0 skipped.
dotnet pack: created artifacts/nuget/Lokad.OoxPdf.0.1.0.nupkg.
Public visual case `pptx-ladder-11-chart-column-clustered-port` passed through the updated harness at
artifacts/visual/pptx-ladder-11-chart-column-clustered-port/20260525-001549.
Chart scene-order slice: `dotnet run --project tests\Lokad.OoxPdf.Tests --tl:off --nologo -v minimal -- --group pptx-model --skip-slow`
passed with 13 passed, 0 failed, 1 skipped; `dotnet run --project tests\Lokad.OoxPdf.Tests --tl:off --nologo -v minimal -- --group pptx-charts --skip-slow`
passed with 8 passed, 0 failed, 0 skipped; full console suite passed with 204 passed, 0 failed, 0 skipped;
`dotnet pack src\Lokad.OoxPdf\Lokad.OoxPdf.csproj --tl:off --nologo -v minimal --no-restore` succeeded.
Classifier family probe: inspected and classified Office/candidate PDFs from public pie, doughnut, radar,
scatter-cluster, and line-marker chart cases under ignored `artifacts/tmp-chart-family-inspect/`. The updated
classifier emits `PlotAreaClipBoxCandidate`, `AxisPairPlotBoxCandidate`, and `PolarPlotBoxCandidate`; strict
reference-vs-reference comparison of those derived kinds passed for all five sampled families.
Chart text oracle probe: `ClassifyPdfChartText.ps1` classified public pie, doughnut, radar, scatter-cluster,
and line-marker text operations relative to derived plot boxes; strict reference-vs-reference comparison of
the chart text buckets passed for all five sampled families. A temporary ignored visual manifest
`artifacts/tmp-chart-text-gate/case.json` exercised the new `maxChartTextStructurePositionDelta` harness path
successfully at `artifacts/visual/tmp-chart-text-gate-line-markers/20260525-002824`.
Chart gridline semantic probe: `ClassifyPdfChartGraphics.ps1` on the clustered-column inspected PDFs emitted
zero Office `HorizontalGridlineCandidate` structures and six candidate structures, while preserving the
passing `AxisPairPlotBoxCandidate` gate. The public visual case passed through the updated classifier at
`artifacts/visual/pptx-ladder-11-chart-column-clustered-port/20260525-015313`.
```

Latest public vertical text probes:

```text
pptx-ladder-04-vertical-text-270 / 20260524-205046:
MAE 0.247899, changed16 0.002204.

pptx-ladder-04-vertical-text-port / 20260524-204933:
MAE 0.791842, changed16 0.005630; effective-matrix/decoded-text inspection shows the second rotated frame is
position-close and the stacked frame no longer drops leading `VERTIC...` content, while
`mongolianVert` glyph grouping/columns remain open.
```

Latest public typography structural probes:

```text
pptx-ladder-04-nonbreaking-space / 20260524-212024:
decoded PDF text gate passed; 2 reference and 2 candidate text operations; B x delta 0.09 pt.

pptx-ladder-04-typography-accent-spacing-probe / 20260524-214440:
corrected true-accent fixture passed; MAE 0.516824, changed16 0.005876, SSIM 0.919226. PDF text line-start
gate passed with four reference and four candidate text lines, decoded text parity, and max start delta
0.00 pt.

glyph-run ownership slice / 2026-05-24:
`dotnet run --project tests\Lokad.OoxPdf.Tests --tl:off --nologo -v minimal -- --group pptx-typography --skip-slow`
passed with 66 passed, 0 failed, 2 skipped after adding per-glyph fallback font splitting and
typeface/resource snapshots to glyph-run inspection.
Non-slow tests passed with 189 passed, 0 failed, 7 skipped; package build succeeded; full tests passed with
196 passed, 0 failed, 0 skipped.

scene-inspection slice / 2026-05-24:
`PptxSceneBuilderBuildsResolvedNodeLists` passed after adding private-safe scene snapshots; non-slow tests
passed with 189 passed, 0 failed, 7 skipped; `dotnet pack src\Lokad.OoxPdf\Lokad.OoxPdf.csproj --tl:off
--nologo -v minimal --no-restore` succeeded; full tests passed with 196 passed, 0 failed, 0 skipped.

ordered text-dispatch split / 2026-05-24:
`PptxSceneBuilderBuildsResolvedNodeLists` passed, and the non-slow suite passed with 189 passed, 0 failed,
7 skipped. A first parallel attempt at the non-slow suite failed during build because Windows held the
debug DLL for writing; rerunning serially passed, so this was an environment file lock rather than a code
failure. `dotnet pack` succeeded, and the full suite passed with 196 passed, 0 failed, 0 skipped.

unknown graphic-frame fallback / 2026-05-24:
`PptxUnsupportedFeaturesEmitDiagnostics` now locks `PPTX_UNSUPPORTED_GRAPHIC_FRAME` separately from chart,
table, and SmartArt classification. `PptxSceneBuilderBuildsResolvedNodeLists` passed after threading source
part names through ordered scene dispatch, and the non-slow suite passed with 189 passed, 0 failed, 7 skipped.
One parallel targeted-test attempt hit a SourceLink file lock; the same tests passed when rerun serially.
`dotnet pack` succeeded, and the full suite passed with 196 passed, 0 failed, 0 skipped.
Private run `artifacts/private-visual/lokad-value-based/20260524-221657` compared 84/84 pages with zero
dimension mismatches, deck MAE `9.005915`, changed16 `0.116052`, and only
`PPTX_UNSUPPORTED_IMAGE_RECOLOR`; slide 17 measured MAE `2.880739`, changed16 `0.044888`, SSIM `0.920083`.

ordered rendering with unknown graphic frames / 2026-05-24:
`PptxSyntheticUnknownGraphicFrameDoesNotDisableSiblingOrder` now locks that slide-level unknown graphic
frames do not disable ordered scene rendering for known siblings, while `PptxUnsupportedFeaturesEmitDiagnostics`
continues to lock a single slide-scoped `PPTX_UNSUPPORTED_GRAPHIC_FRAME` diagnostic. The non-slow suite passed
with 190 passed, 0 failed, 7 skipped. After removing the old ordered-render eligibility branch, `dotnet pack`
succeeded and the full suite passed with 197 passed, 0 failed, 0 skipped.

obsolete PPTX fallback traversal removal / 2026-05-24:
After deleting the private XML-wide picture, shape, and chart traversal entry points, the renamed
`PptxSyntheticGroupedTableUsesGroupTransformWithUnknownGraphicFrame` passed and the non-slow suite passed
with 190 passed, 0 failed, 7 skipped. `dotnet pack` succeeded and the full suite passed with 197 passed,
0 failed, 0 skipped.

typed scene table font prepass / 2026-05-24:
`PptxSyntheticTableRendersGridAndText` and
`PptxSyntheticGroupedTableUsesGroupTransformWithUnknownGraphicFrame` passed after replacing the raw XML table
font prepass with a typed scene-node traversal. The non-slow suite passed with 190 passed, 0 failed,
7 skipped. `dotnet pack` succeeded and the full suite passed with 197 passed, 0 failed, 0 skipped.
Private run `artifacts/private-visual/lokad-value-based/20260524-223120` stayed stable: 84/84 compared pages,
zero dimension mismatches, deck MAE `9.005915`, changed16 `0.116052`, only
`PPTX_UNSUPPORTED_IMAGE_RECOLOR`, and slide 17 MAE `2.880739`, changed16 `0.044888`, SSIM `0.920083`.

typed scene shape text font prepass / 2026-05-24:
`PptxSyntheticTextAndShapesUseSiblingOrder` and `PptxSyntheticGroupedTextAppliesTransform` passed after
replacing whole-part XML shape text font discovery with `ReadSceneShapeTextSpans`. The non-slow suite passed
with 190 passed, 0 failed, 7 skipped. `dotnet pack` succeeded and the full suite passed with 197 passed,
0 failed, 0 skipped. Private run `artifacts/private-visual/lokad-value-based/20260524-223511` stayed stable:
84/84 compared pages, zero dimension mismatches, deck MAE `9.005915`, changed16 `0.116052`, only
`PPTX_UNSUPPORTED_IMAGE_RECOLOR`, and slide 17 MAE `2.880739`, changed16 `0.044888`, SSIM `0.920083`.

obsolete PPTX text prepass wrapper removal / 2026-05-24:
After the shape text font prepass moved onto typed scene nodes, the former whole-part text helper wrappers
(`ReadInheritedTextRuns`, `ReadInheritedTextSpans`, `ReadSlideTextRuns`, `ReadSlideTextSpans`, and the
whole-part `ReadTextRuns` overload) had no remaining production or inspection callers. They were deleted so
the supported PPTX render path no longer retains a misleading whole-slide XML text prepass entry point. The
lower-level `ReadTextSpans` pipeline and `ReadSlideTextSpansForInspection` remain because inspection snapshots
still intentionally compare inherited and slide text at the XML level. `PptxSyntheticTextAndShapesUseSiblingOrder`
passed, the non-slow suite passed with 190 passed, 0 failed, 7 skipped, `dotnet pack` succeeded, and the full
suite passed with 197 passed, 0 failed, 0 skipped.

table prepass dummy-PDF removal / 2026-05-24:
`ReadSceneTableTextSpans` no longer creates a disposable `PdfGraphicsBuilder` just to collect table text for
font discovery. The table walk now has separate text-only and drawing entry points over the same geometry and
merge-handling loop, so the prepass cannot accidentally write discarded PDF operators. This is a useful
intermediate step, not the final table architecture: the next long-term split should promote table layout to a
small data model containing cell rectangles, effective fills, border segments, and text boxes, with PDF drawing
as a consumer of that model. `PptxSyntheticTableRendersGridAndText`, the `pptx-tables` non-slow group, and
`PptxSyntheticGroupedTableUsesGroupTransformWithUnknownGraphicFrame` passed. The non-slow suite passed with
190 passed, 0 failed, 7 skipped, `dotnet pack` succeeded, and the full suite passed with 197 passed,
0 failed, 0 skipped.

table layout model split / 2026-05-24:
The table renderer now builds an explicit `TableFrameLayout` containing positioned text spans, cell fill
rectangles, the default-grid plan, and explicit border segments before emitting PDF operators. The font prepass
reads only the layout text spans, while `RenderTableFrameLayout` consumes the same structural layout for fills
and borders. This removes the nullable/draw-or-collect processing mode introduced as an intermediate step and
pushes tables closer to the desired architecture: OOXML inheritance and geometry first, PDF serialization
second. The `pptx-tables` non-slow group, `PptxSyntheticGroupedTableUsesGroupTransformWithUnknownGraphicFrame`,
the non-slow suite, `dotnet pack`, and the full suite passed; the broad suite results stayed at 190 passed,
0 failed, 7 skipped for non-slow and 197 passed, 0 failed, 0 skipped for full.

scene-node text adapter boundary / 2026-05-24:
Production shape-text callers now go through `ReadTextSpansForSceneNode` instead of reaching directly from
dispatch/prepass code into `node.Source`. The adapter still delegates to the XML shape layout pipeline because
text geometry, body properties, group ancestry, and placeholder style lookup are not yet fully represented in
typed layout inputs. Keeping that dependency in one adapter clarifies the next migration step: build text frame
models from `PptxSceneNode` plus typed `PptxSceneTextBody`, then leave XML only as a compatibility source for
unmodeled OOXML edges. The unused `ReadTextRunsForShape` wrappers were removed. `PptxSyntheticTextAndShapesUseSiblingOrder`,
`PptxSyntheticGroupedTextAppliesTransform`, and the non-slow suite passed with 190 passed, 0 failed, 7 skipped;
`dotnet pack` succeeded, and the full suite passed with 197 passed, 0 failed, 0 skipped.

post table/text refactor private validation / 2026-05-24:
Private run `artifacts/private-visual/lokad-value-based/20260524-224755` stayed stable after the table layout
model split and text adapter cleanup: 84/84 compared pages, zero dimension mismatches, deck MAE `9.005915`,
changed16 `0.116052`, and only `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Slide 17 stayed at MAE `2.880739`,
changed16 `0.044888`, SSIM `0.920083`.

scene-owned picture render path / 2026-05-24:
Ordered picture rendering no longer passes `PptxSceneNode.Source` into the image renderer. The supported
picture path now consumes scene-owned relationship id, bounds, crop/fill rectangle, alpha, and recolor data;
the unused SVG renderer XML parameter was removed. This keeps raw `p:pic` XML attached to scene nodes for
future unsupported edges, but the production image path no longer depends on it. The `pptx-images` non-slow
group passed with 13 passed, 0 failed, 0 skipped, and the non-slow suite passed with 190 passed, 0 failed,
7 skipped. `dotnet pack` succeeded, and the full suite passed with 197 passed, 0 failed, 0 skipped.

obsolete XML-only shape render entry point / 2026-05-24:
The unused raw-XML `RenderShape(XElement, ...)` entry point was removed after ordered scene dispatch made
`PptxSceneNode` the required production path for shapes. The remaining scene-node adapter still passes
`shape.Source` into the lower shape renderer on purpose because style references, custom geometry fallback,
picture fills, and other unmodeled DrawingML edges still need the source XML until a richer typed
`ShapeRenderModel` exists. This is therefore a cleanup of a dead fallback path, not a claim that shape
rendering is fully model-owned. The `pptx-shapes` non-slow group passed with 14 passed, 0 failed, 0 skipped,
and the non-slow suite passed with 190 passed, 0 failed, 7 skipped. `dotnet pack` succeeded, and the full
suite passed with 197 passed, 0 failed, 0 skipped.

model-owned table frame layout / 2026-05-24:
`BuildTableFrameLayout` now consumes `PptxSceneTable` rows, cells, dimensions, style, fills, and borders
directly instead of accepting a parallel `a:tbl` XML fallback. The scene table still retains its source XML
for future unsupported DrawingML edges, but table geometry/fill/border/text layout in the renderer now has a
single model-owned input. This preserves the long-term direction: parse OOXML once into scene data, build a
layout model from that scene data, then serialize PDF operators from the layout. The `pptx-tables` non-slow
group passed with 7 passed, 0 failed, 0 skipped; the full suite passed with 197 passed, 0 failed, 0 skipped;
and `dotnet pack` succeeded.

obsolete chart relationship overload removal / 2026-05-24:
After ordered scene dispatch made `PptxSceneChart` the chart entry point, the old `RenderChartFrame` overload
that accepted only relationship id and target part name had no callers. It was removed while leaving the
current chart renderer's XML fallbacks intact: charts still need a staged migration because many plot, axis,
label, and style reads intentionally fall back to chart XML until the scene chart model fully covers those
OOXML surfaces. The `pptx-charts` non-slow group passed with 5 passed, 0 failed, 0 skipped; the non-slow
suite passed with 190 passed, 0 failed, 7 skipped; `dotnet pack` succeeded; and the full suite passed with
197 passed, 0 failed, 0 skipped.

scene-owned SmartArt graphic-frame classification / 2026-05-24:
`PptxSceneNode` now carries `IsSmartArtGraphicFrame`, and ordered dispatch uses that scene flag instead of
peeking back into `node.Source` when deciding whether to suppress unsupported-graphic-frame diagnostics for
SmartArt. The XML predicate moved to `PptxSceneBuilder` and remains shared with the separate diagnostic
preflight, which still scans slide XML by design. This removes one more raw scene-source dependency from the
dispatch layer while preserving existing SmartArt and unsupported graphic-frame behavior. The full suite
passed with 197 passed, 0 failed, 0 skipped, and `dotnet pack` succeeded.

post scene/model ownership private validation / 2026-05-24:
Private run `artifacts/private-visual/lokad-value-based/20260524-230051` stayed stable after the shape, table,
chart, and SmartArt scene-boundary cleanups: 84/84 compared pages, zero dimension mismatches, deck MAE
`9.005915`, changed16 `0.116052`, and only `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Slide 17 stayed at MAE
`2.880739`, changed16 `0.044888`, SSIM `0.920083`.

named table text-height line metric / 2026-05-24:
The table-cell text-height estimator no longer carries a local literal `1.2` line-height multiplier; it now
uses `PptxTextMetricRules.CssNormalLineHeightFallback`, the existing named metric owner. This is still an
approximate table-centering estimate and should eventually be replaced by the common text-frame layout model,
but the fallback is no longer an isolated table heuristic. The `pptx-tables` non-slow group passed with 7
passed, 0 failed, 0 skipped; the non-slow suite passed with 190 passed, 0 failed, 7 skipped; `dotnet pack`
succeeded; and the full suite passed with 197 passed, 0 failed, 0 skipped.

named chart fallback metrics / 2026-05-24:
The chart renderer no longer carries isolated literals for title, legend, category-axis, value-axis, and data
label fallback font sizes, doughnut hole-size clamping, or single-value axis headroom. These values now live
under `PptxChartMetricRules`. This is only an inventory and ownership step: the rules remain approximations
until replaced by Office-PDF-observed chart layout behavior or a richer scene chart model. The `pptx-charts`
non-slow group passed with 5 passed, 0 failed, 0 skipped; the non-slow suite passed with 190 passed, 0 failed,
7 skipped; `dotnet pack` succeeded; and the full suite passed with 197 passed, 0 failed, 0 skipped.

shared default chart plot box / 2026-05-24:
The repeated line/area/scatter default plot-box ratios now flow through `GetDefaultChartPlotBox` and
`PptxChartMetricRules` instead of being duplicated in each renderer. Bar-chart title/legend-specific plot
boxes remain separate because they are different fallback behavior and need Office-PDF evidence before being
merged. The `pptx-charts` non-slow group passed with 5 passed, 0 failed, 0 skipped; the non-slow suite passed
with 190 passed, 0 failed, 7 skipped; `dotnet pack` succeeded; and the full suite passed with 197 passed, 0
failed, 0 skipped.

chart plot-box ownership for area/scatter/bubble / 2026-05-24:
Area, scatter, and bubble chart renderers now consume a `ChartPlotBox` from the shared chart layout path
instead of recomputing frame-relative geometry internally. This extends manual plot-area layout handling
beyond bar/line charts and makes the remaining split explicit: radar and pie/doughnut still derive their own
polar geometry from the frame and need separate Office-PDF-backed layout treatment. The `pptx-charts`
non-slow group passed with 5 passed, 0 failed, 0 skipped; the non-slow suite passed with 190 passed, 0 failed,
7 skipped; `dotnet pack` succeeded; and the full suite passed with 197 passed, 0 failed, 0 skipped.

post chart-layout private validation / 2026-05-24:
Private run `artifacts/private-visual/lokad-value-based/20260524-231325` stayed stable after chart metric and
plot-box ownership cleanup: 84/84 compared pages, zero dimension mismatches, deck MAE `9.005915`, changed16
`0.116052`, and only `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained at MAE `2.880739`, RMSE `19.298084`,
changed16 `0.044888`, changed32 `0.035257`, SSIM `0.920083`, foreground histogram correlation `0.999858`,
with matching dimensions.

area chart manual plot-box public lock / 2026-05-24:
`PptxSyntheticAreaChartManualPlotBoxUsesSharedPath` now locks the shared chart plot-box path for an area chart
with `manualLayout` edge modes. This protects the area/scatter/bubble layout-owner change from silently
falling back to frame-relative default geometry. The `pptx-charts` non-slow group passed with 6 passed, 0
failed, 0 skipped; the non-slow suite passed with 191 passed, 0 failed, 7 skipped; `dotnet pack` succeeded;
and the full suite passed with 198 passed, 0 failed, 0 skipped.

radar chart manual plot-box public lock / 2026-05-24:
Radar chart rendering now consumes the shared `ChartPlotBox` path instead of computing its polar center and
radius directly from the chart frame. `PptxSyntheticRadarChartManualPlotBoxUsesSharedPath` locks the first
spoke against a `manualLayout` edge-mode plot box. Pie and doughnut still use frame-derived polar geometry
and remain the next explicit chart-layout gap. The `pptx-charts` non-slow group passed with 7 passed, 0
failed, 0 skipped; the non-slow suite passed with 192 passed, 0 failed, 7 skipped; `dotnet pack` succeeded;
and the full suite passed with 199 passed, 0 failed, 0 skipped.

pie/doughnut polar plot-box public lock / 2026-05-24:
Pie and doughnut rendering now use a polar `ChartPlotBox`: explicit plot-area `manualLayout` drives center,
radius, and data-label clipping, while the default remains the full chart frame to preserve existing
frame-derived behavior until Office-PDF evidence supports a stronger default polar plot-area rule.
`PptxSyntheticPieChartManualPlotBoxUsesPolarPath` locks a manual-layout pie slice against the old
frame-derived geometry. The `pptx-charts` non-slow group passed with 8 passed, 0 failed, 0 skipped; the
non-slow suite passed with 193 passed, 0 failed, 7 skipped; `dotnet pack` succeeded; and the full suite
passed with 200 passed, 0 failed, 0 skipped.

polar chart metric ownership / 2026-05-24:
Pie/doughnut and radar center/radius ratios now live under `PptxChartMetricRules`, and the chart renderer
uses shared polar geometry helpers instead of repeating local formulas in slice, cutout, label, and radar
paths. This is behavior-neutral naming and ownership, not evidence that the ratios are Office-perfect; the
open long-term task remains replacing these approximations with Office-PDF-observed chart layout behavior.
The `pptx-charts` non-slow group passed with 8 passed, 0 failed, 0 skipped; the full suite passed with 200
passed, 0 failed, 0 skipped; and `dotnet pack` succeeded.

empty endParaRPr paragraph public lock / 2026-05-24:
While comparing `pptx-renderer` typography coverage, `ooxpdf` already had the structural behavior for an
empty paragraph whose only layout content is `a:endParaRPr`, but only trailing-break coverage locked the
behavior. `PptxSyntheticEmptyParagraphUsesEndParagraphFontSize` now verifies that an otherwise empty
paragraph advances by its `endParaRPr` font size before the next visible paragraph. This preserves the
OOXML paragraph model distinction between "no visible runs" and "no layout content" without introducing a
slide-specific placement rule. The `pptx-typography` non-slow group passed with 67 passed, 0 failed, 2
skipped; the full suite passed with 201 passed, 0 failed, 0 skipped; and `dotnet pack` succeeded.

empty paragraph anchor-estimate alignment / 2026-05-24:
The vertical-anchor text-height estimator now mirrors line layout for layout-only paragraphs: first-paragraph
`spcBef` is skipped, later `spcBef` is conditional on prior layout content, and an empty paragraph with
`a:endParaRPr` uses the `endParaRPr` font size as the reference for percentage spacing. This closes a
structural mismatch where anchored text could estimate a shorter height than the actual layout path when an
empty paragraph had percentage spacing. `PptxSyntheticVerticalAnchorUsesEndParagraphFontSizeForEmptySpacing`
locks percent spacing against an equivalent point spacing case. The `pptx-typography` non-slow group passed
with 68 passed, 0 failed, 2 skipped; the full suite passed with 202 passed, 0 failed, 0 skipped; and
`dotnet pack` succeeded.

post empty-paragraph anchor private validation / 2026-05-24:
Private run `artifacts/private-visual/lokad-value-based/20260524-233516` stayed stable after the
empty-paragraph anchor-estimate fix: 84/84 compared pages, zero dimension mismatches, deck MAE `9.005915`,
changed16 `0.116052`, and only `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 stayed at MAE `2.880739`, RMSE
`19.298084`, changed16 `0.044888`, changed32 `0.035257`, SSIM `0.920083`, and foreground histogram
correlation `0.999858`.

cartesian chart data-label metric ownership / 2026-05-24:
Bar and line chart data-label placement still uses fallback ratios, but the slot fill, category inset,
horizontal label width, vertical label width, label-height, and line-label point-span values now live under
`PptxChartMetricRules` instead of anonymous literals inside the rendering loops. This is behavior-neutral
inventory work: the named rules remain approximations until replaced by Office-PDF-observed cartesian chart
layout behavior. The `pptx-charts` non-slow group passed with 8 passed, 0 failed, 0 skipped; the full suite
passed with 202 passed, 0 failed, 0 skipped; and `dotnet pack` succeeded.

chart axis-label metric ownership / 2026-05-24:
Category and value axis label placement still uses fallback sizing and offset rules, but the horizontal/vertical
axis label height, width, clip, side-gap, and baseline-offset factors now live under `PptxChartMetricRules`
instead of anonymous literals inside axis rendering. This does not make the defaults Office-perfect; it makes
the remaining approximations explicit and replaceable with PDF-observed chart layout rules. The `pptx-charts`
non-slow group passed with 8 passed, 0 failed, 0 skipped; the full suite passed with 202 passed, 0 failed,
0 skipped; and `dotnet pack` succeeded.

chart title/legend metric ownership / 2026-05-24:
Chart title and legend fallback placement still use approximate offsets, but their inset, baseline, width,
line-height, marker-size, side-gap, and clip-height factors now live under `PptxChartMetricRules`. This keeps
chart text placement approximations visible and replaceable, rather than scattered as local literals in the
rendering path. The `pptx-charts` non-slow group passed with 8 passed, 0 failed, 0 skipped; the full suite
passed with 202 passed, 0 failed, 0 skipped; and `dotnet pack` succeeded.

chart data-label offset metric ownership / 2026-05-24:
Pie data-label minimum width and height, bar data-label point gaps, and line data-label side/above/below
offsets now live under `PptxChartMetricRules` instead of being embedded in individual label placement
branches. This is another inventory step, not evidence that the offsets match Office exactly; it keeps the
remaining chart text-placement heuristics explicit for later replacement by PDF-observed chart layout behavior.
The `pptx-charts` non-slow group passed with 8 passed, 0 failed, 0 skipped; the full suite passed with 202
passed, 0 failed, 0 skipped; and `dotnet pack` succeeded.

chart manual-layout default gap / 2026-05-24:
Inspection of the current chart layout path found that `PptxSceneChartManualLayout` and the XML fallback require
all four `x`, `y`, `w`, and `h` children before applying a `manualLayout`. This preserves complete explicit
layouts, including edge modes, but it still drops partial manual layouts instead of resolving missing children
against Office's default chart layout position. Do not paper over this with per-chart constants: the next durable
step is to preserve optional manual-layout values in the scene model and resolve them against a default plot-area
box derived from Office/PDF evidence.

chart value-axis text metric ownership / 2026-05-24:
Value-axis label width estimation and vertical clip expansion still use fallback text metrics, but the whitespace,
digit, non-digit, and extra-clip factors now live under `PptxChartMetricRules`. This keeps the value-axis text
measurement approximation visible beside the rest of the chart layout rules. The `pptx-charts` non-slow group
passed with 8 passed, 0 failed, 0 skipped; the full suite passed with 202 passed, 0 failed, 0 skipped; and
`dotnet pack` succeeded.

chart manual-layout factor x/y semantics / 2026-05-24:
Manual plot-box construction now distinguishes `xMode`/`yMode` `factor` from `edge`: factor-positioned `x` and
`y` values are resolved from the default plot-box position, while edge-positioned values remain chart-frame
relative. This follows the Office behavior note in MS-OI29500 for `ST_LayoutMode` and removes a structural
layout-mode collapse without introducing chart-specific coordinates. `PptxSyntheticChartManualLayoutFactorPositionUsesDefaultPlotBox`
locks the PDF plot-area rectangle for this case. The `pptx-model` non-slow group passed with 12 passed, 0 failed,
1 skipped; the `pptx-charts` non-slow group passed with 8 passed, 0 failed, 0 skipped; the full suite passed with
203 passed, 0 failed, 0 skipped; and `dotnet pack` succeeded. Remaining gap: partial manual layouts still need
optional-value scene-model preservation before missing children can be resolved safely against the same default
plot-box anchor.

chart partial manual-layout preservation / 2026-05-24:
`PptxSceneChartManualLayout` now preserves optional `x`, `y`, `w`, and `h` children independently instead of
dropping the whole `manualLayout` unless all four are present. Manual plot-box resolution fills omitted values
from the default plot-box ratios, while still applying explicit `factor` and `edge` modes when present.
`PptxSyntheticChartManualLayoutMissingPositionUsesDefaultPlotBox` locks the case where `w`/`h` are explicit but
`x`/`y` are omitted: the plot area keeps the default position and only changes size. The `pptx-model` non-slow
group passed with 13 passed, 0 failed, 1 skipped; the `pptx-charts` non-slow group passed with 8 passed, 0 failed,
0 skipped; the full suite passed with 204 passed, 0 failed, 0 skipped; and `dotnet pack` succeeded.

post manual-layout private validation / 2026-05-24:
Private run `artifacts/private-visual/lokad-value-based/20260524-235547` stayed stable after chart manual-layout
mode and optional-value handling: 84/84 compared pages, zero dimension mismatches, deck MAE `9.005915`, changed16
`0.116052`, and only `PPTX_UNSUPPORTED_IMAGE_RECOLOR`. Page 17 remained at MAE `2.880739`, RMSE `19.298084`,
changed16 `0.044888`, changed32 `0.035257`, SSIM `0.920083`, foreground histogram correlation `0.999858`, with
matching dimensions.

bar chart default plot-box metric ownership / 2026-05-24:
Bar-chart default plot-box ratios for the no-title/no-legend case and the title/legend-reserved case now live
under `PptxChartMetricRules`, and the remaining line data-label height literal reuses the cartesian data-label
height rule. This is behavior-neutral ownership: the ratios remain fallback defaults until replaced by
Office/PDF-observed chart layout behavior. The `pptx-charts` non-slow group passed with 8 passed, 0 failed,
0 skipped; the `pptx-model` non-slow group passed with 13 passed, 0 failed, 1 skipped; the full suite passed with
204 passed, 0 failed, 0 skipped; and `dotnet pack` succeeded.

chart axis numeric-rule ownership / 2026-05-24:
Axis epsilon, nice-tick target count, and the `1/2/5/10` fallback tick ladder now live under
`PptxChartMetricRules` instead of being embedded in tick/gridline loops. This keeps numeric fallbacks auditable
for later Office/PDF replacement without changing the current formula. The `pptx-charts` non-slow group passed
with 8 passed, 0 failed, 0 skipped; the full suite passed with 204 passed, 0 failed, 0 skipped; and `dotnet pack`
succeeded.

chart title scene-structure preservation / 2026-05-25:
`PptxSceneChartTitle` now preserves title `overlay`, `manualLayout`, fill, and stroke data in addition to text
and text style. This is an intentional model-first slice: rendering did not start consuming title layout yet,
because exact Office title placement still needs public PDF evidence. The focused scene test passed with 1
passed, 0 failed, 0 skipped; the full suite passed with 204 passed, 0 failed, 0 skipped; and `dotnet pack`
succeeded.

chart axis-title scene-structure preservation / 2026-05-25:
`PptxSceneChartAxis` now carries a `PptxSceneChartTitle`, so category and value axes preserve explicit
OOXML axis-title text, overlay, manual layout, fill/stroke, and text-style data before rendering uses it.
The focused scene test passed with 1 passed, 0 failed, 0 skipped; the full suite passed with 204 passed, 0
failed, 0 skipped; and `dotnet pack` succeeded.

chart legend scene-structure preservation / 2026-05-25:
`PptxSceneChartLegend` now preserves legend `manualLayout` and `spPr` fill/stroke data in addition to position,
overlay, visibility, and text style. This is a model-first preservation step: rendering still uses the existing
legend placement path until public Office-PDF evidence defines the layout rule. The focused scene test passed
with 1 passed, 0 failed, 0 skipped; the full suite passed with 204 passed, 0 failed, 0 skipped; and
`dotnet pack` succeeded.

chart text semantic-role oracle / 2026-05-25:
`ClassifyPdfChartText.ps1` now classifies obvious chart text into semantic roles (`CategoryAxisTickLabel`,
`ValueAxisTickLabel`, `DataLabelText`, `LegendText`, and title candidates) instead of only broad relative
position buckets. `CheckVisualCase.ps1` includes these roles in the default chart-text gate kind list while
retaining the earlier fallback bucket names. The public clustered-column case now gates four
`CategoryAxisTickLabel` structures, and `pwsh tools/CheckVisualCase.ps1 -Case
visual-cases/cases/pptx-ladder-11-chart-column-clustered-port/case.json` passed at
`artifacts/visual/pptx-ladder-11-chart-column-clustered-port/20260525-005334`.

chart data-label override layout preservation / 2026-05-25:
`PptxSceneChartDataLabelOverride` now preserves `c:layout/c:manualLayout` factors for explicit per-label
placement. This is a model-first preservation step; rendering does not consume the label layout until public
Office-PDF evidence defines the label-box placement rule. The focused scene test passed with 1 passed,
0 failed, 0 skipped; the full suite passed with 204 passed, 0 failed, 0 skipped; and `dotnet pack` succeeded.

chart series smooth-flag scene ownership / 2026-05-25:
`PptxSceneChartSeries.Smooth` is now nullable so the scene model distinguishes a missing `c:smooth` element
from explicit Office-authored `val="0"` and `val="1"` values. `ReadSceneOrXmlSmoothSeries` now trusts the
typed scene vector when any series in the plot carries an explicit smooth value, instead of falling back to
a broad chart-XML rescan whenever every scene value was `false`. The scene fixture includes a valid line-chart
series with explicit smooth disable. The focused scene test passed with 1 passed, 0 failed, 0 skipped; the
full suite passed with 204 passed, 0 failed, 0 skipped; and `dotnet pack` succeeded.

chart data-label layout adapter preservation / 2026-05-25:
Renderer-side `ChartDataLabelOptions` and `ChartDataLabelOverride` now carry per-label manual-layout metadata
forward from both typed scene data and raw XML fallback parsing. Effective per-point label option resolution
preserves an override layout when one is present, but rendering still does not consume it until public
Office-PDF evidence defines the label-box placement rule. The focused scene test passed with 1 passed,
0 failed, 0 skipped; the full suite passed with 204 passed, 0 failed, 0 skipped; and `dotnet pack` succeeded.

chart stroke-marker structural oracle / 2026-05-25:
`ClassifyPdfChartGraphics.ps1` now emits `StrokeMarkerCandidate` for compact stroked chart paths whose bounds
look like marker outlines rather than axis/gridline segments. This is an opt-in structural oracle kind for
future marker geometry gates; default chart graphics comparisons remain unchanged. The clustered-column public
visual case still passed with the existing `AxisPairPlotBoxCandidate` and `CategoryAxisTickLabel` structural
gates at `artifacts/visual/pptx-ladder-11-chart-column-clustered-port/20260525-010512`.

chart marker presence scene ownership / 2026-05-25:
`PptxSceneChartMarker` now preserves whether a `c:marker` element was actually present, while still carrying
the effective symbol/size defaults used by the current renderer. `ReadSceneOrXmlMarkerStyles` only treats the
typed scene marker vector as authoritative when at least one series in the plot has explicit marker metadata,
so missing-marker plots keep the existing XML fallback path instead of turning scene defaults into false
source ownership. The focused scene test passed with 1 passed, 0 failed, 0 skipped.
The full suite passed with 204 passed, 0 failed, 0 skipped; and `dotnet pack` succeeded.

chart varyColors presence scene ownership / 2026-05-25:
`PptxSceneChartPlot.VaryColors` is now nullable so the scene model distinguishes a missing `c:varyColors`
element from explicit Office-authored true/false values. Rendering keeps the same effective behavior by using
typed values when present and the existing XML/default fallback otherwise, but the scene no longer claims
ownership of defaulted palette-varying semantics. The focused scene test passed with 1 passed, 0 failed,
0 skipped.
The full suite passed with 204 passed, 0 failed, 0 skipped; and `dotnet pack` succeeded.

chart data-label flag presence ownership / 2026-05-25:
`PptxSceneChartDataLabels` now preserves top-level plot/series `c:dLbls` boolean flags as nullable values,
matching the existing nullable override-flag handling. This distinguishes missing inherited/defaulted label
flags from explicit Office-authored true/false values in the scene model while the renderer adapter still
coalesces to the current effective booleans at the last responsible point. The focused scene test passed with
1 passed, 0 failed, 0 skipped.
The full suite passed with 204 passed, 0 failed, 0 skipped; and `dotnet pack` succeeded.

chart title/legend overlay presence ownership / 2026-05-25:
Chart title and legend overlay flags are now nullable in the scene model, so a missing `c:overlay` element is
not confused with explicit `val="0"`. The renderer still coalesces missing overlay metadata to the existing
effective false behavior when computing legend layout, while scene inspection can now distinguish source
presence. The focused scene test passed with 1 passed, 0 failed, 0 skipped.
The full suite passed with 204 passed, 0 failed, 0 skipped; and `dotnet pack` succeeded.

chart automatic value-axis scale unification / 2026-05-25:
Public clustered-column Office evidence showed a chart with data max `61`, no explicit `c:max`, and no
explicit `c:majorUnit` rendering value-axis labels `0, 10, 20, 30, 40, 50, 60, 70`. The candidate previously
chose axis max `80` because `GetNiceChartAxisMax` used a separate five-tick heuristic while labels/gridlines
used `ChooseChartAxisMajorUnit` with the shared ten-tick ladder. The fallback axis maximum now derives from
the same automatic major unit used for ticks, so scale endpoint and tick sequence are one inferred numeric
rule instead of two competing approximations. The clustered-column visual manifest now gates both
`CategoryAxisTickLabel` and `ValueAxisTickLabel` structures; it passes at
`artifacts/visual/pptx-ladder-11-chart-column-clustered-port/20260525-012351`. Remaining gap: value-axis
label X positions are still about `48 pt` right of Office and category labels about `26 pt` low, so the
public gate keeps a broad `50 pt` chart-text tolerance until the Office-PDF plot-box and axis-label placement
rules replace the current fallback geometry.

bar chart no-title plot-box Office rule / 2026-05-25:
The same public clustered-column fixture has a 720 pt by 432 pt chart frame at `(72, 72)` and no visible chart
title or legend. Office's observable PDF axis pair is `(113.45, 111.90)` to `(781.00, 488.01)`, which gives
frame-relative ratios about `0.0576`, `0.0924`, `0.9272`, and `0.8706`. The no-title/no-legend bar chart
fallback plot-box ratios now use those Office-PDF-observed values instead of the prior approximate overlay-only
ratios. The public clustered-column gate tightened `AxisPairPlotBoxCandidate` bounds from `100 pt` to `1 pt`
and chart tick-label positions from `50 pt` to `10 pt`, passing at
`artifacts/visual/pptx-ladder-11-chart-column-clustered-port/20260525-012613`. Remaining gap: category-axis
tick labels are still about `1.7 pt` below Office, and value-axis labels remain about `8.6 pt` right and
`5.7 pt` below Office; those are now isolated text-box/baseline/width placement gaps rather than plot-box
geometry drift.

chart value-axis side-gap Office rule / 2026-05-25:
The public clustered-column gate isolated value-axis tick-label X drift after the plot-box fix: the whole
vertical value-axis label column was about `8.6 pt` right of Office while the axis-pair plot box already
matched within `0.02 pt`. The renderer now separates label-width padding from the side gap between a vertical
value axis and its tick-label box, using the Office-PDF-observed side-gap rule instead of reusing the internal
text-width padding factor. The same public case passes with value-axis tick-label X deltas within `0.04 pt`
at `artifacts/visual/pptx-ladder-11-chart-column-clustered-port/20260525-013246`. Remaining gap: value-axis
tick labels are still about `5.7 pt` below Office and category-axis labels about `1.7 pt` below Office, so the
next Office-alignment slice should target chart text-box vertical placement and baseline semantics rather than
axis scale or plot-box geometry.

chart axis tick-label vertical placement Office rule / 2026-05-25:
The same clustered-column Office-PDF structural gate showed that, after X placement and plot-box geometry were
aligned, the remaining tick-label drift was purely vertical: category-axis labels were about `1.7 pt` low and
vertical value-axis labels about `5.7 pt` low. `PptxChartMetricRules` now uses Office-PDF-observed vertical
label placement factors for category-axis top offset and vertical value-axis baseline ratio. The public gate is
tightened from `10 pt` to `1 pt` for chart text structures and passes at
`artifacts/visual/pptx-ladder-11-chart-column-clustered-port/20260525-013507`, with category-axis Y deltas
around `0.01 pt` and value-axis Y deltas within `0.06 pt`. Remaining chart text work: these are still metric
rules for simple unrotated tick labels; rich text runs, rotation, full non-default tick-label offset ladders,
multi-level categories, and chart-style inherited defaults remain open under the chart text-style ownership
track.

chart category-axis label-offset consumption / 2026-05-25:
The renderer now consumes scene/XML `c:catAx/c:lblOffset` for supported category tick labels as a
default-preserving scale on the label distance from the plot area. A new synthetic render test compares the
single text baseline from `lblOffset=100` and `lblOffset=200`, with value-axis labels hidden, so the stored
OOXML metadata now has observable PDF geometry instead of only scene-model preservation. This is still a
structural first pass: public Office-PDF evidence for non-default offsets, rotated labels, and multi-level
categories remains needed before this can be considered a complete Office rule. The focused render test
passed with 1 passed, 0 failed, 0 skipped; the `pptx-charts` non-slow group passed with 9 passed, 0 failed,
0 skipped; the clustered-column public visual gate passed at
`artifacts/visual/pptx-ladder-11-chart-column-clustered-port/20260525-014112`; the full suite passed with
205 passed, 0 failed, 0 skipped; and `dotnet pack` succeeded.

chart category-axis tick-label skip consumption / 2026-05-25:
The renderer now consumes scene/XML `c:catAx/c:tickLblSkip` for supported category tick labels as the
Office-authored drawing interval, defaulting to `1` when absent or invalid. A synthetic render test hides the
value axis and verifies that a four-category chart with `tickLblSkip=2` emits two category-label text
matrices, so the existing scene-owned metadata now affects PDF text structure rather than staying archival.
This does not consume `tickMarkSkip` yet; tick-mark geometry remains separate axis-line work because the
current renderer has not modeled category tick marks as independent strokes. The focused render test passed
with 1 passed, 0 failed, 0 skipped; the `pptx-charts` non-slow group passed with 10 passed, 0 failed,
0 skipped; the clustered-column public visual gate passed at
`artifacts/visual/pptx-ladder-11-chart-column-clustered-port/20260525-014345`; the full suite passed with
206 passed, 0 failed, 0 skipped; and `dotnet pack` succeeded.

chart gridline axis-scope fallback / 2026-05-25:
Gridline visibility fallback now checks the specific value-axis XML element being rendered instead of scanning
the whole chart for any `majorGridlines` or `minorGridlines`. This prevents category-axis gridline metadata
from enabling value-axis gridline strokes when the scene model is unavailable or incomplete. A synthetic
render test puts `c:majorGridlines` only on `c:catAx` and verifies the default value-gridline stroke is not
emitted. This is a structural cleanup toward axis-owned rendering; exact category-axis gridline rendering, if
needed, remains separate evidence work. The focused render test passed with 1 passed, 0 failed, 0 skipped.
The `pptx-charts` non-slow group passed with 11 passed, 0 failed, 0 skipped after keeping the existing
line-chart first-value-axis fallback scoped to that resolved axis; the clustered-column public visual gate
passed at `artifacts/visual/pptx-ladder-11-chart-column-clustered-port/20260525-014718`; the full suite
passed with 207 passed, 0 failed, 0 skipped; and `dotnet pack` succeeded.

chart gridline PDF-structure alignment / 2026-05-25:
Office evidence from the clustered-column public case showed major value gridlines emitted as one
multi-segment stroked PDF path, with the maximum tick included and the baseline axis excluded; the previous
candidate emitted separate gray line strokes and skipped the maximum tick. The renderer now derives gridline
ticks from the endpoint-inclusive axis sequence, filters the crossing-axis tick, emits all gridline segments
into one current path before a single `S`, and uses the Office-observed default major gridline stroke
(`0 0 0 RG`, `0.75 w`) when no explicit gridline style is present. `ClassifyPdfChartGraphics.ps1` now emits
`HorizontalGridlineGroupCandidate` and `VerticalGridlineGroupCandidate` for multi-segment gridline strokes so
public visual cases can compare Office/candidate structure directly instead of mistaking grouped Office paths
for missing gridlines. The clustered-column manifest now gates both `AxisPairPlotBoxCandidate` and
`HorizontalGridlineGroupCandidate` with line-width delta `0.05`; it passed at
`artifacts/visual/pptx-ladder-11-chart-column-clustered-port/20260525-020019`. Focused chart tests passed
with 12 passed, 0 failed, 0 skipped; the full suite passed with 208 passed, 0 failed, 0 skipped; and
`dotnet pack` succeeded.

chart gridline crossing consumption / 2026-05-25:
The gridline renderer no longer assumes that the excluded endpoint is always the value-axis minimum. For
supported bar and line charts, scene-owned crossing metadata (`crossesAt` and `crosses`) or the same raw XML
fallback now provides the value tick that represents the crossing axis. `crosses=max` excludes the maximum
tick, `crosses=min` excludes the minimum tick, `crossesAt` excludes the explicit crossing value, and
`autoZero` excludes zero when visible or the nearest endpoint when the data range is entirely positive or
negative. This is a structural axis-owned rule, not a visual nudge: the candidate keeps Office-like grouped
PDF gridline paths while changing which segment is present. A synthetic `c:crosses val="max"` chart proves
the minimum tick remains visible as a gridline while the maximum crossing tick is omitted. The focused
`pptx-charts` non-slow group passed with 13 passed, 0 failed, 0 skipped; the clustered-column public visual
gate passed at `artifacts/visual/pptx-ladder-11-chart-column-clustered-port/20260525-020737`; the full suite
passed with 209 passed, 0 failed, 0 skipped; and `dotnet pack` succeeded. Remaining axis-crossing work is
larger than gridlines: axis line placement, category/value baseline choice for bars and lines, reversed
scales, and secondary-axis binding still need Office-PDF evidence and public gates.

chart category-axis crossing-line consumption / 2026-05-25:
The vertical bar/line chart category-axis stroke now uses the same value-axis crossing resolver as gridlines.
For the supported vertical chart path, `c:crosses val="max"` moves the category-axis line to the value-axis
maximum, `c:crosses val="min"` keeps it at the minimum, `c:crossesAt` places it at the explicit value after
clamping to the visible value range, and `autoZero` uses zero when visible or the nearest endpoint otherwise.
This keeps the visible axis stroke structurally tied to OOXML crossing metadata instead of using a zero-line
heuristic. A synthetic structural PDF test verifies that a `crosses=max` category-axis horizontal stroke
shares the top coordinate of the value-axis vertical stroke. The focused `pptx-charts` non-slow group passed
with 14 passed, 0 failed, 0 skipped; the clustered-column public visual gate passed at
`artifacts/visual/pptx-ladder-11-chart-column-clustered-port/20260525-021112`; the full suite passed with
210 passed, 0 failed, 0 skipped; and `dotnet pack` succeeded. Remaining work deliberately stays open for
separate evidence: horizontal bar axis crossing, secondary-axis binding, reversed axis orientation, and
whether series baselines should move with axis crossing or stay tied to zero for each chart family.

chart reversed value-axis orientation consumption / 2026-05-25:
The renderer now consumes scene/XML `c:scaling/c:orientation val="maxMin"` for the supported vertical
bar/line value-position path. A shared value-to-plot-coordinate helper applies the reversed ratio to
gridline coordinates, value-axis tick label baselines, vertical category-axis crossing-line placement,
line-series points, and clustered vertical bar rectangle endpoints. This keeps the renderer aligned with the
typed `PptxSceneChartAxis.IsReversed` metadata instead of silently ignoring it after parsing. A synthetic
line chart with values `0` then `30` proves that `maxMin` places the lower value above the higher value in
the emitted diagonal PDF stroke. The focused `pptx-charts` non-slow group passed with 15 passed, 0 failed,
0 skipped; the clustered-column public visual gate passed at
`artifacts/visual/pptx-ladder-11-chart-column-clustered-port/20260525-021718`; the full suite passed with
211 passed, 0 failed, 0 skipped; and `dotnet pack` succeeded. This is not the end of axis orientation work:
horizontal bar axes are intentionally left unchanged in this slice, stacked bar/column accumulation still
needs a direction-aware model, chart data labels still use the older position formulas, secondary axes need
independent orientation, and a public Office-PDF visual case for reversed axes should be added before calling
this Office-complete.
```

Representative public visual cases already exist for PPTX blank/shapes/text/images/tables/corporate-theme and
DOCX blank/basic paragraphs/numbering/images/tables/headers-footers.

Private validation commands:

```powershell
pwsh tools/CheckPrivateCase.ps1 -Case private-cases/lokad-value-based.json
pwsh tools/CheckPrivateCase.ps1 -Case private-cases/user-requirements-spec.json
```

Do not commit private inputs, private manifests, private rendered pages, private diagnostics, private
comparison HTML, or private assessments.

## Idempotence And Recovery

All build/test/pack commands are safe to rerun. Visual validation writes timestamped directories. If Office
COM automation leaves an Office process running after a failure, close it only after confirming no unrelated
user document is open.

If a private case reveals a missing feature, record only public-safe feature gaps here, then create synthetic
public tests for the implementation. Do not derive public fixtures from private documents.

Revision note, 2026-05-24: An earlier attempt to compact this file too aggressively was reverted because it
discarded valuable open planning context and private-safe historical evidence. This revision is deliberately
conservative: it restores `PLANS.md` section naming, records the current slide-17 connector evidence, updates
the validation baseline, and documents that older private run artifacts are mostly absent locally. Future
trimming should be done in small audited slices and should preserve open progress items by default.
