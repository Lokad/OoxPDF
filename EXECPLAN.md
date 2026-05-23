# Lokad.OoxPdf Current Execution Plan

This ExecPlan is the current working plan for `Lokad.OoxPdf`. It intentionally omits the historical bootstrap
checklist that has already been completed. Keep `Progress`, `Private Evidence`, `Backlog`, `Decisions`, and
`Validation` current as work proceeds.

## Goal

`Lokad.OoxPdf` is a dependency-free .NET library that converts `.pptx` and `.docx` OOXML documents to static
PDF. The library must not call Office, PDFium, PowerShell, external executables, or third-party packages.
Office and PDFium are allowed only in `tools/` for validation.

The project is now past the initial vertical slice. The next phase is fidelity: use Office-exported PDFs from
public synthetic OOXML as the implementation oracle, inspect those PDFs and their raster output before
changing renderer behavior, use private local-only documents only to discover missing Office features, and
keep diagnostics honest when a feature is still missing.

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
- [x] Add a PDF-inspection typography harness that compares Office and candidate text matrices, TJ arrays,
  baseline positions, highlight rectangles, and clipping boxes before relying on raster metrics.
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
- [ ] Next PPTX typography sequence: attack accented Latin and punctuation-adjacent word cases after the
  spacing-control cases, using Office PDF text operations to separate fallback-font splits from positioning.
  Punctuation-adjacent hyphen boundaries are now represented as separate layout spans and locked by
  `pptx-ladder-04-typography-punctuation-boundaries` with line-start parity within `0.25pt`. Accented
  Latin remains open because Office still splits accent-heavy text into additional PDF operations.
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
- [ ] Tighten `pptx-ladder-04-typography-accent-spacing-probe`: Office splits accented Latin into several
  font/text operations for missing glyph fallback and combining marks, while the candidate currently emits
  fewer operations and has large x/y mismatches on Calibri and Arial accent rows.
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
- [ ] Split PPTX rendering dispatch by typed scene node: background, shape, text, picture, table, chart,
  group, and unknown/diagnostic fallback should be separate renderers consuming the same context.
- [ ] Move master/layout rendering into the scene/model pipeline: non-placeholder template nodes render in
  Office order, placeholders provide inherited geometry/body/style only, and show/hide flags are explicit.
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
- [ ] Port `pptx-renderer` shape geometry coverage: preset geometries, custom geometry, rotations, flips,
  group transforms, connectors, arrows, dash/cap/join, and picture-fill clipping.
- [ ] Port `pptx-renderer` image behavior: relationship resolution, crop/fill/stretch, alpha/soft masks,
  SVG or unsupported-image diagnostics, media caching, and reuse across slides.
- [ ] Port `pptx-renderer` chart behavior as a first-class native renderer: parse a typed chart model for
  series, axes, legends, labels, styles, and layouts; emit diagnostics only for unsupported chart features.
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
  is overlay semantics, manual layout, text styling, entry order/filtering, and exact Office spacing.
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
  Remaining title work is exact Office title layout, rich text styling, overlay/manual layout, and inherited
  chart style defaults.
- Bar and line chart renderer now honor `c:plotArea/c:spPr` fill and border styling through the shared
  chart shape-style helper. Remaining plot-area work is manual layout, rounded corners/effects, and extending
  exact plot bounds to area/scatter/radar/pie/doughnut families.
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
  stroke styles as the plotted series. Remaining legend work is Office layout positions, overlay behavior,
  rich text, hidden/deleted entries, and chart style inheritance.
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
- Vertical text `0033..0034`: ported as `pptx-ladder-04-vertical-text-port` with an explicit
  `PPTX_UNSUPPORTED_TEXT_ORIENTATION` diagnostic gate. Rendering remains a planned capability before
  private vertical text is revisited.
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
- [x] Split fast unit tests from slower visual/oracle tests so routine `dotnet run --project
  tests/Lokad.OoxPdf.Tests` stays fast and visual gates can be run explicitly.
  The custom runner now supports `--skip-slow`, `--only-slow`, and `--list`.
- [ ] Review `pptx-renderer`'s generated oracle and report reuse strategy for ideas that fit the .NET
  dependency-free constraint.

## Progress

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
- [x] Typography experiment: Office PDFs can embed and use math-table TTC faces directly for normal slide
  text when OOXML requests that exact typeface. Remove the family-name workaround and keep exact font
  resolution exact; remaining drift belongs in OpenType shaping/PDF text emission, not substitution.
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
  `0.05pt`; the accented probe remains a targeted gap for font fallback and combining-mark handling.
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
  - [x] Add Office-like automatic chart titles from the first series name when `autoTitleDeleted` is false
    and no explicit title node is present.
  - [x] Port the `pptx-renderer` nice-axis maximum rule for axes without explicit `c:max`, so max data values
    such as `45` expand to an Office-like `50` axis cap.
  - [x] Use document theme accent colors for default/vary-colors bar chart fills instead of a hard-coded
    Excel palette; the public horizontal-bar rung now matches Office's category colors.
  - [x] Increase the default native bar-chart plot height so Office-like value axes use the full plot region
    instead of compressing labels downward. This directly improves the private slide-5 right axis and keeps
    the public bar/column chart gates passing.
  - [x] Restrict implicit chart-title fallback to bar charts. Applying the first-series-name fallback to every
    chart family created false titles in area/bubble cases; the public `pptx-charts` family now passes the
    first eight imported chart gates again.
  - [x] Bind secondary right-axis label rendering to the extra bar-chart group that owns the right value axis,
    so combo charts do not use the primary series as the fallback range for secondary-axis ticks.
  - [ ] Extend combo/multi-axis chart support beyond the first bottom-up slice: bind each chart group to its
    referenced axes, honor axis tick-label formatting, keep primary/secondary scales independent, and place
    non-axis overlays such as the private slide 5 upward green arrow with Office-equivalent transforms.
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
  - [x] Add first-class `PptxTextOrientation` coverage for known `a:bodyPr @vert` variants (`vert`,
    `vert270`, `eaVert`, `mongolianVert`, `wordArtVert`, `wordArtVertRtl`) through a public synthetic unit
    fixture and route them through orientation-aware text-frame flow instead of silently treating them as
    horizontal text.
  - [x] Fix the lower input boxes being too high by using visible run font sizes for line advance and
    middle-anchor height estimation instead of seeding every visible line with the unrelated 18 pt fallback.
  - [ ] Continue with public fixtures for vertical label text and curved connector arrowhead/control-point
    fidelity before marking the slide-9 schema as resolved.
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
    nudged to MAE `0.791131`, changed16 `0.005626`; remaining differences are still dominated by Office's
    exact stacked-glyph positioning.
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

1. Prioritize the `pptx-renderer` track above all other fidelity work. Complete the architectural survey,
   identify the abstractions that explain its higher quality, and port those lessons into `ooxpdf` before
   making more broad renderer patches.
2. Clean-port `pptx-renderer` tests into `ooxpdf`, starting with its typography Office-oracle family, then
   shape/preset, layout/composition, images, tables, charts, and SmartArt. Do not vendor code or private
   assets; recreate minimal public fixtures and keep Office PDF references as validation artifacts.
3. Decide the intermediate-model direction from the `pptx-renderer` survey: what belongs in the PPTX
   scene/model, what remains direct PDF rendering, and how render context, style/color resolvers, diagnostics,
   and assets should be owned.
4. Continue PPTX Ladder 4 only through the `pptx-renderer`-driven path: inspect Office references, add or
   port minimal public typography cases, tighten baselines, advances, tracking, highlights, underlines,
   bullets, and paragraph flow toward near pixel-perfect output.
5. Keep test-loop performance work in service of the `pptx-renderer` test-port track: cached oracles,
   fast/slow/oracle lanes, and richer visual metrics should make hundreds of public cases practical.
   Hot-loop visual checks now avoid rebuilding the CLI, PDFium rasterizer, VisualDiff, and PdfInspect when
   source files outside `bin`/`obj` are older than the output DLLs. A cached typography probe now reruns in
   about 4.5s on this machine.
6. Defer DOCX and private-deck optimization while the PPTX public ladder and `pptx-renderer`-derived tests
   are still incomplete. Private files remain gap discovery and acceptance evidence only.

## Decisions

- The library remains dependency-free. Third-party packages are not allowed in `src/Lokad.OoxPdf`.
- Office and PDFium remain validation-only under `tools/`.
- Private documents remain under ignored `private-cases/`; generated private artifacts remain under ignored
  `artifacts/private-visual/`.
- Public notes from private documents must be anonymized to feature gaps and metrics only.
- Diagnostics must prefer continued conversion over crashing, but omitted visible content must not be treated
  as acceptable final behavior.
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

## Validation

Latest public validation:

```powershell
dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal
dotnet run --project tests/Lokad.OoxPdf.Tests --tl:off
dotnet pack src/Lokad.OoxPdf/Lokad.OoxPdf.csproj --tl:off --nologo -v minimal --no-restore
```

Current expected test result:

```text
104 passed, 0 failed
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
