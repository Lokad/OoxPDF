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
- `tools/SummarizeChartStructureDeltas.ps1`: public chart structural-delta sweep over latest visual runs.
  It accepts repeated PowerShell array values and comma/semicolon-separated `-Case` values for focused
  multi-case comparisons, and `-ShowBounds` prints reference/candidate bounds for singleton structural
  buckets when edge-level plot-box evidence is needed. `-ByRegion` splits the summary by chart-local
  `RegionIndex` when the graphics/text classifiers can attach structures to plot-box-derived regions.
- `tools/SummarizeChartDataLabelLayout.ps1`: focused public PDF structure summary for chart data-label text
  and leader-line layout gaps.
- `tools/SummarizePptxTextStateDeltas.ps1`: private-safe aggregate summary of Office/candidate PPTX text
  emission comparisons, including `Tc` buckets, table/style/frame structure, glyph residuals, and letter-case
  counts. It must not emit private text content.
- `tools/SummarizeDocxTextState.ps1`: private-safe aggregate summary of Office/candidate DOCX text operations
  from visual run directories, including operation counts, `Tc` buckets, `/Tf` sizes, positioned-glyph
  residual buckets, candidate planner segment/advance-profile/glyph-advance-signature buckets when `DocxInspect`
  output is present, source paragraph-index and planner-role buckets, and sequence-paired Office-operation/planner-segment
  buckets when the inspected operation counts match. It must not emit decoded document text.
- `tools/SummarizeDocxRowBoundary.ps1`: private-safe DOCX table-row boundary summary for visual run directories.
  It combines layout-snapshot row bands and PDF text rows near the page bottom, emitting row indices, geometry,
  baseline diagnostics, lengths, and hashes without decoded document text.
- `tools/CompareDocxLayoutPdfFlow.ps1`: private-safe DOCX layout/PDF row matcher that carries source indexes,
  line-height profile fields, paragraph spacing profile fields, and row-advance buckets for Office/candidate
  flow comparisons.
- `tools/SummarizePdfTextPageDeltas.ps1`: private-safe per-page Office/candidate PDF text-operation summary
  for visual run directories. It reports operation counts, `Tc` buckets, text-class buckets, and coordinate
  span deltas without decoded document text.
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

1. Extend chart structural oracle coverage. The first public tooling slices already exist:
   `tools/ClassifyPdfChartGraphics.ps1`, `tools/ClassifyPdfChartText.ps1`,
   `tools/ComparePdfGraphicsOperations.ps1`, and manifest-driven gates in `tools/CheckVisualCase.ps1`
   compare derived plot boxes, axis/gridline lines, legend swatches, chart text buckets, path geometry, line
   attributes, and clipping/fill/stroke structure. The remaining work is not to invent a second oracle, but to
   expand coverage across the ungated public chart cases, add first-class data-label, marker, legend-title,
   plot clipping, and resource-structure classifications where the current buckets are still too generic, and
   use those public gates to retire chart fallback geometry. This work must stay based on public synthetic chart
   fixtures and must not depend on private deck content.
2. Complete the chart scene model. Chart XML reads should move into typed scene/layout records for chart
   kinds, plot areas, axes, series, data labels, markers, title, legend, fills, strokes, and text styles.
   Raw XML may remain attached as source evidence while an OOXML surface is incomplete, but production
   rendering should increasingly consume typed chart data.
3. Replace chart fallback geometry. Named constants under `PptxChartMetricRules` are useful inventory, not
   the destination. Each chart ratio, offset, and text metric should either be replaced by an Office-PDF
   observed rule or explicitly classified as a temporary approximation with a public case that demonstrates
   the remaining gap.
4. Continue the private-deck typography path through public evidence. Slide 17 has mostly moved from graphics
   parity to small text-emission residuals, while high-error pages such as page 50 now point to broader
   text-fit/wrap and grouped cropped-picture structure. Use private slide evidence only to identify generic
   missing behavior, then lock that behavior with public synthetic Office-PDF-backed fixtures.
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
- Its feature survey confirms that the useful OOXPDF imports are structural, not mechanical:
  `src/model/Presentation.ts` resolves placeholder geometry and inherited `bodyPr` from layout/master nodes;
  `src/renderer/TextRenderer.ts` keeps a visible multi-layer paragraph/run cascade including default text
  styles, placeholder list styles, paragraph properties, bullet properties, line spacing, tabs, capitalization,
  baseline, no-fill, outline, highlight, and hyperlink handling; `src/renderer/StyleResolver.ts` applies
  layout/master color-map remapping, `phClr`, system/preset/HSL/SCRGB colors, and modifier chains; group
  rendering propagates `a:grpFill` through render context; image rendering inventories video/audio, crop,
  alpha/luminance/duotone/bi-level effects, SVG/BMP/EMF/WMF distinctions, and media URL lifetime; table parsing
  keeps grid spans, merge flags, table style IDs, cell properties, and cell text bodies; SmartArt support uses
  `ppt/diagrams/drawing*.xml` fallback drawings as grouped shape trees; preset shape tests and baseline
  evaluation emphasize spec-derived geometry before parameter tuning.
- Feature gaps relative to OOXPDF's long-term PDF target are equally important: the TypeScript chart renderer
  builds browser/ECharts/SVG structures rather than Office-like PDF plot boxes and text matrices; many feature
  parsers keep `SafeXmlNode` raw fields as renderer inputs; image handling uses browser blob URLs and optional
  PDF/EMF helpers that cannot become runtime dependencies; SmartArt fallback rendering contains
  layout-specific diagram heuristics that are useful as an inventory but must be replaced by public
  Office-PDF structural evidence before entering OOXPDF.
- Its slide renderer treats master/layout non-placeholder shapes as template content rendered behind slide
  nodes, while placeholders are inheritance templates rather than directly rendered content.
- Its model parsers keep raw XML attached to typed nodes. This is a useful compromise for `ooxpdf`: parse
  stable fields into records while keeping source XML available until every OOXML edge case has a typed home.
- Its chart support is not a direct architectural target for `ooxpdf`. `src/model/nodes/ChartNode.ts` only
  resolves a `chartPath`, `PresentationData.charts` stores chart XML as `SafeXmlNode`, and
  `src/renderer/ChartRenderer.ts` turns that XML into ECharts options. That is useful as an OOXML feature
  inventory, but it does not provide Office-like PDF structure for plot boxes, data-label boxes, gridlines,
  legend reserves, clipping, or resource reuse. OOXPDF's typed chart scene model is already closer to the
  long-term goal and should keep moving chart semantics out of renderer-local XML reads.
- Its testing documentation is more reusable than its chart renderer for OOXPDF's PDF target: the useful
  import is the discipline of generated public cases, PowerPoint ground truth, structural checks before
  visual thresholds, and explicit support catalogs, not browser/ECharts layout behavior.
- Comparing against current OOXPDF shows the migration has already started, not merely been proposed:
  `PptxScene`, scene snapshots, first-class text flow/layout/glyph records, and typed chart scene records
  already provide more PDF-relevant structure than `pptx-renderer`'s DOM/ECharts render layer.
- The main OOXPDF architectural gap is now ownership, not absence of models. `PptxRenderContext` still
  carries freshly loaded slide XML, inherited XML, ad hoc relationship dictionaries, and a nullable
  `SceneSlide`; long-term work should make the scene/context boundary own resolved part, inheritance,
  relationship, and resource state so specialized PDF renderers no longer need renderer-local XML reads.

High-priority actions:

- [ ] 2026-05-30: Continue page-81 text-state work from the generic typography branch, without adding
  page-specific geometry logic. Inspection showed Office and candidate have matching text operation counts
  (`83/83`) and matching graphics operation counts (`107/107`), with identical graphics operator buckets
  (`95` clips, `8` strokes, `2` nonzero fills, `2` even-odd fills). The visible residual is therefore not a
  missing shape. The color-transform midpoint mismatch is closed; the remaining dominant branch is the
  already-open Office text-state issue (`Tc`, secondary `/Tf`, and text-operation splitting). Keep this work on
  the public typography ladder and avoid private text or coordinate shortcuts.
- [ ] 2026-05-31: Fold the remaining page-44 residual into the shared Office text-emission track, not the
  chart-geometry track. The page-44 chart pass closed the major structural mismatches through Office-like
  stacked-column plot padding, image-backed diagonal pattern tiles, denser value-axis ticks, and packed
  horizontal legend spacing. Private run `20260531-134932` improved page 44 to `1.281596258` MAE,
  `0.032060185` changed16, and `0.973682740` SSIM. Fresh PDF inspection shows the chart value-axis labels,
  plot box, bars, slanted connector strokes, hatch fills, and bottom legend baseline are now close enough that
  the dominant changed-pixel bins are title/body/chart text glyph areas rather than chart geometry. Keep page
  44 as private acceptance evidence for the existing `Tc`/secondary-`Tf`/operation-splitting typography work;
  do not add more private chart offsets unless a new PDF graphics inspection shows a real structural geometry
  mismatch.
- [ ] 2026-06-01: Continue DOCX vertical-flow work from Office-observed page text and table structure, not
  from a broad `docGrid` shortcut. Private DOCX inspection showed the acceptance document has section
  `w:docGrid/@w:linePitch` and no explicit `w:snapToGrid` overrides, so the reader now preserves
  `docGrid`/`snapToGrid` tokens in the model. A local Office probe with ordinary Latin paragraphs, zero
  paragraph spacing, `w:docGrid w:linePitch="480"`, explicit `w:snapToGrid`, and `w:useFELayout` still showed
  Office baseline advances of about `13.32` pt rather than the `24` pt grid pitch. A trial candidate line-grid
  rule improved aggregate private MAE but spilled to an extra page, so it was rejected. Keep this item open:
  the next acceptable implementation path is public Office probes for table row height/pagination, paragraph
  spacing collapse, and final-page table flow before applying any grid-derived layout behavior.
  2026-06-01 follow-up: public `docx-ladder-02-table-cell-margins` showed Office aligns row text to the
  maximum top cell margin across the row, not each cell's own smaller top margin independently. `DocxLayout`
  now uses row-level top padding for cell text and inline-image placement. Public cell-margin MAE improved
  `0.45 -> 0.39`; `docx-ladder-03-table-row-heights` and `docx-ladder-03-table-pagination-margins` stayed in
  the same metric band. Private DOCX run `20260601-170045` stayed neutral at `16/16` pages, zero dimension
  mismatches, no diagnostics, `MAE=13.666634`, changed16 `0.126275`.
  2026-06-01 follow-up: source-indexed layout snapshots now expose private-safe `SourceBlockIndex` and
  `SourceLineIndex` on body text lines. The acceptance document's worst private pages `9..11` are ordinary
  paragraph flow, not table flow: candidate body lines start at the same top coordinate as Word but continue
  lower on the page, and long body paragraphs fit into suspiciously few candidate lines. Continue this branch
  through Office-observed text wrapping and typeface/advance measurement before changing widow/orphan or
  page-bottom rules.
  2026-06-01 follow-up: added public `docx-ladder-02-font-table-alternate-wrapping` to cover the private
  document's dominant font-table-alternate pattern without private content. The probe uses a missing primary
  Latin face with a declared `w:font/w:altName` alternate and a direct-alternate control paragraph. Office and
  candidate line breaks already matched, but PDF inspection exposed a structural emission gap: candidate text
  lines appended an extra standalone terminal-space operation when the wrapped line text already ended in
  whitespace. `DocxRenderer.RenderTerminalLineSpace` now only appends the terminal space when the last visible
  segment does not already end in whitespace. Public run `20260601-171530` passes with matching text-line
  starts (`15/15`, zero deltas at `0.2pt` tolerance); private DOCX run `20260601-171546` stayed neutral at
  `16/16` pages, zero dimension mismatches, no diagnostics, `MAE=13.666634`, changed16 `0.126275`. This
  improves PDF-level structural alignment but does not close the private raster gap; keep the open branch on
  actual advance/wrapping differences across the worst private pages.
  2026-06-01 follow-up: rejected a DOCX trial that measured wrap widths using the Office PDF 600-DPI emitted
  font-size grid. It preserved ordinary paragraph tests but regressed the public font-table alternate
  structural text-line gate: the same-line secondary text start moved from about `0.12pt` off to `0.83pt`.
  Keep nominal font-size measurement for DOCX wrapping unless a public Office case proves otherwise. Private
  block-level mapping now points the page-10 drift back to earlier page-9 paragraphs that consume extra
  candidate lines and push a later block to the next page. `DocxInspect` block summaries now include
  paragraph indent points and character buckets; the suspect blocks have no paragraph indents and ordinary
  Latin/space/punctuation profiles, so the next branch is precise Calibri advance/word-break behavior rather
  than page-bottom, indent, or charset logic.
  2026-06-01 follow-up: added public `docx-ladder-02-calibri-body-wrapping` for A4-width 10pt Calibri body
  paragraphs with `w:spacing line=276 lineRule=auto` and 8pt after-spacing, matching the private failure
  surface without private text. Office and candidate both produce `15` text rows with first starts within
  `0.02pt`; the remaining PDF difference is Office splitting some rows into additional text operations.
  `ComparePdfTextLineStarts.ps1` now supports `-FirstStartOnly`, and visual manifests can set
  `compareFirstTextLineStartOnly` to gate wrapping/baseline structure separately from intra-line operation
  splitting. This public case does not reproduce the private over-wrapping by itself, so keep searching for
  the missing discriminator rather than widening body text or adding private page-flow constants.
  2026-06-01 follow-up: added public `docx-ladder-02-compact-before-spacing` for the private-like compact
  paragraph rhythm (`before=36`, `after=0`, `line=276`, 10pt Calibri). Office and candidate produce six
  matching line rows with first starts within `0.02pt`, so the page-9 drift is not caused by the compact
  before-spacing rule alone. A trial that split DOCX wrap tokens after hyphens/slashes also failed to help
  the private run (`MAE=13.698394` vs the `13.666634` baseline), so it was rejected pending a public case
  that specifically proves Word break behavior beyond whitespace.
  2026-06-02 follow-up: public `docx-ladder-02-long-token-wrapping` now supplies that missing case. PDF
  inspection showed body paragraph rows already matched, while Office decomposed hyphenated body words into
  smaller text operations and split overwide alphabetic table-cell tokens (`CellAlphaPlanningTok`/`en21`,
  `BoundaryMarkerOme`/`ga34`). `DocxLayout` now keeps preferred punctuation break opportunities first, then
  falls back to a general Unicode-safe character boundary only when a non-whitespace token itself exceeds the
  available table-cell line width. This improved the public visual run `20260602-003527` to `MAE=1.188618`,
  changed16 `0.017627`, SSIM `0.723631` from the previous `MAE=1.282270`, changed16 `0.018130`, SSIM
  `0.682413`.
  2026-06-02 follow-up: `DocxWrappedTextLine` and `DocxTextLineLayout` now preserve whether a line ended at
  an intra-token emergency break, and `DocxRenderer` uses that structural flag to avoid synthesizing
  Office-like terminal line-space operations after split token prefixes. Public-safe bottom-up coverage checks
  both the layout break reason and the text-emission snapshot without exposing document text. Public visual
  run `20260602-020006` stayed raster-neutral (`MAE=1.188618`, changed16 `0.017627`, SSIM `0.723631`), while
  parsed PDF inspection confirmed the candidate sequence now emits `CellAlphaPlanningTok` directly followed by
  `en21...`, and `BoundaryMarkerOme` directly followed by `ga34...`, with no standalone space operation between
  split fragments. This still left body hyphenated words as a PDF text-operation decomposition issue rather
  than a visible wrapping failure.
  2026-06-02 follow-up: DOCX text emission now splits Unicode dash punctuation into separate Office-like PDF
  text operations, using the same resolved text measurer that owns layout origins so wrapping and visual
  placement stay unchanged. Public-safe coverage asserts `word-break` emits visible operation lengths
  `4/1/5`, and public run `20260602-020347` improved the long-token case to `MAE=1.121729`, changed16
  `0.017126`, SSIM `0.745218`. Parsed Office/candidate text-operation order now matches through the body
  hyphenated words and overwide table token fragments. Treat this long-token operation-order branch as closed;
  any remaining raster gap on this case belongs to broader glyph metrics/positioning, not wrapping, terminal
  spaces, or dash decomposition.
  2026-06-02 follow-up: DOCX body hyperlinks now lower through explicit PDF link annotations instead of being
  only parser/structure metadata. The PDF writer has a first-class `PdfLinkAnnotation` primitive with stable
  object numbering, `/Annots` page wiring, rectangle emission, URI actions, and PDF-string escaping. DOCX
  layout now carries source text-run indexes through span slicing, wrapping, justification, and text-operation
  splitting so body hyperlink annotations are anchored to placed rendered segments. Table-cell and selected
  static header/footer text lines use the same annotation path for external URI links. Keep internal PDF
  destinations open for a later story-owned annotation model.
  2026-06-01 follow-up: added private-safe `tools/CompareDocxLayoutPdfFlow.ps1`, which maps candidate layout
  source block/line indices to Office/candidate PDF text rows using decoded text internally but emits only
  lengths, hashes, pages, and coordinates. The first all-page private flow map shows page shifts recurring
  around bottom-of-page paragraph/table transitions, not a single bad table row. A public
  `docx-ladder-03-widow-page-boundary` probe then tested the tempting hypothesis that a four-line paragraph
  with only one line of remaining page space should split; Office and candidate both move the paragraph to the
  next page. Keep the current widow-control rule and continue with public probes for accumulated paragraph
  line-height, empty-paragraph, table-boundary, and table-row flow instead of weakening widow/orphan handling.
  2026-06-01 follow-up: added public `docx-ladder-03-table-empty-paragraph-boundary` after the private flow map
  highlighted empty body paragraphs after tables. Office does preserve the empty paragraph as vertical flow in
  this public table -> empty paragraph -> body paragraph sequence; candidate and Office match operation counts
  with only a small baseline residual (`MAE=0.165895`, changed16 `0.002063`). Do not suppress after-table empty
  paragraphs as a shortcut. The remaining vertical-flow work should focus on row/page transition heights and
  accumulated paragraph/table spacing, not dropping authored empty blocks.
  2026-06-01 tooling correction: `tools/ComparePdfTextLineStarts.ps1` now groups and matches text rows per page
  before comparing starts. The previous page-agnostic grouping could hide pagination mistakes when two pages had
  similarly placed rows. The existing DOCX wrapping gates still pass with the page-aware comparer, so future
  pagination probes can safely use line-start gates without masking page shifts.
  2026-06-01 settings inventory: private-safe inspection now exposes DOCX document settings and Word compatibility
  facts such as `characterSpacingControl`, `defaultTabStop`, `useFELayout`, and named `compatSetting` entries.
  The current private acceptance document carries modern compatibility mode, enabled FE layout, and
  non-compressing character-spacing control. These are now preserved in `DocxDocument.Settings` and emitted by
  `DocxInspect`, but rendering remains unchanged until public Office probes show which setting changes layout
  or PDF text-state behavior.
  2026-06-01 default-tab progress: DOCX layout now uses `w:settings/w:defaultTabStop` for default tab advances
  instead of the fixed 36pt fallback when no explicit positioning tab stop applies. Public
  `docx-ladder-02-default-tab-stop-settings` locks a 72pt default tab grid against Office, with visible text
  starts matching within `0.02pt`; the line-start gate now has an option to ignore whitespace-only spacer
  operations because Word emits the tab gap as a separate whitespace `TJ`. Private acceptance stays neutral
  because its default tab stop is the normal 36pt (`20260601-180758`: `16/16` pages, no diagnostics,
  `MAE=13.666634`, changed16 `0.126275`).
  2026-06-01 row-minimum progress: private-safe PDF graphics inspection traced the next large DOCX page-flow
  drift to table rows without `w:tblPrEx`: Office fit more compact rows on the page, while the candidate
  injected a hard-coded 401-twip default row minimum and pushed later text down by roughly 60pt. Added public
  `docx-ladder-03-table-no-row-exceptions`, plus companion `docx-ladder-03-table-cell-before-spacing` and
  `docx-ladder-03-table-compact-row-pagination`, then removed the implicit default row minimum so auto-height
  rows are governed by measured cell content unless OOXML declares a row height. The no-row-exception public
  case improved from MAE `3.065771`, changed16 `0.102824` to MAE `1.107443`, changed16 `0.010343`; existing
  public `docx-ladder-03-table-row-heights` stayed at MAE `0.700136`, and `docx-tables` stayed at MAE
  `0.455760`. Private acceptance run `20260601-182935` improved from prior `MAE=13.666634` to
  `MAE=9.982157`, with `16/16` pages, no diagnostics, and zero dimension mismatches. Keep this branch open:
  row geometry is closer, but remaining worst pages point to broader table/text flow alignment rather than a
  single-row minimum.
  2026-06-01 follow-up: the next private-safe flow comparison kept the accepted private run at `16/16` pages
  and showed the remaining page-9/page-10 drift is accumulated compact bullet paragraph rhythm before a later
  page break, not a fresh table-row minimum. A public two-line widow probe (`docx-ladder-03-widow-two-line-boundary`)
  showed Word allows a four-line paragraph to split `2/2`, so do not broaden widow/orphan control for the
  private block split. Public `docGrid` probes (`docx-ladder-03-docgrid-line-pitch`,
  `docx-ladder-03-docgrid-use-fe-layout`, and `docx-ladder-03-docgrid-list-use-fe`) were visually neutral,
  including with `useFELayout`, so `docGrid`/FE layout alone is not the missing compact-list spacing rule.
  A public compact bullet fixture (`docx-ladder-03-compact-bullet-spacing`) exposed one real horizontal
  numbering bug: Word places the marker at `left - hanging` and uses the numbering tab/left position for the
  following paragraph text. The renderer now follows that structure. The fixture improved slightly
  (`MAE=0.810522` to `0.803030`), `docx-numbering` and `docx-ladder-03-docgrid-list-use-fe` stayed unchanged,
  and private run `20260601-185704` stayed neutral at `MAE=9.982157`, changed16 `0.103396`. Keep the open
  vertical-flow target on compact bullet paragraph advance/spacing and Word text-state decomposition; do not
  turn numbering tabs, `docGrid`, FE layout, font names, or private style names into shortcuts.
  2026-06-01 follow-up: added public `docx-ladder-03-compact-bullet-alt-bottom` to isolate compact bullet
  pagination when the requested primary Latin face is missing and `w:font/w:altName` resolves the actual face,
  matching the dominant private-safe font-table-alternate pattern without private content. The public case
  reproduces a real compact-list pagination/line-pitch discrepancy: candidate page 1 is still high-error
  (`MAE=15.916448`, changed16 `0.154725`) while page 2 is much closer (`MAE=0.750764`, changed16 `0.007434`).
  A trial that replaced adjacent compact-list before-spacing with a `docGrid`-derived total pitch was rejected:
  it worsened this public fixture (`MAE=18.261426` on page 1), worsened `docx-ladder-03-compact-bullet-spacing`
  (`MAE=0.914388` vs accepted `0.803030`), and worsened the private aggregate (`MAE=10.578667` first page in
  the rejected run). `DocxLayoutSnapshot` now emits private-safe `LineHeightPoints`,
  `AppliedBeforeSpacingPoints`, and `IsFirstParagraphLine` for text lines so the next step can compare actual
  layout advances against Office PDF rows instead of guessing from style names or font names. Keep this branch
  open on deriving Word's effective single-line metrics and paragraph pitch for compact lists under
  font-table-alternate resolution; do not use a broad `docGrid`, FE-layout, or named-font shortcut.
  2026-06-01 tooling follow-up: `tools/CompareDocxLayoutPdfFlow.ps1` now carries those line-height and applied
  before-spacing fields into its private-safe flow map and emits aggregate candidate line-advance buckets. The
  current private flow map has dominant candidate advance buckets at `14.038`, `20.038`, `22.038`, and
  `15.838` pt, with the compact-list first lines in the `15.838` bucket. The public compact-bullet-alt-bottom
  repro has `17.059` pt candidate compact-list advances while Office rows are closer to `16.4..16.6` pt.
  This confirms the remaining rule is not a single broad page-fit constant; continue by deriving Office's
  paragraph pitch from public rows and font metrics, then apply only a structural rule that also explains the
  private compact-list bucket.
  2026-06-02 follow-up: public compact-bullet flow comparison traced the first-list baseline error to the
  default paragraph model, not to a list/font/style exception. Paragraphs with no explicit
  `w:spacing/@w:line` now use an Office-observed `1.2` auto-line factor, and paragraphs with no authored
  after-spacing side use an `8pt` default after spacing instead of the previous `6pt`. This explains both the
  untokened title paragraph and the compact list body without naming a style or font. Public
  `docx-ladder-03-compact-bullet-alt-bottom` improved from page-1 `MAE=15.916448`, changed16 `0.154725` to
  `MAE=11.188275`, changed16 `0.120160`; page 2 stayed at the low-error profile (`MAE=0.253264`, changed16
  `0.002815`). Public `docx-ladder-03-compact-bullet-alt-line115` improved from page-1 `MAE=15.195190` to
  `12.164815`, and `docx-ladder-03-compact-bullet-spacing` improved from the accepted `MAE=0.803030` to
  `0.626532`. Flow run `20260602-033223` has exact `43/43/43` layout/reference/candidate row counts, zero
  missing or ambiguous matches, candidate effective factor `1.2` for all rows, title advance `22.648pt` vs
  Office `22.680pt`, and compact row-advance mean without max `16.44845` vs Office `16.4755`. Keep the branch
  open for residual baseline/advance quantization and text-state decomposition, but do not broaden this into
  a docGrid, FE-layout, font-name, or style-name rule.
  2026-06-01 negative evidence: rejected an OpenType metric trial that applied the automatic line-spacing
  factor to the typographic ascender/descender body and added typographic line gap only once. It improved the
  public font-table-alternate compact-list stress case (`docx-ladder-03-compact-bullet-alt-bottom` page 1
  `MAE=15.916448 -> 13.390454`) but regressed the existing compact-bullet fixture (`MAE=0.803030 -> 0.978554`)
  and badly regressed private acceptance (`MAE=9.982157 -> 13.879306`, worst page shifted to page 9). Do not
  replace the global DOCX auto-line metric this way; any future metric change needs a discriminator visible in
  OOXML/font state and must improve both compact-list public cases plus the private acceptance aggregate.
  2026-06-01 follow-up: `DocxFontPlanSnapshot` and `DocxInspect` now emit private-safe resolved-font metric
  buckets using resolved-family hashes, source buckets, and font sizes, plus block-level `LineSpacingFactor`.
  This showed the public font-table-alternate compact case and the private dominant alternate path share the
  same resolved-family hash; the visible discriminator is line factor (`1.25` in the original public stress
  case, `1.15` in the private compact blocks). Added public `docx-ladder-03-compact-bullet-alt-line115` with
  font-table alternate plus `line=276 lineRule=auto`: Office renders two pages while the candidate fits all
  rows on one page, reproducing the private failure shape. A trial list-only minimum auto factor fixed that
  public page-count mismatch but still worsened private acceptance (`MAE=9.982157 -> 10.156234`), so it was
  rejected. Keep the fixture; the next implementation must explain why Word applies a larger effective pitch
  only at some compact-list boundaries without shifting the whole private document.
  2026-06-01 progress: the accepted rule is a narrower auto-spacing floor, not a page-bottom reserve and not a
  global font metric change. For list paragraphs with an authored positive before-spacing and auto line
  spacing below `1.19`, `DocxLayout` now uses `1.19` as the minimum effective auto-line factor. This preserves
  the existing `1.25` compact-list and zero-before docGrid cases while matching the private-like `line=276`
  compact-list page split. Public `docx-ladder-03-compact-bullet-alt-line115` improved from a page-count
  mismatch and page-1 `MAE=17.230239` to matching page count with page-1 `MAE=11.440381`; page 2 improved from
  missing to `MAE=0.253264`. A `1.20` probe improved the public page-1 metric further (`MAE=11.440381`) but
  was worse on private acceptance (`MAE=9.386077`), while a `1.18` probe regressed both private (`MAE=8.977391`)
  and the public page-1 metric (`MAE=18.212437`), so `1.19` is the current accepted structural floor. Guards
  stayed neutral: `docx-ladder-03-compact-bullet-alt-bottom` remained at
  page-1 `MAE=15.916448`, `docx-ladder-03-compact-bullet-spacing` at `MAE=0.803030`, and
  `docx-ladder-03-docgrid-list-use-fe` at `MAE=0.019271`. Private acceptance improved from `MAE=9.982157` to
  `MAE=8.915684`, with pages 9..11 notably lower; page 15 is now the worst remaining private page.
  2026-06-02 architecture follow-up: `DocxLayout` now resolves line height through an explicit
  `DocxLineHeightProfile` and carries private-safe `SingleLineHeightPoints`, `EffectiveLineSpacingFactor`,
  and `LineSpacingFactorFloorApplied` into body and table text-line snapshots. A bottom-up
  `DocxLayoutSnapshotReportsLineHeightProfileFacts` test locks the accepted compact-list floor predicates
  (list paragraph, authored positive before-spacing, auto line rule, requested factor below `1.19`) and proves
  that zero-before lists and non-list paragraphs do not report the floor. This is diagnostic/architectural
  groundwork, not a rendering change. Use it next to compare public compact-list and table-cell line pitches
  against Office PDF rows before changing metrics, row fragmentation, or page-bottom rules.
  2026-06-02 tooling follow-up: `tools/CompareDocxLayoutPdfFlow.ps1` now includes the same line-height profile
  fields in mapped rows and emits effective-factor/floor-applied buckets. Public
  `docx-ladder-03-compact-bullet-alt-line115` run `20260602-021217`, with fresh PDF and DOCX inspections,
  shows `42` candidate layout lines at effective factor `1.190000` with the floor applied and `1` control line
  at `1.250000` without the floor. This makes the accepted list-floor decision visible in the public
  layout/PDF flow map; continue by comparing those candidate pitches to Office row deltas before touching
  metrics or pagination rules.
  2026-06-02 evidence update: the same flow summary now emits PDF row-advance buckets for reference and
  candidate rows. On public `docx-ladder-03-compact-bullet-alt-line115` run `20260602-021217`, candidate
  layout/PDF advances are dominated by `16.326` pt, while Office reference PDF rows split between `16.320`
  (`23` rows) and `16.440..16.470` (`17` rows), plus one larger paragraph-boundary gap. The accepted `1.19`
  floor is therefore necessary but not sufficient: the remaining compact-list discrepancy appears to be an
  Office row-rhythm/line-box alternation that is observable in PDF structure, not a page-bottom reserve or a
  new global line-height constant. Find the OOXML/font/layout discriminator before changing rendering.
  2026-06-02 evidence update: the flow summary now reports row-advance means as well as buckets. On the same
  public run, the title-to-first-list-row gap is `22.68` pt in Office vs `21.259` pt in the candidate, while
  the repeated list rhythm after excluding that largest boundary gap averages `16.3735` pt in Office vs
  `16.32635` pt in the candidate. The fixture's OOXML is intentionally uniform across list paragraphs, so the
  next probe should isolate paragraph-boundary spacing after a non-list paragraph separately from repeated
  same-style list rhythm before changing either default paragraph after-spacing or the auto-line metric.
  2026-06-02 tooling follow-up: `CompareDocxLayoutPdfFlow.ps1` now normalizes equivalent extracted bullet
  encodings (`w:lvlText` private-use Symbol bullets vs Office `•`, and diagnostic spacer tokens) when hashing
  rows. This is comparer-only text canonicalization, not renderer logic. On the same public compact-list run,
  source-indexed mapping improved from `1/43` rows matched to `43/43`; the worst matched deltas now show
  accumulated source-block drift reaching about `3.41` pt by list items 40-41. Use this per-source mapping to
  validate any future compact-list rhythm rule.
  2026-06-02 architecture follow-up: `DocxFontPlanSnapshot` now includes private-safe resolved OpenType
  metric values in each metric bucket (`UnitsPerEm`, typographic ascender/descender/line-gap, Windows
  ascender/descender, and derived point metrics) while still hashing resolved family names. Bottom-up coverage
  verifies these fields against an installed OpenType face without exposing the family. Regenerating
  `docx-ladder-03-compact-bullet-alt-line115` inspection showed the dominant body/title font-table-alternate
  bucket at 10 pt has single-line height `12.20703125` pt, and the current accepted `1.19` floor therefore
  produces candidate repeated pitch `12.20703125 * 1.19 + 1.8 = 16.3263671875` pt. Office's repeated-row mean
  excluding the largest paragraph-boundary gap is `16.3735` pt. The nearby numbering-label Windows
  ascent+descent bucket is about `12.2509765625` pt, so the next implementation branch should explicitly
  investigate line-box metric ownership between paragraph body runs and list-label runs. Do not special-case
  the resolved family,
  the bullet character, or this numeric residual; the rule must follow Word's structural treatment of list
  labels versus paragraph content and be validated against both compact-list public fixtures plus private
  aggregate flow.
  2026-06-02 follow-up: body-vs-list-label metric candidates are now carried directly on text-line layout
  snapshots and through `tools/CompareDocxLayoutPdfFlow.ps1` (`ListLabelSingleLineHeightPoints`,
  `BodyWindowsLineHeightPoints`, `ListLabelWindowsLineHeightPoints`, plus aggregate body/label Windows
  buckets). Public `docx-ladder-03-compact-bullet-alt-line115` confirms body/title rows at
  `BodyWindowsLineHeightPoints=12.207031` and list labels at
  `ListLabelWindowsLineHeightPoints=12.250977`. A structural trial that selected
  `max(body single-line, distinct list-label Windows extent)` for auto list line boxes was rejected: it kept
  the case passing but worsened page-1 MAE to `13.533560`, and after reverting behavior the diagnostic-only
  run `20260602-023605` remained passing at page-1 `MAE=15.195190`, page-2 `MAE=0.253264`. Keep the evidence,
  but do not promote label Windows extents directly into line height until another Office-PDF probe explains
  the row-rhythm alternation and first list boundary gap together.
  2026-06-02 accepted default-model update: missing paragraph spacing now follows the public Office-observed
  Word defaults (`1.2` auto-line factor and implicit `8pt` after spacing), replacing the earlier split between
  untokened and spacing-token paragraphs. Public compact-list probes improved or stayed guarded
  (`docx-ladder-03-compact-bullet-alt-bottom` page 1 `MAE=11.188275`, page 2 `MAE=0.253264`;
  `docx-ladder-03-compact-bullet-alt-line115` page 1 `MAE=12.164815`, page 2 `MAE=0.253264`;
  `docx-ladder-03-compact-bullet-spacing` `MAE=0.626532`). Private DOCX acceptance run `20260602-033549`
  stayed valid at `16/16` pages, zero dimension mismatches, no diagnostics, and improved aggregate
  `MAE=8.927687 -> 8.908377` with changed16 `0.095249`. Keep the active residual on row-advance/baseline
  quantization and Office PDF text-state decomposition; do not reintroduce list-label font ownership, a
  page-bottom reserve, or a new residual constant.
  2026-06-02 architecture follow-up: DOCX text emission snapshots now carry a private-safe segment role
  (`ListLabel`, `ListSeparator`, or `Text`) from layout through PDF emission, and
  `tools/SummarizeDocxTextState.ps1` buckets planner/reference pairs by that role. This removes the need to
  infer list-marker behavior from digits, punctuation, or nonzero `Tc` when comparing Office PDF text state.
  Bottom-up numbering coverage also now asserts the accepted Office-observed numbering-tab structure:
  marker geometry stays at `left - hanging`, while the following paragraph text advances to the
  numbering-tab/left target. Keep this distinction explicit in future compact-list and text-state work.
  2026-06-02 PDF-oracle follow-up: `PdfInspect` now expands Office `/ObjStm` object streams, reads Type0
  descendant `/W` and simple `/Widths` font width maps, and emits `NaturalWidthPoints` plus
  `EmittedAdvancePoints` for text operations. `tools/SummarizeDocxTextState.ps1` now prefers run-local
  `comparison/pdf-text` inspections and buckets Office natural/emitted widths against candidate planner
  font-size and glyph-pair side-advance ranges. Public text-state probes
  `docx-ladder-03-text-state-size-matrix`, `docx-ladder-03-table-text-state`,
  `docx-ladder-03-text-state-context`, and `docx-ladder-03-text-state-font-matrix` all pair cleanly with zero
  missing Office width buckets after refreshed inspection (`65`, `35`, `35`, and `65` planner/reference pairs).
  This strengthens the next DOCX `Tc` branch: compare Office emitted advance decomposition against the
  renderer's planned natural width and spacing state, rather than deriving a lookup from text content, font
  names, table roles, or observed bucket constants.
  2026-06-02 evidence refinement: the summary now also emits Office-minus-planner natural/layout/rounded width
  deltas and per-gap deltas. Across those four public runs, every nonzero Office `Tc` bucket matched
  `ReferenceEmittedAdvance - PlannerRoundedWidth` divided by the planner glyph-gap count after rounding
  (`19`, `2`, `7`, and `17` nonzero per-gap keys respectively, zero mismatches). This is not yet a production
  rule because the renderer does not know Office's emitted advance target; it is the invariant the next
  architecture slice must explain from OOXML/font/PDF state instead of from private data or font names.
  2026-06-02 architecture follow-up: `DocxTextEmissionAdvanceProfile` now decomposes the candidate's planned
  PDF emitted advance into rounded width-array advance, kerning `TJ` adjustment, positioned-spacing `TJ`
  adjustment, and `Tc` gap total. Regenerated public text-state runs
  `20260602-092913`/`092917`/`092921`/`092925` plus fresh `DocxInspect` sidecars show the same invariant against
  the candidate plan directly: for gap-positive nonzero Office `Tc` keys, `(ReferenceEmittedAdvance -
  PlannerEmittedAdvance) / GlyphGapCount == ReferenceTc` across all four probes (`19`, `2`, `5`, and `16`
  keys, zero mismatches). Keep this as the structural target for the next renderer branch; do not turn it into
  an oracle lookup, because production still needs to derive Word's emitted-advance target from document/font
  state available before PDF emission.
  2026-06-02 planner follow-up: `DocxTextEmissionPlanner` now has a target-advance primitive that derives the
  PDF text-state character spacing needed to bridge a current emitted advance to a caller-supplied target over
  the operation's glyph gaps, with separate compensated and uncompensated plan paths. This is intentionally
  target-agnostic: it gives future Word-target logic a structural home without deciding the target from
  private data, fonts, or observed constants. `docx-core --skip-slow` passed `50` tests after a serial rerun
  following one transient parallel compiler output lock.
  2026-06-02 tooling follow-up: `tools/SummarizeDocxTextState.ps1` now auto-runs `InspectDocx.ps1` for public
  visual runs when `comparison/docx-inspect` is missing or when an existing `text-emission-snapshot.json` is
  stale and lacks `PlannedEmittedAdvance`. Private/local runs remain opt-in via existing sidecars. A refresh
  test on older public run `docx-ladder-03-text-state-font-matrix/20260602-091115` regenerated the sidecar and
  produced `65` planner/reference pairs with no unexpected missing emitted-delta fields.
  2026-06-02 architecture follow-up: `DocxLayout` now resolves paragraph boundary spacing through an explicit
  `DocxParagraphSpacingProfile` and carries private-safe `PendingAfterSpacingPoints`,
  `ParagraphBeforeSpacingPoints`, `ParagraphAfterSpacingPoints`, and `ContextualSpacingSuppressed` on first
  paragraph lines across body, static header/footer, and table-cell text. `CompareDocxLayoutPdfFlow.ps1` now
  maps those fields and emits paragraph spacing profile buckets. This is diagnostic groundwork for separating
  the first list-boundary gap from repeated compact-list row rhythm; do not turn it into a new spacing rule
  until the Office PDF row deltas point to a structural discriminator.
  2026-06-02 coverage follow-up: bottom-up DOCX tests now assert the paragraph spacing profile directly on
  contextual-spacing boundaries. Same-style adjacent paragraphs expose nonzero pending/author before-after
  values but zero applied before-spacing with `ContextualSpacingSuppressed=true`; different-style adjacent
  paragraphs keep the authored gap with suppression false. This locks the intended Word-style boundary
  structure before using those fields in any compact-list rule.
  2026-06-01 follow-up: added private-safe table-cell text profiles to DOCX layout snapshots: whitespace,
  punctuation, digit/letter, non-ASCII, and longest whitespace-delimited token counts. Page-15 inspection
  showed the worst table-to-heading residual is not a table width or column-position error; the 6-column table
  has Office-aligned x positions, but Office wraps some narrow punctuation-bearing cells into taller bands
  while the candidate keeps those rows shorter. Added public
  `docx-ladder-03-table-punctuation-wrapping` to make this branch public and Office-observable. A broad
  punctuation word-break trial was rejected even though it slightly improved the private aggregate, because the
  public probe showed candidate overwrapping `North-West/2026`-style cells that Office keeps on one row. Keep
  this branch open on effective table-cell text measurement/fit and Office row-band structure, not on a
  punctuation shortcut or a row-height constant. With the rejected break rule reverted, the new public probe is
  an accepted open mismatch at `MAE=2.575378`, and private acceptance returns to the current baseline
  `MAE=8.915684`.
  2026-06-01 follow-up: PDF inspection of the public probe showed the first real structural miss was table
  width resolution, not punctuation itself: Office expands an underfilled explicit percent grid to the content
  width plus the first/last outer cell insets, while the candidate had clamped all percent tables to the text
  width. `DocxLayout` now applies that inset basis only when an explicit grid is below the normal percentage
  target, preserving private tables whose grid already meets/exceeds the percent target. With that corrected
  width basis, overwide-token breaks are enabled only inside table-cell wrapping and only after proving the
  whole non-space token is too wide; fitting tokens such as `North-West/2026` stay on one row, while an
  overwide `dash-separated-value`-style token can break at the rightmost fitting punctuation opportunity.
  Public `docx-ladder-03-table-punctuation-wrapping` improved from `MAE=2.575378` to `MAE=0.662548`, SSIM
  `0.935016`, and table guards stayed neutral (`docx-ladder-03-table-row-heights` `MAE=0.700136`,
  `docx-ladder-03-table-pagination-margins` pages `0.770109/0.171589`). Private run `20260601-201837`
  stayed valid at `16/16` pages, zero dimension mismatches, no diagnostics, and moved page 15 from
  `MAE=12.157179` to `11.744805`; aggregate MAE moved slightly from `8.915684` to `8.927676` because page 16
  worsened while pages 14/15 improved. Keep the branch open on the post-table/page-16 flow residual and PDF
  text-state decomposition; do not broaden punctuation breaks outside table-cell overwide fitting.
  2026-06-01 follow-up: private-safe PDF graphics inspection of `20260601-201837` narrowed the page-16
  residual to table row allocation, not ordinary post-table paragraph spacing. Office shows five shaded row
  bands on page 16 for the continued table, while candidate shows four; candidate keeps one additional
  unshaded row at the bottom of page 15 with its lowest text baseline around `77.31pt`, below Office's
  observed low text baseline around `81.24pt`. A broad trial that added a font-size-derived table-row bottom
  inset preserved public table guards but badly regressed the private acceptance run (`8.927676 -> 12.078142`
  average MAE, worst pages shifting by `+4..+7` MAE) and did not move the intended table-row boundary in the
  layout snapshot, so it was rejected. The next acceptable slice is a public Office probe that isolates
  table-row page-boundary allocation with small positive bottom slack, alternating fills, 9pt table text, and
  a following empty paragraph before changing pagination rules.
  2026-06-01 follow-up: added public `docx-ladder-03-table-bottom-slack` as that probe. Office/candidate PDF
  inspection showed the generic missing rule is row fragmentation, not a scalar bottom-margin correction:
  with no `w:cantSplit`, Office places row 11's first line and shaded fragment at the bottom of page 1 and
  continues the row on page 2, while the candidate moves the whole row to page 2. This is the opposite
  allocation direction from the private residual but the same structural weakness: `DocxLayout` treats table
  rows as atomic page items. The next implementation slice should preserve `w:cantSplit`, then introduce
  explicit row-fragment layout records so split rows carry per-fragment cell rectangles, clipped text lines,
  and suppressed continuation borders instead of using a hard-coded page slack heuristic.
  2026-06-01 rejection: a first production row-fragment trial split auto-height shaded rows without vertical
  cell padding and improved `docx-ladder-03-table-bottom-slack` page 2 (`MAE 3.551346 -> 2.889362`), but it
  regressed the private DOCX acceptance run (`8.927676 -> 9.641534` average MAE, worst page `15.784794`) and
  initially disturbed `docx-ladder-03-table-pagination-margins` until declared-height rows were excluded.
  Keep the public probe and `w:cantSplit` preservation, but do not reintroduce row fragmentation until the
  layout model can represent line-box-derived fragment heights, declared-height row behavior, and continuation
  border suppression from public Office evidence rather than clipping full-row layouts after the fact.
  2026-06-02 progress: reintroduced row fragments through that narrower model instead of through a bottom
  slack shortcut. `DocxLayout` now only splits non-`w:cantSplit` rows when the cell line boxes prove that text
  exists on both sides of the page boundary, records one `DocxTableRowLayout` per fragment, keeps full-row text
  coordinates for clipping, and suppresses same-logical-row continuation borders in the renderer. Bottom-up
  tests cover default multi-page row splitting and `w:cantSplit` preservation. This closes the earlier
  contract-only phase: split rows are now first-class layout/render/snapshot objects, not an inferred
  page-bottom slack behavior. Keep the rejected-trial evidence above because it still explains why the split
  predicate must stay tied to observed line boxes and why text-bearing top fragments with empty continuations
  remain risky.
  2026-06-02 follow-up: split-row continuation pages now use the same repeated-header-row path as ordinary
  row page breaks, and fragment heights reserve that repeated header height on continuation pages. This keeps
  the table layout model structural: headers remain explicit `DocxTableRowLayout` records, while the carried
  body row stays one logical row with per-page fragments. Added bottom-up layout coverage for a header row plus
  a body row split across a page boundary; validation passed `docx-tables --skip-slow` (`85`). Keep the open
  branch on Office's row-boundary decision itself, plus vertical-merge and inline-image behavior inside
  fragments.
  2026-06-02 follow-up: split-row table cells now keep inline image layouts in full-row coordinates and filter
  them by overlap with each fragment rectangle, relying on the existing renderer cell clip path instead of
  dropping images whenever `FragmentCount > 1`. Added bottom-up coverage for a split row containing text plus
  an inline image; validation passed `docx-tables --skip-slow` (`86`). Keep vertical-merge fragment behavior
  open separately because merged-cell height ownership crosses logical rows, not just fragments of one row.
  2026-06-02 follow-up: split-row fragments now separate vertical-merge coordinate ownership from page
  clipping. A `w:vMerge restart` cell still lays out text/images in the full merged-cell span, but when the
  restart row is physically split across pages, each fragment exposes only its own visible row-fragment
  rectangle to the renderer clip path instead of reusing the full cross-row span as the clip. Added bottom-up
  coverage for a split merged restart row followed by a continuation row; validation passed
  `docx-tables --skip-slow` (`87`). Keep the broader cross-page vertical-merge branch open for continuations
  whose own logical rows cross page boundaries, because that needs an explicit merged-cell fragment model.
  2026-06-02 validation update: the row-fragment architecture is now covered by `docx-tables --skip-slow`
  (`92` passing tests), including default row splits, repeated headers before continuations, inline images
  clipped by row fragments, vertical-merge restart clipping, and `w:cantSplit` preservation. Public
  `docx-ladder-03-table-row-fragment-threshold` run `20260602-024328` passes strict page-count, dimension, and
  empty-diagnostic gates with seven pages, but still has high content-page raster residuals (pages 1/2/4/6
  around `6.93..7.77` MAE). Treat this as solved structural pagination parity, not pixel fidelity: the next
  long-view work is Office-aligned row-boundary selection and fragment-internal text/border geometry, plus an
  explicit merged-cell fragment model for continuations whose logical rows cross pages. Do not replace this
  with a scalar bottom-margin slack, post-table heading gap, or private row-coordinate rule.
  2026-06-02 follow-up: split-row fragments now own only their visible table-cell text lines instead of
  carrying every full-row line into every physical fragment. This is an architecture/PDF-structure correction:
  render-time clipping remains available, but layout snapshots and text-emission enumeration no longer imply
  duplicate hidden text operations for synthetic split rows. Added bottom-up coverage asserting an eight-line
  split row distributes its lines across the two fragments without duplication. Public visual probes stayed in
  the same raster band after the change: `docx-ladder-03-table-row-fragment-threshold` run `20260602-043250`
  page MAE `6.930984/7.482251/0.000000/7.765324/0.011417/7.732639/0.010649`,
  `docx-ladder-03-table-bottom-slack` run `20260602-043237` page MAE `6.534271/2.988830`, and
  `docx-ladder-03-table-heading-table-keepnext` run `20260602-043256` page MAE
  `3.670558/7.759079/1.374267`. Validation passed `docx-tables --skip-slow` (`97`),
  `docx-core --skip-slow` (`43`), and `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal`.
  Keep the row-boundary branch open for the actual Word decision model; this slice only makes fragment
  ownership and downstream diagnostics trustworthy.
  2026-06-02 architecture follow-up: cross-page vertical merges now carry explicit visual ownership. A
  `DocxTableCellLayout` continuation can retain its restart cell as `VerticalMergeOwnerCell`, so renderer
  decisions are no longer based on a single "skip continuation" flag. Same-page continuations remain suppressed
  under the restart span, while page-leading continuation fragments can inherit the restart cell's fill/border
  properties. Bottom-up coverage locks a merged restart at the bottom of page 1 with its continuation on page 2;
  `docx-tables --skip-slow` passes `93`. Keep the broader merged-cell branch open for public Office-PDF
  evidence on continuation-internal border suppression and text carry-over when the merged cell's own content
  crosses pages.
  2026-06-01 follow-up: private-safe page-14..16 flow mapping shifted the page-15 diagnosis away from a
  simple post-table heading gap. Block 208, a `keepNext`/`keepLines` heading between two tables, is `25.665pt`
  higher in the candidate than Office, while the preceding heading block 206 is only `1.414pt` off when it
  follows a paragraph. Deeper page-15 text-row hashing showed the drift starts inside the continued table
  before the heading: candidate row baselines around the same first-column labels are already `~12pt` to
  `~25pt` higher than Office after a narrow row that Office allocates as a taller band. Added public
  `docx-ladder-03-table-heading-table-keepnext` to preserve the structural shape: a multi-page, borderless,
  percentage-width fixed table with compact 9pt cell paragraphs followed by a styled keep-next heading and
  another table. After tightening the probe with generic parenthesized formula-like narrow-cell text matching
  the private-safe row profile, the public case now reproduces a row-boundary pagination discrepancy:
  `20260601-211327` renders three pages in both Office and the candidate, but page 2 is high-error
  (`MAE=9.349741`, changed16 `0.096837`) because Office keeps the last continued-table row on page 2 and
  starts the kept heading at the top of page 3, while the candidate moves that row to page 3 and pushes the
  heading down. Keep the private issue open on Office row-band allocation and table-cell line-break
  measurement in continued narrow tables; do not add a post-table heading spacing constant or a private
  row-coordinate rule. The next production attempt must explain this public bottom-boundary behavior together
  with the rejected row-fragment trial and the private page-15 drift before changing table pagination.
  2026-06-01 rejection: a trial that allowed auto-height rows to cross the bottom margin when their laid-out
  text baselines stayed above the margin did not move `docx-ladder-03-table-heading-table-keepnext`, did not
  improve `docx-ladder-03-table-bottom-slack`, and added another ad hoc fit predicate. Reverted it. The next
  attempt should model Office's row page-boundary decision from row line boxes and continuation fragments,
  including rows whose visible text line boxes can remain inside the body while the row band crosses the
  nominal margin; do not add another scalar slack test.
  2026-06-01 rejection: a second trial that allowed a bottom-boundary auto row when every text-bearing cell's
  first line baseline fit above the margin matched the observation that Office can show only a row's top
  fragment at the page bottom, but it still was not acceptable. It improved the public probe's page 3
  (`MAE 3.494976 -> 1.374215`) while worsening page 2 (`9.349741 -> 10.047117`), worsened
  `docx-ladder-03-table-bottom-slack` page 1 (`6.651415 -> 6.849153`), and regressed private acceptance
  from the `8.927676` baseline to `9.160214` average MAE with page 15 worse (`12.133736`). Reverted. Treat
  visible top-fragment placement, continuation clipping, and next-page carry-over as one model problem rather
  than a permissive one-line boundary rule.
  2026-06-01 tooling follow-up: DOCX layout snapshots now expose per-row `BottomOverflowPoints`,
  `FirstBaselineY`, and `LastBaselineY`, aggregated from private-safe table-cell layout. This makes future
  public/private row-fragment comparisons explicit: a row can now be classified by whether its band crosses
  the body bottom, whether only the first line box fits, and whether lower line boxes require clipping or
  continuation.
  2026-06-01 tooling follow-up: layout page snapshots now also expose page-local margins, and
  `tools/SummarizeDocxRowBoundary.ps1` compares bottom-window layout rows against Office/candidate PDF text rows
  using hashes and lengths only. On public `docx-ladder-03-table-heading-table-keepnext` run `20260601-211327`,
  the reference has two bottom text rows on page 2 that are absent from the candidate page-2 hash set, while
  candidate layout has already pushed the next table row to page 3. On private accepted run `20260601-205142`,
  pages 14..16 show the same boundary class around the worst page without exposing text. This confirms the next
  implementation needs a row-fragment/page-boundary model with text-row carry-over, not another fit predicate.
  2026-06-01 architecture progress: `DocxTableRowLayout` and `DocxTableRowSnapshot` now carry explicit
  `FragmentIndex`/`FragmentCount`, and renderer row adjacency treats same-row continuation fragments as table
  adjacency. Current emitted rows remain single-fragment (`0/1`), but the layout/render/snapshot contracts now
  have a structural home for first/continuation fragments before pagination behavior changes. Keep this open
  until actual split rows carry clipped cell geometry, continuation text lines, and Office-observed border
  suppression.
  2026-06-01 rejection: a text-only auto-row split trial used measured row line boxes to place a bottom-page
  top fragment and, when needed, a continuation fragment. It improved both public boundary probes
  (`docx-ladder-03-table-heading-table-keepnext` page MAE `3.67/7.75/1.37`, and
  `docx-ladder-03-table-bottom-slack` page MAE `6.53/2.89`), but private acceptance regressed from the
  `8.927676` baseline to `9.135465`, with pages 14/15 worse. Inspection showed the regressing private rows were
  text-bearing top fragments with empty continuations. Reverted behavior. The next attempt needs a structural
  discriminator for when Word keeps such bottom text fragments versus moving the row, not just a line-box fit
  rule.
  2026-06-01 rejection: a refined continuation-height trial kept the same text-only row-fragment shape but
  advanced continuation rows by the remaining line boxes plus bottom padding/border instead of a full normal
  row body. This improved the public structural probes further: `docx-ladder-03-table-bottom-slack` reached
  page MAE `6.534271/2.019980`, and `docx-ladder-03-table-heading-table-keepnext` reached
  `3.670558/4.095706/1.374215`; the row-boundary hash summary confirmed candidate page 1 now includes the
  same bottom split-row hash as Office. It still regressed private DOCX acceptance from the accepted
  `8.927676` baseline to `9.134550`, with pages 14/15 worse. Private-safe structure inspection showed only two
  split rows were introduced, so the issue is not random over-application but an unresolved structural
  discriminator: public probes and regressing private rows both have paragraph spacing tokens inside split
  cells. Reverted production behavior and do not reattempt row splitting until the pre-layout/table-structure
  model can express the missing Word decision boundary.
  2026-06-01 follow-up: added public `docx-ladder-03-table-row-fragment-threshold` as an open diagnostic probe
  with repeated compact fixed tables, a wrapped shaded target row, and a `w:cantSplit` contrast. After removing
  a confounding page-top heading before-spacing signal, Office-backed run `20260601-222516` renders seven
  reference pages against six candidate pages, so the probe intentionally disables page/dimension gates until
  the feature is implemented. It demonstrates that the missing model is not just visible row placement on a
  crowded page: Office creates short continuation pages for the overflowing table sequences, while the current
  candidate keeps rows atomic and collapses one continuation page away. A trial that suppressed body paragraph
  before-spacing at page top made this probe worse and badly regressed `docx-ladder-03-table-bottom-slack`
  (`MAE 6.65/3.55 -> 12.64/3.15`), so it was reverted. Keep the case as the public row-fragment/page-count
  oracle. The next implementation must create first/continuation row fragments from row line boxes and
  `w:cantSplit`, emit continuation pages even when the visible carry-over is small, and avoid the previously
  rejected private regression where text-bearing top fragments were paired with empty continuations.
  2026-06-01 progress: the first accepted fix from that probe is the page-break paragraph line-box model, not
  row fragmentation itself. Public PDF inspection of the revised probe showed the missing page came from a
  standalone run page-break paragraph after a low after-table marker: Office consumes the invisible paragraph
  line, creates a blank page when that line does not fit, then applies the break. `DocxPageBreakElement` now
  carries the resolved break paragraph for run-break-only paragraphs, and layout consumes that line box before
  finishing the page. The tightened `docx-ladder-03-table-row-fragment-threshold` manifest now requires page
  and dimension parity; run `20260601-223127` passes with seven pages, including near-zero blank-page matches
  on pages 3, 5, and 7. `DocxInspect` block traces now expose whether a page break consumed a paragraph line
  and the break paragraph's line-spacing tokens. Existing row-boundary probes stayed neutral
  (`docx-ladder-03-table-bottom-slack` `6.651415/3.551346`,
  `docx-ladder-03-table-heading-table-keepnext` `4.160735/9.349741/3.494976`), `docx-tables --skip-slow`
  passed `79`, `docx-page --skip-slow` passed `29`, and private DOCX acceptance run `20260601-223133` stayed
  at `16/16` pages, zero dimension mismatches, no diagnostics, `MAE=8.927676`.
  2026-06-02 tooling follow-up: `tools/InspectDocx.ps1` now emits the model-level DOCX text-emission snapshot
  alongside layout, structure, font-plan, source-block, and page summaries. The new
  `text-emission-snapshot.json`, `text-emission-summary.json`, and `source-block-summary.json` files expose
  page/source-block line counts, terminal-space emissions, PDF `Tc` use, font resource IDs, and positioned
  spacing state without decoded document text. Keep this as the default bottom-up diagnostic path before adding
  more pagination or text-state rules, so future changes can be checked against the renderer's own structural
  model instead of loose PDF-text scripts.
  2026-06-02 architecture follow-up: layout `SourceBlocks` now classify their private-safe block kind
  (`Paragraph`, `Table`, `InlineImage`, `Mixed`, or `Unknown`) in addition to page span, text-line counts, row
  counts, and consumed height. This removes another implicit join between layout and structure snapshots when
  diagnosing page-flow residuals, and keeps the bottom-up DOCX model closer to the PPTX scene/snapshot pattern.
  2026-06-02 architecture follow-up: body-level inline image layout items now carry their owning
  `SourceBlockIndex`, and the layout source-block summary can classify true image-only paragraphs as
  `InlineImage` while text/image paragraphs remain `Mixed`. Table-cell images remain owned through table-row
  and cell snapshots. This closes a diagnostic ownership gap for image-bearing DOCX paragraphs without changing
  rendering behavior.
  2026-06-02 tooling follow-up: layout source-block summaries now include private-safe vertical bounds
  (`VerticalTop`/`VerticalBottom`) in addition to consumed height. This lets row-boundary and paragraph-flow
  diagnostics locate each block's emitted page span directly from `source-block-summary.json`, instead of
  re-walking every layout item in downstream scripts.
- [x] 2026-05-31: Investigate private slide 42 as a high-priority PPTX schema/text-layout issue. On the left
  schema, Office places the numbers centered inside their rectangles, while the candidate places the numbers
  incorrectly and emits the wrong color. Treat this as a generic shape/text-frame alignment and inherited text
  color problem, not as a private-content coordinate tweak: inspect the Office/candidate PDF text operations,
  rectangle geometry, fill/stroke/text color states, and OOXML style inheritance, then reproduce the underlying
  behavior with public synthetic fixtures before changing production rendering.
  2026-05-31 progress: PDF and OOXML inspection showed the first reproducible structural gap is in chart
  data-label overrides, not ordinary slide text frames: point-level `c:dLbl/c:delete` was not preserved in the
  scene model, so deleted chart labels could still be rendered. `PptxSceneChartDataLabelOverride` now carries
  the raw delete state, the chart renderer suppresses deleted point labels, and `pptx-charts` has a public
  synthetic guard. Private run `20260531-172242` improved slide 42 slightly (`2.245136960` MAE,
  `0.045451389` changed16, `0.954406876` SSIM), but the item remains open: remaining numeric chart labels
  still differ in baseline/position and likely need Office-aligned bar/stacked data-label box geometry and
  chart-label text color inheritance, not another private coordinate tweak.
  2026-05-31 progress: the remaining right-schema numeric labels were traced to stacked-column data-label
  geometry. The renderer was using clustered series slots for data labels even when the bars themselves used
  cumulative stacked segment geometry. `RenderBarDataLabels` now receives the bar grouping and gap width,
  computes positive/negative stacked segment start/end values with the same normalization as bar rendering,
  and defaults missing stacked-bar label positions to centered segment labels instead of outside-end labels.
  A public `pptx-charts` synthetic guard now checks that stacked-column labels share the stacked column rather
  than spreading into clustered slots. Private run `20260531-183004` kept slide 42 in the same aggregate metric
  band (`2.254341242` MAE, `0.045626929` changed16, `0.954226790` SSIM), but PDF text inspection shows the
  right-schema numeric labels moved from clustered-slot positions to within about 2-3 pt of Office positions
  and their fill color states now match Office's white/dark label sequence. Remaining residual is lower-level
  chart text emission structure: Office emits these bold labels as fill+stroke text, while the candidate uses
  normal filled text when a bold font is resolved. Keep that residual with the shared Office PDF text-emission
  track instead of adding slide-specific chart offsets.
  2026-05-31 progress: the latest right-schema discrepancy was traced to automatic vertical value-axis tick
  density for a stacked column chart without an explicit `c:majorUnit`. Office emits the dense 0..50 axis in
  five-unit steps, while the renderer's generic nine-target nice tick rule selected ten-unit steps and made
  the numeric scaffold look misplaced. `GetValueAxisAutoTickTargetCount` now uses a named vertical-value-axis
  Office default while preserving the existing horizontal/manual-layout policy, and `pptx-charts` has a public
  synthetic guard for the 0,5,10,...,50 sequence. Private run `20260531-191225` improved slide 42 to
  `2.241075907` MAE, `0.045509259` changed16, and `0.954718719` SSIM. The specific slide-42 chart-number
  placement/color concern is now closed; any remaining glyph-structure differences should stay with the shared
  Office PDF text-emission track, not with slide-local chart offsets.
- [ ] 2026-05-31: Continue the private page-36 typography branch as an Office text-emission model problem,
  not as a font-family rule. Private run `20260531-002604` has page 36 as the worst slide (`6.045371817`
  MAE) and the page is now narrowed to text-state decomposition: candidate and Office text positions are close
  enough for `ComparePptxTextEmission.ps1`, but Office emits a mixture of `Tc`, secondary `/Tf`, and text
  operation splits that the candidate still mostly leaves on the main font-size grid. The current page-36
  comparison has `135` Office text operations versus `133` candidate glyph runs, with branch counts
  `main-grid=111`, `secondary-0.024=18`, `secondary-2.04=1`, `secondary-2.064=1`,
  `secondary--2.04=2`, and `2` unmatched reference operations. The dominant private three-column frame uses
  `noAutofit`, `wrap=square`, `numCol=3`, `spcCol=108000`, default `7.2pt/3.6pt` insets, and `12pt` text.
  Office emits mostly negative character spacing there (`-0.036`, `-0.0476`, `-0.0574`, plus smaller
  `-0.0287`/`-0.012` buckets), while the candidate currently emits `Tc=0` for that frame.

  Rejected shortcuts from the 2026-05-31 public/private sweep: do not special-case `Cambria Math`, MATH-table
  fonts, or a column-only condition. The same `secondary-0.024` branch appears in public synthetic probes and
  private page 36 with both `Cambria Math` and `Aptos`; it appears in three-column `noAutofit` frames and also
  in one-column `noAutofit`/default frames. The public `pptx-ladder-04-typography-unspaced-column-tc-probe`
  repeats the three-column structure and shows `secondary-0.024=12`, but it also shows Office using zero,
  positive, and negative `Tc` within the same frame. The durable target is therefore a PDF-emission profile
  that models Office's line/paragraph/frame text-state decomposition after layout, including secondary `/Tf`
  selection and compensated `TJ`, without altering layout widths or using typeface names as conditions.
  Rejected trial: tightening overwide first-segment chunk fit tolerance for every `noAutofit` frame improved
  the public unspaced-column probe dramatically (`10.423376 -> 1.394638` MAE, SSIM `0.3738 -> 0.9597`) but
  regressed private page 36 (`6.045372 -> 7.735627` MAE) and deck MAE (`2.933659 -> 2.976731`) in run
  `20260531-004108`. Preserve this as evidence that overwide-run wrapping is part of the problem, but do not
  land a global no-autofit tolerance change. The next attempt needs a narrower structural discriminator, likely
  tied to the exact Office line-building state for unspaced paragraphs/columns rather than the autofit enum
  alone.
  Follow-up, 2026-05-31: extended `SummarizePptxTextStateDeltas.ps1` with reliable-position placement buckets
  so page-36 text-state work does not accidentally turn into a baseline nudge. Rechecking run `20260531-014545`
  shows the dominant three-column explicit `noAutofit` frame is already vertically aligned: `41` reliable
  matches average `DeltaBaselineY=+0.016632pt` with range `-0.020441..+0.055559pt`. The larger baseline
  residuals are in separate one-column frames (`18pt` default-autofit around `-0.326pt`, and `14.04pt`
  noAutofit numbered/default frames around `-0.17..-0.91pt`). Therefore do not pursue a page-36 global
  baseline constant, `0.974` fallback change, column baseline offset, or font-family baseline rule. The
  rendering-impact branch remains Office PDF text-state decomposition: `Tc`, secondary `/Tf`, and operation
  splitting, with public probes needed before changing production behavior.
- [ ] 2026-05-30: Generalize the new emission-only `Tc` hook beyond highlighted `spAutoFit`. Private page 81 now
  shows the real target shape of the problem without exposing private text: Office and candidate both emit
  `83` text operations, but Office buckets `Tc` as `0:30`, `-0.048:20`, `0.0173:13`, `-0.0535:11`, and
  `-0.024:9` while the candidate remains `0:83`. The nonzero buckets cluster by text frame and line group,
  not by highlight, and include both `spAutoFit` and `noAutofit` frames. Private page 79 remains table-heavy
  and still shows Office `Tc` buckets across the table while the candidate is `0:249`. The next implementation
  should derive a reusable PDF-state split from positioned glyph residuals, frame/line context, and Office-like
  font-grid branch selection, then lock that behavior with public synthetic frames/tables before accepting it
  against the private deck.
  2026-05-30 follow-up: added private-safe glyph category counts to `PptxInspect` and threaded them through
  `ComparePptxTextEmission.ps1`. Page 81's nonzero `Tc` buckets do not collapse to a glyph-class rule:
  category shapes are mixed inside each bucket, while the strongest signal remains frame/paragraph/line
  structure. A disposable public probe with generic text, `spAutoFit`, matching frame geometry, and numbered
  paragraphs reproduced Office `Tc` buckets `-0.048`, `-0.024`, and `0` without using Cambria Math or private
  content. A separate public `startAt` sweep rejected the simpler numbering-start ladder: four numbered
  `spAutoFit` frames with starts 1-4 all emitted `Tc=-0.048`. The next renderer slice must therefore model
  Office's text-body/glyph-emission decomposition, not a font-name, page, or numbering-value shortcut.
  The public reproduction is now tracked as
  `pptx-ladder-04-typography-spautofit-numbered-tc-probe`: Office emits `Tc` buckets `0:7`, `-0.048:12`,
  and `-0.024:16`, while the current candidate still emits `Tc=0:35` and has raster metrics MAE `5.162535`,
  changed16 `0.058189`, SSIM `0.648233`.
  2026-05-30 renderer slice: the candidate now carries paragraph bullet kind and auto-number metadata from
  the text model into positioned spans, glyph-run snapshots, `PptxInspect`, and
  `ComparePptxTextEmission.ps1`. The emission layer uses that structural metadata, plus `spAutoFit` frame
  state and zero OOXML layout tracking, to reproduce the public numbered-autofit probe's Office `Tc` buckets:
  candidate PDF text state moved from `0:35` to `0:7`, `-0.048:12`, and `-0.024:16`. The implementation keeps
  layout character spacing at zero and compensates `Tc` inside `TJ`, so this is still an Office-like PDF
  decomposition slice rather than a raster nudge. Private page 81 moved from candidate `Tc=0:83` to
  `0:41`, `-0.048:20`, and `-0.024:22`; that closes the whole `-0.048` frame bucket but deliberately leaves
  the item open because Office still has an uncaptured non-numbered `-0.0535:11` bucket and a later frame-10
  split where Office changes from `-0.024:9` to `+0.0173:13`. Full private run `20260530-195414` compared all
  `84/84` pages with empty diagnostics and unchanged raster metrics (`2.947966` deck MAE, changed16
  `0.053572`), as expected for an emission-compensated text-state change. Next work should isolate the
  remaining frame-10 line/paragraph split and the non-numbered `spAutoFit` bucket with additional public
  probes before touching page 79's table-heavy `Tc` branch.
  2026-05-31 counterexample: added
  `pptx-ladder-04-typography-spautofit-numbered-synthetic-bold-tc-probe` to isolate a tempting page-81
  hypothesis: a numbered `spAutoFit` frame with a mid-line synthetic-bold continuation. Office still emits
  `Tc=-0.024:17` throughout that public frame while using secondary `12.024pt` font-size branches on `3`
  operations, so do not promote the page-81 `+0.0173` bucket as a generic synthetic-emphasis continuation
  rule. The remaining target still needs a broader PDF text-state decomposition model, not a font-name,
  bold-run, or local line-position shortcut.
  2026-05-31 non-numbered bucket probe: disposable public probes matching a wide zero-inset `noAutofit`
  frame at `20pt`, both with default and `90%` line spacing, emitted Office `Tc=0` rather than the private
  `-0.0535` bucket. This rules out the simple noAutofit/zero-inset/line-spacing explanation; keep the
  production renderer unchanged until a public fixture reproduces the bucket through additional inherited
  placeholder/style or text-state structure.
  2026-05-30 continuation: added
  `pptx-ladder-04-typography-spautofit-numbered-run-split-tc-probe` after private page 81 showed that a
  continuation paragraph can use the dense `-0.002em` bucket even when the public frame has only `14` visible
  lines. This purged the previous visible-line-count guard, which was a magic threshold, and replaced it with
  the structural condition actually supported by public evidence so far: explicit auto-number start followed
  by body-continuation paragraphs. The public run-split probe's Office reference emits `Tc=-0.024` for the
  whole frame; the candidate now emits the same bucket while keeping layout spacing at zero. This still does
  not explain private page 81's later `+0.0173` suffix, so the next public probe must vary paragraph/run
  segmentation, Office font-size grid branch, and line-local operation splitting rather than introducing any
  typeface discriminator. Validation: `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed,
  `pptx-typography --skip-slow` passed (`131` passed, `2` skipped), the new visual case ran at
  `20260530-200506`, and private `lokad-value-based` run `20260530-200639` compared all `84/84` pages with
  empty diagnostics and unchanged deck metrics (`2.947966` MAE, changed16 `0.053572`).
  2026-05-30 source-run evidence: extended glyph-run snapshots, `PptxInspect`, and
  `ComparePptxTextEmission.ps1` with private-safe source-run indices. Page 81 frame 10 now shows that the
  Office `+0.0173` bucket starts inside paragraph 2, line 1, before the OOXML run boundary, then continues
  through later source runs and the following paragraph. A disposable public variant that changed the
  run-split probe to Cambria Math still emitted only `Tc=-0.024`, so the remaining branch should not be
  modeled as a font-family rule. The next reproduction attempt should target Office's line-internal text
  operation segmentation around punctuation/short spans and font-size-grid branch changes.
  2026-05-30 non-numbered branch probe: page 81's `-0.0535` bucket spans several non-numbered frames, including
  one short `spAutoFit` frame with a manual line break in the second paragraph. Existing public same-geometry
  non-numbered `spAutoFit` frames emit `Tc=0`, and a disposable public variant with the same geometry plus a
  manual break emitted only a post-break `Tc=-0.024` tail. That is a real public gap but not the private
  `-0.0535` rule. Continue probing manual-break/short-frame/line-internal segmentation combinations before
  adding any non-numbered autofit emission rule.
  2026-05-31 fractional-size probe: a disposable public no-autofit textbox sweep at `20pt`, `20.04pt`,
  `20.06pt`, `14.04pt`, `15.96pt`, and `12pt` emitted Office `Tc=0` for all operations. This rules out the
  simple fractional-authored-font-size explanation for page 81's non-numbered `-0.0535` bucket; keep looking
  for inherited placeholder/style, bodyPr, or operation-segmentation structure before changing production.
  2026-05-30 page-36 evidence: the current worst private page is a no-table text-state case rather than a
  table-row case. Office emits `135` page-36 text operations with widespread nonzero `Tc` buckets
  (`-0.036`, `-0.0476`, `-0.012`, `-0.0574`, `0.0166`, and related splits) and secondary font-size branches
  (`12.024`, `14.064`, `18.024`), while the candidate initially emitted `133` operations with `Tc=0` for all
  operations. A trial that generalized average glyph-residual promotion beyond tables produced some matching
  nonzero `Tc` buckets for no-autofit text, but many Office buckets had no corresponding glyph residual in the
  current layout and the private deck slightly regressed (`2.947933 -> 2.948011` MAE, page 36
  `6.045372 -> 6.046293`), so the trial was reverted. Do not retry a blanket non-table residual promotion;
  the next page-36 slice needs a first-class Office PDF text-state decomposition model that can account for
  font-grid branches and operation splitting even when layout residuals are zero.
- [ ] 2026-05-30: Continue page-79 table/text work from structural table text evidence, not a new private
  coordinate rule. Page 79 is table-heavy (`42` table text frames, no effects/pictures/transparency), and
  inspection of run `20260530-160928` showed Office/candidate graphics are broadly the same table/grid class
  while text remains divergent: Office emits `258` text operations against candidate `249`, mostly `11.04pt`
  table text, with Office-only `Tc` families (`0.00384`, `0.0509`, `-0.0216`, `-0.116`, and related small
  clusters) while candidate emits `Tc=0`. Text comparison shows most runs are sub-point close, but a few
  wrapped cells have large row-internal line-placement mismatches. The next acceptable page-79 push needs a
  public Office-authored table ladder that isolates wrapped middle-anchored cell line placement and text-state
  emission; do not add a private row/column coordinate shortcut.
  Follow-up, 2026-05-30: the existing public
  `pptx-ladder-10-table-center-explicit-wrapped` case does isolate part of the table-row slack family. Its
  current baseline run `20260530-171826` has MAE `8.201281`; preserving declared row heights instead of
  scaling large row-height slack to the full graphicFrame improved that public case to MAE `3.713218` in
  run `20260530-172034` and kept non-slow `pptx-tables` passing (`15` passed). The same broad rule is still
  rejected for the private deck: run `20260530-172055` worsened deck MAE `2.949517 -> 2.958489`, made page 21
  worst (`5.943711 -> 6.697341`), and left page 79 unchanged at report precision. Conclusion: the future table
  row-height rule needs a discriminator between the public high-slack wrapped-center table and private page 21's
  accepted content-minimum row expansion; do not retry a blanket large-slack declared-row preservation rule.
- [ ] 2026-05-30: Treat page-79's `258` vs `249` text-operation delta as a PDF text-operation decomposition
  gap until proven otherwise. The combined private-deck compare initially overmatched across pages because it
  used only position; `ComparePptxTextEmission.ps1` now records reference page and candidate slide numbers and
  constrains position/text matching to equal page/slide when those fields exist. With page-constrained
  matching, pages 36/79/81 retain their prior per-page counts (`36: 130 delta, 2 missing, 3 ok`; `79:
  249 delta, 9 missing`; `81: 44 delta, 39 ok`). The page-79 "missing" rows are clustered in table-cell text
  bands where Office emits extra same-line text operations with distinct `Tc` buckets; they should drive a
  structural public probe for Office's text-state splitting rather than a content-loss or coordinate patch.
  Follow-up, 2026-05-30: fresh run `20260530-184739` keeps this diagnosis after the raster-effect opacity
  split. Slide 79 has one `6 x 7` table, no pictures, no effects, no transparency, and ordinary left/right
  table paragraph alignment; it is not a graphics, justification, or distributed-text problem. Office emits
  `258` text operations and candidate emits `249`; candidate table glyph spans group into rows with `8`,
  `39`, `39`, `77`, `36`, and `45` operations, while Office's nine extra operations remain inside the same
  table bands. The slide XML has many multi-run cells at the same declared size/typeface, so the next public
  probe should isolate same-style run-boundary preservation and text-state decomposition inside table cells,
  including `Tc` extraction from residual glyph positioning. Do not fix this by keying on the observed private
  font name or by adding row/column coordinate rules.
  Rejected trial, 2026-05-30: broadening text-operation boundary punctuation from dash punctuation to quote
  punctuation was supported by one public run-boundary PDF count (`5 -> 7` candidate text operations against
  Office's `6`), but it regressed the private deck (`2.947966 -> 2.999904` MAE) and page 79 only moved
  structurally (`249 -> 256` candidate text operations) while still emitting `Tc=0`. The split changed layout
  behavior before PDF emission, so the next attempt must move operation decomposition into the PDF emission
  layer or into a layout-preserving glyph-span split, not into word wrapping/line layout segmentation.
  Rejected trial, 2026-05-30: an emission-only quote-boundary split avoided the large layout regression and
  made the public run-boundary fixture match Office's operation count (`6/6`), but it still regressed the
  private deck slightly (`2.947966 -> 2.948299` MAE). Page 79 moved only partway toward Office
  (`249 -> 252` candidate text operations versus Office `258`), while page 81 moved away from the already
  matching text-operation count (`83 -> 91` candidate operations versus Office `83`). This shows quote
  boundary decomposition is entangled with Office's `Tc`/secondary-`Tf` text-state decomposition; standalone
  punctuation splitting over-decomposes cases that Office represents with character spacing. Keep future
  attempts layout-preserving, but derive operation boundaries and `Tc` together from the positioned glyph
  residuals rather than from punctuation categories alone.
  Follow-up, 2026-05-31: refreshed page-79 with reliable-position placement buckets after the table clip and
  following-space-flow fixes. The current exact comparison has `308` Office rows, `202` matched rows, `106`
  missing rows, and `192` reliable position matches. Reliable table baseline deltas are stable within cells
  and vary by row/cell rather than by a single page coordinate: row `3` averages about `-0.554pt`, row `2`
  about `-0.319pt`, row `1` about `-0.328pt`, row `5` about `-0.233pt`, row `4` about `-0.200pt`, and row
  `0` about `-0.199pt`. The largest missing Office text-state buckets are still `11.04pt` table text with
  nonzero `Tc` (`0.0509`, `0.00384`, `-0.0216`, `-0.0379`, `-0.116`) plus unmatched rows where no candidate
  cell match is reliable. The page-79 table has explicit `0.62748pt` top/bottom insets and declared row heights
  equal to rendered row heights, so the existing default-inset magic constants are not active. Do not add
  row/column offsets or a flat middle-anchor adjustment; the next acceptable renderer slice needs a public
  table-cell fixture that separates vertical text-height/anchor computation from PDF text-state splitting.
  Follow-up, 2026-05-31: added public Office-authored
  `pptx-ladder-10-table-middle-small-insets` as the no-slack/small-inset/middle-anchor table counterexample
  needed before touching private page 79. The fixture uses a `6 x 7` table with page-79-like declared row and
  column dimensions, explicit tiny margins, middle anchoring, neutral wrapped text, and some multi-run cells.
  It passes as a discovery visual case (run `20260531-023220`, MAE `0.966760`, changed16 `0.019591`, empty
  diagnostics). PDF/text inspection shows Office and candidate both emit `79` text operations, all Office
  `Tc=0`, with reliable baseline deltas only about `+0.01..+0.28pt`. This preserves valuable negative evidence:
  page 79's nonzero `Tc` families and larger negative baseline residuals are not caused by no-slack table
  geometry, explicit tiny margins, middle anchoring, or generic wrapped multi-run cells alone.
  Follow-up, 2026-05-31: added public Office-authored
  `pptx-ladder-10-table-font-fragmentation` as the matching positive table-text-state probe. It keeps the same
  page-79-like no-slack geometry and explicit tiny middle-anchored insets as the zero-`Tc` counterexample, but
  changes only public structural text inputs: resolved font profile, heavier run fragmentation, longer wrapped
  neutral cells, and right-aligned trailing value columns. Office emits `125` text operations with `90` matched
  nonzero `Tc` operations (`+0.0499`, `+0.102`, `-0.0115`, `-0.0389` families at `11.04/11.064pt`) while the
  candidate still emits `Tc=0` for all matched operations. This is the first public positive close enough to
  page 79 to be useful: it proves that the private table branch is not private-content-specific and that the
  next renderer slice must align Office's PDF text-state decomposition from structural font/run/line context.
  Do not implement this as a font-name branch; the fixture's font is evidence of a metric/emission profile, not
  a predicate.
  Rejected trial, 2026-05-31: lowering the typographic-ascender threshold for fonts with unusable Windows
  ascenders from `0.93` to `0.75` had no visual effect on either table probe (`pptx-ladder-10-table-font-fragmentation`
  stayed at MAE `2.029893`; `pptx-ladder-10-table-middle-small-insets` stayed at MAE `0.966760`) and
  `pptx-typography --skip-slow` still passed. The middle-anchored table path is not using that baseline-floor
  branch, so do not pursue this as the page-79 fix.
  Completed slice, 2026-05-31: fixed the active middle-table anchor estimate instead. For table-cell vertical
  anchors, `EstimateTextHeight` used typographic font-box line advance for default line spacing while the actual
  laid-out lines still advanced on the normal Office line grid. This under-estimated content height for fonts
  whose Windows ascender is unusably oversized and whose typographic line box is compressed, then over-centered
  wrapped middle-anchored table cells. `ReadEstimatedAnchorLineAdvance` now clamps that table-cell estimate to
  the normal line advance only when the resolved font's Windows ascender exceeds the Office baseline limit and
  the typographic line-box ratio is at most `1.18`. This is metric-driven, not a font-name/page/row rule.
  Public positive `pptx-ladder-10-table-font-fragmentation` improved MAE `2.029893 -> 1.524499` and changed16
  `0.032217 -> 0.026186`; the zero-`Tc`/normal-font counterexample
  `pptx-ladder-10-table-middle-small-insets` stayed unchanged at MAE `0.966760`. Private
  `lokad-value-based` run `20260531-025144` improved deck MAE `2.931757 -> 2.877036`, changed16
  `0.053386 -> 0.052651`, page 79 MAE `4.895694 -> 3.38`, page 21 MAE `5.943039 -> 5.013651`, page 47 MAE
  `3.56 -> 2.77`, and page 6 MAE `3.13 -> 2.34`, with no measured page regressions and no diagnostics.
  Remaining page-79 work is still PDF text-state/operation decomposition: refreshed page-79 text comparison has
  `258` Office operations, `252` matched operations, `6` missing operations, `138` reliable-position matches,
  and `240` matched nonzero Office `Tc` operations while candidate `Tc` remains zero.
  Follow-up, 2026-05-30: extended `PptxInspect` so table text paragraph snapshots are emitted separately as
  `table-text-paragraph-models.json`, avoiding ad hoc private XML parsing when investigating table text.
  Re-inspection of page 79 shows the table paragraphs resolve through the normal paragraph/style cascade
  (`paragraph.pPr`, `shape.lstStyle`, inherited text style, default text style), with ordinary left/right
  alignment and zero OOXML run `spc`; the private `Tc` buckets are therefore not a hidden distributed/justify
  alignment rule and not a typeface discriminator. Public table fixtures already expose the same Office PDF
  behavior: `pptx-table` has Office table-cell `Tc` buckets `0.06`, `0.036`, `-0.024`, and `-0.048` against
  candidate `0`, and `pptx-ladder-10-composite-table-port` has Calibri table-cell buckets `0.0551` and
  `-0.0383` against candidate `0`. Keep the next renderer slice public and structural: isolate how Office
  derives table-cell PDF `Tc` from resolved text metrics and `TJ` residuals, then emit an equivalent
  layout-preserving PDF decomposition. Do not key on page 79's private font family, row/column coordinates, or
  a fixed bucket table.
  Follow-up, 2026-05-30: added PDF-inspection fields for `TextChunkCount`, raw `TJ` adjustment counts/ranges,
  average adjustment in points, and net average character spacing. This made the page-79 structural gap
  observable without private text: Office usually emits nonzero `Tc` with one text chunk, while the candidate
  represented many table-cell residuals as `Tc=0` plus uniform `TJ` adjustments. The renderer now promotes
  uniform table-cell PDF residuals into `Tc` at emission time and compensates the `TJ` array, preserving glyph
  positions by construction. This is structural, font-agnostic, and content-agnostic; it is not keyed to
  page 79, row/column coordinates, or a fixed Office bucket. The same pass exposed and fixed a PDF text-state
  bug: `Tc` must be emitted for every text operation, including `0 Tc`, because it is inherited PDF state.
  Public `pptx-ladder-10-composite-table-port` now shows promoted body rows as nonzero `Tc` with zero residual
  `TJ`; `pptx-ladder-10-basic-table` remains at its pre-existing narrow visual gate miss (`0.0508098` vs
  `0.05`) because its affected candidate rows had no residual to promote. Private run `20260530-204050`
  compared all `84/84` pages with empty diagnostics and stable deck metrics (`MAE=2.947966`, changed16
  `0.053572`); page 79 moved slightly (`5.014204 -> 5.014172` MAE) while structurally gaining nonzero `Tc`
  buckets such as `-0.099`, `-0.025`, and `-0.022` for promoted table operations. Remaining page-79 debt is
  now the non-uniform residual branch and the Office-specific extra text-operation splits, not all-zero table
  `Tc`.
  Follow-up, 2026-05-30: broadened the table-cell residual decomposition from uniform-only to average residual
  promotion. Instead of requiring every interglyph residual to be identical, the renderer now emits the mean
  table-cell residual as PDF `Tc` and leaves only deviations in `TJ`, preserving glyph positions by construction.
  This stays content-agnostic and font-agnostic: no private coordinates, no row/column buckets, and no font-family
  discriminator. Focused validation passed (`pptx-typography --skip-slow`: `132` passed, `2` skipped;
  `pptx-tables --skip-slow`: `17` passed). Public `pptx-ladder-10-composite-table-port` passed at run
  `20260530-205145`; public `pptx-ladder-10-basic-table` remains at the pre-existing narrow gate miss
  (`0.0508098234953704` MAE vs `0.05`). Private `lokad-value-based` run `20260530-204915` compared all `84/84`
  pages with empty diagnostics and improved deck metrics slightly (`MAE=2.947933`, changed16 `0.053571`).
  Page 79 improved (`5.014172 -> 5.012818` MAE), page 21 also improved (`5.943075 -> 5.939469` MAE), and
  PDF inspection of page 79 now shows residual-heavy table branches as nonzero `Tc` with near-zero average `TJ`
  residuals. Remaining page-79 work is the `258` vs `249` Office text-operation split and the line-placement
  clusters, not the existence of table-cell text-state extraction.
  Follow-up, 2026-05-31: accepted a table text-flow correction from refreshed private page-79 evidence.
  Page 79's table cells still had large line-placement collapses after the PDF `Tc` work because the table
  layout path grouped ordinary spaces with the preceding word unless `noAutofit` was explicitly authored.
  The private table cells use default table autofit, not explicit `noAutofit`, and Office's observed line
  counts are closer when word-flow treats the following word as owning the leading space. `BuildTextFlowFrame`
  now names this as `UsesOfficeFollowingSpaceFlow`: explicit `noAutofit` frames keep the existing rule, and
  table-cell frames use the same following-space flow even when their autofit mode is the default. This is a
  structural text-flow rule, not a page, coordinate, row/column, typeface, or `Tc` bucket special case. A
  rejected diagnostic branch remains: the same page still has Office-only nonzero `Tc` families and extra
  text-operation splits, so do not interpret this as solving the whole table text-state problem. Validation:
  `pptx-tables --skip-slow` passed (`19` passed), `pptx-typography --skip-slow` passed (`135` passed,
  `2` skipped), public `pptx-table` passed at `20260531-010216`,
  public `pptx-ladder-10-table-center-explicit-wrapped` passed at `20260531-010216`, public
  `pptx-ladder-10-table-center-explicit-multiline` passed at `20260531-010229`, and private
  `lokad-value-based` run `20260531-010104` compared all `84/84` pages with empty diagnostics. Deck MAE
  improved `2.933648 -> 2.931758`, changed16 improved `0.053397 -> 0.053386`, and page 79 improved
  `5.014204 -> 4.895851`.
  Follow-up, 2026-05-31: accepted the horizontal half of the same table-cell clipping structure from refreshed
  page-79 evidence. The table text-frame model already preserved full cell clip rectangles, but the shared
  column layout code replaced any local clip with the inset text column (`columnStartX/columnWidth`) even for
  table cells. Office's PDF clips for the page-79 table are aligned to full cell boundaries, while candidate
  table text clips were mostly inset by the authored cell margins. `LayoutTextFrame` now keeps per-column clips
  for ordinary local-clipped shape text, but lets table-cell text runs carry the full frame clip produced by the
  table text-frame model; wrapping and glyph placement remain inset-based. The public table regression
  `PptxSyntheticTableWrapsCellTextToColumnWidth` now checks both sides of that contract: multiple lines still
  start at the text inset, and the emitted `W*` rectangle is the full cell (`72 396 72 72`) rather than the
  inset text rectangle (`79.2 396 57.6 72`). Validation: non-slow `pptx-tables` passed (`19` passed), and
  private `lokad-value-based` run `20260531-014545` compared all `84/84` pages with empty diagnostics. Deck MAE
  moved `2.931758 -> 2.931757`; page 79 improved `4.895850694 -> 4.895694444`. This is intentionally a small
  visual change, but it removes another table-local PDF clip mismatch without private coordinates, font-family
  checks, or row/column buckets. Remaining page-79 residuals are still text-state decomposition, extra Office
  text-operation splits, and some clip nesting/count differences, not a solved table rendering branch.
  Follow-up, 2026-05-30: extended the table text-frame inspection path with row-height provenance so the
  private table pages can be compared without reconstructing OOXML geometry by hand. `PptxInspect` table-frame
  records now expose the declared row height, declared row-span height, declared total table height, and
  graphic-frame/declared-height slack factor. This exposed a useful private split after average-`Tc`:
  page 21 and the public `pptx-ladder-10-table-center-explicit-wrapped` case both have about `1.54x`
  frame/declared-row slack and the same authored row ladder, but page 21's rendered rows are content-minimum
  driven while the public case remains the rejected high-slack declared-row counterexample. Page 79 is a
  different family: declared and rendered row heights match at report precision, so its residual belongs to
  the PDF text-state decomposition branch rather than table row allocation. Validation: full solution build
  passed, `pptx-tables --skip-slow` passed (`17` passed), `PptxInspect` build passed, and private
  `lokad-value-based` run `20260530-210736` stayed identical to the rendering baseline (`MAE=2.947933`,
  changed16 `0.053571`, empty diagnostics). This is diagnostic scaffolding, not a rendering change.
  Follow-up, 2026-05-30: refreshed the public/private table-row split after run `20260530-220006`. The public
  `pptx-ladder-10-table-center-explicit-wrapped` case is now known to be a high-slack, explicit-margin,
  middle-anchored table whose current candidate cells are all single-line (`1` line in every row/column);
  private page 21 has the same slack factor and row ladder but many multi-line cells and larger content-minimum
  pressure. This is the missing discriminator behind the rejected broad declared-row preservation trial. The
  next acceptable row-height slice should first create a public Office probe that separates high-slack
  single-line explicit-margin tables from high-slack wrapped/multi-line explicit-margin tables, then apply
  the rule only if it predicts both the public counterexample and private page 21. Do not encode a text-length,
  row-index, or private coordinate proxy for this split.
  Follow-up, 2026-05-30: added public `pptx-ladder-10-table-center-explicit-multiline` as the required
  high-slack, explicit-margin, middle-anchored multiline probe. It is generated from the existing
  Office-authored wrapped table by changing only the cell text in `ppt/slides/slide1.xml`, preserving the same
  authored row-height ladder (`13.99/9.79/13.99/22.38/...pt`) and `1.54x` graphic-frame/declared-height
  slack. `PptxInspect` shows every candidate cell now has two layout lines, while the OOXML row declarations
  remain identical to `pptx-ladder-10-table-center-explicit-wrapped`. Public run `20260530-222314` passes the
  intentionally loose discovery gate with MAE `13.537929`, changed16 `0.129103`, and empty diagnostics.
  Office/candidate text operation counts match (`110/110`), but PDF text-position comparison exposes the
  structural row-allocation mismatch: Office preserves heterogeneous row placement under multiline pressure,
  while the current candidate's content-minimum path collapses the rendered rows to near-uniform `22.76pt`
  bands. This fixture is a public row-allocation oracle, not yet permission for another blanket
  declared-row-preservation rule; page 21 remains the acceptance check for any future discriminator.
  Rejected trial, 2026-05-30: applying Office's 600 DPI PDF font-size grid during horizontal table-cell
  layout looked structurally plausible because page 79 Office emits `11.04pt` text while OOXML stores `11pt`.
  The trial made table wrapping/layout consume the same rounded size that PDF emission already uses, but it
  regressed the private deck (`MAE=2.933667 -> 2.949566`, changed16 `0.053398 -> 0.053533`), worsened page 79
  (`5.012818 -> 5.720985` MAE) and page 48 (`4.431142 -> 4.883708` MAE), and broke the public
  `PptxSyntheticTablePromotesAveragePdfSpacingResidualToTextState` expectation by removing the residual the
  emission layer is supposed to decompose. Conclusion: table-cell PDF font-grid alignment must stay in
  emission/text-state decomposition unless a future Office-authored probe proves a narrower layout metric
  rule. Do not retry a blanket table layout font-size rounding rule.
  Superseding follow-up, 2026-05-31: public table evidence overturned the average table-`Tc` promotion branch.
  `pptx-ladder-10-rich-text-cell` Office reference emits rich table text as `Tc=0` while carrying positioning
  differences in `TJ`; the promoted candidate emitted nonzero `Tc` buckets for the same public table spans.
  Private page 21 showed the same structural mismatch: the promoted candidate had table `Tc` buckets around
  `+0.018pt` and `-0.052pt`, while Office table-body text is mostly `9.96pt/Tc=0` with only a smaller
  secondary `Tc=-0.00888` family. The renderer therefore removed `PromoteAverageTablePdfCharacterSpacing`;
  residuals remain in the glyph positioning array and table text no longer invents a PDF text-state bucket.
  Public regression `PptxSyntheticTableKeepsAveragePdfSpacingResidualInPositioningArray` replaces the old
  heuristic-locking unit. Validation: solution build passed; `pptx-tables --skip-slow` passed (`18` passed);
  public `pptx-ladder-10-rich-text-cell` passed in run `20260531-001800`; the full `pptx-tables` visual
  family still has the same three pre-existing narrow gate failures (`basic-table`, `border-alpha`, and
  `explicit-borders`). Private `lokad-value-based` run `20260531-001809` compared all `84/84` pages with empty
  diagnostics; deck MAE was effectively neutral/slightly better (`2.933667 -> 2.933659`) while page 21 worsened
  slightly (`5.939469 -> 5.943068`). Keep this change because it removes a renderer-local heuristic that public
  Office PDFs contradict, but keep the page-21/page-79 table text-state item open: Office's small nonzero table
  `Tc` families and extra text-operation splits still need a public structural rule rather than a resurrected
  average-residual promotion.
  Follow-up, 2026-05-31: added `tools/SummarizePptxTextStateDeltas.ps1` to compare Office `Tc` branches
  against candidate glyph-run structure without emitting private text. A current sweep over private page 36,
  private page 79, public `pptx-table`, public `pptx-ladder-10-composite-table-port`, and public
  `pptx-ladder-10-rich-text-cell` confirms the constraint that should govern the next implementation. Page 36
  is mostly a non-table text-state branch (`131/133` matched text operations have nonzero Office `Tc`), page 79
  is mostly table text-state (`240/252` matched operations have nonzero Office `Tc`), public simple/composite
  tables also require nonzero table `Tc`, but the public rich-text table is a hard zero-`Tc` counterexample
  even when a candidate span carries residual `TJ` adjustments. Therefore do not resurrect average residual
  promotion, do not key on table-ness alone, and do not key on private font family. The next acceptable
  renderer change must explain both the simple/composite table positives and rich-text zero-`Tc` counterexample,
  while also covering the non-table page-36 branch.
  Follow-up, 2026-05-31: extended `PptxInspect`, `ComparePptxTextEmission.ps1`, and
  `SummarizePptxTextStateDeltas.ps1` with private-safe uppercase/lowercase/titlecase counters to test whether
  the private-deck `Tc` buckets are a letter-case decomposition rule. They are not. Refreshed private summaries
  for pages 36, 21, and 79 show incompatible positives and counterexamples: page 36 has `131/133` matched
  operations with nonzero Office `Tc`, dominated by lowercase-only single-span frame-15 rows such as
  `20` operations with `upper=0`, `lower=27`, `spaces=0`, and `Tc=-0.036`; page 21 has many zero-Office-`Tc`
  rows with nonzero candidate residuals and similar mixed-case text shapes; page 79 has table-heavy nonzero
  Office `Tc`, including large buckets with no letters after text-operation matching. Public probes confirm the
  rejection outside the private deck: `pptx-ladder-04-typography-capital-spacing-probe`,
  `pptx-ladder-04-typography-dense-column-probe`, and `pptx-ladder-04-typography-run-boundaries` all emit
  Office `Tc=0` despite nonzero candidate residuals and varied uppercase/lowercase shapes, while
  `pptx-ladder-04-typography-unspaced-column-tc-probe` emits nonzero `Tc` for uppercase-only unspaced rows.
  Therefore do not add letter-case, all-caps, lowercase-only, or residual-magnitude rules. The viable path is
  still an Office PDF text-state model that combines frame/paragraph/line state, secondary `/Tf` selection,
  glyph residual decomposition, and operation splitting, validated by public probes before changing private
  deck rendering.
  Tooling note, 2026-05-31: `ComparePptxTextEmission.ps1` also has an explicit
  `-MatchByTextShapeThenPosition` diagnostic mode that matches by private-safe text category counts before
  position when candidate text is omitted. This mode is useful for checking whether position-only matching is
  overmatching repeated private runs, but it is intentionally not the default because Office can split or merge
  text operations in ways that make strict shape matching report extra unmatched rows. Treat it as a second
  lens, not as the sole source of page-level counts.
  Follow-up, 2026-05-31: accepted the content-minimum overflow row-allocation rule for slack tables. Public
  `pptx-ladder-10-table-center-explicit-multiline` showed that Office lets row content minima exceed the
  `graphicFrame` height instead of compressing all rows back into the frame when the minimum total is too tall.
  Returning the computed minimum row heights in that case moved the public multiline table from MAE
  `13.537929` to `10.442563` (run `20260531-002537`) and produced the Office-like `27.37pt` row cadence
  seen in `PptxInspect`, while `pptx-ladder-10-table-center-explicit-wrapped` stayed unchanged at MAE
  `8.201281`. Private `lokad-value-based` run `20260531-002604` stayed neutral versus the table-`Tc` cleanup
  checkpoint (`2.933659` deck MAE, changed16 `0.053398`; page 21 remained `5.943068`). Keep this as a
  structural table-layout correction, not a page-21 fix: page 21 remains dominated by table text-state,
  smaller-row baseline residuals, and extra Office text-operation splitting. Public regression
  `PptxSyntheticTableKeepsOverflowingContentMinimumRowsUnscaled` locks the branch; validation passed with
  solution build and non-slow `pptx-tables` (`19` passed).
- [ ] 2026-05-30: Pursue the `Tc`/secondary-`Tf` branch as font-metric text-state decomposition, not a
  font-family shortcut. Public probes now reproduce the private pages' family: `pptx-ladder-04-typography-
  spautofit-tracking-probe` has Office `12.024pt` plus `Tc=-0.036` where the candidate emits `12pt/Tc=0`,
  and `pptx-ladder-04-typography-spautofit-tracking-narrow-probe` has Office `9.984pt` and `Tc=-0.036`
  against candidate `9.96pt/Tc=0`. A temporary public-safe font-agnostic probe under
  `artifacts/tmp/tracking-font-agnostic-probe` showed the same highlighted-boundary structure producing
  different Office `Tc` buckets across fonts (`0.02`, `0.012`, `-0.036`) and secondary font-size buckets
  on only some spans. Local font table inspection did not find an OpenType `trak` table in the common Windows
  fonts tested, so the next acceptable implementation path is resolved font metrics/GPOS/kerning plus
  Office-like run-boundary PDF text-state splitting. Do not key this branch on `Cambria Math` or any other
  family name.
  Follow-up, 2026-05-30: isolated page 36's three-column no-autofit text-state branch with public-safe probes.
  A generic three-column Aptos probe with spaces reproduced only `Tc=0`; a space-free Aptos variant still
  reproduced only `Tc=0` while exposing Office's secondary `12.024pt` font-size branch. Replacing only the
  resolved typeface with the Windows math-profile face reproduced the private family publicly:
  `pptx-ladder-04-typography-unspaced-column-tc-probe` has Office `Tc` buckets `0`, `0.0195`, and `-0.036`
  with secondary `12.024pt` font-size operations, while the candidate still emits `Tc=0`. This is evidence
  for a resolved font-metric/profile and PDF text-state decomposition rule, not permission to special-case
  the family name. Validation: `CheckVisualCase.ps1` passed for the new public probe in run `20260530-234639`.
- [ ] 2026-05-30: Resolve the secondary Office PPTX `/Tf` emission branch before adding another private-deck
  text-spacing rule. Two ignored public-safe probes under `artifacts/tmp/office-probes/` now isolate the
  private page-55 signal without private content. `multicol-synth-bold-math` uses a three-column
  `noAutofit` Cambria Math frame with synthetic bold text; Office emits no nonzero `Tc`, but emits
  `12.024 Tf` on selected wrapped rows while OOXPDF emits `12 Tf` throughout. `multicol-mixed-math` mirrors
  the private frame's alternating bold/regular/italic run structure; Office still emits no nonzero `Tc`, but
  increases the `12.024 Tf` population (`27/150` text operations in the probe reference) while the candidate
  remains at `12 Tf` for all `150` operations. The simple 10 pt Arial probe still confirms the first-order
  Office 600-DPI grid (`9.96 Tf`), so the open rule is the secondary wrapped/text-frame emission branch, not
  a replacement for the base grid. Do not add a broad implicit `Tc` rule from page 55 until a public probe
  reproduces the non-authored `Tc` trigger; current evidence says the safer next rendering slice is
  context-aware `/Tf` emission for wrapped PPTX text.
  Follow-up private page-36 inspection after the chart reserve/pattern work keeps this item active: graphics
  are no longer the dominant p36 gap, while Office emits selected secondary font sizes (`12.024`, `14.064`,
  `15.984`, and related branches) and non-authored text-state families where OOXPDF still emits integer
  sizes and mostly `Tc=0`. The relevant visible frame is an ordinary three-column `noAutofit` overflow text
  frame with explicit column spacing; Office's secondary-size rows cluster by wrapped line position rather
  than by any private text content. Do not shortcut this with private slide coordinates or a blanket `Tc`
  rule; the next acceptable change needs a public probe that predicts the multicol and y-sweep evidence
  together.
- [ ] 2026-05-30: Extend the secondary `/Tf` investigation with the public-safe y-sweep probe under
  `artifacts/tmp/office-probes/font-grid-y-sweep/`. While checking current private top-five pages, slide 20
  and slide 59 both reduced to text-state divergence: their Office/candidate fill and stroke buckets match,
  but Office emits non-authored `Tc` families and secondary font sizes while OOXPDF emits `Tc=0` and only the
  first-order 600-DPI font grid. The y-sweep probe uses plain 10 pt Arial text boxes with no table, chart, or
  private content; Office still emits `9.984 Tf` for selected boxes while the base grid remains `9.96 Tf`.
  This rejects a table-only explanation for the secondary font branch and also rejects any private-slide
  coordinate shortcut. Follow-up public-safe probes under `font-grid-y-fine/` and `font-grid-xy-grid/` further
  rejected a simple position-only rule: a finer y sweep triggered one secondary `Tf` while a comparable x/y
  grid triggered none. The next step is to derive the Office export rule from public probes that vary frame
  top, frame height, body insets, wrap/autofit mode, baseline grid, text length, run splitting, and text matrix
  placement, then encode the rule in `PptxPdfTextEmissionProfile` only after it predicts the y-sweep,
  fine-grid, x/y-grid, and multicol probes.
  2026-05-30 refresh: the current public-safe evidence still does not support a renderer rule. The y-sweep
  probe has `24` rows with `2` Office secondary-size rows; the fine y sweep has `80` rows with `1` secondary
  row; the x/y grid has `72` rows and no secondary branch. The multicol mixed probe still has `27/150`
  Office secondary-size rows, while the synthetic bold multicol probe has `15/150` secondary-size rows. These
  positives and negatives cross font/style/position categories, so the next step remains probe design and
  structural comparison, not a typeface, synthetic-style, or fixed-coordinate predicate.
- [ ] 2026-05-30: Private page 36 remains a high-impact text-state alignment gap. The page is text-heavy
  (`16` text frames, `133` candidate glyph runs, no tables/charts, graphics operation counts essentially
  matched), and the reference PDF emits `135` text operations with nonzero `Tc` almost everywhere while the
  candidate emits `0 Tc` for all glyph runs. The source slide has no authored run-level `spc`; most frames are
  `noAutofit`, one frame has `numCol=3`, and some paragraphs have auto-number bullets or `spcAft=6 pt`, but
  the Office `Tc` values vary by frame (`-0.012`, `-0.036`, `-0.0476`, `-0.0574`, small positives) and do not
  reduce to a font-family, font-size, bullet, or column-count shortcut. Keep this open as a public-probe
  requirement: compare Office `Tc`, `TJ`, glyph positions, and line/frame context together, because scalar
  `Tc` promotion alone would mostly alter PDF structure while `TJ` can compensate the visual positions.
- [ ] 2026-05-30: Continue page-21 table work from explicit-margin centered cells, not default anchoring. Slide
  inspection shows the page-21 table cells use `anchor="ctr"` with explicit symmetric margins (`marL/marR`
  and `marT/marB`) on most cells, and Office still emits small non-authored `Tc` families (`-0.00888` and
  `0.0331`) while OOXPDF emits `Tc=0`. The next publicization step is a valid Office-authored synthetic table
  probe with centered cells, explicit symmetric margins, short 10 pt text, and the same kind of non-uniform row
  ladder, so the explicit-margin anchor and row-height behavior can be separated from private content. Do not
  extend the default-margin inset adjustment to explicit margins until that probe or an equivalent Office PDF
  inspection supports it.
- [ ] Align cropped raster-image PDF resources with Office's cropped-XObject strategy:
  Office's inspected PDFs for the private recurring picture pattern draw the cropped image at the final frame
  matrix, while OOXPDF still emits the full source image scaled larger and clipped to the frame. The visible
  pixels are already stable in public crop cases, so this should not become a coordinate heuristic. The durable
  target is a public synthetic crop/fill/group fixture that compares image draw matrices, clip counts, and image
  resource identity, then a dependency-free image-resource pipeline that can materialize cropped image XObjects
  when Office does so without breaking alpha/recolor/SVG paths.
- [ ] Replace sampled curved-connector filled outlines with analytical Office-like curve-rich paths:
  the public connector transform probe now matches Office at the high-level fill/clip/stroke operator count, but
  PDF inspection still shows candidate filled connector paths as sampled outlines. Arrow-tail connectors now emit
  an Office-like two-subpath fill structure, with a line-sampled connector body plus separate curved arrowhead
  subpath (`move=2,close=2`), instead of the previous single smoothed sampled path. This is still not Office's
  exact connector geometry (`seg=250,line=244,curve=4` in Office versus candidate `seg=244,line=239,curve=3` in
  the public probe), so the remaining work is structural parity rather than a raster emergency.
- [ ] Track the remaining private page-17 text/font-size gap while preserving the closed clip history:
  this item started when private page 17 still had one missing high-level `W*` clip and dominant derived font-size
  differences: Office reported `9,9.024,9.96,12,12.96,12.984,14.04,15.96,18`, while the candidate reported
  `9,10,12,13,14,16,18`. Subsequent public-backed clipping and font-grid work closed the graphics count and the
  dominant `10/13/14/16` quantization mismatch; the still-open private signal is now only the secondary Office
  `/Tf +0.024 pt` branch (`9.024` and `12.984`) plus ordinary small text-position deltas. Keep future changes tied
  to public-safe font-emission probes rather than a private slide-specific rule. Public image evidence already rules
  out normal stretched raster pictures as the former clip source.
- [ ] Resolve the recurring high-error private page text-fit pattern through public-safe probes:
  the latest private deck baselines put page 50, 32, 36, 49, 13, 21, 83, 53, 82, and 81 above the remaining
  deck error floor. Private-safe inventory shows the top cluster is picture/group/connector heavy, but page-50
  inspection split the evidence: image frame clips match Office after rounding, while reference text has more
  operations (`161` versus `143`) and many negative `Tc` values on no-explicit-spacing text where the candidate
  emits `0 Tc`. The public `spAutoFit`/tracking probes do not reproduce this negative-spacing branch, so do not
  add a blanket character-spacing compression rule. The next useful public probes should isolate no-autofit,
  inherited centered anchor, Cambria Math, `spcPct=90%`, multi-run highlight, long-line wrapping, and multi-column
  body properties until the Office condition for export-time negative `Tc` and line splitting is observable.
- [ ] Complete the Office PPTX-to-PDF text font-size emission profile:
  ignored Office-generated probes under `artifacts/probes/font-size-quantization*` show the generic export rule
  outside the private deck (`7->6.96`, `8->8.04`, `10->9.96`, `13->12.96`, `14->14.04`, `16->15.96`,
  `19->18.96`, `20->20.04`, while `6`, `9`, `12`, `18`, and `30` remain exact). Do not hide this behind
  per-size lookups. The dominant 600-DPI emission grid is now implemented, but the remaining secondary
  `+0.024 pt` anomalies in public probes and the private `9.024` / `12.984` sizes still need a structural
  explanation before adding another rule.
- [ ] Explain the secondary Office `/Tf +0.024 pt` branch with public-safe probes before changing emission again:
  a new ignored probe under `artifacts/probes/font-size-quantization-cambria` rewrote the no-autofit quantization
  deck to use Cambria Math, then rendered it through Office. The result still mapped `9 pt -> 9` and `13 pt -> 12.96`,
  so the private page-17 secondary `9.024` and `12.984` sizes are not explained by Cambria/Cambria Math alone. The
  follow-up ignored wrap probes under `artifacts/probes/font-size-quantization-wrap*` reproduced both branches:
  a wrapped 9 pt run emitted a mix of `9` and `9.024`, and a narrow wrapped 13 pt run emitted a mix of `12.96` and
  `12.984`. The remaining rule is therefore not font-family-specific; it is tied to Office's wrapped/split text-line
  PDF emission. Do not change `PptxPdfTextEmissionProfile` until the per-line condition is known well enough to lock
  with a public probe instead of a private slide-specific branch.
- [ ] Explain Office `/Tf` dependence on text-frame geometry before extending the font-size profile:
  denser ignored public-safe font-size probes showed that the secondary `+0.024 pt` branch is not a pure function
  of requested point size. The same nominal sizes can emit exact values or `+0.024`/nearby branches depending on
  frame geometry, insets, generated textbox height, and wrapped-line state. Treat the current 600-DPI grid as the
  stable first-order export approximation, but do not add per-size exceptions for `15`, `21`, `24`, `36`, or wrapped
  `9`/`13` until the renderer can derive the same context from the text-frame layout/export model. A follow-up
  inspection pass confirmed that default body insets alone do not trigger the secondary branch: a dense/default-inset
  probe still emits exact `15`, `21`, `24`, `30`, and `36`, while the wide/no-autofit probes emit `15.024`,
  `21.024`, `24.024`, exact `30`, and `36.024`. The next probe must vary generated frame height and line count while
  keeping shape width and body insets fixed, otherwise a renderer rule would still be a hidden geometry heuristic.
  Two ignored public-safe height scans under `artifacts/probes/font-size-quantization-height-scan*` kept width,
  default insets, `noAutofit`, and single-line text fixed, but they varied shape height and absolute page Y together.
  They still ruled out text length and wrapping as the discriminator: a one-character scan maps `21`/`24`/`36` to
  `+0.024 pt` only at selected rows (`15` remains exact), while adjacent rows with the same size remain exact. This
  points to an Office export quantization interaction involving absolute text-frame placement/baseline placement,
  not a font-size table.
- [ ] Reverse-engineer the remaining PPTX text formulas from Office/PDF evidence and `pptx-renderer`
  semantics, replacing local constants with named general rules for hidden advances, baseline offsets,
  line advances, highlight rectangles, and text-operation boundaries.
- [ ] Continue replacing text constants with formula-owned measurements, starting with baseline/line-box
  offsets and highlight/strike geometry, and lock each rule with Office-PDF text-operation or rectangle
  probes before broad visual MAE gates.
- [ ] Extend the cascade model from paragraph defaults to a full named seven-level resolver with separate
  paragraph, run, bodyPr, placeholder geometry, and theme font/color fallback stages.
  2026-05-28 progress: `PptxParagraphStyleCascade` now owns the Office-order merge that produces paragraph
  default properties, so the text model no longer reconstructs default paragraph XML directly from `Sources`.
  The duplicated placeholder `txStyles`/`defaultTextStyle` lookup in renderer and scene inspection was also
  centralized in `PptxTextStyleInheritance`, reducing the chance that the PDF path and inspection path drift while
  the full resolver is staged. Run style resolution now enters through `PptxRunStyleCascade`, with the current
  flattened default run properties carried by the cascade to preserve behavior while giving the next slice a named
  resolver boundary. Keep this item open: scene inspection still has its own `ResolveDefaultParagraphProperties`
  merge chain, run defaults are not yet re-merged from all named run layers, and theme font/color fallback remains
  distributed across resolver helpers.
  2026-05-28 progress: the duplicated paragraph-property merge primitives are now centralized in
  `PptxParagraphPropertyMerger`, with separate named renderer and scene strategies because source inspection found
  they are not semantically identical. The renderer keeps its existing Office-order default paragraph merge that
  overlays `a:defRPr`, while scene inspection keeps its recursive child merge. This is not a rendering change and
  it intentionally extends the open item: the long-term resolver must decide, with public evidence, whether the
  scene snapshot should adopt the renderer's paragraph merge semantics or the renderer should consume a richer
  scene-owned merge. Validation: `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed; focused
  non-slow `pptx-typography` passed (`99 passed, 0 failed, 2 skipped`); focused non-slow `pptx-model` passed
  (`25 passed, 0 failed, 1 skipped`); full non-slow console runner passed
  (`409 passed, 0 failed, 7 skipped`).
  2026-05-29 progress: scene paragraph-property merging now overlays child elements instead of replacing them
  wholesale. This preserves lower-priority `a:defRPr` attributes such as master body-style `b="1"` while higher
  layers supply only their own tokens, for example a layout placeholder `i="1"` or `sz="2600"`. The fix keeps
  scene inspection aligned with the renderer's structural merge direction and removes a null edge in chart
  style effect-reference snapshot counting. Keep this item open: the merge is still XML-structural rather than
  a fully named seven-stage resolver. Validation: `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal`
  passed; focused `PptxSceneBuilderBuildsResolvedNodeLists` passed; full non-slow console runner passed
  (`414 passed, 0 failed, 7 skipped`).
- [ ] Tighten justified paragraph parity against Office: current public probe still has large raster drift,
  so inspect Office text operations and refine baseline, line advance, and per-word spacing strategy.
- [ ] Introduce a renderer-level indexed chart data vector before removing the remaining chart value/category
  XML fallbacks. The scene and workbook layers already preserve `ptCount`, sparse point indices, blank or
  non-numeric points, multi-area/structured-reference workbook cell indices, and category string-point
  presence. Bar/column, line, area, pie/doughnut, Cartesian data labels, category-axis labels, radar labels,
  polar legends, series-name labels, and series-name legend sizing now consume indexed vectors or
  provenance-preserving records at their current renderer boundaries. Remaining open work is narrower:
  chart text/layout decisions still need to carry the same typed data records all the way to
  Office-PDF-backed placement, formatting, and source/cache freshness decisions.
- [ ] Extend the chart data-label scene model beyond the currently consumed subset: leader-line geometry,
  exact Office label-box geometry/auto-fit, and richer position semantics still need typed renderer support
  before automatic data-label layout can be aligned structurally with Office instead of renderer heuristics.
  - [ ] Resolve the polar data-label manual-layout coordinate basis before changing leader-line visibility:
    Office positions the same four custom pie labels near the chart sides while OOXPDF still derives some
    boxes outside the frame. Investigate whether data-label `x`/`y` factors are relative to chart frame,
    plot area, default label anchor, text box, or an extension-owned layout coordinate, then lock the rule
    with text-operation and leader-line structural checks before adding a cardinality gate.
    2026-05-28 update: added `tools/SummarizeChartDataLabelLayout.ps1` so this gap can be measured from
    Office/candidate PDF structures instead of visual inspection. On public run `20260528-125303`, the
    summary reports `DataLabelText` reference/candidate counts `8/4`, max nearest text bounds delta
    `492.681pt`, `DataLabelLeaderLineCandidate` counts `1/4`, and nearest leader-line bounds delta
    `82.054pt`. The Office-authored chart XML stores per-label `c:manualLayout` `x`/`y` values without
    explicit modes, including negative factors; those factors came from COM `DataLabel.Left`/`Top` writes
    and are not behaving like ordinary absolute plot-box factor boxes. Keep the next renderer change tied to
    this public structural evidence, not to a per-fixture coordinate constant.
    2026-05-28 update: extended `tools/SummarizeChartDataLabelLayout.ps1` with an optional `-ChartXml`
    input so the same public summary now emits the source `c:dLbl` manual-layout records (`idx`, `dLblPos`,
    `x/y/w/h`, mode fields, and label visibility flags) beside the PDF text/leader-line nearest-match
    deltas. On current run `20260528-150942`, it reports four source manual-layout records with omitted
    modes and mixed negative/positive `x`/`y` values, while the PDF structural deltas remain
    `DataLabelText` `8/4`, max text bounds delta `492.681pt`, and leader-line counts `1/4`. This is
    preserved as correlation tooling for the next layout-rule change, not a renderer approximation.
    2026-05-28 update: split generated polar data-label components into separate PDF text runs when Office's
    separator is whitespace-only. This keeps comma/default separators as emitted glyphs for existing
    compatibility, but aligns the public Office probe where category and percent are separate PDF text
    structures separated by positioned whitespace. Public leader-line probe run `20260528-151601` improved
    `DataLabelText` counts from reference/candidate `8/4` to `8/8` and reduced max nearest text bounds delta
    from `492.681pt` to `262.481pt`; leader-line counts remain intentionally open at `1/4` until the manual
    label-box coordinate basis is solved. Validation: focused non-slow `pptx-charts` passed
    (`117 passed, 0 failed, 0 skipped`); full non-slow console runner passed
    (`370 passed, 0 failed, 7 skipped`); private run `20260528-151650` stayed at 84/84 compared pages, zero
    dimension mismatches, deck MAE `6.715278`, changed16 `0.093542`, and only
    `PPTX_UNSUPPORTED_IMAGE_RECOLOR`.
    2026-05-28 update: tightened the public leader-line probe manifest to compare `DataLabelText` structures
    with a deliberately loose `300pt` bounds tolerance. Run `20260528-151851` passed this new text gate with
    reference/candidate counts `8/8`, while still leaving `DataLabelLeaderLineCandidate` ungated. This prevents
    regressions back to combined polar label strings without pretending the unresolved manual-layout geometry
    is solved.
    2026-05-28 update: tested and rejected the simple "omitted data-label layout modes are absolute factors"
    hypothesis on the same public probe. That experiment moved lower labels farther outside the chart and
    failed the `DataLabelText` structural gate in run `20260528-152413`, so the renderer change was reverted.
    The durable conclusion is negative but useful: the missing rule is not a toggle between absolute plot-box
    factors and current auto-label offsets. The summary tool now also groups split `DataLabelText` operations
    into public-safe label clusters, reporting reference/candidate cluster counts `4/4` and max nearest
    cluster bounds delta `262.481pt` on run `20260528-151601`; this keeps future coordinate-basis work focused
    on whole labels instead of individual category/percent text fragments.
    2026-05-28 update: extended the same summary tool to correlate each source manual-layout record with
    reference/candidate label clusters by quadrant around the polar plot box. Run `20260528-152722` now records
    public-safe relative cluster centers beside the source factors: for example the top-left label has manual
    `x=-0.7702` and Office relative X `-0.7979`, while top-right has manual `x=0.6602` and Office relative X
    `0.6778`; Y factors and at least one lower-left X value do not line up as directly. This is enough to
    reject a single symmetric x/y scaling shortcut and points the next investigation at Office's conversion
    from COM label positions into `c:manualLayout`, including label-box size and quadrant-specific anchoring.
    2026-05-28 update: `tools/SummarizeChartDataLabelLayout.ps1` can now read the source chart XML directly
    from a PPTX via `-Pptx`/`-ChartPart`, eliminating the manual unzip step from this public probe workflow.
    The direct-PPTX run over `20260528-152722` reproduced the four manual-layout records and the same
    cluster/leader-line deltas as the pre-extracted `chart1.xml` path.
    2026-05-28 update: extended the summary tool again to compute source-side expected label component hashes
    from the chart XML and correlate Office/candidate clusters by content hash set, not just by polar quadrant.
    This preserves a stronger invariant for the coordinate-basis work: on run `20260528-152722`, hash-matched
    clusters show index `0` (`Alpha`/`48%`) at Office relative center roughly `(-0.708, 0.470)` and candidate
    `(-0.967, 0.371)`, while index `2` (`Gamma`/`17%`) sits at Office `(-0.798, -0.528)` and candidate
    `(-0.439, -0.699)`. The earlier quadrant-only evidence remains useful for side placement, but it is not
    strong enough to bind a `c:dLbl idx` to its displayed label content; future renderer changes should use
    the hash-matched rows before changing missing-mode semantics or leader-line cardinality.
    2026-05-28 update: extended `tools/NewChartProbeFixtures.ps1` so regenerating the public pie data-label
    leader-line probe also writes ignored COM metadata under `artifacts/office-probe-metadata/`. The sidecar
    records the chart shape bounds plus each label's requested `DataLabel.Left`/`Top`, observed
    `Left`/`Top`/`Width`/`Height`, applied position, and text. This fills the current evidence gap between the
    generator's requested COM coordinates, the saved `c:manualLayout` factors, and Office's PDF label boxes
    without changing renderer placement from an underdetermined four-point sample.
    2026-05-28 update: `tools/SummarizeChartDataLabelLayout.ps1` now accepts that sidecar through
    `-ComMetadataJson` and joins COM requested/observed label rectangles onto `ChartManualLayouts` and
    `DataLabelManualLayoutCoordinateEvidence`. A temporary public-safe sidecar over run `20260528-152722`
    verified the join path with four COM label records while preserving the existing PDF structural deltas
    (`DataLabelText` `8/8`, cluster max delta `262.481pt`, leader-line counts `1/4`). The next real Office
    regeneration can now compare COM coordinates, saved XML factors, and Office PDF boxes in one summary.
    2026-05-28 update: the summary tool also auto-discovers
    `artifacts/office-probe-metadata/<pptx-basename>/com-metadata.json` when `-Pptx` is supplied, keeping
    explicit `-ComMetadataJson` as the override. A temporary ignored sidecar verified the default path reports
    `ChartComMetadataDataLabelCount=4`, so future Office regeneration does not need an extra manual argument
    to keep COM/PDF/XML evidence together.
    2026-05-28 update: the summary tool now normalizes COM-observed label centers against the COM chart-shape
    frame and reports deltas to the hash-matched Office and candidate PDF clusters. On public run
    `20260528-155031`, the public-safe summary still reports text clusters `4/4`, max cluster bounds delta
    `262.481pt`, and leader-line counts `1/4`; the COM-relative centers do not collapse into a simple direct
    rule against the Office PDF clusters, with X only partly aligned and Y sign/basis disagreements across the
    four labels. Keep the next renderer change blocked on a multi-document Office-derived coordinate model,
    not on a four-label frame-relative shortcut.
    2026-05-29 update: added a second public Office-authored pie data-label leader-line probe,
    `pptx-ladder-11-chart-pie-data-label-leader-lines-offset-probe`, with a different chart frame and manual
    COM label coordinates. `tools/NewChartProbeFixtures.ps1 -DataLabelsOnly` now emits this fixture plus its
    ignored COM sidecar so coordinate-basis work can compare multiple Office geometries instead of fitting the
    original four-label sample. The new probe passes its loose visual/structural gates in run
    `20260529-013658`; the refreshed original leader-line probe and Cartesian legend-key probe also pass in
    runs `20260529-013716` and `20260529-013729`. The offset summary reports text clusters `4/4`, max cluster
    bounds delta `453.645pt`, manual-layout records `4`, COM label records `4`, and leader-line counts `0/4`.
    This extends the negative evidence: Office may emit no visible connectors for a custom-layout pie even
    when `HasLeaderLines` is true, so the eventual rule must combine final label boxes, clipping, and
    Office-visible connector routing rather than trimming candidate leader lines by a fixed count.
    2026-05-29 update: `tools/SummarizeChartDataLabelLayout.ps1` now reports content-hash matched text and
    label-cluster deltas beside nearest-neighbor deltas. This matters for the multi-probe model: the offset
    run's nearest cluster delta is `453.645pt`, but its hash-matched cluster delta is `214.585pt`, proving
    nearest geometry alone can pair the wrong labels once multiple custom labels are off by large distances.
    The original probe remains hash-stable at `4/4` clusters with max hash delta `262.481pt`.
  - [ ] Derive Office leader-line visibility and cardinality from the final label-layout model before gating:
    the first renderer consumption pass deliberately draws all visible polar labels that request leader lines,
    while Office emits only one visible connector in the current public custom-layout probe. Structural report
    `20260527-122348` shows reference/candidate `DataLabelLeaderLineCandidate` counts `1/4`; keep the kind
    ungated until label-box placement, clipping, and Office's per-label leader visibility rule are understood.
    2026-05-28 update: `tools/SummarizeChartDataLabelLayout.ps1` now attaches each leader-line structure to
    the nearest hash-identified label cluster. On public run `20260528-152722`, Office's only visible leader
    line is nearest source label index `2` (`Gamma`/`17%`) at distance `49.91pt`, while the candidate emits
    four connectors nearest indices `3`, `2`, `0`, and `1`. This makes the remaining gate content-aware: do
    not reduce candidate leader lines by a count-only or quadrant-only heuristic; derive visibility from the
    final Office-aligned label boxes and then assert the surviving connector against the matched label index.
    2026-05-29 update: the offset-frame public probe adds a second visibility case where Office emits zero
    classified data-label leader-line structures while the candidate still emits four. This keeps the open
    cardinality rule explicitly multi-fixture: one public custom-layout probe is `1/4`, the offset-frame probe
    is `0/4`, and neither supports a hard-coded "draw one" or "draw none" shortcut.
- [ ] Extend chart area/plot area style modeling beyond direct solid fill/no-fill, pattern fill, and simple
  line: gradient fill, theme style references, transparency groups, and effect inheritance need typed
  ownership before chart background and plot-box styling can be considered structurally Office-aligned.
  - [ ] Add a real PDF soft-mask or equivalent transparency-function backend before rendering non-uniform
    per-stop gradient alpha. The current PDF axial shading model is RGB-only, so varying stop alpha must
    remain structurally represented but not approximated by averaged or stop-flattened opacity.
- [ ] Extend chart text-style ownership beyond simple `defRPr` font/color/size and default-run bold/italic:
  rotation, full non-default tick-label offset ladders, multi-level category labels, default-placement axis-title
  cascades, and richer chart-style text dimensions still need structural modeling before chart text can match
  Office without renderer heuristics.
  - [ ] Derive and render default-placement axis titles:
    the first default-placement category/value axis-title renderer is in place for supported Cartesian native
    chart branches, but the item stays open until plot-box reservation, top/right permutations, overlay
    behavior, and chart-style inherited defaults are structurally proven against Office PDFs. This remains
    distinct from tick-label styling and should continue with Office-PDF-backed placement evidence, not
    frame-relative text nudges.
  - [ ] Replace the initial default-axis-title placement approximation with structural Office layout:
    current horizontal and rotated title positions use named Office-observed ratios over the existing chart
    frame/plot-box reserves. The next slices should add public probes for top and right axis titles,
    horizontal-bar category/value permutations, scatter/bubble explicit titles, overlay/reserve interaction,
    rich rotated multi-run titles, and chart-style-inherited title defaults, then feed those observations back
    into plot-box reservation rather than widening the title-position tolerance.
- [ ] Extend plot-area layout ownership beyond current `x/y/w/h` factor and right/bottom edge support:
  `layoutTarget`, x/y edge semantics, inner-vs-outer plot area semantics, title/legend overlay interactions,
  and reuse across area/scatter/radar/pie/doughnut chart families still need structural modeling before plot
  bounds can be treated as Office-aligned instead of approximate geometry.
- [ ] Extend data-label rendering to consume the remaining richer scene metadata: leader-line geometry,
  automatic Office label-box geometry/auto-fit, and richer position semantics still need renderer support and
  visual cases.
- [ ] Replace fixed chart auto tick target constants with an Office-aligned axis layout model that derives
  target tick density from axis length, label text extents, number format, orientation, and available label
  bands. The new horizontal-axis target is useful evidence, but it is still a named metric rule; the long-term
  state should select major units from structural axis-layout constraints rather than chart-family constants.
- [ ] Replace frame-relative fallback chart-title placement with an Office-aligned title layout model that
  preserves/consumes title `overlay`, explicit/manual layout, title box geometry, inherited text style, and
  measured text extents. The new horizontal-bar title baseline is a useful structural anchor, but it is still
  a constrained fallback rule; long-term title placement should derive both X and Y from the Office chart
  layout object model and font metrics rather than fixed inset/width ratios.
- [ ] Finish secondary-axis structural alignment for chart families beyond the current supported bar/line
  paths: crossing geometry still needs to consume the preserved scene metadata, and exact Office spacing still
  needs explicit model-to-renderer plumbing.
- [ ] Continue data-label rendering alignment with Office: leader-line geometry, label layout/auto-fit, and
  exact Office label-box geometry remain approximate.
- [ ] Consume plot-area manual-layout target and mode semantics in geometry: `layoutTarget=inner`,
  edge/factor modes, title/legend overlay interactions, and non-bar/line chart-family plot boxes still need
  Office-evidenced rendering rules.
  2026-05-29 update: removed the renderer-local chart manual-layout XML parser. Raw XML fallback for
  plot-area, title, legend, and data-label manual layouts now materializes `PptxSceneChartManualLayout`
  through `PptxSceneBuilder.ReadChartManualLayout`, so the scene parser remains the single owner of
  layout-target/mode/raw-factor token semantics. This is intentionally behavior-neutral and does not close
  geometry parity: the public inner/outer plot-layout target probes still pass their bounds/text gates, but
  both continue to show Office `PlotAreaClipBoxCandidate` segment count `5` versus candidate `8`. The next
  long-term step is a shared chart-layout oracle for chart area, plot area, inner plot, axes, title, legend,
  and clip boxes, not another fixture-specific coordinate rule. Validation: `dotnet build Lokad.OoxPdf.slnx
  --tl:off --nologo -v minimal` passed; focused non-slow `pptx-charts` passed (`141 passed, 0 failed,
  0 skipped`); plot-layout visual probes passed in runs `20260529-024411` and `20260529-024423`.
- [ ] Keep raw XML on scene nodes until typed coverage is complete, but make new renderer code prefer typed
  fields and resolver outputs instead of repeated ad hoc descendant queries.
- [ ] Decide whether `ooxpdf` needs an intermediate presentation scene/model between OOXML parsing and PDF
  generation before more large changes to `PptxRenderer`.
- [ ] Port `pptx-renderer` package/model boundary patterns: relationship target resolution, media lookup,
  slide/layout/master/theme ownership, and raw XML retention on typed nodes.
- [ ] Port `pptx-renderer` render-context ownership: one context object should expose slide identity,
  inheritance, relationships, media caches, theme/color resolvers, diagnostics, and group-fill state.
  2026-05-28 update: `PptxRenderContext` now exposes slide/master/layout part names, relationship maps, and
  color maps directly from the scene slide, and render/diagnostic call sites that need slide identity consume
  `context.SlidePartName` instead of reaching back through the raw `PptxSlide`. This does not finish the context
  split, but it makes the intended ownership explicit for future specialized renderers and reduces ad hoc
  document/scene reach-through.
- [ ] Port `pptx-renderer` specialized renderer split: background, shape, text, table, image, group, chart,
  and fallback renderers should consume the same context instead of ad hoc XML/document parameters.
- [ ] Port `pptx-renderer` slide inheritance behavior: render master/layout non-placeholder template content
  behind slide content, and treat placeholders as inheritance templates for geometry/body/text styles.
- [ ] Port `pptx-renderer` placeholder matching rules: match by `idx` before type, handle default/body/title
  fallback consistently, and keep placeholder geometry separate from text-body inheritance.
- [ ] Port `pptx-renderer` theme font resolution: major/minor Latin, East Asian, complex script, and symbol
  font fallback must be explicit diagnostics-bearing stages before glyph mapping.
  2026-05-28 progress: `PptxTheme.ResolveTypefaceWithSource` now returns a source-bearing typeface result
  for direct, default-minor, major/minor Latin, East Asian, complex-script, and Latin-fallback branches, and
  text run model snapshots expose `TypefaceSource`. Keep this item open: bullet and chart text paths still
  mostly consume plain resolved typeface strings, symbol/bullet font policy is not represented as a diagnostic
  stage, and glyph fallback remains downstream in font mapping rather than a theme-font stage.
  2026-05-28 progress: scene-owned chart text style overrides now preserve the requested OOXML typeface token
  and nullable `PptxThemeTypefaceSource` alongside the existing flattened font family. This covers chart-level
  `txPr`, data-label/title/legend/axis `txPr`, rich chart text runs, and chart-style-part `fontRef`/`defRPr`
  text roles without changing PDF rendering. Keep this item open: renderer chart text still consumes only the
  flattened font family, chart text has surface-local run-property readers instead of a shared cascade object,
  bullet font selection is still not represented as an explicit source-bearing theme-font stage, and glyph
  fallback remains downstream in font mapping. Validation: `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo
  -v minimal` passed; focused non-slow `pptx-charts` passed (`128` passed, `0` failed, `0` skipped); focused
  non-slow `pptx-model` passed (`17` passed, `0` failed, `1` skipped); full console runner passed (`396`
  passed, `0` failed, `0` skipped).
  2026-05-28 progress: bullet font selection is now source-bearing at the paragraph model boundary.
  `PptxParagraphBulletModel` preserves the requested `a:buFont @typeface`, resolved typeface, and
  `PptxThemeTypefaceSource`, and `ReadBulletStyle` consumes that typed result instead of resolving the theme
  font during layout. Keep this item open for chart renderer consumption, shared chart run cascades, symbol
  font diagnostics, and glyph fallback staging. Validation: focused non-slow `pptx-typography` passed
  (`98` passed, `0` failed, `2` slow skips); full non-slow console runner passed
  (`402` passed, `0` failed, `7` slow skips).
  2026-05-28 progress: shape text in the scene model now preserves source-bearing typeface resolution too.
  `PptxSceneParagraphStyle` and `PptxSceneRunStyle` carry `PptxThemeTypefaceSource` beside the resolved
  typeface, and the theme EA/CS fixture now checks the scene run style in addition to the renderer text-model
  snapshot. Keep this item open for chart renderer consumption, chart run cascades, symbol font diagnostics,
  and glyph fallback staging. Validation: focused non-slow `pptx-model` passed (`25` passed, `0` failed,
  `1` skipped); focused non-slow `pptx-typography` passed (`98` passed, `0` failed, `2` slow skips); full
  non-slow console runner passed (`402` passed, `0` failed, `7` slow skips).
  2026-05-28 progress: chart renderer text styles now preserve typeface provenance after consuming scene
  chart text styles. The renderer-local `ChartTextStyle` and `ChartTextStyleOverride` records carry the
  requested OOXML typeface token and nullable `PptxThemeTypefaceSource` beside the flattened font family, and
  chart text merges preserve that source when later overrides do not supply a font. Keep this item open for a
  shared chart run cascade object, symbol font diagnostics, and downstream glyph fallback staging. Validation:
  `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed; focused non-slow `pptx-charts`
  passed (`132` passed, `0` failed, `0` skipped); full non-slow console runner passed (`403` passed,
  `0` failed, `7` slow skips).
- [ ] Port `pptx-renderer` color resolver coverage: color maps, theme colors, `phClr`, scheme colors,
  preset colors, HSL/scrgb colors, alpha/lum/tint/shade modifiers, and fallback colors.
  2026-05-28 progress: solid color parsing is now centralized in `PptxColorResolver` and shared by the
  direct PDF renderer and typed scene builder instead of maintaining copied OOXML color logic in both paths.
  The shared resolver covers current `srgbClr`, `schemeClr`/`phClr`, `sysClr`, `prstClr`, `scrgbClr`,
  `hslClr`, alpha, and lum/tint/shade transforms, with renderer image/effect helpers delegating back to the
  same byte/alpha routines. Keep this item open: color-map overrides, source-bearing diagnostics, fallback
  provenance, and full format-scheme ordering are still not represented as model-visible resolver stages.
  Validation: `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed with zero warnings;
  focused non-slow `pptx-model` passed (`13 passed, 0 failed, 1 skipped`), `pptx-typography` passed
  (`96 passed, 0 failed, 2 skipped`), and `pptx-charts` passed (`121 passed, 0 failed, 0 skipped`).
  2026-05-28 progress: introduced a typed `PptxColorMap` owner for the default Office scheme aliases
  (`bg1->lt1`, `tx1->dk1`, `bg2->lt2`, `tx2->dk2`, accent and hyperlink slots) plus parsed `clrMap`
  attributes, and added color-resolver overloads that can resolve `schemeClr` through an explicit map while
  preserving current default behavior. Keep this item open: the scene/render context still needs to carry
  master/layout/slide `clrMap` and `clrMapOvr` provenance into every color-consuming path. Validation:
  `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed; focused non-slow `pptx-model`
  passed (`15` passed, `0` failed, `1` skipped).
  2026-05-28 progress: scene slides now preserve master/layout/slide color-map provenance as typed
  `PptxColorMap` values, layout/slide `clrMapOvr` uses inherited slot mappings instead of sparse standalone
  maps, scene inspection exposes deterministic private-safe map snapshots, and background colors use the
  map for their own source level. Keep this item open: shape, text, table, chart, image recolor, placeholder,
  and format-scheme color consumers still need to receive the effective map systematically instead of falling
  back to the default aliases. Validation: `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal`
  passed; focused non-slow `pptx-model` passed (`16` passed, `0` failed, `1` skipped).
  2026-05-28 progress: direct scene text-body parsing now passes the current source `PptxColorMap` into
  paragraph default-run colors, run colors, and shape `fontRef` fallback, so slide/layout/master text owned by
  the current source resolves `schemeClr` aliases through that source's effective color map. Keep this item
  open and extend it with a specific source-provenance gap: inherited placeholder text-style APIs still return
  bare `XElement`/`XDocument` values, so default properties copied from layout/master placeholders do not yet
  carry their owning color map. Validation: focused non-slow `pptx-model` passed (`17` passed, `0` failed,
  `1` skipped).
  2026-05-28 progress: direct scene shape parsing now passes the current source `PptxColorMap` through
  explicit solid fills, fill/line format-scheme references, explicit lines, theme line references, gradients,
  supported pattern fills, glow, and outer shadow, while the renderer-facing helper overloads keep their
  default-map behavior. Keep this item open: table/chart/image recolor and direct renderer paths still need the
  effective map, and inherited placeholder style maps still require source-aware style ownership. Validation:
  focused non-slow `pptx-model` passed (`17` passed, `0` failed, `1` skipped).
  2026-05-28 progress: direct scene table parsing now passes the current source `PptxColorMap` into explicit
  cell fills and explicit cell border colors, with default-map overloads preserved for existing helper callers.
  Keep this item open: built-in table style colors, chart colors, image recolor, direct renderer paths, and
  inherited placeholder style maps still need source-aware color-map ownership. Validation: focused non-slow
  `pptx-model` passed (`17` passed, `0` failed, `1` skipped).
  2026-05-28 progress: scene picture parsing now passes the current source `PptxColorMap` into duotone image
  recolor color resolution, while keeping default-map helper overloads for current direct renderer callers. Keep
  this item open: chart colors, direct renderer paths, built-in table style color aliases, and inherited
  placeholder style maps still need source-aware color-map ownership. Validation: focused non-slow `pptx-model`
  passed (`17` passed, `0` failed, `1` skipped).
  2026-05-28 progress: scene chart parsing now threads the current source `PptxColorMap` into chart-owned
  chart-space, plot-area, title, and legend style consumers: explicit fills, lines, pattern fills, gradients,
  glow, outer shadow, and chart text-style overrides resolve `schemeClr` aliases through the slide/layout/master
  source map instead of the default aliases. Keep this item open: chart series, point, marker, data-label,
  axis, color-style, and chart-style-part colors still need source-aware map ownership; direct renderer paths,
  built-in table style aliases, and inherited placeholder style maps also remain open. Validation:
  `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed; focused non-slow `pptx-model` passed
  (`17` passed, `0` failed, `1` skipped).
  2026-05-28 progress: chart direct formatting now carries the same source `PptxColorMap` through typed plot
  series, point overrides, markers, data labels, per-label overrides, leader lines, custom label runs, axis
  lines, gridlines, axis text overrides, and axis titles. The model regression now maps those surfaces through
  an overriding slide color map alongside chart-space/title/legend coverage. Keep this item open for chart
  color-style/style-part cascade semantics, direct renderer-only color paths, built-in table style aliases, and
  inherited placeholder style maps. Validation: `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal`
  passed; focused non-slow `pptx-model` passed (`17` passed, `0` failed, `1` skipped).
  2026-05-28 progress: chart color-style part parsing now resolves root palette entries and typed
  declaration records through the chart source `PptxColorMap`, so `colorsN.xml` scheme aliases such as
  `bg1`/`tx1` follow the owning slide/layout/master map instead of the default Office alias map. This is only
  source-map propagation for the already scene-owned color-style part; keep the item open for the actual
  chart-style/color-style cascade, variation-branch selection and precedence, direct renderer-only color paths,
  built-in table style aliases, and inherited placeholder style maps. Validation: `dotnet build
  Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed; focused non-slow `pptx-model` passed (`17` passed,
  `0` failed, `1` skipped); focused non-slow `pptx-charts` passed (`128` passed, `0` failed, `0` skipped).
  2026-05-28 progress: chart-style-part entries now receive the chart source `PptxColorMap` for role-local
  `spPr` fills/lines, `lnRef` theme-line resolution, `cs:defRPr/a:solidFill`, and `cs:fontRef` colors. The
  existing color-map scene regression now covers chart-style `defRPr` and `fontRef` colors plus role-local
  shape fill/line colors. Keep this item open for chart-style/color-style cascade ordering, variation-branch
  selection and precedence, direct renderer-only color paths, built-in table style aliases, and inherited
  placeholder style maps. Validation: `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed;
  focused non-slow `pptx-model` passed (`17` passed, `0` failed, `1` skipped); focused non-slow
  `pptx-charts` passed (`128` passed, `0` failed, `0` skipped).
  2026-05-28 progress: direct line parsing now uses a shared `PptxLineStyleReader` that accepts an explicit
  `PptxColorMap`; scene paths continue to call it with the owning source map, while direct renderer wrappers
  deliberately keep default-map behavior. This removes another copied line-width/color parser and makes the
  remaining direct-renderer color-map gap explicit instead of hidden behind duplicate code. Keep this item open
  for direct renderer-only color paths, built-in table style aliases, inherited placeholder style maps, and
  chart-style/color-style cascade ordering. Validation: `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v
  minimal` passed; focused non-slow `pptx-shapes` passed (`18 passed, 0 failed, 0 skipped`); focused non-slow
  `pptx-charts` passed (`136 passed, 0 failed, 0 skipped`); focused non-slow `pptx-model` passed
  (`25 passed, 0 failed, 1 skipped`); full non-slow console runner passed
  (`409 passed, 0 failed, 7 skipped`).
  2026-05-28 progress: built-in table-style fill accents now resolve through the table source
  `PptxColorMap` instead of the default alias map. The color-map scene regression now gives the table a
  supported built-in style whose accent is `tx1` and asserts that the resolved style fill follows the slide
  override to `accent6`; it also locks the current first-row text style color as a direct `lt1` theme slot
  rather than treating that absolute scheme component as a remapped placeholder alias. Keep this item open for
  direct renderer-only color paths, inherited placeholder style maps, and chart-style/color-style cascade
  ordering. Validation: `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed; focused non-slow
  `pptx-model` passed (`25 passed, 0 failed, 1 skipped`); focused non-slow `pptx-tables` passed
  (`10 passed, 0 failed, 0 skipped`); focused non-slow `pptx-charts` passed
  (`136 passed, 0 failed, 0 skipped`).
  2026-05-28 progress: direct PPTX text rendering now receives the source `PptxColorMap` instead of resolving
  text colors through renderer-local default aliases. The source map is threaded through slide/master/layout
  render dispatch, text layout/model inspection, table-cell text frames, run/default-run solid fills,
  hyperlink theme colors, text outlines, paragraph bullet colors, and shape `fontRef` fallback colors. The
  color-map regression now asserts both regular shape text and table-cell text through `InspectTextRuns`, so
  the renderer/model path is covered in addition to scene parsing. Keep this item open: inherited placeholder
  default text styles still pass around bare XML without the owning source map, and chart-style/color-style
  cascade ordering remains a separate Office-alignment problem. Validation: `dotnet build
  Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed; focused non-slow `pptx-model` passed
  (`25 passed, 0 failed, 1 skipped`); focused non-slow `pptx-typography` passed
  (`99 passed, 0 failed, 2 skipped`); focused non-slow `pptx-tables` passed
  (`10 passed, 0 failed, 0 skipped`); full non-slow console runner passed
  (`409 passed, 0 failed, 7 skipped`).
  2026-05-28 progress: `PptxSceneChart` now retains the effective `PptxColorMap` used when parsing the chart
  from its owning slide/layout/master source. The color-map regression asserts the chart-owned map alongside
  resolved chart colors, which gives future renderer fallback paths a structural owner instead of inferring
  chart color ownership from the current slide. Keep this item open: direct chart XML fallback helpers still
  need to consume `sceneChart.ColorMap`/source maps systematically rather than using default-map overloads.
  Validation: `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed; focused non-slow
  `pptx-model` passed (`25 passed, 0 failed, 1 skipped`).
  2026-05-28 progress: chart title and manual axis-title fallback rendering now consumes the chart/source
  `PptxColorMap` for direct XML shape styles, `txPr` text styles, rich title runs, and default `tx1` chart
  text color. This removes another default-alias path without changing chart layout formulas. Keep this item
  open for series/point/marker styles, legend/data-label fallback text, default axis-title fallback rendering,
  and the broader chart-style/color-style cascade ordering. Validation: `dotnet build
  Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed; focused non-slow `pptx-charts` passed
  (`136 passed, 0 failed, 0 skipped`).
  2026-05-28 progress: native chart rendering now carries the chart/source `PptxColorMap` through
  `TryRenderChart` into default axis-title fallback rendering, legend text style resolution, and the
  line/bubble right-legend reserve calculations that depend on legend text measurement. Scene chart text-style
  defaults also use `sceneChart.ColorMap`, eliminating another renderer-local default-alias path. Keep this
  item open for series/point/marker style fallback readers, data-label fallback text and shapes, direct
  XML axis-label text styles when no scene chart is available, and the broader chart-style/color-style cascade
  ordering. Validation: `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed; focused
  non-slow `pptx-charts` passed (`136 passed, 0 failed, 0 skipped`).
  2026-05-28 progress: native chart series, point, polar-point, doughnut-point, and marker fallback readers
  now receive the chart/source `PptxColorMap` for solid fills, pattern foreground/background colors, line
  colors, and marker shape styles. The renderer color wrapper now exposes the existing color-map-aware solid
  resolver, so chart style extraction no longer needs to bypass through default aliases for these paths. Keep
  this item open for data-label fallback text and shapes, direct XML axis-label text styles when no scene chart
  is available, and the broader chart-style/color-style cascade ordering. Validation: `dotnet build
  Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed; focused non-slow `pptx-charts` passed
  (`136 passed, 0 failed, 0 skipped`).
- [ ] Port `pptx-renderer` format-scheme fill/line resolution: `fillRef`, `lnRef`, style lists, `phClr`
  replacement, and default shape style resolution should be model-visible.
  2026-05-28 progress: shape `fillRef`/`lnRef` lookup now flows through `PptxFormatSchemeResolver` and a
  typed `PptxFormatSchemeReference` shared by the direct renderer and typed scene builder. This preserves the
  current fill/line drawing behavior while eliminating another duplicated style-reference lookup path. Keep
  this item open: the resolved format-scheme stages are still not fully model-visible, chart-style format
  references remain separately parsed, and default shape style/fallback ordering still needs Office-aligned
  provenance instead of renderer-local decisions. Validation: `dotnet build Lokad.OoxPdf.slnx --tl:off
  --nologo -v minimal` passed with zero warnings; focused non-slow `pptx-model` passed (`13 passed, 0 failed,
  1 skipped`) and `pptx-charts` passed (`121 passed, 0 failed, 0 skipped`).
  2026-05-28 progress: shape scene records now preserve resolved `fillRef`/`lnRef` provenance directly as
  `PptxFormatSchemeReference` values, including the raw reference node, parsed index, and resolved theme
  style when available. Scene inspection exposes private-safe fill/line reference index and resolved-state
  fields, and `ReadShape` now resolves each reference once before feeding shape fill/line construction. Keep
  this item open: chart-style format references still use their separate typed entry model, `effectRef` and
  `fontRef` are not structurally resolved through the same format-scheme stage, and default shape style/fallback
  ordering still needs an Office-aligned resolver rather than scattered renderer-local defaults. Validation:
  `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed; focused non-slow `pptx-shapes` passed
  (`18` passed, `0` failed, `0` skipped); focused non-slow `pptx-model` passed (`17` passed, `0` failed,
  `1` skipped).
- [ ] Port `pptx-renderer` text body handling: insets, anchors, vertical overflow, fit modes, wrapping,
  text direction, vertical text, multi-column text, and unsupported diagnostics where rendering is absent.
  2026-05-28 audit: shape text already has model-owned source-tagged `bodyPr` handling for insets, anchors,
  wrap, vertical overflow, columns, autofit children/scales, compatible line spacing, and text-body rotation,
  with later regression notes covering the direct/inherited/default cascade and layout consumption. Keep this
  item open, but narrow it: unsupported PPTX text diagnostics still scan raw slide XML instead of resolved
  text-frame/body-property model state, and table-cell text still has direct `bodyPr`/`tcPr` readers without
  the same source-tagged model/provenance surface as shape text. Those two gaps are the long-term structural
  work here, not another pass over already-closed scalar body-property inheritance.
  2026-05-28 progress: table-cell text insets now preserve source-tagged provenance in the scene model:
  each resolved edge records whether it came from table-cell `tcPr` margins, cell text-body `bodyPr` insets,
  or the Office default, plus the raw token when present. This keeps table text diagnostics from having to
  rediscover `bodyPr`/`tcPr` precedence in renderer code. Keep this item open for unsupported text diagnostics
  that still scan raw slide XML, and for converging full table-cell text body-property handling onto the same
  source-rich resolver used by shape text. Validation: `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v
  minimal` passed; focused non-slow `pptx-model` passed (`17` passed, `0` failed, `1` skipped); focused
  non-slow `pptx-tables` passed (`9` passed, `0` failed, `0` skipped).
- [ ] Port `pptx-renderer` run style behavior: font family, size, bold/italic, underline, strike, color,
  highlight, character spacing, kerning thresholds, caps, superscript/subscript, and hyperlink style.
  2026-05-28 audit: shape text already preserves and renders most run-style features through the text model,
  including raw underline/strike/caps tokens, highlight, character spacing, baseline shifts, small caps,
  per-run typefaces, and source-bearing theme typeface resolution in the direct text-model path. Keep this
  item open for the long-term architecture: chart text style parsing still has local run-property readers and
  only partially source-bearing typeface resolution, and the scene run-style path still does not expose the
  same diagnostics surface as direct text snapshots. 2026-05-28 progress: chart text style overrides now keep
  raw requested typeface tokens and `PptxThemeTypefaceSource` for direct `txPr`/rich runs and chart-style-part
  `fontRef` roles; rendering still consumes the flattened font family. Keep this item open: chart run-style
  parsing should move behind a shared cascade/theme-source resolver, and chart/table text should converge on
  the same run-style diagnostic surface as shape text instead of adding more per-consumer XML reads.
- [ ] Port `pptx-renderer` whitespace behavior: regular spaces, repeated spaces, non-breaking spaces,
  tabs, soft hyphens, explicit line breaks, fields, and end-paragraph runs must remain observable.
  2026-05-28 audit: regular wrapping, manual line breaks, fields, default tab stops, explicit tab stops,
  no-break spaces, narrow no-break spaces, hidden advances, leading spaces at style boundaries, and
  end-paragraph run sizing are already covered by shape-text model/layout tests. Keep this item open but
  narrower: soft-hyphen behavior still needs an explicit Office-backed fixture, and chart/table text should
  keep using the same observable whitespace segmentation instead of gaining surface-local parsers as those
  paths are converged onto shared text-frame layout.
- [ ] Port `pptx-renderer` bullet and numbering behavior: bullet suppression for metadata placeholders,
  `buChar`, `buAutoNum`, `buFont`, `buClr`, `buSz`, hanging indents, and inherited bullet defaults.
  2026-05-28 progress: shape-text bullet layout now consumes the resolved paragraph style cascade for
  marker text and bullet style, so `a:lstStyle`/inherited default bullet markers participate in layout
  instead of only locally declared paragraph `a:pPr` markers. A synthetic layout test locks inherited
  `lvl1pPr/a:buChar` emission. Keep this item open for metadata-placeholder bullet suppression and for
  avoiding future chart/table text bullet logic that bypasses the shared text-frame cascade.
- [ ] Port `pptx-renderer` table text behavior: cell text style inheritance, table style overrides,
  vertical alignment, margins, merged cells, and per-cell text diagnostics.
  - [ ] Replace `PptxTableStyleResolver`'s supported-style formulas with a real Office table-style cascade:
    parse table style parts/theme style matrices, conditional formatting priority, `phClr` replacement, and
    unsupported-style diagnostics instead of expanding GUID-specific logic.
    - [ ] Investigate the current `pptx-tables` visual gate drift instead of treating the table ladder as closed:
      `pptx-ladder-10-basic-table` is barely over its MAE gate and `pptx-ladder-10-vertical-align` has regressed
      from the recorded `0.013784` MAE to `0.028289`. The long-term fix should come from shared text-frame
      metrics for table cells and structural table draw-order alignment, not from widening gates or adding
      fixture-specific offsets.
- [ ] Port `pptx-renderer` shape geometry coverage: preset geometries, custom geometry, rotations, flips,
  group transforms, connectors, arrows, dash/cap/join, and picture-fill clipping.
  2026-05-28 audit: most source-bearing geometry inputs are already scene-owned: preset names and adjustment
  guides, custom geometry guides/paths/commands, group transforms, line dash/compound/cap/join tokens, and
  connector head/tail tokens are exposed through `PptxSceneShape` and locked by model/render tests. Keep this
  item open but narrower. The long-term work is now to make unsupported/partial geometry diagnostics consume
  the scene model instead of broad raw slide XML scans, to replace remaining Office-fit constants for connector
  end geometry with public Office-PDF evidence, and to expand preset/custom path rendering only through shared
  geometry records rather than per-preset narrow logic.
- [ ] Port `pptx-renderer` image behavior: relationship resolution, crop/fill/stretch, alpha/soft masks,
  SVG or unsupported-image diagnostics, media caching, and reuse across slides.
  - [ ] Remaining image work should stay structural: broaden image evidence with Office-authored fixtures for
    tiled fills, SVG limits, recolor/alpha interactions, and cross-slide media cache reuse before retiring this
    item.
- [ ] Port `pptx-renderer` chart behavior as a first-class native renderer: parse a typed chart model for
  series, axes, legends, labels, styles, and layouts; emit diagnostics only for unsupported chart features.
  - [ ] Extend chart data-label rendering to cover leader-line geometry, Office label-box geometry/auto-fit,
    and richer position semantics before attempting finer Office-aligned data-label layout.
  - [ ] Extend chart area and plot area style records to cover the remaining shape-style family instead of
    only direct solid fill/no-fill, pattern fill, and simple line.
  - [ ] Extend chart text-style records to cover the remaining rich-text surfaces, rotation, full non-default
    tick-label offset ladders, multi-level category labels, data-label box style consumption, and chart-style
    inherited defaults.
  - [ ] Continue consuming axis crossing/orientation metadata for remaining series coordinate baselines,
    label placement, and secondary-axis scale/geometry instead of relying on right-side XML/layout
    assumptions.
- [ ] Extend chart plot-area layout records to cover `layoutTarget`, x/y edge semantics, inner/outer plot
  semantics, title/legend overlay effects, and non-bar/line chart-family consumers in rendered geometry. The
  bar/line `ChartLayout` now carries a separate outer plot-area box and inner data-plot box, but both still
  resolve to the same rectangle until axis-label/title reservation rules are derived structurally.
- [ ] Keep SmartArt as a separate diagnostics-first feature until a real SmartArt renderer exists.
- [ ] Port `pptx-renderer` error isolation: one unsupported or malformed node should emit a diagnostic with
  slide/node context instead of aborting the whole render pass when recovery is possible.
- [ ] Port `pptx-renderer` generated public visual-suite organization: normalized case names, grouped
  fixtures, Office reference caching, and separate approximate, needs-review, and locked thresholds.
- [ ] Port `pptx-renderer` oracle tooling ideas that fit `.NET`: compact text-op diffs, visual metrics,
  cached Office references, deterministic artifact paths, and fast focused case selection.
- [ ] Port `pptx-renderer` performance lessons: avoid repeated ZIP/XML/theme/font parsing, cache immutable
  resources per render pass, and measure large-deck hot spots before private-deck tuning.
- [ ] Consider porting the testing strategy, especially generated Office-oracle case families and SSIM plus
  color-histogram metrics, while keeping `src/Lokad.OoxPdf` dependency-free.
## PPTX Renderer Test Port Plan

`pptx-renderer` has a much broader and better organized PPTX test corpus. Treat it as a test-design asset:
clean-port the cases and expected behaviors into `ooxpdf`; do not vendor TypeScript/Python code or private
assets.

Porting priorities:

- [ ] Port broader shape/preset oracle families next, preserving Office-authored/generated public fixtures
  and adding PDF/raster inspection notes for every accepted gate.
- [ ] Port layout-composition cases: master/layout placeholders, grouped transforms, z-order, image crop,
  table placement, native chart rendering, SmartArt fallback/diagnostics, and mixed slide content.
- [ ] Continue porting generated-case manifests and richer parity reporting for the remaining
  `pptx-renderer` oracle families.
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
- Chart color-style scene ownership now preserves more than the flattened palette: `PptxSceneChart.ColorStyle`
  records the resolved color-style part name plus the Office `meth` and `id` attributes alongside the resolved
  colors. This does not close inherited chart-style defaults; it removes a structural blind spot so future
  palette/default resolution can reason from typed scene data instead of reopening relationship parts in the
  renderer. Validation: focused `pptx-model` passed (`14 passed, 0 failed, 1 skipped`), focused
  `pptx-charts` passed (`40 passed, 0 failed, 0 skipped`), and the full non-slow suite passed
  (`244 passed, 0 failed, 7 skipped`).
- Chart style-part scene ownership is now explicit: `PptxSceneChart.StylePart` records the package-resolved
  `chartStyle` relationship target, the style-part `id`, and the loaded style XML. This still does not interpret
  Office chart-style defaults; it gives future default resolvers a single typed source instead of renderer-local
  relationship discovery. Validation: focused `pptx-model` passed (`14 passed, 0 failed, 1 skipped`), focused
  `pptx-charts` passed (`40 passed, 0 failed, 0 skipped`), and the full non-slow suite passed
  (`244 passed, 0 failed, 7 skipped`).
- Chart style-part scene ownership now includes line-reference role entries. The style part remains the source
  of truth for raw XML, but `PptxSceneChartStyle.Entries` decodes role-local `lnRef` indices and resolves their
  theme line styles into `PptxSceneLineStyle`, giving the future chart-style cascade a typed input instead of
  another renderer-local XML scan. This does not yet apply inherited defaults to gridlines or other chart roles;
  the unresolved work is the Office precedence/cascade model. Validation: focused `pptx-model` passed
  (`19 passed, 0 failed, 0 skipped`) and focused non-slow `pptx-charts` passed
  (`52 passed, 0 failed, 0 skipped`).
- Chart style-part role entries now also preserve direct `cs:spPr/a:ln` role lines as `ShapeLine`. This closes
  a structural gap discovered in local style parts where gridline roles carry the useful stroke in role-local
  shape properties while `lnRef` remains zero. Rendering still waits on a real chart-style cascade and precedence
  model. Validation: focused `pptx-model` passed (`19 passed, 0 failed, 0 skipped`) and focused non-slow
  `pptx-charts` passed (`52 passed, 0 failed, 0 skipped`).
- Chart axes now carry non-rendered chart-style gridline candidates. `PptxSceneChartAxis` exposes
  `MajorGridlineStyleLine` and `MinorGridlineStyleLine` from the `gridlineMajor`/`gridlineMinor` style-part
  roles, keeping inherited style evidence adjacent to direct gridline XML while avoiding a premature rendering
  behavior change. Validation: focused `pptx-model` passed (`19 passed, 0 failed, 0 skipped`) and focused
  non-slow `pptx-charts` passed (`52 passed, 0 failed, 0 skipped`).
- Chart style-part entries now preserve role text defaults from `fontRef` and `defRPr` as
  `PptxSceneChartStyleEntry.TextStyle`. This gives the future chart-style text cascade typed defaults for
  titles, axes, legends, and labels without reopening the style XML in the renderer. Validation: focused
  `pptx-model` passed (`19 passed, 0 failed, 0 skipped`) and focused non-slow `pptx-charts` passed
  (`52 passed, 0 failed, 0 skipped`).
- Chart external-data scene ownership now preserves workbook provenance: `PptxSceneChart.ExternalData` records
  `c:externalData/@r:id`, the package-resolved embedded-workbook target, and the `autoUpdate` flag. Workbook
  parsing still remains a renderer-side bridge and does not yet provide stale-cache reconciliation, blank-cell
  policy, table/name support, or typed value vectors in the scene. Validation: focused `pptx-model` passed
  (`14 passed, 0 failed, 1 skipped`), focused `pptx-charts` passed (`40 passed, 0 failed, 0 skipped`), and the
  full non-slow suite passed (`244 passed, 0 failed, 7 skipped`).
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
  Explicitly styled line markers without `c:size` now use a distinct Office-observed default envelope shared by
  the scene and renderer. Remaining marker work is Office's full marker preset set, automatic marker
  inheritance from chart style/color-style parts, and explicit marker-size unit/envelope conversion.
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
  `PptxSceneChartTitle` preserves title overlay/manual layout and shape styling. Rendering now consumes
  explicit title fill/stroke in the currently computed title rectangle, but still needs Office-PDF evidence
  before changing exact placement, overlay reserve, or title box sizing. Remaining title work is exact Office
  title layout and inherited chart style defaults.
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

- [ ] Review `pptx-renderer`'s generated oracle and report reuse strategy for ideas that fit the .NET
  dependency-free constraint.

## Progress

- [ ] 2026-05-25: Continue chart-structure classification toward legend swatches, data-label text positions,
  and polar/radar shape semantics. The new derived candidates improve the structural oracle surface, but they
  still do not classify chart text matrices, legend entries, leader lines, or Office radar polygon strokes as
  first-class chart structures.
  - [ ] Promote radar geometry into a typed radar layout model instead of leaving the filled-vs-marker split as
    renderer metric constants. Preserve the current path-geometry gates as the oracle while deriving center,
    radius, label frame, and plot-box reserves from chart layout context.
- [ ] 2026-05-25: Continue chart text classification beyond the first semantic roles. Remaining gaps include
  robust chart-title disambiguation, top/bottom legend containers, data labels outside the plot box,
  annotations, and multi-chart pages. Value-axis origin labels are now classified structurally, but value-axis
  gates should wait until candidate tick generation no longer emits extra tick labels against Office.
- [ ] 2026-05-25: Complete the chart scene model so chart kinds, plot areas, axes, series, data labels,
  markers, title, legend, fills, strokes, and text styles are represented as typed data before PDF emission.
  - [x] 2026-05-31: Add chart text character spacing to the typed scene and renderer style pipeline.
    `a:defRPr/@spc` and rich text run `a:rPr/@spc` now survive chart scene parsing, style merging, chart text
    measuring, and PDF text-run emission. This closes an architectural omission even though the latest private
    slide-44 run showed no movement: the slide-44 chart/style XML has no nonzero authored `spc`, while Office
    still emits small chart-text `Tc` values in PDF. That remaining gap belongs to the shared Office PDF
    text-emission profile, not to chart geometry or a private chart offset.
- [ ] 2026-05-25: Replace chart fallback geometry by turning each named `PptxChartMetricRules`
  approximation into an Office-PDF-observed rule or an explicitly classified temporary gap with a public
  visual case.
  - [x] 2026-05-31: Replace the remaining stacked-column per-rectangle fill emission with Office-like
    compound series paths. Private slide 44 and public stacked-column cases showed Office emitting each
    stacked series as a compound filled path (`16` segments / `4` moves for four categories). OOXPDF now emits
    stacked column/bar fills as series-major compound paths for simple fills, with patterned fills still using
    the existing per-rectangle pattern fallback. Public stacked-column and 100% stacked-column fixtures now
    gate `FilledRegion` operators, segment counts, and path-command counts.
  - [ ] Bubble chart layout still uses Office-observed title/right-legend plot-box and bubble headroom
    constants, now including separate bubble plot-width and right-legend swatch placement ratios. Keep these
    as explicit temporary `PptxChartMetricRules` inventory until chart structural oracle tooling can derive
    plot box, legend reserve, and bubble extent from Office PDF structures across more public bubble variants.
    Also close the remaining PDF paint-rule gap: Office emits bubble fills as `f*` while the candidate emits
    `f`, so the current marker guard verifies geometry/path shape but not the exact fill operator.
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
- [ ] Continue typed picture migration by moving any remaining SVG-specific paint/path decisions that belong
  to OOXML interpretation into `PptxScenePicture` or a dedicated SVG picture model.
- [ ] Continue reducing renderer XML fallbacks by adding non-linear/path gradient support, gradient alpha
  handling, richer scene-owned effect families, and a JPEG recolor strategy without format-specific shortcuts.
- [ ] JPEG recolor strategy: the remaining private warning is a three-component JPEG with duotone plus alpha.
  Do not special-case the private slide. Choose between a dependency-free JPEG pixel decoder, a principled PDF
  color-transform approach that matches Office output, or an explicit documented limitation with a public
  Office-authored rung that keeps the diagnostic stable.
  2026-05-28 audit: `PptxRenderer.Images.CreateImage` still emits `PPTX_UNSUPPORTED_IMAGE_RECOLOR` before
  returning `PdfImageXObject.Jpeg(...)` for JPEG inputs, while PNG/BMP recolor routes through decoded RGB
  buffers and `ApplyImageRecolor`. `JpegInfo` is deliberately only a header reader, so closing the private
  warning requires either a real JPEG pixel decoder or a structurally equivalent PDF colorization path; a
  per-document or content-type branch would reintroduce the hard-coded logic this plan is eliminating.
  2026-05-28 public limitation rung: `CheckVisualCase.ps1` now supports
  `expected.requiredDiagnostics`, and `pptx-ladder-07-jpeg-duotone-recolor-diagnostic` locks the current JPEG
  duotone-plus-alpha fallback as an explicit public visual case instead of a private-only warning. The fixture
  is derived from the valid public JPEG image deck with a narrow `a:duotone`/`a:alphaModFix` blip effect, and
  the visual run `20260528-120756` passed with the required `PPTX_UNSUPPORTED_IMAGE_RECOLOR` diagnostic,
  MAE `0.582500`, changed16 `0.020000`, and SSIM `0.992915`. This does not close the strategy item: the
  long-term target remains decoded-pixel or PDF-structural recolor, not permanent JPEG passthrough.
  2026-05-28 backend audit: `PdfImageXObject` and `PdfDocumentWriter` already support soft masks and raw
  PDF color-space strings, but the exposed image constructors are RGB Flate streams and DCT passthrough.
  A three-component JPEG duotone cannot be made Office-equivalent by wrapping the compressed RGB DCT stream in
  a different color space, because Office's current decoded-RGB recolor path derives luma per pixel before
  mapping dark/light duotone colors. The long-term implementation path should therefore be a real
  dependency-free JPEG decoder boundary or an explicitly proven PDF-level transform over decoded samples, not
  a renderer-local color-space shortcut. Keep the public limitation rung as the guardrail until that backend
  exists.
  2026-05-28 update: made `JpegInfo` classify SOF frame profiles (`baseline DCT`, `progressive DCT`, and the
  other supported header markers) and threaded that profile into the unsupported JPEG recolor diagnostic. This
  does not decode pixels or remove the warning, but it moves the future decision point to the image backend:
  a baseline-only decoder can now be introduced explicitly without conflating progressive/lossless JPEGs with
  the common baseline path. Validation: focused `imaging` passed (`9 passed, 0 failed, 0 skipped`); focused
  `pptx-images` initially hit a parallel build file lock and then passed serially (`17 passed, 0 failed,
  0 skipped`); full non-slow runner passed (`371 passed, 0 failed, 7 skipped`). Private run
  `20260528-153316` remained stable at 84/84 compared pages, zero dimension mismatches, deck MAE `6.715278`,
  changed16 `0.093542`, and the single `PPTX_UNSUPPORTED_IMAGE_RECOLOR` diagnostic.
  2026-05-29 investigation: added a dependency-free PDF luminosity soft-mask primitive to the PDF writer, then
  tested using it as a structural JPEG duotone route without decoding JPEG pixels. That activation was rejected
  and left disabled: the public JPEG-duotone rung rendered without diagnostics but worsened from the guarded
  fallback run `20260528-120756` (MAE `0.582500`, changed16 `0.020000`, SSIM `0.992915`) to experimental run
  `20260529-095717` (MAE `1.000000`, changed16 `0.020000`, SSIM `0.916703`), and the private deck page carrying
  the remaining JPEG recolor gap worsened from page MAE `14.28`, changed16 `0.46`, SSIM `0.62` to page MAE
  `89.26`, changed16 `0.98`, SSIM `0.06` in experimental run `20260529-095948`. Office inspection of the public
  reference showed materialized decoded RGB plus a grayscale `/SMask`, not a luminosity mask over the original
  JPEG stream. Keep `PPTX_UNSUPPORTED_IMAGE_RECOLOR` for JPEG/DCT recolor until OOXPDF can materialize decoded
  image samples or prove an Office-equivalent PDF structure on both the public rung and the private acceptance
  deck. After disabling the activation, public run `20260529-100154` returned to the guarded fallback metrics
  (MAE `0.582500`, changed16 `0.020000`, SSIM `0.992915` with the required diagnostic), and private run
  `20260529-100255` returned to 84/84 compared pages, zero dimension mismatches, deck MAE `7.167255`,
  changed16 `0.098107`, the two text-overflow ellipsis diagnostics, and the one JPEG baseline-DCT recolor
  diagnostic. Private page 84 returned to MAE `14.28`, changed16 `0.46`, SSIM `0.62`.
- [ ] Trim this ExecPlan conservatively: first add missing `PLANS.md`-required sections and current evidence,
  then consolidate only completed historical detail that is already represented by checked-in fixtures,
  tests, or tool support. Do not remove open checkboxes during this cleanup unless a direct duplicate is
  found and noted.
- [ ] 2026-05-31: Pivot the active architecture track to DOCX after the slide-44 PPTX cleanup. Build a DOCX
  model/pipeline comparable in spirit to PPTX: typed document parts and relationships first, explicit style
  cascade, block layout/pagination as a separate stage, then PDF emission. Use public synthetic Word/Office
  reference cases as the primary oracle and private DOCX files only for public-safe gap discovery.
  - [x] 2026-05-31: Added `DocxLayoutEngine` and layout records for pages, text lines, inline images, and
    table rows/cells, then changed `DocxRenderer` to emit PDF from that positioned layout instead of mixing
    pagination and drawing in the same loop. This is intentionally behavior-preserving and is the first DOCX
    pipeline boundary needed before improving Word-compatible pagination and table layout.
  - [ ] 2026-05-31: Continue moving DOCX toward typed intermediate stages: split style-resolved paragraph/table
    content from raw document reading, promote paragraph spacing/keep decisions into the block-pagination
    stage, then enrich private-safe layout traces so pagination drift can be diagnosed without inspecting
    private text.
  - [x] 2026-05-31: Added a private-safe `DocxLayoutSnapshot` inspection surface over the layout model. It
    reports page dimensions, item counts, item kinds, bounds, cell counts, and text lengths, but not document
    text. This gives private DOCX pagination/table work a restartable trace target before adding Word layout
    rules.
  - [x] 2026-06-01: Added a private-safe `DocxStructureSnapshot` pre-layout inspection surface in the library,
    exposed through `DocxRenderer.InspectStructure` and `tools/Lokad.OoxPdf.DocxInspect` as
    `structure-snapshot.json`. This is the DOCX analogue of the PPTX scene/snapshot discipline at the document
    block boundary: it records stable block indices, neighboring block kinds, paragraph spacing/keep/list/tab
    facts, page-break and section-break tokens, table width/style/layout facts, row/cell feature counts,
    vertical merges, grid spans, shading, borders, and table adjacency without emitting document text. Public
    unit coverage `DocxStructureSnapshotReportsPreLayoutBlockAndTableFacts` locks the bottom-up contract.
    Validation passed `docx-tables --skip-slow` (`80` after the new test) and `docx-core --skip-slow` (`25`);
    one parallel dotnet run hit the known compiler output lock and the serial rerun passed.
    2026-06-01 follow-up: extended the same pre-layout structure snapshot with DOCX story ownership. It now
    inventories the body story, document-scope header/footer variants, and section-scope static stories tied to
    the source section-break block index. This moves DOCX closer to the proven PPTX scene discipline where
    slide/layout/master ownership is explicit before rendering. The same public unit test now checks body,
    document default header, document even footer, and section first-header story facts. Validation passed
    `docx-tables --skip-slow` (`80`) and full solution build.
    2026-06-01 follow-up: extended the pre-layout structure snapshot with floating drawing ownership:
    `FloatingDrawingCount` and per-drawing wrap, relative positioning, extent, distance, overlap, behind-doc,
    and layout-in-cell tokens. This is intentionally diagnostic/structural only; anchored drawing rendering
    remains open, but future DOCX drawing work now has a bottom-up feature inventory instead of inspecting raw
    XML or private content. The same public unit test covers a square-wrap, column/paragraph-relative drawing.
    Validation passed full solution build and serial `docx-tables --skip-slow` (`80`); the first parallel test
    attempt hit the known compiler output lock.
    2026-06-01 follow-up: `tools/Lokad.OoxPdf.DocxInspect` now writes `block-sequence.json` and
    `table-adjacency-summary.json` from `DocxStructureSnapshot` instead of maintaining duplicate tool-local
    parsers and record types. This keeps private-safe DOCX diagnostics aligned with the library model boundary.
    Validation passed full solution build, serial `docx-tables --skip-slow` (`80`), and a public
    `InspectDocx` run that produced `structure-snapshot.json`, `block-sequence.json`, and
    `table-adjacency-summary.json`.
    2026-06-01 follow-up: extended `DocxStructureSnapshot` tables with per-row and per-cell private-safe
    profiles: row index, logical grid span, header/cantSplit/height facts, table-property exceptions, grid
    spans, vertical merges, shading, preferred widths, visible border counts, paragraph/run/text/image counts,
    numbering/keep counts, spacing-token counts, max font size, and character-class counts. This is the
    bottom-up structural surface needed before revisiting row fragmentation. Public unit coverage locks these
    fields. Validation passed serial `docx-tables --skip-slow` (`80`), full solution build, and a private-safe
    inventory run. The private aggregate has `129` table rows and `422` cells; all `129` rows have before/after
    spacing tokens, `67` have shading, `13` have visible borders, and none have `cantSplit` or declared row
    heights. This confirms paragraph spacing tokens alone cannot be the missing row-fragment discriminator.
    2026-06-01 follow-up: added private-safe style/list usage buckets to `DocxStructureSnapshot`, covering
    paragraph styles, table styles, and numbering/list labels by number id, level, format, and suffix without
    emitting text. Public unit coverage checks paragraph, table-cell paragraph, table-style, and list-label
    buckets. Validation passed serial `docx-tables --skip-slow` (`80`), full solution build, and a private-safe
    inventory run. The current private DOCX is driven primarily by paragraph styles `Compact` (`472`
    paragraphs) and `BodyText` (`76`), table styles `PlainTable5` (`12` tables) and `Table` (`1`), plus nine
    level-0 bullet list ids using tab suffixes; it has no floating drawings or inline images. Use these buckets
    to prioritize public probes for compact paragraph rhythm, bullet/tab ownership, and simple table flow before
    anchored drawing or merge-heavy table rendering.
    2026-06-01 follow-up: table layout snapshots now carry the table's source body block index on both
    `DocxTableSnapshot` and `DocxTableRowSnapshot`. This connects the pre-layout structure snapshot to the
    rendered layout trace without relying on table order, which is necessary before comparing representative
    private tables against public probes. Public unit coverage locks a paragraph/table/paragraph/table body
    stream with source block indexes `1` and `3`. Validation passed `docx-tables --skip-slow` (`81`) and full
    solution build.
    2026-06-01 follow-up: the generic per-page `DocxLayoutItemSnapshot` stream now also reports
    `SourceBlockIndex` for `TableRow` items, so page-level drift summaries can line up paragraph text lines and
    table rows through the same private-safe ownership field. The same public test covers the item stream.
    Validation passed `docx-tables --skip-slow` (`81`) and full solution build.
    2026-06-02 follow-up: `DocxLayoutPageSnapshot` now summarizes source-block ownership per page with
    distinct source block count plus first/last source block indexes. This keeps private pagination drift
    inspection at page scope without exposing text or requiring consumers to scan every item. Public coverage
    checks a paragraph/table/paragraph/table stream on one page; validation passed `docx-core --skip-slow`
    (`29`).
    2026-06-02 follow-up: `DocxLayoutSnapshot` now also includes source-block summaries grouped across pages:
    first/last page index, item counts, text-line/table-row/inline-image counts, text length, consumed height,
    and applied before-spacing sums. This is the layout-side counterpart to `DocxStructureSnapshot` block
    ownership and gives future private-safe pagination comparisons a per-block join key without exposing
    content. Public coverage checks paragraph and table source-block summaries; validation passed
    `docx-core --skip-slow` (`29`).
    2026-06-01 follow-up: extended paragraph and table row/cell structure profiles with private-safe
    whitespace-delimited token counts and longest-token lengths. This fills a bottom-up diagnostic gap for
    compact paragraph and narrow table-cell wrapping without exposing text. Public structure tests cover body
    paragraphs plus table, row, and cell aggregates. Validation passed `docx-core --skip-slow` (`25`),
    `docx-tables --skip-slow` (`81`), full solution build, and a private-safe inspect run. The current private
    DOCX has `198` paragraph blocks (`185` with tokens), max body token length `27`, `4` paragraph blocks with
    token length `>=20`, `422` table cells all with tokens, max cell token length `22`, and `5` cells with token
    length `>=20`; future public probes should cover both paragraph and narrow-cell long-token wrapping.
    2026-06-01 follow-up: added public `docx-ladder-02-long-token-wrapping` to cover that token-risk shape with
    a constrained Calibri body paragraph plus a narrow fixed table. Public-safe inspection reports `2` paragraph
    blocks, max body token length `29`, one body paragraph with token length `>=20`, one table with `3` cells,
    max cell token length `24`, and two cells with token length `>=20`. Visual run `20260602-002745` completed
    with `1/1` page, matching dimensions, no diagnostics, `MAE=1.282270`, and changed16 `0.018130`; keep it as
    a bottom-up Office-backed probe for paragraph and table-cell word-break behavior before changing wrapping
    logic.
  - [x] 2026-05-31: Preserved DOCX header/footer reference types (`default`, `first`, `even`) instead of
    flattening every referenced part into one static paragraph list. Static header/footer rendering now selects
    the first-page or even-page part only when the corresponding Word settings are active, otherwise it uses
    the default part. Private-safe inventory found the current private DOCX carries three header references;
    this closes the structural model gap without using header text or document-specific logic.
    2026-06-01 follow-up: section-local `w:headerReference` and `w:footerReference` maps now travel with
    `DocxPageSettings`, so pages render static content from their owning section rather than from the
    document-wide last reference for a type. `DocxFontPlan` also includes section-scoped static runs, preventing
    section-only header/footer typefaces from falling back to a resource planned from another section.
    2026-06-01 follow-up: layout now builds an explicit effective static-part chain for section settings:
    omitted section header/footer references inherit previously defined references by type, while the renderer
    no longer treats an empty section map as permission to backfill from document-global references. Public
    coverage guards the important non-backfill edge where the first section omits static references and the
    final section owns them.
  - [x] 2026-05-31: Preserved DOCX `w:pgMar/@w:header` and `w:pgMar/@w:footer` distances and used those
    authored page-margin tokens for static header/footer baselines. This removes the prior half-margin
    placement fallback for documents that provide Word's header/footer distances, with public coverage that
    checks both token preservation and the emitted PDF text matrices. Private DOCX run `20260531-233331`
    stayed neutral at `16/16` pages, zero dimension mismatches, MAE `13.648284`, changed16 `0.125542`.
  - [x] 2026-05-31: Preserved parsed paragraph models inside DOCX table cells while keeping the existing
    flattened cell text for current rendering. This closes a model gap that blocked future row-height,
    paragraph spacing, inherited paragraph/character styling, per-run styling, and numbering layout inside cells.
  - [x] 2026-05-31: Moved DOCX table-cell text emission onto layout-owned cell text lines. The renderer now
    draws `DocxTextLineLayout` records attached to each `DocxTableCellLayout`, so multi-paragraph cells,
    paragraph alignment, font size, color, bold/italic, underline, and legacy flattened-cell compatibility are
    represented before PDF emission instead of being recomputed inside `RenderTableRow`.
  - [x] 2026-05-31: Added styled text segments to DOCX text lines and used them for table-cell paragraphs.
    Mixed runs inside one cell paragraph now keep separate segment text, style, color, and x-position before
    PDF emission instead of collapsing the paragraph to the first run's style.
  - [x] 2026-05-31: Preserved DOCX `w:tcMar` table-cell margin tokens and used resolved `dxa` margins to
    define the table-cell text box. Authored left/right margins now affect text line width and x-position, and
    authored top margins move the first baseline before PDF emission.
    2026-06-01 follow-up: replaced the zero/default horizontal table-cell padding assumption with Word's
    missing-`w:tcMar` horizontal inset (`5.4pt`) while keeping missing top/bottom padding at `0pt`. The new
    `docx-ladder-03-table-row-heights` fixture exposed this from Office's text matrices: reference cell text
    starts around `5.66pt` inside the cell, while the candidate had been flush to the left border. Public
    `docx-ladder-02-table-explicit-font` improved from the prior `MAE=0.229378`, changed16 `0.003197` to
    run `20260601-032007` at `MAE=0.211186`, changed16 `0.003175`; the row-height ladder shifted x closer but
    stayed dominated by vertical-baseline drift (`MAE=2.435244`, changed16 `0.018702`). Private DOCX run
    `20260601-032030` stayed unchanged at `16/16` pages, zero dimension mismatches, no diagnostics,
    `MAE=13.047852`, changed16 `0.120850`.
  - [x] 2026-05-31: Applied DOCX table-cell vertical alignment in layout. `w:vAlign` `center` and `bottom`
    now shift the layout-owned cell text block inside the row after paragraph wrapping/spacing and margin
    resolution, keeping the geometry visible to layout inspection before PDF emission.
  - [x] 2026-05-31: Made DOCX table rows expand to measured cell text content. Row layout now computes cell
    text block height from the resolved text box, wrapping, paragraph line heights, spacing, and margins, then
    uses the maximum of declared/default row height and content height before placing cells or checking page
    breaks.
  - [x] 2026-05-31: Added public coverage for numbered paragraphs inside DOCX table cells. Table-cell
    paragraph parsing now has an explicit guard that the document numbering context is applied inside cells,
    and layout verifies that the list label is consumed into the cell text line before PDF emission.
  - [x] 2026-05-31: Promoted DOCX table-cell rendering from flattened cell text to the preserved paragraph
    model for current supported cell content. Table-cell layout/rendering now covers paragraph boundaries,
    inherited styling, mixed runs, margins, vertical alignment, content-driven row height, numbering, and inline
    images through layout-owned text/image records before PDF emission.
  - [x] 2026-05-31: Tightened DOCX table row content-height growth to measure actual text/image content,
    margins, and spacing rather than treating the legacy baseline placement inset as required row height. This
    keeps ordinary one-line default rows compact while still expanding rows that wrap or contain taller content.
  - [x] 2026-05-31: Preserved DOCX paragraph spacing and keep-rule source tokens in the paragraph model after
    style/default/direct-property resolution. `DocxParagraph` now carries its `w:pStyle`, raw `w:spacing`
    tokens including line/autospacing/line-count variants, `w:contextualSpacing`, and keep/widow on/off
    values. This intentionally does not change pagination yet; it removes the reader-local XML dependency so
    the next Word-like pagination slice can use typed structural facts instead of diagnostics-only heuristics.
  - [x] 2026-05-31: Applied typed DOCX `keepLines` and `keepNext` facts during block pagination. The layout
    stage now estimates kept paragraph blocks before drawing lines and moves a paragraph, or a paragraph plus
    its next paragraph/table target, to the next page when the kept block would otherwise split at the page
    bottom. This is a generic Word-compatible pagination step backed by public synthetic tests, not a
    private-document page-count rule. `widowControl` remains open because it needs line-level widow/orphan
    decisions rather than whole-block movement.
  - [ ] 2026-05-31: Finish DOCX paragraph pagination rules: implement `widowControl` with line-level orphan
    checks, chain consecutive `keepNext` paragraphs across multiple following blocks, and refine exact
    keep-with-table behavior against public Word/Office PDF fixtures before downgrading the keep-rule
    diagnostics.
    - [x] 2026-05-31: Chained consecutive DOCX `keepNext` paragraphs during kept-block estimation instead of
      stopping after the first following paragraph/table target. Public synthetic coverage now exercises a
      three-paragraph keep chain that only fits after moving the whole chain to the next page. Private DOCX
      run `20260531-184720` stayed dimension-stable at 16/16 pages and improved slightly versus the current
      DOCX baseline (`15.849350` MAE, changed16 `0.141574`). `widowControl` and exact keep-with-table
      behavior remain open.
    - [x] 2026-05-31: Added first-pass DOCX `widowControl` pagination for multi-line body paragraphs. Before
      drawing a paragraph, layout now starts it on the next page when the current page would fit only one
      paragraph line or would leave only one paragraph line for the next page. Public coverage checks a
      three-line paragraph that previously split 2/1 across pages; private DOCX run `20260531-185646` stayed
      page- and metric-neutral (`15.849350` MAE, changed16 `0.141574`). Keep-with-table behavior and
      Office-backed edge cases remain open before downgrading the keep-rule diagnostic.
    - [x] 2026-05-31: Aligned DOCX keep-with-table estimation with the actual table layout width for indented
      tables. Kept-block estimation now subtracts nonnegative `w:tblInd` before resolving table preferred
      width and first-row wrapping, matching the layout stage instead of underestimating indented table targets.
      Public pagination coverage checks that a `keepNext` paragraph moves with an indented following table
      whose first row wraps taller after the indent is applied. Private DOCX run `20260531-224646` stayed
      page- and metric-neutral (`16/16`, MAE `14.791805`, changed16 `0.133893`). The parent item remains open
      for Office-backed keep-with-table edge cases and exact diagnostic narrowing.
  - [x] 2026-06-01: Started section-owned page geometry in DOCX layout. Paragraph-level `sectPr` already
    preserved page settings; layout now uses the upcoming section break's page size and margins for the
    preceding section, then switches to the following section after a `nextPage`/`oddPage`/`evenPage` break.
    This fixes the ownership direction without applying prior-section settings to following pages. Public
    layout coverage checks that the first section uses the paragraph-section page settings and the next page
    returns to the final body section settings. Keep `DOCX_UNSUPPORTED_SECTION_BREAK` open for continuous
    section geometry, header/footer selection per section, odd/even blank pages, and columns. Validation
    passed `docx-page --skip-slow` (`24`), `docx-tables --skip-slow` (`65`), and full solution build. Private
    DOCX run `20260601-103219` stayed neutral at `16/16` pages, zero dimension mismatches, no diagnostics,
    `MAE=13.388935`, changed16 `0.124264`.
  - [x] 2026-06-01: Extended section-owned geometry through static header/footer placement. `DocxLayoutPage`
    now carries the section page margins and `DocxPageSettings` that owned the body layout, and static
    header/footer rendering uses those page-local margins, body width, page height, and header/footer
    distances instead of the document-level final section. Public coverage checks both the layout page
    metadata and emitted PDF text matrices for a two-section document. This is structural page ownership,
    not section-specific header/footer relationship resolution: keep that open for true per-section selected
    parts, odd/even inserted blank pages, continuous-section geometry, and columns. Validation passed
    `docx-page --skip-slow` (`25`), `docx-text --skip-slow` (`36`), and full solution build; public
    `docx-headers-footers` improved to run `20260601-135640` (`MAE=0.073352`, changed16 `0.002110`). Private
    DOCX run `20260601-135549` stayed page- and diagnostic-stable (`16/16`, zero dimension mismatches, no
    diagnostics) but moved to `MAE=13.838763`, changed16 `0.126851`, so keep private aggregate monitoring
    active when the next section/header slice lands.
  - [x] 2026-06-01: Replaced document-global static-content fallback with layout-owned effective section
    static content. `DocxLayoutEngine` resolves section settings in document order, merges inherited
    header/footer references by type, and applies the final body `sectPr` over the inherited chain. Static
    header/footer rendering now consumes only `layoutPage.PageSettings`, which prevents a later/final section
    header or footer from appearing on an earlier section that omitted the reference. This is still not a full
    section model: continuous sections, Word's odd/even blank-page insertion, columns, and any future explicit
    "link to previous" surface remain open. Validation passed `docx-page --skip-slow` (`26`), `docx-text
    --skip-slow` (`36`), and full solution build; public `docx-headers-footers` run `20260601-141617` stayed
    at `MAE=0.073352`, changed16 `0.002110`, SSIM `0.982572`. Private DOCX run `20260601-141624` stayed
    unchanged at `16/16` pages, zero dimension mismatches, no diagnostics, `MAE=13.838763`, changed16
    `0.126851`.
  - [ ] 2026-06-01: Separate DOCX table-style paragraph-property precedence from table-cell text rendering.
    Public Office inspection confirms table-style `w:rPr` must sit below paragraph-style and character-style
    run properties; that run-property precedence belongs in the active table-text slice. A broader table-style
    `w:pPr` precedence trial changed private pagination/page count, so paragraph-property precedence needs its
    own public Word/Office PDF pagination fixture before changing the production cascade.
    Follow-up: a narrower run-only precedence trial also changed the private candidate from 16 pages to 15
    (`20260601-013756`, MAE `13.645890`, changed16 `0.124541`), despite passing public `docx-tables` and
    `docx-text`. Keep this open until a public Word-backed table-style text precedence fixture explains the
    pagination interaction; do not merge a theoretically cleaner cascade without that guard.
    Follow-up: added public `docx-ladder-03-table-style-text-precedence` as that guard. Office renders the
    first-column conflict text in all caps but at the higher-priority `14.04pt` paragraph/character size,
    while the candidate keeps table-style `11pt`; page count is stable and run `20260601-014145` reports
    `MAE=0.317789`, changed16 `0.003980`. This confirms the cascade order problem is real while preserving the
    private-page-count rejection as the implementation constraint.
    2026-06-01 follow-up: added public `docx-ladder-03-table-style-conditional-run-precedence` to split base
    table-style `w:rPr` from conditional `w:tblStylePr/w:rPr` behavior. Office renders first-row and
    first-column conflict text at the paragraph/character `15pt` size even when conditional table-style regions
    declare smaller sizes; non-conflicting caps/bold/italic still apply. The current candidate keeps the
    table-style `9/10/11pt` sizes and run `20260601-015815` reports `MAE=0.387931`, changed16 `0.004464`.
    A trial that moved table-style run properties below paragraph/character styles improved the earlier
    precedence case (`MAE=0.262365`) but again collapsed the private candidate from 16 pages to 15
    (`20260601-014605`, `MAE=13.645890`, changed16 `0.124541`). Private-safe DOCX layout inspection now shows
    the loss is concentrated in eight first table rows whose max font size drops from `13pt` to `10pt`; most
    shrink from `20.05pt` to the current default `16pt`, with one wrapped first row shrinking by `22.30pt`.
    Do not land the correct run cascade until the table-row vertical composition gap below is fixed.
    2026-06-01 progress: moved table-style run properties below paragraph/character styles after adding the
    row-minimum fix below. Public `docx-ladder-03-table-style-text-precedence` improved to `MAE=0.262365`,
    changed16 `0.003536`, and the private DOCX stayed accepted at `16/16` pages with zero dimension
    mismatches while improving to `MAE=12.716572`, changed16 `0.118448`.
  - [ ] 2026-06-01: Fix DOCX table-row vertical composition so correct table-style run precedence does not
    rely on oversized table-style fonts to preserve page count. The accepted private baseline currently keeps
    16 pages partly because table-style font-size precedence is wrong; after the Office-confirmed cascade fix,
    first rows become smaller and expose missing Word row-height/baseline semantics. Investigate public
    Office-backed probes for default table row minimums, line-height plus paragraph spacing inside cells, and
    first-row/header-row spacing before retrying the run cascade. The new private-safe `tools/InspectDocx.ps1`
    layout snapshot includes per-row and per-cell max font sizes to diagnose this without private text.
    2026-06-01 progress: introduced `WordDefaultTableRowMinimumTwips = 401` as the generic auto-row floor used
    when no explicit `w:trHeight` owns the row. This restores the private DOCX to 16 candidate pages after the
    correct run cascade and improves aggregate metrics, but keep the item open for a broader Office-authored
    ladder around empty rows, multi-line rows, explicit `auto/atLeast/exact` heights, and first-row/header-row
    baseline placement before treating the 401-twip default as fully explained.
    2026-06-01 follow-up: kept-block preflight now moves a `keepNext` block when the estimated block lands
    exactly on the bottom margin, not only when it crosses below it. This keeps the preflight consistent with
    subsequent row placement after the 401-twip row minimum and avoids splitting an exactly-fitted
    paragraph/table keep pair across pages.
    2026-06-01 follow-up: added Office-authored public fixture `docx-ladder-03-table-row-heights` to make the
    row-height evidence restartable. Word writes the at-least row as `w:trHeight w:val="360"` without
    `hRule`, the exact row as `w:hRule="exact" w:val="360"`, and omits row-height tokens for natural auto
    rows. Public run `20260601-031548` is page- and dimension-stable (`MAE=2.409006`, changed16 `0.018633`);
    layout inspection shows the candidate matching the basic row-height classes (`21.8678pt` natural rows,
    `35.7355pt` at-least wrapped row, `18pt` exact row). Keep the parent open for first-row/header-row
    baseline placement, row clipping semantics, and broader pagination cases rather than the basic
    `auto`/missing/`exact` token distinction.
    2026-06-01 follow-up: removed the default whole-table preflight from `LayoutTable`; ordinary
    DOCX tables now paginate by row unless an explicit keep/header rule says otherwise, instead of
    treating every table as an indivisible block. A measured-whole-table preflight trial was rejected
    because it produced a 17-page private candidate with a dimension mismatch (`20260601-075711`),
    proving that whole-table movement was the wrong abstraction even when the measured height was
    more accurate than the old minimum-row estimate. The landed row-splitting behavior keeps the
    private case page-stable at `16/16`, zero dimension mismatches, `MAE=13.388935`, changed16
    `0.124264` (`20260601-075941`); public `docx-ladder-03-table-row-heights` remains stable at
    `MAE=2.435244`, changed16 `0.018702` (`20260601-075940`). Added a layout unit test that locks
    the default split-by-row behavior. Keep this parent open for table `keepNext` interactions,
    header repetition, row clipping, and remaining baseline/row-origin drift.
    2026-06-01 follow-up: added public Office-backed fixture
    `docx-ladder-03-table-paragraph-adjacency` after the hard-coded post-table `6pt` gap looked
    suspicious. The fixture shows a deeper row-height issue instead: Word compresses short table rows
    with explicit zero cell paragraph after-spacing to about `15.12pt` baseline spacing, while the
    candidate still applies the generic `401twip` auto-row floor and leaves the following paragraphs
    too low (`MAE=1.108290`, changed16 `0.009178`, run `20260601-081230`). A content-owned natural
    auto-row trial improved that public probe to `MAE=0.815674`, changed16 `0.007123`, but collapsed
    the private DOCX to `15` candidate pages with one dimension mismatch (`20260601-081115`), so do
    not land it as-is. The next structural fix must distinguish Word's non-empty natural row height
    from the empty/default row floor without using the private page count as the only guard.
    2026-06-01 follow-up: promoted private-safe cell paragraph spacing counts into the DOCX layout
    snapshot (`ParagraphsWith*SpacingToken`, zero-spacing counts, and min/max resolved spacing). The
    new trace confirms the public adjacency probe has explicit zero cell after-spacing, and the private
    DOCX also has zero after-spacing in all currently laid-out table cells (`422/422`), explaining why
    the naive content-owned trial affected pagination globally. Keep the row-height item open until the
    snapshot can also explain the remaining Office-vs-candidate distinction, likely through paragraph
    line metrics, cell margin exception handling, table style/default paragraph inheritance, or row
    grid/border accounting rather than through `after=0` alone.
    2026-06-01 follow-up: preserved row-level `w:tblPrEx/w:tblCellMar` as a distinct `DocxTableRow`
    property and exposed its presence in the private-safe layout snapshot. This is the structural
    discriminator missing from the rejected content-owned row trial: the public
    `docx-ladder-03-table-paragraph-adjacency` fixture has row property-exception cell margins, while
    private-safe XML inspection of the current DOCX case found `129` rows and `0` rows with `tblPrEx`.
    Auto/default rows with such row property-exception margins now use measured content height instead of
    the generic `401twip` floor; the public adjacency probe improved from run `20260601-081230`
    (`MAE=1.108290`, changed16 `0.009178`) to run `20260601-083833` (`MAE=0.815674`, changed16
    `0.007123`), while the private DOCX run `20260601-083926` stayed at `16/16` pages with zero
    dimension mismatches, no diagnostics, `MAE=13.388935`, changed16 `0.124264`. Neighboring public
    probes remain bounded: `docx-ladder-03-table-row-heights` stayed at `MAE=2.435244`, changed16
    `0.018702` (`20260601-083900`), and `docx-ladder-02-table-cell-margins` is `MAE=0.565499`,
    changed16 `0.007963` (`20260601-083914`). Keep this parent open: a naive horizontal-border
    row-height trial improved the row-height ladder to `MAE=2.350264` but worsened adjacency to
    `MAE=0.927411` and duplicated renderer border-width logic inside layout, so border participation
    needs a proper Word row line-box/border model rather than another local height adjustment.
    2026-06-01 follow-up: refreshed the private-safe layout snapshot after the section-break slice. The private
    DOCX has `129` laid-out table rows, all auto-height, `422` cells with zero resolved cell paragraph
    after-spacing, and only a small set of center/bottom vertical-aligned cells. Public
    `docx-ladder-03-table-row-heights` run `20260601-100156` remains at `MAE=2.435244`,
    changed16 `0.018702`; PDF text inspection shows Office emitting `27` text operations while the candidate
    emits `14`, with row-baseline deltas growing down the table. Do not address this with another row-height
    constant; the next acceptable slice needs to separate row-origin/baseline ownership from Office's finer
    text-operation decomposition on public fixtures.
    2026-06-01 follow-up: DOCX automatic line height now uses the larger of the OpenType typographic line box
    and Windows ascender/descender line box. This is a structural font-metric rule rather than a table-row
    offset: Office's public row-height fixture showed wrapped table-cell line steps larger than the candidate
    typographic-only line box, and using the larger authored font metric moves row origins and baselines closer
    without adding a row-height constant. Public `docx-ladder-03-table-row-heights` improved from
    `MAE=2.435244`, changed16 `0.018702` to `MAE=2.329670`, changed16 `0.017873` (run `20260601-101829`).
    The neighboring table-adjacency fixture stayed bounded at `MAE=0.873157`, changed16 `0.007549`, and the
    private DOCX run `20260601-101840` stayed neutral at `16/16` pages, zero dimension mismatches, no
    diagnostics, `MAE=13.388935`, changed16 `0.124264`. Keep this parent open for row-border participation,
    row clipping, and Office's finer per-cell text-operation decomposition.
    Rejected follow-up: removing the remaining hard-coded post-table `6pt` drop and letting the next paragraph
    own table/paragraph adjacency spacing worsened the public adjacency fixture to `MAE=0.949619`,
    changed16 `0.008026` (run `20260601-102458`) and regressed the private DOCX to `MAE=13.877948`,
    changed16 `0.127199` (run `20260601-102514`). The `6pt` value is still suspicious, but the replacement
    needs an Office-backed table-adjacent block-spacing model, not a removal of the gap in isolation.
    2026-06-01 follow-up: explicit `dxa` table widths now own the layout table width even when they exceed
    the body text width, while percentage/fallback widths remain bounded by the available body width. Public
    PDF inspection of `docx-ladder-03-table-row-heights` showed Office drawing the `w:tblW dxa=9400` table
    past the right margin; the candidate had clamped it to the body width and compressed columns. Public
    `docx-ladder-03-table-row-heights` improved from `MAE=2.329670`, changed16 `0.017873` to run
    `20260601-104626` at `MAE=2.053192`, changed16 `0.016159`; neighboring
    `docx-ladder-03-table-paragraph-adjacency` improved to `MAE=0.846704`, changed16 `0.006517`, and
    `docx-ladder-02-table-cell-margins` stayed unchanged at `MAE=0.565499`, changed16 `0.007963`. Private
    DOCX run `20260601-104626` stayed stable at `16/16` pages, zero dimension mismatches, no diagnostics,
    `MAE=13.286569`, changed16 `0.123783`. Keep the row-height parent open for row-border participation,
    row clipping, and Office's finer per-cell text-operation decomposition.
    2026-06-01 follow-up: DOCX single-line automatic line boxes now have a structural Word minimum of `1.15em`
    after the larger OpenType typographic/Windows line-box metric is applied. This is not a font-name, table,
    or document rule: the public paragraph-spacing probe showed the first baseline already aligned while the
    next paragraph/table origin remained too high, meaning the line box itself was too short. Public
    `docx-ladder-01-paragraph-spacing` improved from failing run `20260601-111037` (`MAE=0.058484`,
    changed16 `0.000818`) to run `20260601-111301` (`MAE=0.014739`, changed16 `0.000394`), matching the
    exact line-height fixture run `20260601-111558`. Public `docx-ladder-02-character-spacing` improved from
    the empty-paragraph baseline (`MAE=1.505741`, changed16 `0.012793`) to run `20260601-111329`
    (`MAE=0.833100`, changed16 `0.008161`), and `docx-ladder-03-table-row-heights` improved from
    `MAE=2.053192`, changed16 `0.016159` to run `20260601-111354` (`MAE=1.469789`, changed16 `0.013068`).
    Neighboring `docx-ladder-03-table-paragraph-adjacency` improved slightly to `MAE=0.778729`, changed16
    `0.006298` (run `20260601-111558`). Keep the parent open because row-border participation, row clipping,
    and Office's per-cell text-operation decomposition are still not structurally explained.
    Rejected follow-up: snapping emitted DOCX PDF font sizes to a 600-DPI text-state grid matched the public
    row-height fixture's Office `11.04pt`/`21.96pt` text states and slightly improved that fixture
    (`MAE=2.025428`, changed16 `0.016094`, run `20260601-105119`), but it worsened
    `docx-ladder-02-character-spacing` from `MAE=1.685749`, changed16 `0.013842` to `MAE=1.699206`,
    changed16 `0.014035`, and regressed the private DOCX aggregate to `MAE=13.349049`, changed16
    `0.124461` (run `20260601-105129`). Do not land a blanket DOCX font-size grid rule; the acceptable path
    needs text-state decomposition evidence that explains when Word chooses secondary font-size branches.
    2026-06-01 follow-up: aligned DOCX table border emission with Office's filled-rectangle structure.
    Vertical borders now fill from the cell boundary to the right instead of being centered on shared
    boundaries or pulled inside the right edge, and horizontal border fills start after the left vertical
    border instead of painting underneath it. This is PDF-structure work, not a row-height constant: the public
    row-height reference PDF showed Word emitting separate vertical strips and horizontal strips between those
    strips. Public `docx-ladder-03-table-row-heights` improved from `MAE=0.778754`, changed16 `0.008189`,
    SSIM `0.792710` to run `20260601-142401` at `MAE=0.700136`, changed16 `0.007778`, SSIM `0.820054`.
    Public `docx-ladder-03-table-paragraph-adjacency` improved from `MAE=0.568657`, changed16 `0.005616` to
    run `20260601-142411` at `MAE=0.545870`, changed16 `0.005497`. Private DOCX run `20260601-142431` stayed
    neutral at `16/16` pages, zero dimension mismatches, no diagnostics, `MAE=13.838763`, changed16
    `0.126851`. Keep this parent open for row clipping, header repetition, and remaining row-origin/text
    operation decomposition.
  - [x] 2026-05-31: Applied DOCX `w:contextualSpacing` for adjacent body paragraphs with the same resolved
    paragraph style. The layout stage now suppresses inter-paragraph spacing in that structural case instead
    of treating contextual spacing as diagnostics-only metadata. Private impact was neutral for the current
    document, but this removes another style-spacing heuristic gap.
  - [ ] 2026-05-31: Finish DOCX paragraph spacing variants: refine exact
    `beforeAutospacing`/`afterAutospacing` values, implement exact Word adjacent-spacing collapse around
    tables/sections, and keep public Office-backed fixtures for each variant before downgrading
    `DOCX_STYLE_PARAGRAPH_SPACING`.
    - [ ] 2026-06-01: Refine empty/styled paragraph line-box ownership against Office PDF evidence. Public
      `docx-ladder-02-character-spacing` proved that a body `<w:p/>` between text and a table must remain in
      the typed model and consume vertical layout space; dropping it kept the table roughly one paragraph too
      high. The reader now preserves empty paragraphs as zero-length paragraph-mark runs resolved through the
      normal run cascade, and layout consumes empty paragraph line boxes in body flow and table cells without
      emitting visible text. Public `docx-ladder-02-character-spacing` improved from `MAE=1.685749`,
      changed16 `0.013842` to run `20260601-110750` at `MAE=1.505741`, changed16 `0.012793`; neighboring
      `docx-ladder-03-table-paragraph-adjacency` stayed unchanged at `MAE=0.846704`, changed16 `0.006517`.
      Keep this item open: the private DOCX remained accepted at `16/16` pages, zero dimension mismatches, no
      diagnostics, but aggregate error worsened to `MAE=13.615949`, changed16 `0.125050` (run
      `20260601-110750`), so styled empty paragraphs still need finer Word-compatible paragraph-mark
      metrics/spacing collapse instead of broad blank-line assumptions.
      2026-06-01 follow-up: the structural single-line minimum greatly improved public body/table origins but
      did not undo the private regression introduced by preserving styled empty paragraphs: private run
      `20260601-111329` stayed accepted at `16/16` pages, zero dimension mismatches, no diagnostics,
      `MAE=13.615949`, changed16 `0.125050`. Keep this nested item open for paragraph-mark style/default
      inheritance and spacing-collapse evidence; do not special-case private style names or fonts.
      2026-06-01 follow-up: aligned empty DOCX paragraph PDF structure with Office by turning preserved
      paragraph-mark runs into positioned blank text lines instead of consuming vertical space invisibly.
      A temporary public-safe styled-empty probe showed Word emits a space text operation at each empty
      paragraph baseline, including `BodyText`/`Compact`-style paragraphs analogous to the private-safe empty
      paragraph inventory (`13` empty body paragraphs: `7` BodyText, `5` Compact, `1` unstyled; all with
      paragraph run properties and no direct spacing). Public `docx-ladder-02-character-spacing` raster metrics
      stayed unchanged at `MAE=0.833100`, changed16 `0.008161` (run `20260601-112943`), and private DOCX run
      `20260601-113037` stayed unchanged at `16/16` pages, zero dimension mismatches, no diagnostics,
      `MAE=13.615949`, changed16 `0.125050`. Keep this open: the structural blank-line baselines are now
      visible to PDF inspection, but the remaining private regression still needs real paragraph-mark
      spacing/style evidence rather than style-name exceptions.
    - [x] 2026-05-31: Converted preserved `beforeLines`/`afterLines` tokens into resolved paragraph spacing
      points when no explicit twip spacing or autospacing owns that paragraph side. The reader now derives
      hundredths-of-line spacing from the resolved paragraph line height, and a public synthetic reader test
      locks the conversion. The parent item remains open because autospacing, style/default spacing cascade
      ownership, adjacent collapse around tables/sections, and Office-authored fixtures are still unresolved.
    - [x] 2026-05-31: Corrected DOCX automatic line spacing to use WordprocessingML's 240ths-of-a-line value
      directly instead of multiplying it by an extra renderer factor. A public reader test now pins
      `w:spacing w:line="276" w:lineRule="auto"` as `1.15` line spacing. Private DOCX run
      `20260531-192811` stayed page-stable at `16/16` with zero dimension mismatches but moved slightly
      against the aggregate (`15.856962` MAE, `0.141649` changed16 versus `15.849350`/`0.141574`), so the
      remaining spacing item stays open for autospacing, table/section collapse, and Office-backed fixtures.
    - [x] 2026-05-31: Tightened `DOCX_STYLE_PARAGRAPH_SPACING` diagnostics so supported style-level
      `beforeLines`/`afterLines` and adjacent same-style `contextualSpacing` no longer look like unsupported
      renderer gaps. The warning still fires for `beforeAutospacing`/`afterAutospacing`; exact Word
      autospacing values, table/section adjacency collapse, and Office-authored fixtures remain open. Private
      DOCX run `20260531-222149` stayed at `16/16` pages with zero dimension mismatches and unchanged aggregate
      metrics (`14.818900` MAE, `0.134293` changed16), while `DOCX_STYLE_PARAGRAPH_SPACING` dropped from the
      private diagnostic list because the remaining private style-spacing triggers were already-supported forms.
    - [x] 2026-06-01: Added public visual fixture `docx-ladder-02-paragraph-autospacing` and moved
      `beforeAutospacing`/`afterAutospacing` from the previous fixed `0pt`/`6pt` placeholders to resolved
      line-height-owned spacing. The spacing cascade is now side-aware: a higher-priority `before*` or `after*`
      token replaces lower-priority twip, line-based, or auto tokens on that same side instead of merging
      incompatible alternatives. Public run `20260601-021353` is dimension-stable and improves the fixture to
      `MAE=0.426487`, changed16 `0.003888` from baseline `0.432740`/`0.003942`; private DOCX run
      `20260601-021406` stayed neutral at `16/16` pages, zero dimension mismatches, no diagnostics,
      `MAE=12.716572`, changed16 `0.118448`. Keep the parent open because Office's exact autospacing amount
      appears slightly lower than the current line-height approximation and table/section collapse is still
      unmodeled.
      2026-06-01 follow-up: the new structural `1.15em` single-line minimum slightly regressed the public
      autospacing probe from `MAE=0.355702`, changed16 `0.003403` to run `20260601-111415`
      (`MAE=0.377856`, changed16 `0.003540`). Keep this open as evidence that `beforeAutospacing` and
      `afterAutospacing` should not simply reuse the normal single-line box; they need their own Office-backed
      spacing model while ordinary line boxes retain the Word-compatible minimum.
      2026-06-01 follow-up: split `beforeAutospacing`/`afterAutospacing` away from line-height-derived
      spacing and resolved them to Word's observed automatic paragraph spacing (`14pt`). Office PDF probes on
      the public autospacing fixture and a temporary 24pt variant both showed the auto gap staying near `14pt`
      rather than scaling with the run size or the normal line box. Public `docx-ladder-02-paragraph-autospacing`
      improved from run `20260601-111415` (`MAE=0.377856`, changed16 `0.003540`) to run `20260601-112135`
      (`MAE=0.362752`, changed16 `0.003422`), while `docx-ladder-01-paragraph-spacing`,
      `docx-ladder-02-character-spacing`, and `docx-ladder-03-table-row-heights` stayed unchanged. Private
      DOCX run `20260601-112215` stayed accepted at `16/16` pages, zero dimension mismatches, no diagnostics,
      `MAE=13.615949`, changed16 `0.125050`. Keep the parent open because the exact Word autospacing ladder
      still needs more font/default/style combinations before downgrading the spacing diagnostic.
      Rejected follow-up: switching missing DOCX line-spacing fallback from `1.25` to `1.15` while also using
      the larger OS/2 Windows/typographic line box looked plausible from the public autospacing baseline drift,
      but it worsened `docx-ladder-02-paragraph-autospacing` to `MAE=0.411331`, changed16 `0.003791` and broke
      line-spacing unit tests. Do not replace the current default line-box rule with that broad combination;
      the remaining drift needs a narrower Office-backed explanation, likely style/default line-spacing
      ownership or PDF text-state decomposition rather than a blanket multiple change.
    - [x] 2026-05-31: Removed the stale style-level keep-rule diagnostic for `w:keepNext`/`w:keepLines` after
      those tokens became part of the resolved paragraph model and block-pagination stage. Public coverage now
      asserts that supported style keep rules do not emit `DOCX_STYLE_PARAGRAPH_KEEP_RULE`. Private DOCX run
      `20260531-233722` stayed raster-neutral at `16/16` pages, zero dimension mismatches, MAE `13.648284`,
      changed16 `0.125542`, and dropped `DOCX_STYLE_PARAGRAPH_KEEP_RULE`; keep the parent open for
      Office-backed keep-with-table and line-level edge cases rather than broad supported-token warnings.
  - [ ] 2026-05-31: Build a real DOCX font-resolution/substitution stage instead of relying on the first run's
    resolved font as a document-wide layout/drawing font. Private-safe inspection shows the worst remaining
    private DOCX pages are dominated by a missing corporate font fallback that renders with a condensed face;
    a trial stable fallback to generic Office text fonts improved glyph appearance but broke private pagination
    to `18` candidate pages (`20260531-191957`/`20260531-192357`), proving that font selection, font-table
    alternates, theme fonts, per-run resources, and layout measurement must be solved together. Preserve this
    as an architecture task: parse `word/fontTable.xml` `w:altName`/family/signature metadata, consume theme
    major/minor font declarations, resolve per-run font resources, and make `DocxLayoutEngine` measure with
    the same run-level fonts that PDF emission uses before changing production fallback policy.
    - [x] 2026-05-31: Added behavior-preserving DOCX font metadata to the document model. `DocxFontCatalog`
      now carries `word/fontTable.xml` entries with `w:altName`, family, pitch, and Panose tokens, plus theme
      major/minor Latin typefaces from the related theme part. Public coverage verifies the reader preserves
      these facts without changing rendering. Private DOCX run `20260531-193148` stayed page-stable at `16/16`
      with the same metrics as the line-spacing run (`15.856962` MAE, `0.141649` changed16), confirming this
      is a safe foundation for the next per-run font-resolution slice.
    - [x] 2026-05-31: Preserved direct run-level DOCX `w:rFonts` tokens on `DocxTextRun` without changing the
      current renderer font choice. The model now carries explicit ascii/high-ANSI/east-Asia/complex-script
      names and their theme counterparts (`asciiTheme`, `hAnsiTheme`, `eastAsiaTheme`, `csTheme`) so the
      future font resolver can reason from OOXML structure instead of a single flattened `FontFamily` string.
      Public coverage verifies token preservation; private DOCX run `20260531-193423` remained page-stable at
      `16/16` with unchanged metrics from the font-catalog slice.
    - [x] 2026-05-31: Promoted run-font token preservation into the DOCX style/default cascade. Resolved run
      properties now carry merged `w:rFonts` tokens from document defaults, paragraph styles, character
      styles, and direct run properties while leaving the legacy flattened `FontFamily` behavior unchanged.
      Public coverage verifies the cascade order; private DOCX run `20260531-193723` stayed page-stable and
      metric-neutral against the previous font-token slice.
    - [x] 2026-05-31: Added a behavior-preserving DOCX typeface candidate resolver over the preserved font
      metadata. It now resolves Latin primary typefaces from cascaded run fonts, document-authored alternates
      from `word/fontTable.xml`, and major/minor theme typefaces from the theme part into ordered candidates.
      Public coverage pins this structure before renderer consumption; production drawing still uses the
      legacy single-font path until layout measurement and PDF emission can consume the same per-run font
      resource map.
    - [x] 2026-05-31: Promote the DOCX renderer from one document-wide font to a run-level font resource and
      measurement map before changing production fallback policy. A generic trial that consumed `fontTable.xml`
      `w:altName` after an exact primary miss, with no named-font exceptions, correctly avoided the arbitrary
      resolver fallback but regressed the private DOCX run `20260531-194952` from `16/16` to `18` candidate
      pages with 2 dimension mismatches (`16.094264` MAE). This confirms the long-term rule: font selection,
      glyph embedding, and line/table measurement must share the same per-run resolved typeface state before
      document-authored alternates can safely affect rendering.
      2026-06-01 update: completed through the shared `DocxFontPlan` measurement/resource path below. The
      broader font architecture item remains open for header/footer run resources, script-slot shaping, and
      exact Word fallback behavior, not for a document-wide font-resource blocker.
    - [x] 2026-05-31: Added a behavior-preserving DOCX font plan that resolves every run through the ordered
      OOXML candidates (primary run typeface, `fontTable.xml` alternate, then theme major/minor typeface) and
      records whether the resolver chose a primary, alternate, theme, or fallback face. Public coverage uses a
      synthetic resolver instead of installed font names, proving the candidate order without hard-coded font
      exceptions. Private DOCX run `20260531-195324` stayed page-stable at `16/16`, zero dimension mismatches,
      and unchanged metrics (`15.856962` MAE, `0.141649` changed16), so the plan can now become the shared
      contract for the upcoming per-run metric and PDF-resource stages.
    - [x] 2026-05-31: Fixed DOCX font embedding to honor `FontResolution.FontFaceIndex` when the resolver
      returns a face from a TrueType collection. The old path loaded collection index 0 regardless of the
      resolved face, which would undermine any future per-run font plan even when the resolver had selected
      the structurally correct typeface. Public coverage dynamically finds an installed collection face with no
      named-font dependency and verifies the emitted PDF uses that face. Private DOCX run `20260531-195628`
      stayed page-stable and metric-neutral (`16/16`, zero dimension mismatches, `15.856962` MAE).
    - [x] 2026-05-31: Extended the DOCX font plan to cover legacy plain table-cell text that is represented
      as `DocxTableCell.Text` rather than paragraph runs. This mirrors the layout engine's synthetic cell-text
      paragraph path so the future per-run measurement/resource map will not silently skip table cell content.
      Public coverage pins the table-cell text path without relying on any installed font.
    - [x] 2026-05-31: Added a private-safe DOCX font-plan inspection snapshot that reports only run counts,
      source counts, and distinct candidate/resolved-family counts, not document text or font names. This gives
      private DOCX investigations a way to see whether pages are dominated by primary, font-table alternate,
      theme, resolver fallback, or missing branches before changing rendering behavior.
    - [x] 2026-05-31: Split DOCX layout measurement behind an `IDocxTextMeasurer` interface while preserving
      the current single embedded-font adapter. Body paragraphs, keep-block estimation, table row sizing,
      table-cell text, numbered-label measurement, and wrapping now call the same measurement abstraction that
      a future per-run font-plan measurer can implement. Public `docx-core` and `docx-tables` groups pass, and
      private DOCX run `20260531-200355` stayed page-stable and metric-neutral (`16/16`, zero dimension
      mismatches, `15.856962` MAE), so this removes the direct layout dependency on `PdfEmbeddedFont` without
      changing output.
    - [x] 2026-05-31: Added a DOCX font-plan text measurer that measures with the resolved `OpenTypeFont`
      face from `DocxFontPlan`, including TrueType collection face index, instead of a document-wide or named
      fallback font. This is still a bridge for tests and the next renderer slice, not production fallback
      policy: production output remains unchanged until line layout and PDF emission share the same run-level
      resource map. Public coverage dynamically selects an installed usable face without hard-coded family
      names and verifies the measurement matches raw OpenType advances plus kerning.
    - [x] 2026-05-31: Promoted DOCX production layout inspection and rendering to use the `DocxFontPlan`
      measurer for run-level width and line-height measurement, while keeping PDF emission on the existing
      single document fallback font resource. The measurer now takes the same resolved fallback resource used
      by emission for runs with no explicit OOXML font family, avoiding zero-width "missing" runs without
      introducing font-name exceptions. Validation passed `docx-text --skip-slow` (`16`), `docx-tables
      --skip-slow` (`43`), `docx-numbering --skip-slow` (`9`), and full solution build. Private DOCX run
      `20260531-230150` stayed page-stable at `16/16`, zero dimension mismatches, with changed16 slightly
      better (`0.133893` -> `0.133867`) and MAE slightly worse (`14.791805` -> `14.806507`), so this is a
      structural layout-alignment step rather than a completed visual win.
    - [x] 2026-05-31: Fixed DOCX non-numbered body/table line widths and segment positions to use the
      line's run-level text segments instead of measuring the whole line with the first run. This removes a
      structural first-run flattening gap that could misplace mixed-style text and lose later run color/style
      during body paragraph emission. Public `docx-core --skip-slow` passed `12`, and `docx-tables
      --skip-slow` passed `38`.
    - [x] 2026-05-31: Finished the first DOCX mixed-run line-building slice by making wrapping own run spans,
      not just final line segments. Body and table-cell paragraph wrappers now preserve run boundaries through
      tokenization, line breaking, numbering continuation lines, and final segment placement, so mixed
      fonts/styles can influence break points before PDF emission. Public coverage includes a wide second-run
      case that would not break under first-run measurement. Private DOCX run `20260531-201414` stayed
      page-stable at `16/16` with zero dimension mismatches and improved aggregate MAE from `15.856962` to
      `14.577536` (changed16 `0.136194`). Keep the broader font item open for production run-level PDF font
      resources, tabs, bullet fonts, and exact Word line-break edge cases.
    - [x] 2026-05-31: Replaced the DOCX renderer's document-base font selection with `DocxFontPlan`
      resolution instead of choosing the first literal family and letting the resolver fall through. This
      preserves Word's structural candidate order: direct/default/style run font, font-table alternate,
      theme face, then resolver fallback. The reader now walks paragraph/character `w:basedOn` chains before
      direct properties, so inherited corporate style fonts flow into the same font plan without hard-coded
      family names. Public coverage pins theme-only defaults, font-table alternate-before-fallback, and
      based-on style inheritance with dynamically discovered installed fonts. Office/candidate PDF inspection
      on the private DOCX showed both sides using Calibri-family fonts after this change, so the remaining
      private mismatch is no longer a concrete font-name selection issue.
    - [x] 2026-05-31: Preserved DOCX numbering-marker `w:lvl/w:rPr` as a generic run style and threaded it
      into the font plan and list-label layout measurement. `DocxListLabel` now carries authored marker run
      color/style/font tokens, bullet labels use authored `w:lvlText` instead of a blanket bullet glyph
      replacement, and marker runs participate in `DocxFontPlan` without named-font branches. Public
      `docx-numbering --skip-slow` passed `9` tests and `fonts --skip-slow` passed `20`; private DOCX run
      `20260531-221725` stayed at `16/16` pages with zero dimension mismatches and slightly improved MAE
      `14.818900`, changed16 `0.134293`. Keep the broader item open because production PDF emission still
      uses one font resource per document, so fully correct symbol glyph emission requires the shared per-run
      layout/resource map below.
    - [x] 2026-05-31: Promoted DOCX production text emission to run-level PDF font resources derived from
      `DocxFontPlan`, after layout measurement was already using the same run-level resolved typeface plan.
      PDF resources are now grouped by resolved font file and TrueType collection face index, with fallback
      resources used only for unresolved runs; public coverage verifies distinct run typefaces become distinct
      `/F*` resources without hard-coded family names. The stale `DOCX_NUMBERING_MARKER_FONT` diagnostic was
      removed because numbering marker pseudo-runs now participate in both run-level measurement and PDF
      glyph emission. Private DOCX run `20260531-235343` improved to `16/16` pages, zero dimension
      mismatches, MAE `12.509698`, changed16 `0.112673`, and only reports
      `DOCX_STYLE_TABLE_COMPLEX_SCRIPT_RUN`.
      Prior failed attempts remain useful evidence: an emission-only trial stayed page-stable but regressed
      the private aggregate from the span-wrapper run's `14.577536` MAE to `16.256186`
      (`20260531-201939`, changed16 `0.144954`), proving layout measurement and glyph emission had to switch
      together.
      2026-05-31 update: a second no-font-name trial switched layout measurement and PDF emission together
      through per-run `DocxFontPlan` resources and used resolver metadata to avoid synthetic bold/italic when
      the selected face already carried the style. It still regressed the private DOCX aggregate
      (`20260531-225826`, `MAE=14.809827`, changed16 `0.134021`) against the current single-resource baseline,
      so it was not landed; the completed implementation keeps the structural resource map while preserving
      the already-aligned font-plan measurer.
    - [x] 2026-06-01: Modeled DOCX complex-script run style slots from table styles instead of warning on
      `w:bCs`/`w:iCs`. Resolved run properties now carry `bCs`/`iCs`, text is split into Latin versus
      complex-script spans before layout/emission, complex-script spans prefer `w:rFonts/@w:cs` and
      `csTheme` (`majorBidi`/`minorBidi`), and Latin spans do not inherit complex-script bold/italic. Public
      table coverage uses mixed Latin/Hebrew text to prove script-slot behavior without font-name
      special-casing. Private DOCX run `20260601-000124` stayed at `16/16` pages with zero dimension
      mismatches, MAE `12.509698`, changed16 `0.112673`, and no diagnostics.
    - [ ] 2026-06-01: Replace the current DOCX script-slot classifier with a Word-compatible shaping and
      bidi stage when complex-script visual gaps become visible. The present implementation is a structural
      slot-selection step over Unicode ranges; it is enough to route `w:bCs`/`w:iCs`, `w:cs`, and Bidi theme
      faces into the existing run model, but it is not a full OpenType shaping engine, bidi reordering model,
      or per-script fallback ladder. Keep this as an explicit long-term gap so future fixes do not devolve
      into font-name or document-specific exceptions.
    - [ ] 2026-06-01: Add an OpenType shaping/positioning stage for DOCX Latin text before chasing residual
      text-state differences with font-name or document-specific rules. Public `docx-numbering` is visually
      close (`20260601-142558`, `MAE=0.019271`, changed16 `0.000744`), but Office's reference PDF emits
      intra-word `TJ` positioning adjustments for ordinary Latin text while the current candidate usually
      emits one glyph chunk per segment. The existing measurement/emission path consumes simple glyph advances
      and legacy `kern` pairs, but it does not apply GPOS pair positioning, mark positioning, glyph
      substitution, or a true shaping buffer. Keep this as a structural text-emission gap shared by body,
      table, numbering, and static header/footer text.
      2026-06-01 evidence: public text-state probes show a real PDF-state decomposition mismatch, but not yet
      a safe renderer rule. Office emits small nonzero `Tc` for short Arial tokens such as `42`, `Q1`, and
      `AB` while the candidate emits `Tc=0`; however the effective rendered advance is already nearly aligned
      (for example the public context probe advances `42` to the following space at about `12.24pt` in Office
      versus `12.235pt` in the candidate). Treat this as evidence for a future shaping/positioning and
      Office-PDF emission planner, not as permission to bucket by token length, text class, table context, or
      font name.
    - [ ] 2026-05-31: Split DOCX static header/footer rendering into run-level line segments instead of
      concatenating all runs and drawing them with the first run's resource/style. This is now the clearest
      remaining font-resource architecture gap: body/table text can use the run resource map, but static
      header/footer text still collapses mixed runs, `{PAGE}` substitution, color, and font resource ownership
      into the first run. Add public fixtures before touching private documents.
      2026-06-01 discovery: a straightforward fallback-free static/table text rendering trial is not safe yet.
      Enabling static/table text whenever run-level resources exist, and adding all first/even/default
      header/footer parts to the global font plan, preserved private page count and diagnostics but worsened
      the private DOCX aggregate from `12.509698` MAE / `0.112673` changed16 to `13.773663` MAE / `0.126283`
      changed16. The trial was reverted. The next acceptable slice is a real static header/footer layout stage
      with selected-part ownership, Word-compatible header/footer line boxes and baselines, and public
      Office-PDF-backed fixtures; do not re-enable static fallback-free rendering as a side effect of the
      document-level fallback resource.
      2026-06-01 progress: static header/footer emission now preserves run-level segmentation for the selected
      header/footer part while keeping the existing static placement model. Each run resolves its own PDF font
      resource, color, italic state, and `{PAGE}` substitution before line alignment. Public unit coverage
      checks mixed red/blue header runs emit separately. Private run `20260601-020720` stayed neutral against
      the table-style/row-minimum baseline (`16/16`, zero dimension mismatches, `MAE=12.716572`, changed16
      `0.118448`). Keep the parent open for selected-part layout snapshots, wrapping, and exact Word
      header/footer line boxes rather than first-run resource collapse.
      2026-06-01 follow-up: replaced the remaining raw static header/footer baseline heuristic with resolved
      font metric anchoring. Header baselines are now inset from the OOXML header-distance top by the maximum
      run Windows ascender, footer baselines are inset from the footer-distance bottom by the maximum run
      Windows descender, and static runs keep their authored font sizes instead of being capped at `12pt`.
      Public `docx-headers-footers` improved from `MAE=0.498914`, changed16 `0.006279` to `MAE=0.415463`,
      changed16 `0.005343` (`20260601-082246`); inspected PDF baselines moved to header `746.95pt` versus
      Office `746.64pt` and footer `38.12pt` versus Office `38.06pt`. Private DOCX run `20260601-082318`
      stayed accepted (`16/16`, zero dimension mismatches, no diagnostics, `MAE=13.388935`, changed16
      `0.124264`). Keep this open for true static header/footer layout records, selected-part snapshots,
      wrapping, and multi-paragraph line-box behavior.
      2026-06-01 follow-up: static header/footer draw order now matches Office's observable PDF structure by
      emitting selected static content before body content on each page. The previous body-first order was
      visually neutral for non-overlapping fixtures but disagreed with the `docx-headers-footers` reference
      content stream. Public run `20260601-143229` stayed raster-identical (`MAE=0.073352`, changed16
      `0.002110`, SSIM `0.982572`), while candidate PDF text operations now start with the header/footer text
      like Office. Private DOCX run `20260601-143252` stayed neutral at `16/16` pages, zero dimension
      mismatches, no diagnostics, `MAE=13.838763`, changed16 `0.126851`.
      2026-06-01 follow-up: selected static header/footer text lines now belong to the DOCX layout model as
      `DocxLayoutPage.StaticTextLines` instead of being positioned entirely inside `DocxRenderer`. The layout
      stage resolves first/even/default part selection, `{PAGE}`/`{NUMPAGES}` cached placeholders, alignment,
      segment x-positions, and font-metric baselines; the renderer consumes those line records through the
      same run-segment emission path used by body text. The private-safe layout snapshot now reports
      `StaticTextLineCount`, closing the inspection gap without exposing header/footer text. Public
      `docx-page --skip-slow` passed (`27`), `docx-text --skip-slow` passed (`38`),
      `docx-tables --skip-slow` passed (`77`), and post-cleanup `docx-headers-footers` stayed raster-identical at
      `MAE=0.073352`, changed16 `0.002110` (`20260601-154031`). Private DOCX run `20260601-153607` stayed
      neutral at `16/16` pages, zero dimension mismatches, no diagnostics, `MAE=13.855991`, changed16
      `0.127419`. Keep this item open for static wrapping, multi-paragraph line boxes, true field evaluation,
      and selected-part snapshot details beyond counts.
      2026-06-01 follow-up: static header/footer paragraphs now wrap inside the selected page body width at
      the layout stage and still emit ordinary `DocxTextLineLayout` records per visual line. The wrapper uses
      the same whitespace/soft-break tokenization shape as body text, but measures static spans with each run's
      own authored font size so the earlier mixed-run resource work is not collapsed back to a paragraph-wide
      size. Public `docx-page --skip-slow` passed (`28`), `docx-text --skip-slow` passed (`38`),
      `docx-tables --skip-slow` passed (`77`), and `docx-headers-footers` stayed unchanged at `MAE=0.073352`,
      changed16 `0.002110` (`20260601-154404`). Private DOCX run `20260601-154426` stayed neutral at `16/16`
      pages, zero dimension mismatches, no diagnostics, `MAE=13.855991`, changed16 `0.127419`. Keep this
      parent open for exact footer stacking direction, paragraph spacing inside headers/footers, true Word
      field evaluation, and richer private-safe snapshots of selected static parts beyond counts.
      2026-06-01 follow-up: static header/footer layout now applies the same max-collapse paragraph
      before/after spacing shape used by body/table text and creates invisible paragraph-mark lines for empty
      static paragraphs instead of silently dropping their line boxes. Public `docx-page --skip-slow` passed
      (`29`), `docx-text --skip-slow` passed (`38`), `docx-tables --skip-slow` passed (`77`), and
      `docx-headers-footers` stayed unchanged at `MAE=0.073352`, changed16 `0.002110` (`20260601-154643`).
      Private DOCX run `20260601-154658` stayed neutral at `16/16` pages, zero dimension mismatches, no
      diagnostics, `MAE=13.855991`, changed16 `0.127419`. Keep the parent open for exact footer stacking
      direction, contextual-spacing edge cases in static parts, true field evaluation, and richer private-safe
      static-part snapshots.
      2026-06-02 follow-up: wrapped static header/footer lines now carry `SourceLineIndex` and correct
      `IsFirstParagraphLine` ownership instead of marking every visual line as a first paragraph line. This is
      a model/snapshot correction: rendering geometry is unchanged, but selected static parts now line up with
      the body/table source-line contract used by private-safe diagnostics. Extended the existing wrapped
      header test to assert `SourceParagraphIndex=0`, `SourceLineIndex=0/1`, and first/continuation line flags.
      Validation passed `docx-page --skip-slow` (`30`), `docx-text --skip-slow` (`45`), and public
      `docx-headers-footers` run `20260602-043616` stayed at `MAE=0.073352`, changed16 `0.002110`, SSIM
      `0.982572`. Keep the parent open for exact footer stacking direction, contextual-spacing edge cases,
      true field evaluation, and selected-part snapshots beyond counts and source-line ownership.
      2026-06-02 follow-up: layout snapshots now expose selected static text through `StaticItems`, using the
      same private-safe `DocxLayoutItemSnapshot` shape as body items. Static item `Kind` distinguishes
      `StaticHeaderTextLine` from `StaticFooterTextLine`, so header/footer paragraph and line indexes no
      longer collide in `layout-snapshot.json`. Public `docx-headers-footers` inspection regenerated for run
      `20260602-043616` and reports header/footer static items with `SourceParagraphIndex=0`,
      `SourceLineIndex=0`, first-line flags, and text lengths only. Validation passed `docx-page --skip-slow`
      (`30`). Keep the parent open for exact footer stacking direction, contextual-spacing edge cases, and
      true field evaluation; selected static line snapshots now exist beyond aggregate counts.
    - [x] 2026-06-01: Close the fallback-free DOCX table-cell text emission gap without a private regression.
      Body text already renders through run-level resources, but table-cell text is still gated by the
      document fallback resource. Removing that gate rendered additional private table text and regressed the
      aggregate metrics above, which means the missing piece is not the gate itself but exact table-cell
      resource/layout ownership: selected run resource, row height, vertical alignment, clipping/overflow, and
      Word-compatible cell text baselines must be validated together with a public table fixture before the
      fallback gate can be removed.
      2026-06-01 progress: added public visual fixture `docx-ladder-02-table-explicit-font`, a table-only DOCX
      whose cell runs use authored font resources and therefore expose the current fallback-free table-text
      suppression without private content. Office reference PDF inspection has text operations while the
      candidate has none; use this case as the public gate before retrying production table text emission.
      2026-06-01 follow-up: a reverted trial removed the fallback gate and moved table-cell baselines onto the
      body paragraph baseline formula. The public fixture improved from MAE `0.263375` / changed16 `0.003477`
      to MAE `0.241225` / changed16 `0.003272`, and PDF matrices showed the cell top-to-baseline offset was
      nearly aligned (`~11.30pt` Office versus `~11.28pt` candidate). The private DOCX still regressed from
      the current `12.509698` MAE / `0.112673` changed16 baseline to `13.993723` MAE / `0.128607` changed16.
      On a representative private page, candidate text operations rose from `60` to `93` while Office had
      `140`, so the trial added incomplete table text bands instead of closing the Office text coverage gap.
      Keep the gate until table-cell text ownership is complete enough to reduce private raster mismatch.
      2026-06-01 closure: after run-level font resources, table-cell paragraph ownership, row minimums, and
      table-style run precedence landed together, table text no longer depends on the document fallback gate.
      Public `docx-ladder-02-table-explicit-font` remains stable at `MAE=0.229378`, changed16 `0.003197`,
      and the private DOCX run `20260601-020048` stayed `16/16` pages with zero dimension mismatches while
      improving aggregate MAE/changed16. Keep static header/footer fallback-free rendering as a separate open
      item because it still lacks a selected-part line-layout stage.
    - [ ] 2026-05-31: Resolve the DOCX pagination gap exposed by structural font/style alignment. Private
      DOCX run `20260531-203336` improved aggregate MAE to `13.852449` and changed16 to `0.125076`, but the
      candidate now paginates as `14` pages against Office's `16` reference pages with `2` dimension mismatches.
      Do not respond by restoring a document-wide hard-coded font fallback; the inspected reference and
      candidate PDFs both point at Calibri-family output. The next acceptable work is Word-compatible vertical
      composition: paragraph spacing collapse/context, keep-with-next/keep-lines, table style paragraph/run
      precedence, table-cell spacing, repeated headers, and exact row pagination.
      2026-05-31 update: Office PDF inspection showed the dominant 10 pt text baseline step is `14.04pt`, while
      OOXPDF used `11.5pt` from `fontSize * 1.15`. Auto line height now uses the resolved OpenType line box
      before applying the OOXML auto factor; the same inspected bucket moved to `14.04pt`, and private run
      `20260531-204257` improved pagination from `14` candidate pages to `15` against the `16`-page reference
      with `1` dimension mismatch (`14.118472` MAE, changed16 `0.128327`). Keep this item open for the remaining
      one-page deficit; the next likely causes are table style paragraph/run precedence, repeated-header/table
      pagination, and keep/spacing interactions, not font-family selection or raw auto line pitch.
      2026-05-31 update: table-cell layout now applies paragraph `spacingBefore` with the same max-collapse
      shape used by body paragraphs, instead of only consuming `spacingAfter`. Private run `20260531-204834`
      stayed at `15/16` pages but improved aggregate MAE to `13.983783` and changed16 to `0.127071`, so the
      remaining one-page deficit is not explained by the compact table-cell before-spacing alone.
      2026-05-31 update: widow/orphan control now defaults on unless `w:widowControl` is explicitly false,
      matching Word's default paragraph behavior. Public tests cover default-on, explicit-on, and explicit-off
      pagination. Private run `20260531-205119` stayed at `15/16` pages and slightly worsened aggregate metrics
      (`14.058788` MAE, changed16 `0.127630`), so default widow control is structurally correct but not the
      missing page's primary driver.
  - [x] 2026-05-31: Preserved DOCX numbering-level indent tokens and applied a first layout-stage indent
    approximation for numbered paragraphs. `DocxListLabel` now carries typed left/right/first-line/hanging
    indent values from `w:lvl/w:pPr/w:ind`, and body/table-cell paragraph layout uses those values to shift
    numbered line starts and reduce wrapping width. This improves private DOCX aggregate fidelity while keeping
    `DOCX_NUMBERING_INDENT` open as an approximation because Office's true label tab stop, bullet fonts,
    mixed-run wrapping, and style-owned numbering behavior are not yet structurally modeled.
  - [x] 2026-05-31: Split DOCX numbering layout into label run, tab stop, text start, and continuation-line
    geometry. The current indent application moves the flattened label+text line as one unit; long-term Word
    parity needs a layout record that separates the list marker from paragraph text and uses numbering
    tab/hanging rules consistently in body text and table cells.
  - [x] 2026-05-31: Split DOCX body numbering emission into separate marker and paragraph-text segments.
    Body numbered paragraphs now place the label in the hanging area and place text at the numbering `left`
    start, with continuation lines using the text start. This improved the private DOCX aggregate while
    preserving the open numbering diagnostic for exact tab-stop, bullet-font, and table-cell numbering parity.
  - [ ] 2026-05-31: Complete DOCX numbering fidelity after marker/text separation: preserve and apply
    numbering tab stops, symbol/bullet fonts, level suffix behavior, continuation-line wrapping across
    mixed-style runs, restart overrides, and the same segmented model inside table cells.
    2026-05-31 progress: numbering levels now preserve `w:suff` and layout distinguishes tab, space, and
    nothing suffix behavior for first-line text placement. Numbered table-cell paragraphs now use the same
    label/text segment split as body paragraphs instead of flattening the label into the cell text run. Public
    coverage checks suffix preservation, space-suffix geometry, and table-cell numbering segmentation. Private
    DOCX run `20260531-173037` stayed neutral (`15.889775` MAE, `0.141949` changed16), and
    `DOCX_NUMBERING_INDENT` remains appropriate because exact Word tab-stop ownership, bullet fonts, restarts,
    and mixed-style continuation wrapping are still open.
    2026-05-31 progress: numbering-marker run properties from `w:lvl/w:rPr` are now preserved as typed
    `DocxTextRunStyle` data and applied to the separate marker segment for body/table-cell layout. Bullet
    marker text now comes from authored `w:lvlText` when present instead of a hard-coded glyph substitution,
    while an absent `w:lvlText` still receives a generic bullet fallback. Public coverage verifies marker
    font/color/bold token preservation and marker participation in `DocxFontPlan`; private DOCX run
    `20260531-221725` stayed page-stable at `16/16`, zero dimension mismatches, MAE `14.818900`, changed16
    `0.134293`. Keep `DOCX_NUMBERING_INDENT` open for exact numbering tabs and the production per-run PDF
    font resource switch needed for full symbol-glyph parity.
    2026-05-31 progress: DOCX numbering levels now preserve explicit `w:pPr/w:tabs/w:tab w:val="num"`
    positions and layout uses that authored marker position instead of always deriving label x from
    `left - hanging`. Public reader/layout coverage forces the numbering tab to disagree with the derived
    indent formula and verifies only the marker moves while the paragraph text still starts at the numbering
    left indent. Private DOCX run `20260531-231855` stayed neutral at `16/16` pages, zero dimension mismatches,
    MAE `13.648284`, changed16 `0.125542`. Keep the parent open for production per-run PDF font resources,
    bullet/symbol glyph parity, and any remaining Word numbering restart/override semantics; do not replace
    these with font-name or document-specific branches.
    2026-05-31 progress: narrowed numbering diagnostics so supported left/hanging indents and explicit
    `w:tab w:val="num"` positions no longer emit `DOCX_NUMBERING_INDENT`; unsupported `right`/`firstLine`
    indent forms keep that warning. Bullet levels with authored marker fonts now emit
    `DOCX_NUMBERING_MARKER_FONT`, which better describes the remaining structural risk: marker runs already
    participate in the font plan, but production PDF emission still needs the same per-run font resource map
    before symbol/bullet glyph parity can be considered complete. Private DOCX run `20260531-234039` stayed
    raster-neutral at `16/16` pages, zero dimension mismatches, MAE `13.648284`, changed16 `0.125542`, and
    replaced `DOCX_NUMBERING_INDENT` with `DOCX_NUMBERING_MARKER_FONT`.
    2026-05-31 progress: after the production run-level PDF resource map landed, the marker-font warning was
    removed as stale rather than kept as a broad heuristic. Numbering marker pseudo-runs now flow through the
    same font-plan measurement and PDF glyph embedding as body/table runs. Keep this parent open only for
    remaining numbering semantics: exact tab-stop ownership, restart/override edges, and Word fixture coverage
    for suffix and continuation-line behavior.
    2026-06-01 correction: the earlier interpretation that an explicit `w:tab w:val="num"` position can move
    the marker itself was too broad. Public `docx-ladder-03-compact-bullet-spacing` shows Office PDF output
    placing the bullet marker at the hanging position (`left - hanging`) and placing following paragraph text
    at the tab/left position. `DocxLayout` now treats the numbering tab as post-marker text geometry rather
    than marker origin. This keeps exact tab-stop ownership open for richer list fixtures, but rejects the
    prior marker-at-tab rule as a structural mismatch.
  2026-05-31 progress: paragraph wrapping now carries separate first-line and continuation-line widths, so
    numbered body and table-cell paragraphs wrap continuation lines against the hanging text column rather than
    reusing the wider first-line space-suffix box. Public `docx-numbering --skip-slow` passed `8` tests with a
    direct hanging-continuation layout check, and `docx-tables --skip-slow` passed `35`; private DOCX run
    `20260531-190121` stayed neutral at `16/16` pages, zero dimension mismatches, MAE `15.849350`, changed16
    `0.141574`. Keep the diagnostic open for exact tab stops, bullet fonts, style inheritance, and mixed-run
    numbering text segmentation.
  2026-05-31 progress: parsed `w:num/w:lvlOverride/w:startOverride` and applied the num-specific start value
    before incrementing list counters. Public coverage verifies labels restart from `5.`/`6.` while sharing the
    abstract level. Private DOCX run `20260531-175443` stayed neutral (`15.889775` MAE, `0.141949` changed16).
    2026-05-31 progress: multilevel label text now resolves `%1` through `%9` from the active counter state and
    resets deeper counters when a higher level advances. Public coverage verifies `1.`, `1.1.`, `1.2.`, `2.`,
    `2.1.` label progression. Private DOCX run `20260531-175714` stayed neutral (`15.889775` MAE, `0.141949`
    changed16).
  - [x] 2026-05-31: Preserved DOCX `w:tblHeader` row tokens and repeated contiguous header rows when a table
    page break occurs. This moves repeating headers out of diagnostics-only handling and into the table layout
    stage with public coverage. The diagnostic remains as an approximation because multi-row/header-group
    semantics, interaction with table styles, and exact Word page-break behavior are not fully modeled.
  - [ ] 2026-05-31: Continue the DOCX table style track: parse table style definitions, conditional style
    regions, inherited table/cell borders, shading, margins, and header-row formatting before attempting more
    private-document table fixes. The private case still reports table-style diagnostics and did not improve
    from header-row repetition alone, so the next table work should target style resolution rather than row
    repetition.
  - [x] 2026-05-31: Preserved DOCX table style IDs and applied the simplest whole-style cell shading from
    table style definitions. `DocxTable` now carries `w:tblStyle`, `DocxStyleSet` keeps table-style records,
    and cells without direct shading inherit style-level `w:tcPr/w:shd` fill/color tokens before the existing
    PDF cell-fill path. This is a first style-resolution rung, not full table-style support.
  - [ ] 2026-05-31: Extend DOCX table styles beyond whole-style cell shading: parse `w:tblPr`, `w:tblBorders`,
    `w:tblCellMar`, `w:tblLook`, and `w:tblStylePr` conditional regions (`firstRow`, `lastRow`, banding,
    first/last column), then merge them with direct row/cell properties in Word priority order.
    - [x] 2026-05-31: Promoted table-style `w:tblPr/w:tblBorders` into the existing per-cell border ladder.
      Style table borders now map outer and inside horizontal/vertical edges onto cells, with direct table
      borders, conditional/cell-style borders, and direct cell borders layered above them. Public coverage
      checks a 2x2 styled table with distinct outer and inside colors. The parent remains open for `tblLook`
      toggles, full conditional precedence, and row/table exception properties.
  - [x] 2026-05-31: Applied DOCX conditional table-style cell shading for first/last row, first/last column,
    corner cells, and first horizontal/vertical bands. Conditional `w:tblStylePr/w:tcPr/w:shd` now flows into
    cell fill/shading tokens before PDF emission. The private DOCX metrics improved, confirming conditional
    table style resolution is an active rendering branch.
  - [ ] 2026-05-31: Refine conditional table-style precedence against Word fixtures. The current merge order is
    structural but incomplete: it does not yet honor `w:tblLook` toggles, second band regions, or the full Word
    priority ladder across whole-table style, conditional style, direct row/cell properties, and table
    exceptions.
    - [x] 2026-05-31: Preserved `w:tblLook` source tokens as typed table metadata (`val`, first/last row,
      first/last column, and horizontal/vertical band suppression flags) with public reader coverage. A naive
      behavior change that gated conditional regions directly from these flags regressed the private DOCX run
      (`20260531-174932`: MAE `16.580381`, changed16 `0.14746` versus baseline `15.889775`/`0.141949`), so
      `tblLook` remains a fixture-required Office semantics task rather than a guessed rendering rule.
    - [x] 2026-05-31: Preserved DOCX cell-level `w:cnfStyle` conditional-format tokens and used them as the
      authoritative table-style region source when present. The fallback positional region inference remains
      only for cells without `cnfStyle`. Public coverage verifies that an explicitly banded first-row cell uses
      the band style instead of a guessed first-row style, and that the raw `cnfStyle` token is retained in the
      cell model. Private DOCX run `20260531-184012` stayed neutral against the current baseline
      (`15.864724` MAE, `0.141689` changed16), which is acceptable because the slice removes a style-cascade
      guess without changing the visible page count.
    2026-05-31 rejected trial: applying `w:tblLook` as a direct gate over fallback positional conditional
    regions still regressed the private DOCX case (`20260531-184208`: MAE `16.582670`, changed16 `0.147486`).
    Keep `tblLook` as preserved metadata until public Word fixtures clarify how it interacts with `cnfStyle`,
    default table-style toggles, and band-region priority.
  - [ ] 2026-05-31: Resolve DOCX table-style paragraph/run property precedence with Office-backed public
    fixtures before downgrading table-style diagnostics. Private-safe inventory shows table-style conditional
    `w:pPr` and `w:rPr` exist, but naive inherited text-style layers for alignment/bold/font-size worsened
    private aggregate metrics (`20260531-160500`/`20260531-160557` regressed from the `tcW` run). Treat this as
    a cascade-order/model gap, not as evidence to ignore table-style text properties.
    - [x] 2026-05-31: Enabled table-style paragraph/run properties through the actual DOCX style cascade
      instead of post-processing parsed paragraphs. Table-style `w:pPr`/`w:rPr` now merge after paragraph and
      character styles but before direct paragraph/run properties, and fallback conditional regions respect
      preserved `w:tblLook` flags when no cell-level `w:cnfStyle` is present. Public coverage verifies
      inherited paragraph alignment, run italic/color/size, `tblLook` first-column gating, whole-style spacing,
      and direct paragraph spacing override. Private DOCX run `20260531-220413` restored `16/16` pages with
      zero dimension mismatches after the structural font/pagination work, but aggregate MAE rose to
      `14.825006` and changed16 to `0.134342` versus the prior committed `15/16` run. Keep the parent open for
      Word-backed precedence fixtures, diagnostics cleanup, and visual tuning; do not replace this with
      document- or font-specific rules.
    - [x] 2026-05-31: Added table-style cell `w:vAlign` to the same resolved conditional-cell style cascade.
      Whole-style and `w:tblStylePr` conditional `w:tcPr/w:vAlign` now flow into `DocxTableCell` unless a direct
      cell `w:vAlign` overrides them. Private-safe inventory found one first-row `bottom` alignment atom in the
      private DOCX table style, so this closes a real dropped style property without adding a document-specific
      rule.
    - [x] 2026-05-31: Resolved DOCX table-style `w:basedOn` chains before table cells consume style data.
      Whole-style cell properties, table borders, margins, paragraph/run properties, and conditional
      `w:tblStylePr` regions now merge base-to-child with child declarations overriding inherited values.
      Public coverage verifies inherited fill, margin, paragraph alignment, run color/bold, and conditional
      region merging. Private DOCX run `20260531-223326` improved aggregate metrics to `14.791805` MAE and
      `0.133893` changed16 while staying page-stable at `16/16`; keep this parent open for the remaining Word
      precedence ladder and table-style diagnostics.
    - [x] 2026-05-31: Added table-style `w:tblPr` table-property cascade for `tblLayout`, `tblW`, `tblInd`,
      and `tblCellSpacing`. Styled tables now inherit these table-owned geometry properties unless direct
      table properties override them, matching the same source-order model used for direct DOCX table
      properties. Public coverage verifies inherited width/indent/cell-spacing/layout and direct override
      behavior. Private DOCX run `20260531-223803` was metric-neutral against the `basedOn` baseline
      (`16/16`, zero dimension mismatches, `14.791805` MAE, `0.133893` changed16), but this removes another
      structural table-style blind spot found by private-safe style inventory.
    - [x] 2026-05-31: Preserved and applied table-style row/column band sizes from
      `w:tblStyleRowBandSize` and `w:tblStyleColBandSize` during fallback conditional-region inference.
      Public coverage now checks a two-row band before switching from `band1Horz` to `band2Horz`. Private
      DOCX run `20260531-224113` stayed metric-neutral (`16/16`, zero dimension mismatches, `14.791805` MAE,
      `0.133893` changed16), as expected because the private styles declare the Office-default band size of
      `1`.
    - [x] 2026-05-31: Preserved and applied DOCX run all-caps from `w:rPr/w:caps` through the shared run-style
      cascade, including table-style conditional `w:rPr`. Resolved runs now carry an `AllCaps` flag and
      materialize uppercase display text before layout measurement and PDF emission, so the renderer does not
      measure one string and draw another. Private-safe table-style inventory found `w:caps` in the remaining
      style surface; public coverage verifies a first-column table-style `w:caps` run. Private DOCX run
      `20260531-225034` stayed metric-neutral (`16/16`, zero dimension mismatches, `14.791805` MAE,
      `0.133893` changed16). Keep the parent open for script-aware `bCs`/`iCs` and exact table-style
      diagnostics.
  - [ ] 2026-05-31: Resolve DOCX table-cell paragraph spacing semantics with public Word fixtures before
    applying `spacing before` inside table cells. The private table-cell paragraph style carries
    `w:spacing before="36" after="0"`, but a direct cell-layout application regressed private run
    `20260531-161411`, indicating Word suppresses or contextualizes this spacing in table cells.
  - [x] 2026-05-31: Promoted DOCX table borders from stored tokens to PDF emission and conditional style
    inheritance. Cells now inherit conditional `w:tcBorders` when direct borders are absent, and the renderer
    draws authored top/bottom/left/right border colors and widths per edge instead of always stroking every
    cell with a uniform black rectangle. This produced a large private improvement and confirms table border
    style resolution is a high-impact branch.
  - [ ] 2026-05-31: Continue DOCX border fidelity: implement table-level `w:tblBorders`, inside horizontal/
    vertical borders, conflict resolution between adjacent cell borders, nil/none suppression across shared
    edges, and Word's exact border width/unit rules with public Office PDF fixtures.
    2026-05-31 progress: direct table-level `w:tblBorders` now maps outer and inside horizontal/vertical
    borders onto cell edges when stronger cell/style borders are absent, keeping the existing cell-border PDF
    emission path. Public coverage checks a 2x2 table where outer colors and `insideH`/`insideV` land on the
    expected cell edges. Private DOCX run `20260531-173508` was neutral (`15.889775` MAE, `0.141949`
    changed16), so the remaining border work is still the harder Word conflict ladder: adjacent-edge conflict
    resolution, start/end directionality, nil/none suppression across shared edges, and exact width/unit rules.
    2026-05-31 progress: logical `w:start`/`w:end` borders now participate in the default left-to-right DOCX
    table path. Cell-level logical borders render as physical left/right strokes, and table-level logical
    outer borders are resolved onto first/last cell edges when physical `left`/`right` are absent. Public
    coverage verifies both PDF stroke emission and table-border inheritance. Private-safe inventory found
    `start` borders in the current private DOCX; private run `20260531-232213` stayed neutral at `16/16` pages,
    zero dimension mismatches, MAE `13.648284`, changed16 `0.125542`. Keep the parent open for explicit
    right-to-left direction metadata, adjacent-edge conflict rules, and exact Word width/style handling.
    2026-05-31 progress: style-level table borders now use the same mapping path before direct table/cell
    overrides. Public `docx-tables --skip-slow` passed `27`; private DOCX run `20260531-174458` stayed neutral
    (`15.889775` MAE, `0.141949` changed16).
    2026-05-31 progress: removed the renderer fallback that drew a black rectangle around every table cell
    with no resolved `tcBorders`/`tblBorders`/style border. Absence of border data now emits no stroke, while
    explicit table borders keep using resolved edge strokes. Public coverage now separates explicit-border
    table rendering from a no-border table that must not invent a grid. Validation passed `docx-tables
    --skip-slow` (`44`), `docx-text --skip-slow` (`16`), `docx-page --skip-slow` (`14`), and full solution
    build. Private DOCX run `20260531-230652` stayed page-stable at `16/16`, zero dimension mismatches, and
    improved aggregate metrics from `MAE=14.806507`, changed16 `0.133867` to `MAE=13.648284`, changed16
    `0.125542`. Keep this item open for border conflict resolution, nil/none suppression, and exact width
    units; the completed part is only removing the invented default grid.
    2026-05-31 progress: corrected the DOCX border cascade so direct table-level `w:tblBorders` override
    table-style conditional cell borders per edge, while unaffected inherited style-table edges still survive.
    A direct `w:val="nil"` table edge now suppresses a lower-priority conditional edge instead of being
    overwritten by it. Public coverage locks direct table `nil`/right-edge overrides against first-row
    conditional `w:tcBorders`. Validation passed `docx-tables --skip-slow` (`45`) and full solution build.
    Private DOCX run `20260531-231047` stayed page-stable at `16/16`, zero dimension mismatches, and unchanged
    aggregate metrics (`MAE=13.648284`, changed16 `0.125542`), so this is a structural correctness fix rather
    than a current private raster win. Keep this item open for adjacent-edge conflict resolution,
    start/end directionality, and exact border width/unit rules.
    2026-06-01 progress: DOCX table border PDF emission now uses filled rectangle strips instead of stroked
    lines. Public Office PDF inspection of `docx-ladder-03-table-pagination-margins` showed Word emits table
    borders as nonzero fills with the authored border color, while the candidate used stroke operations. The
    renderer now emits top/bottom/left/right/start/end borders through `re f` strips in the resolved edge path,
    with public unit coverage checking that explicit and logical cell borders no longer produce `l S` strokes.
    Public visual run `20260601-023556` improved page 1 from `MAE=0.927605` to `0.888769` and page 2 from
    `MAE=0.253998` to `0.210261`; candidate PDF inspection shows no stroke graphics operations in that case.
    Keep the parent open: shared-edge conflict resolution, border precedence, nil/none suppression across
    adjacent cells, RTL start/end mapping, and exact width/style semantics still need a table-border model
    before emission rather than relying on per-cell paint order.
    2026-06-01 follow-up: vertical shared edges within a row now go through a small border-emission model
    instead of drawing both adjacent cell edges. A shared vertical edge is emitted once, the stronger width wins
    when both sides declare visible borders, and an explicit `nil`/`none` on either side suppresses that shared
    edge. Public unit coverage locks both the one-strip case and the nil-suppression case. This is deliberately
    not marked as full border conflict resolution: horizontal shared edges across rows, page-break/header-row
    repeated borders, full Word style precedence across adjacent cells, and RTL logical edge ownership remain
    open for Office-backed fixtures.
    2026-06-01 follow-up: horizontal shared edges between consecutive same-page table rows now use the same
    one-strip boundary model. The renderer resolves the current row's bottom edge against the next row's top
    edge, emits only one centered fill strip, and lets explicit `nil`/`none` on either side suppress the shared
    border. Public unit coverage now locks both shared `insideH` de-duplication and top-edge nil suppression.
    Private DOCX run `20260601-024417` improved slightly to `MAE=12.494853`, changed16 `0.116738`, while page
    count and diagnostics stayed stable. Keep the parent open for non-contiguous/page-break row boundaries,
    row/column spans, RTL ownership, and the full Word border conflict ranking beyond width-first selection.
    2026-06-01 follow-up: horizontal shared-edge matching now uses cell geometry overlap instead of row-local
    cell indexes. This keeps `gridSpan` row shapes from dropping part of an `insideH` boundary when one row has
    a spanning cell and the adjacent row has multiple cells. Public unit coverage locks a one-cell-over-two
    row boundary as two emitted segments. This closes the first span-related emission miss, but full merged-cell
    conflict resolution still remains open for vertical merges, non-rectangular spans, and page-break
    continuation boundaries.
    2026-06-01 discovery/progress: table/cell border values beyond `single`, `nil`, and `none` are now
    surfaced as `DOCX_TABLE_BORDER_STYLE` diagnostics instead of being silently flattened into solid filled
    strips. The diagnostic is scoped to `w:tblBorders`/`w:tcBorders` in document and style parts, with public
    coverage proving supported `single`/`nil`/`none` remain quiet and a style-level `double` border points at
    `/word/styles.xml`. Private DOCX run `20260601-025057` stayed stable (`16/16`, zero dimension mismatches,
    no diagnostics, `MAE=12.494853`, changed16 `0.116738`), so the next high-impact border work remains
    Office-backed rendering of non-single border styles and full conflict ranking rather than a private
    regression response.
  - [x] 2026-05-31: Applied DOCX table-style `w:tblCellMar` as inherited cell margins. Style-level table
    cell margins now merge with direct `w:tcMar` and feed the existing layout-owned cell text box calculation.
    Private impact was small but positive, and the implementation keeps margins in the same structural path as
    direct cell properties.
    2026-06-01 follow-up: the layout engine no longer invents a hard-coded horizontal cell padding when no
    margin source exists. Word's normal `5.4pt` inset is now represented structurally by the default table
    style (`w:style w:type="table" w:default="1"` / `TableNormal` with `w:tblCellMar`), while direct
    `w:tblPr/w:tblCellMar`, row `w:tblPrEx/w:tblCellMar`, and direct `w:tcPr/w:tcMar` merge through the same
    margin cascade before layout. This came from public `docx-ladder-03-table-pagination-margins` PDF
    inspection: the no-style/no-`tblCellMar` second row in Word starts at the cell content rectangle
    (`x=72.504`) while the old candidate used the baked-in padding (`x=77.88`). After the change candidate
    table text is structurally aligned (`x=72.48`, baseline `708.00` vs Word `72.504`, `707.74`; first-row
    explicit margin baseline `120.00` vs Word `119.66`). Public run `20260601-144030` improved pagination
    page 1 `MAE=0.878900834 -> 0.770108922` and page 2 `0.235116561 -> 0.171588872`. Guard runs:
    `docx-ladder-02-table-cell-margins` `20260601-144221` improved to `MAE=0.453874`; row-heights
    `20260601-144231` stayed at `MAE=0.700136`; paragraph-adjacency `20260601-144239` stayed at
    `MAE=0.545870`; `docx-tables --skip-slow` passed `73`. Private DOCX acceptance run `20260601-144414`
    stayed structurally stable (`16/16`, no dimension mismatches, no diagnostics) with a small aggregate
    raster move to `MAE=13.855991`, changed16 `0.127419`; keep the structural default-style cascade because
    it removes a hard-coded layout heuristic and matches the public Office PDF evidence.
  - [ ] 2026-05-31: Complete table margin/width fidelity: implement table-level width/preferred width,
    cell spacing, `tblInd`, grid spans/merged cells, and style/direct margin priority with public Office PDF
    fixtures.
    2026-05-31 progress: direct table indentation `w:tblInd` is now preserved as table-owned dxa metadata and
    applied in `DocxLayoutEngine` before placing row/cell layouts. Public tests cover source-token retention
    and table x-position shift. Private DOCX run `20260531-173800` was neutral (`15.889775` MAE, `0.141949`
    changed16), so the high-impact remainder is still cell spacing, grid spans/merged cells, and the full
    style/direct margin priority ladder.
    2026-05-31 progress: preserved `w:gridSpan` on table cells and applied it during row measurement and
    layout by summing the spanned grid columns. Public coverage checks source-token retention and layout
    width/x-position for a two-column span. Private DOCX run `20260531-180144` stayed neutral (`15.889775`
    MAE, `0.141949` changed16).
    2026-05-31 progress: preserved direct `w:tblCellSpacing` as table-owned dxa metadata and applied it as
    horizontal spacing between adjacent layout cell boxes. Public coverage checks source-token retention and
    x-position shift. Private DOCX run `20260531-181847` stayed neutral (`15.889775` MAE, `0.141949`
    changed16).
  - [x] 2026-05-31: Preserved DOCX table preferred-width tokens and applied `w:tblW w:type="dxa"` to table
    grid scaling. The layout stage now scales column widths to the preferred table width, capped by available
    page width, instead of relying only on raw `w:tblGrid` sums. Private impact was neutral, indicating this
    document's grid and preferred widths are already effectively aligned.
  - [x] 2026-05-31: Preserved DOCX cell preferred-width tokens and applied complete first-row
    `w:tcW w:type="dxa"` declarations as the table column basis before page-width scaling. Private-safe
    inventory showed every private table cell has `tcW`, and several first rows intentionally diverge from
    `tblGrid`; the change improved private aggregate metrics without changing page count.
  - [ ] 2026-05-31: Complete the DOCX table width ladder beyond the simple first-row `dxa` case: honor
    `tcW` across rows, auto and percentage widths, `tblLayout` fixed/autofit differences, `gridSpan`,
    `tblInd`, cell spacing, and Word's conflict resolution between `tblGrid`, `tblW`, and cell preferred
    widths with public Office PDF fixtures.
    2026-05-31 progress: layout now resolves `w:tblW w:type="pct"` from fiftieths of a percent against the
    available table width before scaling the grid. Public coverage checks a 50% table width; private DOCX run
    `20260531-182041` improved slightly (`15.864724` MAE, `0.141689` changed16).
    2026-05-31 progress: column-basis resolution now collects unambiguous single-column `tcW dxa`
    constraints across rows instead of only accepting a complete first row. Grid-spanning cells remain
    width consumers, not ambiguous column split sources. Public coverage checks a spanning first row followed
    by a later row with 40/80 point preferred cell widths; private DOCX run `20260531-185014` stayed neutral
    versus the current baseline (`15.849350` MAE, `0.141574` changed16), indicating the private document's
    affected rows were already covered by earlier width rules.
    2026-05-31 progress: layout now resolves single-column `tcW w:type="pct"` from fiftieths-of-a-percent
    values against the resolved table width instead of preserving the token without using it. Public coverage
    checks a 25%/75% preferred-cell split after body-width capping; private DOCX run `20260531-185336` stayed
    neutral (`15.849350` MAE, `0.141574` changed16), so the current private case does not appear to exercise
    pct cell widths in a visually material way.
    2026-06-01 progress: missing `w:tblGrid` inference now counts logical grid columns, including
    `w:gridSpan`, instead of using physical cell count. This keeps span+tail rows from overflowing by treating
    the spanned cell as two columns and the tail as the third. Public coverage includes
    `docx-ladder-03-table-missing-grid-spans` and reader/layout tests for both parser-inferred and
    layout-inferred empty grids. Office/candidate PDF inspection on run `20260601-145200` shows the tail
    column text starts at `X=216.48` candidate vs `X=216.53` reference, confirming the structural column
    ownership. Validation passed `docx-tables --skip-slow` (`76`), `docx-page --skip-slow` (`26`),
    `docx-text --skip-slow` (`36`), full solution build, and the public `docx-layout` family with `25`
    cases; private DOCX run `20260601-145516` stayed page-stable at `16/16` pages, zero dimension mismatches,
    no diagnostics, `MAE=13.855991`, changed16 `0.127419`. Keep this width parent open for `tblLayout`
    fixed/autofit differences, auto widths, cell spacing in the width equation, and Word's full conflict
    resolution between `tblGrid`, `tblW`, and `tcW`.
## Private Evidence

Private evidence is intentionally anonymized. Do not copy private text, screenshots, filenames, or
document-specific business content into public notes.

- Private PPTX text-emission inspection after run `artifacts/private-visual/lokad-value-based/20260530-192341`:
  - Page 81 has matching Office/candidate text-operation counts (`83/83`) but candidate `Tc` is still `0:83`
    while Office emits `0:30`, `-0.048:20`, `0.0173:13`, `-0.0535:11`, and `-0.024:9`.
  - The nonzero buckets are frame/paragraph/line clustered and are not explained by highlight, font family, or
    private text categories. Private-safe glyph category counts were added to the local inspection tooling to
    verify this without recording private text.
  - Page 79 remains table-heavy with Office nonzero `Tc` buckets and candidate `Tc=0:249`; because its text
    operation count is still `249/258`, use it as corroborating evidence after page-81-style one-to-one cases
    are understood.
  - A disposable public probe with generic text reproduced the same class of Office `Tc` decomposition, so the
    follow-up should be public-fixture-backed and structural rather than private-deck or font-name tuning.
- Private PPTX rerun `artifacts/private-visual/lokad-value-based/20260531-130027` after the stacked-column
  bottom-legend plot-box slice:
  - 84/84 pages compared with zero dimension mismatches and empty diagnostics.
  - The targeted chart page improved from the pre-slice run `20260531-124414` MAE `2.439032`, changed16
    `0.046995`, SSIM `0.931097` to MAE `1.702704`, changed16 `0.039896`, SSIM `0.960047`.
  - Focused PDF inspection shows the candidate chart axes now align to Office at about `137.71..296.69`
    versus `137.70..296.70`; the overlay connector strokes were already matching Office. The remaining chart
    gap is stacked-column fill structure: Office emits compound per-series paths, while OOXPDF still emits
    separate rectangles.
- Private PPTX rerun `artifacts/private-visual/lokad-value-based/20260531-131300` after the stacked-column
  compound-fill slice:
  - 84/84 pages compared with zero dimension mismatches and no diagnostics file.
  - The focused chart page improved again from run `20260531-130027` MAE `1.702704`, changed16 `0.039896`,
    SSIM `0.960047` to MAE `1.657135`, changed16 `0.039641`, SSIM `0.962169`.
  - The generic remaining slide-44 work is no longer per-segment stacked-column fill emission. The bottom
    legend probe still carries loose raster/text gates because it also exposes broader chart text and layout
    differences.
- Private PPTX rerun `artifacts/private-visual/lokad-value-based/20260531-141742` after chart text spacing
  was carried through the typed scene and renderer styles:
  - 84/84 pages compared with zero dimension mismatches and empty diagnostics.
  - Page 44 remained unchanged from run `20260531-140556` at MAE `1.281596258`, changed16 `0.032060185`,
    and SSIM `0.973682740`.
  - Focused PDF text inspection still shows Office using `Tc=-0.006` for part of the chart text, while the
    candidate mostly emits `Tc=0` for chart text. Inspecting the private source chart showed no nonzero
    authored `spc` in `chart7.xml` and only `spc=0` in the chart style. Treat this as evidence for the shared
    Office PDF text-emission profile, not for a slide-specific chart spacing rule.
- Private PPTX rerun `artifacts/private-visual/lokad-value-based/20260529-035838` before the image recolor
  renderer-boundary cleanup:
  - 84/84 pages compared with zero dimension mismatches.
  - Mean absolute error: `7.167206`; max mean absolute error: `16.511236`; mean changed-pixel ratio at
    threshold 16: `0.098106`.
  - Diagnostics were limited to two `PPTX_UNSUPPORTED_TEXT_OVERFLOW` warnings and one
    `PPTX_UNSUPPORTED_IMAGE_RECOLOR` warning.
  - The recolor warning remains a generic JPEG/DCT picture-recolor architecture gap. No private slide text,
    image content, or screenshots were inspected or copied.
- Private PPTX rerun `artifacts/private-visual/lokad-value-based/20260530-151818` after the chart
  tiling-pattern resource slice:
  - 84/84 pages compared with zero dimension mismatches.
  - Mean absolute error: `3.043284`; max mean absolute error: `6.045372`; mean changed-pixel ratio at
    threshold 16: `0.054661`.
  - Diagnostics were empty.
  - The targeted stacked-bar page now uses PDF `/Pattern` resources for diagonal chart fills instead of
    page-level hatch strokes. Office's reference PDF uses image-backed type-1 tiling patterns with a `16x16`
    pattern cell and `0.375` scale matrix; OOXPDF now matches the tiling-pattern resource class and scale, but
    still uses vector pattern content and has residual chart rectangle geometry offsets. This remains an open
    structural chart/PDF-resource gap, not a private-content note.
- Private PPTX rerun `artifacts/private-visual/lokad-value-based/20260530-152925` after the opposite-side
  chart value-axis reserve slice:
  - 84/84 pages compared with zero dimension mismatches.
  - Mean absolute error: `3.009260`; max mean absolute error: `6.045372`; mean changed-pixel ratio at
    threshold 16: `0.054138`.
  - Diagnostics were empty.
  - The targeted stacked-bar page improved from MAE `4.591801` to `3.849802` because the plot area no longer
    applies the same-side multi-axis label reserve to a chart whose value-axis labels live on opposite sides.
    The candidate patterned-stack X/width is now close to Office; remaining evidence points to vertical
    plot/value-scale residuals and image-backed pattern cells.
- Private PPTX rerun `artifacts/private-visual/lokad-value-based/20260528-141612` after the text-height
  estimator and schema-inspection pass:
  - 84/84 pages compared with zero dimension mismatches.
  - Mean absolute error: `6.731583`; max mean absolute error: `15.165670`; mean changed-pixel ratio at
    threshold 16: `0.093688`.
  - Diagnostics remain limited to one `PPTX_UNSUPPORTED_IMAGE_RECOLOR`.
  - Private-safe inspection of the affected middle-anchored text frame found 3 paragraphs, 4 runs, and
    40 text code units. The frame model reported `TextHeight=347.862047`, `VerticalOffset=141.516766`,
    and line baselines at `296.748162` and `275.148162` using OS/2 `usWinAscent` metrics. Compared with
    Office PDF text operations on the same page, the vertical text drift is now about `0.35 pt`, down from
    roughly `2.09 pt` before the font-box rule.
  - The remaining local mismatch is no longer the frame vertical anchor: the rightmost same-line text segment
    still differs by about `3.63 pt` horizontally and by decoded text-operation length. This is recorded as a
    generic glyph advance/segmentation follow-up, not as private-content tuning.
- Private PPTX rerun `artifacts/private-visual/lokad-value-based/20260527-223838` after the chart enum/fallback
  boundary cleanups:
  - 84/84 pages compared with zero dimension mismatches.
  - Mean absolute error: `7.702155`; max mean absolute error: `16.412422`; mean changed-pixel ratio at
    threshold 16: `0.103230`.
  - Diagnostics remain limited to one `PPTX_UNSUPPORTED_IMAGE_RECOLOR`.
  - Worst-page public-safe metric summary was pages 53, 32, 50, 36, and 49 by MAE; no private slide content
    was copied into this plan.
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
- Private DOCX rerun `artifacts/private-visual/user-requirements-spec/20260531-152007` after typed
  `keepLines`/`keepNext` pagination:
  - Reference output had 16 pages; candidate output had 16 pages; all compared page dimensions matched.
  - Paired-page MAE was `17.339630`, and mean changed-pixel ratio at threshold 16 was `0.151959`.
  - Diagnostics remain `DOCX_NUMBERING_INDENT`, `DOCX_STYLE_PARAGRAPH_KEEP_RULE`,
    `DOCX_STYLE_PARAGRAPH_SPACING`, `DOCX_STYLE_TABLE_STYLE`, `DOCX_UNSUPPORTED_TABLE_HEADER_ROW`, and
    `DOCX_UNSUPPORTED_TABLE_STYLE`.
  - This closes the coarse page-count mismatch but not pixel-level fidelity; remaining DOCX work should focus
    on table style/header behavior, numbering indents, exact paragraph spacing, and line-level widow/orphan
    decisions.
- Private DOCX rerun `artifacts/private-visual/user-requirements-spec/20260531-152705` after preserving and
  applying numbering-level indents:
  - Reference output had 16 pages; candidate output had 16 pages; all compared page dimensions matched.
  - Paired-page MAE improved to `16.995226`, and mean changed-pixel ratio at threshold 16 improved to
    `0.149573`.
  - Diagnostics remain the same six public-safe categories; `DOCX_NUMBERING_INDENT` is still intentionally
    open because the current implementation shifts flattened list lines rather than modeling Office label
    tab stops and continuation-line hanging geometry.
- Private DOCX rerun `artifacts/private-visual/user-requirements-spec/20260531-153115` after preserving and
  repeating DOCX table header rows:
  - Reference output had 16 pages; candidate output had 16 pages; all compared page dimensions matched.
  - Aggregate metrics were unchanged from the numbering-indent run: MAE `16.995226`, changed16 `0.149573`.
  - This suggests the private table residual is not primarily omitted repeated header rows; the remaining
    table work should focus on table style resolution, conditional formatting, borders, shading, and width
    inheritance.
- Private DOCX rerun `artifacts/private-visual/user-requirements-spec/20260531-153536` after preserving table
  style IDs and applying whole-style cell shading:
  - Reference output had 16 pages; candidate output had 16 pages; all compared page dimensions matched.
  - Aggregate metrics were unchanged from the prior table-header run: MAE `16.995226`, changed16 `0.149573`.
  - The private table-style gap is therefore unlikely to be explained by a simple style-level cell fill alone;
    prioritize conditional table-style regions, borders, margins, and widths.
- Private DOCX rerun `artifacts/private-visual/user-requirements-spec/20260531-153922` after conditional
  table-style shading:
  - Reference output had 16 pages; candidate output had 16 pages; all compared page dimensions matched.
  - MAE improved from `16.995226` to `16.817625`; changed16 improved from `0.149573` to `0.148589`.
  - This confirms conditional table-style resolution is present in the private document and should continue
    into borders, margins, `tblLook`, and conditional priority rules.
- Private DOCX rerun `artifacts/private-visual/user-requirements-spec/20260531-154335` after conditional
  table-style borders and per-edge border emission:
  - Reference output had 16 pages; candidate output had 16 pages; all compared page dimensions matched.
  - MAE improved from `16.817625` to `16.004348`; changed16 improved from `0.148589` to `0.142372`.
  - The worst page shifted away from the previously table-dominated page bucket, confirming table border
    structure is one of the dominant remaining private DOCX differences.
- Private DOCX rerun `artifacts/private-visual/user-requirements-spec/20260531-154620` after contextual
  paragraph spacing:
  - Reference output had 16 pages; candidate output had 16 pages; all compared page dimensions matched.
  - Aggregate metrics were unchanged from the table-border run: MAE `16.004348`, changed16 `0.142372`.
  - The contextual-spacing branch is implemented for correctness, but it is not a dominant private metric
    driver for this document; continue prioritizing table style/border and numbering geometry.
- Private DOCX rerun `artifacts/private-visual/user-requirements-spec/20260531-154933` after table-style cell
  margins:
  - Reference output had 16 pages; candidate output had 16 pages; all compared page dimensions matched.
  - MAE moved slightly from `16.004348` to `16.003190`; changed16 moved from `0.142372` to `0.142369`.
  - Table-style margins are structurally supported now, but the remaining private table gap is dominated by
    other style, width, and text geometry.
- Private DOCX rerun `artifacts/private-visual/user-requirements-spec/20260531-155308` after applying `dxa`
  table preferred widths:
  - Reference output had 16 pages; candidate output had 16 pages; all compared page dimensions matched.
  - Aggregate metrics were unchanged from the table-margin run: MAE `16.003190`, changed16 `0.142369`.
  - `w:tblW` is now structurally handled for `dxa`, but it is not a dominant private metric driver here.
- Private DOCX rerun `artifacts/private-visual/user-requirements-spec/20260531-155930` after applying complete
  first-row `dxa` cell preferred widths:
  - Reference output had 16 pages; candidate output had 16 pages; all compared page dimensions matched.
  - MAE improved from `16.003190` to `15.928048`; changed16 improved from `0.142369` to `0.142037`.
  - The remaining worst pages are still table-heavy, so continue the width/layout ladder before treating table
    styling as complete.
- Private DOCX rerun `artifacts/private-visual/user-requirements-spec/20260531-161037` after separating
  DOCX body numbering markers from paragraph text:
  - Reference output had 16 pages; candidate output had 16 pages; all compared page dimensions matched.
  - MAE improved from `15.928048` to `15.888467`; changed16 improved from `0.142037` to `0.141944`.
  - The numbering diagnostic remains open because exact Word tab-stop, bullet-font, and continuation-line
    behavior is not yet modeled.
- Private DOCX rerun `artifacts/private-visual/user-requirements-spec/20260531-162451` after preserving plain
  `w:br` soft line breaks:
  - Reference output had 16 pages; candidate output had 16 pages; all compared page dimensions matched.
  - MAE moved slightly from `15.888467` to `15.889775`; changed16 moved from `0.141944` to `0.141949`.
  - Keep the feature because dropping authored soft breaks is structurally wrong; exact break spacing remains
    part of the text layout ladder.
- Private DOCX rerun `artifacts/private-visual/user-requirements-spec/20260531-162756` after preserving
  authored whitespace during wrapping:
  - Reference output had 16 pages; candidate output had 16 pages; all compared page dimensions matched.
  - Aggregate metrics were unchanged from the soft-break run: MAE `15.889775`, changed16 `0.141949`.

## Backlog

### Release-Blocking Fidelity

- [ ] Extend PPTX chart rendering beyond the current loose native renderer: labels, legends, axes,
  marker styles, theme/chart style colors, and tighter plot-area layout fidelity.
- [ ] Fix DOCX page geometry and pagination fidelity: section page size/margins, paragraph spacing, manual
  page/column breaks, and keep/widow page-break decisions.
- [ ] Text layout: preserve spaces, tabs, line breaks, soft line breaks, kerning-like advances, font
  fallback, mixed run spacing, character spacing, superscript/subscript, and baseline offsets.
- [x] 2026-06-01: Preserved DOCX run-level `w:vertAlign` tokens through the resolved run-property cascade,
  character-style inheritance, `DocxTextRunStyle`, and emitted `DocxTextRun` records. Public coverage checks
  direct `subscript`, inherited `superscript`, and null source tokens; validation passed `docx-text
  --skip-slow` (`20`) and full solution build. Private DOCX run `20260601-085143` stayed page-stable at
  `16/16`, zero dimension mismatches, no diagnostics, `MAE=13.388935`, changed16 `0.124264`, worst page 9
  `17.162084`. Keep superscript/subscript layout open for an Office-backed ladder that derives Word's exact
  font-size scaling, baseline shifts, underline/highlight interaction, and line-box impact from reference PDFs
  instead of adding a guessed offset.
- [x] 2026-06-01: Moved DOCX text segment rendering toward a run-resolved layout contract instead of
  line-wide text emission. `DocxTextSegmentLayout` now carries an optional resolved font size and baseline
  offset; the renderer consumes those values for glyphs, backgrounds, decorations, and terminal line spaces.
  Follow-up, 2026-06-01: body, table, and justified text measurement now use each text span's resolved
  run font size for wrapping, advances, and emitted `DocxTextSegmentLayout.FontSize`, while paragraph line
  boxes still use the maximum run size for baseline and line-height decisions. Public unit coverage now
  proves that mixed-size runs keep their own measured widths and that a small run followed by a larger run
  does not wrap as if the whole paragraph used the maximum font size. This closes the body/table flattening
  gap without adding font-name or token-string conditions. Static header/footer segments already pass each
  run's authored size, matching how those segments are measured and avoiding a hidden renderer-side
  flattening step. This is intentionally only the structural landing zone for `w:vertAlign`:
  superscript/subscript scaling and shifts remain open until a public Office-backed ladder derives the Word
  behavior from PDF evidence rather than guessed constants.
  Follow-up, 2026-06-01: added public `docx-ladder-02-vertical-align` and implemented the first DOCX
  `w:vertAlign` layout rule from Office PDF evidence rather than font-name logic. Superscript/subscript
  segments now resolve through the same run-segment contract as mixed font sizes: Word-style reduced
  half-point-grid font sizes, per-segment baseline offsets, matching body/table/static-header measurement,
  and terminal line spaces emitted at the line baseline rather than inheriting the final shifted segment.
  The new ladder improved from `MAE=0.363558`, changed16 `0.004239`, SSIM `0.766175` before the layout
  rule to `MAE=0.121473`, changed16 `0.002011`, SSIM `0.922178` at run `20260601-161408`. PDF inspection
  shows sampled vertical-align glyph sizes and matrices now line up structurally, for example `2` at
  `10.56pt`, `Y=710.50` reference vs `10.56pt`, `Y=710.46` candidate; `sup` at `10.56pt`, `Y=646.54`
  reference vs `10.56pt`, `Y=646.71` candidate. Keep exact DOCX text-state spacing and residual
  decoration/background rectangle interaction open; the vertical-align font-size/baseline layer is no
  longer blocked on a missing public ladder. Validation passed `docx-text --skip-slow` (`41`),
  `docx-page --skip-slow` (`29`), `docx-tables --skip-slow` (`77`), full solution build, and the public
  `docx-ladder-02-vertical-align` visual case.
  Validation passed `docx-page --skip-slow` (`29`), `docx-text --skip-slow` (`40`), `docx-core --skip-slow`
  (`23`), `docx-tables --skip-slow`
  (`77`), and `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal`. Public `docx-headers-footers`
  run `20260601-155239` stayed unchanged (`MAE=0.073352`, changed16 `0.002110`, SSIM `0.982572`). Private
  DOCX run `20260601-155830` stayed page-stable at `16/16`, zero dimension mismatches, no diagnostics,
  `MAE=13.666634`, changed16 `0.126275`.
- [x] 2026-06-01: Preserved DOCX run-level strike tokens through the same run-property cascade:
  `w:strike`, explicit off values such as `w:val="0"`, and inherited `w:dstrike` now survive into
  `DocxTextRun`/`DocxTextRunStyle` with source token values. Validation passed `docx-text --skip-slow`
  (`21`) and full solution build. Private DOCX run `20260601-085424` stayed page-stable at `16/16`, zero
  dimension mismatches, no diagnostics, `MAE=13.388935`, changed16 `0.124264`, worst page 9 `17.162084`.
  Keep strikethrough drawing open for a public Office PDF ladder that derives the actual line rectangle
  position, thickness, double-strike separation, and interaction with baseline shifts and run boundaries before
  changing renderer geometry.
  2026-06-01 follow-up: underline rendering no longer uses the old `fontSize * 0.12` / `fontSize / 18`
  heuristic, and strike/double-strike tokens now render. Body and static header/footer text decorations use
  OpenType `post` underline metrics and `OS/2` strikeout metrics from the resolved run font, emitting filled
  rectangles like Office rather than stroked heuristic lines. Public `docx-ladder-02-text-decorations` run
  `20260601-150426` is page-stable with no diagnostics (`MAE=0.229168`, changed16 `0.002961`); PDF inspection
  shows close Office/candidate decoration boxes, for example underline `Y=702.10..703.30` reference vs
  `702.68..703.85` candidate. Validation passed `docx-text --skip-slow` (`37`), `docx-page --skip-slow`
  (`26`), full solution build, and the public `docx-layout` family with `26` cases. Private DOCX run
  `20260601-150717` stayed page-stable at `16/16` pages, zero dimension mismatches, no diagnostics,
  `MAE=13.855991`, changed16 `0.127419`. Keep this open for baseline-shift interaction, wrapped mixed-run
  decoration segmentation, and exact double-strike separation across fonts.
- [x] 2026-06-01: Preserved DOCX run-level highlight and run-shading tokens through the resolved
  run-property cascade. `w:highlight @w:val`, inherited character-style highlights, and `w:rPr/w:shd`
  `fill`/`val`/`color` now survive into `DocxTextRun` and `DocxTextRunStyle`. Validation passed
  `docx-text --skip-slow` (`22`) and full solution build. Private DOCX run `20260601-085653` stayed
  page-stable at `16/16`, zero dimension mismatches, no diagnostics, `MAE=13.388935`, changed16 `0.124264`,
  worst page 9 `17.162084`.
- [x] 2026-06-01: Rendered DOCX run backgrounds for `w:highlight` and run `w:shd` from the resolved
  run model instead of dropping the parsed tokens. Named Word highlight values now map to their OOXML
  RGB colors, `w:shd w:val="clear"` uses the fill color, and percentage shading such as `pct20`
  blends foreground `w:color` with background `w:fill` to match the Office PDF evidence (`D9EAD3`
  plus `112233` at 20% becomes `B1C2B3`). Background rectangles use the resolved run font's Windows
  ascender/descender line box and are emitted before text/decorations for body text and static
  headers/footers. Added public `docx-ladder-02-text-backgrounds` plus unit coverage for highlight,
  clear shading, and percentage shading. Public run `20260601-151700` improved from the first
  highlight-only attempt (`MAE=0.726385`, changed16 `0.009851`) to `MAE=0.481045`, changed16
  `0.006061`; PDF inspection shows the same fill colors and near-matching background rectangles
  with residual height difference around `18.384pt` Office vs `17.875pt` candidate on 16pt text.
  Validation passed `docx-text --skip-slow` (`38`), full solution build, the public `docx-layout`
  family with `27` cases, and private DOCX run `20260601-151717` stayed page-stable at `16/16`,
  zero dimension mismatches, no diagnostics, `MAE=13.855991`, changed16 `0.127419`. Keep this open
  for exact Word run-background rectangle height/vertical offset, non-percent pattern families,
  `auto` foreground/background resolution, wrapped runs, and mixed baseline-shift interactions.
- [x] 2026-06-01: Preserved DOCX run-level `w:smallCaps` tokens through direct run properties,
  character-style inheritance, `DocxTextRunStyle`, and emitted `DocxTextRun` records. Public coverage checks
  explicit off, inherited val-less on, and missing-token behavior; validation passed `docx-text --skip-slow`
  (`27`) and full solution build. Private DOCX run `20260601-091708` stayed page-stable at `16/16`, zero
  dimension mismatches, no diagnostics, `MAE=13.388935`, changed16 `0.124264`, worst page 9 `17.162084`.
  Keep small-caps rendering open for Office-backed glyph casing, size scaling, baseline interaction, and PDF
  text-operation evidence rather than treating it like ordinary all-caps.
- [x] 2026-06-01: Preserved DOCX run-level `w:vanish` hidden-text tokens and suppress hidden runs from body
  layout and static header/footer rendering. Direct off values such as `w:val="0"` keep the run visible, while
  inherited val-less hidden runs remain in the source-bearing model but do not emit visible text. Validation
  passed `docx-text --skip-slow` (`30`) and full solution build. Private DOCX run `20260601-092656` stayed
  page-stable at `16/16`, zero dimension mismatches, no diagnostics, `MAE=13.388935`, changed16 `0.124264`,
  worst page 9 `17.162084`. Keep hidden-text mode options, hidden drawings/fields, paragraph spacing edge
  cases for hidden-only paragraphs, and Office-backed diagnostics open.
- [x] 2026-06-01: Preserved cached DOCX `w:fldSimple` result runs in paragraph order for non-PAGE fields
  instead of dropping them. The reader now walks paragraph child elements in document order, keeps cached
  simple-field result text and run styles, and still maps simple PAGE fields to the existing page placeholder.
  Validation passed `docx-text --skip-slow` (`23`) and full solution build. Private DOCX run
  `20260601-090014` stayed page-stable at `16/16`, zero dimension mismatches, no diagnostics,
  `MAE=13.388935`, changed16 `0.124264`, worst page 9 `17.162084`. Keep complex-field instruction/result
  range handling open: `w:fldChar` begin/separate/end runs still need a first-class field model so cached
  results, dynamic fields, and unsupported field diagnostics are separated without relying on instruction-text
  shortcuts.
- [x] 2026-05-31: Preserved DOCX plain `w:br` soft line breaks in run text and layout wrapping. The reader now
  keeps authored soft breaks as line separators, and the layout stage wraps each segment onto separate
  baselines instead of silently concatenating text across the break.
- [ ] 2026-05-31: Complete DOCX break and whitespace fidelity: distinguish soft line breaks from page/column
  breaks, tabs, preserved spaces, empty break-only lines, and Office's exact line-break baseline advance.
  - [x] 2026-06-01: Promoted DOCX run-level page-break-only paragraphs (`w:r/w:br w:type="page"`) to
    `DocxPageBreakElement` without attempting mixed-content paragraph splitting. This handles the safe
    structural case that was previously dropped as an empty paragraph while preserving the open item for
    inline page/column breaks inside text-bearing paragraphs. Validation passed `docx-page --skip-slow`
    (`18`) and `docx-core --skip-slow` (`19`); private DOCX run `20260601-032728` stayed unchanged at `16/16`
    pages, zero dimension mismatches, no diagnostics, `MAE=13.047852`, changed16 `0.120850`.
  - [x] 2026-06-01: Preserved ordinary DOCX `w:tab` tokens as `\t` in run text and expanded them in the
    layout stage to Word's default 36 pt tab-stop grid instead of measuring or emitting a tab glyph. The
    visible glyph runs are now separate `DocxTextSegmentLayout` records around the structural tab advance,
    so PDF text output stays glyph-only while layout width and segment positions observe the tab. Validation
    passed `docx-text --skip-slow` (`19`), `docx-numbering --skip-slow` (`11`), `docx-tables --skip-slow`
    (`65`), `docx-page --skip-slow` (`18`), and `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal`;
    private DOCX run `20260601-084645` stayed page-stable at `16/16`, zero dimension mismatches, no
    diagnostics, `MAE=13.388935`, changed16 `0.124264`. Keep the parent item open for authored
    paragraph/style tab-stop lists (`w:tabs/w:tab`), leader/alignment variants, and exact Office trimming
    around tab/space boundaries.
  - [x] 2026-06-01: Added the first authored DOCX paragraph tab-stop layer. `w:pPr/w:tabs/w:tab` now produces
    typed `DocxTabStop` records with position, kind, and leader source tokens; layout and wrapping advance
    ordinary `w:tab` characters to the next authored position before falling back to Word's default 36 pt
    grid. Public coverage checks a 1440-twip authored left tab stop and source token preservation; validation
    passed `docx-text --skip-slow` (`24`), `docx-numbering --skip-slow` (`11`), `docx-tables --skip-slow`
    (`65`), `docx-page --skip-slow` (`18`), and full solution build. Private DOCX run `20260601-090624`
    stayed page-stable at `16/16`, zero dimension mismatches, no diagnostics, `MAE=13.388935`, changed16
    `0.124264`, worst page 9 `17.162084`. Keep the parent item open for tab-stop clearing,
    right/center/decimal alignment, leader glyph emission, style/list tab merging, and exact indent-relative
    coordinate rules from Office PDF probes.
  - [x] 2026-06-01: Preserved explicit DOCX hyphen run tokens: `w:noBreakHyphen` now survives as U+2011 and
    `w:softHyphen` survives as U+00AD instead of being dropped by the run-text reader. Validation passed
    `docx-text --skip-slow` (`25`) and full solution build. Private DOCX run `20260601-090842` stayed
    page-stable at `16/16`, zero dimension mismatches, no diagnostics, `MAE=13.388935`, changed16
    `0.124264`, worst page 9 `17.162084`. Keep the parent item open for Office-backed soft-hyphen layout
    semantics: conditional glyph emission only at selected line breaks, no-break behavior in wrapping, and PDF
    text operation evidence for visible versus hidden discretionary hyphens.
  - [x] 2026-06-01: Made DOCX text wrapping distinguish breakable whitespace from nonbreaking space code
    points. U+00A0, U+202F, and U+2007 no longer create wrap tokens, so nonbreaking spans stay together while
    ordinary spaces remain break opportunities. Validation passed `docx-text --skip-slow` (`26`) and full
    solution build. Private DOCX run `20260601-091229` stayed page-stable at `16/16`, zero dimension
    mismatches, no diagnostics, `MAE=13.388935`, changed16 `0.124264`, worst page 9 `17.162084`. Keep exact
    Word trimming/carryover of spaces at wrap boundaries open for Office PDF probes.
- [x] 2026-05-31: Made DOCX text wrapping preserve authored whitespace tokens instead of splitting with
  `RemoveEmptyEntries`. Leading, trailing, and repeated spaces now remain in layout line text and contribute
  to measured line width; tab stops and Word's exact whitespace trimming rules remain open.
- [x] 2026-06-01: Promoted DOCX paragraph `w:jc w:val="both"` to a first-class
  `DocxTextAlignment.Justified` value instead of collapsing the typed alignment to `Left` while only keeping
  the raw token. Validation passed `docx-text --skip-slow` (`26`) and full solution build. Private DOCX run
  `20260601-091420` stayed page-stable at `16/16`, zero dimension mismatches, no diagnostics,
  `MAE=13.388935`, changed16 `0.124264`, worst page 9 `17.162084`. Keep actual justification layout open for
  Office-backed inter-word/character expansion, last-line handling, tabs, mixed runs, and PDF `TJ`/spacing
  evidence.
- [ ] Text emission: derive Office's implicit PDF `Tc` text-state behavior for presentation text where OOXML
  has no explicit `a:rPr @spc`. Private pages 24, 36, 39, and 48 show Office exporting nonzero positive or
  negative `Tc` while OOXPDF currently encodes equivalent or residual advances only through `TJ` positioning.
  Do not add a private-deck threshold or a font-name exception; build a public Office-authored probe that
  compares decoded text matrices, `Tc`, `TJ`, and font-size-grid choices for the same layout.
- [ ] Text layout: replace the temporary default-spacing empty-paragraph middle-anchor estimate with a
  structural Office rule. Public Office probes show visible lines use the resolved OS/2 Windows font box,
  while a trailing empty paragraph sits between the normal paragraph advance and the same font box. The
  current midpoint is an evidence-preserving approximation only; collect more public probes across fonts,
  `endParaRPr`, explicit `lnSpc`, and empty-paragraph counts before promoting a final rule.
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
- [ ] Charts: cached chart images, chart XML rendering, secondary-axis scale binding beyond the first
  bar-first line combo path, combo labels, non-default-axis-side label evidence, legends, series styling,
  grouped/stacked families, line charts, combo charts, and embedded chart data.
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

- [ ] Ladder 4: styled text runs: bold, italic, underline, color, highlight, mixed fonts, bullet glyphs,
  bullet hanging indents, paragraph spacing, and line spacing.
- [ ] Add normalized typography rungs for Office-authored kerning words by font family: Arial, Aptos/Calibri,
  Cambria/Cambria Math, and Segoe UI.
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
- [ ] Reverse-engineer Office's `spAutoFit` text fitting for narrow PPTX columns. The refined probe shows
  Office applying about `-0.036pt` character spacing to 12pt body lines, which changes wraps around short
  highlighted runs; candidate currently emits `Tc=0`.
- [ ] Tighten the slide-3-inspired Cambria probes next: compare Office and candidate line breaks, `TJ`
  arrays, highlight rectangles, and font metrics before rerunning the private deck.
- [ ] Add normalized typography rungs for Office PDF text-object structure: compare candidate `TJ` arrays
  and text matrices against Office for simple lines before accepting near-pixel raster gates.
- [ ] Ladder 4 remaining combined-stack gaps are finer glyph/font details after basic bullet font selection.
- [ ] Ladder 6: preset and connector shapes cover arrows, connector endpoints, arrowheads, dashes,
  line caps/joins, callouts, and common freeform/custom path fallbacks.
- [ ] Ladder 6 remaining subcases should isolate additional visual callout rendering and other preset
  geometries.
- [ ] Ladder 8: grouped content covers nested group transforms, grouped pictures, grouped text, grouped shapes,
  child coordinate scaling, z-order, and clips.
- [ ] Ladder 8 remaining subcases should isolate z-order with charts and clipping.
- [ ] Ladder 9: slide inheritance covers placeholders, master/layout text styles, hidden placeholders,
  footer/date/slide number placeholders, theme fonts, and theme color transforms.
- [ ] Ladder 10: tables cover fixed grid, per-edge borders, fills, cell margins, vertical alignment, merged
  cells, rich text inside cells, and table styles.
- [ ] Ladder 10 remaining subcases should isolate broader table style variants.
- [ ] Ladder 11: charts: cached image fallback, basic bar/line/pie rendering, axes, labels, legends, series
  styles, stacked/grouped variants, and chart diagnostics.
- [ ] Ladder 12: effects and advanced fills cover transparency, gradients, pattern fills, shadows, glows, soft
  edges, picture fills, and explicit diagnostics for unsupported effects.
- [ ] Ladder 12 remaining work is visual rendering for each effect/fill family.
- [ ] For every ladder rung, keep public synthetic fixture content artificial and minimal. Do not derive
  fixture text, images, layout, or styling from private documents.
- [ ] Run the relevant public visual case after each rung change; run private PPTX only as feature-discovery
  smoke evidence until the public ladder is much more complete.
- [ ] Revisit PPTX unit tests under the Office-PDF-first workflow: keep parser/safety/API tests, keep
  deterministic low-level PDF writer tests, but replace brittle renderer operator-position assertions with
  public visual gates or assertions derived from inspected Office reference geometry.

### PPTX Private Deck Recovery Plan

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
- [ ] Private slide 5 visible remaining problem: the right-side chart has an incorrectly placed secondary
  value axis and an incorrectly placed upward green arrow.
  Inspect whether this is a value-axis title, rotated axis label text, tick-label formatting, or chart-style
  inheritance, line-width/theme inheritance, or overlay transform geometry, then reproduce with a minimal
  public chart-axis/overlay fixture before changing renderer logic. The separate thin graph-arrow branch is
  closed by the public straight-stealth connector fixture and Office's 6 pt minimum straight-line stealth
  marker size.
  - [ ] Replace the vector approximation inside chart tiling-pattern cells with Office-like image-backed
    pattern-local XObjects, and derive the pattern matrix phase/translation from chart/shape coordinates
    instead of using a zero-offset cell matrix.
  - [ ] Extend combo/multi-axis chart support beyond the first bottom-up slice: bind each chart group to its
    referenced axes, honor axis crossing/orientation, keep primary/secondary scales independent, and place
    non-axis overlays such as the private slide 5 upward green arrow with Office-equivalent transforms.
  - [ ] Private slide 25 repeats the remaining multi-axis chart gap: the left schema's right-side value-axis
    labels and the overlaid green upward arrow are still vertically off. Treat this as the same generic
    secondary-axis/overlay-transform problem as slide 5, not as a static chart fallback.
- [ ] Private slide 6 visible remaining problem: a centered line of text is vertically misaligned inside its
  grey box. Reproduce with a public text-box fixture covering vertical anchor, body insets, line height, and
  shape fill/stroke context.
- [ ] Private slide 10 visible remaining problem: one headline line is positioned too high. Map it to public
  typography baselines, paragraph spacing, inherited bodyPr insets, or placeholder geometry before fixing.
- [ ] Private slide 11 visible remaining problem: a lower-left image overlaps preceding text, and adjacent
  left-side title text formatting mismatches Office. Reproduce as public picture/text z-order and inherited
  text-style fixtures.
  - [ ] Latest inspection suggests the overlap is caused by preceding text wrapping/flowing too tall, not by
    the timeline graphic being placed too high. Isolate with public typography fixtures for requested-font
    metrics, mixed bold/regular run widths, and default line advance before moving the graphic.
  - [ ] Current layout diagnostics after the break-space wrapping fix show the text frame ending above the
    timeline graphic, while the graphic matrix itself already matches the Office PDF. Regenerate the private
    visual case and keep this open for any remaining title/font/style mismatch.
- [ ] Private slide 19 visible remaining problem: lower-left logo rendering, paired-arrow geometry, and
  center/right font selection differ from Office. Split into public logo/image recolor or crop diagnostics,
  arrow geometry fixtures, and font fallback/style inheritance fixtures.
- [ ] Private slide 23 visible remaining problem: top text block formatting is off and an emphasized fragment
  overlaps the horizontal separator. Reproduce with public mixed-run typography, paragraph spacing, and
  separator/z-order fixtures.
- [ ] Private slide 30 visible remaining problem: the lower center/right date schema is badly off, likely a
  geometry or grouped-transform issue. Inventory shapes/connectors/transforms and isolate public geometry
  fixtures before changing renderer logic.
- [ ] Private slide 61 visible remaining problem: a shape-built line graph uses non-circular `prstGeom arc`
  shapes with `stealth` line ends, not chart XML. The preset-arc visual-angle conversion, OOXML Y-down arc
  basis, filled line-end outline emission, and Office default `adj1`/`adj2` values for empty public
  `avLst` arcs are now public-covered. Private slide 61 stays at about MAE `2.72` after the default-guide
  fix because its arc guides are explicit; the fix instead closes a public Office-authored arc collapse where
  missing guide fallbacks were interpreted as raw degrees. The remaining acceptable follow-up is to derive
  Office's exact preset-arc stroke-widening and flattening/segment cardinality from public reference PDFs.
  Public COM probe evidence shows Office emits filled arc outlines with more line segments than the current
  smooth-normal outline for default Office arcs (`185/86/86/101/101` fill segments in the reference versus
  `103/49/49/117/121` candidate on the disposable probe), while private slide 61 explicit arcs still show
  per-arc segment-count mismatches after their bounds and fill/stroke kinds align. Do not replace this with a
  private-slide coordinate shortcut or a per-font/per-shape special case.
  2026-05-31 follow-up: a scoped trial using a larger `stealth` marker geometry for all preset-arc filled
  outlines improved the public default-arc probe (MAE `0.0148676 -> 0.0089822`, SSIM `0.987148 -> 0.996857`)
  and matched Office's wider marker base on that probe, but it slightly worsened private slide 61
  (`2.7182668 -> 2.7209587` MAE) because the private explicit-guide arc bounds were already closer with the
  existing marker size. Reverted the trial. The remaining work is to find the OOXML/PDF structural
  discriminator for default-arc versus explicit-guide arc line-end sizing and flattening before changing
  production marker geometry.
  2026-05-31 public guardrail: added `pptx-ladder-06-explicit-arc-stealth`, a synthetic Office-referenceable
  fixture for the slide-61 shape family. It passes visually at low MAE, but PDF inspection still shows Office
  and candidate path decomposition drift (`34/66/66/20/34/62/62` filled-path line counts versus
  `51/73/65/39/51/77/75`). Use this case to validate any structural arc-flattening or marker-size change
  before rechecking the private slide.
- [ ] Private slide 9 visible remaining problem: left-side schema geometry is visibly broken. Survey the
  involved shapes/connectors/group transforms on public-safe diagnostics, then reproduce with minimal public
  geometry fixtures before changing renderer logic.
  - [ ] New review note: right-side text expands enough to overlap the lower boxing line, and a one-line
    label wraps into two lines. Map this to text metrics, autofit, body insets, and shape bounds with public
    text-fit fixtures.
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
  - [ ] Continue slide-12 parity with public fixtures for exact table vertical alignment, cell text anchoring,
    image placement, and any remaining bottom-right clipping/overlap after table wrapping.
- [ ] Private slide 13 visible remaining problem: text overflows on the bottom-right content. Reproduce via
  public autofit/overflow fixtures before adjusting text layout.
- [ ] Private slide 7 visible remaining problem: curves on the left render as straight horizontal lines.
  Survey the shape presets/path data structurally, then reproduce with public curve/connector fixtures before
  changing renderer logic.
- [ ] Private slide 17 visible remaining problem: left-side schema geometry has issues. Inventory the involved
  shapes, groups, connectors, and transforms with public-safe diagnostics, then isolate public geometry
  fixtures.
  - [ ] Continue residual slide-17 text parity from public PDF evidence: the small-label probe is now tightly
    bounded, but remaining page drift still includes broader text metrics, non-label typography, and the
    typed scene structure now exposed by the private-safe diagnostic.
    2026-05-26 update: page-filtered PDF inspection now shows the residual slide-17 structural gap without
    exposing content. Office and candidate both emit 44 text operations, so the next public probe should not
    chase missing text. The remaining structural differences are fractional Office font sizes versus candidate
    integer sizes, plus Office's even-odd clipping/fill operators around text regions. The explicit OOXML
    fractional-size path and text-frame `W*` clipping path now have public guards; keep the connector/group/picture
    inventory intact, and steer the next public fixture toward the remaining filled-region `f*` structure.
  - [ ] Continue the remaining public-safe slide-17 PDF-structure probe from the 2026-05-28 inspection: after
    gray operator alignment, the next residuals are fractional Office font-size variants that the candidate
    still rounds away in some inherited/autofit paths, plus ordering/segment-count differences around filled
    clipped regions. Build public fixtures for those structural cases before touching private-deck behavior.
- [ ] Private slide 15 visible remaining problem: weird mirror artifact in rendering. Inspect transforms,
  flips, and group/image drawing order, then create public transform fixtures if coverage is missing.
  - [ ] Private slide 15 visible remaining problem: left-side images and their matching text items are
    vertically misaligned. Inspect picture bounds, text-frame bounds, z-order, and any shared grouping
    assumptions, then reproduce with a public image-plus-text alignment fixture.
  - [ ] If slide-15 issues remain after text flow improves, isolate public fixtures for connector flips,
    picture flips, and grouped transform edge cases.
- [ ] Private slide 56 visible remaining problem: text is incorrectly boxed. Inspect whether the issue comes
  from shape fill/stroke, text highlight, clipping, or placeholder/text-frame bounds, then lock the generic
  behavior with public synthetic fixtures.
  - [ ] Continue slide-56 text-list parity: bold list emphasis and residual 18 pt positioning/advance
    differences still differ from the Office PDF even after run-boundary parity. Isolate those as public
    typography fixtures rather than treating the slide as resolved; likely next probes are Office baseline
    placement and per-font advance rounding for bold Calibri-like runs.
- [ ] Private-deck sweep loop: iterate over all `lokad-value-based` slides, keep a public-safe issue inventory,
  and for each visible problem add a minimal synthetic public case before implementing the generic fix.
  - [ ] Add public synthetic effect rungs for `outerShdw` and `glow` before attempting private-slide shadow
    parity. Start with no-blur/low-blur cases, then add blur/alpha/direction/distance variants only when
    the simpler Office PDF paths are understood.
- [ ] Architecture initiative: whenever a fix touches shared PPTX behavior, improve class composition and
  first-class intermediate models rather than piling more ad hoc logic into rendering code.
- [ ] Implementation-gap initiative: when an incomplete OOXML enum, preset, transform, or layout rule is
  discovered, add it to the survey/backlog and prefer filling the general gap over patching a single deck.
- [ ] PPTX typography ladder: add Office-PDF-backed visual gates for all known `a:bodyPr @vert` variants.
  Unit coverage now routes `vert`, `vert270`, `eaVert`, `mongolianVert`, `wordArtVert`, and
  `wordArtVertRtl` through first-class orientation handling, but the ladder must still lock glyph stacking,
  anchoring, clipping, and exact baseline placement before vertical labels are considered pixel-close.
  - [ ] Continue vertical text parity with Office text-operation inspection: stacked-letter orientation,
    column order, per-column x positions, and baseline placement remain visibly approximate.
- [ ] For every generic capability fixed from a private slide, add a small public synthetic test. Do not
  derive public fixtures from private slide content.
- [ ] Run `pwsh tools/CheckPrivateCase.ps1 -Case private-cases/lokad-value-based.json` after each scoped PPTX
  fix and summarize only counts, diagnostics, worst-page numbers, and private ratings.

### DOCX Feature Survey

Long-term DOCX work should mirror the successful PPTX direction without copying PPTX implementation details.
The durable target is a Word-like pipeline with clear ownership boundaries:

- Package/document model: read `word/document.xml`, styles, numbering, settings, sections, headers/footers,
  relationships, media, and theme data into typed records while keeping raw XML as source evidence until each
  OOXML surface has a typed home.
- Style cascade: resolve document defaults, table styles, paragraph styles, character styles, numbering
  levels, direct paragraph properties, and run properties before layout. The renderer should not repeatedly
  rediscover style inheritance while emitting PDF.
- Block layout and pagination: convert paragraphs, tables, drawings, headers, and footers into positioned
  page content with explicit page-size, margin, section, keep-rule, and break decisions. Pagination must be a
  first-class model output that can be traced without exposing private text.
- PDF emission: draw the positioned layout through text, table, image, shape, and field emitters that preserve
  Office-like PDF structures where observable, such as text matrices, clipping, fills/strokes, image masks,
  and reusable resources.
- Test ladder: every implemented DOCX capability should have a small public synthetic case with an Office PDF
  reference before private documents are used as acceptance evidence.

Current architecture status, 2026-05-31: `src/Lokad.OoxPdf/Docx/DocxLayout.cs` now owns a first layout stage
that turns `DocxDocument.BodyElements` into `DocxLayoutPage` records with positioned text lines, inline
images, and table rows/cells. `src/Lokad.OoxPdf/Docx/DocxRenderer.cs` still owns font subsetting, static
header/footer emission, image decoding, and PDF drawing, but it no longer decides paragraph/table page
placement while drawing. `DocxLayoutSnapshot` exposes public-safe counts and bounds for this layout without
copying text. Table cells now preserve their parsed paragraph lists in addition to the previous flattened
text, and `DocxTableCellLayout` carries layout-owned text lines that the PDF renderer consumes. This is still
an early boundary: the reader mixes style resolution into document parsing, and text lines now carry styled
segments but do not yet model font-run shaping. The next architectural work should introduce style-resolved
block models and richer section/pagination layout before adding more Word pagination behavior.

- [ ] Canonical DOCX block inventory: reduce the long-term duplication between `DocxDocument.BodyElements`
  and the separate `Paragraphs`/`Tables` inventories. A 2026-06-02 text-emission test exposed that layout can
  see manually constructed body paragraphs while font planning sees only `document.Paragraphs`; reader-built
  documents keep these lists in sync, but downstream stages should derive paragraphs/tables from the block
  stream or a single normalized inventory. This should be fixed structurally, with tests covering body,
  table-cell, header/footer, section-header/footer, and related-story paragraphs, rather than by adding another
  ad hoc traversal.
  2026-06-02 progress: introduced shared `DocxBlockTraversal` helpers and moved DOCX font planning plus
  private-safe structure style/list summaries to the body/story block stream for main-story and related-story
  paragraphs/tables. Bottom-up tests now cover body paragraphs and body tables with empty parallel
  `Paragraphs`/`Tables` inventories, closing the concrete downstream divergence found by the text-emission
  snapshot work. Validation passed `docx-core --skip-slow` (`29`). Keep the item open until the public model
  no longer needs parallel inventories for compatibility with existing tests and callers.
  2026-06-02 follow-up: related-story structure summaries now also use the shared body/story block traversal,
  and production DOCX code no longer reads `document.Paragraphs`, `document.Tables`, `story.Paragraphs`, or
  `story.Tables` as downstream rendering/diagnostic sources. The parallel inventories remain on the internal
  model for reader/test compatibility, but font planning and private-safe structure diagnostics now derive
  from the block stream consistently. Validation passed `docx-core --skip-slow` (`29`).
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
  - [x] 2026-06-01: Added DOCX `NUMPAGES` support for static header/footer fields. The reader now resolves
    `NUMPAGES` separately from `PAGE` instead of relying on a broad substring match, and static header/footer
    rendering substitutes the layout page count alongside the current page number. Validation passed
    `docx-page --skip-slow` (`18`) and full solution build. Private DOCX run `20260601-092422` stayed
    page-stable at `16/16`, zero dimension mismatches, no diagnostics, `MAE=13.388935`, changed16
    `0.124264`, worst page 9 `17.162084`. Keep body-field replacement, complex-field ranges, dynamic dates,
    document properties, and PDF link/field structure open.
  - [x] 2026-06-01: Extended `PAGE`/`NUMPAGES` placeholder substitution from static headers/footers to body
    and table-cell text emission. The layout stage still measures placeholders approximately, but PDF output
    no longer emits literal field placeholders for normal rendered lines. Validation passed `docx-page
    --skip-slow` (`18`) and `docx-tables --skip-slow` (`65` after a serial rerun of a transient compiler file
    lock), plus full solution build. Private DOCX run `20260601-092932` stayed page-stable at `16/16`, zero
    dimension mismatches, no diagnostics, `MAE=13.388935`, changed16 `0.124264`, worst page 9 `17.162084`.
    Keep layout-time field sizing and complex-field ranges open.
- [ ] Fields: `PAGE`, `NUMPAGES`, `DATE`, `REF`, `HYPERLINK`, `TOC`, `SEQ`, form fields, and cached-field
  fallback semantics.
  - [x] 2026-06-02: DOCX paragraphs now preserve a first-class field-reference inventory instead of relying
    only on placeholder text. `DocxFieldReference` records field kind, simple-vs-complex instruction source,
    optional placeholder, source run index, text-run insertion index, emitted text-run count, and emitted text
    length for `w:fldSimple` and complex `w:instrText` paths. Field opcode recognition now uses the first
    instruction token, so unsupported fields such as `PAGEREF` no longer match `PAGE` by substring.
    Private-safe structure snapshots expose aggregate field counts plus `PAGE`/`NUMPAGES`/other breakdowns at
    document, story, and paragraph-block level without exposing instructions. Validation passed the direct
    `DocxReaderPreservesFieldReferencesStructurally` test. Keep dynamic field evaluation, full complex cached
    result range ownership, layout-time field sizing, field-specific PDF structures, and non-page field
    semantics open.
- [ ] Footnotes/endnotes/comments: render bodies or emit precise diagnostics with usable fallback behavior.
  - [x] 2026-06-02: precise unsupported diagnostics now preserve related story-part ownership when the
    referenced body part exists. `DocxReader` resolves `/word/comments.xml`, `/word/footnotes.xml`, and
    `/word/endnotes.xml` through the main document relationships or content types before emitting
    `DOCX_UNSUPPORTED_COMMENTS`, `DOCX_UNSUPPORTED_FOOTNOTE`, and `DOCX_UNSUPPORTED_ENDNOTE`; if a document
    has only main-story references and no related body part, diagnostics remain scoped to `/word/document.xml`.
    This keeps the fallback private-safe and structurally aligned with the OOXML package instead of flattening
    ignored side stories into the body story. Bottom-up coverage `DocxUnsupportedStoryDiagnosticsPreferRelatedPartNames`
    builds a minimal synthetic package with all three story parts and checks their diagnostic part names.
    Validation passed `docx-core --skip-slow` (`26`). Keep this item open for actual footnote/endnote/comment
    body rendering and placement.
  - [x] 2026-06-02: comments, footnotes, and endnotes now have a first-class structural home before rendering.
    `DocxDocument.RelatedStories` stores parsed `DocxRelatedStory` records with kind, owning package part,
    story id, and paragraphs read through the existing paragraph/style/numbering/run pipeline. `DocxStructureSnapshot`
    exposes these as `Comment`, `Footnote`, and `Endnote` stories scoped to `/word/comments.xml`,
    `/word/footnotes.xml`, and `/word/endnotes.xml`, and `DocxFontPlan` includes their runs so future body
    rendering does not start with missing font resources. The existing minimal related-story package test now
    checks reader preservation, private-safe structure story metrics, and font-plan participation. Validation
    passed `docx-core --skip-slow` (`26`). Keep this item open for tables inside side stories, separator
    footnotes, actual placement/rendering, and reference marker output.
  - [x] 2026-06-02: related DOCX stories now preserve direct table blocks as well as paragraphs.
    `DocxRelatedStory` carries a side-story block stream plus paragraph and table inventories; comments,
    footnotes, and endnotes parse direct `w:p` and `w:tbl` children through the existing DOCX paragraph/table
    readers. Structure snapshots now report side-story block count, paragraph count, table count, text length,
    and inline-image count across those blocks, and font planning includes runs from side-story table cells.
    The minimal related-story test now includes a comment-owned table and checks reader/snapshot/font-plan
    preservation. Validation passed `docx-core --skip-slow` (`26`). Keep this item open for separator
    footnotes, nested/richer story content, actual placement/rendering, and reference marker output.
  - [x] 2026-06-02: main-story comment/footnote/endnote reference markers are now preserved as explicit
    inline paragraph structure instead of disappearing during text extraction. `DocxParagraph` carries
    `DocxInlineReference` entries with kind, id, and `customMarkFollows` when authored; `DocxReader` captures
    markers from normal, inserted, hyperlink, and simple-field run paths already consumed by paragraph parsing.
    Private-safe structure snapshots now expose inline-reference counts at body, block, story, table, row, and
    cell levels, with comment/footnote/endnote breakdowns for paragraph blocks. This is intentionally a
    structural alignment slice: renderer glyph choice, numbering display, separator footnotes, and body
    placement remain open until Office PDF inspection supplies the observable rules. Validation passed
    `docx-core --skip-slow` (`26`).
  - [x] 2026-06-02: inline reference markers now carry private-safe source anchors instead of only aggregate
    counts. `DocxInlineReference` records the source run index, child index, and visible in-run text offset
    before the marker, using the same run-child text contribution logic that extracts tabs, soft/no-break
    hyphens, carriage returns, and soft line breaks. `DocxStructureSnapshot` exposes anchored-reference counts
    and maximum in-run marker offset at document and paragraph-block level. Bottom-up coverage includes a mixed
    text/footnote-reference/text run so future marker rendering can consume model-owned insertion points
    instead of guessing placement from flattened text. Validation passed `docx-core --skip-slow` (`32`) and
    `docx-text --skip-slow` (`45`; a first parallel run hit a transient compiler output lock and passed on
    serial rerun). Keep renderer glyph choice, automatic footnote/endnote numbering display, separator
    footnotes, side-story placement, and PDF link/annotation structure open.
  - [x] 2026-06-02: bottom-up coverage now locks inline reference anchors through the same inline containers
    used by other DOCX run content. `DocxReaderPreservesInlineReferencesInsideRunContainers` checks footnote
    markers in final-view inserted runs, comment markers inside hyperlinks, and endnote markers inside simple
    fields, including source run index, run child index, and visible text offsets. This intentionally stops at
    structural preservation: Office-derived marker glyphs, numbering/restart rules, footnote separators,
    side-story placement, and annotation/link PDF structures remain open until those rules are modeled from
    OOXML settings and observed Office output instead of guessed. Validation passed the direct test.
- [ ] Tracked changes: choose final, original, or marked-up view explicitly and document the behavior.
  - [x] 2026-06-01: Added the first final-view tracked-change slice for simple paragraph run wrappers.
    `w:ins/w:r` now flows through the normal run parser and style cascade, while `w:del/w:r` remains absent
    from rendered paragraph text. Validation passed `docx-text --skip-slow` (`28`) and `docx-core
    --skip-slow` (`19` after a serial rerun of a transient compiler output lock), plus full solution build.
    Private DOCX run `20260601-091958` stayed page-stable at `16/16`, zero dimension mismatches, no
    diagnostics, `MAE=13.388935`, changed16 `0.124264`, worst page 9 `17.162084`. Keep the parent open for a
    first-class tracked-change mode covering inserted/deleted paragraphs, table rows/cells, move ranges,
    comments, metadata, and whether diagnostics should remain broad once final-view coverage is complete.
- [ ] Multi-column layout, text boxes, sidebars, bookmarks, hyperlinks, outlines, and document properties.
  - [x] 2026-06-01: Preserved visible DOCX paragraph runs wrapped in `w:hyperlink` by routing hyperlink child
    runs through the normal run parser and style cascade. Public coverage checks document-order text and the
    inherited Hyperlink character style; validation passed `docx-text --skip-slow` (`29`) and full solution
    build. Private DOCX run `20260601-092152` stayed page-stable at `16/16`, zero dimension mismatches, no
    diagnostics, `MAE=13.388935`, changed16 `0.124264`, worst page 9 `17.162084`. Keep actual PDF link
    annotations, internal anchors, visited/unvisited style behavior, and relationship-target modeling open
    under this parent.
  - [x] 2026-06-02: DOCX hyperlink wrappers now have a paragraph-owned structural span instead of being only
    styled text. `DocxHyperlinkSpan` carries `r:id`, anchor, tooltip, history token, authored target,
    target mode, resolved internal target when applicable, source-run range, text-run range, and text length.
    The main document reader now keeps an all-relationships view for hyperlink targets while retaining the
    internal-only relationship view for package parts. Structure snapshots expose hyperlink counts split into
    external and internal link inventories; table, row, and cell snapshots now carry the same hyperlink
    inventory for table-cell paragraphs. Validation passed `docx-text --skip-slow` (`45`),
    `docx-core --skip-slow` (`32`), `docx-tables --skip-slow` (`94`), and full solution build.
  - [x] 2026-06-02: side-story hyperlinks now keep relationship targets as well. Comment/footnote/endnote
    paragraph parsing receives the full story-part relationship map, so external hyperlinks inside related
    stories can populate `DocxHyperlinkSpan` just like main-document hyperlinks. `DocxStructureStorySnapshot`
    now reports hyperlink, external-link, and internal-link counts for body, header/footer, and related
    stories. The related-story package test includes a comment-owned external hyperlink and checks reader,
    structure, and font-plan ownership. Validation passed `docx-core --skip-slow` (`32`),
    `docx-text --skip-slow` (`45`), and full solution build. Keep actual PDF annotations, internal anchor
    destinations, visited/unvisited style behavior, and visual placement open.

### DOCX Synthetic Fidelity Ladder

Build a DOCX ladder comparable to the PPTX ladder. Each rung must be public, synthetic, minimal,
Office-PDF-inspected, visually gated when close, and free of private content.

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
- [x] 2026-05-31: Extended the internal DOCX layout snapshot with public-safe per-page vertical consumption
  and block-kind height sums for text lines, inline images, and table rows. This does not expose document text
  and gives private pagination investigations a stable place to compare block-level drift.
- [x] 2026-06-01: Extended `tools/Lokad.OoxPdf.DocxInspect` with public-safe `block-sequence.json` and
  `table-adjacency-summary.json` outputs. The new summaries expose body block kinds, paragraph spacing/style
  facts, table ordinals, row/column counts, and table-neighbor classifications without private text. Private
  DOCX inspection showed the current acceptance document has `198` paragraph blocks and `13` tables, and all
  tables are paragraph-to-table-to-paragraph transitions. The refreshed private run `20260601-123301` stayed at
  `16/16` pages with zero diagnostics, `MAE=13.791190`, changed16 `0.126277`; worst pages are not exclusively
  table-heavy. Do not add a blanket replacement post-table gap from this evidence: the remaining private drift
  also needs text-line/style/numbering investigation.
- [ ] Implement style-derived paragraph spacing accurately, including exact Office autospacing magnitudes,
  Word-like adjacent paragraph spacing collapse around tables/sections, and diagnostics that distinguish
  supported tokens from unresolved spacing semantics.
- [ ] Implement paragraph and numbering indents: left/right/first-line/hanging indents from paragraph styles
  and numbering levels, with corresponding wrapping-width changes.
  - [x] 2026-06-01: Added typed `DocxParagraphIndent` for ordinary paragraph `w:ind` values and cascaded it
    through defaults/styles/direct paragraph properties. Layout now applies left/right/first-line/hanging
    indents to non-numbered body and table-cell paragraph wrapping, while keeping numbering-specific label/tab
    geometry on `DocxNumberingIndent`. Public unit coverage checks style/direct first-line-vs-hanging cascade
    and wrapping x/width changes. Public visual `docx-ladder-02-paragraph-indents` run `20260601-022308` is
    dimension-stable at `MAE=1.104564`, changed16 `0.011547`. Private DOCX run `20260601-022152` improved to
    `16/16` pages, zero dimension mismatches, no diagnostics, `MAE=12.481106`, changed16 `0.116738`. Keep the
    parent open for `w:start`/`w:end`, char-unit indents, mirror indents, and exact interaction with numbering
    styles.
  - [x] 2026-06-01: Resolved logical paragraph and numbering indents for the current left-to-right layout path:
    `w:start`/`w:end` now feed the effective start/end twip values before falling back to physical
    `w:left`/`w:right`, and character-unit indent attributes now emit a precise
    `DOCX_UNSUPPORTED_CHARACTER_UNIT_INDENT` diagnostic instead of disappearing silently. Public parser
    coverage checks `start`/`end` precedence; public visual `docx-ladder-02-paragraph-logical-indents` run
    `20260601-022902` is dimension-stable at `MAE=0.094507`, changed16 `0.002454`, SSIM `0.986713`. Public
    validation passed `docx-text --skip-slow` (`18`), `docx-core --skip-slow` (`16`), `docx-numbering
    --skip-slow` (`11`), `docx-page --skip-slow` (`17`), `docx-tables --skip-slow` (`58`), and `dotnet build
    Lokad.OoxPdf.slnx --tl:off --nologo -v minimal`. Private DOCX run `20260601-022931` stayed stable at
    `16/16` pages, zero dimension mismatches, no diagnostics, `MAE=12.481106`, changed16 `0.116738`. Keep the
    parent open for character-unit conversion from Word's document grid/font metrics, mirror/Bidi indents, and
    exact numbered first-line/right-indent interactions.
- [ ] 2026-05-31: Resolve DOCX `w:basedOn` style inheritance with pagination-safe Office fixtures before
  enabling it broadly. The private style graph uses chains such as body/table styles based on Normal, but a
  naive recursive paragraph/character merge changed the private candidate from 16 pages to 14 pages in run
  `20260531-161731`; this must be tackled with fixture coverage for inherited spacing, font size, and page
  breaks rather than a blind cascade merge.
- [x] 2026-06-01: Promote DOCX run character spacing (`w:rPr/w:spacing`) through the typed run model,
  resolved style cascade, text measurement, wrapping/segment placement, static header/footer placement, and
  PDF glyph emission. The parser treats the token as signed twentieths of a point, not as a font-family
  special case; measured widths and emitted `TJ` positioning now agree for authored positive/negative
  tracking. Public coverage checks reader parsing, layout segment advances, and PDF positioning emission.
- [x] 2026-06-01: Add an Office-authored public DOCX character-spacing ladder before making broader
  PDF-emission claims. `docx-ladder-02-character-spacing` includes positive and negative `w:spacing` in body
  paragraphs, mixed-run boundaries, table cells, and header/footer ranges. Public run `20260601-030351` is
  page- and dimension-stable (`MAE=1.960983`, changed16 `0.016165`). Office PDF inspection shows the authored
  `+2pt`/`-1pt` spacing as positioned text adjustments near those values, while also using `TJ` text
  operations and small residual adjustments for nominally zero-spacing runs.
- [x] 2026-06-01: Align the first DOCX PDF text-emission layer with Office's observed positioned text
  strategy. The renderer now emits DOCX glyph runs as positioned `TJ` arrays even when authored
  `w:rPr/w:spacing` is zero, so kerning and structural tracking share one PDF operation form instead of
  falling back to plain `Tj` for nominally untracked runs. Public `docx-ladder-02-character-spacing` run
  `20260601-031105` stayed page- and dimension-stable (`MAE=1.960958`, changed16 `0.016143`), and PDF
  inspection now shows candidate text operations as `15` `TJ` operations versus Office's `28` `TJ`
  operations. Private DOCX run `20260601-031243` stayed accepted at `16/16` pages with zero dimension
  mismatches and no diagnostics, but the aggregate moved to `MAE=13.047852`, changed16 `0.120850`; treat
  this as a structural PDF-alignment step, not a completed visual improvement.
- [ ] Continue DOCX text decomposition alignment against Office's observed character-spacing output. The
  public ladder still shows Office emitting more text operations (`28` vs candidate `15`) and small residual
  positioned adjustments on some nominally zero-spacing runs. Investigate run splitting, script/field/static
  part boundaries, kerning decomposition, and line-segment ownership with Office-backed public fixtures before
  changing spacing math. Do not introduce font-name rules or private-document exceptions.
  2026-06-01 progress: Office inspection of `docx-ladder-02-character-spacing` showed a concrete run-boundary
  decomposition gap: for an authored run whose preserved text starts with a space, Office emits the leading
  space as its own `TJ` text operation before the following word, while the candidate emitted `" text"` as one
  operation. `CreateTextSegments` now splits leading regular spaces into separate layout segments and advances
  the following segment through the existing run character-spacing boundary rule, preserving glyph positions
  without font, content, or coordinate special cases. Public `docx-text --skip-slow` passed (`36` passed) and
  `docx-core --skip-slow` passed (`22` passed; one parallel rerun first hit a transient compiler file lock).
  Public `docx-ladder-02-character-spacing` run `20260601-122636` stayed raster-neutral against the prior run
  (`MAE=0.736009`, changed16 `0.007625`, SSIM `0.787077`) while PDF inspection now shows the mixed line as
  separate candidate operations `" "`, then `"text"`, matching Office's structural split for that boundary.
  The item remains open for static header/footer tab decomposition, empty paragraph mark/resource differences,
  and Office's small nominally-zero spacing residuals.
- [ ] Improve numbering layout: render labels in their own hanging-indent area, support level text expansion
  beyond the current simple label prefix, and honor restart/start rules.
- [ ] Improve table layout accumulation: preferred table widths, cell widths, row minimum height from
  content, cell vertical alignment, cell margins, and repeating header rows.
  - [ ] 2026-05-31: Replace the remaining DOCX table-cell first-baseline heuristic only with Office-backed
    evidence. A structural trial that replaced the fixed first-baseline inset with the resolved font's
    OpenType ascender passed public table tests but regressed the private DOCX aggregate from `14.818900` to
    `15.022987` MAE in run `20260531-222543`. Private-safe inspection found explicit table-cell top/bottom
    margins of `0`, so the missing piece is likely Word's table-cell line-box/baseline placement rather than
    margin parsing. Add a small Office-authored public fixture comparing first-baseline placement for
    font-size/table-style/default-margin variants before changing production behavior.
    2026-06-02 follow-up: the current body/table baseline constants now live behind `DocxLineMetrics`
    (`ResolveBodyBaselineOffset`, `ResolveTableCellFirstBaselineInset`) instead of being repeated at layout
    and test call sites. This is a no-behavior-change architectural step so future Office-derived calibration
    has one structural home. Validation passed `docx-tables --skip-slow` (`87`) and `docx-core --skip-slow`
    (`25`). Keep this item open for the actual Word-backed first-baseline fixture and metric replacement.
- [ ] Revisit keep rules only after layout tracing exists: support style-derived `keepNext`, `keepLines`, and
  widow/orphan control with synthetic tests and private page-count checks.
  2026-06-02 follow-up: keep-block preflight now returns a typed `DocxKeepBlockEstimate` with measured
  height plus paragraph and first-table-row target counts instead of a bare height. This preserves the
  current pagination behavior while making the structure being kept explicit for future table-fragment and
  widow/orphan work. Added bottom-up coverage for a `keepNext` paragraph that must move with the first row of
  the following table; validation passed `docx-tables --skip-slow` (`88`). Keep this item open for true
  line-level orphan/widow behavior and richer keep chains across fragmented tables.
- [ ] Reattempt manual page/column break support with a parser change that does not alter paragraphs when no
  matching break exists; previous paragraph-splitting attempts changed the private page count and were
  reverted.
  - [x] 2026-06-02: Added the safe structural slice for run-only DOCX column-break paragraphs:
    `w:br w:type="column"` now becomes a typed `DocxManualBreakElement` in the body stream, structure
    snapshots count and describe it, and layout treats it as an unsupported single-column boundary instead of
    forcing a false page break. `DOCX_UNSUPPORTED_MANUAL_BREAK` remains the diagnostic for column breaks until
    true multi-column flow exists. Keep this item open for inline text-bearing column breaks and actual
    multi-column pagination.
- [ ] For every generic capability fixed from a private page, add a small public synthetic test. Do not
  derive public fixtures from private page content.
- [ ] After each scoped fix, run `pwsh tools/CheckPrivateCase.ps1 -Case
  private-cases/user-requirements-spec.json` and record only page counts, aggregate metrics, diagnostics, and
  worst-page numbers.

### DOCX Table Recovery Plan

- [x] 2026-06-01: Add the first DOCX table trace layer with public-safe row/cell metrics on each layout page.
  `DocxLayoutSnapshot` now includes table-row and table-cell snapshots with page geometry, header-row tokens,
  preferred width tokens, grid spans, vertical alignment, margins, border counts, fill/shading presence,
  conditional-format presence, inline-image counts, text-line counts, and text length only. It intentionally
  does not expose cell text or fill/shading color values. Public validation passed `docx-core --skip-slow`
  (`16`) and `docx-tables --skip-slow` (`49` after a serial rerun because the first parallel attempt hit the
  known compiler output lock). Private DOCX run `20260601-003212` stayed neutral at `16/16` pages, zero
  dimension mismatches, no diagnostics, `MAE=12.509698`, changed16 `0.112673`.
- [ ] Promote the DOCX table trace from layout rows/cells to full table ownership: stable table ordinal,
  page index range, row count, column count, `tblGrid`, preferred table width, resolved table width, row
  height declarations, vertical merges, and resolved column grid. This remains open because the current trace
  is enough to inspect private table bands but does not yet expose table-level width/grid ownership.
  - [x] 2026-06-02: Extended `DocxLayoutSnapshot` table ownership with authored-vs-inferred grid provenance,
    authored header-row count, aggregate row-height declaration counts (`exact`/`atLeast`), and cant-split row
    count while preserving existing table ordinal, source block, page range, resolved grid, preferred width,
    indent, spacing, layout, and vertical-merge facts. Keep this item open for actual fragmented-row and
    fragmented-cell geometry, where repeated header rows and split rows need richer ownership than aggregate
    table counts. Validation passed `docx-tables --skip-slow` (`92`) and full solution build.
  - [x] 2026-06-02: Added table-level row-fragment ownership to `DocxTableSnapshot`: authored fragmented row
    count, laid-out row-fragment count, and max fragments per row. This keeps the existing per-row fragment
    geometry but makes split-row/table pagination drift visible at table scope for private-safe inspection.
    Public coverage checks both the non-fragmented baseline and a tall row split across two pages. Validation
    passed `docx-tables --skip-slow` (`92`).
  - [x] 2026-06-01: Added stable table ordinals and document-level table snapshots derived from layout rows.
    The trace now reports page index range, source row count, laid-out row count, repeated header-row layout
    count, grid column count, grid-column width sum, resolved table width, table X, preferred table width
    tokens, indent, cell spacing, layout token, and vertical-merge presence. While doing this, found and
    fixed a structural parser gap: `w:vMerge` was not preserved at all. `DocxTableCell` now carries generic
    vertical-merge presence/value tokens, with public coverage for both `w:val="restart"` and val-less
    continuation. Validation passed `docx-core --skip-slow` (`16`), `docx-tables --skip-slow` (`50`), and
    full solution build. Private DOCX run `20260601-003853` stayed neutral at `16/16` pages, zero dimension
    mismatches, no diagnostics, `MAE=12.509698`, changed16 `0.112673`.
  - [x] 2026-06-01: Extended the private-safe table-cell trace with first text-line x, first baseline y, and
    last baseline y. The row-height ladder showed that table/cell y drift cannot be diagnosed from row bounds
    alone: candidate cell text now has near-Office x after default horizontal padding, but baseline y remains
    a row/table placement problem. The new fields expose only geometry and text-line counts, not cell text.
    Validation passed `docx-core --skip-slow` (`19`) and `docx-tables --skip-slow` (`62`).
  - [ ] Finish table-level ownership by exposing row height rule/value tokens and the resolved column grid
    itself, then use the trace to select the simple/dense/worst private tables below. Vertical merges now
    affect same-page layout/rendering, but cross-page merged-cell behavior and resolved-grid tracing remain
    open.
  - [x] 2026-06-01: Preserved DOCX table-row height value/rule tokens and exposed them through the
    private-safe layout snapshot. `DocxTableRow` now carries `w:trHeight/@w:val` and `@w:hRule`, and layout
    distinguishes `exact` from at-least/default row heights instead of always expanding exact rows to
    content. Public coverage verifies parser token preservation, exact-height layout behavior, and snapshot
    fields. Validation passed `docx-tables --skip-slow` (`52`), `docx-core --skip-slow` (`16`), and full
    solution build. Private DOCX run `20260601-004602` stayed neutral at `16/16` pages, zero dimension
    mismatches, no diagnostics, `MAE=12.509698`, changed16 `0.112673`.
  - [x] 2026-06-01: Exposed the resolved DOCX table column grid in private-safe table snapshots. The trace
    now carries per-column resolved widths alongside authored grid count/sum and resolved table width, so
    private tables can be compared structurally without inferring columns from cell rectangles. Validation
    passed `docx-core --skip-slow` (`16`), `docx-tables --skip-slow` (`52`), and full solution build.
    Private DOCX run `20260601-004815` stayed neutral at `16/16` pages, zero dimension mismatches, no
    diagnostics, `MAE=12.509698`, changed16 `0.112673`.
- [ ] Select representative private tables for repeated inspection: one simple table, one typical dense
  table, and one worst table. Record only table ordinal/page, private rating, and public-safe feature gaps.
  - [x] 2026-06-01: Selected public-safe representative private DOCX tables from the layout snapshot
    inventory. Use zero-based table/page indexes in code traces: simple table `1` on page index `2`
    (`5` rows, `2` columns, `10` cell layouts, `12` text lines); typical dense table `7` on page index `11`
    (`13` rows, `3` columns, `39` cell layouts, `42` text lines); worst/densest table `11` on page index
    `14` (`19` rows, `6` columns, `114` cell layouts, `137` text lines). All three have no grid-span or
    vertical-merge cells in the current trace; the recurring high private error is therefore more likely
    table width/text wrapping, borders, fills, or pagination than merge handling. Keep notes private-safe:
    do not record table text or screenshots.
- [ ] Fix the table layout model before cosmetic styling: resolve `tblGrid`, `tblW`, `tcW`, page content
  width, percentage/auto widths, and grid scaling consistently.
  - [x] 2026-06-01: Distinguished authored `tblGrid` from parser fallback columns. Tables without an
    explicit grid now distribute inferred columns across the resolved table width instead of treating 72pt
    placeholder columns as authored geometry; explicit grids and preferred cell widths keep their existing
    behavior. Public coverage verifies the no-`tblGrid` layout case. Validation passed `docx-tables
    --skip-slow` (`53`), `docx-core --skip-slow` (`16`), and full solution build. Private DOCX run
    `20260601-005313` stayed neutral at `16/16` pages, zero dimension mismatches, no diagnostics,
    `MAE=12.509698`, changed16 `0.112673`, indicating the selected private tables are explicit-grid cases.
  - [x] 2026-06-01: Extended missing-`tblGrid` inference through `gridSpan` rows. The parser now synthesizes
    fallback grid columns from the maximum logical grid width, and layout has a defensive empty-grid inference
    path. Public `docx-ladder-03-table-missing-grid-spans` isolates a table whose rows have only two physical
    cells but three logical grid columns; run `20260601-145200` is page-stable with no diagnostics
    (`MAE=0.615754`, changed16 `0.009697`) and the follow-up family run `20260601-145433` kept the DOCX
    family passing at `25/25` cases. The remaining visible residual in that case is table text
    baseline/vertical spacing, not column ownership.
- [ ] Compute row heights from actual cell content: wrap text within cell width, include cell margins,
  respect explicit `trHeight` rules, and avoid the current fixed/default row-height behavior for
  content-heavy rows.
  - [x] 2026-06-01: Added explicit `w:hRule="exact"` handling so exact row heights no longer behave like
    `atLeast`; at-least/default rows still expand to measured content. Keep this parent open for cross-page
    row splitting and deeper Office parity around `auto` and pagination.
- [ ] Render cell text as paragraphs instead of flattened cell text: preserve paragraph breaks, basic run
  styling, numbering/bullets inside cells, alignment, and line spacing.
  - [x] 2026-06-01: Removed the stale fallback-resource gate for table-cell text and clipped table-cell
    text/images to the full cell rectangle. Public coverage now includes a fallback-free explicit run font
    table fixture.
  - [ ] 2026-06-01: Replace the remaining authored-margin first-baseline compatibility branch with an
    Office-backed rule. The no-margin public fixture shows first baseline placement follows the first run
    font size rather than the old fixed inset, but applying that rule to authored-margin cells regressed
    private table pages. Add a public Word-authored ladder for explicit `tcMar`/`tblCellMar` top margins,
    absent margins, mixed font sizes, and vertical alignment before changing the production baseline again.
    Follow-up: added public `docx-ladder-02-table-cell-margins`. Office PDF inspection confirms absent-margin
    cells align near the cell edge (`X=72.504/252.53`, `Y=707.74`) while explicit-margin cells honor authored
    left margins (`X=84.024/270.05`) but share a row baseline (`Y=658.66`) despite different top margins. The
    candidate matches the horizontal margins (`X=84/270`) but places the explicit-margin baselines at
    `Y=661/655`, proving the remaining vertical gap is row height/paragraph spacing/vertical baseline
    coupling, not a single replacement constant.
    Follow-up: explicit `atLeast` row heights now expand by the row's maximum authored top cell margin before
    row layout, matching the Office evidence that `w:trHeight` is not the whole row box once top `tcMar` is
    present. This improved the public margin ladder without changing the private DOCX aggregate metrics.
    A tempting cleanup to use the measured row-height sum for the table pre-pagination check changed the private
    candidate to 17 pages, so keep table pre-pagination as an open pagination model task rather than folding it
    into this row-height fix.
    Follow-up: added public `docx-ladder-03-table-pagination-margins` to isolate that pagination edge. Office
    and the candidate both split the two-row table across two pages instead of moving the whole table before
    the first row, confirming the rejected measured-table-height precheck was directionally wrong. Public run
    `20260601-013533` is page- and dimension-stable with no diagnostics; page 1 MAE is `0.927605` and page 2
    MAE is `0.253998`. Text inspection shows the first table row stays on page 1 and the second row starts at
    the top of page 2 in both PDFs.
- [ ] Implement table and cell styling: table styles, conditional first/header row formatting, cell shading,
  per-edge borders, border widths/colors, and vertical alignment.
  - [x] 2026-06-01: Reused the DOCX shading resolver for table-cell backgrounds instead of treating
    `w:tcPr/w:shd @fill` as a literal-only fill. Table cells now honor `clear`/missing values as solid
    fills and resolve percentage shading such as `pct20` by blending foreground `w:color` over background
    `w:fill`, matching the Office evidence already observed for run shading. Added unit coverage and public
    `docx-ladder-03-table-cell-percentage-shading` to lock literal fills, implicit clear, no-fill cells,
    and multiple percentage blends. Public run `20260601-152522` passed (`MAE=1.081647`, changed16
    `0.017205`, foreground histogram `0.999367`); full `docx-layout` passed `28/28`, `docx-tables
    --skip-slow` passed `77`, and full solution build passed. Private DOCX run `20260601-152430` stayed
    unchanged at `16/16` pages, zero dimension mismatches, no diagnostics, `MAE=13.855991`, changed16
    `0.127419`, confirming this was a structural table-styling gap rather than the current private-driver
    mismatch. Keep non-percent pattern families, `auto` foreground/background resolution, and table-style
    conditional shading precedence open under this parent.
- [ ] Implement structural table features: horizontal merges (`gridSpan`), vertical merges (`vMerge`),
  repeating header rows across page breaks, and page-break behavior inside rows.
  - [x] 2026-06-01: Consumed preserved `w:vMerge` tokens in DOCX table layout/rendering for same-page
    merged cells. Restart cells now span continuation row heights, continuation cells remain trace-visible
    but skip fill/border/text/image rendering, and a public layout fixture covers the geometry. This is
    intentionally structural and not font- or document-specific. Validation passed `docx-tables --skip-slow`
    (`51`) and `docx-core --skip-slow` (`16`). Private DOCX run `20260601-004244` stayed neutral at `16/16`
    pages, zero dimension mismatches, no diagnostics, `MAE=12.509698`, changed16 `0.112673`. Full solution
    build passed with transient copy-retry warnings while the private-case CLI process still held its DLL.
    Keep cross-page merged cells open under this parent.
- [ ] Add synthetic public tests for each table capability before using the private document as evidence;
  never derive fixtures from private table content.

### PDF/Infrastructure

- [ ] Audit current PDF generation patterns against Office reference PDFs: text grouping, text matrices,
  clipping regions, image masks, transparency state, path construction, stroke/fill order, resource
  naming/reuse, and page content stream organization.
- [ ] Improve PPTX text-line emission toward Office-like text objects: PPTX now emits positioned `TJ` arrays
  even when a line has no explicit kerning/tracking adjustment, matching Office's common text-object pattern
  without changing raster output. Continue by reducing unnecessary run splitting around spaces when a line
  can be emitted as one positioned text object.
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
## Next Implementation Targets

1. Make Office-PDF structure the primary fidelity oracle. Pixel metrics stay useful, but every serious fix
   should first ask what Office emitted: text matrices, glyph advances, clipping, image XObjects, paths,
   transparency groups, resources, and drawing order. Extend `PdfInspect`/comparison tooling when a mismatch
   cannot be explained structurally; for charts, first try to reuse and widen the existing
   `ClassifyPdfChartGraphics.ps1` and `ClassifyPdfChartText.ps1` gates before creating a new oracle surface.
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

- Observation: Private slide 80's missing grey world-map layer was not a raster fallback issue or a custom
  geometry default-fill rule.
  Evidence: The slide has `792` custom geometries; `785` map fragments use `a:grpFill` with one custom path,
  white `bg1` strokes, and ancestor group `solidFill` of `bg1` with `lumMod=85000`. Office emits `795` fills
  on page 80, including `785` `g:0.851` fills for those fragments; the previous candidate emitted only `10`
  fills and stroked the fragments.
- Observation: Private slide 5's thin graph arrows were a straight-line `stealth` line-end sizing problem,
  not a chart-series stroke-width problem.
  Evidence: The slide uses 1 pt `straightConnector1` connectors with `tailEnd type="stealth"`. PDF inspection
  showed the candidate's horizontal green arrowheads at 3 pt high while Office emitted the same bounds and
  color at 6 pt high. The public `pptx-ladder-06-straight-stealth-connectors` fixture reproduced the same
  Office behavior for 1 pt `line` and `straightConnector1` presets: Office applies a 6 pt minimum stealth
  marker length/width for straight lines, while 2 pt lines already match the existing factor-derived 6 pt
  marker.
- Observation: The dominant private page-36 three-column text-state branch is not a baseline-placement
  problem.
  Evidence: The reliable-position summary for run `20260531-014545` has `41` matched operations in the
  explicit `noAutofit`, three-column, `12pt` frame with average `DeltaBaselineY=+0.016632pt` and range
  `-0.020441..+0.055559pt`. The remaining page-36 rendering-impact branch is therefore PDF text-state and
  operation decomposition, not a global baseline constant.
- Observation: Private page-79 table text has repeatable sub-point baseline residuals by row/cell, but not a
  single row/column offset rule.
  Evidence: The same reliable-position summary reports page-79 table rows around `-0.554pt`, `-0.319pt`,
  `-0.328pt`, `-0.233pt`, `-0.200pt`, and `-0.199pt`, with stable per-cell buckets. The table uses explicit
  small insets and no declared/rendered row-height slack, so the default table-inset constants are not the
  active branch.
- Observation: A public no-slack, small-inset, middle-anchored table with wrapped multi-run cells is a zero-`Tc`
  counterexample.
  Evidence: `pptx-ladder-10-table-middle-small-insets` run `20260531-023220` passes with MAE `0.966760`.
  Its Office and candidate text-operation counts both equal `79`; all Office text operations have `Tc=0`, and
  matched baseline deltas stay within about `+0.01..+0.28pt`.
- Observation: A structurally similar public table can reproduce the private page-79 nonzero table-`Tc` branch.
  Evidence: `pptx-ladder-10-table-font-fragmentation` run `20260531-024147` passes as a discovery fixture with
  MAE `2.029893`, changed16 `0.032217`, and empty diagnostics. PDF/text inspection reports `125` Office text
  operations, `119` matched operations, `6` missing operations, `116` reliable-position matches, and `90`
  matched nonzero Office `Tc` operations. The matching zero-`Tc` counterexample has the same no-slack,
  explicit-small-inset, middle-anchored table geometry, so the discriminator must be in structural
  font/run/line decomposition rather than table geometry alone.
- Observation: Changing slide height while holding 21 pt textbox source coordinates fixed shifts Office's
  secondary `/Tf +0.024 pt` window.
  Evidence: Public-safe ignored variants of `font-size-quantization-y-scan-21pt-fine` report secondary rows at
  source Y `135..153` on a 540 pt slide, `165..183` on a 440 pt slide, and `150..171` on a 640 pt slide. The rule is
  therefore not a fixed source-coordinate, fixed top-origin baseline, or local text-frame condition.
- Observation: The secondary `/Tf +0.024 pt` page-height dependence is not monotonic.
  Evidence: Additional 21 pt fine Y-scan variants show `490 pt` slide height branching at source Y `120..129`,
  `590 pt` at `120..138`, and `690 pt` at `183..198`. Top-origin baseline grid remainders differ by variant, while
  most secondary rows still share bottom-origin remainder `0.833333`; neither value alone separates the branch.
- Observation: The public wrapped-text secondary `/Tf +0.024 pt` branch is not explained by body-property wrap,
  vertical overflow, vertical-overflow source, or autofit mode.
  Evidence: The refreshed `font-size-quantization-wrap13b` branch summary has `32` main-grid rows and `6`
  secondary rows, and every matched candidate row in both branches reports `Wrap=Square`,
  `VerticalOverflow=Overflow`, `VerticalOverflowSource=DefaultValue`, and `Autofit=noAutofit`.
- Observation: Public COM-authored middle-anchor probes split visible and empty paragraph behavior.
  Evidence: A two-line Arial 18 pt text box with no trailing empty paragraph matches Office within about
  `0.05 pt` when the estimator uses OS/2 Windows ascender plus descender. The same probe with a trailing
  empty paragraph overshoots if the empty line uses the full font box and underfits if it uses normal
  paragraph advance, leaving a small residual midpoint behavior that still needs a broader public probe
  family.
- Observation: The private slide-17 issue has moved from vertical frame placement to right-edge glyph
  segmentation/advance.
  Evidence: The affected private frame now emits the two visible baselines within about `0.35 pt` of Office,
  while the final same-line segment remains horizontally offset and carries a different PDF text-operation
  length. That means more vertical-anchor tuning would be the wrong long-term move.
- Observation: Shape text frames and table-cell text frames currently need distinct default-spacing font-box
  metrics for middle anchoring.
  Evidence: Applying the new OS/2 Windows font-box rule globally broke public table-cell anchoring tests.
  Restoring the existing typographic font-box path for table-cell anchors kept `pptx-tables` green while the
  new shape-text font-box regression remained green. This is an explicit surface distinction until public
  Office table probes justify a common rule.
- Observation: Table-cell text-frame model provenance was weaker than the scene provenance after the inset
  and anchor ownership slices.
  Evidence: `PptxSceneTableCell` distinguished explicit `a:tcPr @anchor` from absent anchors and preserved
  inset sources, but `PptxTextFrameModelSnapshot` had no table-cell inspection path and the table adapter
  still mapped an absent anchor to `TableCellStyle`. The new table text-frame inspection test exposes that
  boundary and locks absent anchors as `DefaultValue`.
- Observation: The earlier end-paragraph ownership slice still left layout parsing resolved font metrics from
  `EndParagraphProperties`.
  Evidence: `PptxTextParagraphModel` exposed the raw `a:endParaRPr` owner, but both empty-paragraph layout and
  vertical-anchor height estimation called `ReadFontSize`, `ReadTypeface`, and boolean parsers against that
  XML. The model now exposes the resolved end-paragraph style, and the end-paragraph vertical-anchor fixture
  asserts the resolved `72 pt` font size before layout runs.
- Observation: PPTX tab stops were only partly model-owned.
  Evidence: `ResolvedParagraphTextStyle` already carried `TabStops` and line layout consumed that model field,
  but `ReadTabStops` only inspected direct `a:pPr/a:tabLst`; inherited `a:lvlNpPr/a:tabLst` from the resolved
  list-style/default paragraph cascade was ignored. The model now reads the direct tab list before falling back
  to resolved defaults, and inspection exposes the inherited tab-stop positions.
- Observation: PPTX bullet marker text was still a layout-time XML decision.
  Evidence: The paragraph model had the resolved paragraph style cascade, but `BuildTextFrameLayout` called
  `ReadBulletText` against merged paragraph XML to decide `buChar`, symbol mapping, `buAutoNum`, and `buNone`.
  The paragraph model now carries the marker facts directly; the remaining XML dependency is bullet visual
  style (`buFont`, `buClr`, `buSz*`) while text marker selection has moved to typed state.
- Observation: PPTX bullet visual style was the remaining bullet-specific layout XML dependency.
  Evidence: After bullet marker ownership moved into `PptxParagraphBulletModel`, `ReadBulletStyle` still read
  `buFont`, `buClr`, `buSzPct`, and `buSzPts` from merged paragraph XML during line layout. The model now
  carries the resolved color and raw size tokens, and layout only combines them with the first visible run's
  fallback text style.
- Observation: Empty PPTX paragraphs still had a paragraph-spacing XML back edge.
  Evidence: The text model owned `EndParagraphStyle`, but empty paragraph layout and vertical-anchor height
  estimation still called `ReadParagraphSpacing(paragraph.Properties, paragraph.DefaultProperties, ...)` using
  the end-paragraph font size. The model now resolves those empty spacing values once, alongside the
  end-paragraph style they depend on.
- Observation: The final empty-paragraph layout XML probe was a boolean, not a numeric formula.
  Evidence: After empty spacing moved into the model, layout still called a helper that checked
  `paragraph.Properties is not null || paragraph.EndParagraphProperties is not null` to decide whether an empty
  paragraph participates in height. `HasLayoutContent` now records that decision during model construction.
- Observation: Small positive table row slack has at least two Office behaviors that look similar if inspected
  only as frame-height surplus.
  Evidence: Private page 12 has about `1.0264x` positive table slack and one filled row; Office paints that
  row at its declared height, and the accepted candidate now matches the three filled rectangles. Private page
  78 has about `1.0198x` positive slack but a dense filled table; the same declared-row rule regressed that page
  and had to be rejected. Future table work must classify sparse decoration rows separately from dense filled
  grids before changing row allocation.
- Observation: Office exports `triangle` line ends differently for `line` and `straightConnector1` presets.
  Evidence: The existing public vertical `line` connector test still matches Office with the `1.5x` triangle
  half-width branch. Private page 20's `straightConnector1` triangles, with the same `triangle` line-end token
  and `1.5 pt` stroke width, match Office only with a `2x` half-width branch. Treating all straight triangle
  line ends as one geometry family would either break the existing public `line` evidence or leave page 20's
  connector markers narrow.

- Observation: The dependency-free console test runner does not support a `--filter` option even though it
  supports capability groups and slow-test switches.
  Evidence: Running `dotnet run --project tests\Lokad.OoxPdf.Tests --tl:off --nologo -v minimal -- --filter
  PptxSceneBuilderBuildsResolvedNodeLists` executed the whole suite, including slow tests, and completed with
  `263 passed, 0 failed, 0 skipped`. Future focused runs should use `--group`, `--skip-slow`, `--only-slow`, or
  a supported test-specific switch if one is added later.
- Observation: A nonexistent test group used to be accepted by the console runner and select no tests; this has
  been fixed so future validation cannot silently pass an empty selection.
  Evidence: Running `dotnet run --project tests\Lokad.OoxPdf.Tests --tl:off --nologo -v minimal -- --group
  pptx-scene --skip-slow` originally completed with `0 passed, 0 failed, 0 skipped`. After the runner fix, the
  same command exits non-zero and prints the available groups, including `pptx-shapes`, `pptx-charts`,
  `pptx-model`, and `pptx-typography`.
- Observation: The Office baseline floor is not a universal "Arial baseline" rule.
  Evidence: Applying the floor broadly fixed top-anchored rectangular wrap cases, but moved the public
  non-rect small-label probe by about `-1.24 pt` and moved middle/bottom anchored rectangular text in
  `pptx-ladder-03-text-anchor-overflow` to `-1.00/-1.97 pt` deltas. The passing rule is scoped to
  rectangular, top-anchored, default-line-spacing text frames; explicit/absolute line spacing still shows a
  separate uniform `~0.68-0.71 pt` Y residual in `pptx-ladder-04-line-spacing-port`.
- Observation: The current chart text classifier is strong enough for tick labels and legend entries in several
  line-chart cases, but not yet robust enough to distinguish every title-like text placement from data labels.
  Evidence: `pptx-ladder-11-chart-line-trend-port` has tight category/value/legend text parity, yet the
  reference classifies its upper text as `ChartTitleText` while the candidate classifies the comparable item as
  `DataLabelText`. The manifest now gates the stable tick/legend structures and leaves this title/data-label
  ambiguity as future classifier/model work instead of hiding it behind a broad text gate.
- Observation: Pie and doughnut labels should not reuse the category-axis text bucket just because their labels
  sit outside the polar plot box.
  Evidence: The polar graphics classifier already exposes `PolarPlotBoxCandidate` and `PolarSliceCandidate`
  structures. After title and legend detection, remaining polar text is data-label text, not axis text. The
  updated classifier makes the public pie-exploded label gateable as `DataLabelText`; doughnut labels remain
  open because legend detection still disagrees on some legend entries.
- Observation: The chart title scene model preserved title text and `txPr` style, but not the sibling title
  structure that will matter for Office-like placement: `c:overlay`, `c:layout/c:manualLayout`, and `c:spPr`.
  Evidence: `PptxSceneChartTitle` had only `Text`, `IsAutoDeleted`, and `TextStyle`; the scene-builder test now
  locks title overlay, manual-layout factors, fill, and stroke as model data before any renderer consumption.
- Observation: Axis titles are structurally the same chart-title sub-tree under each chart axis, but were not
  available to rendering or inspection as axis-owned scene data.
  Evidence: `PptxSceneChartAxis` preserved tick label style and number format but had no title record; the
  scene-builder fixture now locks category-axis and value-axis titles through `PptxSceneChartAxis.Title`.
- Observation: Default-placement axis titles now have both public Office evidence and initial supported
  Cartesian rendering, but plot-reservation fidelity is still not solved. Evidence: the 2026-05-28
  default-axis-title probe originally exposed horizontal category-title text as `CategoryAxisTickLabel` and
  rotated value-title text as `ValueAxisTickLabel`; `ClassifyPdfChartText.ps1` now emits `AxisTitleText` for
  those axis-lane structures without relying on literal title strings, and the renderer emits matching
  `AxisTitleText` hashes for the supported clustered-column case without the unsupported-layout diagnostic.
  The remaining gap is not whether to draw the titles, but how Office reserves plot and axis-title lanes across
  top/right axes, horizontal bars, overlay modes, rich rotated runs, and inherited chart-style title defaults.
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
- Observation: Raw PDF chart graphics buckets were too weak for clustered column charts: repeated clipping
  rectangles and filled bar unions could look like plot-area or polar-plot evidence even when Office's
  structural signal was the multi-segment gridline stroke plus the matching axis baseline.
  Evidence: `ClassifyPdfChartGraphics.ps1` now derives `GridlineAxisPlotBoxCandidate` from that gridline/axis
  pair, suppresses wide filled-region unions as polar plot boxes, and avoids comparing meaningless line width
  on fill-only marker rectangles. The public composite chart gate now compares the derived plot box and legend
  markers instead of raw clip boxes.
- Observation: The line-chart right-legend classifier is not yet reliable enough for an enforced text gate:
  the same chart whose plot box and axis tick labels match Office within 1 pt still classifies right-legend
  text at different vertical regions in Office and candidate PDFs.
  Evidence: `pptx-ladder-11-chart-line-3series-port` now gates `GridlineAxisPlotBoxCandidate`,
  `HorizontalGridlineGroupCandidate`, `CategoryAxisTickLabel`, and `ValueAxisTickLabel`, but intentionally
  leaves `LegendText` out until right-legend container detection is structural rather than clip-region
  incidental.
- Observation: Horizontal bar charts needed both a transposed axis-pair oracle and a distinct Office-observed
  title/no-legend plot-box rule before they could get an honest structural gate.
  Evidence: `ClassifyPdfChartGraphics.ps1` now also derives `AxisPairPlotBoxCandidate` when the horizontal
  axis aligns with the vertical axis' far end, and the renderer now uses a horizontal-bar title/no-legend plot
  box plus Office-style bottom value-axis placement. `pptx-ladder-11-chart-bar-clustered-port` now gates the
  plot box, the 10-segment vertical gridline group, axis strokes, and bottom value-axis tick labels within
  1 pt. The auto-generated series title now also classifies as `ChartTitleText` after its baseline is anchored
  above the resolved plot box, and the side category labels now gate within 1 pt after their label band was
  realigned. Title X/text-width alignment remains ungated.
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
- Observation: Chart value-axis side placement had already been parsed into the scene model as `Position`,
  but axis stroke rendering still ignored it.
  Evidence: `PptxSceneChartAxis.Position` stores `c:axPos`, and label placement has side-aware helpers, yet
  `RenderBarChart` and `RenderLineChart` still drew primary vertical value-axis strokes at `plotX` and
  secondary strokes at `plotX + plotWidth`. The new synthetic right-side-axis chart would have failed before
  the stroke renderer consumed the same scene/XML side metadata.
- Observation: Secondary value-axis identity and secondary value-axis side are separate concerns.
  Evidence: a synthetic combo chart with primary `valAx` on the right and secondary `valAx` on the left showed
  that selecting secondary axes by `axPos="r"` finds the primary axis or misses the secondary axis entirely.
  Selecting the non-primary `axId` first, then applying `axPos`, produces both vertical axis strokes at their
  declared sides.
- Observation: Horizontal bar chart axis strokes reuse the same renderer blocks as vertical charts with
  swapped axis meaning.
  Evidence: in `RenderBarChart`, the horizontal bar vertical stroke uses `categoryAxisStroke` inside the
  `ValueAxisVisible` branch. This made the existing left-edge coordinate look like value-axis placement until
  the code was read in context. The new category-axis side field is intentionally consumed only for the
  horizontal-bar vertical category-axis stroke.
- Observation: Horizontal bar value-axis strokes were using a value coordinate as an axis-side proxy.
  Evidence: before the current slice, the horizontal bar path drew the horizontal value-axis stroke at
  `zeroY`, so `valAx @axPos="b"` or `valAx @axPos="t"` could not affect the visible axis line. The renderer
  now keeps value-axis bottom/top side as structural axis metadata, while value-to-coordinate mapping remains
  responsible for series, gridlines, and value label anchors.
- Observation: Bar-first combo chart dispatch preserved secondary-axis metadata but skipped sibling line
  plots.
  Evidence: `TryRenderChart` entered the bar renderer as soon as it found a `c:barChart`, rendered additional
  `barChart` siblings, then returned before the later line-chart branch could see any `c:lineChart` sibling.
  This meant a column+line combo could expose secondary value-axis labels while omitting the line series
  itself. Rendering sibling line plots inside the bar-first combo path closes that dispatch gap for the
  supported vertical line geometry and makes plot geometry axis-bound by each chart element's own `axId`.
- Observation: Office line charts with category axes plot data points at category band midpoints, not at the
  left and right edges of the plot box.
  Evidence: the Office PDF for `pptx-ladder-11-chart-line-3series-port` uses plot clip
  `131.64 111.24 553.44 376.8 re`, draws category-axis ticks at about `131.7`, `223.8`, ..., `684.42`, and
  starts the first series at `177.76`, halfway between the first two tick boundaries. The candidate had been
  drawing the first and last data points at the plot edges. After the change, candidate points start at
  `177.738` and end at `638.958`, matching the Office midpoint structure.
- Observation: Default line-series styling is not the hardcoded modern chart palette when the deck theme
  provides Office 2007-style accents.
  Evidence: the public line-chart fixture has no explicit series `spPr` and no chart color-style relationship,
  but its theme accents are `4F81BD`, `C0504D`, and `9BBB59`; Office strokes are blue/red/green at `2.25 w`.
  The old fallback emitted the no-theme palette and `1.5 w`. The renderer now passes theme/chart palette into
  line-series stroke fallback, and the inspected candidate stream emits theme-colored `2.25 w` line strokes.
- Observation: The public no-title/right-legend line chart exposes a real Office plot-box rule that the
  generic chart default could not approximate.
  Evidence: Office draws the plot box at about `x=131.64`, `y=111.24`, `w=553.44`, `h=376.8` inside the
  `72 72 720 432` chart frame. The previous generic default was `x=158.4`, `y=141.12`, `w=547.2`, `h=293.76`.
  The new named metric rule moves the candidate to `131.616`, `111.226`, `553.464`, `376.79`, and the public
  visual gate drops from the pre-existing MAE `5.42615294656636` to `1.9852012201003086`.
- Observation: Office's default clustered-column chart with no title and a bottom legend reserves a compact
  bottom legend band below category labels, and its automatic value scale uses whole major units for a data
  maximum of `5`.
  Evidence: the Office PDF for `pptx-ladder-11-composite-chart-port` draws the plot clip at about
  `95.88 126.36 541.8 293.4 re`, emits value-axis labels `0` through `6`, places category labels at
  `y=108.22`, and places legend text at `y=84.816`. The previous candidate used a narrower plot box, emitted
  a `5.5` top label, and placed legend text at the same vertical band as category labels. After the change,
  candidate inspection shows labels `0` through `6`, no `5.5`, category labels at `107.29`, and legend text
  at `84.88`; the public visual gate drops from MAE `13.6832895688657` to `3.3798655719521604`.

## Decision Log

- Decision: Introduce a behavior-preserving DOCX layout model before attempting more pagination or table
  fidelity fixes.
  Rationale: The private DOCX case still has document-wide page-count drift and table defects. A separate
  layout stage gives those failures a structural home: future fixes can adjust page/block geometry and inspect
  public-safe layout traces without coupling every change to PDF drawing side effects.
  Date/Author: 2026-05-31 / Codex.
- Decision: Make DOCX layout inspection public-safe by construction.
  Rationale: Private DOCX pagination debugging needs per-page and per-block evidence, but private document
  text must not leak into notes or fixtures. Reporting item kind, bounds, counts, and text length is enough to
  locate pagination/table drift while keeping content out of the trace.
  Date/Author: 2026-05-31 / Codex.
- Decision: Preserve DOCX table-cell paragraphs before changing table rendering behavior.
  Rationale: Word table fidelity depends on paragraph boundaries, inherited paragraph style, character style,
  run properties, spacing, numbering, and inline objects inside cells. Keeping the old flattened cell text as a
  compatibility field while adding parsed cell paragraphs gives the renderer a migration path without losing the
  current behavior or baking another table-text shortcut into PDF emission.
  Date/Author: 2026-05-31 / Codex.
- Decision: Make DOCX table-cell text layout-owned before adding Word table pagination rules.
  Rationale: Cell text placement, alignment, wrapping, styling, margins, vertical anchoring, and row-height
  growth are layout decisions. Rendering those decisions directly from `RenderTableRow` would keep table
  fidelity coupled to PDF drawing and hide pagination evidence; attaching text lines to `DocxTableCellLayout`
  gives later Word-compatible table rules a typed place to land.
  Date/Author: 2026-05-31 / Codex.
- Decision: Represent DOCX text lines as styled segments rather than a single first-run style.
  Rationale: Word paragraphs and table cells can change run properties inside one visual line. A line-level
  first-run shortcut loses color, emphasis, underline, and font evidence before PDF emission; segment records
  preserve the structure needed for Office-compatible mixed-run rendering and later font-run shaping.
  Date/Author: 2026-05-31 / Codex.
- Decision: Store DOCX table-cell margins as model data and consume them in layout, not PDF drawing.
  Rationale: `w:tcMar` defines the cell text box and therefore changes wrapping, alignment, and pagination.
  Treating it as a draw-time x-offset would hide structural geometry from layout inspection and prevent later
  row-height and vertical-alignment rules from sharing the same content box.
  Date/Author: 2026-05-31 / Codex.
- Decision: Apply DOCX table-cell vertical alignment after building the cell text block.
  Rationale: `w:vAlign` positions the rendered block inside the cell; applying it before wrapping, margins, or
  paragraph spacing would make later row-height work unstable. Shifting layout-owned text lines preserves a
  reusable content-box model and avoids renderer-only baseline tweaks.
  Date/Author: 2026-05-31 / Codex.
- Decision: Let DOCX table row height grow from measured cell content before pagination.
  Rationale: Word does not treat a too-small row height as permission to clip normal visible cell paragraphs.
  Measuring cell text blocks before row placement gives pagination a structural row height and prevents
  private-case table drift from being handled with downstream PDF clipping or baseline heuristics.
  Date/Author: 2026-05-31 / Codex.
- Decision: Remove the implicit 401-twip default DOCX table row minimum for auto-height rows.
  Rationale: Public no-`tblPrEx` Office fixtures and private-safe graphics inspection show Word lays out
  auto-height rows from measured cell content when no row height is declared. The 401-twip fallback was a
  renderer-side shortcut that added false row height and caused page-flow drift; declared `exact` and
  non-auto row heights still apply through their OOXML tokens.
  Date/Author: 2026-06-01 / Codex.
- Decision: Treat DOCX numbering tab stops as post-marker text placement, not marker origin.
  Rationale: Public compact-bullet Office PDF inspection places the bullet marker at the hanging position
  (`left - hanging`) and the paragraph text at the numbering tab/left position. Using `w:tab w:val="num"` as
  the marker origin moved the label into the text column, which is a structural numbering geometry mismatch.
  The remaining list work stays on tab-stop ownership for text starts, suffix behavior, restarts, and
  continuation lines rather than font or private-style exceptions.
  Date/Author: 2026-06-01 / Codex.
- Decision: Render DOCX table-cell inline images through cell-owned layout records.
  Rationale: Inline images inside cells participate in row height and pagination just like text. Routing them
  through `DocxTableCellLayout` keeps their geometry visible to layout inspection and reuses the shared PDF
  image resource path instead of adding a table-specific image drawing shortcut.
  Date/Author: 2026-05-31 / Codex.
- Decision: Keep the slide-44 chart residual on the shared PPTX Office text-emission track after adding
  structural chart text spacing support.
  Rationale: Reading authored `spc` into chart text styles is a correct model gap to close, but the private
  slide and the public stacked-legend probe show Office can emit chart-text `Tc` even when the authored chart
  style has no nonzero spacing. A fixed chart or slide adjustment would be another heuristic; the long-term
  solution is an Office PDF emission profile that explains text-state splitting across shape text, tables, and
  chart text.
  Date/Author: 2026-05-31 / Codex.
- Decision: Shift the active architecture focus to DOCX using the same Office-PDF-first ladder discipline as
  PPTX.
  Rationale: The PPTX renderer now has a clearer model-first direction, while DOCX still needs a comparable
  typed document/style/layout/pagination/PDF pipeline and bottom-up public coverage. Private DOCX evidence
  remains useful for prioritization, but production changes should be driven by public synthetic Word output.
  Date/Author: 2026-05-31 / Codex.
- Decision: Keep DOCX short-run `Tc` work in diagnostics until Word's text-state decomposition is explained by
  public structural evidence.
  Rationale: Public and private-safe joined summaries now show candidate advance residuals are useful
  observables but do not directly equal Word's emitted `Tc`, and Office nonzero `Tc` spans multiple text
  classes. Promoting raw residuals, digit classes, table roles, font names, or private buckets into rendering
  logic would repeat rejected heuristics. The next rendering change must be a planner-level decomposition that
  can be justified against public Office PDF probes.
  Date/Author: 2026-06-02 / Codex.
- Decision: Resolve `a:grpFill` through ancestor group solid fill instead of adding any custom-geometry
  default-fill or private-slide color fallback.
  Rationale: A public synthetic custom-geometry fixture shows omitted shape fill stays unfilled in Office, but
  the same shape under a group with `a:grpFill` inherits the group fill and emits the same light-gray PDF fill
  structure seen on private slide 80. This is an OOXML inheritance rule, not a map-specific special case.
  Date/Author: 2026-05-31 / Codex.
- Decision: Apply the Office-observed 6 pt minimum only to straight-line `stealth` marker length/width.
  Rationale: Public PDF inspection shows 1 pt `line` and `straightConnector1` stealth markers use 6 pt
  marker geometry, while 2 pt straight lines already land at 6 pt through the existing factor. Preset arcs and
  curved connector stealth outlines remain on their separate structural path because the slide-61 guardrail
  showed that a broad marker-size change can regress explicit-guide arcs.
  Date/Author: 2026-05-31 / Codex.
- Decision: Keep the refreshed page-36/page-79 placement evidence in diagnostic mode and do not change
  baseline fallback constants, table row/column coordinates, or table default-inset adjustments from this
  private deck evidence.
  Rationale: Page 36's dominant frame is already vertically aligned, while page 79's table residuals vary by
  row/cell and are entangled with missing Office text operations and `Tc` buckets. A production change needs a
  public fixture that isolates table-cell vertical anchoring or text-state splitting first.
  Date/Author: 2026-05-31 / Codex.
- Decision: Treat `pptx-ladder-10-table-middle-small-insets` as a counterexample fixture, not as permission for
  a renderer change.
  Rationale: The fixture intentionally shares the no-slack table geometry, explicit small margins, middle
  anchoring, wrapped text, and multi-run cells that looked tempting from private page 79, but Office keeps
  `Tc=0` and has tight baselines. The remaining private branch therefore needs another structural input before
  code changes.
  Date/Author: 2026-05-31 / Codex.
- Decision: Do not implement the secondary Office `/Tf +0.024 pt` branch as a fixed Y-band, top-origin baseline
  band, or local text-frame rule.
  Rationale: Page-height variants of the public-safe fine 21 pt Y-scan shift the secondary source-Y window even
  though the local textbox coordinates and formatting are unchanged. Additional intermediate heights show a
  non-monotonic branch cycle, so a renderer rule keyed to one coordinate band or a simple page-height offset would
  fail across page geometry and would not be structural Office alignment.
  Date/Author: 2026-05-29 / Codex.

- Decision: Keep chart cache point lexing in the scene builder, but keep fallback data-vector validity policy at
  the renderer boundary for now.
  Rationale: Scene construction preserves authored chart cache provenance, including ordinal fallback state, while
  renderer XML fallback vectors have an existing policy that rejects negative point indices for renderable data
  expansion. Sharing the token reader without moving that policy avoids a behavior change and still removes duplicate
  low-level parsing.
  Date/Author: 2026-05-29 / Codex.

- Decision: Keep the secondary Office `/Tf +0.024 pt` work in diagnostic/evidence mode after adding vertical
  overflow provenance.
  Rationale: The refreshed public-safe `wrap13b` branch summary shows the main-grid and secondary branches share
  the same wrap, vertical-overflow source, and autofit modes. Adding a renderer rule keyed to those body-property
  tokens would therefore be a hidden heuristic rather than structural Office alignment.
  Date/Author: 2026-05-29 / Codex.

- Decision: Do not blanket-preserve declared DrawingML table row heights for positive frame slack.
  Rationale: The accepted page-12 fix is limited to sparse filled-row decoration because the dense page-78
  counterexample regressed under the broader declared-row rule. Office can preserve declared row heights for a
  single filled decoration row while still behaving closer to frame-height allocation for dense styled tables.
  Until a public Office ladder explains the full row-allocation rule, keep the branch narrow and evidence-bound.
  Date/Author: 2026-05-30 / Codex.

- Decision: Split straight triangle line-end marker geometry by connector preset.
  Rationale: Public evidence already protected the `line` preset's narrower triangle marker width. Private
  page-20 PDF structure shows `straightConnector1` uses a wider Office marker with the same line-end token and
  stroke width. A preset-specific branch is more structural than tuning by private coordinates, and it preserves
  both observable Office behaviors.
  Date/Author: 2026-05-30 / Codex.

- Decision: Render default-placement chart axis titles from inside the native chart branches that own the
  resolved `ChartLayout`, rather than from the older frame-only post-render axis-title pass.
  Rationale: Office placement is tied to the plot box and the surrounding axis/title reserves. Keeping default
  title emission next to branch-specific plot-box calculation gives future Office-PDF-backed reservation rules
  a structural home, while the post-render path remains appropriate for manual-layout titles and diagnostics
  on unsupported chart paths.

- Decision: Use resolved OS/2 Windows ascender plus descender as the Office-aligned visible-line height for
  middle-anchor estimation under default line spacing.
  Rationale: Public Office PDF probes match this structural font-table rule closely, and the same font-box
  metrics are already observable through the text layout model. This avoids another private-deck coordinate
  adjustment and keeps anchoring tied to resolved font state.
  Date/Author: 2026-05-28 / Codex.
- Decision: Keep the empty-paragraph middle-anchor rule explicitly temporary and evidence-bounded.
  Rationale: Public trailing-empty probes do not yet justify a clean Office formula; using a named compromise
  is better than baking a narrow private-slide adjustment into the renderer. The backlog now requires a
  public probe matrix before this can be considered complete.
  Date/Author: 2026-05-28 / Codex.
- Decision: Keep table-cell vertical anchoring on the prior typographic font-box metric while shape text uses
  the new OS/2 Windows font-box metric for visible default-spaced lines.
  Rationale: The public Office probe that motivated the Windows font-box rule covered shape text frames, not
  table cells. Existing table-cell tests encode separate evidence and failed when the new rule was applied
  globally, so the shared estimator now makes the surface distinction explicit instead of silently changing
  table-cell behavior.
  Date/Author: 2026-05-28 / Codex.
- Decision: Add a table-specific text-frame inspection entry point instead of broadening
  `InspectTextFrameModels` to include table cells.
  Rationale: Existing text-frame model tests use shape-only counts and placeholder behavior. A separate
  `InspectTableTextFrameModels` method makes the table adapter boundary testable without changing the meaning
  of shape text snapshots, and it gives future table-text migration a stable place to prove source ownership.
  Date/Author: 2026-05-28 / Codex.
- Decision: Keep raw `EndParagraphProperties` as source evidence but make layout consume a resolved
  end-paragraph style record.
  Rationale: Raw OOXML remains useful for diagnostics and for future unsupported-token ladders, but layout
  should not rediscover font metrics from XML once the text model has enough context to resolve them. Carrying
  the resolved end-paragraph font size/typeface/bold/italic state keeps the model/layout boundary explicit
  without changing empty-paragraph rendering.
  Date/Author: 2026-05-28 / Codex.
- Decision: Resolve paragraph tab stops through the same direct-before-default paragraph style ladder used by
  the rest of the PPTX text model.
  Rationale: Tab expansion is layout behavior, but the authored tab-stop list is paragraph style state. Keeping
  the resolved point positions in `PptxTextParagraphModelSnapshot` makes the PDF-level tab outcome traceable to
  DrawingML structure and avoids compensating for inherited tabs with ad hoc span offsets during rendering.
  Date/Author: 2026-05-28 / Codex.
- Decision: Treat bullet marker identity as paragraph model state, not line-layout state.
  Rationale: Whether a paragraph has no bullet, a character bullet, an auto-numbered bullet, a symbol-font
  mapping, or an unsupported blip bullet is determined by the resolved paragraph style cascade before line
  construction starts. Moving that marker record into `PptxTextParagraphModel` keeps future Office numbering
  and bullet parity tied to OOXML structure while leaving the still-open color/size bullet style path explicit.
  Date/Author: 2026-05-28 / Codex.
- Decision: Treat bullet visual style tokens as paragraph model state, with run style only as the documented
  fallback.
  Rationale: `buFont`, `buClr`, and `buSz*` are authored paragraph properties. Layout needs the current run
  style only when those bullet properties are absent, so carrying the bullet style tokens in
  `PptxParagraphBulletModel` keeps the source decision structural and prevents future Office bullet tuning
  from becoming another renderer-local XML branch.
  Date/Author: 2026-05-28 / Codex.
- Decision: Store empty-paragraph spacing as resolved paragraph model values.
  Rationale: Empty paragraph spacing is a paragraph style decision whose numeric value depends on the resolved
  end-paragraph font size. Computing it during text-model construction keeps vertical anchoring and line layout
  on the same typed facts and removes another layout-stage XML read without changing the Office-like spacing
  formula.
  Date/Author: 2026-05-28 / Codex.
- Decision: Treat empty-paragraph layout participation as paragraph model state.
  Rationale: Whether an empty paragraph contributes vertical space follows from authored paragraph/end-run
  structure. Layout should consume that already-classified fact, not inspect XML presence, so future Office
  empty-paragraph work can refine the model record without adding more layout-side source probes.
  Date/Author: 2026-05-28 / Codex.
- Decision: Extend `PptxInspect` with frame/paragraph/layout schemas instead of overloading glyph-run JSON.
  Rationale: The slide-17 investigation needed frame bodyPr sources, paragraph spacing, line-box metrics, and
  glyph counts together. Separate private-safe schema files keep those layers inspectable without exposing
  private text by default and without weakening the existing glyph-run contract.
  Date/Author: 2026-05-28 / Codex.

- Decision: Remove renderer-local XML fallback once the scene already owns an equivalent typed record.
  Rationale: The scene builder and renderer background fallback used the same solid-color-and-alpha resolver.
  Keeping both paths lets incomplete scene records be silently repaired during PDF emission, which works
  against the long-term goal of making parse/model/render ownership explicit and testable. Unsupported
  background variants should be added to `PptxSceneBackground` rather than reparsed in `RenderBackground`.
  Date/Author: 2026-05-27 / Codex.
- Decision: Keep the legacy inherited XML source list only as an explicit text fallback input, not as a
  separate inheritance owner.
  Rationale: Text snapshots and layout still need raw inherited XML until the style cascade is fully typed, but
  the scene already owns the loaded master and layout documents. Removing `PptxSlideInheritance` avoids a
  second ownership abstraction while preserving behavior for the remaining text migration.
  Date/Author: 2026-05-27 / Codex.
- Decision: Use the current `PptxScene` spine as the migration target instead of introducing a second normalized
  presentation model.
  Rationale: Source inspection shows the scene already owns slide inheritance, relationship maps, draw-order
  nodes, backgrounds, image targets, typed table/chart/text records, and inspection snapshots. The long-term
  gap is field ownership and fallback retirement, not a missing top-level model. Adding another model would
  duplicate ownership and delay removal of renderer XML/package heuristics.
  Date/Author: 2026-05-27 / Codex.
- Decision: Preserve raw OOXML enum tokens in scene records whenever the renderer also needs a normalized or
  PDF-ready value.
  Rationale: A normalized value such as PDF line cap `1` is useful for emission but hides whether the source
  token was `rnd`, unknown, or absent. Keeping the raw token next to the normalized value makes unsupported
  enum variants visible in snapshots and public tests, which supports systematic ladders instead of one-off
  rendering branches.
  Date/Author: 2026-05-27 / Codex.
- Decision: Fail unknown console test groups instead of treating them as empty selections.
  Rationale: Focused validation is only useful when the selector is known to exercise tests. A mistyped or
  nonexistent group that reports `0 passed, 0 failed, 0 skipped` can make architecture migrations look covered
  while no assertions ran, so the runner now reports the known group list and exits non-zero.
  Date/Author: 2026-05-27 / Codex.
- Decision: Scope the PPTX Office baseline floor to rectangular top-anchored text frames using default line
  spacing, and keep non-rect, vertically anchored, explicit-spacing, and absolute-spacing frames on resolved
  font metrics until separate Office-PDF evidence justifies a broader rule.
  Rationale: Public evidence split the problem by text-frame context. The floor collapses `~1.25 pt` ordinary
  top-anchored rectangular residuals to near zero, but it regresses non-rect and vertically anchored cases and
  does not explain explicit line-spacing residuals. Encoding the context in `PptxTextFrameModel` keeps the rule
  structural while preventing a global metric nudge from hiding distinct Office behaviors.
  Date/Author: 2026-05-26 / Codex.
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
- Decision: Treat the public composite chart failure as an Office-PDF structural chart layout slice, not as
  a private-deck tuning problem.
  Rationale: The failure was explained by inspected Office PDF structure: plot-box geometry, whole-number
  value-axis ticks, packed bottom legend placement, and legend marker/text geometry. Naming these as
  chart-layout metric rules keeps the current approximation auditable while future chart oracle tooling is
  built.
  Date/Author: 2026-05-25 / Codex.
- Decision: Chart structural gates should compare derived Office-like primitives before raw graphics
  operations whenever the higher-level primitive can be recovered from PDF structure.
  Rationale: Raw clips, filled bars, and graphics-state attributes can be incidental or reused across text and
  chart drawing. Derived structures such as gridline-plus-axis plot boxes, legend marker rectangles, and typed
  tick/legend text buckets are closer to the chart semantics needed to eliminate renderer heuristics.
  Date/Author: 2026-05-25 / Codex.
- Decision: Preserve transformed PDF path commands in the inspection oracle when bounds are not enough, and use
  those commands to derive semantic chart geometry before changing renderer constants.
  Rationale: Radar spoke bounds obscure the actual polar center because a five-sided polygon is asymmetric.
  Repeated `m`/`l` path commands reveal the Office center and radius directly. This showed that filled radar is
  already structurally aligned while two-series radar is not, so a shared center/radius tweak would regress a
  passing public case and would not be an Office-aligned rule.
  Date/Author: 2026-05-25 / Codex.
- Decision: Split current radar center/radius metric rules by resolved radar style, not by fixture identity.
  Rationale: The path-command oracle shows Office uses different observable spoke geometry for filled radar and
  marker/default radar on the public cases. Encoding this as a `c:radarStyle`-driven rule closes the marker case
  while preserving the filled-radar gate; the next architecture step is still a typed radar layout resolver that
  owns these values instead of renderer constants.
  Date/Author: 2026-05-25 / Codex.
- Decision: Treat chart structural oracle work as coverage and classification expansion, not as a new tooling
  greenfield.
  Rationale: `ClassifyPdfChartGraphics.ps1`, `ClassifyPdfChartText.ps1`, `ComparePdfGraphicsOperations.ps1`, and
  `CheckVisualCase.ps1` already provide public manifest-driven chart structure gates. The 2026-05-26 inventory
  found 23 of 37 public chart cases graphics-gated and 8 of 37 text-gated. The durable architecture work is to
  extend those gates to ungated chart families and refine semantic buckets until renderer constants can be
  replaced by Office-observed structure, while preserving existing passing gates.
  Date/Author: 2026-05-26 / Codex.
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
- Decision: Treat chart axis-line side as axis-owned metadata, separate from tick-label side.
  Rationale: OOXML `c:axPos` locates the axis line, while `c:tickLblPos` can independently place labels high,
  low, next to the axis, or hide them. Keeping these as separate renderer inputs prevents a label-placement
  shortcut from becoming a hardcoded axis geometry rule.
  Date/Author: 2026-05-25 / Codex.
- Decision: Discover secondary value axes by value-axis identity before applying side placement.
  Rationale: Office chart XML can place primary and secondary axes on either side. Using `axPos="r"` as the
  definition of secondary bakes a layout convention into model selection and fails when the primary axis is
  right-sided or the secondary axis is left-sided.
  Date/Author: 2026-05-25 / Codex.
- Decision: Keep category-axis side as a separate chart-axis style field instead of reusing value-axis side.
  Rationale: In horizontal bar charts, the category axis is the vertical axis, while in column and line charts
  the value axis is vertical. Separate fields make the orientation swap explicit and avoid a new heuristic
  that treats all vertical strokes as value axes.
  Date/Author: 2026-05-25 / Codex.
- Decision: Represent horizontal value-axis bottom/top side separately from vertical value-axis left/right side.
  Rationale: `c:valAx/c:axPos` uses the same OOXML field for both vertical and horizontal value axes, but the
  consumed geometry is different after chart orientation is known. Keeping bottom/top as its own resolved
  style field prevents a left/right axis-line rule from leaking into horizontal bar charts and keeps value
  coordinates out of axis-side selection.
  Date/Author: 2026-05-25 / Codex.
- Decision: Combo plot geometry is owned by each chart element's axis bindings, not by the first chart type in
  dispatch.
  Rationale: Office combo charts can mix chart families while sharing categories and using independent value
  axes. If the renderer lets the first matched chart type decide all series geometry, it bakes a dispatch
  artifact into chart semantics and drops or mis-scales sibling plots. Each plot element should resolve its
  own `axId` references before geometry, labels, and legend entries are emitted.
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
- Decision: Treat line-chart category positions as category band centers for the supported category-axis path.
  Rationale: Office's PDF places line-series points between category tick boundaries. Drawing them at plot-box
  edges was a structural axis error, not a pixel offset. This keeps category labels and line geometry on the
  same band model.
  Date/Author: 2026-05-25 / Codex.
- Decision: Add a narrowly named no-title/right-legend line-chart plot-box rule, but keep it in
  `PptxChartMetricRules` as an inventory item rather than treating metric constants as the final architecture.
  Rationale: The public Office PDF provides exact plot-box evidence for this common layout, and the generic
  default was visibly wrong. The rule is constrained by title/legend state and still yields to manual layout;
  longer term, chart plot layout should be resolved from typed chart layout records and structural oracle data.
  Date/Author: 2026-05-25 / Codex.
- Decision: Line-chart default stroke fallback must resolve chart palette and theme before hardcoded palette
  colors.
  Rationale: Office used theme accent colors for the public line-chart fixture even without explicit series
  `spPr` or a chart color-style part. Keeping the fallback theme-aware removes a hardcoded narrow palette from
  the line and legend path.
  Date/Author: 2026-05-25 / Codex.

## Outcomes & Retrospective

- 2026-05-28: The middle-anchor slice produced a structural improvement rather than a private coordinate
  patch. Visible default-spaced text now uses resolved font-table metrics for anchoring, the private-safe
  inspector exposes the model layers needed to audit schema issues, and the affected private frame is
  vertically close to Office. The remaining gap is deliberately narrower: explain trailing empty-paragraph
  height and same-line emphasized-run glyph segmentation from public Office PDF structure.

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
- The public `pptx-ladder-11-chart-line-3series-port` gate now passes through structural chart alignment:
  default line colors/width, category midpoint positions, and the no-title/right-legend plot box are aligned
  to Office PDF operators, and the manifest now guards the derived gridline/axis plot box, gridline group,
  category tick labels, and value tick labels. The public `pptx-ladder-11-composite-chart-port` gate also now
  passes after aligning the default clustered-column no-title/bottom-legend plot box, whole-number value
  scale, and packed bottom legend placement to Office PDF evidence. The composite chart is now protected by
  structural gates for the derived gridline/axis plot box, legend markers, category tick labels, value tick
  labels, and legend text. The public clustered horizontal-bar chart now additionally gates its derived plot
  box, two axis strokes, 10-segment vertical gridline group, and bottom value-axis tick-label text positions.
  These remove both cases from the pre-existing failure list and turn the bar case into a structural-regression
  surface, while broader chart-family layout work remains open because the current named metric rules still
  need to be replaced by systematic chart structural oracle tooling.

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

Current validation baseline:

- DOCX numbering label-start slice:
  `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed; `docx-numbering --skip-slow` passed;
  public `docx-ladder-03-compact-bullet-spacing` run `20260601-185503` improved from MAE `0.810522` to
  `0.803030` while moving the bullet marker to the Office-observed hanging position; public `docx-numbering`
  and `docx-ladder-03-docgrid-list-use-fe` stayed unchanged. Private DOCX acceptance run `20260601-185704`
  stayed at `16/16` pages, zero dimension mismatches, no diagnostics, `MAE=9.982157`, changed16 `0.103396`.
- DOCX border diagnostics/fidelity:
  `dotnet run --project tests\Lokad.OoxPdf.Tests --tl:off --nologo -v minimal -- --group docx-core --skip-slow`
  passed `16`; `docx-tables --skip-slow` passed `62`; `docx-page --skip-slow` passed `17`; full solution
  build passed. Private DOCX run `20260601-025057` compared `16/16` pages with zero dimension mismatches,
  no diagnostics, `MAE=12.494853`, and changed16 `0.116738`. This locks the table-border diagnostic addition
  as behavior-neutral on the current private document while keeping non-single Word border styles visible as
  open rendering work.
- Private deck: `pwsh tools\CheckPrivateCase.ps1 -Case private-cases\lokad-value-based.json` run
  `20260531-123236` compared `84/84` pages. Page 80 improved from run `20260531-120513` MAE `3.62`,
  changed16 `0.11`, SSIM `0.91` to MAE `0.96`, changed16 `0.02`, SSIM `0.96`; PDF inspection now shows the
  candidate page has `795` fill operations, including the `785` `g:0.851` group-fill map fragments that match
  Office's structural bucket.
- Private deck: `pwsh tools\CheckPrivateCase.ps1 -Case private-cases\lokad-value-based.json` run
  `20260531-124414` compared `84/84` pages. Page 5 aggregate metrics stayed effectively unchanged
  (`MAE=2.42`, changed16 `0.05`, SSIM `0.93`) because the remaining slide residuals are elsewhere, but PDF
  inspection now matches the Office 6 pt bounds for the green horizontal stealth arrowheads and the vertical
  stealth connector.
- Private deck: `pwsh tools\CheckPrivateCase.ps1 -Case private-cases\lokad-value-based.json` run
  `20260531-130027` compared `84/84` pages with empty diagnostics. The focused slide 44 chart page improved
  from run `20260531-124414` MAE `2.439032`, changed16 `0.046995`, SSIM `0.931097` to MAE `1.702704`,
  changed16 `0.039896`, SSIM `0.960047`; focused PDF inspection shows the candidate chart plot axes now align
  to Office (`137.71..296.69` vs `137.70..296.70`).
- Public chart bottom-legend probe: `pptx-ladder-11-chart-column-stacked-bottom-legend-probe` run
  `20260531-130415` passed with structural plot-box/axis gates. Its intentionally loose raster/text gates
  now reflect remaining chart-text and bottom-legend layout gaps rather than stacked-column compound-fill
  emission.
- Public stacked-column fixtures: `pptx-ladder-11-chart-column-stacked-port` and
  `pptx-ladder-11-chart-column-100-stacked-port` run `20260531-131506` passed after adding `FilledRegion`
  structural gates. Both now match Office's stacked series fill path operators and path-command counts
  (`16` segments / `4` moves for four-category stacked columns, `12` segments / `3` moves for the 100%
  stacked probe).
- PPTX chart/text cleanup:
  `dotnet run --project tests\Lokad.OoxPdf.Tests --tl:off --nologo -v minimal -- --group pptx-charts --skip-slow`
  passed `148` tests; `dotnet run --project tests\Lokad.OoxPdf.Tests --tl:off --nologo -v minimal -- --group
  pptx-typography --skip-slow` passed `138` tests with `2` skipped. The public
  `pptx-ladder-11-chart-column-stacked-bottom-legend-probe` visual case passed in run `20260531-141703`.
  Private deck run `20260531-141742` compared `84/84` pages with empty diagnostics; page 44 stayed at MAE
  `1.281596258`, changed16 `0.032060185`, SSIM `0.973682740`, confirming the remaining chart-text `Tc`
  difference is not authored OOXML spacing.
- DOCX layout split validation:
  `dotnet run --project tests\Lokad.OoxPdf.Tests --tl:off --nologo -v minimal -- --group docx-core --skip-slow`
  passed `3`; `docx-page` passed `8`; `docx-text` passed `6`; `docx-numbering` passed `3`;
  `docx-images` passed `2`; and `docx-tables` passed `10`. All public `docx-*` visual cases passed in the
  sweep ending with `docx-tables` run `20260531-143206`. Private DOCX run
  `artifacts/private-visual/user-requirements-spec/20260531-142940` compared against 16 reference pages with
  14 candidate pages and the known DOCX diagnostics (`DOCX_NUMBERING_INDENT`,
  `DOCX_STYLE_PARAGRAPH_KEEP_RULE`, `DOCX_STYLE_PARAGRAPH_SPACING`, `DOCX_STYLE_TABLE_STYLE`,
  `DOCX_UNSUPPORTED_TABLE_HEADER_ROW`, `DOCX_UNSUPPORTED_TABLE_STYLE`), confirming pagination/table fidelity
  remains the next high-impact DOCX target after the behavior-preserving layout boundary.
- DOCX layout snapshot validation:
  after adding the private-safe layout snapshot, the DOCX group sweep passed again (`docx-core` `4`,
  `docx-page` `8`, `docx-text` `6`, `docx-numbering` `3`, `docx-images` `2`, `docx-tables` `10`), and
  `docx-ladder-01-plain-paragraph` passed in visual run `20260531-143530`.
- DOCX table-cell paragraph preservation validation:
  `docx-tables --skip-slow` passed `11` tests after preserving parsed table-cell paragraphs with inherited
  paragraph and character styles. The full DOCX group sweep passed (`docx-core` `4`, `docx-page` `8`,
  `docx-text` `6`, `docx-numbering` `3`, `docx-images` `2`, `docx-tables` `11`), and the public `docx-tables`
  visual case passed in run `20260531-143856`.
- DOCX table-cell text-line layout validation:
  `docx-tables --skip-slow` passed `12` tests after moving cell text drawing to layout-owned text lines.
  Public `docx-tables` visual case passed in run `20260531-144643`. The full DOCX group sweep passed
  (`docx-core` `4`, `docx-page` `8`, `docx-text` `6`, `docx-numbering` `3`, `docx-images` `2`,
  `docx-tables` `12`).
- DOCX styled text-segment validation:
  after adding styled segments to DOCX text lines, `docx-tables --skip-slow` passed `12` tests, public
  `docx-tables` visual case passed in run `20260531-145023`, and the full DOCX group sweep passed
  (`docx-core` `4`, `docx-page` `8`, `docx-text` `6`, `docx-numbering` `3`, `docx-images` `2`,
  `docx-tables` `12`).
- DOCX table-cell margin validation:
  after preserving `w:tcMar` tokens and applying `dxa` margins to cell text layout, `docx-tables --skip-slow`
  passed `14` tests, public `docx-tables` visual case passed in run `20260531-145403`, and the full DOCX group
  sweep passed (`docx-core` `4`, `docx-page` `8`, `docx-text` `6`, `docx-numbering` `3`, `docx-images` `2`,
  `docx-tables` `14`).
- DOCX paragraph indent validation:
  after preserving cascaded paragraph `w:ind` left/right/first-line/hanging tokens and applying them to ordinary
  paragraph wrapping, `docx-text --skip-slow` passed `18` tests, `docx-page --skip-slow` passed `17`,
  `docx-tables --skip-slow` passed `58`, `docx-numbering --skip-slow` passed `11`, and `dotnet build
  Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed. Public `docx-ladder-02-paragraph-indents` visual run
  `20260601-022308` is dimension-stable at `MAE=1.104564`, changed16 `0.011547`; private DOCX run
  `20260601-022152` compared `16/16` pages with zero dimension mismatches, no diagnostics, `MAE=12.481106`,
  changed16 `0.116738`.
- DOCX logical indent validation:
  after resolving `w:start`/`w:end` as effective left-to-right start/end indents and diagnosing character-unit
  indent attributes explicitly, `docx-text --skip-slow` passed `18`, `docx-core --skip-slow` passed `16`,
  `docx-numbering --skip-slow` passed `11`, `docx-page --skip-slow` passed `17`, `docx-tables --skip-slow`
  passed `58`, and `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed. Public visual
  `docx-ladder-02-paragraph-logical-indents` run `20260601-022902` is dimension-stable at `MAE=0.094507`,
  changed16 `0.002454`, SSIM `0.986713`; private DOCX run `20260601-022931` compared `16/16` pages with zero
  dimension mismatches, no diagnostics, `MAE=12.481106`, changed16 `0.116738`.
- DOCX table-cell vertical alignment validation:
  after shifting layout-owned cell text blocks for `w:vAlign`, `docx-tables --skip-slow` passed `15` tests,
  public `docx-tables` visual case passed in run `20260531-145618`, and the full DOCX group sweep passed
  (`docx-core` `4`, `docx-page` `8`, `docx-text` `6`, `docx-numbering` `3`, `docx-images` `2`,
  `docx-tables` `15`).
- DOCX table row content-height validation:
  after expanding row height from measured cell text blocks, `docx-tables --skip-slow` passed `16` tests,
  public `docx-tables` visual case passed in run `20260531-145846`, and the full DOCX group sweep passed
  (`docx-core` `4`, `docx-page` `8`, `docx-text` `6`, `docx-numbering` `3`, `docx-images` `2`,
  `docx-tables` `16`).
- DOCX table-cell numbering validation:
  after adding explicit coverage for numbered paragraphs inside table cells, `docx-tables --skip-slow` passed
  `17` tests and the full DOCX group sweep passed (`docx-core` `4`, `docx-page` `8`, `docx-text` `6`,
  `docx-numbering` `3`, `docx-images` `2`, `docx-tables` `17`).
- DOCX table-cell inline image validation:
  after adding cell-owned inline image layouts and rendering them through the shared PDF image path,
  `docx-tables --skip-slow` passed `18` tests, `docx-images --skip-slow` passed `2` tests, public
  `docx-tables` visual case passed in run `20260531-150408`, and the full DOCX group sweep passed
  (`docx-core` `4`, `docx-page` `8`, `docx-text` `6`, `docx-numbering` `3`, `docx-images` `2`,
  `docx-tables` `18`).
- DOCX row content-height correction validation:
  after removing the legacy baseline inset from required row content height, `docx-tables --skip-slow` passed
  `18` tests, public `docx-tables` visual case passed in run `20260531-150846`, and the full DOCX group sweep
  passed (`docx-core` `4`, `docx-page` `8`, `docx-text` `6`, `docx-numbering` `3`, `docx-images` `2`,
  `docx-tables` `18`). Private DOCX run `20260531-150625` showed the too-tall-row regression
  (`18` candidate pages vs `16` reference pages); after the correction, private run `20260531-150813` moved to
  `15` candidate pages vs `16` reference pages with the same known diagnostic categories, confirming the next
  pagination gap is not a blanket table-row growth problem.
- DOCX paragraph spacing/keep-token preservation validation:
  after adding typed paragraph spacing and keep-rule records, the full DOCX group sweep passed again:
  `docx-core` `4`, `docx-page` `8`, `docx-text` `7`, `docx-numbering` `3`, `docx-images` `2`, and
  `docx-tables` `18`. The new public synthetic test preserves style/default/direct `w:spacing` source tokens,
  `w:contextualSpacing`, and keep/widow on/off values without changing rendering behavior.
- DOCX keep-rule pagination validation:
  after applying typed `keepLines` and `keepNext` in the layout stage, the full DOCX group sweep passed
  (`docx-core` `4`, `docx-page` `10`, `docx-text` `7`, `docx-numbering` `3`, `docx-images` `2`,
  `docx-tables` `18`). All public `docx-*` visual cases passed in the sweep ending with `docx-tables` run
  `20260531-152113`. Private DOCX run `20260531-152007` matched the reference page count (`16/16`) with zero
  dimension mismatches; MAE was `17.339630` and changed16 was `0.151959`.
- DOCX numbering-indent validation:
  after preserving `w:ind` on numbering levels and applying first-pass numbered-paragraph offsets, the full
  DOCX group sweep passed (`docx-core` `4`, `docx-page` `10`, `docx-text` `7`, `docx-numbering` `4`,
  `docx-images` `2`, `docx-tables` `18`). Public `docx-numbering` visual run `20260531-152648` and
  `docx-tables` visual run `20260531-152654` passed. Private DOCX run `20260531-152705` kept `16/16` pages
  with zero dimension mismatches and improved MAE to `16.995226`, changed16 to `0.149573`.
- DOCX table-header validation:
  after preserving `w:tblHeader` and repeating contiguous table header rows across table page breaks,
  `docx-core --skip-slow` passed `4`, `docx-tables --skip-slow` passed `20`, public `docx-tables` visual run
  `20260531-153109` passed, and private DOCX run `20260531-153115` stayed at `16/16` pages, zero dimension
  mismatches, MAE `16.995226`, changed16 `0.149573`.
- DOCX table-style shading validation:
  after preserving table style IDs and applying whole-style cell shading, the full DOCX group sweep passed
  (`docx-core` `4`, `docx-page` `10`, `docx-text` `7`, `docx-numbering` `4`, `docx-images` `2`,
  `docx-tables` `21`), public `docx-tables` visual run `20260531-153530` passed, and private DOCX run
  `20260531-153536` stayed at `16/16` pages, zero dimension mismatches, MAE `16.995226`, changed16 `0.149573`.
- DOCX conditional table-style shading validation:
  after applying `w:tblStylePr` conditional cell shading, the full DOCX group sweep passed (`docx-core` `4`,
  `docx-page` `10`, `docx-text` `7`, `docx-numbering` `4`, `docx-images` `2`, `docx-tables` `21`), public
  `docx-tables` visual run `20260531-153916` passed, and private DOCX run `20260531-153922` stayed at `16/16`
  pages, zero dimension mismatches, MAE `16.817625`, changed16 `0.148589`.
- DOCX table-border validation:
  after applying conditional `w:tcBorders` and rendering per-edge borders, the full DOCX group sweep passed
  (`docx-core` `4`, `docx-page` `10`, `docx-text` `7`, `docx-numbering` `4`, `docx-images` `2`,
  `docx-tables` `22`), public `docx-tables` visual run `20260531-154329` passed, and private DOCX run
  `20260531-154335` stayed at `16/16` pages, zero dimension mismatches, MAE `16.004348`, changed16 `0.142372`.
- DOCX contextual-spacing validation:
  after applying `w:contextualSpacing` for adjacent same-style body paragraphs, the full DOCX group sweep
  passed (`docx-core` `4`, `docx-page` `10`, `docx-text` `8`, `docx-numbering` `4`, `docx-images` `2`,
  `docx-tables` `22`). Private DOCX run `20260531-154620` stayed at `16/16` pages, zero dimension mismatches,
  MAE `16.004348`, changed16 `0.142372`.
- DOCX table-style margin validation:
  after applying style-level `w:tblCellMar`, the full DOCX group sweep passed (`docx-core` `4`, `docx-page`
  `10`, `docx-text` `8`, `docx-numbering` `4`, `docx-images` `2`, `docx-tables` `22`). Private DOCX run
  `20260531-154933` stayed at `16/16` pages, zero dimension mismatches, MAE `16.003190`, changed16 `0.142369`.
- DOCX table preferred-width validation:
  after applying `w:tblW` `dxa` preferred widths to grid scaling, the full DOCX group sweep passed
  (`docx-core` `4`, `docx-page` `10`, `docx-text` `8`, `docx-numbering` `4`, `docx-images` `2`,
  `docx-tables` `23`). Private DOCX run `20260531-155308` stayed at `16/16` pages, zero dimension mismatches,
  MAE `16.003190`, changed16 `0.142369`.
- DOCX cell preferred-width validation:
  after preserving `w:tcW` and using complete first-row `dxa` widths as the column basis, the full DOCX group
  sweep passed (`docx-core` `4`, `docx-page` `10`, `docx-text` `8`, `docx-numbering` `4`, `docx-images` `2`,
  `docx-tables` `24`). Private DOCX run `20260531-155930` stayed at `16/16` pages, zero dimension mismatches,
  MAE `15.928048`, changed16 `0.142037`.
- DOCX segmented numbering validation:
  after separating body numbering markers from paragraph text starts, the full DOCX group sweep passed
  (`docx-core` `4`, `docx-page` `10`, `docx-text` `8`, `docx-numbering` `4`, `docx-images` `2`,
  `docx-tables` `24`). Private DOCX run `20260531-161037` stayed at `16/16` pages, zero dimension mismatches,
  MAE `15.888467`, changed16 `0.141944`.
- DOCX soft line-break validation:
  after preserving plain `w:br` as soft line breaks, the full DOCX group sweep passed (`docx-core` `4`,
  `docx-page` `10`, `docx-text` `9`, `docx-numbering` `4`, `docx-images` `2`, `docx-tables` `24`).
  Private DOCX run `20260531-162451` stayed at `16/16` pages, zero dimension mismatches, MAE `15.889775`,
  changed16 `0.141949`.
- DOCX authored-whitespace validation:
  after preserving leading/trailing/repeated spaces in wrapping, the full DOCX group sweep passed
  (`docx-core` `4`, `docx-page` `10`, `docx-text` `10`, `docx-numbering` `4`, `docx-images` `2`,
  `docx-tables` `24`). Private DOCX run `20260531-162756` stayed at `16/16` pages, zero dimension mismatches,
  MAE `15.889775`, changed16 `0.141949`.
- DOCX line-based paragraph-spacing validation:
  after converting `beforeLines`/`afterLines` into resolved paragraph spacing, `docx-text --skip-slow` passed
  `11` tests and `docx-page --skip-slow` passed `10`. Private DOCX run `20260531-174219` stayed at `16/16`
  pages, zero dimension mismatches, MAE `15.889775`, changed16 `0.141949`; diagnostics still include
  `DOCX_STYLE_PARAGRAPH_SPACING`, which remains appropriate for autospacing, cascade, and collapse semantics.
- DOCX table-style border validation:
  after applying table-style `w:tblPr/w:tblBorders`, `docx-tables --skip-slow` passed `27`. Private DOCX run
  `20260531-174458` stayed at `16/16` pages, zero dimension mismatches, MAE `15.889775`, changed16 `0.141949`;
  table-style diagnostics remain for `tblLook`, conditional precedence, and table-style paragraph/run layers.
- DOCX table-look token validation:
  after preserving `w:tblLook` tokens without applying unverified gating semantics, `docx-tables --skip-slow`
  passed `28`. Private DOCX run `20260531-175135` stayed at `16/16` pages, zero dimension mismatches, MAE
  `15.889775`, changed16 `0.141949`. The rejected gating experiment is documented as an open Word-fixture
  requirement.
- DOCX numbering start-override validation:
  after applying `w:lvlOverride/w:startOverride`, `docx-numbering --skip-slow` passed `6`. Private DOCX run
  `20260531-175443` stayed at `16/16` pages, zero dimension mismatches, MAE `15.889775`, changed16 `0.141949`;
  `DOCX_NUMBERING_INDENT` remains open for tab-stop ownership, bullet fonts, and multilevel/restart semantics.
- DOCX multilevel numbering validation:
  after resolving `%1` through `%9` from active counters and resetting deeper counters on parent increments,
  `docx-numbering --skip-slow` passed `7`. Private DOCX run `20260531-175714` stayed at `16/16` pages, zero
  dimension mismatches, MAE `15.889775`, changed16 `0.141949`.
- DOCX table grid-span validation:
  after preserving and applying `w:gridSpan`, `docx-tables --skip-slow` passed `30`. Private DOCX run
  `20260531-180144` stayed at `16/16` pages, zero dimension mismatches, MAE `15.889775`, changed16 `0.141949`.
- DOCX table cell-spacing validation:
  after preserving and applying direct `w:tblCellSpacing`, `docx-tables --skip-slow` passed `31`. Private DOCX
  run `20260531-181847` stayed at `16/16` pages, zero dimension mismatches, MAE `15.889775`, changed16
  `0.141949`.
- DOCX percentage table-width validation:
  after applying `w:tblW w:type="pct"`, `docx-tables --skip-slow` passed `32`. Private DOCX run
  `20260531-182041` stayed at `16/16` pages, zero dimension mismatches, and improved to MAE `15.864724`,
  changed16 `0.141689`.
- DOCX numbering continuation-width validation:
  after wrapping numbered paragraphs with separate first-line and continuation widths, `docx-numbering
  --skip-slow` passed `8` and `docx-tables --skip-slow` passed `35` after a serial rerun. Private DOCX run
  `20260531-190121` stayed at `16/16` pages, zero dimension mismatches, MAE `15.849350`, changed16
  `0.141574`; `DOCX_NUMBERING_INDENT` remains open for exact tab-stop ownership, bullet fonts, style
  inheritance, and mixed-run segmentation.
- DOCX table-header diagnostic validation:
  after removing the stale unsupported diagnostic for repeating table header rows, public diagnostics now
  assert that `w:tblHeader` is not reported as unsupported because the reader preserves the token and the
  layout stage repeats contiguous header rows after page breaks. Exact table pagination remains covered by
  the broader table recovery items. Private DOCX run `20260531-222823` stayed pixel-neutral (`16/16` pages,
  zero dimension mismatches, `14.818900` MAE, `0.134293` changed16) while dropping
  `DOCX_UNSUPPORTED_TABLE_HEADER_ROW` from the diagnostics.
- DOCX table-style paragraph/run cascade validation:
  after merging table-style `w:pPr`/`w:rPr` into the DOCX style cascade and gating fallback conditional
  regions through `w:tblLook`, the targeted public reader test passed, `docx-tables --skip-slow` passed `40`,
  `docx-text --skip-slow` passed `15`, and `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal`
  passed. Private DOCX run `20260531-220413` matched the Office reference page count (`16/16`) with zero
  dimension mismatches; aggregate MAE was `14.825006`, changed16 `0.134342`. This is a pagination/structure
  improvement but not a pure raster win, so table-style precedence and diagnostics remain open.
- DOCX table-style cell vertical-alignment validation:
  after merging whole-style and conditional table-style `w:tcPr/w:vAlign` into resolved cells,
  `docx-tables --skip-slow` passed `41`. Private DOCX run `20260531-221021` stayed at `16/16` pages with zero
  dimension mismatches and unchanged aggregate metrics (`MAE=14.825006`, changed16 `0.134342`), confirming the
  slice is a structural cascade fix rather than a dominant private raster driver.
- DOCX table-style basedOn validation:
  after resolving table-style `w:basedOn` chains, `docx-tables --skip-slow` passed `42`,
  `docx-text --skip-slow` passed `16`, `docx-numbering --skip-slow` passed `9`, and
  `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed. Private DOCX run
  `20260531-223326` stayed at `16/16` pages with zero dimension mismatches and improved aggregate metrics
  (`MAE=14.791805`, changed16 `0.133893`). Remaining diagnostics are `DOCX_NUMBERING_INDENT`,
  `DOCX_STYLE_PARAGRAPH_KEEP_RULE`, `DOCX_STYLE_TABLE_STYLE`, and `DOCX_UNSUPPORTED_TABLE_STYLE`.
- DOCX table-style table-property validation:
  after applying style-level `tblLayout`, `tblW`, `tblInd`, and `tblCellSpacing`, `docx-tables --skip-slow`
  passed `43`, `docx-page --skip-slow` passed `13`, `docx-text --skip-slow` passed `16`, and
  `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed. Private DOCX run `20260531-223803`
  stayed metric-neutral at `16/16` pages, zero dimension mismatches, `MAE=14.791805`, changed16 `0.133893`.
- DOCX table-style band-size validation:
  after applying `tblStyleRowBandSize`/`tblStyleColBandSize` to fallback conditional-region inference,
  `docx-tables --skip-slow` passed `43`, `docx-text --skip-slow` passed `16`, and
  `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed. Private DOCX run `20260531-224113`
  stayed neutral at `16/16` pages, zero dimension mismatches, `MAE=14.791805`, changed16 `0.133893`.
- DOCX keep-with-indented-table validation:
  after making kept-block estimation subtract table indent before resolving preferred width and first-row
  wrapping, `docx-page --skip-slow` passed `14`, `docx-tables --skip-slow` passed `43`, and
  `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed. Private DOCX run `20260531-224646`
  stayed neutral at `16/16` pages, zero dimension mismatches, `MAE=14.791805`, changed16 `0.133893`.
- DOCX all-caps run-style validation:
  after applying `w:rPr/w:caps` through resolved DOCX run styles, `docx-tables --skip-slow` passed `43`,
  `docx-text --skip-slow` passed `16`, `docx-numbering --skip-slow` passed `9`, and
  `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed. Private DOCX run `20260531-225034`
  stayed neutral at `16/16` pages, zero dimension mismatches, `MAE=14.791805`, changed16 `0.133893`.
- DOCX numbering-tab validation:
  after preserving authored numbering-level `w:tab w:val="num"` positions and using them for marker layout,
  `docx-numbering --skip-slow` passed `10`, `docx-tables --skip-slow` passed `45`, and
  `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed. Private DOCX run
  `20260531-231855` stayed neutral at `16/16` pages, zero dimension mismatches, `MAE=13.648284`, changed16
  `0.125542`; remaining private diagnostics are `DOCX_NUMBERING_INDENT`, `DOCX_STYLE_PARAGRAPH_KEEP_RULE`,
  `DOCX_STYLE_TABLE_STYLE`, and `DOCX_UNSUPPORTED_TABLE_STYLE`.
- DOCX logical table-border validation:
  after rendering cell-level `w:start`/`w:end` as default left-to-right physical edges and resolving
  table-level logical outer borders, `docx-tables --skip-slow` passed `47`, `docx-numbering --skip-slow`
  passed `10`, and `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed. Private DOCX run
  `20260531-232213` stayed neutral at `16/16` pages, zero dimension mismatches, `MAE=13.648284`, changed16
  `0.125542`.
- DOCX typed header/footer validation:
  after preserving default/first/even header/footer reference types and selecting static parts per page,
  `docx-page --skip-slow` passed `15`, `docx-tables --skip-slow` passed `47`, and
  `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed. Private DOCX run
  `20260531-232937` stayed neutral at `16/16` pages, zero dimension mismatches, `MAE=13.648284`, changed16
  `0.125542`.
- DOCX header/footer distance validation:
  after preserving `w:pgMar/@w:header` and `@w:footer` and using them for static header/footer baselines,
  `docx-page --skip-slow` passed `16`, `docx-tables --skip-slow` passed `47`, `docx-numbering --skip-slow`
  passed `10`, and `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed. Private DOCX run
  `20260531-233331` stayed neutral at `16/16` pages, zero dimension mismatches, `MAE=13.648284`, changed16
  `0.125542`.
- DOCX style keep-rule diagnostic validation:
  after narrowing diagnostics so supported style-level `keepNext`/`keepLines` no longer warn, `docx-text
  --skip-slow` passed `17`, `docx-page --skip-slow` passed `16`, `docx-numbering --skip-slow` passed `10`,
  `docx-tables --skip-slow` passed `47`, and `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal`
  passed. Private DOCX run `20260531-233722` stayed neutral at `16/16` pages, zero dimension mismatches,
  `MAE=13.648284`, changed16 `0.125542`, and now reports only `DOCX_NUMBERING_INDENT`,
  `DOCX_STYLE_TABLE_STYLE`, and `DOCX_UNSUPPORTED_TABLE_STYLE`.
- DOCX numbering diagnostic validation:
  after narrowing `DOCX_NUMBERING_INDENT` to unsupported `right`/`firstLine` indent forms and adding
  `DOCX_NUMBERING_MARKER_FONT` for bullet levels with authored marker fonts, `docx-numbering --skip-slow`
  passed `11`, `docx-text --skip-slow` passed `17`, `docx-tables --skip-slow` passed `47`, and
  `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed. Private DOCX run `20260531-234039`
  stayed neutral at `16/16` pages, zero dimension mismatches, `MAE=13.648284`, changed16 `0.125542`, and now
  reports `DOCX_NUMBERING_MARKER_FONT`, `DOCX_STYLE_TABLE_STYLE`, and `DOCX_UNSUPPORTED_TABLE_STYLE`.
- DOCX table-style use diagnostic validation:
  after removing the duplicate `DOCX_UNSUPPORTED_TABLE_STYLE` warning for merely using `w:tblStyle`, public
  diagnostics still keep `DOCX_STYLE_TABLE_STYLE` on table-style definitions until exact Word precedence and
  unsupported complex-script atoms are narrowed further. `docx-tables --skip-slow` passed `47`,
  `docx-numbering --skip-slow` passed `11`, `docx-text --skip-slow` passed `17`, and
  `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed. Private DOCX run `20260531-234350`
  stayed neutral at `16/16` pages, zero dimension mismatches, `MAE=13.648284`, changed16 `0.125542`, and now
  reports only `DOCX_NUMBERING_MARKER_FONT` and `DOCX_STYLE_TABLE_STYLE`.
- DOCX table-style complex-script diagnostic validation:
  after replacing the broad table-style-definition warning with `DOCX_STYLE_TABLE_COMPLEX_SCRIPT_RUN` for
  unsupported table-style `w:bCs`/`w:iCs`, supported table-style atoms no longer warn just because they are
  part of a table style. Public validation passed `docx-tables --skip-slow` (`48`), `docx-numbering
  --skip-slow` (`11`), `docx-text --skip-slow` (`17`, after a serial rerun because the first parallel build
  hit a transient compiler output lock), and `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal`.
  Private DOCX run `20260531-234639` stayed neutral at `16/16` pages, zero dimension mismatches,
  `MAE=13.648284`, changed16 `0.125542`, and now reports `DOCX_NUMBERING_MARKER_FONT` plus the specific
  `DOCX_STYLE_TABLE_COMPLEX_SCRIPT_RUN`.
- DOCX run-level PDF font-resource validation:
  after grouping `DocxFontPlan` runs into PDF font resources by resolved font file and collection face index,
  and removing the stale broad marker-font diagnostic, public validation passed `docx-core --skip-slow` (`16`),
  `docx-numbering --skip-slow` (`11`), `docx-text --skip-slow` (`17`), `docx-tables --skip-slow` (`48`), and
  `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal`. A parallel test attempt hit the known
  transient `VBCSCompiler` output lock, then serial reruns passed. Private DOCX run `20260531-235343`
  improved to `16/16` pages, zero dimension mismatches, `MAE=12.509698`, changed16 `0.112673`, and now reports
  only `DOCX_STYLE_TABLE_COMPLEX_SCRIPT_RUN`.
- DOCX complex-script table-style validation:
  after modeling table-style `w:bCs`/`w:iCs` as script-slot run properties and routing complex-script spans
  through `w:cs`/Bidi theme font candidates, public validation passed `docx-tables --skip-slow` (`49`),
  `docx-core --skip-slow` (`16`), `docx-text --skip-slow` (`17`), `docx-numbering --skip-slow` (`11`), and
  `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal`. Private DOCX run `20260601-000124` stayed at
  `16/16` pages, zero dimension mismatches, `MAE=12.509698`, changed16 `0.112673`, and reports no diagnostics.
- DOCX explicit-font table fixture:
  added `docx-ladder-02-table-explicit-font` and ran `pwsh tools/CheckVisualCase.ps1 -Case
  visual-cases/cases/docx-ladder-02-table-explicit-font/case.json` as run `20260601-001941`. The comparison is
  dimension-stable (`MAE=0.235933`, changed16 `0.003113`) but PDF text inspection shows the structural gap:
  the Office reference has table-cell text operations and the candidate has none because table-cell text is
  still gated on the document fallback font resource.
- DOCX fallback-free table-text and no-margin placement validation:
  removing the stale fallback-resource gate makes table-cell text render when runs resolve to their own font
  resources, and table-cell text/images are clipped to the cell rectangle. The table-cell placement path no
  longer invents `4pt` horizontal padding when `tcMar`/`tblCellMar` are absent; public Office inspection of
  `docx-ladder-02-table-explicit-font` shows reference text at the cell edge and first baseline. The accepted
  branch keeps the existing authored-margin first-baseline inset behind a named compatibility constant because
  applying the no-margin font-size baseline to all cells worsened private table pages; add an Office-backed
  explicit-margin ladder before replacing that remaining constant. Validation passed `docx-tables --skip-slow`
  (`55`) and `docx-core --skip-slow` (`16`). Public visual `docx-ladder-02-table-explicit-font` run
  `20260601-012032` improved to `MAE=0.229378`, changed16 `0.003197`, with candidate text at `X=72/252`,
  `Y=708` versus Office `X=72.504/252.53`, `Y=707.74`. Public visual
  `docx-ladder-02-table-cell-margins` run `20260601-012305` is dimension-stable at `MAE=0.701522`,
  changed16 `0.016586`, and records the authored-margin baseline gap for the next row-height/text-spacing
  slice. The accepted row-height follow-up expanded non-`exact` authored row heights by the row's maximum
  authored top cell margin; public visual `docx-ladder-02-table-cell-margins` run `20260601-013230` improved to
  `MAE=0.545123`, changed16 `0.007648`, while `docx-tables --skip-slow` passed `56`. Private DOCX run
  `20260601-013928` matched the previous clipped-run metric (`16/16` pages, zero dimension mismatches, no
  diagnostics, `MAE=13.589200`, changed16 `0.124354`). Rejected trial: using measured row heights for the
  table pre-pagination check changed the private candidate to `17` pages (`20260601-013155`), so page-breaking
  still needs a separate Word-backed table pagination fixture. Public `docx-ladder-03-table-pagination-margins`
  run `20260601-013533` now covers that edge: Office and candidate both keep the first row on page 1 and move
  only the following row to page 2, with no diagnostics (page 1 `MAE=0.927605`, page 2 `MAE=0.253998`). The
  table-style `w:rPr` precedence probe remains open: Office shows table-style run properties should sit below
  paragraph/character styles, but the broader cascade change was not kept because it changed private
  pagination/page count. A narrower run-only cascade trial was also rejected after private run
  `20260601-013756` changed the candidate to `15` pages (`MAE=13.645890`, changed16 `0.124541`), despite
  passing public `docx-tables` and `docx-text`.
- DOCX paragraph autospacing validation:
  added `docx-ladder-02-paragraph-autospacing` and changed `beforeAutospacing`/`afterAutospacing` to resolve
  from paragraph line height after fixing spacing cascade ownership per side. `docx-text --skip-slow` passed
  `17`, `docx-page --skip-slow` passed `17`, `docx-tables --skip-slow` passed `58`, `docx-numbering
  --skip-slow` passed `11`, and `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed. Public
  visual run `20260601-021353` is dimension-stable at `MAE=0.426487`, changed16 `0.003888` versus baseline
  `0.432740`/`0.003942`. Private DOCX run `20260601-021406` stayed neutral at `16/16` pages, zero dimension
  mismatches, no diagnostics, `MAE=12.716572`, changed16 `0.118448`.
- DOCX Office-like table-border emission validation:
  after changing DOCX table borders from stroked lines to filled rectangle strips, `docx-tables --skip-slow`
  passed `58`, `docx-page --skip-slow` passed `17`, `docx-core --skip-slow` passed `16`, and
  `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed. Public
  `docx-ladder-03-table-pagination-margins` run `20260601-023556` improved to page 1 `MAE=0.888769`,
  changed16 `0.011461`, and page 2 `MAE=0.210261`, changed16 `0.003116`; candidate PDF inspection shows the
  table graphics as fill operations rather than strokes. Private DOCX run `20260601-023645` stayed at `16/16`
  pages, zero dimension mismatches, no diagnostics, `MAE=12.503007`, changed16 `0.116841`.
- DOCX shared vertical table-border validation:
  after emitting each shared vertical border once and honoring adjacent `nil`/`none` suppression,
  `docx-tables --skip-slow` passed `59`, `docx-page --skip-slow` passed `17`, and
  `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed. Public
  `docx-ladder-03-table-pagination-margins` run `20260601-024057` stayed unchanged from the filled-border
  baseline (page 1 `MAE=0.888769`, page 2 `MAE=0.210261`). Private DOCX run `20260601-024114` stayed at
  `16/16` pages, zero dimension mismatches, no diagnostics, `MAE=12.503007`, changed16 `0.116841`.
- DOCX shared horizontal table-border validation:
  after resolving consecutive same-page row boundaries as one horizontal strip, `docx-tables --skip-slow`
  passed `60`, `docx-page --skip-slow` passed `17`, and `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v
  minimal` passed. Public `docx-ladder-03-table-pagination-margins` run `20260601-024402` stayed unchanged
  from the filled-border baseline (page 1 `MAE=0.888769`, page 2 `MAE=0.210261`). Private DOCX run
  `20260601-024417` stayed at `16/16` pages, zero dimension mismatches, no diagnostics, and improved to
  `MAE=12.494853`, changed16 `0.116738`.
- DOCX grid-span shared horizontal border validation:
  after matching shared horizontal border segments by overlapping cell geometry, `docx-tables --skip-slow`
  passed `61`, `docx-page --skip-slow` passed `17`, and `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v
  minimal` passed. Public `docx-ladder-03-table-pagination-margins` run `20260601-024621` stayed unchanged
  from the filled-border baseline (page 1 `MAE=0.888769`, page 2 `MAE=0.210261`). Private DOCX run
  `20260601-024621` stayed at `16/16` pages, zero dimension mismatches, no diagnostics, `MAE=12.494853`,
  changed16 `0.116738`.
- DOCX collapsed-border row-advance validation:
  Office's table row boundaries in `docx-ladder-03-table-row-heights` showed non-`exact` row advances include
  the resolved horizontal collapsed-border width, while `exact` rows keep the authored height. The layout stage
  now adds the visible bottom horizontal border width, falling back to top when no bottom border exists, through
  the same shared `w:sz` eighth-point border geometry helper used by PDF border emission. Public row-height run
  `20260601-114414` improved from run `20260601-113307` (`MAE=1.469789`, changed16 `0.013068`) to
  `MAE=1.001114`, changed16 `0.010631`; inspected table text baselines now sit within roughly `0.1pt` of
  Office through the table. Collateral checks were acceptable but not free: `docx-ladder-03-table-paragraph-adjacency`
  moved from `MAE=0.778729`, changed16 `0.006298` to run `20260601-114024` at `MAE=0.814874`, changed16
  `0.006597`; `docx-ladder-03-table-pagination-margins` run `20260601-114004` moved page 1 from `0.888769` to
  `0.872023` and page 2 from `0.210261` to `0.232394`. Validation passed `docx-tables --skip-slow` (`69`),
  `docx-core --skip-slow` (`21`), `docx-text --skip-slow` (`36`), and full solution build. Private DOCX run
  `20260601-114151` stayed at `16/16` pages, zero dimension mismatches, no diagnostics, and improved from
  `MAE=13.615949`, changed16 `0.125050` to `MAE=13.577699`, changed16 `0.124754`. Keep a follow-up open for
  border-width calibration: Office emits the same nominal `w:sz=4` horizontal strips at about `0.48pt` while
  the candidate emits `0.50pt`, so the remaining row-grid miss belongs to shared border-width normalization,
  not another row-origin offset.
  2026-06-01 follow-up: the shared DOCX table-border geometry helper now applies Word's PDF border-width
  scale of `0.96 * (w:sz / 8pt)`, observed on public Office PDFs for both `w:sz=4` (`0.48pt`) and `w:sz=8`
  (`0.96pt`). This keeps layout row advances and PDF border strips on the same structural width. Public
  validation passed `docx-tables --skip-slow` (`69`), `docx-core --skip-slow` (`21`), and full solution build
  (one parallel build had transient CLI copy-lock warnings but succeeded). Visuals improved further:
  `docx-ladder-03-table-row-heights` run `20260601-114904` reached `MAE=0.867597`, changed16 `0.009963`;
  `docx-ladder-02-table-cell-margins` run `20260601-114904` reached `MAE=0.535590`, changed16 `0.005766`;
  `docx-tables` run `20260601-114904` reached `MAE=0.539255`, changed16 `0.004836`; adjacency and pagination
  stayed at the post-row-advance metrics. Private DOCX run `20260601-114952` stayed at `16/16` pages, zero
  dimension mismatches, no diagnostics, and improved again to `MAE=13.562167`, changed16 `0.124579`.
- DOCX typographic auto-line-height validation:
  Office evidence from `docx-ladder-02-paragraph-logical-indents` showed the logical-indents geometry itself
  was correct, but wrapped Aptos 11pt lines were advancing by about `13.44pt` while the candidate advanced by
  about `14.13pt` from the resolved font's Windows bounding box. Cross-checking `docx-ladder-02-paragraph-indents`
  showed Arial `w:line=276` auto spacing was already aligned, so this was not an indent bug and not a
  font-specific exception. DOCX automatic line height now uses the resolved OpenType typographic line box
  (`sTypoAscender - sTypoDescender + sTypoLineGap`) with the existing Word minimum floor instead of taking the
  maximum of typographic and Windows boxes. This keeps the rule structural and avoids hard-coded font names.
  Validation passed `docx-text --skip-slow` (`36`), `docx-core --skip-slow` (`21`), `docx-tables --skip-slow`
  (`69`), and full solution build; two parallel runs hit transient compiler/Defender output locks and passed
  on serial rerun. Public visuals improved or stayed neutral: `docx-ladder-02-paragraph-logical-indents`
  improved from run `20260601-115121` (`MAE=0.690059`, changed16 `0.007701`) to run `20260601-120526`
  (`MAE=0.475814`, changed16 `0.006385`); `docx-ladder-02-character-spacing` improved from run
  `20260601-115112` (`MAE=0.876254`, changed16 `0.008339`) to run `20260601-120516` (`MAE=0.770167`,
  changed16 `0.007879`); `docx-ladder-02-paragraph-indents` stayed neutral at `MAE=0.162103`, changed16
  `0.004023`. Full `docx-layout` run `20260601-120449` passed all `21` cases. Private DOCX run
  `20260601-120450` stayed neutral at `16/16` pages, zero dimension mismatches, no diagnostics,
  `MAE=13.562167`, changed16 `0.124579`. Keep the follow-up open for baseline anchoring and paragraph/table
  adjacency: the remaining top DOCX misses are still table pagination/row heights/paragraph adjacency rather
  than logical indentation.
- DOCX table-to-paragraph adjacency validation:
  Office PDF inspection of `docx-ladder-03-table-paragraph-adjacency` showed table borders and cell baselines
  were already close, but body paragraphs following the table were about `5.4pt` too low. The candidate table
  bottom was only about `0.6pt` above Office, so the gap came from a renderer-owned fixed `6pt` table trailing
  advance rather than OOXML paragraph spacing or row height. The layout stage no longer injects a synthetic
  post-table gap; following paragraphs now start from the table bottom and use their own before/after spacing.
  Public validation improved the adjacency case from run `20260601-120538` (`MAE=0.814874`, changed16
  `0.006597`, SSIM `0.517943`) to run `20260601-121054` (`MAE=0.543831`, changed16 `0.004823`, SSIM
  `0.683283`). Related table visuals stayed neutral: `docx-ladder-03-table-row-heights` remained at
  `MAE=0.867597`, changed16 `0.009963`; `docx-ladder-03-table-pagination-margins` remained at `MAE=0.552209`,
  changed16 `0.006630`. Validation passed `docx-tables --skip-slow` (`70`), `docx-text --skip-slow` (`36`),
  `docx-core --skip-slow` (`21`), `docx-page --skip-slow` (`24`), and full solution build. Private DOCX run
  `20260601-121146` stayed structurally valid at `16/16` pages, zero dimension mismatches, and no diagnostics,
  but worsened to `MAE=13.791252`, changed16 `0.126278`; keep a follow-up open to classify table-adjacent
  private pages by table/paragraph sequence before adding any replacement spacing rule.
- DOCX bordered-cell content inset validation:
  Office PDF inspection of `docx-ladder-03-table-row-heights` showed bordered table text starts about one half
  of the visible collapsed border width farther inside each cell than the candidate. This was not a font or
  row-height issue: row baselines were already within roughly `0.1pt`, while the table text X positions were
  consistently about `0.24pt` left with `w:sz=4` borders rendered as `0.48pt` strips. The DOCX cell content
  box now adds half of the resolved visible left/right border width to the horizontal content insets used by
  measurement, text layout, and inline images. Unbordered cells keep the Word default `5.4pt` horizontal
  padding; the extra inset is derived from actual border structure. Public validation improved
  `docx-ladder-03-table-row-heights` from run `20260601-121040` (`MAE=0.867597`, changed16 `0.009963`, SSIM
  `0.787208`) to run `20260601-121747` (`MAE=0.631193`, changed16 `0.006442`, SSIM `0.823925`), improved
  `docx-tables` from `MAE=0.539255`, changed16 `0.004836` to `MAE=0.507449`, changed16 `0.004190`, and
  slightly improved table/paragraph adjacency. The explicit-margin and explicit-font ladders moved slightly
  worse (`docx-ladder-02-table-cell-margins` from `MAE=0.535590` to `0.546795`; explicit-font from
  `0.198498` to `0.199818`), so keep a follow-up open for Office's exact interaction between authored cell
  margins and collapsed border inside edges. Validation passed `docx-tables --skip-slow` (`71`),
  `docx-text --skip-slow` (`36`), `docx-core --skip-slow` (`21`), and full solution build. Private DOCX run
  `20260601-121847` stayed at `16/16` pages, zero dimension mismatches, no diagnostics, `MAE=13.791252`,
  changed16 `0.126278`.
- DOCX unsupported table-border-style diagnostic validation:
  after emitting `DOCX_TABLE_BORDER_STYLE` only for visible non-`single`/`nil`/`none` table and cell border
  styles in document/style parts, `docx-tables --skip-slow` passed `62`, `docx-core --skip-slow` passed `16`
  after a serial rerun of a transient output lock, `docx-page --skip-slow` passed `17`, and
  `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` passed. Private DOCX run `20260601-025057`
  stayed at `16/16` pages, zero dimension mismatches, no diagnostics, `MAE=12.494853`, changed16 `0.116738`.
- DOCX run character-spacing validation:
  after promoting `w:rPr/w:spacing` as signed twip run geometry through parsing, style application, text
  measurement, segment placement, header/footer placement, and spacing-only `TJ` PDF emission, public
  validation passed `docx-core --skip-slow` (`19`), `docx-text --skip-slow` (`18`), `docx-tables --skip-slow`
  (`62`), `docx-page --skip-slow` (`17`), `docx-numbering --skip-slow` (`11`), and
  `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal`. Two early parallel validation attempts hit
  transient compiler/Defender output locks and passed on serial rerun. Private DOCX run `20260601-025951`
  stayed at `16/16` pages, zero dimension mismatches, no diagnostics, `MAE=12.494853`, changed16 `0.116738`.
- DOCX character-spacing visual ladder:
  added Office-authored `docx-ladder-02-character-spacing` and manifest. Public run `20260601-030351` compared
  `1/1` pages with matching dimensions, `MAE=1.960983`, changed16 `0.016165`, foreground histogram
  correlation `1.0`. Office PDF text inspection shows `28` positioned text operations with spacing buckets
  near authored `+2pt` and `-1pt`, plus residual positioned adjustments on nominally zero-spacing runs. The
  candidate renders the default header/footer references from `header2.xml`/`footer2.xml` with the authored
  spacing, but its text-operation decomposition remains less Office-like than the reference.
  2026-06-01 follow-up: DOCX text emission now applies synthetic bold and italic only when the resolved font
  face does not already provide the requested bold/italic attribute. This keeps structural font resolution as
  the source of truth instead of double-emitting already-bold text or shearing already-italic glyphs. Public
  `docx-ladder-02-character-spacing` improved from run `20260601-103426` (`MAE=1.714594`, changed16
  `0.014529`) to run `20260601-103906` (`MAE=1.685749`, changed16 `0.013842`). Private DOCX run
  `20260601-103906` stayed stable at `16/16` pages, zero dimension mismatches, no diagnostics, and improved
  from `MAE=13.388935`, changed16 `0.124264` to `MAE=13.286569`, changed16 `0.123783`.
  2026-06-01 follow-up: the DOCX layout stage now justifies wrapped non-final `w:jc="both"` lines by
  expanding ordinary inter-word spaces structurally in layout instead of leaving trailing-space width in the
  line box. Final paragraph lines remain natural width, and list-label first lines are kept out of the new
  path until public Office evidence covers that interaction. Validation passed `docx-core --skip-slow` (`23`),
  `docx-text --skip-slow` (`36`), `docx-tables --skip-slow` (`71`), and full solution build. Private DOCX run
  `20260601-124141` stayed unchanged at `16/16` pages, zero dimension mismatches, no diagnostics,
  `MAE=13.791190`, changed16 `0.126277`; inspection found no direct body `w:jc` tokens in that private case,
  so justification is closed as a generic missing DOCX feature but is not the current private-driver gap.
  2026-06-01 follow-up: the Office PDF font-size emission grid is now shared between PPTX and DOCX through
  `OfficePdfTextEmissionProfile` instead of being a PPTX-private duplicate. DOCX `/Tf` sizes and glyph
  positioning arrays now use the same 600-DPI Office grid already observed on PPTX. This structurally moves
  candidate DOCX font-size buckets toward Word's PDF output: in private run `20260601-124523`, candidate
  buckets became `9.96:847`, `9:323`, `14.04:24`, `18:9`, `11.04:8`, `12:1`, `26.04:1`, while the reference
  uses the same size grid with higher operation counts. Raster impact was mixed: aggregate private MAE moved
  from `13.791190` to `13.838763`, but several worst pages improved (`9`, `10`, `11`, and `8`). Public
  `docx-ladder-02-character-spacing` run `20260601-124606` and `docx-numbering` run `20260601-124606` also
  moved slightly worse in raster while gaining the shared `/Tf` structure. Keep the shared grid; the remaining
  driver is text-state and operation decomposition, not a reason to restore pre-grid font sizes.
- [ ] DOCX Office PDF text-state decomposition:
  private DOCX inspection after the shared font grid still shows the candidate emits only `Tc=0` (`1213`
  operations), while Word emits widespread nonzero character-spacing buckets (`0`, `-0.003`, `-0.010`,
  `-0.050`, `0.110`, `0.031`, `0.051`, `0.058`, and related small values). Public
  `docx-ladder-02-character-spacing` shows authored run spacing should remain encoded through positioned
  glyph adjustments rather than blindly promoted to `Tc`, because Word uses `Tc=0` there; the mismatch is a
  separate Office PDF text-state/operation-splitting behavior. Next work must survey existing public DOCX
  reference PDFs and, if needed, add small Office-authored probes for ordinary paragraphs, table cells,
  headers/footers, multi-run paragraphs, and style-driven font sizes before changing emission. Do not
  hard-code fonts, private bucket constants, or private page/document conditions.
  2026-06-01 progress: public PDF inspection split the problem into two parts. Word emits separate text
  operations for terminal line spaces and for list-label suffix spaces; `ooxpdf` now mirrors that decomposition
  without changing layout width or underline geometry. Public runs `docx-ladder-01-plain-paragraph`
  (`20260601-130012`), `docx-numbering` (`20260601-130021`), and `docx-tables` (`20260601-130029`) now match
  Word's text-operation counts exactly: plain `2/2`, numbering `14/14`, tables `21/21`. The remaining open gap
  is specifically context-derived PDF character spacing: numbering reference uses `Tc=0.048` on list/body
  operations while candidate remains `Tc=0`; table reference uses `Tc=0.0509` for header cells and `Tc=-0.0182`
  for most body cell text while candidate remains `Tc=0`. Public `docx-ladder-02-character-spacing` still shows
  authored `w:rPr/w:spacing` is not the answer because Word keeps `Tc=0` there and encodes spacing through
  positioned glyph arrays. Next work should derive the Word `Tc` rule from public structure such as line/cell/list
  context, font-size grid, and PDF text-state reuse, not from fonts, private bucket constants, or private cases.
  Private DOCX acceptance run `20260601-130157` stayed stable at `16/16` pages, zero dimension mismatches, no
  diagnostics, `MAE=13.838763`, changed16 `0.126851`. Aggregate private text-operation inspection improved
  candidate decomposition from the pre-change `1213` operations to `2184` operations, closer to Word's `2388`,
  but all candidate operations still use `Tc=0` while Word uses many small positive and negative buckets. This
  reinforces that the next DOCX fidelity step is the generic Word `Tc` rule, not more operation splitting.
  2026-06-01 negative result: an emission-only font-grid width-residual `Tc` experiment improved public
  `docx-numbering` raster metrics (`MAE` about `0.01927` -> `0.00814`) and slightly improved
  `docx-ladder-02-character-spacing`, but it introduced nonzero candidate `Tc` in a public case where Word's
  reference keeps every operation at `Tc=0`. It also slightly worsened `docx-tables`. Do not use font-size-grid
  residual compensation as the generic DOCX `Tc` rule; the rule must be derived from Word's actual text-state
  choices, not from raster improvement alone.
  2026-06-01 follow-up: static header/footer rendering now uses the same terminal-space operation split as body
  paragraphs, preserving layout geometry while aligning Word's PDF text-operation decomposition. Public
  `docx-ladder-02-character-spacing` run `20260601-131532` now matches Word exactly on text operations and text
  state for this case: reference `28` `TJ` operations, candidate `28`; reference nonzero `Tc=0`, candidate
  nonzero `Tc=0`; `/Tf` size buckets also match (`9.96`, `12`, `14.04`, `21.96`). Raster metrics stayed
  `MAE=0.747573`, changed16 `0.007952`, confirming this was a structural PDF alignment step rather than a
  raster shortcut. The new `tools/SummarizeDocxTextState.ps1` records these private-safe aggregate checks and
  should be used before any further DOCX `Tc` work.
  2026-06-01 correction: the first numbered-text slice exposed a validation-tool gap. `PdfInspect` text
  extraction tracked `cm` through `q`/`Q`, but did not save/restore font and `Tc` text-state. Restacking those
  fields changed the public evidence: Word's `docx-numbering` reference uses `Tc=0.048` only on the three
  numbered-label operations (`1.`, `2.`, `3.`), while list suffix spaces and body text stay at `Tc=0`.
  `ooxpdf` now matches that narrower structure: public run `20260601-133234` has reference/candidate `14/14`
  `TJ` operations, `Tc=0` on `11`, `Tc=0.048` on `3`, matching `/Tf` buckets (`12`, `21.96`), and unchanged
  raster metrics (`MAE=0.019271`, changed16 `0.000744`, `SSIM=0.996098`). The label bucket is emitted as real
  text-state spacing, not compensated away; body and suffix text remain `Tc=0`. Public
  `docx-ladder-02-character-spacing` run `20260601-133112` remains the guard against overgeneralization:
  reference/candidate both emit `28/28` `TJ` operations, all with `Tc=0`, unchanged at `MAE=0.747573` and
  changed16 `0.007952`. Restacked public `docx-tables` run `20260601-133059` also narrows the table target:
  Word has only `6` nonzero `Tc` operations, not the previous erroneous `17` (`0.0509` on two header value
  cells, `-0.0182` on four body value cells). Keep table/body `Tc` open and add public probes before changing
  rendering; the next rule must come from table structure and graphics/text-state lifetime, not from text
  length, column number, per-cell constants, or the obsolete unstacked summary.
  2026-06-02 bullet-label correction: public `docx-ladder-03-compact-bullet-alt-line115` showed the decimal
  numbering `Tc` branch must not apply to bullet-format list labels. Office and the corrected candidate both
  emit `170/170` text operations with `NonzeroTcCount=0` and a single `Tc=0` bucket in run `20260602-031750`;
  raster metrics stayed unchanged (`MAE=7.724227`, changed16 `0.076554`), as expected for a PDF text-state
  structural alignment step. Keep decimal/numbered label `Tc` guarded by the existing public numbering case,
  and keep the broader short-run/table `Tc` item open.
  2026-06-01 table-text-state probe: added Office-authored public fixture
  `docx-ladder-03-table-text-state` to vary table header/body text classes without using private content. Run
  `20260601-133647` matches operation count (`35/35`) but confirms the remaining structural table text-state
  gap: Word emits `Tc=-0.0182` on six 11.04pt digit-only body value operations while the candidate emits
  `Tc=0` everywhere. The existing `docx-tables` run still shows the companion positive header/value bucket
  (`Tc=0.0509` on two alphanumeric header value operations). `tools/SummarizeDocxTextState.ps1` now includes
  private-safe `TextClassByTc` and `NonzeroTcByTextClass` buckets, so future private and public inspections can
  distinguish digits/letters/alphanumeric/empty operations without exposing decoded text. Keep rendering
  unchanged until another probe separates text class from table role and header/body state; do not collapse
  this into a digit-only or column-number shortcut.
  2026-06-01 context probe: added `docx-ladder-03-text-state-context` to put the same short digit,
  alphanumeric, and letter tokens in normal paragraphs, table header cells, and table body cells. Run
  `20260601-133921` shows Word's short-token `Tc` behavior crosses paragraph/table boundaries and is not a
  pure digit rule: paragraph `42` uses `Tc=-0.0182`, paragraph `Q1` uses `Tc=0.0509`, table `42`/`128` use
  `Tc=-0.0182`, table `Q1` uses `Tc=0.0509`, while short alphabetic/alphanumeric pairs can use
  `Tc=-0.0509` or `-0.0437` depending on the glyph pair and row context. Longer text generally stays at
  `Tc=0` with residuals in positioning arrays. This suggests the long-term target is a generic Office text
  emission decomposition that can promote uniform short-run advance residuals into `Tc` where Word does, not a
  table-specific or text-class-specific special case. Candidate still has a decomposition gap here (`35`
  reference operations vs `33` candidate operations, all candidate `Tc=0`).
  2026-06-01 size-matrix probe: added Office-authored public fixture
  `docx-ladder-03-text-state-size-matrix` to vary the same short tokens across `8`, `9`, `10`, `11`, `12`, and
  `14` pt Arial table rows. Run `20260601-134305` shows the Word rule is not a single font-size-grid residual:
  reference emits `19` nonzero `Tc` operations across all tested sizes, while candidate emits `63` operations
  all at `Tc=0`. Public decoded reference examples are `42`: `-0.0302`, `0.036`, `-0.0178`, `-0.0182`,
  `0.048`, `-0.00624`; `Q1`: `-0.0151`, `-0.042`, `0.0511`, `0.0509`, `0.024`, `-0.00312`; and `AB`:
  `0.0373`, `-0.003`, `-0.0433`, `-0.0437`, `0.036`, `-0.00468`. For these nonzero operations Word uses
  one text chunk, zero `TJ` numeric adjustments, and `NetAverageCharacterSpacing` equal to `Tc`; longer strings
  and separator fragments still use `Tc=0` and may carry residuals through positioned glyph arrays. The
  durable implementation target is therefore an Office PDF text-emission planner that can decompose a measured
  run into a uniform text-state component plus residual `TJ` adjustments after layout, with public evidence for
  when Word chooses that decomposition. Do not encode token strings, font names, table roles, or the observed
  buckets as renderer conditions. `tools/SummarizeDocxTextState.ps1` now records `TextChunkCountByTc`,
  `AdjustmentCountByTc`, and `AverageAdjustmentByTc` so this split is visible in future public/private-safe
  summaries.
  2026-06-01 PDF chunking progress: the shared DOCX glyph-positioning encoder now coalesces consecutive glyphs
  inside one hex string until a real numeric `TJ` adjustment is needed, instead of emitting one PDF text chunk
  per glyph. This is a structural alignment step, not a raster tuning step: `docx-ladder-03-text-state-size-matrix`
  rerun `20260601-134643` kept raster metrics unchanged (`MAE=1.559806`, changed16 `0.012952`) while candidate
  chunk buckets moved to `chunks=1|tc=0` on `62/63` operations. Guard runs stayed visually unchanged:
  `docx-numbering` `20260601-134708` remained at `MAE=0.019271`, changed16 `0.000744`, and
  `docx-ladder-02-character-spacing` `20260601-134708` remained at `MAE=0.747573`, changed16 `0.007952`.
  The remaining gap is the actual uniform `Tc` selection and residual split; long strings still differ in
  Office/candidate chunk buckets because Office sometimes introduces intra-run `TJ` segmentation where candidate
  has no measured residual yet.
  2026-06-02 text-state architecture follow-up: the proven decimal-numbered-list `Tc` branch now lives in the
  shared `OfficePdfTextEmissionProfile` instead of as a DOCX-layout-local numeric expression, and terminal
  line-space emissions are forced back to neutral `Tc=0` so future nonzero text-state branches cannot leak into
  synthetic trailing spaces. Public `docx-numbering` run `20260602-034714` stayed at `MAE=0.019271`,
  changed16 `0.000744`, dimensions matched. A trial table-cell digit-only branch was also evaluated against
  public `docx-ladder-03-table-text-state`: run `20260602-034629` matched Word's structural buckets
  (`35/35` operations, both with six nonzero value operations; reference `Tc=-0.0182`, candidate `Tc=-0.018`)
  with unchanged raster metrics, but this was rejected and removed because the public context and size-matrix
  probes already show Word's short-run `Tc` behavior crosses paragraph/table boundaries, includes positive
  alphanumeric buckets, varies by glyph/size, and is not a digit-only or table-role rule. The accepted state is
  the honest open gap in rerun `20260602-034924`: reference `35` operations with six `Tc=-0.0182`, candidate
  `35` operations all at `Tc=0`, raster `MAE=0.632046`, changed16 `0.006503`. Validation after removal passed
  `docx-tables --skip-slow` (`95`), `docx-core --skip-slow` (`37`), `docx-text --skip-slow` (`45`), and full
  solution build. Private DOCX acceptance rerun `20260602-035051` stayed valid at `16/16` pages, zero dimension
  mismatches, no diagnostics, aggregate `MAE=8.908377`, changed16 `0.095249`. Next implementation must be a
  generic Office text-emission planner that decomposes measured short runs into uniform `Tc` plus residual
  positioning where public evidence supports it; do not key on token strings, digit-only text,
  table/body/header role, font names, or a single observed bucket.
  2026-06-02 planner boundary progress: added `DocxTextEmissionPlanner` as the DOCX counterpart to the PPTX
  Office PDF text-state planning pass. The planner now owns Office-grid `/Tf` size, emitted PDF `Tc`,
  compensation state, and positioned-glyph character spacing for both regular runs and terminal line spaces;
  `DocxRenderer` consumes the plan for snapshots and glyph drawing instead of recomputing those values inline.
  This is behavior-neutral but removes another renderer-local text-state calculation before implementing the
  generic short-run decomposition. Bottom-up coverage checks compensated authored spacing, numbered-list `Tc`
  without positioning compensation, and neutral terminal-space `Tc`. Validation: `docx-core --skip-slow` passed
  (`40`), public `docx-numbering` run `20260602-035348` stayed at `MAE=0.019271`, public
  `docx-ladder-03-table-text-state` run `20260602-035358` stayed at the known open-gap raster
  (`MAE=0.632046`, changed16 `0.006503`), and full solution build passed.
  2026-06-02 planner operation-splitting progress: moved DOCX Office text-operation part splitting out of
  `DocxRenderer` and into `DocxTextEmissionPlanner`, next to the text-state plan. The planner now owns the
  dash-punctuation operation boundary rule plus the no-measurer fallback, so future short-run `Tc`
  decomposition can reason over operation parts before PDF drawing. This remains behavior-neutral and does not
  add any token/font/table shortcut. Bottom-up coverage checks dash punctuation splitting with measured part
  coordinates and whole-operation fallback without a measurer. Validation: `docx-core --skip-slow` passed
  (`42`), public `docx-ladder-02-long-token-wrapping` run `20260602-035625` stayed valid at `MAE=1.121729`,
  changed16 `0.017126`, public `docx-ladder-03-table-text-state` run `20260602-035636` stayed at
  `MAE=0.632046`, changed16 `0.006503`, and full solution build passed.
  2026-06-02 private-safe emission-profile progress: `DocxTextEmissionPlanner` now classifies emitted text
  operations into Unicode digit/letter/whitespace/punctuation/symbol/other counts, and
  `DocxTextEmissionSegmentSnapshot` carries that profile without exposing decoded text. `DocxInspect`
  aggregates the same profile at document and page scope in `text-emission-summary.json`, aligning candidate
  DOCX diagnostics with the PDF `TextClassByTc` evidence used for the open short-run `Tc` branch. Public
  `docx-ladder-03-table-text-state` inspect output now shows `35` candidate operations with aggregate
  `DigitCount=14`, `LetterCount=100`, `WhitespaceCount=27`, and `NonzeroPdfCharacterSpacingSegmentCount=0`.
  Validation passed `docx-core --skip-slow` (`43`), public `InspectDocx` on
  `docx-ladder-03-table-text-state`, and full solution build. The next `Tc` implementation should compare
  these private-safe candidate classes against Office PDF classes before selecting any uniform text-state
  component.
  2026-06-02 text-state summary integration: `tools/SummarizeDocxTextState.ps1` now attaches candidate planner
  text-emission summaries when a run contains `comparison/docx-inspect/text-emission-summary.json` (or
  `docx-inspect/text-emission-summary.json`). This keeps the existing PDF-only summary behavior for ordinary
  visual runs, while allowing public/private-safe joined evidence between Office PDF `Tc` buckets and candidate
  planner character profiles. Validation on public `docx-ladder-03-table-text-state` run `20260602-035636`
  with generated `comparison/docx-inspect` shows `CandidatePlannerPresent=true`, planner `SegmentCount=35`,
  planner `DigitCount=14`, candidate PDF `NonzeroTcCount=0`, and reference PDF `NonzeroTcCount=6`; full
  solution build passed.
  2026-06-02 advance-profile progress: candidate DOCX text-emission snapshots now include per-operation
  advance profiles measured from the actual embedded PDF font at Office-grid `/Tf`: mapped glyph count,
  glyph-gap count, natural PDF width, layout width, layout-to-natural residual, and residual per glyph gap.
  `DocxInspect` aggregates the same facts at document/page scope. This gives the future uniform `Tc`
  decomposition a structural input instead of token/font/table predicates. Public
  `docx-ladder-03-table-text-state` inspect output now reports `GlyphCount=141`, `GlyphGapCount=106`,
  `NaturalPdfWidth=850.266563`, `LayoutWidth=793.433105`, aggregate residual `-56.833457`, and
  `UniformResidualPerGap=-0.536165`; this aggregate is diagnostic only, not a rule. Validation passed
  `docx-core --skip-slow` (`43`), public `InspectDocx` on `docx-ladder-03-table-text-state`, and full solution
  build.
  2026-06-02 joined planner-bucket progress: `tools/SummarizeDocxTextState.ps1` now reads
  `text-emission-snapshot.json` when available and emits private-safe candidate planner segment buckets by
  text class, Office-grid `/Tf`, glyph-gap count, residual-per-gap, terminal-space status, and weighted advance
  profile aggregates. This keeps the next `Tc` rule investigation structural without exposing text. On public
  `docx-ladder-03-table-text-state`, refreshed inspect output shows Word has six `digits|tc=-0.0182`
  operations, while candidate planner sees the matching digit segments as four `digits|gaps=1` operations with
  residual-per-gap `-0.044492` and two `digits|gaps=2` operations with residual-per-gap `-0.033369`; the
  aggregate digit residual-per-gap is `-0.038931`. This is deliberately not equal to Word's `-0.0182` `Tc`,
  so the rejected raw-residual rule remains rejected. On the private DOCX acceptance run `20260602-035051`,
  Office has `81` nonzero `Tc` operations out of `2388`, spanning alphanumeric (`4`), digits (`25`), letters
  (`16`), mixed (`33`), and punctuation (`3`), while candidate still has `0` nonzero `Tc` operations across
  `2339` operations. Candidate planner aggregate residual-per-gap is `-0.020238`, but class-level residuals
  include positive and negative signs; this confirms the next rendering change must model Word's text-state
  decomposition and residual `TJ` split from public probes, not key on digit/table/font/style names or apply an
  aggregate residual as a global rule. Validation passed `docx-core --skip-slow` (`43`), refreshed public
  `InspectDocx`, public/private-safe `SummarizeDocxTextState`, and full solution build.
  2026-06-02 sequence-pair progress: the same summary tool now pairs Office PDF text operations with candidate
  planner segments by operation sequence when counts match exactly, and refuses to fuzzy-match when they do not.
  On public `docx-ladder-03-table-text-state`, `35/35` operations pair cleanly; the six Office nonzero `Tc`
  operations pair to four `digits|gaps=1|refTc=-0.0182` segments and two
  `digits|gaps=2|refTc=-0.0182` segments, with paired residual buckets
  `resGap=-0.044492|refTc=-0.0182` (`4`) and `resGap=-0.033369|refTc=-0.0182` (`2`). This gives a direct
  public structural oracle for the next planner experiment while preserving the negative finding that candidate
  residual-per-gap is not itself Word's `Tc`. On the private DOCX acceptance run, operation counts do not match
  (`2388` Office operations vs `2323` candidate planner segments), so the tool records `CountsMatched=false`
  and leaves pairing for public or fixed-decomposition cases. Validation passed `docx-core --skip-slow` (`43`),
  public/private-safe `SummarizeDocxTextState`, and full solution build.
  2026-06-02 fresh paired-evidence update: reran current public `docx-tables` and
  `docx-ladder-03-text-state-context` before using the new pair buckets, because older candidate artifacts
  still reflected a rejected residual experiment. Current `docx-tables` run `20260602-041314` stays at
  `MAE=0.455760`, changed16 `0.003840`, matches `21/21` candidate/reference PDF operations, and pairs Office
  nonzero `Tc` to `alphanumeric|gaps=1|refTc=0.0509` (`2`) and `digits|gaps=1|refTc=-0.0182` (`4`). The paired
  residual buckets are `resGap=-0.053359|refTc=0.0509` (`2`) and `resGap=-0.044492|refTc=-0.0182` (`4`), so
  positive Office `Tc` can arise from a negative candidate residual. Current
  `docx-ladder-03-text-state-context` run `20260602-041314` stays at `MAE=0.493604`, changed16 `0.004763`,
  matches `35/35` candidate/reference PDF operations, and shows Office nonzero `Tc` across
  `alphanumeric|gaps=1` (`-0.0437`, `-0.0509`, `0.0509`), `digits|gaps=1/2` (`-0.0182`), and
  `letters|gaps=1` (`-0.0437`, `-0.0509`), while current candidate emits all `Tc=0`. The same candidate
  residual bucket `resGap=-0.053359` pairs to both `refTc=-0.0437` and `refTc=0.0509`, so even paired
  residual-per-gap is insufficient. The next bottom-up step should expose a private-safe glyph-advance
  signature or public-only glyph-pair oracle for short runs before any renderer `Tc` selection is changed.
  2026-06-02 glyph-signature follow-up: candidate DOCX text-emission snapshots now expose a private-safe
  glyph-advance signature per segment: mapped glyph count, glyph-pair count, summed advance units, summed
  kerning units, and a fixed-width hash over glyph/advance/kerning structure. `SummarizeDocxTextState.ps1`
  carries those signatures into planner buckets and sequence-paired Office `Tc` buckets. On current public
  `docx-tables` run `20260602-041314`, the six nonzero Office `Tc` pairs split into six structural signatures:
  two alphanumeric signatures map to `refTc=0.0509`, and four digit signatures map to `refTc=-0.0182`. On
  current public `docx-ladder-03-text-state-context` run `20260602-041314`, repeated signatures stay stable in
  the public probe (`glyphSig=99A58E54284E5D99` maps three times to `refTc=0.0509`, and
  `glyphSig=4579009E93D630E7` maps three times to `refTc=-0.0182`), while the remaining nonzero pairs separate
  by signature into `-0.0437`, `-0.0509`, and `-0.0182` buckets. This is diagnostic evidence only, not a
  renderer rule: the next renderer change still needs an Office-like decomposition model, but it should use
  glyph-advance structure rather than text class, font name, table context, or residual-per-gap alone.
  2026-06-02 pair-signature follow-up: candidate DOCX text-emission snapshots now expose private-safe
  glyph-pair advance structure in addition to whole-operation glyph signatures: pair advance sum, min/max pair
  advance, and a separate fixed-width pair hash. `SummarizeDocxTextState.ps1` carries those fields into planner
  buckets and sequence-paired Office `Tc` buckets, giving the next public probe a lower-level oracle when whole
  residuals or whole-segment signatures are not enough. This remains diagnostic-only and deliberately does not
  select a renderer `Tc`; keep the implementation branch open for an Office-like decomposition model derived
  from glyph advance/shaping structure rather than text strings, font names, table roles, or observed bucket
  constants. Validation passed `docx-core --skip-slow` (`49`) and full solution build; an initial parallel
  test/build attempt hit a transient output DLL lock and passed on serial rerun.
  2026-06-02 refreshed evidence: reran public `docx-ladder-03-table-text-state` (`20260602-085036`) and
  `docx-ladder-03-text-state-context` (`20260602-085135`) with current candidate DOCX inspect snapshots and
  paired summaries. The table case remains at `MAE=0.632046`, changed16 `0.006503`; all six Office nonzero
  operations are digit segments at `refTc=-0.0182` and pair range `pairMin=2278|pairMax=2278`, with six
  distinct pair signatures. The context case remains at `MAE=0.493604`, changed16 `0.004763`; Office nonzero
  operations span alphanumeric (`5`), digits (`4`), and letters (`2`). Pair advance range separates some
  buckets (`2278 -> -0.0182`, `2618 -> -0.0509`, `2505/2732 -> -0.0437`, `2732 -> 0.0509`), but the same
  `pairMin=2732|pairMax=2732` range maps to both positive and negative Office `Tc` depending on pair
  signature. This confirms pair range is useful evidence but still not a renderer rule; the next discriminator
  must preserve pair-signature/shaping structure and avoid text-class, table-role, font-name, or observed-range
  shortcuts.
  2026-06-02 normalized-pair follow-up: glyph-pair signatures now carry `UnitsPerEm` plus pair advance totals
  and min/max normalized to em units, and `SummarizeDocxTextState.ps1` emits paired Office `Tc` buckets by
  normalized pair-advance range. The context rerun exposes the same structural split in portable units
  (`1.112305em -> -0.0182`, `1.223145em -> -0.0437`, etc.), while still preserving the negative evidence that
  pair range alone cannot choose the sign. Validation passed `docx-core --skip-slow` (`49`), full solution
  build, refreshed `InspectDocx`, and text-state summary on `docx-ladder-03-text-state-context` run
  `20260602-085135`.
  2026-06-02 PDF-width residual follow-up: candidate text-emission advance profiles now also expose the
  residual against the actual rounded PDF `/W` width-array model, not only raw OpenType+kerning natural width.
  `DocxInspect` aggregates `RoundedPdfWidth`, `LayoutToRoundedResidual`, and `RoundedResidualPerGap`, and the
  text-state summary pairs those rounded residual buckets against Office `Tc`. Refreshed public evidence:
  `docx-ladder-03-table-text-state` nonzero Office pairs sit at rounded residual-per-gap `-0.030846` (`2`) and
  `-0.041128` (`4`) for `refTc=-0.0182`; `docx-ladder-03-text-state-context` confirms rounded residual alone
  is still insufficient because `roundResGap=-0.053532` maps to both `refTc=-0.0437` and `refTc=0.0509`.
  Keep this as PDF-level structural evidence; do not turn rounded residuals, pair ranges, text classes, or
  public bucket values into renderer conditions without a deeper Office-like shaping/decomposition model.
  2026-06-02 pair-side advance follow-up: candidate glyph-pair signatures now expose the aggregate and min/max
  left-side and right-side glyph advances for each adjacent pair, with normalized em-unit counterparts, and
  `SummarizeDocxTextState.ps1` pairs Office `Tc` against those side ranges with and without PDF font size.
  Refreshed public runs `docx-ladder-03-text-state-size-matrix` (`20260602-090220`), `docx-ladder-03-table-text-state`
  (`20260602-090232`), and `docx-ladder-03-text-state-context` (`20260602-090241`) show why this matters:
  the previous ambiguous `pairMin=2732|pairMax=2732` context bucket splits structurally into
  `1366+1366 -> refTc=-0.0437` and `1593+1139 -> refTc=0.0509`. A mechanical check over those three public
  cases found zero ambiguous `PdfFontSize + pair-side advance range -> Office Tc` keys. Treat this as stronger
  shaping/decomposition evidence, not as permission to encode a bucket table; the next renderer branch should
  derive a uniform text-state component from Office-like font-size and glyph-pair structure, with public guards
  against token strings, font names, table roles, text classes, and observed constants.
  2026-06-02 PDF text-operation decomposition follow-up: `PdfInspect` now exposes decoded rune count,
  character-spacing gap count, total `Tc` gap contribution, total `TJ` adjustment contribution, and the net
  gap-spacing total for every text operation. `SummarizeDocxTextState.ps1` carries those totals into PDF-only
  and planner-paired buckets. Regenerated inspections for the three public text-state runs show the short
  two-rune nonzero Office operations in `docx-ladder-03-text-state-size-matrix` express the whole spacing
  through `Tc` (`AdjustmentTotalPoints=0`), including the 9pt and 12pt rows where the candidate rounded-width
  residual is near zero. This closes another diagnostic blind spot: the remaining renderer work is to decide
  Office's uniform text-state component from font-size and glyph-pair/shaping structure, not merely to move an
  existing `TJ` adjustment into `Tc` or to replay candidate layout residuals.
  2026-06-02 font-matrix coverage follow-up: added Office-authored public
  `docx-ladder-03-text-state-font-matrix`, which keeps the same short text-state tokens at 11pt while varying
  common Office fonts (`Arial`, `Calibri`, `Times New Roman`, `Courier New`, `Georgia`, `Verdana`). Run
  `20260602-091115` is valid (`MAE=0.766012`, changed16 `0.008298`) and shows the same open renderer gap:
  Office emits `19` nonzero `Tc` operations, candidate emits none. The public summary found zero ambiguous
  `PdfFontSize + pair-side advance range -> Office Tc` keys in this cross-font probe as well. Keep the case as
  an oracle guard against font-name rules: it broadens the structural evidence for glyph-pair/shaping driven
  text-state decomposition, but it still does not justify a lookup table of observed font/pair constants.
  2026-06-01 negative result: a narrower two-encodable-glyph residual split was tested and reverted. The
  rule computed `Tc` from the difference between the already-laid-out segment width and the natural PDF width
  at Office's rounded export font size, applying it only when there was exactly one glyph gap and no authored
  run character spacing. This did move the candidate into the same broad structural class as Word on public
  short-token probes, but it produced systematically wrong signs and magnitudes: `docx-ladder-03-text-state-context`
  candidate nonzero buckets became only negative (`-0.044`, `-0.049`, `-0.051`, `-0.053`, `-0.058`) while Word
  requires both negative and positive buckets (`-0.0182`, `-0.0437`, `-0.0509`, `0.0509`), and
  `docx-ladder-03-text-state-size-matrix` worsened from MAE `0.894614` to `0.899455`. `docx-tables` also
  matched the nonzero operation count but not the Word buckets (`candidate -0.044/-0.053` vs Word
  `-0.0182/0.0509`). Do not use layout-width minus rounded-PDF-width as the DOCX `Tc` rule. The next viable
  branch needs a lower-level public oracle for glyph advance quantization or shaping/spacing decomposition,
  not another residual rule layered on the current measured segment width.
  2026-06-02 follow-up: plain table-cell text normalization is now centralized in `DocxTableCellContent` and
  consumed by font planning plus table-cell height/text-line layout, removing three duplicated synthetic
  paragraph constructions. Added focused bottom-up coverage so legacy `DocxTableCell.Text` cells still enter
  layout as 11pt paragraph text. Validation passed `docx-tables --skip-slow` (`89`), `docx-core --skip-slow`
  (`25`), and `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal`. Keep the broader table-text
  convergence item open for richer cell paragraph ownership, inline images, vertical merge continuations, and
  eventual shared text-frame layout.
  2026-06-02 follow-up: table-cell text line ownership now exposes private-safe `SourceParagraphIndex` through
  `DocxTextLineLayout`, layout snapshots, text-emission line snapshots, text-emission segment snapshots, and
  `SummarizeDocxTextState.ps1` buckets. This does not change rendering, but it removes an inspection blind spot
  where two cell paragraphs with the same line index were indistinguishable once lowered to PDF text operations.
  Added bottom-up coverage for a two-paragraph table cell in both layout and text-emission snapshots, including
  segment-level ownership assertions. Public `docx-tables` diagnostics regenerated for run `20260602-041314`;
  its single-paragraph cells now bucket as `srcPara=0` in `TextClassBySourceParagraphIndex` and paired
  Office-`Tc` summaries. Validation passed `docx-tables --skip-slow` (`97`), `docx-core --skip-slow` (`43`),
  and `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal`. Keep the rendering branch open for actual
  Office row/paragraph allocation and text-state decomposition.
  2026-06-02 follow-up: the same normalized table-cell paragraph stream now also feeds private-safe
  `DocxStructure` table metrics, style/list usage, and `DocxLayoutSnapshot` table-cell text profiles. This
  closes a diagnostic blind spot where legacy plain `DocxTableCell.Text` could render but disappear from
  pre-layout/snapshot counters. Added bottom-up structure and layout snapshot coverage for plain table cells.
  Validation passed `docx-tables --skip-slow` (`91`), `docx-core --skip-slow` (`25`), and full solution build.
  2026-06-02 follow-up: DOCX rendering now has an inspectable private-safe text-emission snapshot on the same
  path as PDF rendering. `DocxRenderer.InspectTextEmission` exposes line/segment counts, terminal line-space
  emissions, source block/line indexes, resolved PDF font resources, Office-grid `/Tf` sizes, authored layout
  character spacing, emitted PDF `Tc`, glyph-positioning spacing, compensation flags, and synthetic face flags
  without exposing document text. Bottom-up coverage asserts both authored run spacing kept in positioned
  glyph advances and numbered-label spacing emitted through PDF text state. Validation passed
  `docx-core --skip-slow` (`27`). Keep the actual generic Word `Tc` selection rule open.
  2026-06-02 follow-up: DOCX body external hyperlinks now emit PDF `/Link` annotations from placed text
  geometry. `DocxTextSpan`, `DocxTextSegmentLayout`, and `DocxTextEmissionSegment` carry source text-run
  indexes, so annotations survive wrapping, justification segmentation, punctuation splitting, and terminal
  line-space synthesis without text scanning. Added PDF-writer coverage for link annotation object emission and
  DOCX renderer coverage for external body links plus internal/bookmark links staying out of URI annotations.
  Validation passed `pdf --skip-slow` (`17`), `docx-core --skip-slow` (`34`), `docx-text --skip-slow` (`45`),
  and full solution build. Follow-up in the next slice extended `DocxTextLineLayout` with source paragraph
  ownership and covered table-cell external hyperlinks through the same placed-segment annotation path;
  `docx-tables --skip-slow` passed (`95`). Static header/footer text spans now also preserve source run indexes
  and source paragraph ownership, and the shared annotation pass includes static text lines; `docx-page
  --skip-slow` passed (`30`). 2026-06-02 PDF-target groundwork: `PdfLinkAnnotation` now carries either a URI
  target or a structural page destination, and `PdfDocumentWriter` emits internal `/Dest [... /XYZ ...]`
  annotations with bottom-up PDF coverage while preserving external `/URI` output. Keep DOCX internal
  destinations open for bookmark inventory and story-owned destination mapping; non-rendered related-story
  links also remain open. 2026-06-02 DOCX bookmark inventory: `DocxReader` now preserves
  `w:bookmarkStart` anchors as paragraph-owned private-safe source/text positions, and `DocxStructureSnapshot`
  exposes bookmark-anchor counts across body, static, table-cell, and related-story paragraph traversal.
  Bottom-up coverage pairs a bookmark anchor with an internal hyperlink target. Keep rendering open until
  bookmark anchors are mapped to placed layout coordinates and emitted through the structural PDF destination
  target. 2026-06-02 rendering follow-up: the DOCX annotation pass now builds a bookmark destination map from
  placed text-emission geometry and emits internal hyperlink annotations as PDF page destinations instead of
  URI actions when the anchor resolves. Bottom-up coverage locks the URI-vs-destination split and clickable
  rectangle geometry. 2026-06-02 table-cell coverage: internal bookmark hyperlinks inside table cells now
  have bottom-up renderer coverage through the same placed table-cell text-emission geometry, proving the
  destination path is story-agnostic for rendered body/table text instead of a body-paragraph special case.
  Keep non-rendered related-story links open, and keep deeper bookmark mapping open for offsets within a
  wrapped text run. 2026-06-02 inline-container follow-up: `DocxReader` now uses a shared
  inline-child dispatcher for hyperlink and inserted-run containers, so `w:bookmarkStart` anchors inside those
  containers are preserved in source/text order instead of being skipped by run-only traversal. Focused
  coverage locks bookmark anchors inside `w:hyperlink`.
- DOCX carriage-return break validation:
  `w:cr` is now preserved as the same soft line-break token as plain `w:br`, instead of being dropped during
  run text extraction. Focused `docx-text --skip-slow` passed `31`, `dotnet build Lokad.OoxPdf.slnx --tl:off
  --nologo -v minimal` passed, and private DOCX run `20260601-093234` stayed neutral at `16/16` pages, zero
  dimension mismatches, no diagnostics, `MAE=13.388935`, changed16 `0.124264`.
- DOCX simple tracked-change final-view validation:
  simple paragraph-child `w:ins` runs are already rendered and direct `w:del` content is hidden, matching the
  final document view instead of exposing revision markup. The `DOCX_UNSUPPORTED_TRACKED_CHANGES` diagnostic is
  now limited to unsupported move/range revision structures and insertions whose content is not direct runs.
  Validation passed `docx-text --skip-slow` (`33`), `docx-core --skip-slow` (`19`), and
  `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal`. Private DOCX run `20260601-095044` stayed
  neutral at `16/16` pages, zero dimension mismatches, no diagnostics, `MAE=13.388935`, changed16 `0.124264`.
- DOCX complex-field cached-result validation:
  complex-field diagnostics now distinguish malformed/evaluation-only fields from closed fields that can be
  rendered from cached result runs. PAGE/NUMPAGES placeholders remain supported, and closed fields with an
  authored cached result no longer warn just because `w:fldChar`/`w:instrText` exists. This keeps static PDF
  rendering aligned with Word's stored field result without adding field-evaluation logic. Validation passed
  `docx-text --skip-slow` (`33`), `docx-core --skip-slow` (`20`), and
  `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal`. Private DOCX run `20260601-095525` stayed
  neutral at `16/16` pages, zero dimension mismatches, no diagnostics, `MAE=13.388935`, changed16 `0.124264`.
  2026-06-02 inline-container follow-up: complex-field diagnostics now validate cached result ranges inside
  the same visible inline containers consumed by the paragraph reader (`w:hyperlink`, `w:fldSimple`, and
  inserted runs), instead of treating non-direct paragraph-run `w:fldChar`/`w:instrText` as unsupported by
  parent shape alone. Bottom-up coverage locks a hyperlink-owned cached result field as rendered text plus
  hyperlink span metadata without `DOCX_UNSUPPORTED_COMPLEX_FIELD`, while the malformed-field diagnostic guard
  still fires. Keep first-class field ownership open for dynamic evaluation, nested fields, and field-specific
  PDF/link semantics.
- DOCX non-positioning tab-stop validation:
  authored `w:tabs` records with `w:val="bar"` or `w:val="clear"` remain preserved in the paragraph model,
  but no longer act as text-positioning stops during layout tab advance. This keeps non-positioning tab
  structure from becoming a false text offset. Focused `docx-text --skip-slow` passed `32`,
  `docx-tables --skip-slow` passed `65`, `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal`
  passed, and private DOCX run `20260601-093551` stayed neutral at `16/16` pages, zero dimension mismatches,
  no diagnostics, `MAE=13.388935`, changed16 `0.124264`.
- DOCX inline page-break validation:
  direct body-paragraph `w:br w:type="page"` tokens inside non-empty runs now split the body stream into
  paragraph fragments separated by a `DocxPageBreakElement`, so text before and after the authored break lands
  on distinct layout pages instead of silently merging. Page breaks are no longer reported as unsupported;
  `DOCX_UNSUPPORTED_MANUAL_BREAK` remains for column breaks. Keep the broader item open for column breaks and
  page breaks inside richer inline containers such as hyperlinks or fields. Validation passed
  `docx-page --skip-slow` (`19`), `docx-core --skip-slow` (`19`), and
  `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal`. Private DOCX run `20260601-093956` stayed
  neutral at `16/16` pages, zero dimension mismatches, no diagnostics, `MAE=13.388935`, changed16 `0.124264`.
  2026-06-01 follow-up: extended the same structural split through visible inline containers already consumed
  by the paragraph reader (`w:hyperlink`, `w:fldSimple`, and inserted runs) while keeping deleted/unsupported
  branches out of final-view pagination. Public coverage now checks hyperlink- and simple-field-contained
  page breaks. Keep the item open for column breaks and richer complex-field/container normalization.
  Validation passed `docx-page --skip-slow` (`21`), `docx-text --skip-slow` (`32`),
  `docx-core --skip-slow` (`19`), and `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal`.
  Private DOCX run `20260601-094513` stayed neutral at `16/16` pages, zero dimension mismatches, no diagnostics,
  `MAE=13.388935`, changed16 `0.124264`.
  2026-06-02 follow-up: run-only column-break paragraphs are now preserved as `DocxManualBreakElement`
  body-flow blocks and surfaced in structure snapshots without turning them into page breaks. Layout clears
  paragraph adjacency across the boundary but remains single-column, so `DOCX_UNSUPPORTED_MANUAL_BREAK` remains
  for actual column-flow support. Validation passed `docx-page --skip-slow` (`29`), `docx-core --skip-slow`
  (`25`), `docx-tables --skip-slow` (`92`), and `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal`.
  A broad unfiltered test run also exercised `DocxReaderPromotesRunColumnBreakOnlyParagraphAsManualBreak` but
  still failed unrelated existing PPTX/font/numbering tests. Private DOCX run `20260602-010936` stayed neutral
  at `16/16` pages, zero dimension mismatches, no diagnostics, `MAE=8.927687`, changed16 `0.095427`.
- DOCX section-break page-boundary validation:
  `DocxSectionBreakElement` is now consumed by layout instead of being skipped. `nextPage`, `oddPage`,
  `evenPage`, and default paragraph section breaks force a page boundary when the current page already has
  content; `continuous` remains in-flow. Section-owned page size, margins, and header/footer distances now
  flow through each layout page into static header/footer placement. Keep `DOCX_UNSUPPORTED_SECTION_BREAK`
  open because continuous-section geometry, true per-section header/footer relationship selection, odd/even
  blank-page parity, and columns are not yet fully applied.
  Validation passed `docx-page --skip-slow` (`22`) and `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo
  -v minimal`. Private DOCX run `20260601-095811` stayed neutral at `16/16` pages, zero dimension mismatches,
  no diagnostics, `MAE=13.388935`, changed16 `0.124264`.
  2026-06-01 follow-up: `DocxLayoutPage` now stores page-local margins and page settings; static header/footer
  rendering uses those instead of document-level final-section geometry. Validation passed `docx-page
  --skip-slow` (`25`), `docx-text --skip-slow` (`36`), full solution build, and public `docx-headers-footers`
  run `20260601-135640` (`MAE=0.073352`, changed16 `0.002110`). Private DOCX run `20260601-135549` stayed
  `16/16` pages with zero dimension mismatches and no diagnostics, with aggregate `MAE=13.838763`,
  changed16 `0.126851`.
  2026-06-01 follow-up: section-local static header/footer reference maps now live on the page settings used by
  layout/rendering, and the DOCX font plan includes those section-scoped static runs. The public
  `DocxSyntheticSectionHeadersUsePageLocalGeometry` test now distinguishes first-section vs final-section
  static content with PDF fill-color operators as well as page-local matrix positions. Validation passed
  `docx-page --skip-slow` (`25`), `docx-text --skip-slow` (`36`), and full solution build; the first parallel
  `docx-text` and private-case attempts hit transient compiler output locks and passed on serial rerun. Public
  `docx-headers-footers` run `20260601-140626` stayed unchanged (`MAE=0.073352`, changed16 `0.002110`).
  Private DOCX run `20260601-140640` stayed at `16/16` pages with zero dimension mismatches and no diagnostics,
  `MAE=13.838763`, changed16 `0.126851`.
- DOCX header/footer font-plan validation:
  the DOCX font plan now includes every referenced header/footer variant, not only the default-selected
  paragraph lists. This prevents first/even static header/footer runs from falling back to a font resource
  whose glyph subset was built without their text or requested typeface. Validation passed `docx-page
  --skip-slow` (`23`), `docx-core --skip-slow` (`20`), and `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo
  -v minimal`. Private DOCX run `20260601-100716` stayed neutral at `16/16` pages, zero dimension mismatches,
  no diagnostics, `MAE=13.388935`, changed16 `0.124264`.
- Public straight stealth connector fixture: `pptx-ladder-06-straight-stealth-connectors` run
  `20260531-124414` passed with tightened gates (`MAE=0.000717`, changed16 `0.00000868`), locking the 6 pt
  minimum marker geometry for 1 pt straight-line stealth ends.
- Public arc guardrail: `pptx-ladder-06-explicit-arc-stealth` run `20260531-124330` passed after the straight
  marker change, confirming the arc/curved path was not broadened.
- Public custom-geometry group-fill fixture: `pptx-ladder-06-custom-geometry-default-fill` run
  `20260531-123236` passed after tightening the gate to MAE `0.05` and changed16 `0.001`.
- Focused shape tests: `dotnet run --project tests\Lokad.OoxPdf.Tests --tl:off --nologo -v minimal -- --group
  pptx-shapes --skip-slow` passed `33` tests.
- Private deck: `pwsh tools\CheckPrivateCase.ps1 -Case private-cases\lokad-value-based.json` run `20260531-025144` compared `84/84` pages, had zero dimension mismatches, no diagnostics, deck MAE `2.877036`, changed16 `0.052651`. Worst pages were page 36 (`6.045372`), page 21 (`5.013651`), page 81 (`4.540970`), page 48 (`4.441990`), and page 20 (`4.406351`). Page 79 improved to about `3.38` MAE after the compressed table-anchor line-box fix. Page 17 is no longer a high-impact schema blocker in that run.
- Public table text-state fixtures: `pptx-ladder-10-table-font-fragmentation` run `20260531-025104` passed at MAE `1.524499`, changed16 `0.026186`; `pptx-ladder-10-table-middle-small-insets` run `20260531-025127` passed unchanged at MAE `0.966760`, changed16 `0.019591`.
- Focused table tests: `dotnet run --project tests\Lokad.OoxPdf.Tests --tl:off --nologo -v minimal -- --group pptx-tables --skip-slow` passed `19` tests.
- `tools\NewOfficeVisualFixtures.ps1` parses after the latest table fixture additions.

Preferred commands for the next slices:

    dotnet run --project tests\Lokad.OoxPdf.Tests --tl:off --nologo -v minimal -- --group pptx-typography --skip-slow
    dotnet run --project tests\Lokad.OoxPdf.Tests --tl:off --nologo -v minimal -- --group pptx-shapes --skip-slow
    pwsh tools\CheckVisualCase.ps1 -Case visual-cases\cases\<case>\case.json
    pwsh tools\CheckPrivateCase.ps1 -Case private-cases\lokad-value-based.json

## Idempotence And Recovery

All build/test/pack commands are safe to rerun. Visual validation writes timestamped directories. If Office
COM automation leaves an Office process running after a failure, close it only after confirming no unrelated
user document is open.

If a private case reveals a missing feature, record only public-safe feature gaps here, then create synthetic
public tests for the implementation. Do not derive public fixtures from private documents.


Historical revision notes were removed during the 2026-05-31 EXECPLAN trim. Open progress items above remain the authoritative backlog.
