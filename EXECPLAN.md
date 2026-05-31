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
- [ ] 2026-05-31: Investigate private slide 42 as a high-priority PPTX schema/text-layout issue. On the left
  schema, Office places the numbers centered inside their rectangles, while the candidate places the numbers
  incorrectly and emits the wrong color. Treat this as a generic shape/text-frame alignment and inherited text
  color problem, not as a private-content coordinate tweak: inspect the Office/candidate PDF text operations,
  rectangle geometry, fill/stroke/text color states, and OOXML style inheritance, then reproduce the underlying
  behavior with public synthetic fixtures before changing production rendering.
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
    authored top margins move the first baseline before PDF emission; the legacy 4 pt inset remains only for
    cells without explicit margins.
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
  - [x] 2026-05-31: Preserved DOCX numbering-level indent tokens and applied a first layout-stage indent
    approximation for numbered paragraphs. `DocxListLabel` now carries typed left/right/first-line/hanging
    indent values from `w:lvl/w:pPr/w:ind`, and body/table-cell paragraph layout uses those values to shift
    numbered line starts and reduce wrapping width. This improves private DOCX aggregate fidelity while keeping
    `DOCX_NUMBERING_INDENT` open as an approximation because Office's true label tab stop, hanging text start,
    and continuation-line behavior are not yet structurally modeled.
  - [ ] 2026-05-31: Split DOCX numbering layout into label run, tab stop, text start, and continuation-line
    geometry. The current indent application moves the flattened label+text line as one unit; long-term Word
    parity needs a layout record that separates the list marker from paragraph text and uses numbering
    tab/hanging rules consistently in body text and table cells.
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

## Backlog

### Release-Blocking Fidelity

- [ ] Extend PPTX chart rendering beyond the current loose native renderer: labels, legends, axes,
  marker styles, theme/chart style colors, and tighter plot-area layout fidelity.
- [ ] Fix DOCX page geometry and pagination fidelity: section page size/margins, paragraph spacing, manual
  page/column breaks, and keep/widow page-break decisions.
- [ ] Text layout: preserve spaces, tabs, line breaks, soft line breaks, kerning-like advances, font
  fallback, mixed run spacing, character spacing, superscript/subscript, and baseline offsets.
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
