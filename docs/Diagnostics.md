# Diagnostics

Diagnostics use stable code prefixes such as `OOXML_`, `PPTX_`, `DOCX_`, `PDF_`, `FONT_`, and `IMAGE_`.

Diagnostics are emitted through `OoxPdfOptions.DiagnosticSink` and through the CLI `--diagnostics` JSON file. Warnings mean conversion continued with a fallback or omission. Errors mean conversion failed or the output cannot be trusted.

## OOXML

Top-level package or dialect issues:

- `OOXML_STRICT_DIALECT`: Strict OOXML (ISO 29500) content was detected; only the transitional dialect is supported and content may be missing.
- `OOXML_MUST_UNDERSTAND`: content marked must-understand uses unsupported namespaces and was ignored. Main parts warn once per document; slides, shared masters/layouts, and DOCX story and infrastructure parts warn once per part with their part names.

## PPTX

Unsupported feature warnings:

- `PPTX_UNSUPPORTED_ANIMATION`: slide timing or animation content was detected and ignored.
- `PPTX_UNSUPPORTED_AUDIO`: audio content was detected and ignored.
- `PPTX_UNSUPPORTED_CHART`: an unsupported chart kind was detected, a chart part was missing/unresolvable, or a supported chart referenced formula-only data without cached numeric values (see also `PPTX_CHART_MISSING_CACHED_DATA`). Native rendering covers bar/column, line, area, pie/doughnut, scatter, bubble, and radar charts with cached values.
- `PPTX_UNSUPPORTED_OLE_OBJECT`: embedded OLE content was detected and ignored.
- `PPTX_UNSUPPORTED_SMARTART`: SmartArt or DrawingML diagram content was detected and ignored.
- `PPTX_UNSUPPORTED_TRANSITION`: slide transition content was detected and ignored.
- `PPTX_UNSUPPORTED_VIDEO`: video content was detected and ignored.
- `PPTX_CHART_MISSING_CACHED_DATA`: a supported chart referenced formula-only data without chart-side cached numeric values; workbook provenance is preserved but not used for layout.
- `PPTX_UNSUPPORTED_CHART_AXIS_TITLE_LAYOUT`: a chart axis title with manual positioning that cannot be honored (default-placement titles render natively).
- `PPTX_UNSUPPORTED_CHART_AXIS_TITLE_AXIS_POSITION`: a default-placement chart axis title has an unsupported or missing axis kind/position.
- `PPTX_UNSUPPORTED_CHART_TRENDLINE`: a chart series defines a trendline, which is not rendered.
- `PPTX_UNSUPPORTED_CHART_NUMBER_FORMAT`: a chart number format uses unsupported syntax (locale rules, colors, native digit shapes, `!` escapes, `@` placeholders, or unknown brackets) that is dropped; one warning per construct per chart.
- `PPTX_UNSUPPORTED_GRAPHIC_FRAME`: an unsupported graphic frame was detected and ignored.
- `PPTX_UNSUPPORTED_GRADIENT_FILL`: an unsupported gradient fill was detected and ignored.
- `PPTX_UNSUPPORTED_PATTERN_FILL`: an unsupported pattern fill was detected and ignored.
- `PPTX_UNSUPPORTED_TEXT_ORIENTATION`: an unrecognized text orientation was detected and ignored (horz, vert, vert270, eaVert, mongolianVert, and the wordArt variants render).
- `PPTX_UNSUPPORTED_TEXT_OVERFLOW`: text vertical overflow uses ellipsis; local clipping was applied but the ellipsis marker is not rendered.
- `PPTX_UNSUPPORTED_PICTURE_FILL`: an unsupported picture fill was detected and ignored.
- `PPTX_UNSUPPORTED_IMAGE_TILE`: a tiled image fill was detected and ignored.
- `PPTX_UNSUPPORTED_IMAGE_RECOLOR`: an unsupported image recolor was detected.
- `PPTX_UNSUPPORTED_TRANSPARENCY`: unsupported transparency was detected and ignored.
- `PPTX_UNSUPPORTED_EFFECT`: an unsupported effect was detected and ignored.
- `PPTX_UNSUPPORTED_CUSTOM_GEOMETRY`: unsupported custom geometry was detected and ignored.
- `PPTX_UNSUPPORTED_CALLOUT`: a callout shape was detected and ignored.
- `PPTX_UNSUPPORTED_TABLE_STYLE`: an unsupported table style was detected.
- `PPTX_UNSUPPORTED_HYPERLINK`: a shape, frame, text-run, or chart title hyperlink pointed at an unresolvable, empty, unknown-slide, or unsupported-typed relationship and was ignored. Repeated uses of the same relationship on one slide report once.
- `PPTX_UNSUPPORTED_HYPERLINK_ACTION`: a hyperlink carried only an action (no relationship target) and was ignored. Repeated uses of the same action on one slide report once.
- `PPTX_NODE_RENDER_FAILED`: a slide node failed to render.

These warnings are slide-scoped when a slide index is available. Duplicate occurrences of the same unsupported feature on one slide are aggregated into one warning for that slide.

## DOCX

Unsupported feature warnings:

- `DOCX_UNSUPPORTED_COMMENTS`: comment markup was detected and ignored, usually in compatibility/default final output when markup printing was not requested.
- `DOCX_UNSUPPORTED_COMPLEX_FIELD`: a complex field was malformed in a way that prevents cached-result or placeholder rendering, nested inside another field instruction, or had no supported dynamic placeholder and no cached result to render.
- `DOCX_UNSUPPORTED_ENDNOTE`: endnote references were detected and ignored.
- `DOCX_UNSUPPORTED_EQUATION`: Office Math content was detected and ignored.
- `DOCX_UNSUPPORTED_FLOATING_DRAWING`: floating DrawingML content was malformed or used an unsupported payload such as a chart, SmartArt, external image, or anchor positioning outside the supported image/text-box placement model.
- `DOCX_UNSUPPORTED_FOOTNOTE`: footnote references were detected and ignored.
- `DOCX_UNSUPPORTED_MACRO`: a VBA project was detected and ignored.
- `DOCX_UNSUPPORTED_MULTI_COLUMN`: a multi-column section was detected and rendered as a single column.
- `DOCX_UNSUPPORTED_OLE_OBJECT`: embedded OLE content was detected and ignored.
- `DOCX_UNSUPPORTED_TRACKED_CHANGES`: tracked insertion or deletion markup was detected but the selected mode did not request visible markup support.
- `DOCX_UNSUPPORTED_VML`: VML drawing content outside the supported inline image subset was detected and ignored.

Approximation warnings:

- `DOCX_APPROXIMATED_COMMENTS`: comment markup was rendered through the selected markup mode with first-pass markers or balloons.
- `DOCX_APPROXIMATED_FORMATTING_REVISIONS`: formatting revisions were detected and surfaced through private-safe property-family provenance, first-pass formatting balloons, change bars, or approximation diagnostics.
- `DOCX_NUMBERING_INDENT`: a numbering indentation variant outside the supported twip-based left/right/first-line/hanging/tab model was detected, such as character-unit list indents.
- `DOCX_STYLE_PARAGRAPH_SPACING`: a paragraph style used a spacing variant outside the supported before/after, line-unit, automatic, contextual, exact, auto, and at-least line-spacing model.
- `DOCX_TABLE_BORDER_STYLE`: a table or table style used a border style outside the supported solid, thick, double, triple, dotted, dashed, dash-small-gap, dash-dot-stroked, dot-dash, dot-dot-dash, thin/thick compound, wave/double-wave, 3D emboss/engrave, outset, inset, nil, and none set.
- `DOCX_APPROXIMATED_TRACKED_CHANGES`: tracked changes were rendered through the selected markup mode with final/original filtering, change bars, inline revision styling, grouped first-pass revision balloons, or compact overflow continuations, but without full Word-style review-pane geometry.

These warnings are document-scoped. Duplicate occurrences of the same unsupported feature in one document are aggregated into one warning.

## Fonts

- `FONT_UNSUPPORTED_OUTLINES`: a CFF/OpenType-CFF typeface was detected and not embedded (native CFF embedding is unsupported). DOCX substitutes the document fallback typeface for primary runs, skips CFF per-character fallback candidates, and degrades CFF-only fallbacks to the missing-font path; PPTX regroups uses through its per-glyph fallback. One warning per typeface per conversion.

The CLI writes diagnostics JSON when `--diagnostics <file>` is provided. The JSON is an array of diagnostic entries with these fields when available:

- `Id`: stable diagnostic code.
- `Severity`: `Info`, `Warning`, or `Error`.
- `Message`: human-readable explanation.
- `PartName`: OOXML package part where the issue was detected.
- `SlideIndex`: one-based slide index for slide-scoped PPTX diagnostics.
- `PageIndex`: one-based page index when available.
- `Feature`: short unsupported or approximated feature name.
- `Fallback`: short description of the fallback, such as `ignored` or `approximated`.

Exit codes:

- `0`: conversion succeeded.
- `1`: conversion failed and an `OOXML_CONVERSION_FAILED` error diagnostic is emitted when a diagnostics path was supplied.
- `2`: invalid arguments.
- `3`: conversion succeeded, but `--strict` saw at least one warning or error diagnostic.

## Resource Budgets and Host Admission

Per-conversion cumulative budgets (`OoxPdfOptions.ConversionLimits`) bound
covered work so many individually legal expansions cannot jointly exhaust a
shared process. Defaults are generous multiples of the per-site caps:

- `MaxChartRangeCellsPerConversion` (default 2,000,000): total chart workbook
  range cells expanded, including blank cells materialized during union expansion,
  plus chart dense slots materialized from cached/literal vectors (R04).
- `MaxTableFragmentsPerConversion` (default 20,000): total DOCX table row
  fragments constructed.
- `MaxXmlNodesPerConversion` (default 40,000,000): total XML elements plus
  attributes parsed across every part; cached parses do not recharge (R04).
- `MaxWorkbookCellsPerConversion` (default 1,000,000): total chart workbook
  cells read across every embedded workbook; cached models do not recharge (R04).
- `MaxSceneNodesPerConversion` (default 2,000,000): total PPTX scene nodes
  built across slides, masters, and layouts, including nested group children (R04).
- `MaxNestedPackageBytesPerConversion` (default 512 MiB): total retained bytes
  across nested (embedded workbook) packages, each already capped individually (R04).
- `MaxWorkbookModelsPerConversion` (default 100): total retained chart workbook
  models; cached models do not recharge (R04).
- `MaxPagesPerConversion` (default 10,000): total PDF pages serialized (R06).
- `MaxPdfContentBytesPerConversion` (default 1 GiB): total encoded page-content
  bytes serialized (R06).
- `MaxOutputBytesPerConversion` (default 2 GiB): total PDF output bytes admitted before each write
  while the conversion scope is still open; a zero budget writes nothing (R06.1).
- `MaxPdfFontBytesPerConversion` (default 256 MiB): total embedded font program
  plus ToUnicode bytes serialized, complementing the font-work count (R06).
- `MaxPdfImageBytesPerConversion` (default 512 MiB): total encoded image and
  soft-mask bytes serialized, complementing the image-decode count (R06).
- `MaxRetainedImageBytesPerConversion` (default 512 MiB): total encoded image
  plus soft-mask bytes of newly created image resources at the owning producer
  (R06.2). Cache hits create nothing and do not recharge; pruned resources stay
  charged. Intentionally separate from the serialized image cap.
- `MaxRetainedFontBytesPerConversion` (default 256 MiB): total subset (or whole
  fallback) font program bytes at subset construction (R06.2). Identical-merge
  fast paths build nothing and do not recharge. Intentionally separate from the
  serialized font cap.
- `MaxImagesDecodedPerConversion` (default 500): total content images decoded,
  including PPTX crop/recolor variants and effect rasters, through the shared
  decoder boundary (R03). Each image is still individually pixel-capped; cache
  hits do not recharge.
- `MaxFontWorkPerConversion` (default 5,000): total font program loads and
  subset builds through the shared loader/subsetter boundary, covering DOCX and
  PPTX (R03). Ordinary conversions resolve dozens; thousands indicate reference
  churn. Cache hits do not recharge.
- `MaxLiveImageBytesPerConversion` (default 512 MiB): peak image-decode
  reservation using per-format working-set estimates held across
  decode/transform/compress by decoded-pixel ownership (R02; R01 reserves
  before component-plane allocation; JPEG passthrough holds none).

Crossing any budget throws `OoxPdfLimitExceededException` before further
expansion: no partial PDF is published (file output stays atomic) and the
failure escapes per-node recovery, so it always aborts the conversion.
Limit exceptions escape tolerant loader/diagnostic catch filters.

With `OoxPdfOptions.ReportResourceUsage`, each successful conversion emits one
informational `CONVERSION_RESOURCE_SUMMARY` diagnostic reporting cumulative
counters plus the peak live reservation
(`pages`, `chartRangeCells`, `tableFragments`, `imagesDecoded`, `fontWork`,
`xmlNodes`, `workbookCells`, `peakLiveImageBytes`, plus `pdfPages`,
`pdfContentBytes`, and `pdfOutputBytes` writer-stage fields). File conversions
snapshot after serialization, so the page/content/output fields carry post-write
values before atomic publication; stream conversions snapshot after rendering but
before serialization, so page/content fields carry render-stage values while output
stays zero even when the emitted PDF contains those bytes. Both omit
charged domains without a summary field (font/image bytes, scene nodes, nested
package bytes, workbook models). Informational diagnostics
never affect CLI `--strict` exit codes.

Scope limits (R19): `peakLiveImageBytes` is a reservation peak, not total live
or process memory. It covers per-format working-set estimates held across
decode/transform/compress (R02) but omits pixels retained outside any live
operation scope, compressed PDF resources, fonts, XML DOMs, pages, and writer
work. Retained image/font production carries admission caps plus `retainedImageBytes`/
`retainedFontBytes` summary fields (R06.2); serialized font/image bytes charge
without summary fields (R19 follow-up). Do not size
hosts from `peakLiveImageBytes` plus image headroom alone.

Domain map (R06.2/V01) — every charged domain against its admission bound, or its
named residual:

- Chart range cells, dense slots, table fragments, XML nodes, workbook cells and
  models, nested package bytes, scene nodes: aggregate cumulative caps (R04).
- Images: decode counts at the shared decoder boundary (R03); live working set as
  a reservation peak (R02); retained encoded bytes at the owning producer with
  no recharge on cache hits and no refunds (R06.2); serialized bytes at the
  writer (R06).
- Fonts: load/subset counts at the shared loader boundary (R03); retained programs
  at subset construction with identical-merge fast paths exempt (R06.2);
  serialized bytes plus ToUnicode maps at the writer (R06).
- Pages, content, output: per-page admission during rendering (R06.2),
  serialization totals (R06), pre-write output admission at every write (R06.1).
- Vector, shading, and annotation writer data: no byte quota (residual; writer
  validation still applies).
- XML/text payload heaps, font subset/compress scratch, writer scratch: counts
  bound work, not heap bytes (residual; V01 measures).
- DOCX repagination passes: fragment/table budgets bound the inputs, but the
  passes themselves have no separate control (residual).

Host admission: measure conversion-only live/process peaks across
page/image/font/chart breadth with `ReportResourceUsage`, separate discovery,
transient allocation, retained ownership, and output buffering, and size
admission from total peaks plus host baseline/concurrency headroom. Set tighter
`ConversionLimits` for shared processes and tune them against measured
corpora; `ConvertAsync` offloading to `Task.Run` is not admission control.

## Code Conventions

Use stable prefixes by subsystem:

- `OOXML_`: package, ZIP, relationship, XML, or top-level conversion issues.
- `PPTX_`: PowerPoint-specific parser or renderer issues.
- `DOCX_`: Word-specific parser or renderer issues.
- `PDF_`: PDF writer issues.
- `FONT_`: font resolution, parsing, or embedding issues.
- `IMAGE_`: image parsing or rendering issues.

Prefer one durable code per observable behavior. Aggregate repeated unsupported features when the exact count is not useful to the caller.
