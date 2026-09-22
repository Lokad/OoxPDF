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

Per-conversion cumulative budgets (`OoxPdfOptions.ConversionLimits`, PLAN Q01)
bound total work so many individually legal expansions cannot jointly exhaust a
shared process. Defaults are generous multiples of the per-site caps:

- `MaxChartRangeCellsPerConversion` (default 2,000,000): total chart workbook
  range cells expanded, including blank cells materialized during union expansion.
- `MaxTableFragmentsPerConversion` (default 20,000): total DOCX table row
  fragments constructed.
- `MaxImagesDecodedPerConversion` (default 500): total content images decoded
  (each image is still individually pixel-capped).
- `MaxFontWorkPerConversion` (default 5,000): total font program loads and
  subset builds.
- `MaxLiveImageBytesPerConversion` (default 512 MiB): peak transient image
  decode scratch plus pixel planes reserved at any one time (conservative
  width-by-height-by-4 estimate per pixel decode; JPEG passthrough holds none).

Crossing any budget throws `OoxPdfLimitExceededException` before further
expansion: no partial PDF is published (file output stays atomic) and the
failure escapes per-node recovery, so it always aborts the conversion.

With `OoxPdfOptions.ReportResourceUsage`, each successful conversion emits one
informational `CONVERSION_RESOURCE_SUMMARY` diagnostic reporting cumulative
counters plus the peak live reservation
(`pages`, `chartRangeCells`, `tableFragments`, `imagesDecoded`, `fontWork`,
`peakLiveImageBytes`). Informational diagnostics never affect CLI `--strict`
exit codes.

Host admission recipe: run a representative corpus with `ReportResourceUsage`
enabled, take the maximum observed `peakLiveImageBytes` plus headroom for the
per-image pixel cap, and size concurrent conversions as
`maxConcurrent ≈ hostByteBudget / perConversionPeak`. Set tighter
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
