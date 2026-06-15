# Diagnostics

Diagnostics use stable code prefixes such as `OOXML_`, `PPTX_`, `DOCX_`, `PDF_`, `FONT_`, and `IMAGE_`.

Diagnostics are emitted through `OoxPdfOptions.DiagnosticSink` and through the CLI `--diagnostics` JSON file. Warnings mean conversion continued with a fallback or omission. Errors mean conversion failed or the output cannot be trusted.

## PPTX

Unsupported feature warnings:

- `PPTX_UNSUPPORTED_ANIMATION`: slide timing or animation content was detected and ignored.
- `PPTX_UNSUPPORTED_AUDIO`: audio content was detected and ignored.
- `PPTX_UNSUPPORTED_CHART`: chart content was detected and ignored.
- `PPTX_UNSUPPORTED_OLE_OBJECT`: embedded OLE content was detected and ignored.
- `PPTX_UNSUPPORTED_SMARTART`: SmartArt or DrawingML diagram content was detected and ignored.
- `PPTX_UNSUPPORTED_TRANSITION`: slide transition content was detected and ignored.
- `PPTX_UNSUPPORTED_VIDEO`: video content was detected and ignored.

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

## CLI Behavior

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

## Code Conventions

Use stable prefixes by subsystem:

- `OOXML_`: package, ZIP, relationship, XML, or top-level conversion issues.
- `PPTX_`: PowerPoint-specific parser or renderer issues.
- `DOCX_`: Word-specific parser or renderer issues.
- `PDF_`: PDF writer issues.
- `FONT_`: font resolution, parsing, or embedding issues.
- `IMAGE_`: image parsing or rendering issues.

Prefer one durable code per observable behavior. Aggregate repeated unsupported features when the exact count is not useful to the caller.
