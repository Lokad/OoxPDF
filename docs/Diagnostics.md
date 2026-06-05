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

- `DOCX_UNSUPPORTED_COMMENTS`: comment markup was detected and ignored.
- `DOCX_UNSUPPORTED_COMPLEX_FIELD`: a non-page-number field was detected and ignored or approximated.
- `DOCX_UNSUPPORTED_ENDNOTE`: endnote references were detected and ignored.
- `DOCX_UNSUPPORTED_EQUATION`: Office Math content was detected and ignored.
- `DOCX_UNSUPPORTED_FLOATING_DRAWING`: floating DrawingML content was detected and ignored.
- `DOCX_UNSUPPORTED_FOOTNOTE`: footnote references were detected and ignored.
- `DOCX_UNSUPPORTED_MACRO`: a VBA project was detected and ignored.
- `DOCX_UNSUPPORTED_MULTI_COLUMN`: a multi-column section was detected and rendered as a single column.
- `DOCX_UNSUPPORTED_OLE_OBJECT`: embedded OLE content was detected and ignored.
- `DOCX_UNSUPPORTED_TRACKED_CHANGES`: tracked insertion or deletion markup was detected and approximated.

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
