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
