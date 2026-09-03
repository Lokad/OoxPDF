# Rendering Model

Lokad.OoxPdf is organized as a small pipeline:

1. Open the OOXML ZIP package safely.
2. Resolve content types and relationships.
3. Parse a PPTX or DOCX document model.
4. Lay out document content into PDF pages.
5. Emit deterministic PDF objects and page content streams.

The production library does not shell out to Office, PDFium, PowerShell, or external tools.

## Package Layer

The OOXML layer reads ZIP entries by normalized package part name. It rejects path traversal and applies entry-count and uncompressed-size limits before higher-level parsers touch XML or binary media.

XML is read with DTD and external entity resolution disabled. Package relationships are resolved relative to the source part, so document parsers can ask for related slides, themes, media, headers, footers, numbering parts, and styles without reimplementing ZIP path logic.

## Document Parsers

The PPTX parser discovers the presentation part, slide order, slide size, slide relationships, slide layouts, slide masters, themes, media relationships, tables (including merged-cell continuations), and chart parts with cached numeric values plus embedded workbook provenance. Rendering then walks slide, layout, and master shape trees in draw order for the common static features currently supported, including native bar/column, line, area, pie/doughnut, scatter, bubble, and radar charts.

The DOCX parser discovers the main document part, styles, numbering, section properties, headers, footers, paragraphs, runs, inline images, and fixed-grid tables. It applies a simple style cascade for document defaults, paragraph styles, character styles, and direct formatting.

Unsupported features are detected as early as practical and emitted through `OoxPdfOptions.DiagnosticSink`.

## Layout Layer

PPTX layout is mostly absolute: shapes, images, text boxes, groups, and tables already carry Office coordinates in EMUs. The renderer converts EMUs to PDF points, composes group transforms, resolves theme colors/fonts, and writes drawing operations to the target page.

DOCX layout is flow-based. The renderer converts twips to points, computes the content rectangle from page size and margins, performs simple greedy Latin line breaking, handles keep rules, split table rows with repeated headers, floating-drawing exclusion geometry, and markup balloon contexts, advances a page cursor, and creates new pages when content crosses the bottom margin. Headers and footers are drawn on each produced page.

Current text layout intentionally avoids complex shaping. Bold/italic use a hybrid: prefer a real font face on exact non-fallback matches, otherwise synthesize (stroked bold pass, oblique shear). Font programs flow through FontRequest, FontFaceResolution, IFontProgramSource, OpenTypeFont, and the subsetted embedded PDF font (see README Fonts And Portability).

## PDF Layer

The PDF writer emits a static PDF with deterministic object ordering, stable resource names, subsetted embedded TrueType/CID fonts, ToUnicode maps, path drawing, text drawing, and JPEG/PNG image XObjects. PNG alpha is represented with a soft mask when needed.

When deterministic conversion is requested, the output avoids unstable metadata and resource naming so the same input package produces byte-identical output in tests.

## Diagnostics

Diagnostics carry a stable `Id`, severity, message, optional package part name, optional slide/page index, feature name, and fallback description. CLI `--strict` treats warning or error diagnostics as exit code `3` after a successful conversion.
