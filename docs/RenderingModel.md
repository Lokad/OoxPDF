# Rendering Model

Lokad.OoxPdf is organized as a small pipeline:

1. Open the OOXML ZIP package safely.
2. Resolve content types and relationships.
3. Parse a PPTX or DOCX document model.
4. Lay out document content into PDF pages.
5. Emit deterministic PDF objects and page content streams.

The production library does not shell out to Office, PDFium, PowerShell, or external tools.

Chart painting requires a resolved scene. Plot and secondary axes are selected
from that scene, preserving declared ID order and explicit absence. The selected
axis owns its resolved properties and explicit absence. The chart retains its
original XML for properties still read from markup; a second document cannot
supply a missing axis.

Chart data, names, labels, paints, markers and pie/doughnut slice offsets consume
their parsed plots and series. Cached/workbook values retain source indices and
per-point overrides. Numeric bounds, units, orientation, crossing, gridlines and number formats use
the selected parsed axis. Axis measurement and painting carry the parsed
chart-root date system; source XML remains available for properties not yet
represented in the scene. Missing axes keep their explicit
defaults rather than acquiring declarations from another XML source.
Reciprocal category-axis IDs own crossing and reversed label sides; absent axes
retain automatic label defaults.

## Package Layer

The OOXML layer reads ZIP entries by normalized package part name. It rejects path traversal and applies entry-count and uncompressed-size limits before higher-level parsers touch XML or binary media.

XML is read with DTD and external entity resolution disabled. Package relationships are resolved relative to the source part, so document parsers can ask for related slides, themes, media, headers, footers, numbering parts, and styles without reimplementing ZIP path logic.

## Document Parsers

The PPTX parser discovers the presentation part, slide order, slide size, slide relationships, slide layouts, slide masters, themes, media relationships, tables (including merged-cell continuations), and chart parts with cached numeric values plus embedded workbook provenance. Rendering then walks slide, layout, and master shape trees in draw order for the common static features currently supported, including native bar/column, line, area, pie/doughnut, scatter, bubble, and radar charts.

The DOCX parser discovers the main document part, styles, numbering, section properties, headers, footers, paragraphs, runs, inline images, and fixed-grid tables. It applies a simple style cascade for document defaults, paragraph styles, character styles, and direct formatting.

Note markers retain source-character or generated-run ownership independently of
XML run grouping. Numbering and admission use those owners; custom marks bind only
to an immediately following source character.

Stored field results recognize text, pictures and mathematical inline items through one
source-content rule. The diagnostic scan follows the reader's transparent
metadata containers. Parsed pictures and equations share result notification through nested
fields; computed counters keep their replacement owner. Layout consumes the
materialized inline sequence without a field-specific geometry correction.
Complete stored results may be empty: their source flags remain empty, while
neighboring supported inline content remains eligible. Computed placeholders and
instructions missing a result boundary retain their existing paths.

Paragraph-property resolution retains its unique parsed change owner separately
from aggregate revision metadata. Inline eligibility consumes that same owner;
run, mark, range and inherited revisions still require their own source. This
preserves equations beside formatting history without adding a geometry correction.

Unsupported features are detected as early as practical and emitted through `OoxPdfOptions.DiagnosticSink`.

## Layout Layer

DOCX layout resolves omitted header/footer distances to a half inch, independently
of body margins. Declared distances and raw source inspection metadata remain intact.

PPTX layout is mostly absolute: shapes, images, text boxes, groups, and tables already carry Office coordinates in EMUs. The renderer converts EMUs to PDF points, composes group transforms, resolves theme colors/fonts, and writes drawing operations to the target page.

Measured chart category layouts carry the gap used by admission and painting.
For rotated outer captions, mirroring preserves that clearance and authored
offsets add normal-line spacing. Inner and automatic layouts retain their
separate baseline rules; signed-angle horizontal admission remains approximate.

For ordinary clustered columns, one ordinal coordinate transform places data,
captions and data-label/link slots. Reversal mirrors the within-cluster series
offsets while preserving source indices, sparse values and point formatting.
Plain labels, automatic labels and legend-key units use actual bar centers.

DOCX layout is flow-based. The renderer converts twips to points, computes the content rectangle from page size and margins, performs simple greedy Latin line breaking, handles keep rules, split table rows with repeated headers, floating-drawing exclusion geometry, and markup balloon contexts, advances a page cursor, and creates new pages when content crosses the bottom margin. Headers and footers are drawn on each produced page.

Word-compatible body origins belong to each printed page. Painting, inspection,
bookmark/link geometry and following-note boundaries use that page's origin;
mapped static stories and bottom-anchored notes retain their own frames.

Run and paragraph-mark formatting descriptions use the same property resolver.
The mark and aggregate metadata share one resolved revision object. Caption
recognition is separate from placement. Qualified footer captions and their
frames share the source strut and physical story transform. Custom Review paper
origins and remaining font/frame combinations retain the documented approximations.

In qualified unscaled document-end notes, measured source paragraphs own the
start and adjacent paragraph boundaries even when later notes need another page.
Complete table/image blocks retain their trailing paragraph's ownership; the
container still admits each following head against the available page frame.
Whole-group fitting, typography, counter binding and clipping remain separate
decisions. Unsupported source/static/scaled frames retain their existing fallbacks.
Endnote counter binding admits a complete measured prefix against the selected
footer edge. It stops at the first unfit or unsupported block; later notes do
not change the reference page of an already admitted source block. Section-end
and document-end placement use the same proof, independently of whole-group flow.

Supported mixed DOCX content uses ordered source items for text, tabs, breaks,
pictures and equations. Measurement produces inline fragments without page
coordinates; body, cell and related-story containers own wrapping, admission and
placement. Ordered final lines retain their container's resolved paragraph
after-spacing so annotations include the same gap as layout. Painting and
annotation bounds consume those placed frames. The reader distinguishes root-row
math links from argument links: root fragments merge per placed line, while
argument selections retain local paint bounds. An authored
body page break preserves source ownership on both sides, including linked
formulas and the note associated with a reference. Unsupported combinations keep
their explicit fallbacks; this is not unrestricted Word reflow.

Current text layout intentionally avoids complex shaping. Bold/italic use a hybrid: prefer a real font face on exact non-fallback matches, otherwise synthesize (stroked bold pass, oblique shear). Font programs flow through FontRequest, FontFaceResolution, IFontProgramSource, OpenTypeFont, and the subsetted embedded PDF font (see [fonts and portability](Fonts.md)).

## Shape, media, and table pipeline (PPTX)

Shape rendering is a dispatch-plus-lanes pipeline, not one module. `PptxRenderer.Shapes.cs` owns
dispatch (`RenderBackground`, two `RenderShape` overloads, group transforms) and fans out to
single-home lanes:

- `ShapeShared.cs`: cross-file shared infra - group-transform application, pattern-fill stroking,
  shape line and font-color readers. These live in Shared (not a Readers file) because shape
  dispatch, text, and scene readers all call them.
- `ShapeStyles.cs`: scene-to-renderer style converters (`ToShapePatternFill`, `ToGradientFill`,
  `ToShapePictureFill`, `ToGlow`, `ToOuterShadow`, line-end/dash/cap/join readers, `ToLineStyle`,
  `ToFillStyle`).
- `ShapeGeometry.cs`: custom-geometry engine (guide/formula evaluation, path readers, bezier builders).
- `ShapeConnectors.cs`: line-end markers and arrowheads plus the curved-connector bezier engine.
- `ShapeEffects.cs`: effect emission (outer shadow, glow, gradient and preset fills, preset arcs).
- `Images.cs` (plus `Images.Svg.cs`): the media home - picture-frame rendering, image cache and
  recolor, crop/fill-rect conversion, plus picture-fill IO (`TryReadShapePictureFill`,
  `DrawImageFill`). Picture fills ride the shape fill-resolution path: `RenderShape` resolves them
  mid-dispatch and draws through the image cache, so fill IO stays with the cache instead of
  splitting across shape files.
- Tables: `Tables.cs` (frame pipeline), `Tables.Borders.cs` (border-stroking lane), `TableStyles.cs`
  (style resolution). Scene-side readers mirror the split: `PptxScene.Shapes.cs` core plus the
  `Shapes.Lines`, `Shapes.Fills`, `Shapes.Effects`, and `Shapes.Pictures` lanes.

## PDF Layer

The PDF writer emits a static PDF with deterministic object ordering, stable resource names, subsetted embedded TrueType/CID fonts, ToUnicode maps, path drawing, text drawing, and JPEG/PNG image XObjects. PNG alpha is represented with a soft mask when needed.

Generated PDFs avoid volatile metadata and use stable resource names. Tests check repeated conversion bytes with the same inputs and font configuration.

## Diagnostics

Diagnostics carry a stable `Id`, severity, message, optional package part name, optional slide/page index, feature name, and fallback description. CLI `--strict` treats warning or error diagnostics as exit code `3` after a successful conversion.
