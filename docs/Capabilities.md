# Rendering scope

Lokad.OoxPdf renders static DOCX and PPTX content. Support is bounded by the rules
below; it does not imply complete Office compatibility. Diagnostics identify many
fallbacks, but visible content and geometry must also be checked independently.

## Input and PDF output

- Reads OOXML ZIP relationships, content types, themes and supported markup-
  compatibility alternatives with bounded archive/XML/image/font processing.
- Supports ordinary Transitional DOCX/PPTX. Common Strict DOCX content uses the
  existing pipeline with `OOXML_STRICT_DIALECT` / Partial; Strict PPTX is unsupported.
- Produces deterministic PDF 1.7 with embedded/subset TrueType fonts, Unicode maps,
  raster pictures, supported vector paint, transparency and hyperlink annotations.
- Does not produce tagged PDF, PDF/A or PDF/UA, document outlines, interactive
  forms, scripts, encryption or embedded source files. See [architecture](Architecture.md).

## DOCX

| Area | Supported surface | Important limits |
|---|---|---|
| Text | Paragraph/run styles, common formatting, tabs, Latin wrapping, alignment, spacing, lists and supported fonts | General complex-script shaping and full Word typography are not provided. |
| Lists | Decimal/padded decimal, Roman 1–3999, Latin letters, referenced-level formats, overrides and legal decimal references | Unknown formats use decimal fallback; locale alphabets and custom `lvlRestart` policies are unqualified. |
| Pagination | Page/column breaks, section frames, supported keep/widow rules, repeated table headers and page fields | Column section starts can share a page when their frames and page-level ownership agree. Other mixed sections, oversized/fragmented content and complex page ownership retain restrictions. |
| Tables | Fixed grids, supported widths, merges, borders, padding, vertical alignment, nested tables and mixed cell content | Automatic sizing is approximate. Explicit widths can extend past the page edge and be cropped. Exact heights keep their clips. Arbitrary row splitting is unsupported. |
| Pictures and boxes | Supported inline/anchored pictures and text boxes, transforms, clipping and first-pass wrap exclusions | Nested inline boxes are unsupported. Complex floating overlap/wrap remains approximate. |
| Stories | Section-selected headers/footers, footnotes/endnotes, supported fields and navigation | Note continuation/numbering eligibility is restricted; general note-row fragmentation is unsupported. |
| Review | Final, Original, SimpleMarkup and AllMarkup, supported revisions, comment indicators and balloons | Word print fitting, mixed balloon collision and custom-paper Review placement remain approximate. |
| Navigation | External links and supported bookmark/note targets using placed content, including body-level bookmarks around paragraphs and tables | Some unplaced or fragmented owners remain unsupported. |

Supported inline text, equations and pictures retain source order through body,
table and supported static-story containers. Bookmarks, links and supported
revision metadata preserve their owners. This does not promise every composition
of otherwise supported objects.

Symbol runs retain printable BMP character codes and their declared font.
Malformed or nonprinting symbol payloads use a diagnosed replacement character;
font/glyph availability still controls whether the authored symbol can be painted.

Native equations support bounded common OMML: runs, fractions, scripts, roots,
accents/bars, delimiters, n-ary operators and supported arrays/matrices. Structural
pieces can use PDF ActualText. Literal math operators keep their symbols and use
ordinary spacing. Supported equations also render in placed floating
text boxes, within their authored and ancestor clips. Placement, baselines, line breaks, large operators
and font-dependent geometry remain approximate. Unsupported constructs/placements
are diagnosed and can be omitted. Discretionary hyphens inside equations stay
invisible. Equation wrapping uses the existing supported operator breaks. Math
fonts and coverage matter.

Floating charts, SmartArt, OLE, macros, external images and general Word field
execution are unsupported. Supported counters are computed; other fields retain
usable stored results rather than executing arbitrary instructions.

## PPTX

- Uses authored slide dimensions and absolute shape/group transforms, supported
  slide/layout/master inheritance, themes, text, bullets, tables, pictures,
  fills, strokes, transparency and external/internal hyperlinks.
- Highlights and explicit bullet colors retain supported DrawingML colors and
  opacity transforms. Follow-text bullet opacity retains the current opaque
  fallback; its Office behavior is uncalibrated.
- Renders cached-data bar/column, line, area, pie/doughnut, scatter, bubble and radar
  charts, including supported axes, labels, legends and explicit styles. Numeric
  chart-side caches are required. Embedded workbooks provide source provenance and
  row/column visibility; formulas are not recalculated and formula-only series
  do not fall back to workbook values.
- Complete unique nonnegative series orders resolve within a plot. Missing,
  negative, duplicate or malformed orders retain XML sibling order, preserving
  series identities, explicit styles and data.
- Measured automatic legend/label placement preserves supported names and keys.
  Left and right radar legends render when their entries fit inside the chart frame.
  Non-overlay legends need a separate band, and an authored plot stays in place.
  Per-entry deletions apply to series legends on single-plot bar and column
  charts when the authored indices, series order and legend slots agree.
  Automatic right bar legends reserve plot space when their names would cross the
  physical page and the chart itself fits. Authored manual, overlay and off-page
  placement retain their constraints and clipping.
- Font metrics, text overflow, automatic chart scales, point colors, dense/manual
  charts, rich label wrapping and exact Office geometry remain approximate.
- Source text clips follow shape placement when the transformed frame remains an
  exact rectangle. General oblique text crops retain approximate placement.
- Stock/surface/3D charts, SmartArt, OLE, audio/video, animation and transitions
  are unsupported and may be omitted with diagnostics.

## Images, fonts and remaining correctness gaps

PNG, supported JPEG/BMP and a bounded SVG subset share managed rendering. Ordinary
SVG text can become glyph outlines. Complex SVG text, filters, masks/blends and
unsupported constructs retain limitations. TrueType outlines embed; CFF faces
require an embeddable fallback. Missing glyphs use diagnosed replacement.
Local SVG gradient templates reuse the authored stops and their source colors.
Missing or wrong-type local paint references use the authored fallback color or `none`.
SVG link wrappers preserve transforms, numeric opacity attributes, and inline
zero opacity on their shapes and text. Embedded SVG link actions retain an
unsupported-content diagnostic.
SVG paint accepts opaque numeric `rgb()` colors and separate opacity properties.
Supported inline style declarations honor `!important` precedence, including
zero opacity on SVG roots and groups.
Basic shape geometry, stroke widths, dash lengths and dash offsets accept viewport percentages.
Line elements use stroke paint and ignore fill declarations.
Object opacity can compose a visible fill with a supported linear-gradient stroke.
Translucent solid strokes retain the fill beneath them.
Rectangles and ellipses resolve an automatic radius from the other radius.
SVG paths accept absolute and relative cubic and quadratic Bézier commands,
including smooth forms. An unsupported command reports a warning and stops the
path at that command. A solid stroke can retain the preceding supported segments.

PPTX picture opacity composes with supported SVG opacity. Usable authored raster
previews retain precedence. Unsupported SVG picture recoloring retains original
colors with `PPTX_UNSUPPORTED_IMAGE_RECOLOR` / Original image; managed raster
recoloring remains available.

Known consequential or unresolved cases include partially clipped compressed
standalone cell equations; severe displacement on some custom-paper Review pages;
constrained exact-height/overflowing picture and table crops; and unsupported note
fragmentation. Retained long/top/manual legend geometry failures also remain.
Whole-page visual gates still fail for some readable DOCX documents. These are
not waived by extraction success or green unit tests.

Office application-view calibration is unfinished. Cached Office exports,
source-semantic checks, native paint inspection and whole-page image comparisons
provide different evidence. A passing image score must not hide missing words,
labels or pictures.
