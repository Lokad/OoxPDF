# Capabilities

This document tracks implemented rendering behavior. Anything not listed as supported or partial should be treated as unsupported unless a diagnostic says otherwise.

## PPTX

Supported:

- Slide discovery, slide order, and slide size.
- Blank slide pages with the corresponding PDF media box.
- Solid slide backgrounds.
- Solid-fill rectangles, ellipses, and straight lines.
- Basic shape rotation and horizontal or vertical flips.
- Slide master and slide layout backgrounds and visible shapes for common inherited cases.
- Theme color resolution for common scheme colors.
- Theme Latin font resolution when text runs request theme fonts.
- JPEG image passthrough as PDF `/DCTDecode` image XObjects.
- PNG image embedding for RGB and RGBA images, including alpha masks.
- Picture relationships, placement, sizing, and basic crop clipping.
- Grouped shapes with nested translation and scaling.
- Fixed-grid tables with cell fills, black grid borders, and cell text.

Partial or approximated:

- Text boxes support paragraphs, runs, font size, color, bold, italic, underline, and left, center, or right alignment, with simple Latin greedy wrapping.
- Bold and italic are PDF drawing approximations and do not yet select matching font faces.
- Table rendering does not yet honor merged cells, per-edge border styles, vertical alignment, or rich table styles.
- Shape rendering supports only a small preset geometry set.

Unsupported or ignored:

- Charts, SmartArt, videos, audio, OLE objects, transitions, animations, macros, and ActiveX.
- Complex effects such as shadows, gradients, transparency, 3D, and most custom geometry.
- Complex scripts, bidirectional text, text shaping, fallback font selection, and OpenType layout features.

Charts, SmartArt, videos, audio, OLE objects, transitions, and animations produce stable warning diagnostics when detected on slides.

## DOCX

Supported:

- Document package discovery.
- Basic page size extraction.
- Page margin extraction.
- Basic body paragraphs and text runs.
- Document defaults, paragraph styles, and character styles for common run and paragraph properties.
- Simple paragraph page breaking.
- Simple decimal and bullet list labels from `numbering.xml`.
- Inline JPEG and PNG images.
- Fixed-width tables with cell fills, black borders, and cell text.
- Default headers and footers with simple text and `PAGE` field approximation.

Partial or approximated:

- Paragraph text supports font size, color, bold, italic, underline, left, center, and right alignment, spacing before/after, and simple Latin greedy wrapping.
- Bold and italic are PDF drawing approximations and do not yet select matching font faces.
- List indentation, hanging indents, tab stops, and advanced numbering formats are approximate.
- Inline images are rendered as block-level content at the paragraph cursor; surrounding text flow is approximate.
- Table merges, cell margins, table styles, and per-cell text formatting are not yet preserved.
- Header/footer distance, odd/even variants, first-page variants, and complex fields are approximate or unsupported.

Unsupported or ignored:

- Tracked-change rendering as revision markup.
- Floating image wrapping, anchored text boxes, charts, SmartArt, equations, live OLE content, comments, footnote/endnote bodies, multi-column layout, and macros.
- Section variants beyond the simple page setup used by the current renderer.
- Complex scripts, bidirectional text, text shaping, fallback font selection, and OpenType layout features.

Comments, tracked changes, complex fields, equations, OLE objects, floating drawings, footnotes, endnotes, multi-column sections, and macros produce stable warning diagnostics when detected.

## PDF Output

Supported:

- Deterministic object ordering and resource naming for stable inputs.
- Embedded TrueType/CID fonts with ToUnicode maps for searchable copied text in supported runs.
- JPEG passthrough and PNG RGB/RGBA embedding.
- Basic path, fill, stroke, transform, clipping, and text operators.

Partial or approximated:

- PDF metadata is intentionally minimal in deterministic mode.
- Font embedding currently favors whole-font simplicity over compact subsetting.

Unsupported or ignored:

- Interactive PDF features such as forms, annotations, outlines, links, tagged PDF structure, video, audio, and JavaScript.
