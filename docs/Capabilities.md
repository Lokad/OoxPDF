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
- Charts, SmartArt, videos, audio, OLE objects, transitions, and animations produce stable warning diagnostics when detected on slides.

## DOCX

Supported:

- Document package discovery.
- Basic page size extraction.
- Blank PDF page output for simple documents.

Unsupported or ignored:

- Paragraph text layout.
- Numbering and lists.
- Inline images.
- Tables.
- Headers and footers.
- Unsupported DOCX features are not yet surfaced through stable diagnostics.
