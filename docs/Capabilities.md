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
- Fixed-grid tables with cell fills, explicit grid borders, merged-cell continuations (horizontal merges and row spans), vertical anchoring approximations, and cell text, plus first-pass built-in table styles (light/medium/dark header, banding, first/last row/column accents).
- Native chart rendering for bar/column, line, area, pie/doughnut, scatter, bubble, and radar charts with cached numeric values, including titles, legends, category/value axes, tick labels, and first-pass data labels. Tier-1 (clustered/stacked bar/column, line with markers, plain pie) is the close-parity target; Tier-2 (area, scatter/bubble, doughnut/radar, secondary axes, leader-line labels, percent-stacked, trendlines) remains approximate.

Partial or approximated:

- Text boxes support paragraphs, runs, font size, color, bold, italic, underline, and left, center, or right alignment, with simple Latin greedy wrapping.
- Bold/italic use a hybrid: the resolver prefers a real font face on exact non-fallback matches; otherwise bold is synthesized with a stroked second pass and italic with an oblique shear (agreed PLAN.md policy).
- Table rendering honors merges and explicit borders with vertical-anchor approximations, but per-edge border styles, rich table styles beyond the first-pass built-ins, and fine vertical metrics remain approximate.
- Shape rendering supports only a small preset geometry set.

Unsupported or ignored:

- Stock, surface, and 3D charts, SmartArt, videos, audio, OLE objects, transitions, animations, macros, and ActiveX.
- Complex effects such as shadows, gradients, transparency, 3D, and most custom geometry.
- Complex scripts, bidirectional text, text shaping, fallback font selection, and OpenType layout features.

Unsupported chart kinds, SmartArt, videos, audio, OLE objects, transitions, and animations produce stable warning diagnostics when detected on slides. PPTX_UNSUPPORTED_CHART is now reserved for unsupported kinds, missing parts, formula-only data without cached values, and unrendered default axis-title layouts (see Diagnostics.md).

## DOCX

Supported:

- Document package discovery.
- Basic page size extraction.
- Page margin extraction.
- Basic body paragraphs and text runs.
- Document defaults, paragraph styles, and character styles for common run and paragraph properties, including paragraph spacing inheritance, contextual spacing, line-unit spacing, automatic before/after spacing, and exact/at-least line heights.
- Simple paragraph page breaking.
- Simple decimal and bullet list labels from `numbering.xml`, including twip-based left, right, first-line, hanging, and numbering-tab indentation.
- Inline JPEG and PNG images.
- Anchored floating images and text boxes with page/margin/column/paragraph positioning, z-order, behind-document rendering, clipping, and first-pass wrap exclusion geometry.
- Fixed-width tables with cell fills, common collapsed border styles, conflict resolution, nil/none suppression, inside/outer table borders, repeated headers, split rows, and cell text.
- Default headers and footers with simple text and `PAGE` field approximation.
- DOCX markup mode selection for final, original, simple markup, and all markup views.
- Simple fields and complex fields with cached results, including nested cached-result fields and cached cross-references inside hyperlinks.

Partial or approximated:

- Paragraph text supports font size, color, bold, italic, underline, left, center, and right alignment, spacing before/after, and simple Latin greedy wrapping.
- Bold/italic use a hybrid: the resolver prefers a real font face on exact non-fallback matches; otherwise bold is synthesized with a stroked second pass and italic with an oblique shear (agreed PLAN.md policy).
- Advanced numbering formats, character-unit list indents, and complex bidirectional list layout are approximate.
- Inline images are rendered as block-level content at the paragraph cursor; surrounding text flow is approximate.
- Floating drawing wrap effects on nearby body text are still approximate even when anchor placement and exclusion geometry are inspected.
- Some decorative table border styles, table merges, cell margins, table styles, and per-cell text formatting are not yet preserved.
- Header/footer distance, odd/even variants, first-page variants, and dynamic field evaluation are approximate or unsupported.
- `Final` and `Original` DOCX markup modes filter inserted/deleted and moved content before layout.
- `SimpleMarkup` renders final text with page-margin change bars and compact comment markers.
- `AllMarkup` renders inline revision styling plus first-pass comment and tracked-change balloons with metadata or revision-kind summaries, preview text, table/image fallback markers, and connectors.
- Markup modes keep the authored PDF media box and text-column geometry; no expanded review-pane margin is created.

Unsupported or ignored:

- Full Word-style tracked-change balloon content and collision-aware markup margin pagination.
- Floating charts, SmartArt, equations, live OLE content, footnote/endnote bodies, multi-column layout, macros, and full Word-style text reflow around floating objects.
- Section variants beyond the simple page setup used by the current renderer.
- Complex scripts, bidirectional text, text shaping, fallback font selection, and OpenType layout features.

Comments, tracked changes, formatting revisions, complex fields, equations, OLE objects, floating drawings, footnotes, endnotes, multi-column sections, and macros produce stable warning or approximation diagnostics when detected.

## PDF Output

Supported:

- Deterministic object ordering and resource naming for stable inputs.
- Embedded TrueType/CID fonts with ToUnicode maps for searchable copied text in supported runs.
- JPEG passthrough and PNG RGB/RGBA embedding.
- Basic path, fill, stroke, transform, clipping, and text operators.

Partial or approximated:

- PDF metadata is intentionally minimal in deterministic mode.
- Embedded TrueType fonts are subset to the glyphs, metrics, and Unicode map entries used by the rendered document (see CHANGELOG 0.1.1); whole-font embedding is no longer the default.

Unsupported or ignored:

- Interactive PDF features such as forms, annotations, outlines, links, tagged PDF structure, video, audio, and JavaScript.
