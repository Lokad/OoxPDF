# Capabilities

This document tracks implemented rendering behavior. Anything not listed as supported or partial should be treated as unsupported unless a diagnostic says otherwise.

## PPTX

Supported:

- Slide discovery, slide order, and slide size.
- Hidden slides (`show="0"`) are excluded from export, matching PowerPoint; per-slide master-shape suppression (`showMasterSp`) is honored.
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
- Fixed-grid tables with cell fills, explicit grid borders, merged-cell continuations (horizontal merges and row spans), vertical anchoring approximations, and cell text, plus calibrated built-in table styles (MediumStyle2 linear-light band tints at 0.60/0.80, DarkStyle1 black header with 0.60 linear-shade bands and raw unbanded rows, LightStyle1 raw accent bands; first/last row/column accents).
- Clickable hyperlink annotations for shape, picture, connector, group, table, chart, and unknown-frame clicks plus shape and table body-text runs and chart title, axis-title, and data-label runs, resolving external URLs and internal slide targets with transformed bounds.
- Native chart rendering for bar/column, line, area, pie/doughnut, scatter, bubble, and radar charts with cached numeric values, including titles (single-series charts gain the series name as an auto title), legends, category/value axes on any side, tick labels, measured horizontal-bar label strips and outer-manual plot reserves, and first-pass data labels. Major tick marks on bar/column, line, area, scatter, and bubble charts scale with tick-label size (Office-calibrated 0.315x) and honor explicit majorTickMark settings; unstyled chart series/gridline/marker/key strokes default to round caps/joins; vertical single-series vary-colors bar/column fills follow per-count variation rows below 96 points and the Office-calibrated linear-light recipe over theme-live accents at 96 or more points. Bar/column plot clips extend 0.69pt past the axis-bounded plot rect on the bottom and right; unstyled line, scatter, and radar series strokes plus legend key lines and glyphs render the theme accent darkened to 0.975x luminance while fills keep raw colors, except gallery line styles and smoothed series which stroke the raw base: effective style-18 series at 5pt and style-26 at 7pt with end-chord Hermite smoothing (gap-split ends eased at sixths, span joints at symmetric mean-thirds with span starts at sixths and full chart ends, zero-mode blanks plotting contiguously at zero), style-18 forcing auto diamond/square/triangle markers at 12.96pt with per-marker axial gradient fills from the shared bar stop table and raw 1pt rims (explicit symbols, nominal-5 sizeless defaults, sizes, circle gradients, quantized rect dots and dash bars, stroked asterisks, and explicit marker fills/1pt outlines honored; explicit line colors win with gallery-width inheritance when width is unspecified; series-level symbol=none wins; style 26 draws no markers), explicit smooth series at raw 3pt with forced flat 9pt markers plot-wide (explicit smooth=false keeps straight legacy; plot-level smooth=0 yields to explicit series smooth), and gallery legend keys matching with raw series-width swatches and forced 9.9 gradient markers (explicit marker fills win in keys); unstyled axes and ticks render black, except single-plot bar/column charts carrying gallery style 2 without a style part, whose unstyled axis-family strokes render gray at width 1. Unstyled axis titles render bold black. Tier-1 (clustered/stacked bar/column, line with markers, plain pie) is the close-parity target; Tier-2 (area, scatter/bubble, secondary axes, leader-line labels, percent-stacked, trendlines) remains approximate; filled-radar series under effective style 18 paint Office-derived per-channel five-stop vertical gradients (rank-mapped endpoints with span-law interiors); doughnut ports hold no open tripwires.
- Chart number formats: sign and conditional sections, 1900/1904 dates and datetimes, scientific and fraction rendering, accounting skip/fill runs, scaling commas, quoted/escaped/bare literal runs, locale currency symbols, and source-linked workbook formats including dates; axes and data labels share one formatter.
- Markup-compatibility Choice/Fallback selection renders exactly one AlternateContent representation (first understood Choice, else Fallback).
- Failed slide nodes rewind partial paint and annotations, then report PPTX_NODE_RENDER_FAILED while neighboring nodes render normally.

Partial or approximated:

- Plain ASCII distributed PowerPoint paragraphs retain nominal text advances, including active kerning, in horizontal single-column, unrotated/unflipped non-table frames with explicit noAutofit and wrap=none, no bullets, tabs or manual breaks. Other scripts, wrapping, autofit, column, rotation, table and bullet paths retain their existing distribution approximation.
- PowerPoint automatic numbering supports the existing Arabic, alphabetic and Roman period/right-parenthesis forms and Arabic paired parentheses outside tables, with valid levels/starts and default or overridden bullet font/size. Equal effective format/start settings continue a level; changed settings and plain paragraphs reset it, and a parent advances deeper levels to a new sequence. Empty numbered paragraphs update settings without emitting or advancing a label; run-free plain paragraphs clear the qualified sequence. Alphabetic labels follow Office’s repeated-letter cycle (a..z, aa..zz, up to thirty repeated letters before cycling). Qualified numbering keeps paragraph pitch at the body font size when labels are larger. Bounded labels follow the first body run’s font, preserve authored size and use nominal advances including trailing character spacing for first-fragment clearance in left horizontal single-column, unrotated/unflipped non-table frames without tabs or normal autofit, with nonpositive hanging indents and a label that fits. Font/size overrides additionally require room for the first drawable body fragment. Continuation lines retain the authored margin. For admitted wrapping with explicit noAutofit, printable ASCII text runs, no fields/manual breaks and room for the first body fragment, word fitting uses the authored line edge instead of the general bullet overrun allowance. Excluded frames, insufficient space, unsupported formats/starts, character bullets and table numbering retain their existing font/kerning/placement fallback; wrapping, manual-break baselines and broader font metrics remain approximate.
- Text boxes support paragraphs, runs, font size, color, bold, italic, underline, and left, center, or right alignment, with simple Latin greedy wrapping. For plain left-aligned, non-bullet shape paragraphs in unrotated horizontal single-column no-wrap/noAutofit frames, explicit or inherited zero kerning disables pair adjustments in both measurement and glyph emission. Positive thresholds and absent settings keep their rules; excluded frames retain their kerning approximation. Emergency splitting of unbroken first-token text in horizontal two/three-column frames uses the column edge. Single/four-plus columns and rotated/no-wrap paths retain their prior approximations; general font metrics, column balancing and autofit remain partial.
- Bold/italic use a hybrid: the resolver prefers a real font face on exact non-fallback matches; otherwise bold is synthesized with a stroked second pass and italic with an oblique shear (hybrid policy: a real face wins on exact non-fallback matches, otherwise bold is stroked and italic is sheared).
- Table rendering honors merges and explicit borders with vertical-anchor approximations, but per-edge border styles, rich table styles beyond the first-pass built-ins, and fine vertical metrics remain approximate.
- Shape rendering supports only a small preset geometry set.
- SVG linear gradient strokes use native PDF shading patterns for uniformly numeric linear-pad stops, forward axes and unrotated/unflipped pictures. Admission requires 2..256 usable stops, represented intervals at least 0.001, an identity gradient transform, identity source-unit or positive axis box-path transforms, and serialized projection error at most 0.001. Plain numeric stop/node/stroke alpha multiplies once; supported root/group isolation retains local pattern resources. Fill masks restore before the stroke and later solid paint. Scaled box-path color/width and dash phase remain approximate. Radial, varying-alpha, reverse/repeated/reflected or excluded transformed strokes, unsupported opacity syntax/depth, nested SVG viewports and filled elements with partial node opacity retain diagnosed omission.
- Rasterized outer shadows and glows, axial-gradient chart style fills and shape gradients, SVG gradients, and alpha transparency; other effects remain unsupported. SVG radial fills with distinct representable stop intervals use smooth PDF shading, including bounded repeat/reflect spread through the full clipped shape. Uniform/opaque and excluded linear gradients, hard/near-stop radial fills and dense repetition retain sampled approximations. Native repetition is capped at 128 cycles and 256 expanded normalized color stops. Fill-only radial transparency follows the user-selected smooth PowerPoint slide preview; its PDF export has different circular transparency artifacts. Uniform numeric stop-opacity attributes multiply fill/node alpha; numeric zero-opacity root/group containers do not paint. Numeric partial opacity on the outer SVG and group containers composites overlapping supported children once through isolated PDF Forms, with at most 32 nested groups. SVG picture viewport clips follow picture rotation/flip transforms, preserving the rotated footprint for the qualified opaque/isolated, cropped and grouped controls. Partial CSS/percentage container alpha and nested SVG viewport alpha retain diagnosed fallbacks. Bounded numeric varying stop alpha on radial pad fills uses a vector luminosity mask and multiplies fill/node alpha. Admission requires 2..256 usable stops, representable native radii/intervals, an identity gradient transform, no off-center focus, an identity user-space path transform and a total Form/mask depth of at most 32. Bounded numeric varying linear-pad stop alpha uses native axial color and a vector mask. It preserves normalized object-box/user-unit projection under viewport stretch. Admission requires forward vector components, an identity gradient transform, identity user-space or positive axis box path transforms, the same stop/count/depth bounds and serialized corner projection error at most 0.001. Reversed/excluded linear, sampled radial, repeated/reflected, transformed/focal and depth-excluded varying alpha, CSS/percentage stop opacity and path-level fill/stroke compositing retain their prior fallbacks. Shapes inside SVG definitions do not paint directly; supported gradient resources remain available through references. SVG strokes preserve directional widths, caps and dashes under picture viewport stretching, including combined solid fill/stroke paint where the PDF transform survives number precision; extreme aspect ratios retain scalar stroke approximation. Element transforms use the Office uniform stroke-width rule (largest singular value), including shear. Dash offsets follow Office stroke-width units; dash-array lengths remain source geometry. Solid square-cap dash lengths compensate the cap extent when every normalized painted segment exceeds the stroke width. Gradient strokes retain Office cap extension into their gaps. Qualified opaque round-dashed strokes on separate straight subpaths use flat interior ends and outward semicircles when a visible dash end lies within half a stroke width of an original endpoint. Opaque straight round-dashed paths that fall entirely in gaps use Office solid rounded fallback; a painted subpath suppresses fallback for the whole path. Single visible dashes use this rule when at least one end falls outside the endpoint region; matching native rounded dashes retain their original paint. Equal/short/zero square dashes, other round-dash paths, standard SVG affine stroke outlines, dash-offset semantics and path-level opacity compositing remain approximate.
- Missing-glyph and missing-font fallback renders deterministic standard-14 faces with capped diagnostics (WinAnsi coverage; other scripts substitute a diagnosed mark).
- Chart number formats keep invariant separators and English names (locale-specific rules warn instead), and axis labels use chart-side formats rather than linked workbook formats.

Unsupported or ignored:

- Stock, surface, and 3D charts, SmartArt, videos, audio, OLE objects, transitions, animations, macros, and ActiveX.
- Complex effects such as 3D, most custom geometry, and effects beyond the supported subset below.
- Strict OOXML (ISO 29500) content: only the transitional dialect is supported; strict parts warn OOXML_STRICT_DIALECT and may render missing.
- Complex scripts, bidirectional text, text shaping, and OpenType layout features beyond coverage fallback.

Unsupported chart kinds, SmartArt, videos, audio, OLE objects, transitions, and animations produce stable warning diagnostics when detected on slides. PPTX_UNSUPPORTED_CHART is now reserved for unsupported kinds, missing parts, and formula-only data without cached values (see Diagnostics.md).

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
  `REF` cross-references retain stored results; bookmark text is not reevaluated.
- Markup-compatibility Choice/Fallback selection renders exactly one AlternateContent representation (first understood Choice, else Fallback).

Partial or approximated:

- Paragraph text supports font size, color, bold, italic, underline, left, center, and right alignment, spacing before/after, and simple Latin greedy wrapping.
- Bold/italic use a hybrid: the resolver prefers a real font face on exact non-fallback matches; otherwise bold is synthesized with a stroked second pass and italic with an oblique shear (hybrid policy: a real face wins on exact non-fallback matches, otherwise bold is stroked and italic is sheared).
- Advanced numbering formats, character-unit list indents, and complex bidirectional list layout are approximate.
- Inline images with recorded run affinity paint mid-line at run position across body, table-cell, related-story, and static header/footer paths (tall images shift the line down keeping the advance below; justified lines add distributed stretch before the image offset), and overwide images overflow past the margin. Office calibration of image baseline and line growth remains approximate.
- Missing-glyph and missing-font fallback renders deterministic standard-14 faces with capped diagnostics, including comment and revision balloon text (word-compatible balloon text still needs embedded faces).
- Floating drawing wrap effects on nearby body text are still approximate even when anchor placement and exclusion geometry are inspected.
- Some decorative table border styles, table merges, cell margins, table styles, and per-cell text formatting are not yet preserved.
- Header/footer distance, odd/even variants, first-page variants, and dynamic field evaluation are approximate or unsupported.
- `Final` and `Original` DOCX markup modes filter inserted/deleted and moved content before layout.
- `SimpleMarkup` renders final text with page-margin change bars and compact comment markers.
- `AllMarkup` renders inline revision styling plus first-pass comment and tracked-change balloons with metadata or revision-kind summaries, preview text, table/image fallback markers, and connectors.
- Footnote and endnote bodies place first-pass as related stories; split-note continuation and complex note flow stay approximate.
- Simple paragraph-only section sequences with equal page geometry use the following section's start type for next-page/continuous boundaries. Actual current-page footnote markers with page-bottom/default placement complete the page before a continuous section; beneath-text placement and endnotes do not force it. Tables, columns, static stories, rich/revision content, changed geometry, terminal section elements and partially placed or overflowing notes retain the prior flow. Raw section metadata remains inspectable.
- Ordinary, unscaled, one-column body paragraphs rendered on one line include direct/inherited before- and after-spacing in note-reference hit areas for automatic/exact line slots. The predecessor owns its after-spacing; only the remaining collapsed before-gap belongs to the note paragraph, including section starts and empty predecessors. For multiline paragraphs, only marks on the first rendered line own a positive before-gap; later-line marks retain their line slot and no multiline mark includes after-spacing. Wrapped/contextual/rich predecessors, contextual/auto/line-unit spacing, tables, rich/static/scaled content and split/manual/page-break flow retain the prior hit-area fallback; glyphs, horizontal bounds and note destinations remain unchanged.
- Explicit false custom-mark flags retain automatic note labels/counters. True custom marks keep authored text and numbering; a single-reference source run links its immediately following first visible BMP character to the rendered note. Empty, whitespace, separate-run, multiple-reference and revision/field/comment marks retain the prior annotation fallback. Source ownership and a visible note destination are required.
- Automatic note labels in simple section sequences support continuous and `eachSect` numbering, section starts/formats and their defaults, including minimum two-digit `decimalZero`. Continuous labels use document occurrence counts with the current section start; section restarts count from the section boundary. Custom marks do not advance automatic numbering, and authored document settings remain inspectable. Each-page numbering, unsupported or malformed settings, revision bodies and hidden/nested section boundaries retain the prior numbering fallback.
- Multi-column sections flow explicit breaks with first-pass exclusion geometry while continuous balancing stays approximate.
- Markup modes keep the authored PDF media box and text-column geometry; no expanded review-pane margin is created.

Unsupported or ignored:

- Full Word-style tracked-change balloon content and collision-aware markup margin pagination.
- Floating charts, SmartArt, equations, live OLE content, macros, and full Word-style text reflow around floating objects.
- Section variants beyond the simple page setup used by the current renderer.
- Strict OOXML (ISO 29500) content: only the transitional dialect is supported; strict parts warn OOXML_STRICT_DIALECT and may render missing.
- Complex scripts, bidirectional text, text shaping, and OpenType layout features beyond coverage fallback.

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
- CFF/OpenType-CFF outlines are never embedded; affected runs fall back to an embeddable typeface with a `FONT_UNSUPPORTED_OUTLINES` warning (see Diagnostics).

Unsupported or ignored:

- Interactive PDF features such as forms, non-link annotations, outlines, tagged PDF structure, video, audio, and JavaScript. Hyperlink annotations (/Link with URI actions and internal destinations) are supported for DOCX and PPTX links.

## Static-content fallback matrix

Each row links an unsupported family to its detection diagnostic, fallback, and covering tests (synthetic in-code fixtures). Severities and dedup scope are defined in Diagnostics.

| Family | Detection | Fallback | Covering tests |
|---|---|---|---|
| PPTX stock, surface, and 3D charts | `PPTX_UNSUPPORTED_CHART` Warning | Ignored | `PptxStockChartKindWarnsUnsupportedChart` (pptx-model) |
| PPTX SmartArt | `PPTX_UNSUPPORTED_SMARTART` Warning | Ignored | `PptxUnsupportedFeaturesEmitDiagnostics` (pptx-core) |
| PPTX OLE objects | `PPTX_UNSUPPORTED_OLE_OBJECT` Warning | Ignored | `PptxUnsupportedFeaturesEmitDiagnostics` (pptx-core) |
| PPTX audio, video, animation, transitions | matching `PPTX_UNSUPPORTED_*` Warning | Ignored | `PptxUnsupportedFeaturesEmitDiagnostics` (pptx-core) |
| PPTX missing image parts | `IMAGE_MISSING_PART` Error | Ignored, conversion continues | `PptxMissingImagePartDiagnosesErrorAndContinues` (pptx-images) |
| Undecodable images | `IMAGE_UNSUPPORTED_FORMAT` Error | Ignored | DocxImagesTests, ImagingTests, PptxImagesTests |
| Cropped-image decode | `IMAGE_CROP_UNSUPPORTED_FORMAT` Warning | PDF clipping | PptxImagesTests |
| SVG pictures | `SVG_UNSUPPORTED_CONTENT` Error/Warning | Ignored / Partial | PptxImagesTests SVG diagnostics |
| DOCX floating charts | `DOCX_UNSUPPORTED_CHART` Warning | Ignored | `DocxUnsupportedChartInHeaderIsDiagnosed` (docx-core) |
| DOCX SmartArt | `DOCX_UNSUPPORTED_SMARTART` Warning | Ignored | `DocxUnsupportedFeaturesEmitDiagnostics` (docx-core) |
| DOCX equations | `DOCX_UNSUPPORTED_EQUATION` Warning | Ignored | `DocxUnsupportedEquationInFooterIsDiagnosed` (docx-core) |
| DOCX OLE objects | `DOCX_UNSUPPORTED_OLE_OBJECT` Warning | Ignored | `DocxUnsupportedFeaturesEmitDiagnostics` (docx-core) |
| DOCX external images | `DOCX_UNSUPPORTED_EXTERNAL_IMAGE` Warning | Ignored | `DocxUnsupportedFeaturesEmitDiagnostics` (docx-core) |
| DOCX macros | `DOCX_UNSUPPORTED_MACRO` Warning | Ignored | `DocxUnsupportedFeaturesEmitDiagnostics` (docx-core) |
| Strict OOXML | `OOXML_STRICT_DIALECT` Warning | May render missing | OoxmlTests dialect probes (ooxml) |
| Font packs | `FONT_PACK_*` throw codes | Conversion fails with DiagnosticId | FontTests pack suites (fonts) |
