# Supported PDF profile

This document defines the PDF this library emits, so compatibility claims
stay checkable without implying PDF/A or PDF/UA conformance. Every statement
below traces to `PdfDocumentWriter`, `PdfObjectWriter`, `PdfGraphicsBuilder`,
or `PdfEmbeddedFont` in `src/Lokad.OoxPdf/Pdf`, unless noted otherwise.

## Document structure

- Header `%PDF-1.7`, body objects, plain cross-reference table,
`trailer<</Size /Root>>`, `startxref`, `%%EOF`.
- Object graph: catalog, page tree, one page plus one content object per
page, five objects per embedded font (Type0, CIDFont, descriptor, font
program, ToUnicode), image and soft-mask image objects, shadings, tiling
patterns, and link annotations. Object numbers are deterministic.
- No object streams, no linearization, no encryption, no outlines, no name
tree, no embedded files, no forms, no JavaScript.
- Document information dictionary exists only when
`OoxPdfOptions.FixedCreationDate` is set (Producer, CreationDate, ModDate);
otherwise the trailer carries no `/Info` and output stays deterministic.

## Content streams

- Page, pattern, and luminosity soft-mask contents are uncompressed ASCII,
written inline with explicit `/Length`.
- Text uses `Tj`/`TJ` with `Tm`/`Tf`/`Tc` positioning; glyphs address a
`/Identity-H` CID space (see Fonts).
- Images arrive as `/DCTDecode` passthrough for baseline JPEG or
`/FlateDecode` for re-encoded raster (PNG/BMP/recolored JPEG) and grayscale
soft masks. Color spaces are device gray, RGB, or CMYK only.
- Transparency uses ExtGState fill alpha and luminosity soft masks.
- The full operator inventory is whatever `PdfInspect` parses; anything it
cannot tokenize is outside the supported profile by construction.

## Fonts

- Every embedded font is Type0 with `/Identity-H`, a CIDFontType2
descendant, `/CIDToGIDMap /Identity`, a Widths array, a descriptor with
TrueType metrics, a Flate-compressed FontFile2 subset, and a ToUnicode
CMap. Subsets cover exactly the used glyphs.
- Only TrueType outlines embed. CFF/OpenType-CFF faces fall back to an
embeddable typeface with `FONT_UNSUPPORTED_OUTLINES` (see Diagnostics);
they are never written as FontFile3/CIDFontType0.
- ToUnicode maps preserve source scalars per emitted CID. Glyphs shared by several source scalars resolve to one scalar: spaces win over NBSP (matching Word extraction; pinned by SharedSpaceGlyphExtractsAsPlainSpace and PdfEmbeddedFontMapsSharedSpaceGlyphToPlainSpace). Private-use scalars pass through (pinned by the docx-symbols text-content gate).
- Unbroken soft hyphens never reach the font layer: DOCX layout strips them before emission (pinned by DocxParagraphLayoutSuppressesUnbrokenSoftHyphens).
Extraction fidelity is gated by `tools/CheckPdfTextContent.ps1`.

## Annotations and metadata

- The only annotation is `/Link`: URI actions for external hyperlinks and
explicit XYZ destinations for internal ones.
- No document outlines, page labels, language entries, structure tree,
viewer preferences, or XMP metadata are emitted.

## What establishes what

- Syntax and object integrity: the writer plus `PdfContentValidator`
(rejects dangling resources, unbalanced state, non-ASCII content) and the
`PdfWriterTests` structural checks (xref offsets, stream lengths).
- Independent parse: `tools/Lokad.OoxPdf.PdfInspect` re-parses objects,
streams, fonts, and text with project-external logic (no shared writer
code paths).
- Extraction: ToUnicode maps plus the text-content gate above.
- Visual fidelity: PDFium rasterization with VisualDiff similarity and
locked visual cases (see VisualValidation). Pixel similarity alone never
establishes the rows above.
