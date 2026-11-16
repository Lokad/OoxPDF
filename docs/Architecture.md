# Architecture and PDF output

Lokad.OoxPdf is a dependency-free managed pipeline:

1. Read bounded OOXML ZIP parts, content types and relationships.
2. Resolve styles, themes, source identities and supported content into document models.
3. Acquire fonts/images and prepare measured geometry.
4. Produce staged PDF pages, then serialize resources and document objects.

The shipped library uses no unsafe code, unchecked memory-access APIs, native
interop or external processes. Office, PDFium, PyMuPDF and fontTools belong to
separate validation tooling.

## Layout ownership

DOCX uses flowing paragraphs, tables, sections and related stories. Ordered inline
content preserves source ownership through measurement, wrapping and paint.
Supported cell content shares a source-order placement cursor and bounded memoized
measurement. Containers own available frames, page admission, alignment and clips;
painting and navigation consume resulting geometry. Authored exact heights and
physical page edges constrain content rather than implying automatic expansion.

PPTX retains absolute slide placement. Shape/group transforms, scene-resolved
chart axes/series, theme properties and measured labels feed its painters. It
shares fonts, image handling and PDF primitives with DOCX, not its pagination engine.
Some rendering policies remain qualified to narrower source shapes; see
[capabilities](Capabilities.md). Models still retain XML for properties not fully
resolved into the scene.

Font programs are immutable and acquired through owned leases. Cache retention,
source acquisition and conversion ownership have distinct lifetimes. Bounded
caches synchronize publication/eviction; source leases survive cache eviction and
shutdown until their readers finish. Models, resources and staging dispose their
funding after final use. [Hosting](Hosting.md) defines the public contracts.

## PDF profile

The writer emits PDF 1.7, numbered objects, a conventional cross-reference table,
trailer and page tree. It supports text, paths, clipping, raster images/soft masks,
transparency groups, shadings/patterns and Link annotations. URI actions and explicit
internal destinations implement supported navigation.

Embedded TrueType fonts use Type0/Identity-H and CIDFontType2 with subset FontFile2,
widths and ToUnicode mappings. Distinct source characters retain separate PDF
character identities while sharing embedded outlines through a CID-to-glyph map.
Shared spaces still normalize to plain space when mixed with NBSP. Structural
equation pieces can use ActualText. Unembedded standard Type1 fallback depends on viewer fonts. CFF
outlines do not embed as FontFile3 and require a supported fallback.

Resource references, encoding and balanced state/nesting are validated; writer
tests separately check object structure, stream lengths and cross-reference offsets.
Independent PDF inspection and native raster tools provide additional evidence.
ToUnicode/ActualText do not establish accessibility conformance.

The output is not tagged PDF, PDF/A or PDF/UA. No encryption, linearization, object
streams, AcroForms, JavaScript, source attachments, document outlines or XMP metadata
are emitted. The generated subset is distinct from the validation inspector's
partial parsing coverage. Syntax, extraction, visible completeness and Office
appearance each need appropriate checks; one cannot substitute for the others.
