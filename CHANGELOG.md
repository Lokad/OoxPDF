# Changelog

## Unreleased

## 0.1.5 - 2026-09-19

- Fixed garbled subset-font text when one typeface is used with different character sets across a document: embedded font resources are now keyed by codepoint set so merged subsets keep valid CID mappings.
- Supported forward-only output streams for stream conversion: PDF cross-reference offsets are now tracked internally instead of reading output Position, so non-seekable host descriptor streams no longer fail with Specified method is not supported.
- Fixed very slow DOCX conversion for run-dense documents: run typeface resolution during text measuring is now indexed instead of scanning all runs per measurement, about 100x faster on a 131-page markup document with byte-identical output.
- Capped XML element count per part during package reading so broad shallow documents fail fast instead of inflating the DOM without bound.
- Improved DOCX pagination, table sizing and borders, headers and footers, and Word-compatible review markup placement.
- Improved PPTX chart layout, labels, legends, number formats, and theme-based varying colors, along with vertical text and text-frame layout.
- Corrected built-in PPTX table styles, including theme text colors and linear-light band, column, and last-row shading.
- Reduced repeated font loading and discovery work, coalesced concurrent font-pack downloads, added transient download retries, and exposed Windows font discovery cache invalidation.
- Hardened package and image parsing, selected a single supported markup-compatibility representation, and added diagnostics for unsupported Strict OOXML and must-understand content.
- Protected existing destination files with atomic replacement after successful conversion. Stream outputs now reject nonempty or nonzero-position seekable streams, and input and output must be distinct.
- Corrected PDF hyperlink escaping for non-ASCII URLs, with bounded stack use even for very long links.
- Added Windows and Ubuntu CI coverage, including package-consumption smoke tests; Ubuntu intentionally runs the portable API, OOXML, PDF, imaging, and font test groups.

## 0.1.4 - 2026-07-02

- Added a built-in HTTP(S) font pack resolver for deterministic packaged font resolution outside local Windows font directories.

## 0.1.3 - 2026-06-15

- Added DOCX markup printing support for comments, insertions, deletions, and move revisions.
- Added DOCX review-mode options for final/original/simple/all-markup output and margin-preserving or expanded comment geometry.
- Added public visual validation coverage and tooling for DOCX markup reference-cache workflows.

## 0.1.2 - 2026-06-09

- Added public stream-based conversion overloads so callers can convert PPTX and DOCX packages to PDF without creating temporary input or output files.

## 0.1.1 - 2026-06-06

- Subset embedded TrueType fonts so generated PDFs carry only the glyphs, metrics, and Unicode map entries needed by rendered PPTX and DOCX text.
- Added coverage for subset font output, deterministic font subsetting, and PPTX text emission through subset font resources.
- Added package release notes metadata and packaged this changelog with the NuGet artifact.

## 0.1.0 - 2026-06-05

- Initial public package for dependency-free PPTX and DOCX to PDF conversion.
- Included managed OOXML readers, PDF writing, font embedding, image handling, diagnostics, CLI conversion, and visual validation tooling.
