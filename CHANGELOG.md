# Changelog

## Unreleased

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
