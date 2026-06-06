# Changelog

## 0.1.1 - 2026-06-06

- Subset embedded TrueType fonts so generated PDFs carry only the glyphs, metrics, and Unicode map entries needed by rendered PPTX and DOCX text.
- Added coverage for subset font output, deterministic font subsetting, and PPTX text emission through subset font resources.
- Added package release notes metadata and packaged this changelog with the NuGet artifact.

## 0.1.0 - 2026-06-05

- Initial public package for dependency-free PPTX and DOCX to PDF conversion.
- Included managed OOXML readers, PDF writing, font embedding, image handling, diagnostics, CLI conversion, and visual validation tooling.
