# Changelog

## 0.3.0 — Unreleased

- Awaits package, staging and PDF output I/O; synchronous calls share that path.
  Font acquisition and CPU rendering remain synchronous.
- Adds host-selected temporary storage. Embedded streams no longer create implicit
  temp files beyond their resident window; file/CLI convenience keeps file backing.
- DOCX: sizes unconstrained automatic tables from content and cell insets,
  preserving fixed/preferred widths. Arithmetic roundoff no longer creates
  spurious text line breaks at an exactly fitting boundary.
- DOCX: reports omitted external images, malformed floating drawings and unsupported
  VML in selected headers/footers using the owning story's relationships.
- Multi-targets .NET 10 and .NET 11. Development uses the pinned .NET 11 RC SDK;
  qualification records framework-specific library and tool identities.
- PPTX: honors series and point negative-bar inversion flags across clustered and
  stacked bar/column charts, including label swatches and explicit point fills.
- Replaces fitted bar-color regimes with shared cycle shading, including
  horizontal charts and theme changes.
- Restores category legends for varying-color bar charts and reserves space for
  wrapped legend rows; crowded frames report clipped entries.
- PPTX: preserves cells following horizontal table merges, sharing source grid
  positions across styles, row-height measurement, text and link placement.
- DOCX: preserves plain comment paragraph breaks beyond five paragraphs, using
  shared measured rows for balloon height and painting.
- DOCX: minimum-line placement uses measured local slots across fonts, styles
  and neighboring tables/wrapped text; links and background paint retain geometry.
- DOCX: continuation pages use their own body print origins, keeping text,
  bookmarks, links and following endnotes independent of page-one font sizes.
- DOCX: resolves table bands, edge styles and corner styles in one precedence
  order for both explicit conditional flags and positional styling, including
  merged edges and single-row/column tables. Unset or zero band sizes suppress
  band formatting, including zero overrides of an inherited positive size.
- DOCX: applies bitmask-only table-look declarations and explicit-flag precedence;
  omitted flags in an explicit declaration no longer acquire unrelated defaults.
- DOCX: continues note lists in document reference order; unused notes and
  reordered notes XML no longer change visible list labels.
- DOCX: places resolved custom-style padding inside the table grid, preserving
  explicit indentation and the existing TableNormal anchor. Table/cell formatting
  inherits earlier parents; empty and text-only styles retain forward references.
- DOCX: recognizes intrinsic Normal Table styles by name and custom-style metadata,
  applying their built-in padding without authored paint or inheritance. Custom
  styles, descendant formatting and direct overrides keep their own properties.
- PPTX: retains readable advances and run/chart hyperlinks when standard PDF
  fonts replace unavailable fonts, including cell text and transformed links.
- Preserves authored chart titles when automatic-title deletion is enabled;
  automatic bar layouts reserve a measured title line inside the frame, and
  inferred titles resolve built-in defaults and authored text sizes before layout.
- Reads common Strict DOCX content and measure forms, including composed tables,
  images, equations and related stories. Strict dialect warnings remain; this
  does not add general Strict OOXML or Strict PPTX support.

## 0.2.0 — 2026-10-29

### Rendering

- DOCX: improves equations combined with text, bookmarks and pictures; table,
  header/footer and note placement; pagination, fields and review markup.
- PPTX: corrects inherited themes/backgrounds, master visibility, bullets and
  table formatting; improves cached-data charts, text clipping and hyperlinks.
- Fonts and images: preserves font identity and language-specific font choices,
  corrects PNG transparency and CMYK JPEG handling, and recovers malformed images
  through diagnostics.
- Diagnostics follow visible content and selected document stories, avoiding
  false warnings from hidden objects and unused headers or footers.

### Hosting and compatibility

- Simplifies bounded font caches and PDF validation while retaining cooperative
  cancellation, input/output bounds, bounded page staging and atomic file output.
- Removes `OoxPdfOptions.Deterministic` and library `Strict`; output remains
  deterministic and CLI `--strict` retains its diagnostic exit policy.
- Removes cumulative work quotas and nonrefunding memory ceilings. Hosts can use
  cancellation and optional coarse shared-memory admission instead.

**Breaking changes require recompiling callers.** See the
[migration guide](docs/PublicApiChanges.md) for removed APIs and hosting contracts.
Rendering remains approximate in documented areas; unsupported content may be
simplified or omitted. See [scope and limitations](docs/PreviewRelease.md) and
[qualification evidence](docs/PreviewQualification.md).

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
