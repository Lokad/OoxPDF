# Changelog

## 0.3.0 — Unreleased

- Targets .NET 10 and .NET 11; the library remains safe managed code without
  runtime dependencies.
- DOCX: improves mixed text/equation/picture composition, tables, notes, related
  stories, pagination, fields, review markup and navigation. Corrects visible
  content lost through wrapping, clipping or off-page frames. Computed fields
  remain intact beside pictures; automatic note references retain source order and link ownership,
  and Review endnote counters retain neighboring equations. Oversized keep-with-next
  groups flow without repeatedly starting new pages; independent keep-lines rules
  and explicit breaks remain respected. Hidden run breaks and inline text boxes
  honor inherited visibility; annotation-only revision scopes retain neighboring
  equations, including scopes containing hyperlinks.
- PPTX: improves shared chart-data ownership, scales, colors, labels and legends;
  qualified authored horizontal legends retain every name; manual pie viewports
  own their radius and use the measured automatic frame for factor offsets in both axes.
  Side plots reserve measured legend space; default legends retain their chart frame
  when plot geometry is authored. Whole and rich chart labels
  share measured clipping; automatic polar labels avoid measured collisions while
  preserving authored/unfit fallbacks. Dense wrapping remains approximate. Merged-table content,
  font fallback, text/link geometry and preset arcs also improve.
- SVG pictures: ordinary text labels now remain visible in DOCX and PPTX as managed
  glyph outlines; unsupported text features retain explicit diagnostics.
- Hosting: awaits package, font, staging and output I/O; introduces immutable font
  leases, explicit temporary storage and host-funded bounded caches. Shared memory
  admission remains coarse and partial. Custom resolvers define the complete font
  universe. See the [caller migration guide](docs/PublicApiChanges.md) for API changes.
- Comment diagnostics follow the resolved view: intentionally unprinted comments
  no longer warn as unsupported; requested markup retains approximation warnings.

General note-row fragmentation remains unsupported. Complex layout, typography,
automatic/dense chart layout and unqualified combinations remain approximate.
[Capabilities](docs/Capabilities.md) and
[development qualification](docs/DevelopmentQualification.md) record scope and
remaining visual failures. Office-PDF evidence does not establish calibrated
application-view parity. Publication remains deferred.

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
