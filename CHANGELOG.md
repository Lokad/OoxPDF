# Changelog

## 0.3.0

- Targets .NET 10 and .NET 11, with safe managed code and no runtime dependencies.
- DOCX: improves mixed text/equation/picture layout, table flow, notes, pagination,
  fields, review markup, numbering, static stories and navigation; fixes lost
  content, clipping and overlap in supported combinations.
- PPTX: improves chart data, axes, labels, legends, series order, merged tables,
  theme colors, percentage values, opacity, font fallback and text/link geometry.
  SVG pictures preserve authored fallbacks
  and supported opacity; text renders as glyph outlines and unsupported recoloring
  reports its fallback.
- Hosting: awaits I/O, uses immutable font leases, explicit stream staging and
  optional shared-memory budgets. Custom resolvers define the complete font universe.
- PDF text preserves distinct Unicode characters that share a glyph outline.
  Plain space and NBSP still normalize to plain space when they share a glyph.
- Breaking API changes require recompilation; see [migration](docs/Migration.md).

See [rendering scope and limitations](docs/Capabilities.md) for approximate
features, known correctness gaps and the status of Office application-view
calibration.

## 0.2.0 — 2026-10-29

- Improves DOCX equations, tables, notes, static stories, pagination and review markup.
- Corrects PPTX theme/master inheritance, tables, charts, text clipping and links.
- Improves font identity, image decoding and diagnostics for visible content.
- Simplifies caches and PDF validation; retains cancellation, input bounds,
  bounded staging and atomic output. Removes ineffective options and cumulative
  work quotas in favor of cancellation and optional host admission.
- Breaking API removals require recompilation; see [migration](docs/Migration.md).

## 0.1.5 — 2026-09-19

- Fixes font subsets, forward-only output, run-dense DOCX performance, image/input
  handling and atomic publication; improves DOCX/PPTX layout and chart rendering.
- Adds Windows/Ubuntu CI and installed-package checks.

## 0.1.4 — 2026-07-02

- Adds HTTP(S) font packs for portable font resolution.

## 0.1.3 — 2026-06-15

- Adds DOCX review views, comments/revisions and markup geometry options.

## 0.1.2 — 2026-06-09

- Adds stream conversion overloads.

## 0.1.1 — 2026-06-06

- Adds TrueType subsetting, Unicode maps and packaged release notes.

## 0.1.0 — 2026-06-05

- Initial dependency-free DOCX/PPTX to PDF library and CLI.
