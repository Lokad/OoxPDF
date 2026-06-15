# DOCX Markup Fidelity Triage

Use this checklist when reviewing cached Office-reference comparisons for DOCX markup cases. Keep private artifacts under ignored directories and commit only anonymized findings.

## Inputs

- [ ] Confirm the reference came from a trusted Office export produced on a setup that was already proven headless and non-interactive.
- [ ] Confirm the comparison was run through `tools/CompareCachedDocxMarkupReference.ps1` or `tools/RunDocxMarkupReferenceGate.ps1`.
- [ ] Confirm `summary.json`, `comparison/gate-summary.json`, `comparison/page-delta-summary.json`, `comparison/region-delta-summary.json`, `comparison/raster-region-summary.json`, PDF text/graphics inspections, annotation comparisons, balloon comparisons, geometry summaries, and raster metrics are present.
- [ ] Confirm no private screenshot, page image, author name, filename, or document text is copied into a tracked file.

## Page Priority

- [ ] Start with the highest `PriorityScore` pages in `page-delta-summary.json`.
- [ ] Check `GateFailures` in `summary.json` or `comparison/gate-summary.json` before reading page-level artifacts, including page-flow first/last baseline and body-height-used gates.
- [ ] Use `RasterRegionDeltas` in `summary.json` or `comparison/raster/diff/region-metrics.json` to decide which visual region needs the next focused inspection.
- [ ] Classify whether the first blocking defect is page count, page size, body-frame geometry, body text flow, table layout, drawing layout, markup margin, comment balloons, revision balloons, annotations, or low-level PDF drawing.
- [ ] Record only private-safe counters: page index, metric names, counts, deltas, and anonymized feature class.
- [ ] If a private-only feature interaction caused the delta, create or update a synthetic public fixture that isolates the same interaction.

## Visual Delta Classes

- [ ] Page geometry: media box, body frame, markup lane side/width, margins, section/mirror/landscape behavior.
- [ ] Pagination: first/last source block drift, paragraph/table fragment drift, keep rules, footnote/endnote placement.
- [ ] Text flow: line breaks, baseline drift, font fallback, glyph advance, kerning, tabs, nonbreaking spaces, soft hyphens.
- [ ] Revisions: inline styling, deleted text, moved text, formatting balloons, author color buckets, grouping.
- [ ] Comments: body marker, visible/hidden/orphaned anchor classification, threaded replies, connector routing.
- [ ] Balloons: ordering, rectangle geometry, typography, padding, overflow, continuation, collisions.
- [ ] Tables/lists: grid widths, cell margins, repeated headers, vertical merges, borders, numbering labels.
- [ ] Drawings/fields/links: anchored drawing placement, text boxes, field results, page fields, hyperlink annotation rectangles.
- [ ] PDF primitives: strokes, fills, dashes, joins, caps, z-order, text operation segmentation, font embedding.

## Promotion Rule

- [ ] A renderer change driven by a private case should get a public synthetic fixture before it is considered complete.
- [ ] A public fixture should be tagged by markup mode and subsystem so it can be run independently.
- [ ] A fix should include either an Office-reference report, a private-safe comparison summary, or a focused unit test that proves the specific behavior.
