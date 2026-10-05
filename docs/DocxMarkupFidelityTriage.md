# DOCX Markup Fidelity Triage

Use this checklist when reviewing cached Office-reference comparisons for DOCX markup cases. Keep private artifacts under ignored directories and commit only anonymized findings.

## Public Reference Snapshot (2026-10-05, RV06-E1)

All 33 public markup manifests have verified cached PDFs at 144 DPI on the
validation workstation, covering 25 distinct input/view identities. Sixteen new
identities closed the 24 case-level misses. Nine public `docx-private-grounded-*`
cases share one complex synthetic fixture and view; they exercise named feature
interactions, not nine independent inputs. This is local cache availability,
not a repository-distributed reference corpus or an Office parity claim.

New references were exported with Word 16.0 under the reference supervisor and
explicit ShowRevisions/RevisionsView/MarkupMode/RevisionsMode/ShowComments flags.
Original mode rejects revisions on a temporary input copy and prints clean text.
All-markup references use inline deletions (RevisionsMode=1). The historical
original-mode reference instead printed comment balloons and shrank the body;
it was backed up and replaced with a clean explicit-view export. Eight other
historical PDFs were preserved. Independent all/final exports confirm their
historical view class; their raster MAE against the old PDFs is below 0.03.
Automatic date fields can change during Word export, so field-result differences
must be distinguished from renderer layout and operation alignment.

The blank-line comparison repair removes nine false gate failures across three
cases, with all 33 candidate/reference PDF pairs unchanged. Hidden-anchor last
baseline drift falls from 24.98 to 0.02 points and its failures from nine to three;
the remaining width/advance deltas of about 2.5 points still fail. Correcting the
original-view reference separately removes eleven failures (27 to 16), without
changing the candidate PDF. Missing or corrupt markup variants now require an
explicit export/import; generic cache rendering cannot silently label Word's
inherited view. Acceptance-test PDFs use unique scratch identities.

All 33 cases still have nonzero parity gates. Page counts match in every case;
22 cases fail the whole-page raster threshold, and region gates catch additional
localized differences. The E1 snapshot tally is 671 failed gates by manifest, or 463
when shared identities are counted once. These are overlapping metrics, not
671 distinct bugs. The comparator repair closes an alignment defect; broader
Word-compatible layout, balloon composition and annotation fidelity remain open.

| Delta class | Failed gates by manifest | Interpretation for the next repair |
|---|---:|---|
| Page geometry | 10 | Media/body-frame geometry signals; page count already agrees |
| Pagination | 64 | Painted baseline/span drift; inspect actual layout and keep/row flow |
| Markup geometry | 43 | Body/lane widths, occupied bounds and connector placement |
| Text | 281 | Baselines, positions, advances and spacing; verify merged-line alignment before changing rendering |
| Graphics | 32 | Visible path-operation differences; confirm with the corresponding pixels |
| Tables | 38 | Grid/column geometry; clean original mode still has a roughly 53-point grid-bound mismatch |
| Annotations | 26 | Rectangle and destination differences; inspect targets independently of visible text |
| Balloons | 27 | Rectangle/order/placement differences in review views |
| Raster | 150 | Whole-page and regional paint/layout differences, with multiple gates per region |

The per-case E1 inventory below classifies the failures into gate phases.
Structural covers page/pagination, markup geometry, tables, annotations and
balloons; operations covers text and graphics. Counts are failing metrics.
Reports and hashes stay under ignored `artifacts/plan-revision-20261005/`:
`markup-baseline`, `markup-after-comparator`, `markup-original-view-corrected`,
`markup-cache-final`, and `markup-final-inventory.json`. Re-run with
`tools/RunDocxMarkupReferenceGate.ps1 -ContinueOnFailure -FailOnDeltas` after
populating trusted references on another workstation.

| Public case | Structural | Operations | Raster |
|---|---:|---:|---:|
| docx-markup-all | 12 | 11 | 6 |
| docx-markup-all-word-compatible | 9 | 10 | 5 |
| docx-markup-balloon-lane-bands | 6 | 10 | 6 |
| docx-markup-comment-hidden-anchors | 0 | 3 | 0 |
| docx-markup-comment-long | 6 | 8 | 3 |
| docx-markup-comment-static-stories | 7 | 10 | 6 |
| docx-markup-comment-table | 3 | 8 | 3 |
| docx-markup-comment-text-box | 0 | 4 | 2 |
| docx-markup-comment-threaded-resolved | 4 | 6 | 3 |
| docx-markup-comment-unresolved | 0 | 10 | 3 |
| docx-markup-final | 3 | 7 | 0 |
| docx-markup-links-fields-all | 10 | 9 | 5 |
| docx-markup-links-fields-final | 4 | 8 | 2 |
| docx-markup-links-fields-original | 4 | 8 | 2 |
| docx-markup-links-fields-simple | 10 | 10 | 6 |
| docx-markup-margin-dense-revisions | 6 | 10 | 6 |
| docx-markup-margin-landscape | 4 | 10 | 5 |
| docx-markup-margin-mirrored | 2 | 10 | 3 |
| docx-markup-margin-multi-column | 6 | 10 | 6 |
| docx-markup-margin-multi-page | 4 | 10 | 5 |
| docx-markup-margin-one-page | 4 | 11 | 5 |
| docx-markup-margin-table-heavy | 8 | 11 | 6 |
| docx-markup-original | 5 | 9 | 2 |
| docx-markup-simple | 10 | 11 | 6 |
| docx-private-grounded-comment-anchors | 9 | 11 | 6 |
| docx-private-grounded-complex-fields | 9 | 11 | 6 |
| docx-private-grounded-dense-balloons | 9 | 11 | 6 |
| docx-private-grounded-floating-drawings | 9 | 11 | 6 |
| docx-private-grounded-formatting-revisions | 9 | 11 | 6 |
| docx-private-grounded-numbering-indentation | 9 | 11 | 6 |
| docx-private-grounded-style-spacing | 9 | 11 | 6 |
| docx-private-grounded-table-borders | 9 | 11 | 6 |
| docx-private-grounded-threaded-comments | 9 | 11 | 6 |

## Original-View Autofit Update (2026-10-05, RV06-L1)

Original mode previously added the width of excluded insertion/move-to text to
autofit columns, even though that text was absent from the rendered page. It now
sizes only the original content. Final/Simple views retain the Office-calibrated
hidden deletion/move-from measurement behavior.

The clean original reference's first column is about 171.39 points wide. The
candidate moves from 118.43 to 171.42 points; the roughly 53-point discrepancy is
removed. Table-grid bounds now pass (maximum delta 0.787 points, including border
and row edges), and regional raster gates pass. The case drops from 16 to 10
failing gates. Date-field result differences, small text spacing, graphics and
link annotation differences remain; this is still an approximate case.

The regression changes an excluded insertion to 200 characters and then to
move-to content, requiring unchanged Original-mode columns and original text.
Affected table/core groups pass, and six held-out cases across Final, Simple,
AllMarkup and an independent Original input preserve candidate bytes and gate
counts. Evidence is under `artifacts/plan-revision-20261005/rv06-l1/`.

## Annotation Comparison Update (2026-10-05, RV06-L2)

Internal link targets now compare logical destination pages, view types and
coordinates rather than hashes of raw PDF destination syntax. Page order comes
from the page tree. Object numbers, numeric spelling and equivalent retained
zoom values no longer create target failures; direct destinations and local
GoTo actions compare together. Positions use the annotation bounds tolerance,
while source/destination pages, annotation subtypes, view types, retained
coordinates and zoom remain strict.
Raw and canonical hashes are both retained, alongside actual position deltas.
Named/indirect destinations and other unresolved actions retain exact checks;
this tool still inspects direct PDF objects rather than compressed object streams.

Annotation-only checks of all 33 existing PDF pairs preserve their failure counts.
Their current target differences are real: the clean Original case differs by
3 points horizontally and about 1.63 vertically. The normalization corrects a
reproduced comparator defect without claiming a renderer improvement or loosening
the remaining parity gates. Evidence is under
`artifacts/plan-revision-20261005/rv06-l2/`; portable adversarial checks run via
`tools/TestMarkupAnnotations.ps1` in both CI jobs.

## Text Whitespace Update (2026-10-05, RV06-L3)

The reader now removes insignificant XML whitespace at the edges of `w:t` and
`w:delText`. A direct `xml:space="preserve"` retains those edges; internal spaces
and nonbreaking spaces remain. Literal tab/newline characters inside text become
spaces, while `w:tab` and `w:br` retain their layout semantics. The default-edge
rule follows the [OOXML TextType remarks](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.texttype.space?view=openxml-3.0.1).
Independent Word 16 exports also establish that preservation attributes on
ancestor runs or paragraphs do not preserve a text element's edges.

The hidden-comment-anchor case previously emitted one extra space, causing
roughly 2.5-point width/advance differences. All its gates now pass. A regression
covers thirteen combinations of preservation, deleted text, nonbreaking spaces
and actual/literal text controls. Synthetic unit inputs that intend significant
edge spaces now declare preservation explicitly; their expected results remain
unchanged. The broader DOCX run passes 847 tests with no failures or skips.

Across all 33 cached-reference cases, gate failures decrease from 665 to 645 and
page counts remain unchanged. Thirty-two cases still have nonzero gates. Two
cases gain failures: the threaded/resolved case exposes a 0.298-point spacing
residual while its largest width mismatch improves from 61.48 to 35.20 points;
the mirrored-margin case changes wrapping on its wider even-page body frame,
moving two balloons and increasing failures from 15 to 21. That geometry defect
is the next bounded repair, rather than a reason to restore insignificant text.
Evidence is under `artifacts/plan-revision-20261005/rv06-l3/`.

## Mirrored Wrap Width Update (2026-10-05, RV06-L4)

Scaled Word-compatible layout now mirrors the body margins before computing its
review reserve. Odd and even pages therefore retain the same wrap width for
identical content, including an inside gutter. The authored odd-page right edge
is retained separately for balloon placement, matching Word's fixed right lane.
The unscaled PreserveDocumentLayout and ReserveMarkupMargin profiles keep their
existing geometry.

Three independent Word 16 probes cover mirrored pages, mirrored pages with a
24-point gutter and a nonmirrored gutter control. Repeated paragraphs keep the
same line breaks and balloon x position on both pages. The new regression
exercises that behavior through layout and balloon inspection. The broader DOCX
run passes 848 tests, with no failures or skips. The public mirrored case returns
from 21 to 15 failures; its even-page last-baseline delta falls from 7.267 to 1.25
points, and its raster-region maximum error returns from 27.72 to 18.35. It still
has spacing, graphics and other parity residuals. The other 32 cached cases
preserve candidate PDF bytes and gate counts; total failures fall from 645 to
639, with unchanged page counts. Evidence is under
`artifacts/plan-revision-20261005/rv06-l4/`.

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
