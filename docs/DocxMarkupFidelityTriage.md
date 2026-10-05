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

## Printed Comment Thread Update (2026-10-05, RV06-L5)

Independent Word 16 exports print only the parent comment for resolved/open
threads with one or two replies. Removing the reply relationship and anchoring
that comment independently makes its text print, confirming the visibility rule.
Word-compatible balloons now use the parent text and its wrapped rows, with no
reply summary, extra height or separator strokes. Thread ownership, dates,
resolved/open state and reply counts remain in the model and inspection. The
other geometry profiles retain their compact reply previews.

A red/green regression compares production page content with a parent-only
control for both resolved states. It also checks metadata retention, the default
profile's reply summary and an empty parent through the diagnosed font fallback.
The broader DOCX suite passes 849 tests with no failures or skips.

The threaded/resolved case drops from 15 to 9 failures. Last-baseline drift falls
from 8.273 to 0.077 points and body-height drift from 8.146 to 0.05 points. Text
spacing/advance, graphics and regional raster differences remain. The other 32
cached cases preserve candidate PDF bytes and gate counts, with unchanged page
counts; total failures fall from 639 to 633. Evidence is under
`artifacts/plan-revision-20261005/rv06-l5/`.

## Bookmark Viewport Update (2026-10-05, RV06-L6)

Eleven independent Word 16 exports establish that ordinary body bookmark
destinations use the containing column's left edge, with three design points of
context. Paragraph indentation and center/right alignment do not move that
viewport; visible text before a midline bookmark does. The context scales with
Word-compatible review printing. Layout now retains the column's position
relative to the line. Table, static-story, note and textbox target geometry keeps
its existing behavior pending separate reference qualification.

The reader also now captures bookmark offsets relative to their source run,
matching the renderer's interpretation. Previously a midline bookmark counted
the preceding paragraph text twice. Merging paragraphs across a deleted mark
shifts source/run indexes while retaining that local offset. Scaled destinations,
hyperlink rectangles and text-emission inspection now follow each target/source
page's actual mirrored left margin.

Three production regressions cover alignment/indentation, column breaks,
mirrored review pages, ordinary/nested reader anchors and deleted-mark merging.
The mirrored inspection is checked against the emitted PDF text matrix and link
rectangle. Across the eleven Office controls, horizontal destination deltas are
at most 0.61 points; the nine ordinary controls have sub-point vertical deltas.
The broader DOCX run passes 852 tests with no failures or skips.
The column-break control retains a 24.58-point flow residual, and the mirrored
control retains a 1.12-point baseline residual. These are not closed by changing
the destination padding.

All 33 cached cases preserve raster bytes and page counts. Twenty-eight preserve
PDF bytes; five change bookmark destinations. The Word-compatible All Markup
target gate passes, reducing total failures from 633 to 632. Its annotation
rectangle still differs, and the other markup views retain target baseline
differences. One case passes all gates and 32 remain partial. Evidence is under
`artifacts/plan-revision-20261005/rv06-l6/`.

## Column-Break Paragraph Update (2026-10-05, RV06-L7)

Break-only body paragraphs now resume their paragraph mark in the next column
or page. Their line pitch is consumed there, followed by normal after-spacing.
Before-spacing carries only the amount beyond the preceding paragraph's
after-gap. The retained empty run now participates in the font plan so automatic
pitch uses its resolved face rather than unqualified fallback metrics. This
preserves the existing body-paragraph inventory and uses the retained break
paragraph model.

Ten independent Word 16 controls cover exact/automatic line height, 12/24-point
mark fonts, before/after spacing and a single-column turn to the next page.
Vertical destination differences decrease from 13.58-37.58 points to at most
0.42 points. The original two-column bookmark control improves from 24.58 to
0.38 points. Two regressions verify collapsed spacing through layout and
resolved-font pitch through production text emission, using a portable test face.
An inline-break control remains unchanged at 28.58 points; missing continuation
paragraphs after inline breaks require separate qualification.

The broader DOCX run passes 854 tests with no failures or skips. All 33 cached
cases retain PDF/raster bytes and gate counts; total failures remain 632. Raw
controls and reference exports are under `artifacts/plan-revision-20261005/rv06-l6/`;
repair verification is under `artifacts/plan-revision-20261005/rv06-l7/`.

## Trailing Inline Column-Break Update (2026-10-05, RV06-L8)

Body paragraphs ending with an inline column break now retain their empty
continuation paragraph. It carries the authored paragraph-mark style, line
height and after-spacing in the new column; before-spacing and numbering are
removed from the continuation as for existing nonempty fragments. Previously
the reader discarded this tail, placing the following paragraph too high.
Table-cell lowering keeps its existing behavior pending separate qualification.

Eight independent Word 16 controls cover default/explicit styles, exact/automatic
spacing, after-spacing and mark/prefix/break-run font sizes. Destination vertical
differences are now at most 0.42 points, compared with 12.58-37.58-point gaps for
the trailing-break controls. The mid-paragraph control remains within 0.22 points.
Changing the break or prefix font does not change an empty continuation's pitch;
changing its paragraph-mark font does. A regression checks direct and hyperlink
tails through the reader and production text emission, with mid-paragraph and
page-break controls.

The broader DOCX run passes 855 tests with no failures or skips. All 33 cached
cases preserve PDF/raster bytes and gate counts; total failures remain 632.
Raw controls are under `artifacts/plan-revision-20261005/rv06-l6/`, with final
verification under `artifacts/plan-revision-20261005/rv06-l8/`.

## Body Hyperlink Line-Box Update (2026-10-05, RV06-L9)

Ordinary body hyperlink rectangles now follow the printed paragraph line box.
Horizontal padding is 2.25 design points per side, scaled with review printing,
independent of glyph size. The layout retains the body's baseline inset and
printed slot height: exact heights need print scaling, while automatic heights
already contain scaled font metrics. A larger neighbouring run therefore controls
the top edge even when the linked run is small. Non-consuming page-break spill
glyphs no longer shorten the preceding clickable slot.

Twelve independent Word 16 controls cover 11/12/24-point glyphs, exact/automatic
spacing, scaled/unscaled printing, mixed font sizes, trailing paragraph spacing
and a following paragraph with a different font size. Every rectangle bound is
within 0.14 points of Word. Before the repair, automatic-spacing bottoms differ
by 2.71-5.86 points, and the mixed-size rectangle is 11.98 points too short.
Two regressions exercise production emission across fonts, print scale, spacing,
mixed runs and break-spill rows. Both fail on the previous renderer and pass
with the repair.

Table cells, static stories, text boxes, notes and at-least spacing retain their
existing geometry pending separate reference qualification. The cached
Word-compatible link also has a separate body-baseline difference; this change
does not claim to correct that text-flow residual. Raw controls and verification
are under `artifacts/plan-revision-20261005/rv06-l8/` and `rv06-l9/`.

The broader DOCX run passes 857 tests with no failures or skips. All 33 cached
cases preserve raster bytes, page counts and gate tallies; total failures remain
632. Nine PDFs change through annotation geometry, while the other 24 remain
byte-identical. The cached Word-compatible rectangle's maximum difference remains
2.267 points, predominantly its body's accumulated baseline drift.

## Review Typographic-Metrics Update (2026-10-05, RV06-L10)

Review print scaling now retains the resolved font's typographic-metrics
selection and scales its typographic line box once. The wrapper previously
dropped this provider interface, so fonts requesting typographic metrics used
larger Windows boxes in the scaled layout. A portable production regression
checks both flag-set and flag-clear synthetic faces: the former uses its
typographic box, while the latter continues to use Windows extents.

In the cached Word-compatible review fixture, ordinary body-baseline differences
fall from 0.89-2.17 points to at most 0.07 points; gate failures decrease from
23 to 19. Correcting pitch exposes separate annotation assumptions: the bookmark
viewport's Windows ascent differs by 1.453 points, and clipping a body hyperlink
against the following table's glyph top leaves a 2.563-point bottom-edge
difference. These are separate follow-up calibrations. The full DOCX run passes
858 tests with no failures or skips. Raw evidence is under
`artifacts/plan-revision-20261005/rv06-l10/`.

The complete 33-case cached comparison reduces total gate failures from 632 to
586. All page counts match; 19 PDF/raster cases remain byte-identical and 14
change with the repaired line metrics. No case's total failure count increases.
The newly exposed bookmark target delta above is retained as a named residual,
despite the main case's net improvement. One case passes and 32 remain partial.

## Body-Link Table-Boundary Update (2026-10-05, RV06-L11)

A table or another story's glyph top no longer clips the preceding body's
qualified hyperlink slot. Its own printed line height and trailing paragraph
spacing define that boundary. A production regression varies the following
table font across 8/12/24 points and requires the body slot to stay unchanged.
It fails on the previous renderer and passes with the repair.

The Word-compatible cached link rectangle's maximum bound difference decreases
from 2.563 to 0.245 points, passing its rectangle gate. The separate bookmark
viewport residual remains. The DOCX suite passes 859 tests with no failures or
skips. Across all 33 cached cases, total failures decrease from 586 to 585;
raster bytes and page counts stay unchanged, and 31 PDFs are byte-identical.
Evidence is under `artifacts/plan-revision-20261005/rv06-l11/`.

## Uniform Typographic-Ascent Update (2026-10-05, RV06-L12)

Uniform automatic body paragraphs whose resolved face requests typographic
metrics now place their baseline using typographic ascent plus line gap. Their
bookmark viewports use the same ascent at the emitted font size. The baseline
inset remains a design coordinate through the review-scaling wrapper; layout
applies print scaling once. Mixed faces/styles/sizes, differing paragraph-mark
sizes, inline images/text boxes, exact/at-least spacing and other stories retain
their existing baseline geometry pending separate qualification.

Eight independent Word 16 controls cover 12/24-point Aptos and Abadi, explicit
paragraph-mark controls and scaled mirrored pages. Destination vertical
differences are at most 0.13 points unscaled and 0.50 points scaled. Previously
Aptos differs by 0.87/1.71 points and Abadi by -1.65/-3.42 points. Abadi's paragraph
pitch already matches Word; its baseline had incorrectly followed horizontal
header ascent. Changing paragraph-mark styling leaves the control unchanged.

Two production regressions cover requested/unrequested typographic metrics,
zero/positive line gap, font size, viewport ascent and review scaling. Both fail
on the previous renderer and pass with the repair. Five typographic tests and
all 861 DOCX tests pass with no failures or skips. Raw qualification and
verification are under `artifacts/plan-revision-20261005/rv06-l12/`.
The complete 33-case cached comparison reduces total failures from 585 to 579.
The main Word-compatible, Final and Original cases each lose two failures,
including their bookmark-target gate; no case's total increases. All page
counts match. Seven raster cases and one PDF remain byte-identical; the
remaining cases change with baseline or annotation geometry. The main
Word-compatible target's maximum position difference is 0.804 points and its
rectangle difference remains 0.245 points. One case passes and 32 remain
partial; these scoped repairs do not establish complete Office parity.

## Autofit Column-Maximum Update (2026-10-05, RV06-L13)

Autofit now measures the widest cell in each column independently. Previously
it selected one row with the largest total content width and used all of that
row's widths, missing wider cells in other rows. A production regression moves
cells between rows within their own column and requires column widths to stay
unchanged. It fails before the repair and passes afterward.

Eight independent Word controls separate clean cell text, comment anchors and
review scaling. In the unscaled insertion-only control, the first-column width
difference falls from 6.66 to 0.07 points; all three unscaled controls have
differences of at most 0.28 points. The scaled insertion-only control exposes a
separate comment-marker measurement difference (15.43 to 20.66 points). In a
separate combined-text control, removing the table comment anchor reduces the
width difference from 19.30 to 0.27 points. That remains a scoped follow-up.

All 862 DOCX tests pass. Across 33 cached comparisons, page counts and case
failure totals remain unchanged at 579; 32 PDF/raster cases are byte-identical.
The table-comment case's maximum text-position difference improves from 16.339
to 9.058 points, and its maximum regional raster MAE improves from 15.714 to
14.536. Its 12 failed gates remain partial. Raw evidence is under
`artifacts/plan-revision-20261005/rv06-l13/`.

## Review Comment-Label Autofit Update (2026-10-05, RV06-L14)

Word-compatible review tables now include hidden comment display labels in
their preferred content widths. Word measures these labels at the paragraph
mark's resolved font face and size, ignoring direct formatting on the comment
reference run. Labels use first-seen comment display order rather than source
IDs; the same per-conversion labels also supply printed balloon titles. This
measurement does not emit additional body glyphs. Other markup geometry modes,
nested tables and tables in other stories retain their existing behavior.

Nineteen trusted Word controls include literal-label substitution, changed
initials, 8/24-point reference formatting, paragraph-mark/text size changes,
source-ID renumbering and three unscaled controls. Eighteen have column-width
differences below 0.3 points. A tight table with both text and paragraph mark
at 24 points retains a 3.76-point difference; its literal-label Office control
has the same width, isolating a separate minimum/maximum-content allocation
follow-up.

The production regression fails before the repair and passes afterward across
32 combinations; it also verifies that hidden-label measurement adds no body
digits. All 863 DOCX tests pass. The 33 cached comparisons reduce total failed
gates from 579 to 576 without increasing any case's total. All page counts
match, and 30 PDF/raster cases remain byte-identical. The table-comment case
falls from 12 to 10 failures and its maximum regional raster MAE improves from
14.536 to 7.847. One case passes and 32 remain partial. Raw evidence is under
`artifacts/plan-revision-20261005/rv06-l14/`.

## Autofit Compression and Minimum-Width Update (2026-10-05, RV06-L15)

Qualified autofit tables now compress the flexible width above each column's
minimum content width. Unbreakable words stay together across styled runs and
nonbreaking spaces; hyphens use the shared wrapping rules. Allocation includes
the cell insets used by layout. A preferred table width below the content
minimum can grow within the available frame. Fixed layout, explicit cell
preferred widths and spans retain their existing paths. Deleted content,
positioned tabs, hard breaks and minima wider than the frame remain separately
scoped.

Thirty Word controls cover width scans, word boundaries, review labels and
fixed layout. Column differences are below 0.3 points; the previous tight-table
review-label difference falls from 3.76 to 0.15 points. The fixed narrow control
matches within 0.06 points. A new regression fails before the repair and passes
across eight variants, covering styled run boundaries, nonbreaking spaces,
preferred-width preservation and growth below the content minimum.

The full DOCX run and focused rechecks cover all 864 tests with no remaining
failures or skips. Four emergency-wrap fixtures now explicitly select fixed
layout; their assertions remain intact. Seventy-two focused checks pass after
that correction. All 33 cached page counts match, 32 PDF/raster cases remain
byte-identical, and total failed gates fall from 576 to 575. The table-heavy
margin case falls from 20 to 19 failures; no case total increases. One case
passes and 32 remain partial. Raw qualification, the original full-run report
and focused rechecks are under `artifacts/plan-revision-20261005/rv06-l15/`.

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
