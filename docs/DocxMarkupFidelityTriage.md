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

## Review Font-Grid Order Update (2026-10-05, RV06-L16)

DOCX embedded-font emission now rounds the nominal font size on the Office
600-DPI export grid before applying review print scaling. The print scale
travels with each emission segment so rendering, inspection, terminal advance
and hyperlink widths use the same font-size plan. Unscaled DOCX and the shared
PPTX font-grid behavior retain their existing results.

Six independent Word size controls reduce the largest font-size difference
from 0.048 to 0.004 points and the largest natural-width difference from 0.379
to 0.032 points. Three unscaled controls remain exact for font size and width.
The remaining small printed-scale quantization difference is explicit. The
regression exercises both inspection and emitted PDF text state across six
sizes and two page widths, using a portable synthetic embedded font.

The large main-case width outlier involves different cached DATE-field results;
it does not establish a font-metric defect. Terminal-space paragraph-mark caps
remain a separate measured follow-up. Three custom-page controls also expose
larger review-scale differences despite matching PDF page sizes; those require
page/print scaling qualification and are outside the Letter font-grid result.

The full DOCX run and focused connector rechecks cover all 865 tests without
remaining failures or skips. Three connector assertions now use layout widths;
their bounds and placement behavior are unchanged. Both PPTX font-grid checks
pass. All 33 cached page counts and failure totals remain unchanged at 575,
with 18 PDF/raster case identities. One case passes and 32 remain partial.
Raw qualification and validation are under
`artifacts/plan-revision-20261005/rv06-l16/`.

## Resolved Paragraph-Mark Size Update (2026-10-05, RV06-L17)

Terminal spaces now retain the resolved paragraph-mark size through review
scaling. The reader already resolves that size through the paragraph style
cascade, independently of the final visible run. The old large-font cap is
retained for model-only paragraphs with no resolved mark size.

Eight independent Word controls vary the body size and explicit or inherited
mark size. The largest terminal font-size difference drops from 10.05 to 0.004
points in the prototype; body baselines and terminal X positions are unchanged.
The production regression checks inspection and emitted PDF font state across
20 variants, including null marks and preserved layout. Terminal X differences
up to 0.43 points remain a separate residual.

The full DOCX run passes 866 tests with no failures or skips. All 33 cached
rasters, page counts and failure totals are unchanged at 575, and 32 PDFs
remain byte-identical. Only the comment text-box case changes PDF text state.
Raw evidence is under `artifacts/plan-revision-20261005/rv06-l17/`.

## Custom-Page Review Scale Qualification (2026-10-06, RV06-L18)

Eleven Word controls cover A4, Letter landscape and custom page dimensions.
PDF page sizes match. The standard-page font-size differences are at most
0.031 points; custom portrait/landscape differences reach 3.104 points, with
coupled position and baseline differences. These are print-scale residuals,
separate from nominal font-grid rounding.

Four additional exports with Word's
[paper-mapping option](https://learn.microsoft.com/en-us/office/vba/api/word.options.mappapersize)
disabled have identical rasters and text geometry. Word identifies the custom
pages as paper size 41.

Eight controls were then exported with each of two installed virtual drivers,
Microsoft Print To PDF and OneNote. Custom-page body font sizes remain identical
between drivers; Letter landscape differs by 0.024 points. All text and page
assignments match, with small coordinate and advance differences. The custom
candidate font-size residual still reaches 3.104 points. Every export preserves
document dimensions and the Windows default printer. Selection uses Word's
hidden [printer setup dialog](https://learn.microsoft.com/en-us/office/vba/api/word.wdworddialog)
with `DoNotSetAsSysDefault=1`; validation exports PDF without submitting a print
job and restores the original application printer.

A read-only [GetPrinter](https://learn.microsoft.com/en-us/windows/win32/printdocs/getprinter)
query finds Letter global defaults for both drivers and no per-user overrides.
This qualifies a driver change with the same default paper. It does not qualify
a different paper canvas. A smaller printer canvas remains an inference that
requires a different-paper control. A section-level OOXML printer-settings part
cannot supply that control: [Microsoft documents that Word discards it](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/ff9e6328-9e35-4396-9651-50cd7cfdfdcb).
The renderer keeps its current formula;
no printer-specific constant was introduced. Evidence is under
`artifacts/plan-revision-20261005/rv06-l18/`, including
`alternate-profile-evidence.json` and `paired-printer-controls.json`.

## Missing-Font Inspection Update (2026-10-05, RV06-L19)

Inspection now reports the font size and spacing actually used by standard-14
fallback glyphs. Those glyphs use unrounded layout sizes and absolute positions,
with zero PDF text-state character spacing. Embedded fonts retain their export
grid and spacing plan. Rendering and missing-font diagnostics are unchanged.

A regression compares inspection against emitted PDF font state and glyph
coordinates across 12 size/spacing/view variants. It fails before the repair and
passes afterward. All 23 text-emission checks and four missing-font checks pass.
Nine public document/view controls retain identical PDF bytes and diagnostic
counts, verified against the actual loaded baseline and candidate DLL hashes in
both prototype and root. Evidence is under
`artifacts/plan-revision-20261005/rv06-l19/`.

## Simple Review-Table Baseline Update (2026-10-06, RV06-L20)

Word-compatible review layout now projects simple table geometry onto the
existing body baseline anchor and scales each cell's first-baseline inset by
the review print scale. The preceding and following body text, row pitch and
pagination retain their existing coordinates. The correction applies only to
complete tables on one page, with automatic row heights and a single
top-aligned text line in each cell. Wrapped or split rows, declared heights,
vertical merges, nested tables, inline graphics, multiple paragraphs and
explicit line heights retain the previous layout.

Twenty-four independent Word 16 controls vary cell font size, preceding
paragraph count, paragraph spacing, borders, following body text and a table
at the document start. All improve their pixel comparison. Unbordered controls
have maximum cell-baseline, fill-top and fill-bottom differences of 0.11, 0.18
and 0.20 points respectively. Bordered controls reach 0.21, 0.57 and 1.22
points; border-driven height differences remain a separate residual. Following
body baselines remain identical to the accepted renderer.

Two production regressions cover 32 font/page/start-position variants and
surrounding-text invariance. The baseline fails the inset check; the proposal
passes on Windows and Linux. Full integration at the accepted L20 revision
passes 2,150 Windows tests, including 869 DOCX checks, with one unconfigured
private probe skipped. Release builds are clean and the exact fresh 0.1.5
package smoke loads the DLL tested by the full suite. The scoped proposal also
passes all 43 API tests. Its PDFs
retain the measured proposal's bytes for all 24 controls and all 33 cached
cases; four fallback controls retain the accepted renderer's bytes. Twenty-four
preserved-layout controls also retain their bytes.

All cached input/reference identities and page counts match. Thirty cached
PDFs and rasters retain their bytes; the three changed cases improve or retain
every measured raster region. Failed gates decrease from 575 to 573, with no
case increase. One case passes and 32 remain partial. The main review fixture
closes two last-baseline gates; wider table/markup parity remains incomplete.
Raw qualification is under `artifacts/plan-revision-20261005/rv06-l20/`.

## Simple Review-Table Border Update (2026-10-06, RV06-L21)

Qualified simple tables now scale solid border widths, junctions and row
advances by the review print scale. Layout already shrinks table coordinates;
border paint uses a local inverse projection and a PDF transform to apply the
width scale once. Following body text consumes the printed border advances.
The complete-table admission from L20 also gates this repair. Declared heights,
wrapped or split cells, other border styles and the remaining unqualified
layouts retain their previous behavior. A shared numeric guard retains the
legacy border advances and paint if the PDF transform cannot represent the
scale within 0.1 percent.

Across 40 Word controls, 22 rasters improve and 17 retain their bytes at 144
DPI. One thin-border control rises from 0.223 to 0.246 MAE at that resolution;
at 288 and 432 DPI its MAE falls from 0.273 to 0.215 and 0.261 to 0.209, with
improved similarity. The maximum border-width difference falls from 0.60 to
0.094 points. New one/three-row controls reduce the maximum following-body
baseline difference from 2.50 to 0.31 points. Existing two-row bordered controls
fall from 0.54/0.50 to 0.018/0.025 points. Wider border quantization remains
approximate.

The initial bounded projection passes 870 DOCX checks with no failures or
skips. The final numeric guard passes four production regressions on Windows
and Linux. Both border-advance and numeric-fallback checks have failing prior
implementations. All 40 qualified, 40 preserved-layout, four fallback and 33
cached PDFs retain their qualified bytes after the guard. The final Release
build is clean and a fresh 0.1.5 package smoke loads the DLL used by the focused
tests. The broader table-border experiment is superseded.

A fresh integration pinned to `49a9e80b` subsequently passes 2,152 Windows
checks with no failures and one unconfigured private-probe skip, including all
871 DOCX checks. Its clean Release build and exact fresh 0.1.5 package smoke
use the same library bytes. This integration predates the balloon repairs below.

The cached corpus keeps 573 failed gates, one passing case and 32 partial
cases. All input/reference identities and page counts match; 30 PDFs and
rasters retain their bytes. The three changed cases improve overall pixel
error and similarity. The main fixture's balloon-region MAE rises from 7.50 to
7.70, while its body, table and connector regions improve; its last-baseline
difference rises from 0.182 to 0.343 points and stays within its gate. These
remain explicit composition/flow residuals. Comparisons are linked to fresh
final candidate PDFs by input, reference and output hashes. Evidence is under
`artifacts/plan-revision-20261005/rv06-l21/`.

## Balloon Placement and Body Coverage (2026-10-06, RV06-L22/L23)

Sixteen independent Word controls place a comment before or after a table,
varying cell font size, row count and border width. Balloon title baselines
differ by at most 0.121 points, vertical box edges by 0.218 points, and the
title offset from its anchor by 0.050 points. These controls support retaining
the existing placement rather than applying another vertical correction.

The controls instead expose missing body letters. Synthetic labels receive
extra glyph coverage, but Word-compatible bodies can use another font subset.
The renderer now ensures coverage in the actual body subset and shares the
label supplement when the original resource is the same. Existing complete
subsets and missing-face fallback retain their paths. The production regression
decodes actual emitted CIDs and checks distinct faces, wrapped text and a
supplementary character. A grouped-balloon assertion now checks bold/regular
faces and its printed summary prefix, including the existing ellipsis, rather
than pinning resource IDs.

The initial broader DOCX run has 871 passing checks and one obsolete resource-ID
assertion. With that assertion replaced, all 69 balloon checks pass and the
portable content regression passes on Linux. The fresh 0.1.5 package smoke
loads the exact library used by the final focused tests. All 33 freshly rendered
cached PDFs retain their bytes and their 573 failed gates. Sixteen controls
restore the full comment text without changing graphics; their page MAE rises
by 0.001–0.002 because body typeface selection remains wrong in that slice.
Raw qualification is under `artifacts/plan-revision-20261005/rv06-l22/` and
`artifacts/plan-revision-20261005/rv06-l23/`.

## Uniform Comment Body Faces (2026-10-06, RV06-L24)

Nine independent Office controls cross Calibri, Arial and Aptos document runs
with inherited, Arial and Courier New comment runs. Word uses the comment's
face in every combination. Qualified regular comments now carry their own
prepared font resource through balloon height measurement and emission.
Admission requires a single plain paragraph, one regular face, and complete
glyph coverage. Matching existing faces retain their resources; mixed faces,
formatting, compound stories and grouped balloons keep their previous path.
The coverage scan observes cancellation without copying the body string.

All sixteen table controls improve after the face correction; seven of the
nine font controls improve and two retain their rasters. First-row advances
in the font matrix differ from Office by at most 0.12 points. Two held-out
wrapped comments improve pixel error and similarity. Arial retains three
printed rows; Courier still prints four where Office prints five, and its
continuation pitch remains approximate. The mixed-face control retains the
content repair and its legacy typeface path. Font selection does not establish
complete comment formatting or wrapping parity.

The final Release build is clean; 70 balloon checks and both portable content
and typeface regressions pass. The fresh 0.1.5 package smoke uses the exact
tested library. All 33 cached PDFs retain their accepted bytes and 573 failed
gates, with comparisons linked by input, reference and candidate hashes.
Final-source identity checks retain all 28 qualified control PDFs. Evidence
is under `artifacts/plan-revision-20261005/rv06-l24/`.

## Terminal Blank Positions (2026-10-06, RV06-L25)

Word-compatible balloons now position terminal spaces using the face that
actually emits the body. Previously the title face measured that position,
which could put a blank text operation far from the end of the visible text.
The production regression uses distinct advances for the two faces and checks
both single-row and wrapped bodies. It fails before the repair and passes
afterward. All 71 balloon checks pass, the Release build is clean, and the
portable regression and exact fresh 0.1.5 package smoke pass.

Twenty-eight independent Office controls retain their raster and graphics
bytes. Their 31 moved spaces reduce the maximum position gap from 43.676 to
0.015 points. Across all 33 cached cases, 18 PDFs retain their bytes and the
15 changed PDFs move only 54 blank operations horizontally. All 36 raster
pages and all graphics remain identical. Failed gates fall from 573 to 546;
nine cases each lose three failures, and no case gains a failure. One case
passes and 32 remain partial. Comparisons retain input, reference and fresh
candidate hash links. This closes a text-operation position defect; wrapped
line breaks, continuation pitch and terminal-space typeface remain residuals.
Evidence is under `artifacts/plan-revision-20261005/rv06-l25/`.

Full integration pinned to `8576f276`, before this terminal-position repair,
passes 2,154 Windows checks with no failures and one unconfigured private
layout probe skipped, including 873 DOCX checks. Its fresh 0.1.5 package
contains the exact full-suite DLL. Evidence is under
`artifacts/plan-revision-20261005/milestone-8576f276/`.

## Wrapped Break Spaces (2026-10-06, RV06-L26)

Wrapped balloon rows now reserve the advance of their separating space when
choosing a line break. The previous width check measured only the visible
words and could keep an extra word despite the emitted space exceeding the
available width. The final word retains its existing measurement. A portable
production regression checks actual emitted rows with deterministic font
advances; it fails before the change and passes afterward. All 72 balloon
checks pass, the Release build is clean, and the fresh 0.1.5 package smoke
contains the exact tested DLL.

Nine new Word controls cross Aptos, Arial and Courier New with inherited,
9-point and 18-point comment sizes. Word prints nominal 9-point balloon text
in all nine; the candidate now matches every line break. Six controls improve
pixel error and similarity and three retain their rasters. Both earlier
wrapped controls also improve: Courier now matches Word's five rows rather
than printing four. The other 26 prior controls retain their rasters, including
the mixed-face fallback, whose composition remains partial. Main-document
text operations stay unchanged in all 37 controls.

Across the 33 cached cases, 32 PDFs retain their bytes. The changed long-comment
case improves page MAE from 0.660 to 0.625 and similarity from 0.831 to 0.844;
its balloon region also improves. All page counts and input/reference identities
hold. Failed gates remain 546, with no case increase. Font-dependent
continuation pitch remains a separate residual. Evidence is under
`artifacts/plan-revision-20261005/rv06-l26/`.

## Uniform Face Continuation Pitch (2026-10-06, RV06-L27)

Uniform regular comments that already carry their own qualified body resource
now use that face's horizontal line metrics for continuation spacing and
balloon height. Both paths share the same calculation. Nonpositive or invalid
metrics retain the existing pitch. Grouped, mixed-face and legacy resource
paths also retain their spacing. This admission rule does not yet cover a
comment whose face is the same as the legacy body face.

The portable production regression uses a body face with distinct line metrics
and checks emitted baseline differences and matching balloon height. It also
checks the fallback for nonpositive metrics. All 73 balloon checks pass, the
Release build is clean, and the exact tested DLL passes fresh 0.1.5 package
smoke. The regression also passes on Linux.

All nine new Word controls and both earlier wrapped controls improve pixel
error and similarity without changing line breaks. The largest baseline gap
across the nine controls falls from 2.00 to 0.18 points. The other 26 controls
retain their rasters, and all 37 retain their main-document text operations.
All 33 freshly rendered cached PDFs retain their bytes and 546 failed gates.
Same-face continuation admission, mixed composition and existing nonzero
markup gates remain residuals. Evidence is under
`artifacts/plan-revision-20261005/rv06-l27/`.

## Same-Face Metric Admission (2026-10-06, RV06-L28)

Qualified comments now carry their line-height ratio independently of a font
resource replacement. A comment that shares the existing body face keeps its
resource and subset while gaining the same metric-based pitch as a distinct
face. Grouped and mixed-face comments retain their previous pitch; invalid
metrics retain the fallback. No public snapshot or API field changes.

The production regression checks actual same-face rows, their box height and
the mixed-face fallback. The older wrapped-height assertion now compares its
four emitted rows with the box rather than assuming a fixed 1.2 factor.
All 877 DOCX checks and 74 balloon checks pass, as do the portable same-face
and distinct-face/fallback regressions. The Release build is clean and the
fresh 0.1.5 package contains the exact tested DLL.

Three new Word controls with matching document/comment faces all improve
pixel error and similarity. Their largest baseline gap falls from 2.00 to
0.18 points; all prior 37 control PDFs retain their bytes. Across 33 cached
cases, 22 PDFs retain their bytes and all page counts/input/reference identities
hold. The 11 changed PDFs retain font resources and every text-operation field
except vertical position: 90 operations move. Failed gates fall from 546 to
544, with no case tally increase; the multi-page case loses two failures.

Pixel results remain mixed. At 144dpi, 12 of 14 changed pages improve error,
two increase by at most 0.0088, and five lose similarity. Higher-resolution
checks of the five affected cases retain small differences: each resolution
has two page-error increases and four similarity decreases. The largest error
increase at 432dpi is 0.0143 in the multi-column case. In the links/fields
case, the pitch gets closer to Office while the existing title position stays
12.71 points above its reference. First-row transitions, anchoring and composed
layout remain residuals; these results do not establish complete parity.
Evidence is under `artifacts/plan-revision-20261005/rv06-l28/`.

Full integration pinned to `dc66b4f2`, through L26 and before these pitch
repairs, passes 2,156 Windows checks with no failures and one unconfigured
private-layout probe skipped, including 875 DOCX checks. Its fresh 0.1.5
package contains the exact full-suite DLL. Evidence is under
`artifacts/plan-revision-20261005/milestone-dc66b4f2/`.

## Final Blank Face (2026-10-06, RV06-L29)

Qualified single-paragraph, uniform regular comments now retain their resolved
paragraph-mark face for the final emitted blank. Intermediate break spaces
keep the body face, and the final blank keeps the body-based X position and
nominal 9-point print size. Mixed, grouped and unsupported comment composition
retain their existing path. Font preparation collects the mark's space before
rendering; the renderer does not create another font resource.

Four new Word controls vary the paragraph-mark face and include an 18-point
mark. Word honors the mark face but keeps nominal 9-point balloon printing.
The candidate matches those faces, with terminal advance differences below
0.01 points. The portable production regression checks both single-row and
wrapped output, including intermediate spaces. It fails before the repair
and passes afterward. All 878 DOCX checks and 75 balloon checks pass, the
Release build is clean, three portable regressions pass on Linux, and fresh
0.1.5 package smoke contains the exact tested DLL.

All 44 Office controls retain visible text, text geometry, graphics and raster
bytes. Nineteen final blank faces change; 25 control PDFs retain their bytes.
All 33 cached PDFs retain their bytes and 544 failed gates. One case passes
and 32 remain partial. This repairs a text-operation font-state difference;
existing pixel, body-flow and composed-layout residuals remain. Evidence is
under `artifacts/plan-revision-20261005/rv06-l29/`.

## REF Result Freshness (2026-10-06, RV06-L30)

The links/fields title's roughly 12.7-point offset comes from an upstream text
change during Office export. Word refreshes the five unlocked REF fields from
bookmark text, adding a wrapped row. The candidate retains the stored field
results and their revisions, as documented by its cached-result capability.
This difference is not evidence of a comment-anchor selection defect.

Three controls use explicit Word all-markup view settings. The unlocked control
retains the content mismatch and 12.763-point title delta. Locking the fields
preserves matching body text and reduces the absolute delta to 0.147 points.
Refreshing their stored text before locking also matches body text and reduces
the delta to 0.113 points. Both matched controls agree on line breaks.

The new public locked-field companion changes only five `w:fldLock` attributes
at the XML level. Its checked-in generator can produce it alone with
`-LockedOnly`, preserving the previous fixture bytes. Public inventory now has
344 manifests across ten families. Its controlled Office reference is cached
under the explicit all/word-compatible variant. The case remains approximate,
with seven failed gates for spacing, graphics, table geometry and pixels;
first/last body-baseline deltas are below 0.12 points. The previous 33 cases
retain their 544 failures; this additional case is separate coverage, not a
reduction of those failures. Evidence is under
`artifacts/plan-revision-20261005/rv06-l30/`.

## Additional Continuation Faces (2026-10-06, RV06-L31)

Four held-out Word controls extend uniform comment-face coverage to Georgia,
Cambria, Consolas and Times New Roman. All preserve body content and match
Office line breaks, including four Consolas rows. The largest baseline gap is
0.177 points. Office's small first-row and continuation differences remain;
these controls support the accepted metric admission without qualifying a
font-specific positioning constant. No runtime change is introduced. Evidence
is under `artifacts/plan-revision-20261005/rv06-l31/`.

## Single-Row Mixed Faces (2026-10-06, RV06-L32)

Plain regular mixed-face comments now retain their prepared font resources when
the complete body fits the first row. Each part follows the preceding part's
measured advance on a shared baseline; geometry and emission share the fit
check. Grouped, styled, compound and overwide comments retain the previous
path. Source glyph coverage and cancellation are checked before admission;
no font is created during emission. Wrapped mixed composition and its final
blank face remain residuals.

The production reproducer fails before the repair and passes afterward,
checking actual font CIDs, complete content, advances, baseline and height.
It also checks styled and overwide fallback. The older mixed-face assertion
now checks the two emitted faces rather than expecting flattened text.
All 879 DOCX checks and 76 balloon checks pass; three portable regressions
pass on Linux and the Release build is clean. Fresh 0.1.5 package smoke
contains the exact tested DLL.

Across 57 Office controls, seven mixed-face cases improve error and similarity,
50 PDFs retain their bytes, and all retain graphics and main-document text.
The repaired operations match Office faces and text, with maximum horizontal
and baseline gaps of 0.126 and 0.177 points. The long, bold and multi-paragraph
controls retain their PDFs and remain partial. Comparisons now distinguish
font-operation counts from rows grouped by their baselines: the earlier short
mixed control has two font operations on one visual row. All 34 cached PDFs
retain their bytes and 551 failed gates. Evidence is under
`artifacts/plan-revision-20261005/rv06-l32/`.

Full integration pinned to `10612c66`, through L29 and before this mixed-face
repair, passes 2,159 Windows checks with no failures and one unconfigured
private-layout probe skipped, including 878 DOCX checks. Its fresh 0.1.5 package
contains the exact full-suite DLL. Evidence is under
`artifacts/plan-revision-20261005/milestone-10612c66/`.

## RV06-L33: wrapped two-face comment tails

Plain regular comments with exactly two prepared source-face parts now retain
the first-row prefix face and wrap the tail in its own face. The prefix must
fit beside the title; every tail word, including its separating space, must
fit the applicable row. Tail font metrics drive both baseline pitch and balloon
height. Styled, grouped, compound, leading-space and individually overwide
paths retain their previous behavior. No font is created during emission.

The production reproducer fails before the repair and passes afterward,
checking source CIDs, complete content, first-row advances, continuation breaks,
pitch and height. All 880 DOCX checks and 77 balloon checks pass, together with
three Linux regressions and a clean Release build. Fresh 0.1.5 package smoke
contains the exact tested DLL from runtime revision `d420fdec`:
`59B3C3B2F3139289B6F30EE12CFF41F42186AF2C893B5C3A4EB104A7C684A556`.

Across 68 Office controls, seven improve error and similarity, 61 retain PDF
and graphics bytes, and all retain main-document text. The earlier long
Arial/Courier control now matches twelve Office rows rather than nine; its
144dpi mean error falls from 1.085 to 0.401 and similarity rises from 0.687 to
0.915. That control and its fresh counterpart cover the same text/font mechanism,
with separate recorded input and reference identities.

All six fresh long font pairs preserve Office faces and body content and
improve pixels; five match line breaks. Cambria/Georgia still produces nine
rows where Office uses ten. Among the five matching controls, horizontal gaps
reach 0.106pt and baseline gaps reach 0.483pt on the longer Courier/Arial body.
These are explicit residuals, alongside mixed final paragraph-mark faces and
more general composition. All five fallback guards retain PDF bytes, as do
all 34 cached cases with 551 failed gates. Evidence, including per-control
hashes and the separate face/content and row audits, is under
`artifacts/plan-revision-20261005/rv06-l33/`.

## RV06-L34: continuation-width probe limits

Twelve fresh Office controls vary Georgia, Courier New and Aptos tails across
72/108/144pt right margins, with uniform companions. Georgia loses one row in
all four applicable controls; Courier and Aptos retain Office row counts.
A symmetric right text inset is rejected: it would wrap three currently
fitting Aptos controls too early. The observed inset scales with the print
profile, but these bounds do not qualify a shared replacement width rule.
Runtime remains unchanged. Input/reference identities, body-content checks
and numeric width constraints are under
`artifacts/plan-revision-20261005/rv06-l34/`.

Full integration at `ca949c0e`, through L32 and the locked-field companion but
before L33, passes 2,160 Windows checks with no failures and one unconfigured
private-layout probe skipped, including 879 DOCX checks. Release is clean;
fresh 0.1.5 package smoke contains the exact full-suite DLL:
`620E02E62AA586027117F7977982EA3E27E06594F4652331476050408B244C5C`.
Evidence is under `artifacts/plan-revision-20261005/milestone-ca949c0e/`.

## RV06-L35: final paragraph-mark faces in mixed comments

Plain two-run comments now retain a prepared paragraph-mark space. The resolved
mark face supplies the final blank when the existing single-row or two-face
continuation gate admits the body; intermediate spaces retain the tail face.
Overwide or otherwise unqualified mixed paths still emit the legacy final
blank. Grouped, styled and compound composition remains outside this repair.

The reader-to-renderer reproducer records both failures: a dropped mark, then
legacy-face emission after reader retention. It now passes short and wrapped
prefix/tail mark faces at a declared 18pt size, together with overwide fallback.
All 881 DOCX checks, 78 balloon checks and three Linux regressions pass.
Release is clean; fresh 0.1.5 package smoke contains the exact tested DLL from
runtime revision `f058786b`:
`5DD7607E9260486B9C3025B586B171B8ECE4879B8F63F44CD7C6D35B5446FF64`.

All 92 Office controls retain visible text state, geometry, graphics, rasters
and pixel metrics. Thirty-two final blank faces change, while 56 PDFs retain
their bytes. Nine fresh admitted controls match Office's final-space face,
including the 18pt mark rendered at nominal balloon size. Three fresh fallback
guards retain PDFs. Four earlier fallback controls additionally register an
unused Aptos subset, adding 7,956 bytes each while preserving all operations
and visible faces. This preparation/serialization overhead is explicit.

All 34 cached PDFs retain bytes and 551 failed gates. Georgia's existing wrap
difference, long-run baseline drift and composed pixel residuals remain.
Evidence, final-blank and font-resource audits, all identities and exact package
proof are under `artifacts/plan-revision-20261005/rv06-l35/`.

## RV06-L36: serialize only used DOCX page fonts

DOCX pages now serialize embedded font resources referenced by surviving text
operators. Content rollback also rewinds newly used font names, including nested
marks and reuse of an earlier face. Font discovery, loading, preparation and CID
coverage remain unchanged, as do standard fallback registration and PPTX resource
policy. This reduces PDF bytes without claiming a prepared-font memory reduction.

The production reproducer covers two page-specific faces and an orphan comment
face; nested rollback verifies both glyph and fallback text. All 882 DOCX,
79 balloon, 65 PDF and 127 image checks pass, together with four Linux regressions.
Release is clean. Fresh 0.1.5 package smoke contains the exact tested DLL from
runtime revision `92fd1809`:
`499D4993FEADA9E7C0FD70246DB9B609735ECC22465CEC635C66DE1F5966F7CA`.

All 92 Office controls retain exact text operations, graphics, rasters and
surviving font metadata, excluding reassigned PDF object identifiers. Removing
104 unused font declarations saves 1,868,575 bytes across this recorded corpus,
with a maximum of 44,023 bytes per document. One PDF retains its bytes; the other
91 change only through removal of unused declarations and consequent object
numbering. Related controls do not constitute independent mechanism counts.

All 34 cached cases retain exact text operations and surviving font metadata,
and all 37 raster pages retain bytes. Eighteen PDFs retain bytes; removing
22 declarations from the remaining PDFs saves 172,611 bytes. Failed gates stay
at 551 with no case increase. Georgia wrapping, long baseline drift and composed
pixel residuals remain. Hashes, audits and package proof are under
`artifacts/plan-revision-20261005/rv06-l36/`.

Full integration pinned to `facef6ad`, through L35 and before this resource
repair, passes 2,162 Windows checks with no failures and one unconfigured
private-layout probe skipped, including 881 DOCX checks. Release is clean;
the fresh 0.1.5 package contains the exact full-suite DLL:
`503BA96F0E64352CFAB842776BE7313C0574CBCC4C9AF6144795B57382B4B376`.
Evidence is under `artifacts/plan-revision-20261005/milestone-facef6ad/`.

## RV06-L37: continuation words and separating spaces

Word-compatible continuations now choose breaks using visible words and allow
the emitted trailing blank into the inset. A measured 4.42pt design inset on each
side gives 221.36pt of visible width in the existing 230.2pt body. First-row
admission retains its break-space reserve. Preparation and font coverage stay
unchanged; height and emission use the same row selection.

Eighteen coarse and nineteen distinct fine Office controls separate Georgia,
Arial and Aptos boundaries. Reserving the blank gives incompatible font bounds;
excluding it gives a common 221.331..221.405pt design interval. Twenty-four
additional margin controls retain the chosen bound across three print scales.
The earlier half-inset proposal improves coarse controls but fails eight fine
controls, so it is rejected.

The portable production regression covers narrow and wide separating spaces,
four print scales and uniform/two-face composition. It fails before the repair
and passes afterward. All 883 DOCX checks, 79 balloon checks, the new wrap check
and four Linux regressions pass. Release is clean; fresh 0.1.5 package smoke
contains the exact tested DLL from runtime revision `f4a2cc44`:
`235B807E9D70A408E06ADAC41C76EEA1BB468A984D5CA90ADE79552E37980366`.

Across 159 Office controls, 31 improve error and similarity, 128 retain PDF and
raster bytes, 152 retain graphics, and all retain main-document text. All 61 new
boundary controls match Office breaks; the earlier Georgia long bodies now do
too. Seventy controls pass row-level source-face/content audits. Audited row-start
X gaps reach 0.838pt and baseline gaps reach 0.177pt; unchanged longer controls
retain their previously recorded drift.

Six whole-word guards retain PDF bytes. Four remain partial because Office splits
overwide Courier words while the current renderer keeps its existing behavior.
Seven older styled/grouped/overwide controls retain break mismatches. All 34
cached PDFs retain bytes and 551 failed gates; composed pixel residuals remain.
Inputs, references, accepted-baseline links, rejected-proposal evidence and final
audits are under `artifacts/plan-revision-20261005/rv06-l37/`.

## RV21-V26: full integration through L37

The isolated Release run at `c6a0b902` passes 2164 checks, fails none and skips
the unconfigured private-layout diagnostic. It includes 883 DOCX checks and
excludes the isolated SVG transparency proposal. Release has no warnings or
errors. Fresh package smoke keeps version 0.1.5 and verifies the exact full-suite
DLL: `FB3BBCA3CBA6D615E9ACEDD091C5EA91C222EA50FC573A7193D9D229BCE22C30`.
Evidence is under `artifacts/plan-revision-20261005/milestone-c6a0b902/`.

## RV06-L38: uniform balloon word splitting

Prepared uniform regular comment faces now split overwide words at Unicode
scalar boundaries. The first-row and continuation widths retain their qualified
rules; within-word breaks emit no invented spaces. Word-boundary separators and
the final paragraph-mark face remain explicit. Height and emission use the same
rows. Mixed, styled, grouped and unprepared fallback branches retain their
previous admission rules.

The production regression fails before the repair and passes for ASCII and
supplementary Unicode in both first-row and continuation lanes. A separate
single-glyph guard preserves progress and spaces even when one glyph exceeds
the lane. All 885 DOCX checks, 81 balloon checks and six Linux regressions pass.
Release is clean; fresh 0.1.5 package smoke contains the exact tested DLL from
runtime revision `4f3812e1`:
`C68D720A22D099E35038366027B667B8205AA47A9608B3FECF819E847326940B`.

Across 159 held-out and three new first-row Office controls, five improve error
and similarity, 157 retain PDF/raster bytes, 161 retain graphics and all retain
main-document text and comment content. The longer first-row control now uses
Office's three rows rather than two. Six uniform controls match exact row text,
spaces and source faces, including six within-word breaks; maximum row-start X
gap is 0.085pt and baseline gap is 0.170pt. All 34 cached PDFs retain bytes and
551 failed gates.

Two mixed-source whole-word controls retain fallback break mismatches, together
with seven older styled/grouped/overwide controls. Prior positioning and composed
pixel residuals remain; 33 cached cases are partial. Input/reference/candidate
hashes and exact emission audits are under
`artifacts/plan-revision-20261005/rv06-l38/`.

## RV06-L39: mixed continuation word splitting

A qualified two-face comment can now split overwide continuation words in the
prepared tail face. The prefix and first tail word retain their admission gates;
within-word breaks share the uniform splitter and emit no invented separators.
The tail face supplies continuation pitch and the resolved paragraph-mark face
supplies the final blank. Height and emission use the same rows, with cancellation
checks during scanning and emission.

The production reproducer fails before the repair and passes for ASCII and
supplementary Unicode at two print scales, while checking distinct source faces,
complete content, tail metrics, final mark size/face and box height. All 886 DOCX
checks, 82 balloon checks and seven Linux regressions pass. Release is clean;
fresh 0.1.5 package smoke contains the exact tested DLL from runtime revision
`87c518b9`:
`6F7EFCE649C1A2BB2D46FE86D7F22979FE19D99EA6ABF03DDC38227A8F5CF6DD`.

Across 162 held-out and nine new Office controls, nine improve error and similarity
and 162 retain PDF, raster and graphics bytes. All retain main-document text and
comment content. Ten audited controls match exact row text, spaces and source
faces, including twelve within-word breaks; maximum row-start X gap is 0.085pt
and baseline gap is 0.177pt. All 34 cached PDFs retain bytes and 551 failed gates.

Six older fallback controls and three new first-word guards retain break
mismatches. The latter expose a word spanning font runs: Office can move the last
prefix word into the continuation row before splitting the tail. This requires
separate row composition and metric evidence. Styled/grouped/unprepared fallbacks,
prior positioning and composed pixel residuals remain. Evidence is under
`artifacts/plan-revision-20261005/rv06-l39/`.

## RV06-L40: a joined first word across source faces

A fitting ASCII letter/digit prefix can now share an overwide first word with its
prepared tail face. The tail splits at Unicode scalar boundaries using the
qualified first-row and continuation widths. Prefixes with spaces or other break
opportunities retain first-tail-word admission; the remaining two-face gates,
including multiple tail words, stay intact. Height, continuation metrics, source
CIDs, final mark face and cancellation use the existing qualified paths.

The production reproducer fails before repair and passes for ASCII and
supplementary Unicode at two print scales. It checks source-face fragments,
complete text with no invented spaces, first-row prefix advance, tail pitch,
final mark size/face and height, plus composed-prefix fallback. All 887 DOCX
checks, 83 balloon checks and eight Linux regressions pass. Release is clean;
fresh 0.1.5 package smoke contains the exact tested DLL from runtime revision
`ccaa4809`:
`2AFA896FFBF90A2C2F206955E65E1F3A0E2C5888200E04728D5B0B231D286E8B`.

All 171 prior Office controls retain PDF, raster and graphics bytes. Of seven new
probes, five joined-word controls improve error and similarity and match Office
breaks across Courier New, Georgia, Arial and Aptos. The two preserved-separator
guards retain PDFs. All 178 controls retain main-document text and comment
content. The five repaired controls match exact row text, spaces and source faces,
including twelve within-word breaks; maximum row-start X gap is 0.085pt and
baseline gap is 0.177pt. All 34 cached PDFs and 551 failed gates stay unchanged.

Joined multi-word prefixes, separated prefixes and the prior fallback/positioning
and composed-pixel residuals remain. The separated-word controls show a different
first continuation pitch, requiring font-transition evidence before admission.
Ten controls retain break mismatches, and 33 cached cases remain partial. Evidence
is under `artifacts/plan-revision-20261005/rv06-l40/`; invalid unqualified fixture
versions from an XML namespace generation error are retained separately from the
trusted corrected input identities.

## RV21-V27: full integration through L39

The isolated Release run at `b2a74995` passes 2167 checks, fails none and skips
only the unconfigured private-layout diagnostic. It includes 886 DOCX checks and
excludes the isolated SVG transparency proposal. Release has no warnings/errors.
Fresh 0.1.5 package smoke verifies the exact full-suite DLL:
`48A0F74DC351C6675858E6AE5350BD37161E03B0F9504DCCE036708CD0725441`.
Evidence is under `artifacts/plan-revision-20261005/milestone-b2a74995/`.

## RV06-L41: separated prefix and first continuation metrics

A fitting single ASCII word prefix with a preserved trailing separator can now
place an overwide first tail word below it in the prepared source face. The first
continuation uses the prefix face's horizontal descent plus the tail face's
horizontal ascent and line gap. Later continuations retain the tail height.
Height and emission share these distinct steps; joined multi-word prefixes and
other admission/fallback rules remain unchanged.

Twelve trusted Office controls cover nine font pairs and three print scales.
Among the measured alternatives, the selected first-pitch model has maximum
error 0.156pt/mean 0.043pt at Office font sizes; using the tail height alone reaches
0.590pt. Including the previous font's line gap performs worse and is rejected.
The production regression fails before repair and passes for ASCII and
supplementary Unicode at three scales with distinct descent/line-gap metrics,
complete text, final mark face/size and box height. All 888 DOCX checks,
84 balloon checks and nine Linux regressions pass. Release is clean; fresh 0.1.5
package smoke contains the exact tested DLL from runtime revision `fc0f4fe0`:
`E9D305077EEFD4481CED65B83E2F816BB03DAC56CF3AC4E23124EBDE6FA7ECD8`.

Across 178 held-out and ten new Office controls, twelve improve error and
similarity, 176 retain PDF/raster/graphics bytes and all retain main-document
text and comment content. All twelve repaired controls match exact row text,
spaces and source faces, including seventeen within-word breaks. Maximum
row-start X gap is 0.838pt at one print scale; maximum baseline gap is 0.187pt.
All 34 cached PDFs retain bytes and 551 failed gates.

Nine earlier fallback/composed-prefix controls retain break mismatches. Rich,
grouped and unprepared paths, prior positioning/composed pixel residuals and 33
partial cached cases remain. Input/reference/candidate identities, selected and
rejected metric models and exact emission audits are under
`artifacts/plan-revision-20261005/rv06-l41/`.

## RV06-L42: joined two-word prefixes and composed row metrics

A fitting plain two-word ASCII prefix can now move its final word onto the
first split tail row in its prepared font. The first step uses prefix descent
plus the larger complete ascent/line-gap metric of the two faces. The second
uses their larger descent plus tail ascent/line gap; later steps use tail height.
Height and emission share these phases. The mixed continuation fits visible
words without reserving the emitted break separator. Existing single-word and
separated-prefix paths, styled/unprepared guards and cancellation remain active.
Wider, punctuated or longer prefixes retain their prior admission rules.

Sixteen trusted metric controls cover ten font pairs and three print scales.
The selected first-step model has maximum error 0.108pt/mean 0.034pt at Office
font sizes; the second has maximum 0.156pt/mean 0.035pt. Maximizing ascent and
line gap separately reaches 1.179pt and is rejected. Two word-boundary probes
confirm that the mixed continuation excludes the emitted blank from visible
width. The production regression fails before repair and passes for ASCII and
supplementary scalars at three scales with distinct ascent/descent/gap metrics,
complete source text, source faces, terminal mark/size and shared box height.

All 889 DOCX checks, 85 balloon checks and ten Linux regressions pass. Release
has no warnings/errors. Fresh 0.1.5 package smoke contains the exact tested DLL
from runtime revision `2530bcb9`:
`437102382717564CD9816F3597D23D36959670C446F52837E00CEB2421C5A20D`.
Across 188 earlier and twelve new Office comparisons, sixteen improve error
and similarity, 184 retain PDF/raster/graphics bytes and all retain main-document
text and comment content. These are 200 distinct input hashes and 199 case IDs;
the reused ID has identical candidate bytes. All sixteen changed comparisons
match exact row text, spaces and source faces, including 27 within-word breaks.
Maximum row-start X gap is 0.838pt and baseline gap is 0.201pt.

All 34 cached PDFs retain bytes and 551 failed gates. Six held-out break
mismatches remain across leading tail separators, wide prefixes, three parts
and multiple paragraphs. Prior positioning/composed pixel residuals and 33
partial cache cases remain. Input/reference/candidate identities, selected and
rejected models and exact emission audits are under
`artifacts/plan-revision-20261005/rv06-l42/`.

## RV21-V28: full integration through L41

The isolated Release run at `d6ec0d8e` passes 2169 checks, fails none and skips
only the unconfigured private-layout diagnostic. It includes 888 DOCX checks
and excludes the isolated SVG transparency proposal. Release has no warnings
or errors. Fresh 0.1.5 package smoke verifies the exact full-suite DLL:
`A799FC52156F4D34A6031C41C5F058E669893787222225DFB18925E29861AD40`.
Evidence is under `artifacts/plan-revision-20261005/milestone-d6ec0d8e/`.

## RV06-L43: leading tail separator and mixed descent

A fitting plain two-word ASCII prefix followed by one leading tail separator
and a fitting first tail word can now wrap in the prepared source faces. The
separator retains its tail face and consumes that face's advance before fitting
the first word. The first continuation uses the larger descent of the two faces
plus tail ascent/line gap; later rows use tail height. Height and emission share
these steps. Overwide first words, doubled separators and prior admission,
styled/unprepared and cancellation guards retain their behavior.

Thirteen trusted metric controls cover ten font pairs and three print scales.
The selected first-step model has maximum error 0.156pt/mean 0.064pt at Office
font sizes; using tail height alone reaches 0.442pt. Using only prefix descent
reaches 0.541pt and is rejected. The production regression fails before repair
and passes for ASCII and supplementary scalars at three scales with distinct
prefix-space/ascent/descent/gap metrics, exact source text, tail-space face,
later word splits, terminal mark/size and box height. All 890 DOCX checks,
86 balloon checks and eleven Linux regressions pass. Release is clean; fresh
0.1.5 package smoke contains the exact tested DLL from runtime `bdd36a67`:
`DFAFEA493173DD815093651AC34FDE8F39F788ED3058ADD31C663611F02AF91D`.

Across 200 earlier and thirteen new Office comparisons, thirteen improve error
and similarity, 200 retain PDF/raster/graphics bytes and all retain main-document
text and comment content. All thirteen changed comparisons match exact row
text, spaces and source faces. The comparison now retains separately emitted
whitespace before grouping rows, avoiding a mismatch caused solely by different
text-operator segmentation. All 213 comparisons were reconciled against the
same candidate/reference hashes; the exact space/face audit is independent.
Maximum row-start X gap is 0.838pt and baseline gap is 0.197pt.

All 34 cached PDFs retain bytes and 551 failed gates. Four earlier break
mismatches remain across wide prefixes, three parts and multiple paragraphs;
two new leading-overwide-first-word guards also retain fallback. Prior
positioning/composed pixel residuals and 33 partial cache cases remain. Evidence
and input/reference/candidate identities are under
`artifacts/plan-revision-20261005/rv06-l43/`.

## RV21-V29: full integration through L43

The isolated Release run at `972fc7c9` passes 2171 checks, fails none and skips
only the private PPTX diagnostic whose input/output environment is absent.
All 890 DOCX checks pass. The clean build and fresh 0.1.5 package contain the
exact full-suite DLL:
`A6EE1A4BBD0F82F3849DBB14CAF7D6D943A7269A3DDEFB6EDF2CFAB622E57798`.
This run excludes E4 transparency and L44. Evidence is under
`artifacts/plan-revision-20261005/milestone-972fc7c9/`.

## RV06-L44: multiple prefix rows before a joined tail

An ordinary ASCII word prefix that exceeds the first row can wrap in its
prepared face before its last word joins a fitting first tail word. Each prefix
word must fit continuation width; an overwide prefix word or first tail word
retains fallback. The first prefix row fits visible words without reserving the
emitted break separator. Prefix-only and tail-only rows use their own metrics.
The incoming mixed row uses prefix descent plus the larger full ascent/line-gap
pair; the outgoing step uses mixed descent plus tail ascent/line gap. Height and
emission share these phases. The source font boundary invents no separator.

Seventeen trusted metric controls cover ten font pairs and three print scales.
Incoming/outgoing maximum prediction errors are 0.131/0.134pt. Using prefix or
tail height alone, or separate maxima of ascent and line gap, is rejected where
it disagrees with those measurements. Sixty-two prefix-only and 131 tail-only
pitches have maximum errors of 0.082/0.110pt. A production regression was red
before admission; an enlarged-space variant was red before the first-row
correction. Seventy-two scalar/gap/space/scale/tail variants and fallback guards
check text, source faces, phase spacing, final mark/size and height. All 891 DOCX
checks, 87 balloon checks and twelve Linux regressions pass. Release is clean;
fresh 0.1.5 package smoke contains the exact tested DLL at runtime `7eb675b2`:
`7F8F2D3CF07A707B415E3781C894B83BDD0BF12E26761C0B21A574237082315E`.

Across 213 earlier and seventeen new comparisons, seventeen improve MAE and
SSIM, 213 retain PDF/raster/graphics bytes and all retain main-document text and
comment content. All seventeen changed controls match exact row text, spaces
and source faces, including the two older Arial controls fixed by excluding the
first prefix row's emitted blank from width reservation. Those two controls
improve MAE from 1.164 to 0.593 versus the initial L44 prototype; all seventeen
new-control PDFs retain their bytes through that correction. Maximum row-start
X gap is 0.838pt and baseline gap is 0.253pt. Actual Office break matches reach
224 of 230 inputs; separately emitted spaces remain part of the audit.

All 34 cached PDFs retain bytes and 551 failed gates. Six break mismatches remain:
three-part and multi-paragraph comments, two leading-overwide-first-word controls
and two new overwide-word guards. Prior positioning/composed pixel residuals and
33 partial cached cases remain. Corrected qualification and identities are under
`artifacts/plan-revision-20261005/rv06-l44-first-row/`; the initial prototype and
rejected first-row reservation evidence remain under `rv06-l44/` beside it.


## RV06-L45: separator-only prefix row before the first tail word

A fitting plain two-word ASCII prefix followed by one authored tail separator
can now keep that separator in its prepared face while the first word starts
below the prefix. Whole and overwide words retain the existing scalar-safe tail
wrapper. An explicitly empty first tail row prevents even a zero-advance first
scalar from painting beside the prefix. The separator itself still consumes its
own advance on that row; no separator is invented inside a split word.

Fourteen trusted metric controls cover ten font pairs and three print scales.
When only a separator shares the prefix row, the first step uses visible prefix
descent plus tail ascent/line gap. Maximum error is 0.131pt/mean 0.042pt. Including
tail descent reaches 0.621pt/mean 0.352pt and is rejected; that mixed-descent rule
still applies to the earlier fitting visible-tail-word case. Twenty-six later
tail-only pitches have maximum error 0.082pt. Both source-face admission and
distinct-descent regressions fail before their repairs. Forty-eight numeric
scalar/zero-first-advance/descent/whole-and-split/scale variants check exact text,
separator face/advance, row spacing, final mark/size, visible width and height.

All 892 DOCX checks, 88 balloon checks and thirteen Linux regressions pass.
Release is clean; fresh 0.1.5 package smoke contains the exact tested DLL at
runtime `e0647231`:
`D2BA3A419DC2A6DDFB6351633B85A8A7686CB3A8F8DE54370D4043BC227B6987`.
Across 230 earlier and fourteen new comparisons, fourteen improve MAE and SSIM,
230 retain PDF/raster/graphics bytes and all retain main-document text and
comment content. All fourteen changed controls match exact row text, spaces and
source faces, including eighteen within-word breaks. Maximum row-start X gap is
0.838pt and baseline gap is 0.207pt; the rejected mixed-descent prototype reaches
0.810pt baseline gap. Actual Office break matches reach 238 of 244 inputs.

All 34 cached PDFs retain bytes and 551 failed gates. Six break mismatches remain:
three-part and multi-paragraph comments, two earlier wide-word guards and the
new doubled-separator/punctuation guards. Prior positioning/composed pixel
residuals and 33 partial cached cases remain. Qualification, input/reference/
candidate identities and the rejected mixed-descent counterfactual are under
`artifacts/plan-revision-20261005/rv06-l45/`.


## RV21-V30: full integration through L44 and SVG transparency

The isolated Release run at `317a694d` passes 2173 checks, fails none and skips
only the private PPTX diagnostic whose input/output environment is absent.
All 891 DOCX checks pass. The clean build and fresh 0.1.5 package contain the
exact full-suite DLL:
`5540144785238FB224C00CA22697B49D881F76DE7D40644BB4F0F68395FB172F`.
This run includes corrected L44 and accepted smooth-preview SVG transparency E4;
it excludes L45/L46. Evidence is under
`artifacts/plan-revision-20261005/milestone-317a694d/`.

## RV06-L46: reflow before splitting a joined first tail word

An ordinary ASCII prefix can now wrap in its prepared face before its first
joined tail word splits across rows. If the final prefix row contains multiple
words, its last word moves onto a fresh row before tail splitting. The authored
separator remains at the preceding prefix row's end in its prefix face; no space
is invented at the font boundary or inside the joined word. The first tail
scalar must fit the mixed row and have positive advance, and every scalar must
fit continuation width. Oversized/zero-first scalars and overwide prefix words
retain fallback. Existing prefix-only/incoming/outgoing/tail-only metrics and
the final paragraph mark remain shared by height and emission.

Fourteen metric controls cover ten font pairs and three print scales. Incoming
and outgoing maximum errors are 0.131/0.083pt; rejected prefix/tail-only or
separate-max-component models retain their counterfactuals. Seventy prefix-only
and seventeen tail-only pitches have maximum errors of 0.082/0.054pt. Admission
and multiword final-prefix-row regressions fail before their repairs. The final
regression checks 144 scalar/gap/space/prefix-count/scale/tail variants plus
glyph/fallback guards, exact text, prepared faces, all spacing phases, visible
width, final mark/size and height.

All 893 DOCX checks, 89 balloon checks and fourteen Linux regressions pass.
Release is clean; fresh 0.1.5 package smoke contains the exact tested DLL at
runtime `e0699172`:
`141C6FAB3F418644507BE9758954BFC8325BDE11A18D8F115AFB2DD9100D2702`.
Across 244 earlier and fourteen new comparisons, fourteen improve MAE and SSIM,
244 retain PDF/raster/graphics bytes and all retain main-document text and
comment content. All fourteen changed controls match exact row text, spaces and
source faces, including 28 within-word breaks. Maximum row-start X gap is 0.838pt
and baseline gap is 0.221pt. Actual Office break matches reach 252 of 258 inputs.

All 34 cached PDFs retain bytes and 551 failed gates. Six break mismatches remain:
three-part and multi-paragraph comments, the earlier overwide-prefix-word guard,
two L45 doubled-separator/punctuation guards and the new oversized-prefix guard.
Prior positioning/composed pixel residuals and 33 partial cached cases remain.
Qualification, identities and the rejected unreflowed-prefix prototype are under
`artifacts/plan-revision-20261005/rv06-l46/`.


## RV06-L47: an overwide prefix word in its prepared face

A single overwide ASCII prefix word now splits in its prepared font before its
joined tail begins. Hard breaks invent no spaces. The first prefix glyph must
have positive advance and fit the first row; every prefix glyph must fit the
continuation width. Existing tail-glyph bounds, four spacing phases and nominal
paragraph-mark size remain shared by height and emission. Punctuation and
non-ASCII prefix words retain fallback.

Fourteen Office metric controls cover ten font pairs and three print scales.
Incoming/outgoing maximum errors are 0.131/0.124pt. Prefix-only and tail-only
pitch errors are at most 0.060/0.082pt across 27/121 measurements. Rejected
single-face and separate-component models remain recorded. The production
regression fails before admission, then checks 72 scalar/gap/length/scale/tail
variants, zero or oversized first prefix/tail glyphs and unsupported prefix
guards. Assertions preserve exact content, all spacing phases, source fonts,
visible widths, final mark and matching balloon height.

All 894 DOCX checks, 90 balloon checks and fifteen Linux regressions pass.
Release is clean; fresh 0.1.5 package smoke contains the exact tested DLL at
runtime `4a59bb73`:
`ED62397274F86AB85121EA09E109322B3D27155D7C9DD79BA1461600349BBCC3`.
Across 258 earlier and fourteen new inputs, fifteen improve MAE and SSIM,
257 retain PDF/raster/graphics identity and all retain main-document text and
comment content. All fifteen changed controls match exact row text, spaces and
source fonts, including 44 within-word breaks. Maximum row-start X gap is
0.838pt and baseline gap is 0.233pt. Office break matches reach 267 of 272 inputs.

All 34 cached PDFs retain bytes and 551 failed gates. Five break mismatches
remain: three-part comments, multiple paragraphs, a doubled tail separator and
two punctuation-prefix guards. Earlier positioning/composed pixel residuals
and 33 partial cached cases remain. Qualification and identities are under
`artifacts/plan-revision-20261005/rv06-l47/`.

## RV06-L48: two authored leading tail spaces

Two leading tail spaces now retain their prepared font and advance beside a
fitting plain two-word ASCII prefix. The available first-row width reserves
both spaces. An overwide first word starts below the prefix; a fitting word
shares its row. Separator-only rows use prefix descent for the first step,
while visible mixed rows use both descents. Later rows retain tail metrics.
Three leading spaces, repeated internal spaces and punctuated prefixes retain
fallback. The exception to repeated-space admission is limited to this shape.

Fourteen Office phase controls cover ten font pairs and three print scales:
twelve separator-only first rows and two visible mixed first rows. The selected
first-step error is at most 0.131pt; using one descent rule for every case has
larger errors. Twenty-eight tail-only pitches have maximum error 0.082pt.
The production reproducer fails before admission. The final regression checks
48 scalar/zero-first-glyph/prefix-descent/length/scale variants, expanded fitting
tail cases and fallback guards, preserving exact spaces, fonts, advances,
baseline steps, visible widths, nominal final mark and matching height.

All 895 DOCX checks, 91 balloon checks and sixteen Linux regressions pass.
Release is clean; fresh 0.1.5 package smoke contains the exact tested DLL at
runtime `28e66b48`:
`81CC4C76DAECDB7B826670A4164F06FC9A8E518E11E1EB4DE8E821660871B617`.
Across 272 earlier and fourteen new inputs, fourteen improve MAE and SSIM,
272 retain PDF/raster/graphics identity and all retain main-document text and
comment content. All fourteen changed controls match exact row text, spaces
and source fonts, including twenty within-word breaks. Maximum row-start X gap
is 0.838pt and baseline gap is 0.207pt. Office break matches reach 281 of 286.

All 34 cached PDFs retain bytes and 551 failed gates. Five break mismatches
remain: three-part comments, multiple paragraphs, two punctuation-prefix guards
and the new three-space guard. Earlier positioning/composed pixel residuals
and 33 partial cached cases remain. Qualification and identities are under
`artifacts/plan-revision-20261005/rv06-l48/`.

## RV21-V31: full integration through L46 and SVG transparency

The isolated Release run at `c0458d1a` passes 2175 checks, fails none and skips
only the private PPTX diagnostic whose input/output environment is absent.
All 893 DOCX checks pass. The clean build and fresh 0.1.5 package contain the
exact full-suite DLL:
`8CA4C0372C9B4A2BF59DF7CB5064957EFBAC078A405B26A5F3A5851221945C37`.
This run includes L45/L46 and accepted smooth-preview SVG transparency E4;
it excludes L47/L48. Evidence is under
`artifacts/plan-revision-20261005/milestone-c0458d1a/`.

## RV06-L49: a terminal comma before a separated tail

A fitting two-word ASCII prefix ending in one comma now retains its prepared
font before one leading tail space. The comma stays with the prefix, the
authored separator stays with the tail, and a fitting first tail word shares
the row while a larger word starts below it. Existing separator-only/mixed
descent rules and later tail metrics remain shared by height and emission.
Colons, internal commas, punctuation-only words and doubled separators beside
a punctuated prefix retain fallback. Wide punctuated prefixes remain outside
this admission rule.

Fourteen Office phase controls cover ten font pairs and three print scales:
twelve separator-only first rows and two visible mixed first rows. The selected
first-step error is at most 0.131pt; twenty-eight tail-only pitches have maximum
error 0.082pt. Rejected descent models remain recorded. The production
reproducer fails before admission. The final regression checks 96 scalar,
zero-first-glyph, prefix-descent, length, scale and fitting-first-word variants
plus unsupported-prefix/separator guards. Exact text, fonts, advances, spacing
phases, visible widths, nominal final mark and matching height remain asserted.

All 896 DOCX checks, 92 balloon checks and seventeen Linux regressions pass.
Release is clean; fresh 0.1.5 package smoke contains the exact tested DLL at
runtime `fb521bdb`:
`A5AEAC236471479ADB25465D4CDDCF35A9DA33AA11FE64A0C8DB672061802A18`.
Across 286 earlier and fourteen new inputs, fourteen improve MAE and SSIM,
286 retain PDF/raster/graphics identity and all retain main-document text and
comment content. All fourteen changed controls match exact row text, spaces
and source fonts, including twenty within-word breaks. Maximum row-start X gap
is 0.838pt and baseline gap is 0.207pt. Office break matches reach 295 of 300.

All 34 cached PDFs retain bytes and 551 failed gates. Five break mismatches
remain: three-part comments, multiple paragraphs, an overwide punctuated prefix
word, the three-space guard and the new colon guard. Earlier positioning/
composed pixel residuals and 33 partial cached cases remain. Qualification and
identities are under `artifacts/plan-revision-20261005/rv06-l49/`.

## RV21-V32: full integration through L48 and SVG transparency

The isolated Release run at `b7d25b43` passes 2177 checks, fails none and skips
only the private PPTX diagnostic whose input/output environment is absent.
All 895 DOCX checks pass. The clean build and fresh 0.1.5 package contain the
exact full-suite DLL:
`1C3ED748DA24274AEC3742EE56AD804D37CA7BC8A1E5710D4E89AD7B79A7283E`.
This run includes L47/L48 and accepted smooth-preview SVG transparency E4;
it excludes L49/L50. Evidence is under
`artifacts/plan-revision-20261005/milestone-b7d25b43/`.

## RV06-L50: an overwide comma-ending prefix word

A single overwide ASCII prefix word ending in one comma now splits in its
prepared font before its joined tail. The comma stays on the final prefix
fragment and hard breaks invent no spaces. First and continuation glyph bounds,
four spacing phases and nominal final mark remain shared by height and emission.
Wrapped multiword punctuation, repeated commas, other punctuation and non-ASCII
prefix words retain fallback.

Seventeen Office metric controls cover ten font pairs and three print scales,
including three overwide-first-tail-word probes. Those probes confirm that the
tail can split beside the final comma-ending prefix fragment. Incoming/outgoing
maximum errors are 0.131/0.124pt; 33 prefix-only and 124 tail-only pitches have
maximum errors 0.060/0.082pt. Rejected single-face and separate-component models
remain recorded. The production reproducer fails before admission, then checks
72 scalar/gap/length/scale/tail variants, zero/oversized first prefix and tail
glyphs and unsupported-prefix guards. Earlier comma-only guards now use
semicolons. Exact text, fonts, all spacing phases, visible widths, final mark
and matching height remain asserted.

All 897 DOCX checks, 93 balloon checks and eighteen Linux regressions pass.
Release is clean; fresh 0.1.5 package smoke contains the exact tested DLL at
runtime `aebca030`:
`E9983F4880802BB3305C99C90BD5C39AEBE5163C07B25B8661E3B66632AFF796`.
Across 300 earlier and seventeen new inputs, seventeen improve MAE and SSIM,
300 retain PDF/raster/graphics identity and all retain main-document text and
comment content. All seventeen changed controls match exact row text, spaces
and source fonts, including 56 within-word breaks. Maximum row-start X gap is
0.838pt and baseline gap is 0.233pt. Office break matches reach 312 of 317 inputs.

All 34 cached PDFs retain bytes and 551 failed gates. Five break mismatches
remain: three-part comments, multiple paragraphs and the three-space, colon and
new semicolon guards. Earlier positioning/composed pixel residuals and 33
partial cached cases remain. Qualification and identities are under
`artifacts/plan-revision-20261005/rv06-l50/`.

## RV06-L51: authored leading tail space spans

Word preserves a fitting run of leading tail spaces beside a plain two-word
ASCII prefix. Admission now counts the complete authored span and reserves its
prepared advance; the span must leave positive first-row width. Separator-only
first rows use the visible prefix descent, while a visible tail word uses both
descents. Later rows keep tail metrics. Internal repeated spaces, punctuation,
tabs and overflowing spans retain the complete fallback.

The regression fails before admission and then covers 384 scalar,
zero-first-glyph, space-count, descent, length, scale and fitting-first-word
variants with explicit fallback guards. Fourteen mixed checks, all 898 DOCX
checks, 94 balloon checks and nineteen Linux regressions pass. Release has no
warnings or errors. Fresh 0.1.5 package smoke contains the exact tested DLL at
runtime `b6e59c96`:
`DCF6350A303972B248596B2DEC382A623B37074FBB5C7B8AB755398E300DD6DB`.

Across 317 prior and eighteen fresh inputs, seventeen improve MAE and SSIM,
318 retain PDF/raster/graphics identity and all 335 retain comment content and
main-document text. Office break matches reach 329 of 335. All seventeen changed
controls match exact row text, spaces and prepared fonts, including 23
within-word breaks. Maximum row-start X gap is 0.838pt and baseline gap 0.207pt.
Fifteen separator-only and two visible mixed first rows support the selected
phase rule (maximum first-pitch error 0.131pt); 34 later pitches retain tail
metrics (maximum error 0.082pt). Single-descent alternatives remain recorded.

All 34 cached PDFs retain bytes and 551 failed gates. Six break mismatches
remain: three parts, multiple paragraphs, colon, semicolon, overflowing-space
and internal-double-space guards. Earlier positioning/composed pixel residuals
and 33 partial cached cases remain. Qualification, identities, selected and
rejected models are under `artifacts/plan-revision-20261005/rv06-l51/`.


## RV06-L52: two fitting comment paragraphs and their marks

Two plain comment paragraphs that each fit one printed row retain their own
prepared body and paragraph-mark faces. The reader now retains marks for one
or two plain paragraphs in all-markup view. Both renderer and reader
regressions fail before their respective changes. First-face descent plus
next-face ascent and line gap measures the shared layout/emission step;
paragraph marks keep nominal balloon size. Explicit spacing, wrapped rows,
decorations and larger stories retain the complete fallback.

Fifteen mixed checks cover 96 scalar/font/mark/metric/scale variants and guards.
All 900 DOCX checks, 96 balloon checks and twenty-one Linux regressions pass.
Release is clean; fresh 0.1.5 package smoke contains the exact tested DLL at
runtime `1c9a3a4f`:
`04845D429E4F2F683B68713A60BEA275457DE962AB4D3AD16CF26D4234DDEC33`.

Across 335 prior and eighteen fresh inputs, sixteen improve MAE and SSIM,
335 retain PDF bytes and 337 retain raster/graphics identity. All 353 retain
main-document text and comment content; 345 match Office breaks. All sixteen
changed controls match exact row text, spaces and prepared fonts, with maximum
row-start X gap 0.838pt and baseline gap 0.181pt. Fifteen metric controls support
the selected step (maximum error 0.131pt); rejected mark-driven and uniform
models remain recorded. Two new fallback PDFs differ only by internal font
aliases: resolved glyph operations, embedded font bytes, declarations,
graphics and rasters retain identity.

All 34 cached PDFs retain bytes and 551 failed gates. Eight break mismatches
remain: three parts, colon, semicolon, overflowing-space, internal-double-space
and new wrapped/spacing/three-paragraph guards. Prior positioning/composed
pixel residuals and 33 partial cached cases remain. Qualification, identities
and the fallback alias audit are under `artifacts/plan-revision-20261005/rv06-l52/`.

## RV21-V33: full integration through L50 and SVG transparency

Frozen integration at `5486a32c`, through L49/L50 and accepted smooth SVG
transparency E4, passes **2179/0/1**, including 897 DOCX checks, with a clean
Release build. The sole skip is the private PPTX diagnostic's missing optional
input/output configuration. Fresh 0.1.5 package smoke contains the exact
full-suite DLL:
`16E704643A9B892CEF072C1003E0E3954F892E42FAACFDCF2C70A4616C7F973A`.
This run excludes L51/L52. Its report, package link and identities are under
`artifacts/plan-revision-20261005/milestone-5486a32c/`.

## RV06-L53: fitting closing run after a wrapped two-face body

A plain two-word prefix, a wrapped second face and a same-face closing run
retain their prepared fonts and terminal mark. When the closing run fits its
last tail word but overflows the current row, that word moves with the closing
run onto a new final mixed row. First, pure-tail and final mixed steps share
layout and emission. The production reader regression fails before changes;
three-run mark retention is exercised through the actual reader.

Sixteen mixed checks cover 192 scalar/ascent/gap/descent/length/closing-length/
scale variants, numeric guards and rejected formatting/spacing shapes. All
901 DOCX checks, 97 balloon checks and twenty-two Linux regressions pass.
Release is clean; fresh 0.1.5 package smoke contains the exact tested DLL at
runtime `cbc07d9c`:
`93C64100A778163678D73689F69D75B8E0C8C3965A151807441DFFCD48AFCD59`.

Across 353 prior and seventeen fresh inputs, fifteen improve MAE and SSIM,
353 retain PDF bytes and 355 retain raster/graphics identity. All 370 retain
main-document text and comment content; 360 match Office breaks. All fifteen
changed controls match exact row text, spaces and prepared fonts, with maximum
row-start X gap 0.838pt and baseline gap 0.303pt. Fourteen metric controls support
the selected first/final steps (maximum errors 0.124/0.141pt); 114 pure-tail
pitches stay within 0.082pt. Uniform tail metrics reach 0.342/0.900pt first/final
errors. The rejected proposal without final-row reflow also retains a measured
long-comment break mismatch. Both counterfactuals are archived. Two new fallback
PDFs differ only by audited font aliases; resolved operations, font bytes,
declarations, graphics and rasters retain identity.

All 34 cached PDFs retain bytes and 551 failed gates. Ten break mismatches
remain: colon, semicolon, overflowing-space, internal-double-space, prior
wrapped/spacing/three-paragraph and new overwide/decorated/separated closing-run
guards. Prior positioning/composed pixel residuals and 33 partial cached cases
remain. Qualification, identities and fallback audits are under
`artifacts/plan-revision-20261005/rv06-l53/`.


## RV06-L54: wrapped second comment paragraph

Two plain comment paragraphs retain their own prepared body and paragraph-mark
faces when the first fits one row and the second wraps. The first transition
uses first-body descent plus second-body ascent and line gap; later rows use
the second body's height. Authored spaces stay in the second body face and
scalar breaks insert no separators. Layout and emission share these steps;
marks stay at nominal balloon size. Both fitting paragraphs retain L52's path.

Seventeen mixed checks cover 192 scalar/zero-width/descent/gap/length/mark-face/
scale variants and wrapped-first, explicit-spacing, mixed/decorated, repeated-
space, tab, three-paragraph and oversized-scalar guards. All 902 DOCX checks,
98 balloon checks and twenty-three Linux regressions pass. Release is clean;
fresh 0.1.5 package smoke contains the exact tested DLL at runtime `7bb4fc67`:
`B2D332E5C8CF887BC55C7569BF8206C945A91760FABCD7E6C79AF62157E8D566`.

Across 370 prior and sixteen fresh inputs, fifteen improve MAE and SSIM and
371 retain PDF/raster/graphics identity. All 386 retain main-document text and
comment content; 375 match Office breaks. All fifteen changed controls match
exact row text, spaces and prepared fonts, including twenty within-word breaks,
with maximum row-start X gap 0.838pt and baseline gap 0.207pt. Fifteen metric
controls support the first transition (maximum error 0.131pt); 38 later pitches
stay within 0.082pt. Mark-driven, maximal-mark and uniform transition models
reach 0.495, 1.250 and 0.621pt errors respectively. The two fresh spacing and
mixed-paragraph guards retain exact PDF bytes.

All 34 cached PDFs retain bytes and 551 failed gates. Eleven break mismatches
remain: colon, semicolon, overflowing-space, internal-double-space, explicit
paragraph spacing, three paragraphs, overwide/decorated/separated closing runs,
and the new spacing/mixed-second-paragraph guards. Prior positioning/composed
pixel residuals and 33 partial cached cases remain. Qualification, identities
and selected/rejected metric audits are under
`artifacts/plan-revision-20261005/rv06-l54/`.


## RV06-L55: three fitting comment paragraphs

Three plain comment paragraphs that each fit one printed row retain their own
prepared body and nominal-size paragraph-mark faces. The reader retains marks
for one to three comment paragraphs in all-markup view. Renderer and reader
regressions fail before their changes. Each paragraph transition uses its
previous body's descent plus the next body's ascent and line gap; layout and
emission share both steps. Existing two-paragraph fitting/wrapped paths remain
qualified. Wrapped bodies, explicit spacing, mixed faces, decorated marks and
four-paragraph stories retain the complete fallback.

Eighteen mixed checks cover 192 scalar/middle-face/last-face/descent/gap/mark/
scale variants and formatting/numeric guards. All 903 DOCX checks, 99 balloon
checks and twenty-four Linux regressions pass. Release is clean; fresh 0.1.5
package smoke contains the exact tested DLL at runtime `921000ef`:
`5D268754269295928A5FC48F1AFB209D1E9415F995F65BEF5244D2991A1A9B7E`.

Across 386 prior and nineteen fresh inputs, sixteen improve MAE and SSIM,
386 retain PDF bytes and 389 retain raster/graphics identity. All 405 retain
main-document text and comment content; 391 match Office breaks. All sixteen
changed controls match exact row text, spaces and prepared fonts, with maximum
row-start X gap 0.838pt and baseline gap 0.187pt. Sixteen Office controls support
separate adjacent metric steps, with maximum errors 0.131/0.137pt. Rejected
mark-driven, maximal-mark and uniform-next-face proposals reach larger errors
and remain archived. Three fallback PDFs differ only by audited font aliases;
resolved operations, decoded font bytes, declarations, graphics and rasters
retain identity. The four-paragraph guard retains exact PDF bytes.

All 34 cached PDFs retain bytes and 551 failed gates. Fourteen break mismatches
remain: colon, semicolon, overflowing-space, internal-double-space, paragraph
spacing, mixed/wrapped paragraphs, four paragraphs and overwide/decorated/
separated closing runs. Prior positioning/composed pixel residuals and 33
partial cached cases remain. Qualification, identities and fallback audits are
under `artifacts/plan-revision-20261005/rv06-l55/`.


## RV21-V34: full integration through L51/L52

Frozen integration at `d03ef3fa`, through L51/L52 and accepted smooth SVG
transparency E4, passes **2182/0/1**, including 900 DOCX checks, with a clean
Release build. The sole skip remains the private PPTX diagnostic's missing
optional input/output configuration. Fresh 0.1.5 package smoke contains the
exact full-suite DLL:
`840DBA4D1C47158BC0366D1FAF2187F9DA1E7D52ECF5BBAE7998A2CC9B94A556`.
This run excludes L53/L54. Its report, package link and identities are under
`artifacts/plan-revision-20261005/milestone-d03ef3fa/`.


## RV06-L56: explicit comment paragraph spacing

Nonnegative integer before/after twips preserve prepared body and paragraph-mark faces in the qualified two-paragraph fitting/wrapped and three-paragraph fitting paths. The Office controls normalize these spacing values. Line spacing, automatic/contextual spacing, negative values and malformed tokens keep existing fallback or intake rejection. The production reproducer fails before the change; nineteen mixed checks cover 108 value/side/body-shape/scale variants and 66 unsupported-token/line/auto/contextual guards. Synthetic maximum-integer tokens remain finite; measured Office parity is bounded to the sampled values and shapes.

Frozen runtime `5804e48b` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `84328e22e85f44b99d6b02e75b0d251e` contains the exact tested DLL, SHA-256 `0F736A54A5EF4A54F9F3C39CBF9F24FAF1A48449AF130ECDD468D10A8C121385`. The Linux source archive emits two SourceLink warnings; its build has zero errors.

Across 405 prior and twenty-five valid fresh controls, twenty-five improve MAE and SSIM and 405 retain PDF/raster/graphics bytes. All 430 retain main-document text and comment content; 416 match Office breaks. All twenty-five changed controls match exact row text, spaces and prepared fonts, with maximum row-start X gap 0.838pt and baseline gap 0.187pt. The three fresh line/auto/negative guards retain exact PDFs. The ignored-spacing metric model has maximum error 0.082pt; rejected maximum/summed-spacing models reach 18.165pt. Nine later row pitches have maximum error 0.029pt.

Office rejects one additional malformed DOCX as corrupt before export, without an orphaned process. Both baseline and candidate CLI reject the identical input with the same twips error and no PDF. This separate intake audit is preserved without inventing an Office reference or counting it among the 430 visual controls.

All 34 cached PDFs retain bytes and 551 failed gates. Fourteen break mismatches remain: terminal colon/semicolon, oversized/internal-double spaces, mixed/wrapped paragraphs, four paragraphs, overwide/decorated/separated closing runs and unsupported paragraph-spacing modes. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, selected/rejected models, input identities and the malformed-input audit are under `artifacts/plan-revision-20261005/rv06-l56/`.



## RV21-V35: full integration through L54 and E4

Frozen source `7bb4fc6774cf74e3ac6af5a010286f44a8a2625e` passes a clean Release build and **2,184 tests, zero failures, one skip**, including 902 DOCX checks. The skipped `PptxPrivateLayoutDiagnosticWhenRequested` requires optional private input/output settings. This run covers accepted behavior through L54 and smooth Office SVG transparency E4; it excludes L55 and later prototypes.

The local 0.1.5 package smoke run `9f884216faeb4112971a409040b870ed` verifies the exact tested library bytes, SHA-256 `ECF3578B52A31C8A15B7CCC541DA3CA66E7634C4CA5D5B43B134C8347FA5D103`. The ignored evidence is under `artifacts/plan-revision-20261005/milestone-7bb4fc67/`; qualification completed 2026-10-06 at 13:32:48 UTC. The scoped Office residuals remain recorded with their individual slices. Version stays 0.1.5 and release preparation remains deferred.

## RV06-L57: terminal colon and semicolon in prepared prefixes

One terminal comma, colon or semicolon stays in the prepared prefix face in the existing fitting two-word/separated-tail and overwide single-word/joined-tail paths. The change retains authored punctuation and spaces and reuses the previously qualified row metrics. Internal/doubled punctuation, wrapped multiword punctuated prefixes and separated overwide prefixes keep complete fallback. Both expanded production regressions fail before the change; nineteen mixed checks cover 504 punctuation/scalar/advance/metric/length/scale variants and formatting/numeric guards.

Frozen runtime `b1a5455f` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `914bf7d93e814feb89bed3d6e6df663a` contains the exact tested DLL, SHA-256 `ADC95B8686892B0920F43A2F2C17DE4E5D8CE6914FCE8F47C9AFD9F57A8F83DF`. The Linux source archive emits two SourceLink warnings; its build has zero errors.

Across 430 prior and twenty-eight fresh controls, twenty-six improve MAE and SSIM and 432 retain PDF/raster/graphics bytes. All 458 retain main-document text and comment content; 442 match Office breaks. All twenty-six changed controls match exact row text, spaces and prepared fonts, including 55 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.207pt. All four fresh fallback guards retain exact PDFs. The original colon case improves MAE 0.329→0.198 and SSIM 0.869→0.945; the original semicolon case improves MAE 1.266→0.386 and SSIM 0.674→0.941.

All 34 cached PDFs retain bytes and 551 failed gates. Sixteen break mismatches remain: overflowing/internal-double spaces, mixed/wrapped paragraphs, four paragraphs, overwide/decorated/separated closing runs, unsupported paragraph-spacing modes and the four new punctuated-prefix guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, exact space/font audits, input/reference identities and fallback comparisons are under `artifacts/plan-revision-20261005/rv06-l57/`.


## RV06-L58: overflowing leading tail space spans

The complete authored leading ASCII space span stays in its prepared tail face beside a fitting plain two-word prefix even when the spaces exceed the available first-row width. The following word starts below the prefix. Layout and emission reuse the qualified separator-only first transition and later tail metrics. Internal repeated spaces, punctuated prefixes and tabs retain complete fallback. The production reproducer fails before the change; nineteen mixed checks cover 576 fitting/overflow/scalar/advance/word-shape/descent/length/scale variants, including 192 new overflow variants, with formatting/numeric guards.

Frozen runtime `8c9351e5` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `0b085d43fb95401d84986a99161e662e` contains the exact tested DLL, SHA-256 `502C8685D9F4162AFB2F6FE44F67E17A8A76F7013B0343113072ED264A3D18CA`. The Linux source archive emits two SourceLink warnings; its build has zero errors.

Across 458 prior and twenty-four fresh controls, twenty-one improve MAE and SSIM and 461 retain PDF/raster/graphics bytes. All 482 retain main-document text and comment content; 464 match Office breaks. All twenty-one changed controls match exact row text, spaces and prepared fonts, including 25 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.207pt. One shorter boundary and all three fresh fallback guards retain exact PDFs. All 22 measured first rows contain only tail separators after the prefix; the selected prefix-descent/tail-ascent model has maximum error 0.083pt, while rejected mixed-descent/uniform-tail models reach 0.621pt. Forty-six later row pitches have maximum error 0.082pt.

All 34 cached PDFs retain bytes and 551 failed gates. Eighteen break mismatches remain: internal-double spaces, mixed/wrapped paragraphs, four paragraphs, overwide/decorated/separated closing runs, unsupported paragraph-spacing modes, four punctuated-prefix guards and the three fresh overflow-shape guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, selected/rejected models, exact space/font audits and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l58/`.


## RV21-V36: full integration through L55/L56 and E4

Frozen source `5804e48b2365a2e7783c15b38e0cbc309ffe0832` passes a clean Release build and **2,186 tests, zero failures, one skip**, including 904 DOCX checks. The skipped `PptxPrivateLayoutDiagnosticWhenRequested` requires optional private input/output settings. This run covers accepted behavior through L56 and smooth Office SVG transparency E4; it excludes L57 and later prototypes.

The local 0.1.5 package smoke run `74424953c8d541c78ba120dbfd85f9bc` verifies the exact full-suite library bytes, SHA-256 `40B25C1F53F26CDCA96D0FEF1B4322B5B0ABA47F969E9BDD207658A8C5520CAF`. The ignored evidence is under `artifacts/plan-revision-20261005/milestone-5804e48b/`; qualification completed 2026-10-06 at 14:14:24 UTC. Scoped Office residuals remain recorded with their individual slices. Version stays 0.1.5 and release preparation remains deferred.


## RV06-L59: one doubled internal tail separator

One doubled internal ASCII separator between two tail words retains its prepared face and advance after a leading space span beside a fitting plain two-word prefix. Wrapping reserves the authored separator width and emission preserves both spaces. Triple internal separators, a single leading space, extra tail words, tabs and punctuated prefixes retain complete fallback. The production reproducer fails before the change; nineteen mixed checks cover 864 fitting/overflow/internal-separator/scalar/advance/word-shape/descent/length/scale variants, including 288 new doubled-separator variants, with formatting/numeric guards.

Frozen runtime `42c5c717` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `4d833f6e615d41d2a4a05857fb27e907` contains the exact tested DLL, SHA-256 `6F8FB86EE0CDCA98726A70868CC00EEE25F70FFE2EB51BF33E10EB2F0A5CD3A5`. The Linux source archive emits two SourceLink warnings; its build has zero errors.

Across 482 prior and thirty-one fresh controls, twenty-eight improve MAE and SSIM and 485 retain PDF/raster/graphics bytes. All 513 retain main-document text and comment content; 492 match Office breaks. All twenty-eight changed controls match exact row text, spaces and prepared fonts, including 23 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.207pt. All five fresh fallback guards retain exact PDFs. Twenty-eight metric controls cover 22 separator-only and six visible mixed first rows; the selected descent/ascent model has maximum error 0.090pt. Thirty-nine later row pitches have maximum error 0.082pt. Short-word boundaries, longer second words and scale controls retain exact Office breaks.

All 34 cached PDFs retain bytes and 551 failed gates. Twenty-one break mismatches remain: mixed/wrapped paragraphs, four paragraphs, overwide/decorated/separated closing runs, unsupported paragraph-spacing modes and punctuated-prefix/tab/triple-separator/single-leading/extra-tail guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, metric models, exact space/font audits and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l59/`. Separate full integration V37 at this runtime is complete; see its qualification below.


## RV06-L60: fitting third paragraph after a wrapped middle

A fitting third paragraph after a wrapped middle body retains its own prepared body and nominal paragraph mark. Height and emission share the first transition, middle row pitch and final transition. Wrapped first or third bodies, mixed paragraph faces, decorated marks, four paragraphs and unsupported spacing retain complete fallback. The production reproducer fails before the change; nineteen mixed checks cover 384 two-/three-paragraph scalar/advance/descent/gap/length/mark/scale variants, including 192 new fitting-third variants, with formatting and numeric guards.

Frozen runtime `874caff6` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `f9f900567e6c4231aeff1185abedb711` contains the exact tested DLL, SHA-256 `F68023F010E5C668DB9ACFD680F3FC1C839110E813C70A04314350DEFB349257`. The Linux source archive emits two SourceLink warnings; its build has zero errors.

Across 513 prior and twenty-six fresh controls, twenty-one improve MAE and SSIM and 518 retain PDF/raster/graphics bytes. All 539 retain main-document text and comment content; 514 match Office breaks. All twenty-one changed controls match exact row text, spaces and prepared fonts, including 22 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.214pt. The five fresh fallback guards and one fitting-middle boundary retain exact PDFs. Twenty-two metric controls qualify first/final transitions with maximum errors 0.083/0.071pt; thirty-nine middle row pitches have maximum error 0.082pt. Reusing uniform middle metrics for the final transition misses by up to 1.076pt.

All 34 cached PDFs retain bytes and 551 failed gates. Twenty-five break mismatches remain across mixed/wrapped paragraph bodies, four paragraphs, overwide/decorated/separated closing runs, unsupported spacing and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, metric models, exact space/font audits and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l60/`. Separate full integration V37 through L59/E4 is complete and excludes L60.

## RV21-V37: full integration through L59 and E4

Frozen revision `42c5c717` passes a clean Release build and the full Windows suite: **2186 passed, zero failed, one environmental skip**, including all 904 DOCX checks. The optional private PPTX layout diagnostic lacks its configured input/output. Scope includes accepted runtime through L59, smooth Office-preview SVG transparency E4, locked-field coverage and test helper P2; L60 and later prototypes are excluded.

Fresh local 0.1.5 package smoke `6154a2e13898461aa0e7b33bb8dfa4fa` contains the exact full-suite DLL, SHA-256 `4481ED7EDFCD8BFFC22EFB112A5085AB3462790A625434BD1B1A0FC9BAEEF473`. Full-suite/package evidence is under `artifacts/plan-revision-20261005/milestone-42c5c717/`; separate L59 and L60 Office-reference qualification remains authoritative for their scoped visual behavior. Version remains 0.1.5 and release preparation remains deferred.

## RV06-L61: wrapped final paragraph after a wrapped middle

A wrapped final paragraph after a wrapped middle retains its own prepared body and nominal paragraph mark. Height and emission share both adjacent paragraph transitions and each body's row pitch. Wrapped first bodies, mixed paragraph faces, decorated marks, four paragraphs and unsupported spacing retain complete fallback. The production reproducer fails before the change; nineteen mixed checks cover 576 two-/three-paragraph scalar/advance/descent/gap/length/mark/scale variants, including 192 new final-wrap variants, with formatting and numeric guards.

Frozen runtime `e4de3dcd` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `94dd44a7476843689dc50b43ca88151c` contains the exact tested DLL, SHA-256 `A792087190FA3EE3AED19D991342C40E64B25EC4EDE222787951E85E005E5EFD`. The Linux source archive emits two SourceLink warnings; its build has zero errors.

Across 539 prior and twenty-six fresh controls, twenty improve MAE and SSIM and 545 retain PDF/raster/graphics bytes. All 565 retain main-document text and comment content; 537 match Office breaks. All twenty changed controls match exact row text, spaces and prepared fonts, including 41 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.270pt. Five fresh fallback/previously-qualified guards and two fitting-final boundaries retain exact PDFs. Twenty-two metric controls qualify first/final incoming transitions with maximum errors 0.083/0.071pt; forty-four middle row pitches and twenty-five final row pitches have maximum errors 0.082/0.086pt. Reusing middle metrics for the final body's later rows misses by up to 1.316pt.

All 34 cached PDFs retain bytes and 551 failed gates. Twenty-eight break mismatches remain across mixed/wrapped paragraph bodies, four paragraphs, overwide/decorated/separated closing runs, unsupported spacing and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, metric models, exact space/font audits and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l61/`. Separate full integration V38 through L61/E4 is complete and excludes later prototypes.

## RV06-L62: four fitting plain comment paragraphs

Four plain paragraphs that each fit one row retain their own prepared body and nominal paragraph mark faces. Height and emission share all three adjacent paragraph transitions. Five paragraphs, wrapping, mixed/decorated bodies, decorated marks and unsupported spacing retain complete fallback. The renderer and reader production regressions fail before the change; nineteen mixed checks cover 384 three-/four-paragraph scalar/face/descent/gap/mark/scale variants, including 192 new four-paragraph variants, and the expanded reader regression preserves four marks in all-markup mode while final view and five paragraphs retain their prior policy.

Frozen runtime `6ca6ef6a` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `941e68094d724401a82588d195002211` contains the exact tested DLL, SHA-256 `BFE02371855BF37672BCA5E20F4951A59A54CAFDDC7C40C276A18828D5F08EA0`. The Linux source archive emits two SourceLink warnings; its build has zero errors.

Across 565 prior and twenty-six fresh controls, twenty-two improve MAE and SSIM and 563 retain PDF bytes. All 569 raster/graphics identities pass; six additional fallback PDFs differ only by audited font aliases, with identical glyph operations, font declarations after resolving aliases, decoded embedded font bytes and pixels. Four are fresh guards and two are prior wrapped-four-paragraph guards. All 591 retain main-document text and comment content; 559 match Office breaks. All twenty-two visually changed controls match exact row text, spaces and prepared fonts, with maximum row-start X gap 0.838pt and baseline gap 0.207pt. Twenty-two metric controls cover sixty-six paragraph transitions at maximum error 0.086pt; reusing uniform middle metrics misses by up to 1.336pt.

All 34 cached PDFs retain bytes and 551 failed gates. Thirty-two break mismatches remain across mixed/wrapped paragraph bodies, five paragraphs, overwide/decorated/separated closing runs, unsupported spacing and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, metric models, exact space/font and fallback-alias audits, and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l62/`. Separate full integration V38 through L61/E4 is complete and excludes L62 and later prototypes.

## RV21-V38: full integration through L61 and E4

Frozen revision `e4de3dcd` passes a clean Release build and the full Windows suite: **2186 passed, zero failed, one environmental skip**, including all 904 DOCX checks. The optional private PPTX layout diagnostic lacks its configured input/output. Scope includes accepted runtime through L61, smooth Office-preview SVG transparency E4, locked-field coverage and test helper P2; L62 and later prototypes are excluded.

Fresh local 0.1.5 package smoke `5280366d4f4d452bb27546ef908c4a49` contains the exact full-suite DLL, SHA-256 `0A49318C9AD6A2D5A9F04C9983785C99ACC7F3D23F23B1A35995F272D625DD7A`. Full-suite/package evidence is under `artifacts/plan-revision-20261005/milestone-e4de3dcd/`; separate L61 and L62 Office-reference qualification remains authoritative for their scoped visual behavior. Version remains 0.1.5 and release preparation remains deferred.

## RV06-L63: wrapped first bodies in three plain comment paragraphs

A first paragraph that exceeds the title row wraps in its prepared body font before a wrapped middle paragraph and a fitting or wrapped third. Each paragraph retains its own nominal mark. Height and emission share each body's continuation pitch and both adjacent transitions. Two-paragraph first wrapping, four-paragraph wrapping, fitting middle bodies, mixed/decorated bodies and unsupported spacing retain complete fallback. The expanded production regression fails before changes; nineteen mixed checks cover 960 two-/three-paragraph scalar/advance/descent/gap/length/mark/scale variants, including 384 new wrapped-first variants and formatting/numeric guards.

Frozen runtime `08594bc8` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `f2f8b3ed308e4f688e2ad0d7d5b5a882` contains the exact tested DLL, SHA-256 `2CE9150FDB11B5E1298DE71F4615C2212AC2141AB0BDC08DCB592D7AF880B4A1`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 591 prior and twenty-nine fresh controls, twenty-five improve MAE and SSIM and 595 retain PDF/raster/graphics bytes, with no image regressions. All 620 retain main-document text and comment content; 585 match Office breaks. All twenty-five changed cases match exact row text, spaces and prepared fonts, including 98 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.204pt. Twenty-five Office metric controls cover 66 first-body, fifty middle-body and four third-body pitches, at maximum errors 0.076/0.066/0.039pt; incoming transitions differ by at most 0.113/0.073pt. Reusing middle metrics for first continuations misses by up to 1.336pt. All six fresh guards/boundaries retain PDF/raster/graphics bytes.

Of 34 cached Office cases, 33 retain PDF bytes and the public long-comment case improves from seventeen failed gates to eleven, reducing the aggregate from 551 to 545 with no gate increases. It remains partial; text advance, geometry and raster residuals remain. Thirty-five control break mismatches remain across mixed/wrapped paragraph bodies, five paragraphs, overwide/decorated/separated closing runs, unsupported spacing and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, metric models, exact space/font audits and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l63/`. Separate full integration V39 through L63/E4 passes 2186/0/1, a clean Release build and an exact package; later prototypes are excluded.

## RV06-L64: wrapped first bodies in two plain comment paragraphs

Two plain paragraphs retain a wrapped first body before a wrapped second, each with its own prepared font and nominal paragraph mark. Height and emission share both continuation pitches and the adjacent paragraph transition. Fitting second bodies after a wrapped first, four-paragraph wrapping, mixed/decorated bodies and unsupported spacing retain complete fallback. The expanded production regression fails before changes; nineteen mixed checks cover 1152 two-/three-paragraph scalar/advance/descent/gap/length/mark/scale variants, including 192 new two-paragraph wrapped-first variants and formatting/numeric guards.

Frozen runtime `38766973` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `7a6cadf0e63a45bc8a3407299cb73ea8` contains the exact tested DLL, SHA-256 `F5FDB0E754F1AAD36C2A4AB844EBC92BFDFCE8EB2BE5FFE44EDFB589A0596A50`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 620 prior and twenty-two fresh controls, 17 improve MAE and SSIM and 625 retain PDF/raster/graphics bytes, with no image regressions. All 642 retain main-document text and comment content; 603 match Office breaks. All 17 changed cases match exact row text, spaces and prepared fonts, including 63 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.19pt. Seventeen Office metric controls cover 42 first-body and 34 second-body pitches at maximum errors 0.076/0.066pt; the transition differs by at most 0.113pt. Reusing second-body metrics for first continuations misses by up to 1.336pt. All six fresh guards/boundaries retain PDF/raster/graphics bytes.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate remains 545 failed gates with no gate increases. 39 control break mismatches remain across mixed/wrapped paragraph bodies, fitting second bodies, five paragraphs, overwide/decorated/separated closing runs, unsupported spacing and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, metric models, exact space/font audits and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l64/`. The held-out validation helper's stale reference path was corrected; completed comparisons were retained only after verifying their frozen DLL hash. Separate full integration V39 through L63/E4 passes 2186/0/1, a clean Release build and an exact package; L64 and later prototypes are excluded.

## RV06-L65: fitting second bodies after wrapped first comment paragraphs

A fitting second paragraph follows a wrapped first in two or three plain paragraphs, with a fitting or wrapped third when present. Each body and nominal mark retains its prepared font. Height and emission share all continuation pitches and adjacent transitions. Four-paragraph wrapping, fitting first and second with wrapped third, mixed/decorated bodies and unsupported spacing retain complete fallback. The expanded production regression fails before changes; nineteen mixed checks cover 1728 two-/three-paragraph scalar/advance/descent/gap/length/mark/scale variants, including 576 new fitting-second variants and formatting/numeric/zero-width-body guards. Two stale fitting-first fallback guards now use unsupported four-paragraph wrapping; their initial failures are retained.

Frozen runtime `4fcd0929` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `78a9c1479fc44fb78faee8f73e3e9370` contains the exact tested DLL, SHA-256 `B7660576BACE2E8A7D8EC53252C48D5D71BE84E595867BB3BF58F3E083747E75`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 642 prior and twenty-eight fresh controls, 24 improve MAE and SSIM and 646 retain PDF/raster/graphics bytes, with no image regressions. All 670 retain main-document text and comment content; 628 match Office breaks. All 24 changed cases match exact row text, spaces and prepared fonts, including 62 within-word break rows, with maximum row-start X gap 0.479pt and baseline gap 0.177pt. Twenty-four Office metric controls cover sixty first-body and three third-body pitches at maximum errors 0.076/0.029pt; transitions differ by at most 0.113/0.071pt. Reusing second-body metrics for first continuations misses by up to 1.336pt. All six fresh guards/boundaries retain PDF/raster/graphics bytes.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 42 control break mismatches remain across mixed/wrapped paragraph bodies, five paragraphs, overwide/decorated/separated closing runs, unsupported spacing and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, metric models, exact space/font audits and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l65/`. Separate full integration V40 through L65/E4 was still in progress when this slice was qualified and excludes later prototypes.

## RV21-V39: full integration through L63 and smooth SVG transparency

Frozen revision `08594bc8` passes a clean Release build and the full Windows suite: **2186 passed, zero failed, one environmental skip**, including all 904 DOCX checks. The optional private PPTX layout diagnostic lacks its configured input/output. Scope includes accepted runtime through L63, smooth Office-preview SVG transparency E4, locked-field coverage and test helper P2; L64 and later prototypes are excluded.

Fresh local 0.1.5 package smoke `733e2de8d79d4fa18a84a5927467b594` contains the exact full-suite DLL, SHA-256 `3C85895AD2F7AA79FBBE99286E9FAB7F2C996D47674B048CE4092C7CFA25E079`. Full-suite/package evidence is under `artifacts/plan-revision-20261005/milestone-08594bc8/`; separate L63, L64 and L65 Office-reference qualifications remain authoritative for their scoped visual behavior. Version remains 0.1.5 and release preparation remains deferred.

## RV06-L66: wrapped third bodies after two fitting comment paragraphs

A wrapped third body follows two fitting plain paragraphs. Each body and nominal mark retains its prepared font; height and emission share the third body's continuation pitch and both adjacent transitions. Four-paragraph wrapping, mixed/decorated bodies and unsupported spacing retain complete fallback. The expanded production regression fails before changes; nineteen mixed checks cover 1920 two-/three-paragraph scalar/advance/descent/gap/length/mark/scale variants, including 192 new third-only wrap combinations and formatting/numeric/zero-width-body guards.

Frozen runtime `97b0c6ad` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `0f6938ab6bf34a2ca02d3bec2f08ab6f` contains the exact tested DLL, SHA-256 `C98F6676BAE8983769174B1BBC0EFB5670A910517BB61DEDBB842BCE6FF8BA18`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 670 prior and twenty-six fresh controls, 19 improve MAE and SSIM and 677 retain PDF/raster/graphics bytes, with no image regressions. All 696 retain main-document text and comment content; 649 match Office breaks. All 19 changed cases match exact row text, spaces and prepared fonts, including 16 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.277pt. Nineteen Office metric controls cover twenty-five third-body pitches at maximum error 0.083pt; transitions differ by at most 0.083/0.089pt. Reusing second-body metrics for third continuations misses by up to 1.346pt. All six guards/boundaries retain PDF/raster/graphics bytes; the separately classified short third body fits one row and also retains PDF/raster/graphics bytes, without changing its input.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 47 control break mismatches remain across mixed/wrapped paragraph bodies, five paragraphs, overwide/decorated/separated closing runs, unsupported spacing and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, metric models, exact space/font and boundary audits and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l66/`. Separate full integration V40 through L65/E4 was still in progress when this slice was qualified and excludes L66 and later prototypes.

## RV06-L67: wrapped fourth bodies after three fitting comment paragraphs

A wrapped fourth body follows three fitting plain paragraphs. Each body and nominal mark retains its prepared font; height and emission share the fourth body's continuation pitch and all three adjacent transitions. Earlier wrapping in four paragraphs, mixed/decorated bodies and unsupported spacing retain complete fallback. The expanded production regression fails before changes; nineteen mixed checks cover 576 leading-fitting/fourth-wrap scalar/advance/descent/gap/length/mark/scale variants, including 192 new wrapping combinations and formatting/numeric/zero-width-body guards.

Frozen runtime `e69e78dd` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `2ade54ca65c3479290ee1a24a203c03c` contains the exact tested DLL, SHA-256 `82665BEF77AD5251224E45C14FB98557D12800D1AE7A0D72204FDDBE83E2D666`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 696 prior and twenty-six fresh controls, 21 improve MAE and SSIM and 701 retain PDF/raster/graphics bytes, with no image regressions. All 722 retain main-document text and comment content; 671 match Office breaks. All 21 changed cases match exact row text, spaces and prepared fonts, including 21 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.264pt. Twenty-one Office metric controls cover 31 fourth-body pitches at maximum error 0.105pt; adjacent transitions differ by at most 0.083/0.059/0.085pt. Reusing third-body metrics for fourth continuations misses by up to 0.59pt. All six unsupported/boundary guards retain PDF/raster/graphics bytes. The original four-paragraph case improves from three candidate rows to eight matching Office rows, with exact spaces/fonts and MAE 0.768→0.181.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 51 control break mismatches remain across earlier wrapping in four paragraphs, mixed paragraph bodies, five paragraphs, overwide/decorated/separated closing runs, unsupported spacing and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, metric models, exact space/font and guard audits and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l67/`. Separate full integration V40 through L65/E4 is complete and excludes L66 and later prototypes.

## RV21-V40: full integration through L65 and smooth SVG transparency

Frozen runtime `4fcd0929` passes a clean Release build and **2186 passed, 0 failed, 1 skipped**, including 904 DOCX checks. The optional private PPTX layout diagnostic lacks configured input/output. Exact local 0.1.5 package smoke `eb28d51a710e4e779b4ba80b57dd45a2` contains the full-suite DLL, SHA-256 `B9ED1318D0CEDC7E60E9C785BFA21522F75350CB2244F228517EE2930657CB7C`. This run covers L20 through L65/E4 and excludes L66 and later prototypes. Evidence is under `artifacts/plan-revision-20261005/milestone-4fcd0929/`. Release preparation remains deferred.

## RV06-L68: wrapped first bodies before three fitting comment paragraphs

A wrapped first body precedes three fitting plain paragraphs. Each body and nominal mark retains its prepared font; height and emission share the first body's continuation pitch and all three adjacent transitions. Other wrapping combinations in four paragraphs, mixed/decorated bodies and unsupported spacing retain complete fallback. The expanded production regression fails before changes; nineteen mixed checks cover 2112 two-/three-/four-paragraph scalar/advance/descent/gap/length/mark/scale variants, including 192 new four-paragraph combinations and formatting/numeric/zero-width-body guards.

Frozen runtime `a7e00d87` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `54b5def2ce7a4c17aab6a5cfe4897648` contains the exact tested DLL, SHA-256 `E61B7E2746894BFB651C79D16E8DB92DB25D8A074D5ABDFBCB3DD9A3A32E06FC`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 722 prior and twenty-six fresh controls, 22 improve MAE and SSIM and 726 retain PDF/raster/graphics bytes, with no image regressions. All 748 retain main-document text and comment content; 694 match Office breaks. All 22 changed cases match exact row text, spaces and prepared fonts, including 59 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.21pt. Twenty-two Office metric controls cover 59 first-body pitches at maximum error 0.075pt; adjacent transitions differ by at most 0.113/0.071/0.082pt. Reusing second-body metrics for first continuations misses by up to 1.336pt. All six unsupported/boundary guards retain PDF/raster/graphics bytes. The earlier fitting-second four-paragraph case also improves and matches exact Office rows, spaces and fonts. The original four-paragraph case improves from two candidate rows to seven matching Office rows, with exact spaces/fonts and MAE 0.563→0.151.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 54 control break mismatches remain across earlier wrapping in four paragraphs, mixed paragraph bodies, five paragraphs, overwide/decorated/separated closing runs, unsupported spacing and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, metric models, exact space/font and guard audits and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l68/`. Separate full integration V40 through L65/E4 is complete and excludes L66 and later prototypes.


## RV06-L69: wrapped second bodies in four comment paragraphs

A wrapped second body follows a fitting first and precedes two fitting plain paragraphs. Each body and nominal mark retains its prepared font; height and emission share the second body's continuation pitch and all three adjacent transitions. Other wrapping combinations in four paragraphs, mixed/decorated bodies and unsupported spacing retain complete fallback. The expanded production regression fails before changes; nineteen mixed checks cover 2304 two-/three-/four-paragraph scalar/advance/descent/gap/length/mark/scale variants, including 192 new four-paragraph combinations and formatting/numeric/zero-width-body guards.

Frozen runtime `c224f31d` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `f6ab484a3b9f42478fb7d0cd046cae36` contains the exact tested DLL, SHA-256 `BDAA89FD5011692C0FB3DB6DB06D2DBEC7F0B2E141440255A6B4A39DFB9C6CA0`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 748 prior and twenty-six fresh controls, 21 improve MAE and SSIM and 753 retain PDF/raster/graphics bytes, with no image regressions. All 774 retain main-document text and comment content; 716 match Office breaks. All 21 changed cases match exact row text, spaces and prepared fonts, including 18 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.2pt. Twenty-one Office metric controls cover 35 second-body pitches at maximum error 0.071pt; adjacent transitions differ by at most 0.083/0.071/0.092pt. Reusing first-body metrics for second continuations misses by up to 1.355pt. All six unsupported/boundary guards retain PDF/raster/graphics bytes. The original four-paragraph case improves from three candidate rows to six matching Office rows, with exact spaces/fonts and MAE 0.427→0.204.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 58 control break mismatches remain across earlier wrapping in four paragraphs, mixed paragraph bodies, five paragraphs, overwide/decorated/separated closing runs, unsupported spacing and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, metric models, exact space/font and guard audits and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l69/`. Separate full integration V41 through L68/E4 was still in progress when this slice was qualified and excludes L69 and later prototypes.


## RV06-L70: wrapped third bodies in four comment paragraphs

A wrapped third body follows two fitting bodies and precedes a fitting fourth paragraph. Each body and nominal mark retains its prepared font; height and emission share the third body's continuation pitch and all three adjacent transitions. Other wrapping combinations in four paragraphs, mixed/decorated bodies and unsupported spacing retain complete fallback. The expanded production regression fails before changes; nineteen mixed checks cover 2496 two-/three-/four-paragraph scalar/advance/descent/gap/length/mark/scale variants, including 192 new four-paragraph combinations and formatting/numeric/zero-width-body guards.

Frozen runtime `cc859424` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `df488ed5b46a47f2a8e9ed3c358aab70` contains the exact tested DLL, SHA-256 `2A62582590A015737CA84D2CB03EC7DAC9640784419E2BD668DEE10797152465`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 774 prior and twenty-six fresh controls, 20 improve MAE and SSIM and 780 retain PDF/raster/graphics bytes, with no image regressions. All 800 retain main-document text and comment content; 738 match Office breaks. All 20 changed cases match exact row text, spaces and prepared fonts, including 18 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.194pt. Twenty Office metric controls cover 22 third-body pitches at maximum error 0.04pt; adjacent transitions differ by at most 0.083/0.059/0.105pt. Reusing second-body metrics for third continuations misses by up to 1.346pt. All six unsupported/boundary guards retain PDF/raster/graphics bytes. The separately classified short Calibri third body fits one row and retains PDF/raster/graphics bytes, without changing its input. The original four-paragraph case improves from three candidate rows to five matching Office rows, with exact spaces/fonts and MAE 0.456→0.149.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 62 control break mismatches remain across earlier wrapping in four paragraphs, mixed paragraph bodies, five paragraphs, overwide/decorated/separated closing runs, unsupported spacing and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, metric models, exact space/font and guard audits and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l70/`. Separate full integration V41 through L68/E4 is complete and excludes L69 and later prototypes.

## RV21-V41: full integration through L68 and smooth SVG transparency

Frozen runtime `a7e00d87` passes a clean Release build and **2186 passed, 0 failed, 1 skipped**, including 904 DOCX checks. The optional private PPTX layout diagnostic lacks configured input/output. Exact local 0.1.5 package smoke `fc8ca2fecb6e42f98e34dadf69618701` contains the full-suite DLL, SHA-256 `EEFFFF3A1F197652C434F1C648E6594D919F1B5E9AE9FAB1728A70A46F5D30A9`. This run covers L20 through L68/E4 and excludes L69 and later prototypes. Evidence is under `artifacts/plan-revision-20261005/milestone-a7e00d87/`. Release preparation remains deferred.

## RV06-L71: wrapped first and second bodies in four comment paragraphs

Wrapped first and second bodies precede fitting third and fourth bodies. Each body and nominal mark retains its prepared font; height and emission share each continuation pitch and all three adjacent transitions. Other wrapping combinations in four paragraphs, mixed/decorated bodies and unsupported spacing retain complete fallback. The expanded production regression fails before changes; nineteen mixed checks cover 2688 two-/three-/four-paragraph scalar/advance/descent/gap/length/mark/scale variants, including 192 new first-and-second combinations and inherited formatting/numeric/zero-width-body guards.

Frozen runtime `d015108c` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `54c7d705881948368e0dc20dc71d153c` contains the exact tested DLL, SHA-256 `9CE7A7BB57C4441BE6160504E2BA3DF4D013503CCD3F79D66054188F411273F5`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 800 prior and twenty-six fresh controls, 24 improve MAE and SSIM and 802 retain PDF/raster/graphics bytes, with no image regressions. All 826 retain main-document text and comment content; 763 match Office breaks. All 24 changed cases match exact row text, spaces and prepared fonts, including 86 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.29pt. 24 Office metric controls cover 65 first-body pitches and 41 second-body pitches at maximum errors 0.075pt and 0.066pt; adjacent transitions differ by at most 0.113/0.141/0.111pt. Reusing second-body metrics for first continuations misses by up to 1.336pt. All six unsupported/boundary guards retain PDF/raster/graphics bytes. The original four-paragraph case improves from four candidate rows to nine matching Office rows, with exact spaces/fonts and MAE 0.669→0.198. Older changed cases are included in the emission and metric inventories.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 63 control break mismatches remain across other wrapping combinations in four paragraphs, mixed paragraph bodies, five paragraphs, overwide/decorated/separated closing runs, unsupported spacing and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, metric models, exact space/font and guard audits and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l71/`. Separate full integration V42 through L70/E4 was still in progress when this slice was qualified and excludes L71 and later prototypes.


## RV06-L72: wrapped first and third bodies in four comment paragraphs

Wrapped first and third bodies alternate with fitting second and fourth bodies. Each body and nominal mark retains its prepared font; height and emission share each continuation pitch and all three adjacent transitions. Other wrapping combinations in four paragraphs, mixed/decorated bodies and unsupported spacing retain complete fallback. The expanded production regression fails before changes; nineteen mixed checks cover 2880 two-/three-/four-paragraph scalar/advance/descent/gap/length/mark/scale variants, including 192 new first-and-third combinations and inherited formatting/numeric/zero-width-body guards.

Frozen runtime `8d0c3dd3` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `73fdc50d0a294b5c97b3a7a6ec5f4e8a` contains the exact tested DLL, SHA-256 `7B9496BC2EBCB1554B9B903314970A9364BAA2F1244C461BCD262134BA4E1297`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 826 prior and twenty-six fresh controls, 21 improve MAE and SSIM and 831 retain PDF/raster/graphics bytes, with no image regressions. All 852 retain main-document text and comment content; 786 match Office breaks. All 21 changed cases match exact row text, spaces and prepared fonts, including 77 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.294pt. 21 Office metric controls cover 58 first-body pitches and 23 third-body pitches at maximum errors 0.075pt and 0.05pt; adjacent transitions differ by at most 0.103/0.061/0.121pt. Reusing second-body metrics for first/third continuations misses by up to 1.336/1.346pt. All six unsupported/boundary guards retain PDF/raster/graphics bytes. The short Calibri third body fits one row after the wrapped first body and retains PDF/raster/graphics bytes, without changing its input. The original four-paragraph case improves from four candidate rows to eight matching Office rows, with exact spaces/fonts and MAE 0.658→0.164. Older changed cases are included in the emission and metric inventories.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 66 control break mismatches remain across other wrapping combinations in four paragraphs, mixed paragraph bodies, five paragraphs, overwide/decorated/separated closing runs, unsupported spacing and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, metric models, exact space/font and guard audits and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l72/`. Separate full integration V42 through L70/E4 was still in progress when this slice was qualified and excludes L71 and later prototypes.


## RV06-L73: wrapped second and third bodies in four comment paragraphs

Wrapped second and third bodies follow a fitting first body and precede a fitting fourth body. Each body and nominal mark retains its prepared font; height and emission share each continuation pitch and all three adjacent transitions. Other wrapping combinations in four paragraphs, mixed/decorated bodies and unsupported spacing retain complete fallback. The expanded production regression fails before changes; nineteen mixed checks cover 3072 two-/three-/four-paragraph scalar/advance/descent/gap/length/mark/scale variants, including 192 new second-and-third combinations and inherited formatting/numeric/zero-width-body guards.

Frozen runtime `8c5224e8` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `bd85138fd19f457886710c89f9f41c13` contains the exact tested DLL, SHA-256 `0FDA73D3E6CA55F2CA3F0699631A7280F56730FF474AE3510A58DDF264A720BD`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 852 prior and twenty-six fresh controls, 22 improve MAE and SSIM and 856 retain PDF/raster/graphics bytes, with no image regressions. All 878 retain main-document text and comment content; 810 match Office breaks. All 22 changed cases match exact row text, spaces and prepared fonts, including 46 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.217pt. 22 Office metric controls cover 42 second-body pitches and 24 third-body pitches at maximum errors 0.071pt and 0.085pt; adjacent transitions differ by at most 0.083/0.071/0.092pt. Reusing first metrics for second continuations and second metrics for third continuations misses by up to 1.355/1.336pt. All six unsupported/boundary guards retain PDF/raster/graphics bytes. The short Calibri third body fits one row after the wrapped second body and retains PDF/raster/graphics bytes, without changing its input. The original four-paragraph case improves from five candidate rows to seven matching Office rows, with exact spaces/fonts and MAE 0.570→0.220. Older changed cases are included in the emission and metric inventories.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 68 control break mismatches remain across other wrapping combinations in four paragraphs, mixed paragraph bodies, five paragraphs, overwide/decorated/separated closing runs, unsupported spacing and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, metric models, exact space/font and guard audits and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l73/`. Separate full integration V42 through L70/E4 is complete and excludes L71 and later prototypes.

## RV21-V42: full integration through L70 and smooth SVG transparency

Frozen runtime `cc859424` passes a clean Release build and **2186 passed, 0 failed, 1 skipped**, including 904 DOCX checks. The optional private PPTX layout diagnostic lacks configured input/output. Exact local 0.1.5 package smoke `8cc2116273d44de9a99f60173246f955` contains the full-suite DLL, SHA-256 `530C0DDD6790821BC65FFB441667B57AF86FC64E592D50BA4D0F3CA63BCEB7E2`. This run covers L20 through L70/E4 and excludes L71 and later prototypes. Evidence is under `artifacts/plan-revision-20261005/milestone-cc859424/`. Release preparation remains deferred.

## RV06-L74: three wrapped leading bodies in four comment paragraphs

Three wrapped leading bodies precede a fitting fourth body. Each body and nominal mark retains its prepared font; height and emission share each continuation pitch and all three adjacent transitions. Other wrapping combinations in four paragraphs, mixed/decorated bodies and unsupported spacing retain complete fallback. The expanded production regression fails before changes; nineteen mixed checks cover 3264 two-/three-/four-paragraph scalar/advance/descent/gap/length/mark/scale variants, including 192 new triple-wrapping combinations and inherited formatting/numeric/zero-width-body guards.

Frozen runtime `0cfe2932` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `415f15407f3547fdab17cb99ff8ea802` contains the exact tested DLL, SHA-256 `01A403A8B76057BBF4E559F3547BDB622D8228E15249D52FC18B4426F25EAAF6`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 878 prior and twenty-six fresh controls, 22 improve MAE and SSIM and 882 retain PDF/raster/graphics bytes, with no image regressions. All 904 retain main-document text and comment content; 835 match Office breaks. All 22 changed cases match exact row text, spaces and prepared fonts, including 107 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.287pt. 22 Office metric controls cover 61 first-body pitches, 42 second-body pitches and 24 third-body pitches at maximum errors 0.075pt/0.066pt and 0.04pt; adjacent transitions differ by at most 0.103/0.071/0.121pt. Reusing second metrics for first continuations, first metrics for second continuations and second metrics for third continuations misses by up to 1.336/1.375/1.336pt. All six unsupported/boundary guards retain PDF/raster/graphics bytes. The short Calibri third body fits one row after two wrapped leading bodies and retains PDF/raster/graphics bytes, without changing its input. The original four-paragraph case improves from six candidate rows to ten matching Office rows, with exact spaces/fonts and MAE 0.795→0.212. Older changed cases are included in the emission and metric inventories.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 69 control break mismatches remain across other wrapping combinations in four paragraphs, mixed paragraph bodies, five paragraphs, overwide/decorated/separated closing runs, unsupported spacing and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, metric models, exact space/font and guard audits and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l74/`. Separate full integration V43 through L72/E4 was still in progress when this slice was qualified and excludes L73 and later prototypes.


## RV06-L75: wrapped first and fourth comment bodies

Wrapped first and fourth bodies enclose two fitting middle paragraphs. Each body and nominal mark retains its prepared font. Fourth-body rows, pitch and continuation height are measured once and shared by balloon height and drawing; the fitting third body remains above that continuation phase. Other wrapping combinations involving the fourth paragraph, mixed/decorated bodies and unsupported spacing retain complete fallback. The expanded production regression fails before changes; nineteen mixed checks cover 3456 two-/three-/four-paragraph scalar/advance/descent/gap/length/mark/scale variants, including 192 new first/fourth combinations. Numeric guards include a nonpositive fourth continuation pitch with positive incoming transition, zero-width/overwide fourth bodies, and unsupported combinations. Third and fourth marks are checked against their actual final body rows.

Frozen runtime `a3adac23` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `26c97c81b30a472483ea953196659564` contains the exact tested DLL, SHA-256 `F648F584DFD0E14F62E2278735949F625ADF2B42E685DE5405B483FAC8F071F4`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 904 prior and twenty-six fresh controls, 21 improve MAE and SSIM and 909 retain PDF/raster/graphics bytes, with no image regressions. All 930 retain main-document text and comment content; 857 match Office breaks. All 21 changed cases match exact row text, spaces and prepared fonts, including 74 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.264pt. 21 Office metric controls cover 56 first-body pitches and 28 fourth-body pitches at maximum errors 0.075pt and 0.105pt; adjacent transitions differ by at most 0.113/0.071/0.082pt. Reusing second metrics for first continuations and third metrics for fourth continuations misses by up to 1.336/0.59pt. All six unsupported/boundary guards retain PDF/raster/graphics bytes. The original case improves from four candidate rows to eight matching Office rows, with exact spaces/fonts and MAE 0.661→0.169.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 73 control break mismatches remain across other fourth-body wrapping combinations, mixed paragraph bodies, five paragraphs, overwide/decorated/separated closing runs, unsupported spacing and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, metric models, exact space/font and guard audits and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l75/`. Separate full integration V43 through L72/E4 is complete and excludes L73 and later prototypes.

## RV21-V43: full integration through L72 and smooth SVG transparency

Frozen runtime `8d0c3dd3` passes a clean Release build and **2186 passed, 0 failed, 1 skipped**, including 904 DOCX checks. The optional private PPTX layout diagnostic lacks configured input/output. Exact local 0.1.5 package smoke `dd6d673051094eea838bd9d55643980f` contains the full-suite DLL, SHA-256 `94D66C5F936C9A071D174132CC5C5E3F98D6920C813007C5CFADE1849C832A5B`. This run covers L20 through L72/E4 and excludes L73 and later prototypes. Evidence is under `artifacts/plan-revision-20261005/milestone-8d0c3dd3/`. Release preparation remains deferred.

## RV06-L76: wrapped second and fourth comment bodies

Wrapped second and fourth bodies alternate with fitting first and third paragraphs. Each body and nominal mark retains its prepared font. The shared paragraph-row model retains each continuation pitch and all three adjacent transitions; the fitting third remains above the fourth continuation phase. Other wrapping combinations involving the fourth paragraph, mixed/decorated bodies and unsupported spacing retain complete fallback. The expanded production regression fails before changes; nineteen mixed checks cover 3648 two-/three-/four-paragraph scalar/advance/descent/gap/length/mark/scale variants, including 192 new second/fourth combinations. Numeric guards include a nonpositive fourth continuation pitch with positive incoming transition, zero-width/overwide fourth bodies, and unsupported combinations. Third and fourth marks are checked against their actual final body rows.

Frozen runtime `bb97751b` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `569b3d3e51744d1d895f0cd8f0f924ed` contains the exact tested DLL, SHA-256 `4DCFAB494A552716F1F2A069B4C3E38FDD0C13744CA386F1C33811ABEA8F8254`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 930 prior and twenty-six fresh controls, 21 improve MAE and SSIM and 935 retain PDF/raster/graphics bytes, with no image regressions. All 956 retain main-document text and comment content; 879 match Office breaks. All 21 changed cases match exact row text, spaces and prepared fonts, including 43 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.257pt. 21 Office metric controls cover 39 second-body pitches and 28 fourth-body pitches at maximum errors 0.071pt and 0.085pt; adjacent transitions differ by at most 0.083/0.071/0.092pt. Reusing first metrics for second continuations and third metrics for fourth continuations misses by up to 1.355/0.59pt. All six unsupported/boundary guards retain PDF/raster/graphics bytes. The original case improves from five candidate rows to seven matching Office rows, with exact spaces/fonts and MAE 0.579→0.237.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 77 control break mismatches remain across other fourth-body wrapping combinations, mixed paragraph bodies, five paragraphs, overwide/decorated/separated closing runs, unsupported spacing and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, metric models, exact space/font and guard audits and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l76/`. Separate full integration V43 through L72/E4 is complete and excludes L73 and later prototypes.


## RV06-L77: wrapped first, second and fourth comment bodies

Wrapped first, second and fourth bodies surround a fitting third paragraph. Each body and nominal mark retains its prepared font. The shared paragraph-row model retains each continuation pitch and all three adjacent transitions; the fitting third remains above the fourth continuation phase. Other wrapping combinations involving the fourth paragraph, mixed/decorated bodies and unsupported spacing retain complete fallback. The expanded production regression fails before changes; nineteen mixed checks cover 3840 two-/three-/four-paragraph scalar/advance/descent/gap/length/mark/scale variants, including 192 new first/second/fourth combinations. Numeric guards include a nonpositive fourth continuation pitch with positive incoming transition, zero-width/overwide fourth bodies, and unsupported combinations. Third and fourth marks are checked against their actual final body rows.

Frozen runtime `9d0352c7` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `d93764e87ff448e8879d9ecb93d4030c` contains the exact tested DLL, SHA-256 `6F9F56867C35DE3031133850F363598B91F6ABE0F6E07CD2BFCCB241928753B1`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 956 prior and twenty-six fresh controls, 23 improve MAE and SSIM and 959 retain PDF/raster/graphics bytes, with no image regressions. All 982 retain main-document text and comment content; 904 match Office breaks. All 23 changed cases match exact row text, spaces and prepared fonts, including 109 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.29pt. 23 Office metric controls cover 62 first-body pitches, 43 second-body pitches and 30 fourth-body pitches at maximum errors 0.075pt/0.066pt and 0.075pt; adjacent transitions differ by at most 0.113/0.141/0.111pt. Reusing second metrics for first continuations, first metrics for second continuations and third metrics for fourth continuations misses by up to 1.336/1.375/0.6pt. All six unsupported/boundary guards retain PDF/raster/graphics bytes. The original case improves from six candidate rows to ten matching Office rows, with exact spaces/fonts and MAE 0.793→0.214. Older changed L75/L76 guards are included in the complete metric and emission inventories.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 78 control break mismatches remain across other fourth-body wrapping combinations, mixed paragraph bodies, five paragraphs, overwide/decorated/separated closing runs, unsupported spacing and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, metric models, exact space/font and guard audits and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l77/`. Separate full integration V44 through L74/E4 is complete and excludes L75 and later prototypes.

## RV21-V44: full integration through L74 and smooth SVG transparency

Frozen runtime `0cfe2932` passes a clean Release build and **2186 passed, 0 failed, 1 skipped**, including 904 DOCX checks. The optional private PPTX layout diagnostic lacks configured input/output. Exact local 0.1.5 package smoke `ee16d0744a6a4c6b835613e1610fb535` contains the full-suite DLL, SHA-256 `D461531B3C14DC8D72458E12A1720912A4DAFAB6ED31DC91BC6BEFDEAAF5A879`. This run covers L20 through L74/E4 and excludes L75 and later prototypes. Evidence is under `artifacts/plan-revision-20261005/milestone-0cfe2932/`. Release preparation remains deferred.

## RV06-L78: wrapped first, third and fourth comment bodies

Wrapped first, third and fourth bodies surround a fitting second paragraph. Each body and nominal mark retains its prepared font. The shared paragraph-row model retains each continuation pitch and all three adjacent transitions; the fitting third remains above the fourth continuation phase. Other wrapping combinations involving the fourth paragraph, mixed/decorated bodies and unsupported spacing retain complete fallback. The expanded production regression fails before changes; nineteen mixed checks cover 4032 two-/three-/four-paragraph scalar/advance/descent/gap/length/mark/scale variants, including 192 new first/third/fourth combinations. Numeric guards include a nonpositive fourth continuation pitch with positive incoming transition, zero-width/overwide fourth bodies, and unsupported combinations. Third and fourth marks are checked against their actual final body rows.

Frozen runtime `b03f86af` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `1763f08fb3ff4ba3b08b88feb3fc83e2` contains the exact tested DLL, SHA-256 `D5B403B83FFAE85AC7E05353BFFFF4ED2FCACCBA9DFA61CDF0941592B272BE14`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 982 prior and twenty-six fresh controls, 21 improve MAE and SSIM and 987 retain PDF/raster/graphics bytes, with no image regressions. All 1008 retain main-document text and comment content; 927 match Office breaks. All 21 changed cases match exact row text, spaces and prepared fonts, including 96 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.291pt. 21 Office metric controls cover 58 first-body pitches, 24 third-body pitches and 28 fourth-body pitches at maximum errors 0.075pt/0.085pt and 0.075pt; adjacent transitions differ by at most 0.103/0.073/0.105pt. Reusing second metrics for first continuations, second metrics for third continuations and third metrics for fourth continuations misses by up to 1.336/1.346/0.6pt. All six unsupported/boundary guards retain PDF/raster/graphics bytes. The short Calibri third body fits one row within six body rows total and retains PDF/raster/graphics bytes; its input is unchanged. Three pre-reference font variants give the third body a different face from the first; reusing first metrics for third continuations misses by up to 0.289pt. The original case improves from six candidate rows to nine matching Office rows, with exact spaces/fonts and MAE 0.816→0.187. The older changed L75 guard is included in the complete metric and emission inventories.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 81 control break mismatches remain across other fourth-body wrapping combinations, mixed paragraph bodies, five paragraphs, overwide/decorated/separated closing runs, unsupported spacing and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, metric models, exact space/font and guard audits and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l78/`. Separate full integration V44 through L74/E4 is complete and excludes L75 and later prototypes.


## RV06-L79: wrapped second, third and fourth comment bodies

Wrapped second, third and fourth bodies follow a fitting first paragraph. Each body and nominal mark retains its prepared font. The shared paragraph-row model retains each continuation pitch and all three adjacent transitions; the fitting third remains above the fourth continuation phase. Other wrapping combinations involving the fourth paragraph, mixed/decorated bodies and unsupported spacing retain complete fallback. The expanded production regression fails before changes; nineteen mixed checks cover 4224 two-/three-/four-paragraph scalar/advance/descent/gap/length/mark/scale variants, including 192 new second/third/fourth combinations. Numeric guards include a nonpositive fourth continuation pitch with positive incoming transition, zero-width/overwide fourth bodies, and unsupported combinations. Third and fourth marks are checked against their actual final body rows.

Frozen runtime `4784a655` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `a31612d7ad544daf809d145b0e26a987` contains the exact tested DLL, SHA-256 `189AE2E05DBB64AA62B997C21EBD919DD46C428CFD9438CF7048A4C1A20773A1`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 1008 prior and twenty-six fresh controls, 21 improve MAE and SSIM and 1013 retain PDF/raster/graphics bytes, with no image regressions. All 1034 retain main-document text and comment content; 950 match Office breaks. All 21 changed cases match exact row text, spaces and prepared fonts, including 63 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.284pt. 21 Office metric controls cover 40 second-body pitches, 24 third-body pitches and 28 fourth-body pitches at maximum errors 0.071pt/0.085pt and 0.075pt; adjacent transitions differ by at most 0.083/0.071/0.092pt. Reusing first metrics for second continuations, second metrics for third continuations and third metrics for fourth continuations misses by up to 1.355/1.336/0.59pt. All six unsupported/boundary guards retain PDF/raster/graphics bytes. The short Calibri third body fits one row within six body rows total and retains PDF/raster/graphics bytes; its input is unchanged. Three font variants give the third body a different face from the first; reusing first metrics for third continuations misses by up to 0.289pt. The original case improves from seven candidate rows to eight matching Office rows, with exact spaces/fonts and MAE 0.718→0.254. The older changed L76 guard is included in the complete metric and emission inventories.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 84 control break mismatches remain across other fourth-body wrapping combinations, mixed paragraph bodies, five paragraphs, overwide/decorated/separated closing runs, unsupported spacing and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, metric models, exact space/font and guard audits and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l79/`. Separate full integration V45 through L77/E4 was still in progress when this slice was qualified and excludes L78 and later prototypes.


## RV06-L80: wrapped third and fourth comment bodies after two fitting paragraphs

Fitting first and second paragraphs precede wrapped third and fourth bodies. Each body and nominal mark retains its prepared font. The shared row model supplies both continuation pitches and all three adjacent transitions to geometry and emission. All-four wrapping, mixed/decorated bodies and unsupported spacing retain complete fallback. The expanded production regression fails before changes; nineteen mixed checks cover 4416 two-/three-/four-paragraph scalar/advance/descent/gap/length/mark/scale variants, including 192 new third/fourth combinations. A stale negative assertion for the admitted combination is removed while positive production coverage and the all-four fallback guard remain. Numeric guards exercise invalid third/fourth pitches and incoming transitions, zero-width and overwide scalars and bodies. Actual final body rows determine third/fourth mark positions.

Frozen runtime `68805a3c` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `28d039522f4c47ca9cec0da2bca3704b` contains the exact tested DLL, SHA-256 `FDE54F58332966DA31761EC04F02F795C933911FCC6FDF2188F2D620118DAAFC`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 1034 prior and twenty-six fresh controls, 21 improve MAE and SSIM and 1039 retain PDF/raster/graphics bytes, with no image regressions. All 1060 retain main-document text and comment content; 974 match Office breaks. All 21 changed cases match exact row text, spaces and prepared fonts, including 38 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.277pt. 21 Office metric controls cover 24 third-body and 28 fourth-body pitches at maximum errors 0.082/0.105pt; adjacent transitions differ by at most 0.083/0.083/0.105pt. Reusing second metrics for third continuations and third metrics for fourth continuations misses by up to 1.346/0.6pt. Three distinct third-body font variants reject first-metric reuse by up to 0.289pt. All six fallback/boundary guards retain PDF/raster/graphics bytes. The short Calibri third body fits one row within five body rows total and retains PDF/raster/graphics bytes with unchanged input. The original case improves from 5 candidate rows to 6 matching Office rows, exact spaces/fonts and MAE 0.577→0.17. The older changed L79 fitting-second guard is included in the complete metric and emission inventories.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 86 control break mismatches remain across all-four wrapping, mixed paragraph bodies, five paragraphs, overwide/decorated/separated closing runs, unsupported spacing and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, metric models, exact space/font and guard audits and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l80/`. Separate full integration V45 through L77/E4 is complete and excludes L78 and later prototypes.

## RV21-V45: full integration through L77 and smooth SVG transparency

Frozen runtime `9d0352c7` passes a clean Release build and **2186 passed, 0 failed, 1 skipped**, including 904 DOCX checks. The optional private PPTX layout diagnostic lacks configured input/output. Exact local 0.1.5 package smoke `2f951cb96bd348299b055d1782fa2364` contains the full-suite DLL, SHA-256 `E5BF885A31475C3FA085B72E73084CA05A91CFFD1F9BED257CBB7EEB8D1371F2`. This run covers L20 through L77/E4 and excludes L78 and later prototypes. Evidence is under `artifacts/plan-revision-20261005/milestone-9d0352c7/`. Release preparation remains deferred.

## RV06-L81: four independently wrapped comment bodies

Four independently wrapped bodies use each prepared font and each continuation pitch. Each body and nominal mark retains its prepared font. The shared row model supplies all four continuation pitches and all three adjacent transitions to geometry and emission. Mixed/decorated bodies and unsupported spacing retain complete fallback. The expanded production regression fails before changes; nineteen mixed checks cover 4608 two-/three-/four-paragraph scalar/advance/descent/gap/length/mark/scale variants, including 192 new all-four combinations. Two stale negative assertions for the admitted combination are removed while positive production coverage, numeric guards and five-paragraph fallback remain. Numeric guards exercise invalid body pitches and incoming transitions, zero-width and overwide scalars and bodies. Actual final body rows determine all four mark positions.

Frozen runtime `1adf6698` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `446ef81fed64404d9a8acecbbde3210c` contains the exact tested DLL, SHA-256 `B431DE33030788614EEB74B28AE7875164111F19A75B5D66F970E8E5975B8E3D`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 1060 prior and twenty-six fresh controls, 24 improve MAE and SSIM and 1062 retain PDF/raster/graphics bytes, with no image regressions. All 1086 retain main-document text and comment content; 1002 match Office breaks. All 24 changed cases match exact row text, spaces and prepared fonts, including 139 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.284pt. 24 Office metric controls cover 67 first/46 second/27 third/31 fourth pitches at maximum errors 0.075/0.066/0.04/0.091pt; adjacent transitions differ by at most 0.103/0.073/0.082pt. Reusing second metrics for first continuations and first metrics for second continuations misses by up to 1.336/1.375pt. Reusing second metrics for third continuations and third metrics for fourth continuations misses by up to 1.336/0.6pt. Three distinct third-body font variants reject first-metric reuse by up to 0.289pt. All six fallback/boundary guards retain PDF/raster/graphics bytes. The short Calibri third body fits one row within seven body rows total and retains PDF/raster/graphics bytes with unchanged input. The original case improves from 8 candidate rows to 11 matching Office rows, exact spaces/fonts and MAE 0.96→0.249. The older changed L77, L78, L79 and L80 guards are included in the complete metric and emission inventories.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 84 control break mismatches remain across mixed paragraph bodies, five paragraphs, overwide/decorated/separated closing runs, unsupported spacing and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, metric models, exact space/font and guard audits and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l81/`. Separate full integration V46 through L79/E4 was still in progress when this slice was qualified and excludes L80 and later prototypes.


## RV06-L82: canonical after-lines spacing in plain comment paragraphs

Plain two-, three- and four-paragraph balloons normalize the exact `afterLines="100"` token while retaining each prepared body and nominal mark face, its own continuation pitch and adjacent transitions. Before-lines, other after-lines values, malformed tokens, automatic/contextual spacing, mixed/decorated bodies and five paragraphs retain complete fallback. The expanded production spacing regression fails before changes; nineteen mixed checks retain 4608 scalar/advance/descent/gap/length/mark/scale body variants and add 108 canonical-spacing combinations across fitting/wrapped bodies, explicit before/after twips and scale. Earlier negative after-lines assertions use the unqualified value 200; numeric and formatting guards remain active.

Frozen runtime `7b61d6f6` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `fd857415c3cc42f3bd08ce72d619b3f5` contains the exact tested DLL, SHA-256 `FDCCE7606BF01DDA6B5652E15A598F49AF126FCA3923B0F7A0CF4C6807ED2970`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 1086 prior and twenty-six fresh controls, 43 improve MAE and SSIM and 1069 retain PDF/raster/graphics bytes, with no image regressions. All 1112 retain main-document text and comment content; 1045 match Office breaks. All 43 changed cases match exact row text, spaces and prepared fonts, including 174 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.284pt. 43 Office metric controls cover 214 continuation pitches and 120 adjacent transitions with maximum errors 0.091/0.141pt. Each changed input is matched to its own hash and cached Office reference; all changed legacy spacing cases are included. All six negative guards retain PDF/raster/graphics bytes. Fresh positive fixtures apply the canonical token to every paragraph, including distinct third-body fonts, marks, right margins and explicit twip spacing. The short Calibri third fits one row within seven body rows total: it is a positive target because the spacing previously forced fallback. Its input is unchanged and its rows, spaces and fonts match Office. The original case improves from 8 candidate rows to 11 matching Office rows, exact spaces/fonts and MAE 0.96→0.249.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 67 control break mismatches remain across unsupported spacing, mixed/decorated paragraph bodies, five paragraphs, overwide/decorated/separated closing runs and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, exact space/font and metric audits, changed legacy inventory, guards and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l82/`. Separate full integration V46 through L79/E4 is complete and excludes L80 and later prototypes.

## RV21-V46: full integration through L79 and smooth SVG transparency

Frozen runtime `4784a655` passes a clean Release build and **2186 passed, 0 failed, 1 skipped**, including 904 DOCX checks. The optional private PPTX layout diagnostic lacks configured input/output. Exact local 0.1.5 package smoke `55f53760d8474a3d8709e71f7154a120` contains the full-suite DLL, SHA-256 `BAFDABB89B9114F34CCBC145F2F2C09C93B7062737C70CE09ECF43424D32EB0C`. This run covers L20 through L79/E4 and excludes L80 and later prototypes. Evidence is under `artifacts/plan-revision-20261005/milestone-4784a655/`. Release preparation remains deferred.

## RV06-L83: canonical before-lines spacing in plain comment paragraphs

Plain two-, three- and four-paragraph balloons normalize the exact `beforeLines="100"` token, alone or with `afterLines="100"` while retaining each prepared body and nominal mark face, its own continuation pitch and adjacent transitions. Other before-lines values and unqualified after-lines values, malformed tokens, automatic/contextual spacing, mixed/decorated bodies and five paragraphs retain complete fallback. The expanded production spacing regression fails before changes; nineteen mixed checks retain 4608 scalar/advance/descent/gap/length/mark/scale body variants and add 216 canonical before-lines combinations (324 canonical-spacing combinations beyond the original twip cases) across fitting/wrapped bodies, explicit before/after twips and scale. The previous negative before-lines assertion uses the unqualified value 200; numeric and formatting guards remain active.

Frozen runtime `96febbca` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `7389a79437dd40de97e65467a28d6bd3` contains the exact tested DLL, SHA-256 `287F2703AC725B24D78CDBE8868C83AF1EBC616173CB9323C21219B279954232`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 1112 prior and twenty-six fresh controls, 21 improve MAE and SSIM and 1117 retain PDF/raster/graphics bytes, with no image regressions. All 1138 retain main-document text and comment content; 1066 match Office breaks. All 21 changed cases match exact row text, spaces and prepared fonts, including 117 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.284pt. 21 Office metric controls cover 146 continuation pitches and 63 adjacent transitions with maximum errors 0.091/0.141pt. Each changed input is matched to its own hash and cached Office reference; all changed legacy spacing cases are included. All six negative guards retain PDF/raster/graphics bytes. Fresh positive fixtures apply beforeLines=100 to every paragraph, with ten before-only and ten combined before/afterLines=100 targets, including distinct third-body fonts, marks, right margins and explicit twip spacing. The short Calibri third fits one row within seven body rows total: it is a positive target because the spacing previously forced fallback. Its input is unchanged and its rows, spaces and fonts match Office. The original case improves from 8 candidate rows to 11 matching Office rows, exact spaces/fonts and MAE 0.96→0.249.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 72 control break mismatches remain across unsupported spacing, mixed/decorated paragraph bodies, five paragraphs, overwide/decorated/separated closing runs and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, exact space/font and metric audits, changed legacy inventory, guards and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l83/`. Separate full integration V47 through L81/E4 was still in progress when this slice was qualified and excludes L82 and later prototypes.


## RV06-L84: canonical after-auto spacing in plain comment paragraphs

Plain two-, three- and four-paragraph balloons normalize the exact `afterAutospacing="1"` token, alone or with canonical `beforeLines="100"`/`afterLines="100"` while retaining each prepared body and nominal mark face, its own continuation pitch and adjacent transitions. Before-auto, other after-auto values and unqualified line values, malformed tokens, contextual spacing, mixed/decorated bodies and five paragraphs retain complete fallback. The expanded production spacing regression fails before changes; nineteen mixed checks retain 4608 scalar/advance/descent/gap/length/mark/scale body variants and add 432 automatic-spacing combinations beyond 324 earlier canonical line-spacing combinations across fitting/wrapped bodies, explicit before/after twips and scale. Existing before-auto and after-auto=0 assertions, numeric and formatting guards remain active.

Frozen runtime `d98a1eb5` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `7555b1e62cba4de6bd41fd4a9707197d` contains the exact tested DLL, SHA-256 `D8D3320B553C402451187F6AC33ED803D30449B2B0A839F34187AE45C479604F`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 1138 prior and twenty-six fresh controls, 21 improve MAE and SSIM and 1143 retain PDF/raster/graphics bytes, with no image regressions. All 1164 retain main-document text and comment content; 1087 match Office breaks. All 21 changed cases match exact row text, spaces and prepared fonts, including 111 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.284pt. 21 Office metric controls cover 139 continuation pitches and 61 adjacent transitions with maximum errors 0.091/0.141pt. Each changed input is matched to its own hash and cached Office reference; the changed legacy automatic-spacing case is included. All six negative guards retain PDF/raster/graphics bytes. Fresh positive fixtures apply afterAutospacing=1 to every paragraph, with five targets for each of four canonical before/afterLines=100 combinations, including distinct third-body fonts, marks, right margins and explicit twip spacing. The short Calibri third fits one row within seven body rows total: it is a positive target because the spacing previously forced fallback. Its input is unchanged and its rows, spaces and fonts match Office. The original case improves from 1 candidate rows to 2 matching Office rows, exact spaces/fonts and MAE 0.188→0.112.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 77 control break mismatches remain across unsupported spacing, mixed/decorated paragraph bodies, five paragraphs, overwide/decorated/separated closing runs and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, exact space/font and metric audits, changed legacy inventory, guards and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l84/`. Separate full integration V47 through L81/E4 is complete and excludes L82 and later prototypes.

## RV21-V47: full integration through L81 and smooth SVG transparency

Frozen runtime `1adf6698` passes a clean Release build and **2186 passed, 0 failed, 1 skipped**, including 904 DOCX checks. The optional private PPTX layout diagnostic lacks configured input/output. Exact local 0.1.5 package smoke `a3b4f10c8fe54bcebed812609069553f` contains the full-suite DLL, SHA-256 `EC5A0CE4484DB31F93201CA157DF00E449B7BA677FE3EAEC2025DFB81B36ADF7`. This run covers L20 through L81/E4 and excludes L82 and later prototypes. Evidence is under `artifacts/plan-revision-20261005/milestone-1adf6698/`. Release preparation remains deferred.

## RV06-L85: canonical before-auto spacing in plain comment paragraphs

Plain two-, three- and four-paragraph balloons normalize the exact `beforeAutospacing="1"` token, alone or with canonical after-auto and `beforeLines="100"`/`afterLines="100"` while retaining each prepared body and nominal mark face, its own continuation pitch and adjacent transitions. Other before-auto values and unqualified after-auto/line values, malformed tokens, contextual spacing, mixed/decorated bodies and five paragraphs retain complete fallback. The expanded production spacing regression fails before changes; nineteen mixed checks retain 4608 scalar/advance/descent/gap/length/mark/scale body variants and add 864 before-auto combinations beyond 756 earlier canonical spacing combinations across fitting/wrapped bodies, explicit before/after twips and scale. Existing automatic-spacing zero assertions, numeric and formatting guards remain active.

Frozen runtime `08376fbd` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `0d2da07739fb46b6a7b7833c99685b28` contains the exact tested DLL, SHA-256 `7EDF278AA43956C224BF7E73D809C419903FD964DECCA0968263DFBBADBE3EBB`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 1164 prior and twenty-six fresh controls, 21 improve MAE and SSIM and 1169 retain PDF/raster/graphics bytes, with no image regressions. All 1190 retain main-document text and comment content; 1108 match Office breaks. All 21 changed cases match exact row text, spaces and prepared fonts, including 117 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.284pt. 21 Office metric controls cover 146 continuation pitches and 63 adjacent transitions with maximum errors 0.091/0.141pt. Each changed input is matched to its own hash and cached Office reference; the changed legacy automatic-spacing case is included. All six negative guards retain PDF/raster/graphics bytes. Fresh positive fixtures apply beforeAutospacing=1 to every paragraph, with twenty targets covering eight accepted after-auto/before/afterLines combinations, including standalone before-auto and distinct third-body fonts, marks, right margins and explicit twip spacing. The short Calibri third fits one row within seven body rows total: it is a positive target because the spacing previously forced fallback. Its input is unchanged and its rows, spaces and fonts match Office. The original case improves from 8 candidate rows to 11 matching Office rows, exact spaces/fonts and MAE 0.96→0.249.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 82 control break mismatches remain across unsupported spacing, mixed/decorated paragraph bodies, five paragraphs, overwide/decorated/separated closing runs and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, exact space/font and metric audits, changed legacy inventory, guards and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l85/`. Separate full integration V48 through L83/E4 is complete and excludes L84 and later prototypes.

## RV21-V48: full integration through L83 and smooth SVG transparency

Frozen runtime `96febbca` passes a clean Release build and **2186 passed, 0 failed, 1 skipped**, including 904 DOCX checks. The optional private PPTX layout diagnostic lacks configured input/output. Exact local 0.1.5 package smoke `39313a6ce41947aebe69959aafd12ed0` contains the full-suite DLL, SHA-256 `0D529E319EBCFC80A867C4F7933DEF5AFB23FE85405A66BAF93AECA5B5D62542`. This run covers L20 through L83/E4 and excludes L84 and later prototypes. Evidence is under `artifacts/plan-revision-20261005/milestone-96febbca/`. Release preparation remains deferred.

## RV06-L86: canonical after-auto-zero spacing in plain comment paragraphs

Plain two-, three- and four-paragraph balloons normalize the exact `afterAutospacing="0"` token, alone or with canonical before-auto and `beforeLines="100"`/`afterLines="100"` while retaining each prepared body and nominal mark face, its own continuation pitch and adjacent transitions. Before-auto zero, lexical after-auto values and unqualified line values, malformed tokens, contextual spacing, mixed/decorated bodies and five paragraphs retain complete fallback. The expanded production spacing regression fails before changes; nineteen mixed checks retain 4608 scalar/advance/descent/gap/length/mark/scale body variants and add 864 after-auto zero combinations beyond 1620 earlier canonical spacing combinations across fitting/wrapped bodies, explicit before/after twips and scale. The before-auto zero assertion remains active; the earlier after-auto zero assertion uses a malformed value. Lexical, numeric and formatting guards remain active.

Frozen runtime `04ceef0c` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `63bb673f3c19477cb67dac7d11b86c4f` contains the exact tested DLL, SHA-256 `2F51D6EBE3A99E74A3944EDD157F5A478796762C53A3E84AA717E849CB7FEE52`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 1190 prior and twenty-six fresh controls, 22 improve MAE and SSIM and 1194 retain PDF/raster/graphics bytes, with no image regressions. All 1216 retain main-document text and comment content; 1130 match Office breaks. All 22 changed cases match exact row text, spaces and prepared fonts, including 123 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.284pt. 22 Office metric controls cover 153 continuation pitches and 66 adjacent transitions with maximum errors 0.091/0.141pt. Each changed input is matched to its own hash and cached Office reference; every changed legacy automatic-spacing case is included. All six negative guards retain PDF/raster/graphics bytes. Fresh positive fixtures apply afterAutospacing=0 to every paragraph, with twenty targets covering eight accepted before-auto/before/afterLines combinations, including standalone after-auto zero and distinct third-body fonts, marks, right margins and explicit twip spacing. The short Calibri third fits one row within seven body rows total: it is a positive target because the spacing previously forced fallback. Its input is unchanged and its rows, spaces and fonts match Office. The original case improves from 8 candidate rows to 11 matching Office rows, exact spaces/fonts and MAE 0.96→0.249.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 86 control break mismatches remain across unsupported spacing, mixed/decorated paragraph bodies, five paragraphs, overwide/decorated/separated closing runs and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, exact space/font and metric audits, changed legacy inventory, guards and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l86/`. Separate full integration V48 through L83/E4 is complete and excludes L84 and later prototypes.


## RV06-L87: canonical before-auto-zero spacing in plain comment paragraphs

Plain two-, three- and four-paragraph balloons normalize the exact `beforeAutospacing="0"` token, alone or with canonical after-auto and `beforeLines="100"`/`afterLines="100"` while retaining each prepared body and nominal mark face, its own continuation pitch and adjacent transitions. Lexical before/after-auto values and unqualified line values, malformed tokens, contextual spacing, mixed/decorated bodies and five paragraphs retain complete fallback. The expanded production spacing regression fails before changes; nineteen mixed checks retain 4608 scalar/advance/descent/gap/length/mark/scale body variants and add 1296 before-auto zero combinations beyond 2484 earlier canonical spacing combinations across fitting/wrapped bodies, explicit before/after twips and scale. Both earlier exact-zero assertions now use malformed values. Lexical, numeric and formatting guards remain active.

Frozen runtime `3b6caf99` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `18012b663b4f4d05a790c9d2d42560ed` contains the exact tested DLL, SHA-256 `CE314A90FC6315F5BCCFB8F1A4BE99F5898B0F7CD7422A06C5253E0B90A36FAD`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 1216 prior and twenty-six fresh controls, 22 improve MAE and SSIM and 1220 retain PDF/raster/graphics bytes, with no image regressions. All 1242 retain main-document text and comment content; 1152 match Office breaks. All 22 changed cases match exact row text, spaces and prepared fonts, including 123 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.284pt. 22 Office metric controls cover 153 continuation pitches and 66 adjacent transitions with maximum errors 0.091/0.141pt. Each changed input is matched to its own hash and cached Office reference; every changed legacy automatic-spacing case is included. All six negative guards retain PDF/raster/graphics bytes. Fresh positive fixtures apply beforeAutospacing=0 to every paragraph, with twenty targets covering twelve accepted absent/0/1 after-auto and before/afterLines combinations, including standalone before-auto zero and distinct third-body fonts, marks, right margins and explicit twip spacing. The short Calibri third fits one row within seven body rows total: it is a positive target because the spacing previously forced fallback. Its input is unchanged and its rows, spaces and fonts match Office. The original case improves from 8 candidate rows to 11 matching Office rows, exact spaces/fonts and MAE 0.96→0.249.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 90 control break mismatches remain across unsupported spacing, mixed/decorated paragraph bodies, five paragraphs, overwide/decorated/separated closing runs and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, exact space/font and metric audits, changed legacy inventory, guards and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l87/`. Separate full integration V49 through L85/E4 is complete and excludes L86 and later prototypes.

## RV21-V49: full integration through L85 and smooth SVG transparency

Frozen runtime `08376fbd` passes a clean Release build and **2186 passed, 0 failed, 1 skipped**, including 904 DOCX checks. The optional private PPTX layout diagnostic lacks configured input/output. Exact local 0.1.5 package smoke `5c39f87424a54bbcae1350ea8b3636f0` contains the full-suite DLL, SHA-256 `59704DC7579DBF07F012D6E1A5D6916783E822AEF882839ECA1B2EF3774E9AD0`. This run covers L20 through L85/E4 and excludes L86 and later prototypes. Evidence is under `artifacts/plan-revision-20261005/milestone-08376fbd/`. Release preparation remains deferred.

## RV06-L88: canonical after-auto-true spacing in plain comment paragraphs

Plain two-, three- and four-paragraph balloons normalize the exact `afterAutospacing="true"` token, alone or with canonical before-auto and `beforeLines="100"`/`afterLines="100"` while retaining each prepared body and nominal mark face, its own continuation pitch and adjacent transitions. Lexical before-auto values and after-auto false and unqualified line values, malformed tokens, contextual spacing, mixed/decorated bodies and five paragraphs retain complete fallback. The expanded production spacing regression fails before changes; nineteen mixed checks retain 4608 scalar/advance/descent/gap/length/mark/scale body variants and add 1296 after-auto true combinations beyond 3780 earlier canonical spacing combinations across fitting/wrapped bodies, explicit before/after twips and scale. Malformed automatic-spacing assertions remain active. Lexical, numeric and formatting guards remain active.

Frozen runtime `15f93d40` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `2d169aa5314e4480ad789e35cbdf6893` contains the exact tested DLL, SHA-256 `5B3CEAF3876020259AB3BE3E1812FD4E5A26520DFD01D63A3F3016F969063689`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 1242 prior and twenty-six fresh controls, 22 improve MAE and SSIM and 1246 retain PDF/raster/graphics bytes, with no image regressions. All 1268 retain main-document text and comment content; 1174 match Office breaks. All 22 changed cases match exact row text, spaces and prepared fonts, including 123 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.284pt. 22 Office metric controls cover 153 continuation pitches and 66 adjacent transitions with maximum errors 0.091/0.141pt. Each changed input is matched to its own hash and cached Office reference; every changed legacy automatic-spacing case is included. All six negative guards retain PDF/raster/graphics bytes. Fresh positive fixtures apply afterAutospacing=true to every paragraph, with twenty targets covering twelve accepted absent/0/1 before-auto and before/afterLines combinations, including standalone after-auto true and distinct third-body fonts, marks, right margins and explicit twip spacing. The short Calibri third fits one row within seven body rows total: it is a positive target because the spacing previously forced fallback. Its input is unchanged and its rows, spaces and fonts match Office. The original case improves from 8 candidate rows to 11 matching Office rows, exact spaces/fonts and MAE 0.96→0.249.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 94 control break mismatches remain across unsupported spacing, mixed/decorated paragraph bodies, five paragraphs, overwide/decorated/separated closing runs and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, exact space/font and metric audits, changed legacy inventory, guards and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l88/`. Separate full integration V50 through L87/E4 was still in progress when this slice was qualified and excludes L88 and later prototypes.


## RV06-L89: canonical after-auto-false spacing in plain comment paragraphs

Plain two-, three- and four-paragraph balloons normalize the exact `afterAutospacing="false"` token, alone or with canonical before-auto and `beforeLines="100"`/`afterLines="100"` while retaining each prepared body and nominal mark face, its own continuation pitch and adjacent transitions. Lexical before-auto values and after-auto on and unqualified line values, malformed tokens, contextual spacing, mixed/decorated bodies and five paragraphs retain complete fallback. The expanded production spacing regression fails before changes; nineteen mixed checks retain 4608 scalar/advance/descent/gap/length/mark/scale body variants and add 1296 after-auto false combinations beyond 5076 earlier canonical spacing combinations across fitting/wrapped bodies, explicit before/after twips and scale. Malformed automatic-spacing assertions remain active. Lexical, numeric and formatting guards remain active.

Frozen runtime `f3bb65f8` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `f57a32bc8d324d04849f0c2cf9cdd280` contains the exact tested DLL, SHA-256 `10E56D1A8DB9DA78DEEA6C54D50679F046115F618B64B5C73235C1067E2272A4`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 1268 prior and twenty-six fresh controls, 21 improve MAE and SSIM and 1273 retain PDF/raster/graphics bytes, with no image regressions. All 1294 retain main-document text and comment content; 1195 match Office breaks. All 21 changed cases match exact row text, spaces and prepared fonts, including 117 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.284pt. 21 Office metric controls cover 146 continuation pitches and 63 adjacent transitions with maximum errors 0.091/0.141pt. Each changed input is matched to its own hash and cached Office reference; every changed legacy automatic-spacing case is included. All six negative guards retain PDF/raster/graphics bytes. Fresh positive fixtures apply afterAutospacing=false to every paragraph, with twenty targets covering twelve accepted absent/0/1 before-auto and before/afterLines combinations, including standalone after-auto false and distinct third-body fonts, marks, right margins and explicit twip spacing. The short Calibri third fits one row within seven body rows total: it is a positive target because the spacing previously forced fallback. Its input is unchanged and its rows, spaces and fonts match Office. The original case improves from 8 candidate rows to 11 matching Office rows, exact spaces/fonts and MAE 0.96→0.249.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 99 control break mismatches remain across unsupported spacing, mixed/decorated paragraph bodies, five paragraphs, overwide/decorated/separated closing runs and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, exact space/font and metric audits, changed legacy inventory, guards and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l89/`. Separate full integration V50 through L87/E4 is complete and excludes L89 and later prototypes.

## RV21-V50: full integration through L87 and smooth SVG transparency

Frozen runtime `3b6caf99` passes a clean Release build and **2186 passed, 0 failed, 1 skipped**, including 904 DOCX checks. The optional private PPTX layout diagnostic lacks configured input/output. Exact local 0.1.5 package smoke `8c9b5e5c6268468289be1a0288d3181e` contains the full-suite DLL, SHA-256 `64BBF6B2314A8E6CADDAEBE4E4DEA81F5387B22A83578D3AB205471186EA30CB`. This run covers L20 through L87/E4 and excludes L89 and later prototypes. Evidence is under `artifacts/plan-revision-20261005/milestone-3b6caf99/`. Release preparation remains deferred.

## RV06-L90: canonical before-auto-true spacing in plain comment paragraphs

Plain two-, three- and four-paragraph balloons normalize the exact `beforeAutospacing="true"` token, alone or with canonical after-auto and `beforeLines="100"`/`afterLines="100"` while retaining each prepared body and nominal mark face, its own continuation pitch and adjacent transitions. Before-auto false, after-auto on, unqualified line values, malformed tokens, contextual spacing, mixed/decorated bodies and five paragraphs retain complete fallback. The expanded production spacing regression fails before changes; nineteen mixed checks retain 4608 scalar/advance/descent/gap/length/mark/scale body variants and add 2160 before-auto true combinations beyond 6372 earlier canonical spacing combinations across fitting/wrapped bodies, explicit before/after twips and scale. Malformed automatic-spacing assertions remain active. Lexical, numeric and formatting guards remain active.

Frozen runtime `b84855e1` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `02d95c1fa1204abdaef4b8095bb4159d` contains the exact tested DLL, SHA-256 `7EC9EB08B8B0C0E919A1350C625E3F4B0C86E879E2D6EECE68FC52430E1F64E6`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 1294 prior and twenty-six fresh controls, 23 improve MAE and SSIM and 1297 retain PDF/raster/graphics bytes, with no image regressions. All 1320 retain main-document text and comment content; 1218 match Office breaks. All 23 changed cases match exact row text, spaces and prepared fonts, including 129 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.284pt. 23 Office metric controls cover 160 continuation pitches and 69 adjacent transitions with maximum errors 0.091/0.141pt. Each changed input is matched to its own hash and cached Office reference; every changed legacy automatic-spacing case is included. All six negative guards retain PDF/raster/graphics bytes. Fresh positive fixtures apply beforeAutospacing=true to every paragraph, with twenty targets covering all twenty accepted absent/0/1/true/false after-auto and before/afterLines combinations, including standalone before-auto true and distinct third-body fonts, marks, right margins and explicit twip spacing. The short Calibri third fits one row within seven body rows total: it is a positive target because the spacing previously forced fallback. Its input is unchanged and its rows, spaces and fonts match Office. The original case improves from 8 candidate rows to 11 matching Office rows, exact spaces/fonts and MAE 0.96→0.249.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 102 control break mismatches remain across unsupported spacing, mixed/decorated paragraph bodies, five paragraphs, overwide/decorated/separated closing runs and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, exact space/font and metric audits, changed legacy inventory, guards and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l90/`. Separate full integration V51 through L89/E4 was still in progress when this slice was qualified and excludes L90 and later prototypes.


## RV06-L91: canonical before-auto-false spacing in plain comment paragraphs

Plain two-, three- and four-paragraph balloons normalize the exact `beforeAutospacing="false"` token, alone or with canonical after-auto and `beforeLines="100"`/`afterLines="100"` while retaining each prepared body and nominal mark face, its own continuation pitch and adjacent transitions. Before-auto on, after-auto on, unqualified line values, malformed tokens, contextual spacing, mixed/decorated bodies and five paragraphs retain complete fallback. The expanded production spacing regression fails before changes; nineteen mixed checks retain 4608 scalar/advance/descent/gap/length/mark/scale body variants and add 2160 before-auto false combinations beyond 8532 earlier canonical spacing combinations across fitting/wrapped bodies, explicit before/after twips and scale. Malformed automatic-spacing assertions remain active. Lexical, numeric and formatting guards remain active.

Frozen runtime `72072402` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `c6180cba1d4b4cf4a2e6d6ae6fb74390` contains the exact tested DLL, SHA-256 `6114C4D4AEB7771B0F673DC244C337409F042F2AAE339E26E9727914F33DF0EB`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 1320 prior and twenty-six fresh controls, 21 improve MAE and SSIM and 1325 retain PDF/raster/graphics bytes, with no image regressions. All 1346 retain main-document text and comment content; 1239 match Office breaks. All 21 changed cases match exact row text, spaces and prepared fonts, including 117 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.284pt. 21 Office metric controls cover 146 continuation pitches and 63 adjacent transitions with maximum errors 0.091/0.141pt. Each changed input is matched to its own hash and cached Office reference; every changed legacy automatic-spacing case is included. All six negative guards retain PDF/raster/graphics bytes. Fresh positive fixtures apply beforeAutospacing=false to every paragraph, with twenty targets covering all twenty accepted absent/0/1/true/false after-auto and before/afterLines combinations, including standalone before-auto false and distinct third-body fonts, marks, right margins and explicit twip spacing. The short Calibri third fits one row within seven body rows total: it is a positive target because the spacing previously forced fallback. Its input is unchanged and its rows, spaces and fonts match Office. The original case improves from 8 candidate rows to 11 matching Office rows, exact spaces/fonts and MAE 0.96→0.249.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 107 control break mismatches remain across unsupported spacing, mixed/decorated paragraph bodies, five paragraphs, overwide/decorated/separated closing runs and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, exact space/font and metric audits, changed legacy inventory, guards and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l91/`. Separate full integration V51 through L89/E4 is complete and excludes L90 and later prototypes.

## RV21-V51: full integration through L89 and smooth SVG transparency

Frozen runtime `f3bb65f8` passes a clean Release build and **2186 passed, 0 failed, 1 skipped**, including 904 DOCX checks. The optional private PPTX layout diagnostic lacks configured input/output. Exact local 0.1.5 package smoke `395f9e630b104098a34adfc9d6a34150` contains the full-suite DLL, SHA-256 `ECA3ED3418E0EBC64EB7068D6EC878CE7044628ADA1B5B617049986597FA25FD`. This run covers L20 through L89/E4 and excludes L90 and later prototypes. Evidence is under `artifacts/plan-revision-20261005/milestone-f3bb65f8/`. Release preparation remains deferred.

## RV06-L92: canonical before-lines200 spacing in plain comment paragraphs

Plain two-, three- and four-paragraph balloons normalize the exact `beforeLines="200"` token, alone or with canonical automatic spacing and `afterLines="100"` while retaining each prepared body and nominal mark face, its own continuation pitch and adjacent transitions. Before-auto on, after-auto on, beforeLines=300, afterLines=200, other unqualified line values, malformed tokens, contextual spacing, mixed/decorated bodies and five paragraphs retain complete fallback. The expanded production spacing regression fails before changes; nineteen mixed checks retain 4608 scalar/advance/descent/gap/length/mark/scale body variants and add 5400 before-line 200 combinations beyond 10692 earlier canonical spacing combinations across fitting/wrapped bodies, explicit before/after twips and scale. Malformed automatic-spacing assertions remain active. Lexical, numeric and formatting guards remain active.

Frozen runtime `4881f61f` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `3e67e286d1bf4b5ab45121bd1435c865` contains the exact tested DLL, SHA-256 `5BFAEB915165882B95274A4ED3DBEFDDDD97EC0EC153982F4D8A1B9C4A8D0311`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 1346 prior and fifty-six fresh controls, 59 improve MAE and SSIM and 1343 retain PDF/raster/graphics bytes, with no image regressions. All 1402 retain main-document text and comment content; 1298 match Office breaks. All 59 changed cases match exact row text, spaces and prepared fonts, including 322 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.284pt. 59 Office metric controls cover 402 continuation pitches and 177 adjacent transitions with maximum errors 0.091/0.141pt. Each changed input is matched to its own hash and cached Office reference; every changed legacy automatic-spacing case is included. All six negative guards retain PDF/raster/graphics bytes. Fresh positive fixtures apply beforeLines=200 to every paragraph, with fifty targets covering all fifty accepted absent/0/1/true/false before/after-auto and absent/100 afterLines combinations, including standalone before-line 200 and distinct third-body fonts, marks, right margins and explicit twip spacing. Three short Calibri-third fixtures each fit one row within seven body rows total under distinct canonical combinations. They are positive targets because the spacing previously forced fallback. Their inputs are unchanged and their rows, spaces and fonts match Office. The original case improves from 8 candidate rows to 11 matching Office rows, exact spaces/fonts and MAE 0.96→0.249.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 104 control break mismatches remain across unsupported spacing, mixed/decorated paragraph bodies, five paragraphs, overwide/decorated/separated closing runs and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, exact space/font and metric audits, changed legacy inventory, guards and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l92/`. Separate full integration V52 through L91/E4 was still in progress when this slice was qualified and excludes L92 and later prototypes.


## RV06-L93: canonical after-lines200 spacing in plain comment paragraphs

Plain two-, three- and four-paragraph balloons normalize the exact `afterLines="200"` token, alone or with canonical automatic spacing and `beforeLines="100"`/`beforeLines="200"` while retaining each prepared body and nominal mark face, its own continuation pitch and adjacent transitions. Before-auto on, after-auto on, beforeLines=300, afterLines=300, other unqualified line values, malformed tokens, contextual spacing, mixed/decorated bodies and five paragraphs retain complete fallback. The expanded production spacing regression fails before changes; nineteen mixed checks retain 4608 scalar/advance/descent/gap/length/mark/scale body variants and add 8100 after-line 200 combinations beyond 16092 earlier canonical spacing combinations across fitting/wrapped bodies, explicit before/after twips and scale. Malformed automatic-spacing assertions remain active. Lexical, numeric and formatting guards remain active.

Frozen runtime `0ef66901` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `d50afc3ed72345959c1c43f4eb6e38ea` contains the exact tested DLL, SHA-256 `14AC3EF6DD87B1754E6C419DA7927D4283DD3FE68A89B0C6FD857652A38BE9D2`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 1402 prior and eighty-one fresh controls, 86 improve MAE and SSIM and 1397 retain PDF/raster/graphics bytes, with no image regressions. All 1483 retain main-document text and comment content; 1384 match Office breaks. All 86 changed cases match exact row text, spaces and prepared fonts, including 479 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.284pt. 86 Office metric controls cover 596 continuation pitches and 258 adjacent transitions with maximum errors 0.091/0.141pt. Each changed input is matched to its own hash and cached Office reference; every changed legacy automatic-spacing case is included. All six negative guards retain PDF/raster/graphics bytes. Fresh positive fixtures apply afterLines=200 to every paragraph, with seventy-five targets covering all seventy-five accepted absent/0/1/true/false before/after-auto and absent/100/200 beforeLines combinations, including standalone after-line 200 and distinct third-body fonts, marks, right margins and explicit twip spacing. Four short Calibri-third fixtures each fit one row within seven body rows total under distinct canonical combinations. They are positive targets because the spacing previously forced fallback. Their inputs are unchanged and their rows, spaces and fonts match Office. The original case improves from 8 candidate rows to 11 matching Office rows, exact spaces/fonts and MAE 0.96→0.249.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 99 control break mismatches remain across unsupported spacing, mixed/decorated paragraph bodies, five paragraphs, overwide/decorated/separated closing runs and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, exact space/font and metric audits, changed legacy inventory, guards and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l93/`. Separate full integration V52 through L91/E4 is complete and excludes L92 and later prototypes.

## RV21-V52: full integration through L91 and smooth SVG transparency

Frozen runtime `72072402` passes a clean Release build and **2186 passed, 0 failed, 1 skipped**, including 904 DOCX checks. The optional private PPTX layout diagnostic lacks configured input/output. Exact local 0.1.5 package smoke `69874a1744e94c299b9c3b01460973a4` contains the full-suite DLL, SHA-256 `20DBE4D5BE873B4EEA0EB2737B830EF62B9AAF876A9EFBC6BC3D9274CA3D3F01`. This run covers L20 through L91/E4 and excludes L92 and later prototypes. Evidence is under `artifacts/plan-revision-20261005/milestone-72072402/`. Release preparation remains deferred.

## RV06-L94: canonical before-auto-on spacing in plain comment paragraphs

Plain two-, three- and four-paragraph balloons normalize the exact `beforeAutospacing="on"` token, alone or with canonical after-auto and `beforeLines="100"`/`beforeLines="200"`/`afterLines="100"`/`afterLines="200"` while retaining each prepared body and nominal mark face, its own continuation pitch and adjacent transitions. Before-auto off, after-auto on, beforeLines=300, afterLines=300, other unqualified line values, malformed tokens, contextual spacing, mixed/decorated bodies and five paragraphs retain complete fallback. The expanded production spacing regression fails before changes; nineteen mixed checks retain 4608 scalar/advance/descent/gap/length/mark/scale body variants and add 4860 before-auto on combinations beyond 24192 earlier canonical spacing combinations across fitting/wrapped bodies, explicit before/after twips and scale. Malformed automatic-spacing assertions remain active. Lexical, numeric and formatting guards remain active.

Frozen runtime `4fc32e90` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `ab7c678910454bf2bd84e70916be1fb9` contains the exact tested DLL, SHA-256 `CE6B5A598F5E52C9B76DDCC94D95B5E0B1BC4782D9D7907A0CE9197680A97508`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 1483 prior and fifty-one fresh controls, 48 improve MAE and SSIM and 1486 retain PDF/raster/graphics bytes, with no image regressions. All 1534 retain main-document text and comment content; 1432 match Office breaks. All 48 changed cases match exact row text, spaces and prepared fonts, including 258 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.284pt. 48 Office metric controls cover 325 continuation pitches and 144 adjacent transitions with maximum errors 0.091/0.141pt. Each changed input is matched to its own hash and cached Office reference; every changed legacy automatic-spacing case is included. All six negative guards retain PDF/raster/graphics bytes. Fresh positive fixtures apply beforeAutospacing=on to every paragraph, with forty-five targets covering all forty-five accepted absent/0/1/true/false after-auto and absent/100/200 before/afterLines combinations, including standalone before-auto on and distinct third-body fonts, marks, right margins and explicit twip spacing. Three short Calibri-third fixtures each fit one row within seven body rows total under distinct canonical combinations. They are positive targets because the spacing previously forced fallback. Their inputs are unchanged and their rows, spaces and fonts match Office. The original case improves from 8 candidate rows to 11 matching Office rows, exact spaces/fonts and MAE 0.96→0.249.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 102 control break mismatches remain across unsupported spacing, mixed/decorated paragraph bodies, five paragraphs, overwide/decorated/separated closing runs and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, exact space/font and metric audits, changed legacy inventory, guards and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l94/`. Separate full integration V53 through L93/E4 was still in progress when this slice was qualified and excludes L94 and later prototypes.


## RV06-L95: canonical after-auto-on spacing in plain comment paragraphs

Plain two-, three- and four-paragraph balloons normalize the exact `afterAutospacing="on"` token, alone or with canonical before-auto and `beforeLines="100"`/`beforeLines="200"`/`afterLines="100"`/`afterLines="200"` while retaining each prepared body and nominal mark face, its own continuation pitch and adjacent transitions. Before-auto off, after-auto off, beforeLines=300, afterLines=300, other unqualified line values, malformed tokens, contextual spacing, mixed/decorated bodies and five paragraphs retain complete fallback. The expanded production spacing regression fails before changes; nineteen mixed checks retain 4608 scalar/advance/descent/gap/length/mark/scale body variants and add 5832 after-auto on combinations beyond 29052 earlier canonical spacing combinations across fitting/wrapped bodies, explicit before/after twips and scale. Malformed automatic-spacing assertions remain active. Lexical, numeric and formatting guards remain active.

Frozen runtime `41c7047f` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `fe212de3455d481d982844dc6c34a35f` contains the exact tested DLL, SHA-256 `C89FF2200DFA6DCC2EAF26A2D90C5D3491F50CB7BB815722F66A22C5B543DDDB`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 1534 prior and sixty fresh controls, 60 improve MAE and SSIM and 1534 retain PDF/raster/graphics bytes, with no image regressions. All 1594 retain main-document text and comment content; 1492 match Office breaks. All 60 changed cases match exact row text, spaces and prepared fonts, including 332 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.284pt. 60 Office metric controls cover 415 continuation pitches and 180 adjacent transitions with maximum errors 0.091/0.141pt. Each changed input is matched to its own hash and cached Office reference; every changed legacy automatic-spacing case is included. All six negative guards retain PDF/raster/graphics bytes. Fresh positive fixtures apply afterAutospacing=on to every paragraph, with fifty-four targets covering all fifty-four accepted absent/0/1/true/false/on before-auto and absent/100/200 before/afterLines combinations, including standalone after-auto on and distinct third-body fonts, marks, right margins and explicit twip spacing. Three short Calibri-third fixtures each fit one row within seven body rows total under distinct canonical combinations. They are positive targets because the spacing previously forced fallback. Their inputs are unchanged and their rows, spaces and fonts match Office. The original case improves from 8 candidate rows to 11 matching Office rows, exact spaces/fonts and MAE 0.96→0.249.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 102 control break mismatches remain across unsupported spacing, mixed/decorated paragraph bodies, five paragraphs, overwide/decorated/separated closing runs and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, exact space/font and metric audits, changed legacy inventory, guards and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l95/`. Separate full integration V53 through L93/E4 is complete and excludes L94 and later prototypes.

## RV21-V53: full integration through L93 and smooth SVG transparency

Frozen runtime `0ef66901` passes a clean Release build and **2186 passed, 0 failed, 1 skipped**, including 904 DOCX checks. The optional private PPTX layout diagnostic lacks configured input/output. Exact local 0.1.5 package smoke `09afa4ba97b849e4a928b8387f0c678b` contains the full-suite DLL, SHA-256 `76A1372C8B5923F4A71CD81E59E59AF76FF99C68D02A54EFF3BB8EF24F406B41`. This run covers L20 through L93/E4 and excludes L94 and later prototypes. Evidence is under `artifacts/plan-revision-20261005/milestone-0ef66901/`. Release preparation remains deferred.

## RV06-L96: canonical before-auto-off spacing in plain comment paragraphs

Plain two-, three- and four-paragraph balloons normalize the exact `beforeAutospacing="off"` token, alone or with canonical after-auto and `beforeLines="100"`/`beforeLines="200"`/`afterLines="100"`/`afterLines="200"` while retaining each prepared body and nominal mark face, its own continuation pitch and adjacent transitions. Before-auto yes, after-auto off, beforeLines=300, afterLines=300, other unqualified line values, malformed tokens, contextual spacing, mixed/decorated bodies and five paragraphs retain complete fallback. The expanded production spacing regression fails before changes; nineteen mixed checks retain 4608 scalar/advance/descent/gap/length/mark/scale body variants and add 5832 before-auto off combinations beyond 34884 earlier canonical spacing combinations across fitting/wrapped bodies, explicit before/after twips and scale. Malformed automatic-spacing assertions remain active. Lexical, numeric and formatting guards remain active.

Frozen runtime `855f2773` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `7c4e6caf83254fde8c71123eff722574` contains the exact tested DLL, SHA-256 `CB92FDADD0CBC0FDA9F6F12E2BDE30B0FB42FEE7BFF58BD52C28689CD6CA0F1A`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 1594 prior and sixty fresh controls, 56 improve MAE and SSIM and 1598 retain PDF/raster/graphics bytes, with no image regressions. All 1654 retain main-document text and comment content; 1548 match Office breaks. All 56 changed cases match exact row text, spaces and prepared fonts, including 308 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.284pt. 56 Office metric controls cover 387 continuation pitches and 168 adjacent transitions with maximum errors 0.091/0.141pt. Each changed input is matched to its own hash and cached Office reference; every changed legacy automatic-spacing case is included. All six negative guards retain PDF/raster/graphics bytes. Fresh positive fixtures apply beforeAutospacing=off to every paragraph, with fifty-four targets covering all fifty-four accepted absent/0/1/true/false/on after-auto and absent/100/200 before/afterLines combinations, including standalone before-auto off and distinct third-body fonts, marks, right margins and explicit twip spacing. Three short Calibri-third fixtures each fit one row within seven body rows total under distinct canonical combinations. They are positive targets because the spacing previously forced fallback. Their inputs are unchanged and their rows, spaces and fonts match Office. The original case improves from 8 candidate rows to 11 matching Office rows, exact spaces/fonts and MAE 0.96→0.249.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 106 control break mismatches remain across unsupported spacing, mixed/decorated paragraph bodies, five paragraphs, overwide/decorated/separated closing runs and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, exact space/font and metric audits, changed legacy inventory, guards and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l96/`. Separate full integration V54 through L95/E4 was still in progress when this slice was qualified and excludes L96 and later prototypes.


## RV06-L97: canonical after-auto-off spacing in plain comment paragraphs

Plain two-, three- and four-paragraph balloons normalize the exact `afterAutospacing="off"` token, alone or with canonical before-auto and `beforeLines="100"`/`beforeLines="200"`/`afterLines="100"`/`afterLines="200"` while retaining each prepared body and nominal mark face, its own continuation pitch and adjacent transitions. Before-auto yes, after-auto yes, beforeLines=300, afterLines=300, other unqualified line values, malformed tokens, contextual spacing, mixed/decorated bodies and five paragraphs retain complete fallback. The expanded production spacing regression fails before changes; nineteen mixed checks retain 4608 scalar/advance/descent/gap/length/mark/scale body variants and add 6804 after-auto off combinations beyond 40716 earlier canonical spacing combinations across fitting/wrapped bodies, explicit before/after twips and scale. Malformed automatic-spacing assertions remain active. Lexical, numeric and formatting guards remain active.

Frozen runtime `cead2616` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `408c21aa2aad41f7bad04221218a9e4f` contains the exact tested DLL, SHA-256 `2428A336B3517A6496D4C1E8BA88B6E93B560A32460C50BC528757AACBE9D27D`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 1654 prior and sixty-nine fresh controls, 65 improve MAE and SSIM and 1658 retain PDF/raster/graphics bytes, with no image regressions. All 1723 retain main-document text and comment content; 1613 match Office breaks. All 65 changed cases match exact row text, spaces and prepared fonts, including 354 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.284pt. 65 Office metric controls cover 445 continuation pitches and 195 adjacent transitions with maximum errors 0.091/0.141pt. Each changed input is matched to its own hash and cached Office reference; every changed legacy automatic-spacing case is included. All six negative guards retain PDF/raster/graphics bytes. Fresh positive fixtures apply afterAutospacing=off to every paragraph, with sixty-three targets covering all sixty-three accepted absent/0/1/true/false/on/off before-auto and absent/100/200 before/afterLines combinations, including standalone after-auto off and distinct third-body fonts, marks, right margins and explicit twip spacing. Four short Calibri-third fixtures each fit one row within seven body rows total under distinct canonical combinations. They are positive targets because the spacing previously forced fallback. Their inputs are unchanged and their rows, spaces and fonts match Office. The original case improves from 8 candidate rows to 11 matching Office rows, exact spaces/fonts and MAE 0.96→0.249.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 110 control break mismatches remain across unsupported spacing, mixed/decorated paragraph bodies, five paragraphs, overwide/decorated/separated closing runs and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, exact space/font and metric audits, changed legacy inventory, guards and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l97/`. Separate full integration V54 through L95/E4 is complete and excludes L96 and later prototypes.

## RV21-V54: full integration through L95 and smooth SVG transparency

Frozen runtime `41c7047f` passes a clean Release build and **2186 passed, 0 failed, 1 skipped**, including 904 DOCX checks. The optional private PPTX layout diagnostic lacks configured input/output. Exact local 0.1.5 package smoke `18d6ebf07463475c9b85209c6943d0b0` contains the full-suite DLL, SHA-256 `30D55D8E226C2F612C7C2687AFC5EA2E6FF486D82F3ABACD6029FD0B7D98FA46`. This run covers L20 through L95/E4 and excludes L96 and later prototypes. Evidence is under `artifacts/plan-revision-20261005/milestone-41c7047f/`. Release preparation remains deferred.

## RV06-L98: canonical beforeLines=300 spacing in plain comment paragraphs

Plain two-, three- and four-paragraph balloons normalize exact `beforeLines="300"`, alone or with canonical before/after-auto and `afterLines="100"`/`afterLines="200"` while retaining each prepared body and nominal mark face, its own continuation pitch and adjacent transitions. Before-auto yes, after-auto yes, beforeLines=400, afterLines=300, other unqualified line values, malformed tokens, contextual spacing, mixed/decorated bodies and five paragraphs retain complete fallback. The expanded production spacing regression fails before changes; nineteen mixed checks retain 4608 scalar/advance/descent/gap/length/mark/scale body variants and add 15876 beforeLines=300 combinations beyond 47520 earlier canonical spacing combinations across fitting/wrapped bodies, explicit before/after twips and scale. Malformed automatic-spacing assertions remain active. Lexical, numeric and formatting guards remain active.

Frozen runtime `871a9fcb` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `209d7ea4a7da48b49b35215b366e1a97` contains the exact tested DLL, SHA-256 `DFF2B9907B08409DCE406A8943676F833444D5A8502346B452B1C7F39F5AFEFD`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 1723 prior and 153 fresh controls, 153 improve MAE and SSIM and 1723 retain PDF/raster/graphics bytes, with no image regressions. All 1876 retain main-document text and comment content; 1766 match Office breaks. All 153 changed cases match exact row text, spaces and prepared fonts, including 840 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.284pt. 153 Office metric controls cover 1054 continuation pitches and 459 adjacent transitions with maximum errors 0.091/0.141pt. Each changed input is matched to its own hash and cached Office reference; every changed legacy automatic-spacing case is included. All six negative guards retain PDF/raster/graphics bytes. Fresh positive fixtures apply beforeLines=300 to every paragraph, with 147 targets covering all 147 accepted absent/0/1/true/false/on/off before/after-auto and absent/100/200 afterLines combinations, including standalone beforeLines=300 and distinct third-body fonts, marks, right margins and explicit twip spacing. Eight short Calibri-third fixtures each fit one row within seven body rows total under distinct canonical combinations. They are positive targets because the spacing previously forced fallback. Their inputs are unchanged and their rows, spaces and fonts match Office. The original case improves from 8 candidate rows to 11 matching Office rows, exact spaces/fonts and MAE 0.96→0.249.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 110 control break mismatches remain across unsupported spacing, mixed/decorated paragraph bodies, five paragraphs, overwide/decorated/separated closing runs and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, exact space/font and metric audits, changed legacy inventory, guards and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l98/`. Separate full integration V55 through L97/E4 was still in progress when this slice was qualified and excludes L98 and later prototypes.


## RV06-L99: canonical afterLines=300 spacing in plain comment paragraphs

Plain two-, three- and four-paragraph balloons normalize exact `afterLines="300"`, alone or with canonical before/after-auto and `beforeLines="100"`/`beforeLines="200"`/`beforeLines="300"` while retaining each prepared body and nominal mark face, its own continuation pitch and adjacent transitions. Before-auto yes, after-auto yes, beforeLines=400, afterLines=400, other unqualified line values, malformed tokens, contextual spacing, mixed/decorated bodies and five paragraphs retain complete fallback. The expanded production spacing regression fails before changes; nineteen mixed checks retain 4608 scalar/advance/descent/gap/length/mark/scale body variants and add 21168 afterLines=300 combinations beyond 63396 earlier canonical spacing combinations across fitting/wrapped bodies, explicit before/after twips and scale. Malformed automatic-spacing assertions remain active. Lexical, numeric and formatting guards remain active.

Frozen runtime `e233cb4d` passes a clean Release build, **904 DOCX checks, 100 balloon checks and twenty-five Linux regressions**. Fresh local 0.1.5 package smoke `9b284c544e3e403abb702a34aee3f58a` contains the exact tested DLL, SHA-256 `E95BDFC07BEFF839707201E3DDBE0CEB6D6FDE22C4CB2705CBAD73C687762A0A`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 1876 prior and 202 fresh controls, 202 improve MAE and SSIM and 1876 retain PDF/raster/graphics bytes, with no image regressions. All 2078 retain main-document text and comment content; 1968 match Office breaks. All 202 changed cases match exact row text, spaces and prepared fonts, including 1121 within-word break rows, with maximum row-start X gap 0.838pt and baseline gap 0.284pt. 202 Office metric controls cover 1402 continuation pitches and 606 adjacent transitions with maximum errors 0.091/0.141pt. Each changed input is matched to its own hash and cached Office reference; every changed legacy automatic-spacing case is included. All six negative guards retain PDF/raster/graphics bytes. Fresh positive fixtures apply afterLines=300 to every paragraph, with 196 targets covering all 196 accepted absent/0/1/true/false/on/off before/after-auto and absent/100/200/300 beforeLines combinations, including standalone afterLines=300 and distinct third-body fonts, marks, right margins and explicit twip spacing. Ten short Calibri-third fixtures each fit one row within seven body rows total under distinct canonical combinations. They are positive targets because the spacing previously forced fallback. Their inputs are unchanged and their rows, spaces and fonts match Office. The original case improves from 8 candidate rows to 11 matching Office rows, exact spaces/fonts and MAE 0.96→0.249.

Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no gate increases. 110 control break mismatches remain across unsupported spacing, mixed/decorated paragraph bodies, five paragraphs, overwide/decorated/separated closing runs and punctuation/whitespace guards. Prior positioning/composed pixel residuals and 33 partial cached cases remain. Qualification, exact space/font and metric audits, changed legacy inventory, guards and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l99/`. Separate full integration V55 through L97/E4 is complete and excludes L98 and later prototypes.

## RV21-V55: full integration through L97 and smooth SVG transparency

Frozen runtime `cead2616` passes a clean Release build and **2186 passed, 0 failed, 1 skipped**, including 904 DOCX checks. The optional private PPTX layout diagnostic lacks configured input/output. Exact local 0.1.5 package smoke `a09f98b4ff8c4afca383af7be4aea3ff` contains the full-suite DLL, SHA-256 `F522EEDBC94482F1FA25DF447020BECCFE74C3B5A17D0FCEDB842693CCC430F8`. This run covers L20 through L97/E4 and excludes L98 and later prototypes. Evidence is under `artifacts/plan-revision-20261005/milestone-cead2616/`. Release preparation remains deferred.

## RV06-L100: five plain comment paragraphs with independent body metrics

Word-compatible all-markup balloons with five plain comment paragraphs retain each prepared body and nominal-size paragraph-mark face. A separate bounded five-body layout supplies each continuation pitch and all four adjacent transitions to both balloon height and text emission. The DOCX reader retains marks for one through five plain comment paragraphs in all-markup mode; Final view and larger stories keep their prior policy. Mixed/decorated bodies, drawings, six or more paragraphs, unqualified spacing and invalid body/transition metrics retain fallback. Existing one-to-four-paragraph rendering paths remain unchanged. Both renderer and reader production regressions fail before their respective fixes. Twenty mixed checks retain 4608 earlier body and 84672 canonical spacing variants and add 256 five-body fitting/wrapped masks, Unicode scalars, mark faces, line gaps and right margins. Zero-width, overwide-scalar, negative-pitch, invalid-transition, formatting and spacing guards remain active.

Frozen runtime `f4b15257` passes a clean Release build, **905 DOCX checks, 101 balloon checks and twenty-six Linux regressions**. Fresh local 0.1.5 package smoke `170fdd2da7064fb88aa76df4e572505a` contains the exact tested DLL, SHA-256 `E2E17A27791D0A749048E72115CA4997B5C25276457D821ADF08A13DF2D834D4`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 2078 prior controls, 42 fresh fixtures and four independent five-font profiles, 74 improve MAE and SSIM and 2050 retain PDF/raster/graphics bytes, with no image regressions. All 2124 retain main-document text and comment content; 2042 match Office breaks. All 74 changed cases are five-paragraph inputs and match exact row text, spaces and prepared fonts, including 383 within-word break rows. Maximum row-start X and baseline gaps are 0.085/0.308pt. Their own prepared-font models cover 453 continuation pitches and 296 adjacent transitions, with maximum errors 0.105/0.166pt. Wrapped Georgia continuation pitch differs by up to 0.106pt; the Calibri→Georgia transition residual reaches 0.166pt in the fresh all-fitting case, beyond the earlier four-paragraph envelope; this remains an explicit approximation. Each changed input is matched by its own SHA-256 to its cached Office reference. All 34 changed legacy five-paragraph inputs are included. Every previously qualified one-to-four-paragraph output retains bytes.

The 36 fresh positive targets cover all 32 fitting/wrapped masks and four distinct-font/canonical-spacing profiles. The all-fitting case has exactly five Office-matching rows; fifth-body wrapping and all four adjacent transitions are measured independently. Six fresh negative guards retain PDF/raster/graphics bytes. Four additional five-font profiles match thirteen Office rows each using frozen L99 baselines and their original Office references. The original case improves from 8 candidate rows to 12 matching Office rows, exact spaces/fonts and MAE 1.008→0.231. Of 34 cached Office cases, 34 retain PDF bytes and the aggregate is 545 failed gates with no increases.

82 control break mismatches remain across unsupported spacing, mixed/decorated bodies, larger comments, closing-run and punctuation/whitespace guards. Prior positioning/composed pixel residuals and partial cached cases remain. Qualification, input/reference identities, exact emission and own-font metric audits are under `artifacts/plan-revision-20261005/rv06-l100/`. Separate full integration V56 through L99/E4 was still in progress when this slice was qualified and excludes L100 and later prototypes.


## RV06-L101: two fitting body faces in a fourth comment paragraph

Four-paragraph word-compatible all-markup balloons retain two plain prepared body faces and the nominal paragraph-mark face on a fitting fourth row. The first three paragraphs retain their own fitting/wrapped rows and continuation pitches. One bounded layout supplies all three adjacent transitions to both geometry and emission; the transition into the fourth row uses the preceding body's descent plus the larger body ascent/gap of its two faces. Reversed-font Office probes discriminate this rule: Courier New first misses by 0.821pt when only its ascent is used, versus 0.019pt using the larger Cambria ascent. Wrapped mixed fourth rows, other mixed/decorated bodies, three fourth-body runs, unqualified spacing and invalid metrics retain fallback. Earlier plain one-to-five-paragraph paths remain unchanged.

The production renderer regression fails before change. A separate corrected numeric regression fails at the fourth-body pitch guard when that guard is removed, then passes when restored. Twenty-one mixed checks preserve 4608 earlier body, 256 five-body and 84672 canonical spacing variants and add 256 mixed-fourth leading-wrap-mask, Unicode, face-order, line-gap, margin and canonical-spacing variants. Each preceding pitch and transition, both fourth-body advances, nominal mark, content and geometry are checked. Zero-width, overwide-scalar, invalid leading/fourth-body pitch, formatting, run-count, wrapping and spacing guards remain active.

Frozen runtime `00eb04fd` passes a clean Release build, **906 DOCX checks, 102 balloon checks and twenty-seven Linux regressions**. Fresh local 0.1.5 package smoke `9cbeaa5f6717476da5119ee09f2aec02` contains the exact tested DLL, SHA-256 `B676326733B5DDF9D1C07FF2C31BD4BE2117CF7CBDCFA0948EFDA25B4EE77EE6`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 2124 prior controls, 42 fresh fixtures and four independent reversed-font probes, 41 improve MAE and SSIM and 2129 retain PDF/raster/graphics bytes, with no image regressions. All 2170 retain main-document text and comment content; 2083 match Office breaks. All 41 changed inputs match exact extracted row text, spaces and prepared fonts, including 130 within-word rows. Maximum row-start X/baseline gaps are 0.085/0.277pt. Own prepared-face models cover 145 preceding continuation pitches and 123 adjacent transitions, with maximum errors 0.085/0.135pt. The separate fourth-run audit covers eighty-two body runs and forty-one terminal marks; their maximum start-X gaps are 0.128/0.095pt. Metric paragraph partitioning normalizes extracted whitespace at body-font changes; the independent emission audit still compares exact Office row spaces/fonts.

The 36 fresh positive targets cover four font profiles across all eight preceding wrap masks plus four canonical-spacing profiles. Four all-fitting targets match four Office rows. All six fresh negative guards retain PDF/raster/graphics bytes. The single fitting legacy mixed-fourth input improves from 2 to 4 matching Office rows, exact spaces/fonts and MAE 0.294→0.135; the 26 cached wrapped mixed-fourth cases retain fallback. Each input/reference pair is linked by its own hash. Every previously qualified plain one-to-five-paragraph output retains bytes. All 34 cached reference PDFs retain bytes, with 545 failed gates and no increases.

87 control break mismatches remain across wrapped mixed fourth rows, other mixed/decorated bodies, larger stories, unqualified spacing and closing-run/punctuation/whitespace guards. Prior positioning/composed pixel residuals and partial cached cases remain. Evidence, exact emission, own-font metrics, fourth-run positions, guards and input/reference identities are under `artifacts/plan-revision-20261005/rv06-l101/`. Separate full integration V57 through L100/E4 was still in progress when this slice was qualified and excludes L101 and later prototypes.

## RV21-V56: full integration through L99 and smooth SVG transparency

Frozen runtime `e233cb4d` passes a clean Release build and **2186 passed, 0 failed, 1 skipped**, including 904 DOCX checks. The optional private PPTX layout diagnostic lacks configured input/output. Exact local 0.1.5 package smoke `7c4d6be725344720bbf754eb837f9309` contains the full-suite DLL, SHA-256 `7DF008682138BEE8BAD97B8BF57AA6F53F68A11787955A84F05AC894E3E69EFD`. This run covers L20 through L99/E4 and excludes L100 and later prototypes. Evidence is under `artifacts/plan-revision-20261005/milestone-e233cb4d/`. Release preparation remains deferred.

## RV06-L102: wrapped fourth comment body with a fitting closing run

Four-paragraph word-compatible all-markup balloons retain the first fourth-body face across Unicode word splits and a distinct fitting closing face on its final row. When their joined row would overflow, the first body's last whole word moves down with the closing run; a joined word that still cannot fit retains fallback. Each preceding paragraph uses its own rows and pitch. The first fourth-row transition uses its first face; ordinary fourth continuations use that face's pitch, and the final mixed row uses its preceding first-face descent plus the larger of both body ascents/gaps. One bounded layout supplies balloon height and emission. A trailing first-body separator, other mixed/decorated bodies, three fourth-body runs and unqualified spacing retain fallback. Earlier fitting mixed-fourth and plain one-to-five-paragraph paths retain bytes.

The production renderer regression fails on the frozen parent. An expanded prefix-length regression also fails without final-word reflow, then passes with it. Twenty-two mixed checks preserve 4608 earlier body, 256 five-body, 256 fitting-fourth and 84672 canonical spacing variants and add 768 wrapped-fourth prefix-length, preceding-wrap-mask, Unicode, face-order, line-gap, margin and canonical-spacing combinations. Content, source faces, body/mark advances, ordinary/final pitches, all paragraph transitions and height are checked. Zero-width, overwide-scalar, invalid body/closing pitch, formatting, nonfitting closing run and trailing-separator guards remain active.

Frozen runtime `47e876e4` passes clean Release, **907 DOCX checks, 103 balloon checks and twenty-eight Linux regressions**. Fresh local 0.1.5 package smoke `4711d5ee258d494c8f59c0ac67808cd3` contains the exact tested DLL, SHA-256 `DEE1815166AEE20B64BD4DCE7B9899754007894C077FF4AC2CB356AA22540F20`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 2170 prior controls, 46 fresh fixtures and four independent Office probes, 71 improve MAE and SSIM and 2149 retain PDF/raster/graphics bytes, with no image regressions. All 2220 retain main-document text and comment content; 2154 match Office breaks. All 71 changed inputs match exact extracted row text, spaces and prepared fonts, including 357 within-word rows. Maximum row-start X/baseline gaps are 0.085/0.29pt. Own prepared-face models cover 451 continuation pitches and 213 adjacent transitions, with maximum errors 0.099/0.135pt. Final body/mark start-X gaps are 0.194/0.116pt. Metric partitioning normalizes extracted whitespace at body-font changes; the independent emission audit compares exact Office row spaces/fonts.

The forty fresh targets cover four font profiles across all eight preceding wrap masks, three prefix lengths, four canonical-spacing profiles and four explicitly preserved leading-space profiles. All six negative fixtures retain PDF/raster/graphics bytes. All 26 cached wrapped mixed-fourth cases and the preceding slice's wrapped guard improve and match Office breaks; the original changes from 3 to 6 matching rows and MAE 0.509→0.156. All 41 qualified fitting mixed-fourth outputs retain PDF/raster/graphics bytes. Input/reference pairs retain their own hashes. All 34 cached reference PDFs retain bytes, with 545 failed gates and no increases.

Independent reversed-font probes reject applying one ordinary pitch to the final mixed row: Courier New→Cambria yields 8.56pt in Office, 7.721pt from the ordinary pitch and 8.523pt from the final-row model. Arial→Cambria retains a measured 0.099pt final-pitch residual. Prior five-paragraph Georgia/Calibri transitions, positioning and composed pixel residuals remain explicit. 66 control break mismatches and partial cached cases remain. Evidence is under `artifacts/plan-revision-20261005/rv06-l102/`. Full integration scope is recorded separately; release preparation remains deferred.

A separate Office-only Cambria→Courier New counterprobe rejects taking the larger descent of both final-row bodies. Office's final pitch is 8.02pt; previous-first-body descent predicts 7.991pt, versus 8.523pt using maximum descent. The latter would improve the Arial/Cambria probe but worsen this reverse pair. Its reference/input identities and rejected-model audit remain separate from the 2220 candidate-control inventory.

## RV21-V57: full integration through L100 and smooth SVG transparency

Frozen runtime `f4b15257` passes clean Release and **2187 passed, 0 failed, 1 skipped**, including 905 DOCX checks. The optional private PPTX layout diagnostic lacks configured input/output. Exact local 0.1.5 package smoke `df55a00d73634a5b8ec7d180273978c8` contains the full-suite DLL, SHA-256 `E9F20FF1EE067CDDFE6A354BDF4F580AC93AB39636AD050A042830F14A5D2D26`. This run covers L20 through L100/E4 and excludes L101 and later prototypes. Evidence is under `artifacts/plan-revision-20261005/milestone-f4b15257/`. Release preparation remains deferred.

## RV06-L103: two fitting terminal body faces in three comment paragraphs

Three-paragraph word-compatible all-markup balloons retain both plain fitting terminal body faces and the nominal paragraph-mark face. Each preceding paragraph retains its fitting/wrapped rows and own continuation pitch. The two adjacent transitions share the geometry/emission model; the final transition uses the preceding body's descent plus the larger terminal-body ascent/gap. The prepared terminal model also preserves the qualified four-paragraph fitting/wrapped paths. Wrapped mixed thirds, mixed intermediate bodies, extra runs, decoration and unqualified spacing retain fallback.

The production renderer regression fails on the frozen parent, then passes across 128 preceding-wrap-mask, Unicode, face-order, line-gap, margin and canonical-spacing patterns. Twenty-three mixed checks preserve 4608 earlier body, 256 five-body, 256 fitting-fourth, 768 wrapped-fourth and 84672 spacing variants. Own pitches, both transitions, both final-body advances, nominal mark, exact content and height are checked. Numeric, width, run-count and formatting guards remain active. An older fitting-third fallback assertion now covers an actually wrapped mixed third, which retains fallback.

Frozen runtime `14af00c0` passes clean Release, **908 DOCX checks, 104 balloon checks and twenty-nine Linux regressions**. Exact local 0.1.5 package smoke `39e09bd8c33242aab8ba99a1cde87875` contains the tested DLL, SHA-256 `A220A2AA82C57647B4A871E603C91260DA10A187C33ED26E86AD2E2B11DE09A6`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 2220 prior controls, thirty fresh fixtures and four independent font pairs, 30 improve MAE and SSIM and 2224 retain PDF/raster/graphics bytes, with no image regressions. All 2254 retain main-document text and comment content; 2184 match Office breaks. All thirty changed inputs match exact extracted row text, spaces and prepared fonts, including 77 within-word rows. Maximum row-start X/baseline gaps are 0.085/0.234pt. Own prepared-face models cover 82 continuation pitches and sixty transitions, with maximum errors 0.075/0.111pt. Terminal body/mark start-X gaps are 0.169/0.108pt.

Twenty-four targets cover four font profiles across all four preceding wrap masks, four canonical-spacing profiles and four preserved-leading-space profiles. All six negative fixtures retain PDF/raster/graphics bytes. Both cached fitting-third inputs improve and match Office's three- and five-row layouts. All 112 qualified fitting/wrapped fourth-body inputs retain PDF/raster/graphics bytes. The exact run-position audit includes separately emitted Office leading blanks in their own body face; exact text/spaces and advances are retained across operator segmentation. Every input/reference pair retains its own hash. All 34 cached reference PDFs retain bytes, with 545 failed gates and no increases.

70 control break mismatches remain. Wrapped mixed thirds and mixed intermediate paragraphs remain guarded. Prior Arial/Cambria final-pitch, five-paragraph Georgia/Calibri transition, positioning and composed pixel residuals remain explicit. The rejected maximum-final-descent model and its independent Cambria/Courier New counterprobe remain preserved under L102. Evidence is under `artifacts/plan-revision-20261005/rv06-l103/`. Full-suite scope is recorded separately and release preparation remains deferred.

## RV21-V58: full integration through L101 and smooth SVG transparency

Frozen runtime `00eb04fd` passes clean Release and **2188 passed, 0 failed, 1 skipped**, including 906 DOCX checks. The optional private PPTX layout diagnostic lacks configured input/output. Exact local 0.1.5 package smoke `015b18cb924740a49da681b4159cb79f` contains the full-suite DLL, SHA-256 `5A2583347F5B3833C4773792C39A41C8D7681D3D01A1C3EBB9528E0D1264460C`. This run covers L20 through L101/E4 and excludes L102 and later prototypes. Evidence is under `artifacts/plan-revision-20261005/milestone-00eb04fd/`. Release preparation remains deferred.
## RV21-V59: full integration through L102 and smooth SVG transparency

Frozen runtime `47e876e4` passes clean Release and **2189 passed, 0 failed, 1 skipped**, including 907 DOCX checks. The optional private PPTX layout diagnostic lacks configured input/output. Exact local 0.1.5 package smoke `c14d15f01db245d2930c7b78501e87ce` contains the full-suite DLL, SHA-256 `59E58C0F36DD42FECA6904317C1D7241A4E0AD0A8AA62678129DDF306365FF92`. This run covers L20 through L102/E4 and excludes L103 and later prototypes. Evidence is under `artifacts/plan-revision-20261005/milestone-47e876e4/`. Release preparation remains deferred.

## RV06-L104: wrapped terminal body before a fitting closing face in three comment paragraphs

Three-paragraph word-compatible all-markup balloons wrap the first terminal body run in its own prepared face, followed by a fitting distinct closing face and nominal paragraph mark. Each preceding paragraph retains fitting/wrapped rows and its own pitch. Ordinary terminal continuations use first-body metrics; the final mixed row uses the preceding first-body descent plus the larger final-body ascent/gap. When needed, the final whole word reflows with the closing run. A single preserved leading separator stays in its closing face on the preceding row while a fitting closing word moves to its own row; that row uses only closing ascent/gap plus preceding-first descent. Height and emission share both adjacent transitions. Nonfitting closing runs, trailing first-body separators, extra runs, decoration, mixed intermediate bodies and unqualified spacing retain fallback.

The wrapped-third production regression fails on frozen L103. A second production regression rejects moving the preceding word with a preserved breakable closing run. Both pass after change, alongside 384 prefix-length, preceding-wrap-mask, Unicode, face-order, line-gap, margin and canonical-spacing patterns. Twenty-five mixed checks also retain 4608 earlier body, 256 five-body, 256 fitting-fourth, 768 wrapped-fourth, 128 fitting-third and 84672 spacing variants. Own pitches, both transitions, exact content, final body advances, nominal mark and height are checked. Numeric, scalar-progress, width, run-count and formatting guards remain active. Former wrapped-third fallback assertions now use closing runs that cannot fit; the newly supported shapes are covered positively.

Frozen runtime `d6baeaeb` passes clean Release, **910 DOCX checks, 106 balloon checks and thirty-one Linux regressions**. Exact local 0.1.5 package smoke `d8206d72d87242fc89353a855c4fbc25` contains the tested DLL, SHA-256 `52D718300EBE96CC3547641CECBEAA49E93DAED3C0B96002DE60A96E0B3A5491`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 2254 prior controls, thirty fresh fixtures and four independent font pairs, 31 improve MAE and SSIM and 2257 retain PDF/raster/graphics bytes, with no image regressions. All 2288 retain main-document text and comment content; 2215 match Office breaks. All 31 changed inputs match exact extracted row text, spaces and prepared fonts, including 119 within-word rows. Maximum row-start X/baseline gaps are 0.085/0.291pt. Independent prepared-face models cover 145 pitches and 62 transitions, with maximum errors 0.089/0.115pt. Final body/mark start-X gaps are 0.194/0.144pt.

Twenty-four fresh targets cover four font profiles across all four preceding wrap masks, four canonical-spacing profiles and four preserved-leading-space profiles. Six negative fixtures retain PDF/raster/graphics bytes. The two cached wrapped-third inputs and the prior fresh wrapped-third guard improve and match Office. All 142 qualified fitting-third, fitting-fourth and wrapped-fourth paths retain PDF/raster/graphics bytes. Exact final-run audits include Office's separately emitted leading blanks in their own body face, preserving text, spaces, starts and advances across operator segmentation. Every input/reference pair retains its own hash. All 34 cached reference PDFs retain bytes, with 545 failed gates and no increases.

73 control break mismatches remain. Nonfitting closing runs, mixed intermediate paragraphs and other composed residuals remain guarded. Prior Arial/Cambria final-pitch, five-paragraph Georgia/Calibri transition, positioning and composed pixel errors remain explicit. The rejected maximum-final-descent model and independent Cambria/Courier New counterprobe remain preserved under L102. Evidence is under `artifacts/plan-revision-20261005/rv06-l104/`. Full integration is recorded separately; release preparation remains deferred.


## RV06-L105: wrapped terminal body before a fitting closing face in two comment paragraphs

Two-paragraph word-compatible all-markup balloons wrap the first terminal body run in its own prepared face, followed by a fitting distinct closing face and nominal paragraph mark. A scoped preparation call admits the single preceding paragraph while ordinary paragraph preparation retains its existing bounds. Both bodies retain their own rows and pitches; the single adjacent transition uses the preceding descent plus terminal-first ascent/gap. The final mixed row retains preceding-first descent plus the larger terminal-body ascent/gap, with whole-word reflow when needed. Height and emission share the model. Fitting mixed seconds, preserved two-paragraph closing-only overflow, nonfitting closing runs, trailing first-body separators, extra runs, decoration, mixed intermediate bodies and unqualified spacing retain fallback.

The production regression fails on frozen L104 because the preceding paragraph loses separate body/nominal-mark emission, then passes across 192 prefix-length, preceding-wrap-mask, Unicode, face-order, line-gap, margin and canonical-spacing patterns. Twenty-six mixed checks retain 4608 earlier body, 256 five-body, 256 fitting-fourth, 768 wrapped-fourth, 128 fitting-third, 384 wrapped-third and 84672 spacing variants plus the three-paragraph preserved closing-only regression. Own pitches, the single transition, final body advances, nominal mark, content and height are checked with finite/numeric/scalar/width/formatting guards. An earlier wrapped-second fallback assertion now uses a closing run that cannot fit. The initial indexing failure in the copied test is retained separately from the actual renderer failure proof.

Frozen runtime `1bbff13a` passes clean Release, **911 DOCX checks, 107 balloon checks and thirty-two Linux regressions**. Exact local 0.1.5 package smoke `fc0eb020908d43118275cd516167a57f` contains the tested DLL, SHA-256 `6DEE67540C48EE7F86A3065DEE60F304D12F1235B0B3F5F17A71A849D12B6F8C`. The Linux source archive emits two SourceLink warnings and zero build errors.

Across 2288 prior controls, twenty-four fresh fixtures and four independent font pairs, 22 improve MAE and SSIM and 2294 retain PDF/raster/graphics bytes, with no image regressions. All 2316 retain main-document text and comment content; 2237 match Office breaks. All 22 changed inputs match exact extracted row text, spaces and prepared fonts, including 68 within-word rows. Maximum row-start X/baseline gaps are 0.085/0.281pt. Independent prepared-face models cover 82 pitches and 22 transitions, with maximum errors 0.082/0.141pt. Final body/mark start-X gaps are 0.194/0.122pt.

Sixteen targets cover four font profiles across both preceding wrap masks, four canonical-spacing profiles and four preserved-leading-space profiles. Eight negative fixtures retain PDF/raster/graphics bytes, including fitting-two and preserved closing-only cases outside this slice. Both cached wrapped-second inputs improve and match Office's four- and seven-row layouts. All 173 earlier fitting/wrapped terminal paths, including L104's two preserved closing-only cases, retain PDF/raster/graphics bytes. Exact run audits preserve leading blanks, face signatures, starts and advances across PDF operator segmentation. Every input/reference pair retains its own hash. All 34 cached reference PDFs retain bytes, with 545 failed gates and no increases.

79 control break mismatches remain. Fitting mixed seconds, preserved two-paragraph closing-only overflow, mixed intermediate paragraphs and composed residuals remain guarded. Prior Arial/Cambria final-pitch, five-paragraph Georgia/Calibri transition, positioning and composed pixel errors remain explicit. The rejected maximum-final-descent model and independent Cambria/Courier New counterprobe remain preserved under L102. Evidence is under `artifacts/plan-revision-20261005/rv06-l105/`. Full integration is recorded separately; release preparation remains deferred.

## RV21-V60: full integration through L103 and smooth SVG transparency

Frozen runtime `14af00c0` passes clean Release and **2190 passed, 0 failed, 1 skipped**, including 908 DOCX checks. The optional private PPTX layout diagnostic lacks configured input/output. Exact local 0.1.5 package smoke `38c2894c637746d5931985e7337292c4` contains the full-suite DLL, SHA-256 `9EBF9438BC6C64632A2A97BAB85778961B762C1E8AD001E7AB867D8037CE97F5`. This run covers L20 through L103/E4 and excludes L104 and later prototypes. Evidence is under `artifacts/plan-revision-20261005/milestone-14af00c0/`. Release preparation remains deferred.

## RV06-L106: two fitting terminal body faces in two paragraphs

Frozen runtime `3d3038c7770eb84da29a48f304ee1662c63acf16` extends the fitting mixed terminal resolver to one preceding paragraph. Existing scoped preparation supplies that paragraph; its own rows and continuation pitch remain independent. The terminal row retains both body faces, their measured advances and the nominal-size mark. The preceding body's descent plus the larger terminal-body ascent/gap defines the single transition used by height and emission. Wrapped-second and three-/four-paragraph terminal behavior retain bytes.

A production regression fails before change because fallback collapses preceding and terminal content. Clean Release, 27 mixed checks and one paragraph-reader check pass, including 64 new wrap-mask/Unicode/font-order/gap/margin/canonical-spacing combinations and numeric/formatting/run-count/width guards. Broader qualification passes **912 DOCX / 108 balloon / 33 Linux checks**. The source archive yields two expected SourceLink warnings on Linux. Exact local 0.1.5 package smoke `53baaa30bca64441bf534e44444b038f` contains the tested Windows DLL, SHA-256 `F667E48925F9B1BED5C475534453184C6197F989C93969373726FC99673FCAD8`.

All **2342 controls** preserve content and main text: 21 improve MAE and SSIM, 2321 retain PDF/raster/graphics bytes, 2258 match Office row breaks and none regress pixels. All 21 changed inputs match exact spaces/fonts; 39 within-word rows, 39 pitches and 21 transitions are independently audited. Maximum pitch/transition errors are 0.0754531249999362/0.14071875000004pt; row start-X/baseline gaps are 0.0849999999999795/0.234000000000037pt. Forty-two body runs and 21 nominal marks retain starts/advances within 0.127999999999986/0.104999999999961pt.

Sixteen fresh targets cover four font profiles, both preceding wrap masks, canonical spacing and preserved leading spaces. Six negative fixtures retain PDF/raster/graphics bytes. The prior fitting-second guard now matches Office's five-row layout; four independent font pairs establish the same behavior. All 195 earlier fitting/wrapped terminal paths retain PDF/raster/graphics bytes. Input/reference identities and exact extracted spaces/fonts remain mandatory. Cache retains all 34 PDFs and 545 failed gates.

84 control break mismatches remain. Preserved two-paragraph closing-only overflow, nonfitting closing runs, mixed intermediate paragraphs, extra runs, decoration and unqualified spacing retain fallback. Prior font-transition, positioning and composed pixel residuals remain explicit. The rejected maximum-final-descent model and independent Cambria/Courier New counterprobe remain preserved under L102. Evidence is under `artifacts/plan-revision-20261005/rv06-l106/`. Release preparation remains deferred.

## RV21-V61: full integration through L104 and smooth SVG transparency

Frozen runtime `d6baeaeb` passes clean Release and **2192 passed, 0 failed, 1 skipped**, including 910 DOCX checks. The optional private PPTX layout diagnostic lacks configured input/output. Exact local 0.1.5 package smoke `33ed2a95553846cfad55f380a8e45f8a` contains the full-suite DLL, SHA-256 `1A027671AE786F07CDE61FC80F63DEBC78B0B96AE8069E2928BCEE778D176C49`. This run covers L20 through L104/E4 and excludes L105 and later prototypes. Evidence is under `artifacts/plan-revision-20261005/milestone-d6baeaeb/`. Release preparation remains deferred.

## RV06-L107: preserved closing-only second comment row

Frozen runtime `2e59b39bdc25d5ef3bf99bf33eed2a4779129025` admits the preserved closing-only reflow branch in two-paragraph word-compatible all-markup balloons. The complete closing word moves to its own row; its authored breakable separator stays on the preceding first-body row in the closing face. That preceding row retains ordinary first-body pitch. Closing-only ascent/gap plus preceding-first descent defines the final transition; the continuation inset and nominal-size mark remain independent. Height and emission share the model. Prior fitting and wrapped two-/three-/four-paragraph paths retain bytes.

A production regression fails before change because fallback moves the preceding terminal word onto the closing row. Clean Release, 28 mixed checks and one paragraph-reader check pass, including 16 preceding-wrap/Unicode/line-gap/canonical-spacing combinations and earlier numeric/width/formatting guards. Broader qualification passes **913 DOCX / 109 balloon / 34 Linux checks**. The source archive yields two expected SourceLink warnings on Linux. Exact local 0.1.5 package smoke `944964cf8fc84c299eafecdb23e6e1e7` contains the tested Windows DLL, SHA-256 `F9A93CE6A6121693EB1A30FC8160D5A71FB5EC4DA26FE67360FAC90CFD71983C`.

All **2364 controls** preserve content and main text: 17 improve MAE and SSIM, 2347 retain PDF/raster/graphics bytes, 2275 match Office row breaks and none regress pixels. All 17 changed inputs match exact spaces/fonts; 45 within-word rows, 62 pitches and 17 paragraph transitions are independently audited. Maximum pitch/transition errors are 0.155781250000014/0.14071875000004pt; row start-X/baseline gaps are 0.0849999999999795/0.281000000000063pt. All 17 closing-only rows retain their preceding own-face separators within 0.07000000000005pt. Thirty-four body runs and 17 marks retain starts within 0.0849999999999795/0.144000000000005pt.

Twelve fresh targets cover four font profiles, both preceding wrap masks and canonical spacing; all closing separators are authored and preserved. Six negative fixtures retain PDF/raster/graphics bytes. The prior preserved-overflow guard matches Office's four rows; four independent font pairs establish closing-only placement, including five rows for Courier New/Cambria. All 216 earlier fitting/wrapped terminal paths retain PDF/raster/graphics bytes. Every input/reference pair retains its hash. Cache retains all 34 PDFs and 545 failed gates.

The initial Calibri/Georgia closing-only counterprobe measures 8.110pt versus the model's 7.954219pt, a 0.155781pt pitch residual. Keep this error and the complete qualification envelope explicit. The prior rejected maximum-final-descent model and independent Cambria/Courier New counterprobe remain preserved under L102. 89 control break mismatches remain; nonfitting or multispace closings, mixed intermediate paragraphs, extra runs, decoration, larger stories and unqualified spacing retain fallback. Prior font-transition, positioning and composed pixel errors remain explicit. Evidence is under `artifacts/plan-revision-20261005/rv06-l107/`. Release preparation remains deferred.

## RV21-V62: full integration through L105 and smooth SVG transparency

Frozen runtime `1bbff13a` passes clean Release and **2193 passed, 0 failed, 1 skipped**, including 911 DOCX checks. The optional private PPTX layout diagnostic lacks configured input/output. Exact local 0.1.5 package smoke `5e11e1994c1249c1b17c09307fd95d8d` contains the full-suite DLL, SHA-256 `E262F88D03022D2A3DF26E4DB395328C58646AD799416EDFD48CF184C8D93C5C`. This run covers L20 through L105/E4 and excludes L106 and later prototypes. Evidence is under `artifacts/plan-revision-20261005/milestone-1bbff13a/`. Release preparation remains deferred.

## RV06-L108: preserved closing-only fourth comment row

Frozen runtime `3771f48dd12aa91e5a77ed9f6195c3e662c345f1` admits preserved closing-only reflow in four-paragraph word-compatible all-markup balloons. The complete closing word moves to its own row; its authored separator stays on the preceding first-body row in the closing face. Ordinary terminal continuations retain first-body pitch. Closing-only ascent/gap plus preceding-first descent defines the final transition; the continuation inset, nominal-size mark and shared height/emission remain independent. Three preceding bodies retain their own rows, pitches and adjacent transitions.

A production regression fails before change because the fourth terminal reflow moves the preceding whole word onto the closing row. Clean Release, 30 mixed checks and one reader check pass, including 64 preceding-wrap-mask/Unicode/line-gap/canonical-spacing combinations. The existing 768 wrapped-fourth patterns now verify either Office-qualified final-row shape, with unchanged numeric, width and formatting guards. A separate 64-case boundary regression prevents the initial prototype from sending a previously rendered multiword closing run to fallback. Inputs outside the new single-closing-word scope retain their prior fourth-paragraph reflow. The initial broader test failure and superseded preliminary proof are archived. Broader qualification passes **915 DOCX / 111 balloon / 36 Linux checks**; Linux source archives yield two expected SourceLink warnings. Exact local 0.1.5 package smoke `45b07710bace486d85cd8a88662ed454` contains tested Windows DLL SHA-256 `33B6189666866ABA369D908CDFA813ADFDBDBA1743C34B897976C44E6227A4AA`.

All **2395 controls** preserve content and main text: 24 improve MAE and SSIM, 2371 retain PDF/raster/graphics bytes, 2299 match Office row breaks and none regress pixels. Twenty fresh targets cover four font profiles, four preceding wrap masks and canonical spacing; every closing separator is preserved. Four independent font pairs establish the initial Office defect and closing-only placement. Seven fresh negative fixtures and all 233 previously qualified terminal paths retain PDF/raster/graphics bytes. The 2364 inherited controls retain bytes. The additional multiword boundary guard retains accepted PDF bytes and the same six-row count as Office; the last two rows retain existing text/face-placement differences outside the new closing-only qualification. Input/reference pairs retain their hashes; shared output IDs are audited.

All 24 changed inputs match exact row spaces and fonts; 108 within-word rows, 141 pitches, 72 paragraph transitions, 48 body runs, 24 marks and 24 preceding separators are independently audited. Maximum pitch/transition errors are 0.165781250000005/0.210640624999909pt, row X/baseline gaps 0.0849999999999795/0.303999999999974pt, separator X gap 0.07000000000005pt and body/mark starts 0.0849999999999795/0.144000000000005pt. Initial Calibri/Georgia closing-only pitch remains 8.110pt versus 7.954219pt modeled, a 0.155781pt residual. The complete matrix reaches 8.120pt versus 7.954219pt on a closing-only row (0.165781pt error), and Georgia-to-Calibri transition from the third to fourth paragraph 7.900pt versus 8.110641pt modeled (0.210641pt error). These residuals remain explicit; the rejected maximum-final-descent model remains preserved. Cache retains all 34 PDFs and 545 failed gates.

96 control break mismatches remain. Fourth-paragraph closing inputs outside the admitted single-word shape retain their prior reflow or fallback. Nonfitting closing runs, trailing first-body separators, mixed intermediate paragraphs, extra runs, decoration, larger stories and unqualified spacing retain fallback. Prior font-transition, positioning and composed pixel errors remain explicit. Evidence is under `artifacts/plan-revision-20261005/rv06-l108/`. Release preparation remains deferred.

## RV21-V63: full integration through L106 and smooth SVG transparency

Frozen runtime `3d3038c7` passes clean Release and **2194 passed, 0 failed, 1 skipped**, including 912 DOCX checks. The optional private PPTX layout diagnostic lacks configured input/output. Exact local 0.1.5 package smoke `dc80d9d09e4f4df4adfe7a0a7a6b522a` contains full-suite DLL SHA-256 `3AFC93A2A12EA4A0C31CC4D3CB4B74C58D32031A3B065A064DE38533B4B2208D`. This run covers L20 through L106/E4 and excludes L107 and later prototypes. Evidence is under `artifacts/plan-revision-20261005/milestone-3d3038c7/`. Release preparation remains deferred.

## RV06-L109: fitting fifth comment body faces

Frozen runtime `7f961288891d8d62299d14226a7426c64072bb1d` preserves two fitting body faces in the fifth paragraph of word-compatible all-markup comment balloons. Four plain preceding bodies retain independent wrapping and continuation pitches. Each adjacent transition uses the preceding body's descent and the next body's ascent/gap; the terminal transition uses the larger ascent/gap of its two faces. Both fifth-body runs retain their own advances and the nominal-size paragraph mark. Height and emission share the same row metrics.

The production renderer and reader fail before this extension. The initial prototype also changes the existing wrapped-fifth PDF because an unused retained mark font shifts surviving font resource numbers; its text, geometry and raster remain unchanged. A separate failing regression reproduces that byte change with an independent unused mark face. Newly retained marks use separate subsets and resource names; ordinary body subsets and surviving fallback font numbers remain stable. Another failing regression exposes a same-face font alias entering the uniform-five-paragraph path. That path now excludes newly admitted mixed terminals; fitting qualification requires distinct prepared fonts. 32 combinations of wrap masks and mark font availability and an actual DOCX font-record alias retain prior PDF bytes. An unloaded mark also caused an unused built-in fallback declaration; its failing regression is fixed by preparing built-in fallback faces from ordinary runs, since newly admitted marks require their own usable embedded faces. Final-mode reading continues to omit these marks. Original and corrected evidence remain archived.

Clean Release, 32 mixed checks and one reader check pass. The new fitting-fifth regression covers 512 preceding-wrap-mask/Unicode/font-order/gap/margin/spacing combinations, all four transitions, body/mark advances, nominal size, height and numeric/formatting/run-count/width guards. Broader qualification passes **917 DOCX / 113 balloon / 38 Linux checks**; Linux source archives yield two expected SourceLink warnings. Exact local 0.1.5 package smoke `54216fe1cae04e7a912222c70fed39fd` contains tested Windows DLL SHA-256 `02BF5ABA8118F77BF0A5CF3FD114FD0F24D214D56C55474A76DF4EDE73B352F1`.

All **2474 controls** preserve content and main text: 72 improve MAE and SSIM, 2402 retain PDF/raster/graphics bytes, 2371 match Office row breaks and none regress pixels. Sixty-eight fresh targets cover four font profiles, all sixteen preceding wrap masks and canonical spacing; four independent Office font pairs establish the initial behavior. Seven fresh negative fixtures, the legacy wrapped-fifth guard, all 257 qualified terminal paths and all 2395 inherited controls retain PDF/raster/graphics bytes. Shared IDs and every actual candidate PDF hash are audited. The wrapped-fifth legacy guard retains ten candidate rows versus eighteen Office rows; that mismatch remains outside this fitting qualification.

All 72 changed inputs match exact row text, spaces and font faces. 330 within-word rows, 380 continuation pitches, 288 transitions, 144 final-body runs and 72 marks are independently audited. Maximum pitch/transition errors are 0.105453125000023/0.165781250000005pt, row X/baseline gaps 0.0849999999999795/0.314000000000078pt and final body/mark starts 0.127999999999986/0.0950000000000273pt. Initial independent probes retain pitch/transition errors 0.105453/0.165781pt and baseline error 0.284pt. Previous closing-only Calibri/Georgia pitch error 0.165781pt, Georgia/Calibri transition error 0.210641pt, rejected maximum-final-descent model and composed positioning residuals remain explicit. All 34 cached PDFs retain bytes with 545 failed gates.

103 control break mismatches remain. Wrapped fifth terminals, mixed intermediate paragraphs, extra runs, decoration, larger stories and unqualified spacing retain fallback. Evidence is under `artifacts/plan-revision-20261005/rv06-l109/`. Release preparation remains deferred.

## RV21-V64: full integration through L107 and smooth SVG transparency

Frozen runtime `2e59b39b` passes clean Release and **2195 passed, 0 failed, 1 skipped**, including 913 DOCX checks. The optional private PPTX layout diagnostic lacks configured input/output. Exact local 0.1.5 package smoke `c7109caf1b274c0a9fca16d57f4c149e` contains full-suite DLL SHA-256 `CC1969E81DEFB8CC679501855B0B22E2CDFD5AD995713D632865F85F00095F4D`. This run covers L20 through L107/E4 and excludes L108 and later prototypes. Evidence is under `artifacts/plan-revision-20261005/milestone-2e59b39b/`. Release preparation remains deferred.
## RV21-V65: full integration through L108 and smooth SVG transparency

Frozen runtime `3771f48d` passes clean Release and **2197 passed, 0 failed, 1 skipped**, including 915 DOCX checks. The optional private PPTX layout diagnostic lacks configured input/output. Exact local 0.1.5 package smoke `6b6e3854347e4c158e22597a9fe94fa7` contains full-suite DLL SHA-256 `6FD063F3BB6F4FE395DCB8D0B5F34DF1BFEFA8F1E5792CF0E31AED0972B45412`. This run covers L20 through L108/E4 and excludes L109 and later prototypes. Evidence is under `artifacts/plan-revision-20261005/milestone-3771f48d/`. Release preparation remains deferred.

## RV06-L110: wrapped fifth comment body before a fitting closing face

Frozen runtime `bb3f65a2ced228ff70bf2644239addb542b0d136` preserves the source faces and advances of a wrapped fifth comment body. Four plain preceding bodies keep their own rows, continuation pitches and all four adjacent transitions. Ordinary fifth rows use the first terminal face; the last mixed row uses its preceding first-face descent and the larger emitted-body ascent/gap. The terminal mark stays at nominal size and height uses the same metrics as emission. Office-proven whole-word reflow without an authored closing separator is retained; overflow with an authored closing separator remains outside this slice.

The production regression fails before the extension. An initial prototype also reflows a closing-only fifth case without qualification; a separate failing regression pins the previous fallback for that boundary. The corrected model passes 1536 prefix-length/wrap-mask/Unicode/font-order/gap/margin/spacing combinations at lengths 80, 88 and 104; a length-96 case independently checks the rejected overflowing final row. Thirty-three mixed checks and one reader check preserve prior fitting, alias, loaded/unloaded mark, separator, formatting, spacing and numeric guards. Broader qualification passes **918 DOCX / 114 balloon / 39 Linux checks** with clean Release; Linux source archives yield two expected SourceLink warnings. Exact local 0.1.5 package smoke `e3dd7b83ba234f7186b197cacbefb2a3` contains tested Windows DLL SHA-256 `25EF0B3773DAC48B996FB4EC331F62FBF0B562CC61CACB11FE08EC54278E123D`.

All **2555 controls** preserve content and main text: 74 improve MAE and SSIM, 2481 retain PDF bytes, 2481 retain raster bytes, 2481 retain graphics operations, 2445 match Office breaks and none regress pixels. Sixty-eight fresh targets cover four profiles, all sixteen preceding-wrap masks and canonical spacing; four independent font pairs and two legacy inputs establish the behavior. The legacy ten-row candidates now match Office's eighteen/seventeen rows with exact spaces and fonts. Nine fresh guards and all 329 qualified terminal paths retain PDF/raster/graphics bytes; all 2472 other inherited controls retain those identities. Shared IDs and every actual candidate PDF hash are checked.

All 74 changed inputs match exact row text, spaces and font faces. 476 within-word rows, 549 pitches, 296 transitions, 148 final-body runs and 74 marks are independently audited. Maximum pitch/transition errors are 0.105453125000023/0.165781250000005pt, row X/baseline gaps 0.0849999999999795/0.314000000000078pt and final body/mark starts 0.173000000000002/0.108999999999924pt. Initial six-input errors 0.105453/0.165781pt, row X/Y 0.085/0.297pt and body/mark starts 0.173/0.109pt remain explicit. The rejected maximum-terminal-ascent prediction for the first fifth-row transition is archived; the actual model uses that row's first face. Previous closing-only pitch/transition errors 0.165781/0.210641pt and composed drift remain explicit. All 34 cached PDFs retain bytes with 545 failed gates.

Two initial negative probes did not express their intended authored spaces: unmarked boundary XML whitespace was insignificant, and a first preservation attempt produced an invalid reserved-namespace prefix. Both inputs are retained in ignored evidence; distinct corrected fixtures explicitly use `xml:space="preserve"` and pass XML round-trip validation. Their hashes and Office references remain separate. The nine qualified guards cover authored closing-only overflow, a nonfitting closing word, a preserved trailing first-body separator, same-face aliases, extra runs, decoration, mixed intermediate bodies, larger stories and unqualified spacing. 110 control break mismatches remain. Evidence is under `artifacts/plan-revision-20261005/rv06-l110/`. Release preparation remains deferred.

## RV21-V66: full integration through L109 and smooth SVG transparency

Frozen runtime `7f961288` passes clean Release and **2199 passed, 0 failed, 1 skipped**, including 917 DOCX checks. The optional private PPTX layout diagnostic lacks configured input/output. Exact local 0.1.5 package smoke `aa7b1e9975a5476e9592f4d7d25bf639` contains full-suite DLL SHA-256 `D54D0B27883652BF024DD823603FA8E769BE2EC73139F483047EE1469E702C88`. This run covers L20 through L109/E4 and excludes L110 and later prototypes. Evidence is under `artifacts/plan-revision-20261005/milestone-7f961288/`. Release preparation remains deferred.

## RV06-L111: preserved fifth closing word on its own row

Frozen runtime `0b4d9517800b11750043dbc67a14e07beca64f3f` follows accepted PNG-S1. Four plain preceding comment bodies retain their own rows and continuation pitches. When a wrapped two-face fifth body has a preserved single leading separator and a complete fitting one-word closing face, the first-body words and separator stay in place and only the closing word moves to its own row. Every adjacent transition and the closing row use the emitted faces; the mark stays at nominal size and height uses the same metrics as emission. L110 whole-word reflow without an authored separator remains qualified separately.

The new production regression fails before the extension, then passes 128 wrap-mask/Unicode/gap/canonical-spacing combinations and seven negative boundaries per combination (896 checks). The former length-96 authored-overflow guard becomes an explicit positive: its last first-body word stays and its closing word moves. All 34 mixed checks and one reader check pass; **919 DOCX / 115 balloon / 40 Linux checks** pass with clean Release. Linux source archives yield two expected SourceLink warnings. Exact local 0.1.5 package smoke `728a2abea7ae46e4a463c7b2fa574225` contains tested Windows DLL SHA-256 `34D3C3EE810D09091049F5EF5B3296A9F77C99AE2268E702936744F6FBA409D7`.

All **2588 controls** preserve content and main text: 25 improve MAE and SSIM, 2563 retain PDF bytes, 2563 retain raster bytes, 2563 retain graphics operations, 2470 match Office breaks and none regress pixels. Twenty fresh targets cover four profiles, preceding-wrap masks 0/1/6/15 and canonical spacing; four independent font pairs and the original legacy authored-closing input provide separate evidence. The legacy ten-row candidate now matches seventeen Office rows. Nine fresh guards and all 2554 other inherited controls retain PDF/raster/graphics identities. The preliminary 403-input terminal-path check verified fresh PDF bytes; the complete inherited-control qualification supplies the separate raster and graphics proof. Shared IDs and every actual candidate PDF hash are checked.

All 25 changed inputs match exact row text, spaces and font faces. 164 within-word rows, 209 pitches, 100 transitions, fifty final-body runs and 25 marks are independently audited. Maximum pitch/transition errors are 0.155781250000014/0.165781250000005pt, row X/baseline gaps 0.0849999999999795/0.283999999999992pt, final body/mark starts 0.0849999999999795/0.144000000000005pt and preserved-separator starts 0.07000000000005pt. The immutable initial four-pair proof records 47 pitches/sixteen transitions and errors 0.155781/0.165781pt; its row X/Y 0.085/0.284pt and body/mark starts 0.085/0.144pt remain explicit. Previous font-transition, positioning and composed residuals remain. All 34 cached PDFs retain bytes with 545 failed gates.

Multiword and doubled-separator authored closings, nonfitting closing words, trailing first-body separators, same-face aliases, mixed intermediate bodies, extra runs, decoration, larger stories and unqualified spacing retain guarded fallback. Fresh fixtures explicitly preserve authored XML boundaries and validate XML round trips. 118 control break mismatches remain. Evidence is under `artifacts/plan-revision-20261005/rv06-l111/`. Release preparation remains deferred.

## RV21-V67: full integration through L110 and smooth SVG transparency

Frozen runtime `bb3f65a2` passes clean Release and **2200 passed, 0 failed, 1 skipped**, including 918 DOCX checks. The optional private PPTX layout diagnostic lacks configured input/output. Exact local 0.1.5 package smoke `68a74f2c84c749f3b96694cc38270357` contains full-suite DLL SHA-256 `02CA905E6A21706EF9ECB7E390375A8228F4524BAAF25DF0AA559DCAD1E94C9C`. This run covers L20 through L110/E4 and excludes L111 and PNG-S1. Evidence is under `artifacts/plan-revision-20261005/milestone-bb3f65a2/`. Release preparation remains deferred.

## RV06-L112: fitting two-word fifth closing run on its own row

Frozen runtime `60ff93970007cc50a0cb9c4ead5a6d4a98fb4b8e` follows accepted L111 and PNG-S1. Four plain preceding comment bodies retain their own rows and continuation pitches. When a wrapped two-face fifth body has a preserved single leading separator and a complete fitting two-word closing face with one internal space whose first word cannot fit on the preceding row, the first-body words and separator stay in place and only the closing run moves to its own row. Every adjacent transition and the closing row use the emitted faces; the mark stays at nominal size and height uses the same metrics as emission. L110 whole-word reflow without an authored separator remains qualified separately.

The production regression fails before the extension; the final fixture and unchanged final test assembly also fail against the exact original frozen L111 library in a separate runner. The corrected source passes 128 wrap-mask/Unicode/gap/canonical-spacing combinations and ten negative boundaries per combination (1280 checks). The prior one-word regression uses a three-word fallback boundary; its former two-word fallback is now separately qualified as a positive. All 35 mixed checks and one reader check pass; **920 DOCX / 116 balloon / 41 Linux checks** pass with clean Release. Linux source archives yield two expected SourceLink warnings. Exact local 0.1.5 package smoke `6fc193af9cea447c84b3118e924a5549` contains tested Windows DLL SHA-256 `1D453C53DA23A838D162A789C6324B46D5FE35917B91BDE859A95F9E7D949D5C`.

All **2623 controls** preserve content and main text: 25 improve MAE and SSIM, 2598 retain PDF bytes, 2598 retain raster bytes, 2598 retain graphics operations, 2495 match Office breaks and none regress pixels. Twenty fresh targets cover four profiles, preceding-wrap masks 0/1/6/15 and canonical spacing; four independent font pairs and the original legacy two-word-closing input provide separate evidence. The original legacy multiword-closing candidate changes from three to seven exact Office rows. Eleven fresh guards and all 2587 other inherited controls retain PDF/raster/graphics identities. The preliminary 428-input terminal-path check verified fresh PDF bytes; the complete inherited-control qualification supplies the separate raster and graphics proof. Shared IDs and every actual candidate PDF hash are checked.

All 25 changed inputs match exact row text, spaces and font faces. 156 within-word rows, 199 pitches, 100 transitions, fifty final-body runs and 25 marks are independently audited. Maximum pitch/transition errors are 0.155781250000014/0.165781250000005pt, row X/baseline gaps 0.0849999999999795/0.283999999999992pt, final body/mark starts 0.0849999999999795/0.0930000000000177pt and preserved-separator starts 0.07000000000005pt. The immutable initial four-pair proof records 47 pitches/sixteen transitions and errors 0.155781/0.165781pt; its row X/Y 0.085/0.284pt and body/mark starts 0.085/0.093pt remain explicit. Previous font-transition, positioning and composed residuals remain. All 34 cached PDFs retain bytes with 545 failed gates.

A separate Office probe keeps a fitting first closing word on the preceding row and moves its second word; the initial whole-run prototype fails an independent regression at that boundary. The corrected source keeps generic fallback there (three rows versus seven Office rows), so no parity is claimed for that case. Three-word and doubled-separator authored closings, nonfitting closing words, trailing first-body separators, same-face aliases, mixed intermediate bodies, extra runs, decoration, larger stories and unqualified spacing retain guarded fallback. Fresh fixtures explicitly preserve authored XML boundaries and validate XML round trips. An initial trailing closing-space unit guard is archived as invalid for this route because preparation trims that boundary before rendering; trailing first-body-space and doubled-separator guards remain meaningful. 128 control break mismatches remain. Evidence is under `artifacts/plan-revision-20261005/rv06-l112/`. Release preparation remains deferred.


## RV06-L113: partially fitting two-word fifth closing

Frozen runtime `36d3999c19d1dc188ef3fbdcf193498edc86c133` follows accepted L112 and PNG-S1. Four plain preceding comment bodies retain their own rows and pitches. When a wrapped two-face fifth body has a preserved single leading separator and a complete fitting two-word closing with one internal space, its fitting first closing word and internal separator stay beside the preceding first-body text; the second word moves onto its own closing-face row. The mixed row uses first-body descent and the larger emitted ascent/gap; the final row uses first-body descent plus closing ascent/gap. Height shares those metrics, and the mark remains nominal. Earlier complete one-/two-word own-row closings retain bytes.

The production regression fails before this extension. The final fixture in an unchanged final test assembly separately fails for the expected row-layout reason against the exact original frozen L112 DLL. The corrected source passes 128 wrap-mask/Unicode/gap/canonical-spacing combinations and ten negative boundaries per combination (1280 checks). All 36 mixed checks and the reader check pass. **921 DOCX / 117 balloon / 42 Linux checks** pass with clean Release. Linux source archives yield two expected SourceLink warnings. Exact local 0.1.5 package smoke `ea6d26d83894477d86857882c15232de` contains tested Windows DLL SHA-256 `5FF05944FF2E798AD841DA42AE13B90B1BE43812DAF09216F3EFC9113A64B5BA`.

All **2664 controls** preserve content and main text: 25 improve MAE and SSIM, 2639 retain PDF bytes, 2639 retain raster bytes, 2639 retain graphics operations, 2526 match Office breaks and none regress pixels. Twenty original fresh targets cover four profiles, preceding-wrap masks 0/1/6/15 and canonical spacing; four initial independent pairs are also retained. Six Courier New/Cambria inputs already match Office through the earlier complete own-row behavior and retain PDF/raster/graphics bytes. Six additional inputs with first-body prefix length 104 provide partial-closing evidence for that pair, covering its independent pair, four masks and canonical spacing. The original first-closing-word-fits input supplies the remaining legacy positive. The original candidate changes from three to seven exact Office rows. Eleven fresh guards and all 2622 other inherited controls retain PDF/raster/graphics bytes. The preliminary 453-input terminal-path proof establishes PDF identity separately; the complete inherited-control run supplies raster and graphics proof. Shared IDs and every actual candidate PDF hash are audited.

All 25 changed inputs match exact row text, spaces and font faces. 156 within-word rows, 199 pitches, 100 transitions, seventy-five terminal-body parts and 25 marks are independently audited. Maximum pitch/transition errors are 0.155781250000014/0.165781250000005pt, row X/baseline gaps 0.0849999999999795/0.283999999999992pt, body/mark starts 0.0849999999999795/0.0419999999999163pt, leading separator 0.07000000000005pt, first closing glyph 0.067260000000033pt and internal separator 0.0913679999999886pt. Immutable preliminary evidence for the original input plus two independent pairs records 25 pitches/twelve transitions, errors 0.079063/0.165781pt, nineteen within-word rows within 0.085/0.260pt, and nine body starts/three marks. It precedes the complete matrix and does not substitute for it. Previous font-transition, positioning and composed residuals remain. All 34 cached PDFs retain bytes with 545 failed gates.

Three-word/doubled-separator/nonfitting complete authored closings, trailing first-body separators, same-face aliases, mixed intermediate bodies, extra runs, decoration, larger stories and unqualified spacing retain fallback. A complete closing wider than a continuation row remains unqualified even when its second word fits. Fixtures preserve authored XML spaces and validate round trips. 138 control break mismatches remain. Evidence is under `artifacts/plan-revision-20261005/rv06-l113/`. Release preparation remains deferred.


## RV06-L114: partially fitting three-word fifth closing

Frozen runtime `4601a1d8efb473466b17df62a6b618133768b914` follows accepted L113 and PNG-S1. Four plain preceding comment bodies retain their own rows and pitches. When a wrapped two-face fifth body has a preserved single leading separator and a complete fitting three-word closing with single internal spaces, its fitting first closing word and internal separator stay beside the preceding first-body text; the second word cannot fit beside the preceding body, so the last two words move together onto their own closing-face row. The mixed row uses first-body descent and the larger emitted ascent/gap; the final row uses first-body descent plus closing ascent/gap. Height shares those metrics, and the mark remains nominal. Earlier complete one-/two-word own-row closings retain bytes.

The production regression fails before this extension. The final fixture in an unchanged final test assembly separately fails for the expected row-layout reason against the exact original frozen L113 DLL. The corrected source passes 128 wrap-mask/Unicode/gap/canonical-spacing combinations and fourteen negative boundaries per combination (1792 checks). All 37 mixed checks and the reader check pass. **922 DOCX / 118 balloon / 43 Linux checks** pass with clean Release. Linux source archives yield two expected SourceLink warnings. Exact local 0.1.5 package smoke `a680eed32fd24804965bcde5fc918408` contains tested Windows DLL SHA-256 `D56BC18157C956BC1CE2C08A85CFF143B15CDA0E5F5BF58A14416A87FF152C61`.

All **2702 controls** preserve content and main text: 25 improve MAE and SSIM, 2677 retain PDF bytes, 2677 retain raster bytes, 2677 retain graphics operations, 2551 match Office breaks and none regress pixels. Twenty fresh targets cover four profiles, preceding-wrap masks 0/1/6/15 and canonical spacing; four independent font pairs and the original three-word-closing input supply separate evidence. The first closing word fits beside the preceding body, the second cannot, and the last two words fit together on the closing row. The original candidate changes from three to seven exact Office rows. Fourteen fresh guards and all 2663 other inherited controls retain PDF/raster/graphics bytes. The preliminary 478-input terminal-path proof establishes PDF identity separately; the complete inherited-control run supplies raster and graphics proof. Shared IDs and every actual candidate PDF hash are audited.

All 25 changed inputs match exact row text, spaces and font faces. 156 within-word rows, 199 pitches, 100 transitions, seventy-five terminal-body parts and 25 marks are independently audited. Maximum pitch/transition errors are 0.155781250000014/0.165781250000005pt, row X/baseline gaps 0.0849999999999795/0.283999999999992pt, body/mark starts 0.0849999999999795/0.0809999999999036pt, leading separator 0.07000000000005pt, first closing glyph 0.067260000000033pt and internal separator 0.0913679999999886pt. The immutable original-input proof records two pitches/four transitions within 0.012375/0.165781pt, exact rows/spaces/faces and MAE/SSIM 0.610969→0.156392 / 0.731643→0.973178. The separate initial four-pair proof records 47 pitches/sixteen transitions within 0.155781/0.165781pt and twelve body parts/four marks. These precede the complete matrix and do not substitute for it. Previous font-transition, positioning and composed residuals remain. All 34 cached PDFs retain bytes with 545 failed gates.

Four-word/doubled-separator/nonfitting complete authored closings, a fitting second word, trailing first-body separators, same-face aliases, mixed intermediate bodies, extra runs, decoration, larger stories and unqualified spacing retain fallback. A complete closing wider than a continuation row remains unqualified even when its last two words fit. The initial three-word boundary failure is archived: earlier unit guards that become admitted now use four-word closings, and the new regression separately guards nonfitting first words and fitting second words. The original prior guard inputs remain in the complete inherited proof. Fixtures preserve authored XML spaces and validate round trips. 151 control break mismatches remain. Evidence is under `artifacts/plan-revision-20261005/rv06-l114/`. Release preparation remains deferred.

## RV21-V68: full integration through L110 and smooth SVG transparency and PNG-S1

Frozen runtime `aabe6d32` passes clean Release and **2205 passed, 0 failed, 1 skipped**, including 918 DOCX checks. The optional private PPTX layout diagnostic lacks configured input/output. Exact local 0.1.5 package smoke `0b74508f9fcc4974b17109608df37bd1` contains full-suite DLL SHA-256 `99BBB455827E860FD1AAB866841A955FA357FF2A1A5A5584C47E2801DABC653B`. This run covers L20 through L110/E4 and PNG-S1 and excludes L111 and later prototypes. Evidence is under `artifacts/plan-revision-20261005/milestone-png-s1/`. Release preparation remains deferred.

## RV06-L115: whole and partial three-word fifth closings

Frozen runtime `68fb5864be5e8defec512216f52dfb4a9b68d91a` follows accepted L114 and PNG-S1. Four plain preceding comment bodies keep independent rows, pitches and transitions. A wrapped two-face fifth body with exactly three closing words and single internal spaces now handles both remaining fitting-closing layouts: the complete closing moves onto its own row when its first word cannot fit beside the preceding body; when two words fit, those words and their internal separator stay on the preceding row and the third word moves below. Closing ascent/gap, preceding first-body descent and the nominal mark share height and emission. Earlier partial three-word and complete/partial one-/two-word paths retain bytes.

Unchanged final regression assemblies fail against exact frozen L114 for both expected row-layout reasons. Corrected source passes 256 wrap-mask/Unicode/gap/canonical-spacing combinations across two regressions and their negative boundaries, all 39 mixed checks and the reader check. **924 DOCX / 120 balloon / 45 Linux checks** pass with clean Windows Release; Linux source archives have two expected SourceLink warnings. Exact local 0.1.5 package smoke `0f9b04bbef864fd6ab40e29c6fb2768d` contains tested Windows DLL SHA-256 `6A0A9FF082AE8D7E87975FC911D9A42CD86A6F0420B723FED7B1B7A2F291E1DC`.

All **2761 controls** preserve comment content and main text: 51 improve MAE and SSIM, 2710 retain PDF bytes, 2710 retain raster bytes and 2710 retain graphics operations. Forty fresh targets span four font profiles, preceding wrap masks 0/1/6/15 and canonical spacing; eight independent font-pair inputs and three inherited inputs supply distinct evidence. The three formerly guarded inputs are `two-word-fifth-three-word-closing-guard`, `three-word-fifth-first-word-does-not-fit-guard` and `three-word-fifth-first-two-words-fit-guard`. Eleven fresh guards and 2699 other inherited controls retain bytes. All 59 fresh Office baselines retain Word 16/Letter and unchanged-default-printer evidence.

Every inherited input is freshly converted. For actual PDF byte identities, raster/graphics evidence transfers from qualified frozen L114 with pinned parent PDF, raster, graphics and reference hashes; three changed inherited PDFs and all fresh inputs receive fresh comparisons. All candidate PDF hashes and shared output IDs are verified. The separate preliminary 503-input terminal proof establishes PDF identity. Original whole-only and combined preliminary proofs remain separate immutable archives.

All 51 changed inputs match exact row text, spaces and font faces. 325 within-word rows, 414 pitches, 204 transitions, 127 terminal body parts and 51 marks are audited. Maximum pitch/transition errors are 0.341916503906305/0.565417480468772pt; row X/baseline gaps 0.444999999999993/0.383000000000038pt; body/mark starts 0.466999999999985/3.99899999999991pt. Whole-closing preserved-separator starts differ by at most 0.07000000000005pt. Partial closing leading-separator, first visible glyph and final internal-separator starts differ by at most 0.466999999999985/0.344399999999951/0.440500000000043pt. One preserved original Office export of `two-fitting-words-fifth-profile2-mask15` uses different glyph widths and origins from a same-input independent repeat; its nominal-mark start differs by 3.99899999999991pt, versus 0.0190000000000055pt for the repeat. The repeat matches the independent paired reference; the primary 51-control proof retains the original and its maxima. Previous font-transition, positioning and composed residuals remain explicit. All 34 cached PDFs retain bytes with 545 existing failed gates.

Four-word and doubled-separator closings, complete closings wider than continuation width, trailing first-body separators, mixed intermediate bodies, extra runs, decoration, larger stories and unqualified spacing retain fallback. 159 control break mismatches remain. Evidence is under `artifacts/plan-revision-20261005/rv06-l115/`. Version remains 0.1.5; release preparation is deferred. Full milestone integration is recorded separately.

## RV21-V69: combined-batch delivery checkpoint

The isolated full integration at frozen runtime `68fb5864be5e8defec512216f52dfb4a9b68d91a` completes through L115/E4 and PNG-S1: **2211 passed, zero failed, one skipped**, including **924 DOCX checks**. The only skip is the unconfigured optional private PPTX diagnostic. Release builds are clean. Exact local 0.1.5 package smoke `d7fb25be34e247998153572ff883ea05` contains the full-suite DLL SHA-256 `64AD48D48572CCA21223B20E8146C7C2DCB86102D7E57C69BF10899D28434220`. Source and package hashes are checked again after the run; later prototypes are excluded.

L115's separate qualification retains its own exact DLL/package hashes, 924 DOCX/120 balloon/45 Linux checks, 2761 controls, 51 MAE/SSIM improvements, 2710 PDF/raster/graphics identities and all content/main identities. All 34 cached PDFs remain identical with 545 existing failed gates. The original Office export variability, positioning/font-transition limits and unsupported scopes remain recorded. This checkpoint completes the integrated implementation and validation work; version remains 0.1.5 and release preparation is deferred.

Future affected qualification uses `tools/RunDocxAffectedTests.ps1` to execute overlapping DOCX/balloon methods once for one frozen assembly. The actual catalogue currently has 924 DOCX methods covering 118 of 120 balloon methods; two missing-font checks complete the balloon group. Source reports and both library/test-assembly hashes are pinned, incomplete reports or ambiguous selectors are rejected, and catalogue listing does not claim tests passed. The helper was checked with the PowerShell parser and the actual frozen catalogue; the completed L115 qualification preserves its original independently executed reports. Full integration runs remain milestone checks. Public cached fixtures guide the next coherent wrapped-table/range-anchor investigation; no representative user documents are required.

Evidence is under `artifacts/plan-revision-20261005/milestone-l115/`. The delivery-workflow commit adds no renderer changes and does not require another full-suite run.

## RV06-L116: combined wrapped-table geometry and comment ends

Automatic single-paragraph, top-aligned review-table cells now retain the printed first-baseline inset when they wrap; all physical lines receive the shared correction, preserving continuation pitch. The shared Word-compatible non-body anchor helper selects the actual complete final source fragment for end markers between runs; this behavior is independently qualified on table cells. Wrapped static/placed-story paragraphs remain outside the table-cell fidelity claim. Explicit end offsets keep their existing path; unknown metadata and incomplete/hidden terminal source runs retain fallback. Ordinary body anchors and preserve-document geometry are unchanged. Declared heights, split rows, vertical merges, nested tables, drawings, multiple paragraphs, non-top alignment and explicit line heights retain the existing table-projection guards.

Frozen runtime `cb3995e1d89fffdade27cb8abf958d8e7f221209` and DLL `1AB89485DF482272B2142A5CB8E0CF32DE3FD6232C90DDC740EC3E6B85528526` pass **927 DOCX / 120 balloon checks**, executing **929 distinct methods**, and **48 Linux regressions**. Windows Release and inspector builds are clean; the git-free Linux archive has the two expected SourceLink warnings. Both production regressions in the unchanged frozen test assembly fail against the exact qualified L115 DLL for the inset/final-line reasons. Exact local 0.1.5 package smoke `47515d7972b64e73b8ff65daf777adc7` contains the tested DLL. Release preparation remains deferred.

**36 regular independent Office controls** cover four fonts, 11/24-point cells, one-/two-/three-/five-line bodies, solid/no borders, preceding text and document-start tables: **28 MAE/SSIM improvements and eight PDF identities**, unchanged visible main content and page counts. Fresh frozen conversions of the initial 24 retain the measured preliminary PDF bytes; pinned pixel/graphics evidence is reused with its original DLL provenance. All **37 preserve-layout PDFs** retain bytes, including the separate long legacy case. **2761 inherited controls** are freshly converted with exact PDF identities and pinned qualified raster/graphics provenance. The prior 28 table controls retain 27 PDF identities; the formerly guarded long wrapped case is measured separately below.

All **34 cached cases** are refreshed. **33 PDFs retain bytes**, and the public table-heavy case improves **16→13 failed gates**, reducing the unchanged gate set from **545→542** with zero increases. Its six actual cell clips improve from 3.768pt maximum error to **0.616pt**; 144-DPI whole-page MAE improves 2.094→1.574 and SSIM 0.690→0.756. Original broad table-like graphics gates remain intact. `tools/SummarizeDocxTableCellGeometry.ps1` adds source-dimension/translation mapping of rectangular cell clips, explicit unresolved/missing counts, hashes and unclassified clips. Five checks using the actual public case and missing/ambiguous/no-source mutations verify that partial matching does not claim complete coverage.

Regular-control residual maxima remain **2.619pt cell bounds**, **3.140pt connector X** and **2.554pt connector Y**; Courier New retains the largest cell-bound residual. The long legacy wrapped control independently shows a first-page MAE improvement **6.090→2.602** and SSIM **0.604→0.875**, while its pre-existing pagination mismatch remains: **Office two pages, candidate one page before and after**. This case is separate from the 36 regular page-matching controls and motivates the next pagination investigation. Remaining cached gates and earlier balloon-positioning/composed limits remain. Full V69 integration excludes L116; affected qualification replaces another immediate full-suite run.

Evidence is under `artifacts/plan-revision-20261005/rv06-l116/`; preliminary, frozen and inherited hashes remain distinct.

## RV06-L117: printed page capacity and table-row fragments

Plain automatic review-table rows now compare their already printed line heights against the printed body-frame capacity. Page-boundary fragments receive the existing qualified table-baseline projection, retaining line pitch and subsequent-row flow. Fragment membership uses the same printed font inset as projection; testing nominal baselines against printed row heights could otherwise discard a terminal line when the font inset exceeded the printed pitch. The complete-range comment anchor path now requires a typed table-cell story. Wrapped header paragraphs retain their established fallback; this guard does not claim Office parity for static stories.

Frozen runtime `9f23553806b9ffac1c34dbe117f1c9ebc2f8c967` and DLL `AEBFC0B9906B3C63F154FF8EB1B376693D261B3ED9F52FE0834ED30453ECD895` pass **932 DOCX /120 balloon checks**, executing **934 distinct methods**, and **53 Linux regressions**, with zero failures or skips. Windows Release and inspector builds are clean; the git-free Linux archive has two expected SourceLink warnings. All five unchanged new regressions fail against the exact accepted L116 DLL. Exact local 0.1.5 package smoke `13299f4fbe814f6eb23f774c4abe2f88` contains the qualified DLL. Release preparation remains deferred.

**24 independent Word 16 / Letter controls** cover Calibri 11/24-point long rows, solid/no borders, fitting and splitting rows, single-line continuations, and a following row moving to page two. Every control matches main text content, total page count, and first-cell physical-line allocation. **12 changed PDFs improve first-page MAE and SSIM; 12 PDFs retain bytes**. All **24 preserve-layout PDFs** retain bytes. First-cell baseline error is at most **0.220pt**. Source-backed geometry maps **58/58 cell fragments**, with no unresolved cells and at most **0.864pt** cell-bound error. Physical lines are grouped by page/baseline; PDF text-operation counts are not line counts because Office can split a word into several operations.

The original long wrapped control now matches **Office two pages**, distributing its **22 first-cell lines as 18/4**, with the second cell on page two. Previously the candidate kept the entire table on one page. First-page 144-DPI MAE improves **2.602→0.440**, and SSIM **0.875→0.985**. The formerly missing second page has MAE **0.180**, SSIM **0.981**, unchanged visible content, and all three source cell fragments mapped.

Fresh frozen conversion preserves **2761 inherited control PDFs**, all **36 previously qualified regular table PDFs and 36 preserve PDFs**, and **27/28 earlier table controls**; the changed earlier control is the qualified long case above. All **34 cached public PDFs** retain bytes and their existing **542 failed gates**. Reused visual evidence retains input/reference/PDF/report/raster/graphics hashes. The cached cases remain partial apart from the previously passing case.

The pagination change admits plain automatic single-paragraph top-aligned cells with implicit or disabled widow control. Explicit keep/widow policies, declared heights, cantSplit/header rows, vertical merges, nested tables, drawings, multiple paragraphs, non-top alignment and explicit line heights retain the prior capacity path. Header/footer/note content and multiple columns also retain prior capacity. Office proof is bounded to the controls above; earlier Courier New geometry, balloon positioning, composed drift and unsupported scopes remain. Full V69 integration through L115 excludes L116/L117; complete affected qualification avoids another immediate full-suite run.

Evidence is under `artifacts/plan-revision-20261005/rv06-l117/`, including `evidence.json`, `office-pagination-summary.json`, `office-cell-geometry-summary.json`, the separate legacy proof, inherited/regular-table identities, and unchanged cached gates. Frozen runtime, package and later documentation integration hashes remain distinct.

## RV06-L118: visible comment stories and review canvas

Word-compatible all-markup now prints comment balloons and range markers only for source paragraphs in the main story, including its table cells. Independent Word 16 references omit header, footer, footnote and endnote comments, including comments in their tables; mixed documents retain the body comment. Textbox comment visibility follows the previously measured Office policy. Source reference identity prevents structurally identical paragraph values in another story from becoming visible main-story anchors. Other geometry modes retain their comment policy.

Hidden-story-only comments no longer trigger fitted print scale or its text-anchor shift. The gray lane requires surviving comment or formatting balloons; formatting balloons in other stories retain their established lane eligibility even though the body print-scale trigger is narrower. An orphan comment part or reference cannot independently sustain the review lane. Four former tests expected static/note markers or an orphan lane; their expectations now follow the independent Office evidence. Body paragraph/table markers remain covered, and preserve-layout balloons remain covered separately.

Frozen runtime `79114a71aa74de163ad68ea9809cdbf8d930818f`, DLL `EC2123AF592C312103AE8952838E143CB3AE8ECCE592FA9E25FEC461447F7E00` and test assembly `9D7DAD817217A544F6882D882D99A8BA8A8F0DF6D6128565FABA0C50A03BEEBF` pass **935 DOCX /121 balloon checks**, **937 distinct methods**, and **56 Linux regressions**, with zero failures or skips. Release and source-inspector builds are clean; the git-free Linux archive retains two expected SourceLink warnings. Three unchanged visibility regressions fail against the exact accepted L117 DLL. The unchanged foreign-formatting lane regression also fails against the rejected intermediate prototype. Exact local 0.1.5 package smoke `d65400b003554be5ade17545d028d0eb` contains the qualified DLL. Release preparation remains deferred.

All **14 independent Word 16 / Letter controls** match comment titles, lane presence, normalized source-story content and page counts. They cover isolated header/footer/footnote/endnote comments, header/note table cells and mixed body/story anchors. **13 changed PDFs improve MAE and SSIM**, the body-only control retains PDF bytes, and all **14 preserve-layout PDFs** retain bytes. These controls qualify visibility and canvas policy. Mixed/body-only cases retain an existing **8.797pt** first-body-baseline error; unscaled generic paragraphs can still wrap an extra short line, and endnote/note-table flow plus palette/connector residuals remain. No complete typography or pagination parity is claimed.

Fresh frozen conversion preserves **2761 inherited control PDFs** and **150 table-control PDFs**: 36 regular controls in both geometry modes, 24 pagination controls in both modes, 28 prior table controls and the legacy long control in both modes. A local batch converter calls the public API with one shared resolver, an existing supported contract, and verifies every generated PDF against its accepted CLI hash. The **2911 fresh conversions complete in 64.004 seconds**. Input/reference/report/raster/graphics hashes remain linked; the library gains no dependency.

The **34 public cached cases improve 542→525 failed gates**, with **32 accepted PDF identities** and no increases. Static-story comments improve **23→9**; text-box comments improve **3→0** after the empty lane is removed. Two cases now pass all gates; 32 remain partial. Pixel/graphics comparisons from the earlier prototype are reused only when the fresh final frozen PDF and input/reference hashes match. The earlier prototype's **930 passed /4 failed** report and its formatting-lane defect remain archived separately. Current frozen evidence is under `artifacts/plan-revision-20261005/rv06-l118-final/`; the earlier prototype remains under `rv06-l118/`.

Remaining frame-width and pagination gaps, dense painting/text-spacing errors, composed drift and earlier fallback boundaries remain explicit. The cached dense reference uses inline revisions, matching the declared renderer view; historical balloon-view artifacts do not establish a defect in that mode. Full V69 through L115 excludes L116-L118. Frozen runtime, package and later documentation integration identities remain distinct.

## RV06-L119: authored frame without printed balloons

Word-compatible all-markup restores the authored body frame when no comment or formatting balloons survive. Hidden header/footer/note comments, unused comment parts and inline insertions cannot narrow the text into an unoccupied review margin. The fitted text shift also uses effective printed-balloon visibility, respecting document settings that hide comments. Mixed main-story comments and visible foreign formatting balloons retain their existing frames and lanes. Other geometry modes retain their behavior.

Frozen runtime `6d473fd30c28487cf7ebd06c45bd91547704b545`, DLL `EC47D04A4521287840A0DE98844ECD681DF4F92909CBEBD78B4BB924778AFFF5` and test assembly `53294324983A2B8018579DC4C055E1832391B9724E56558EBB8BFC0C535ED9C4` pass **939 DOCX /121 balloon checks**, **941 distinct Windows methods** and **60 Linux regressions**, with zero failures/skips. Windows Release and source-inspector builds are clean; the git-free Linux archive has two expected SourceLink warnings. Three unchanged regressions fail against the exact accepted L118 DLL. Mixed-frame and formatting-lane positive guards pass against that parent; the unchanged hidden-comment placement guard also fails against the hash-pinned prior frame-only prototype. Exact local 0.1.5 package smoke `5bc873afbe124e4aa8d83874ccc3f657` contains the qualified DLL. Release preparation remains deferred.

Eight independent Word 16 / Letter controls match total pages, physical body-line allocation and word content across 12/24-point text, two authored margins, fitting/overflowing documents, a mixed body-comment guard, unused comment parts and inline insertion. **Seven changed PDFs improve first-page MAE and SSIM; the mixed PDF retains bytes**, and all eight preserve-layout PDFs retain bytes. Three 24-point controls previously used two candidate pages against one Office page; they now fit on one page. The long two-page control matches Office's 22/2 physical-line allocation. The separate fourteen earlier story controls retain visibility/lane/content/page matches, with **eight pixel improvements, six PDF identities and fourteen preserve identities**.

Every one of the **34 cached public cases** receives fresh source-layout inspection, structural comparisons and raster comparisons from the frozen Release library. Failed gates improve **525→520**, with **33 PDF identities** and no increases; static-story comments improve **9→4**. Two cases pass all gates and 32 remain partial. The ignored gate adapter pins existing validation tools and records successful inspections with no graphics as explicit empty operation sets; the inspector otherwise omits that file when the set is empty. Runtime source and dependencies are unchanged by the adapter.

Fresh public-API batch conversion preserves all **2761 inherited plus 150 table PDFs**, completing 2911 conversions in **58.626 seconds**. Accepted input/reference/PDF/report/raster/graphics hashes remain linked. No representative private documents were needed.

Pixel parity remains incomplete. The mixed 24-point control retains MAE **16.119**, SSIM **0.083** and about **11.36pt** first-body-baseline error; earlier generic mixed/body controls retain **8.797pt**. The long hidden-only control retains first-page MAE **3.960** and the inline-insertion control **5.785**. Dense painting/text spacing, note/table flow, connector/palette differences and prior fallback boundaries remain. The separate L119 milestone covers all 2227 registered methods with 2226 passes, zero failures and one optional private PPTX skip. It combines 941 affected methods and 1286 disjoint original frozen delegates after complete-catalogue validation. Sixteen harness fixture-path failures are retained and pass through the original runner with its normal fixture base; this is reconciled catalogue coverage, not one unfiltered invocation. Evidence is under `artifacts/plan-revision-20261005/rv06-l119/`; frozen runtime, package and subsequent documentation integration identities remain distinct.

## RV06-L120: automatic note marks and forward navigation

Referenced normal footnote/endnote stories retain their explicit automatic marks, using the filtered main-story labels and original run styles. Visible automatic markers in main paragraphs and table cells link to the first actually placed matching note. The source run and offset must match a measurable emitted marker; missing source/target content creates no link. Footnote and endnote identities remain distinct, and a continued note targets its first rendered page. At unit print scale, a complete plain document-endnote block may fit between the body and bottom footnotes. Overflowing, scaled and complex stories retain continuation placement.

Frozen runtime `33c10bcae4be73592d9cbebed7fccfc88a81f49f`, DLL `64FA3A50D4B51684365AA39D1C14C070466ED7572AD4E85B657EA89B85305928` and test assembly `98C0E3E7748F097B82C63D304BB2990A20989C6C7345C06804C9F752DEC7B5CE` pass **944 DOCX /121 balloon checks**, **946 distinct Windows methods** and **65 Linux regressions**, with zero failures/skips. Release/library-inspector builds are clean; the git-free Linux archive has two expected SourceLink warnings. Four unchanged frozen checks fail against the exact accepted L119 DLL, while the missing-source/target guard passes against that parent. The exact local 0.1.5 package contains the qualified DLL; release preparation stays deferred.

Eight independent Word 16 / Letter controls verify **11 actual links** with matching source/destination pages, including table markers, a second-page endnote, shared XML runs, reordered note IDs and separate footnote/endnote identities. All eight match page counts; the mixed-note document now matches Word's one page instead of two. Seven match the decoded word multiset. Nine of eleven source rectangles lie within 1pt of Office. Four controls improve both MAE and SSIM. The fourteen prior story controls preserve titles, review lanes, content, page counts and every raster metric. Their PDFs, including preserve-layout output, intentionally change for restored links.

All **34 cached public cases** receive fresh frozen conversion, layout inspection, structural and raster comparison. Failed gates improve **520→518**, with **33 PDF identities** and no increases. Static-story comments improve **4→2**, closing both missing-link gates; two raster gates remain. Fresh public-API conversion retains all **2911 inherited/table PDFs** in **54.850 seconds**, with accepted evidence hashes linked.

Remaining errors are explicit. Endnote content sits about **8.03pt too low**, and its destination about **7.3pt too low**. Pure-endnote MAE worsens by **0.00166**; mixed-note first-page MAE worsens **0.674→0.766** while fixing the page-count mismatch. A settings-only numbering control retains the accepted parent's **105/106 versus Office 1/2** mismatch; new inside-note marks follow those body labels consistently. Its two annotation bounds reach **7.582pt** error, so it qualifies ID binding rather than numbering-setting parity. Custom marks, back-links, section/restart numbering, separator calibration and complete Office parity are outside the new claim. L119's complete-catalogue milestone excludes L120. Evidence is under `artifacts/plan-revision-20261005/rv06-l120/`; runtime, package and documentation integration identities remain separate.

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








