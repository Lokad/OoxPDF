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








