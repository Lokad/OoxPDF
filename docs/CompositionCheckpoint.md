# Composed-document reliability checkpoint

The fixed [portfolio](../tests/Lokad.OoxPdf.Tests/Cases/Composition/portfolio.json)
contains four synthetic DOCX documents and two existing composite decks. The
report, letter and specification exercise mixed content; the technical memo and
second deck are held out from fitting. The DOCX documents have different lengths,
and the memo moves the table after its supporting discussion. This is a bounded
reliability checkpoint; version 0.1.5 release preparation remains deferred.

Each DOCX source has six equations across its body, table and header, an inline
image, numbering, bookmark, external link, footer, footnote and comment. Both
default Final/preserved layout and AllMarkup/Word-compatible output are measured.
Expected prose and source counts are pinned before layout corrections. Source
counts do not establish that every item painted or remained visible.

The initial baseline retains all expected prose but ignores equations in all
eight DOCX mode outcomes. All eight fail existing DOCX numerical visual gates;
the report and held-out memo also differ from Word in AllMarkup page count. Both
decks pass their inherited numerical gates. Passing a raster gate alone does not
establish complete fidelity. Missing supported content and significant pagination,
overlap or clipping are checkpoint blockers. Typography and cosmetic residuals
remain explicit and thresholds are unchanged.

The first correction removes bookmarks from math exclusion policy: source
annotations have no measured or painted extent. Transformation regressions cover
standalone and inline body equations, a cell and direct header/footer paragraphs,
in both modes, including unchanged-format prose run splitting. Nineteen capability and neighbor methods pass in one invocation; the new
regressions fail on the previous library. Seventeen frozen assets transfer with
four identical integration smoke PDFs and no root rebuild. Independent bookmarked
Word controls still fail foreground/SSIM gates; no complete visual pass is claimed.
Equation-only bookmark navigation remains unresolved; this correction does not
complete the checkpoint.

Reusable commands:

- `NewCompositionPortfolio.ps1` authors fixtures; do not regenerate pinned inputs
  during qualification. ZIP metadata may differ on regeneration.
- `MeasureCompositionPortfolio.ps1` converts using an explicit existing CLI and
  inspector, verifies identities and reports content findings separately by mode.
- `CompareCompositionPortfolio.ps1` uses independent Office PDFs and existing
  numerical thresholds. Cache references for unchanged source/mode identities.
- `RunDocxAffectedTests.ps1 -Capability docx-inline-annotations` runs a tracked
  capability plus neighbors as one exact method set. `--test-set file.json` also
  accepts an explicit JSON array directly in the console test runner.

Cold CLI wall time includes process startup. Calling-thread allocation and stage
measurements belong to the separate allocation probe; no speedup is claimed.
Retained local baseline evidence is under `artifacts/composition-c1/`.


The next component retains ordered text, image, equation, tab and break items for
math-bearing source paragraphs. Inline equation offsets and visible-prose policy
consume this view. Original object identities and image positions within a run
are preserved; zero-width annotations stay metadata. Fields, references, textboxes
and revisions retain the legacy source path. Replacing a legacy content collection
invalidates its borrowed view; placement-only changes preserve it. This source foundation preserves the baseline output identities.

Twenty-seven selected methods pass in one invocation; six portable source/bookmark
methods also pass on the exact artifact in Ubuntu. Seventeen assets and four
integration PDFs transfer without root rebuild. All ten portfolio PDFs
retain their baseline bytes, so their Office comparisons are reused. Mixed-document
warm allocation medians rise by approximately 11–25KB for the DOCX sources; deck
allocation medians are unchanged. Three-iteration timings are noisy and establish
no speedup. These measurements precede any repeated-layout optimization.


The bounded mixed-content path prepares each equation by its original identity.
Ordered fragments retain measured text segments, advances, baseline insets and
occupied heights. Body, cell and direct header/footer containers use the same
wrapping and placement functions, including inline pictures and multiple fitting
inline equations. Placement translates measured geometry instead of measuring
text again. Original text-run affinity survives hidden-run filtering; character
spacing and pair kerning survive tokenization. Long words reuse the existing
bounded Unicode wrapping rules. Keep-next reserves the complete mixed frame.

Eight separate probes place two equations and one picture in the body, a cell,
a header or a footer, with and without review scaling. They supplement the fixed
portfolio and do not replace its held-out documents. Source-to-page mapping moves
text, equations and pictures together; one existing vertical-frame owner applies
to all three. Independent Office comparisons retain their numerical failures.
Content preservation does not establish Office geometry or unclipped painting.

Display combinations, tabs, fields, revisions, list labels and rich source owners
retain explicit fallbacks. The older single-equation marker path remains until a
separate measured migration preserves its Office comparisons. This is a bounded
migration, not completion of the composed-document checkpoint.


The mixed-content Release artifact passes 59 capability and neighboring methods
in one invocation with no skips. Eight DOCX portfolio outcomes now retain all
expected prose and have no ignored-equation diagnostic. The development report's
native checks verify all six equations and an actual image drawing in both modes;
sixteen placement/mode checks verify two equations and one image drawing each.
All 22 placement PDFs match their independently compared prototypes exactly, so
the Office comparisons are reused. Both deck PDFs also retain their baseline bytes.
All eight DOCX portfolio visual comparisons still fail; the report and memo retain
their AllMarkup page-count mismatch. These residuals block checkpoint completion.


The `rendering-portable` console group adds three small actual-rendering checks:
DOCX mixed content in both modes and a PPTX composite chart. All use the same
pinned DejaVu Sans font, license and provenance under `Cases/PortableFonts`.
The dedicated DOCX fixture authors normal mathematical text because this font
lacks the portfolio's default mathematical italic alphabet. Ubuntu CI runs this
group alongside its existing portable groups. Windows retains the full suite.


The next paragraph-frame correction compares already scaled advances with the
printed body capacity for single-column review output. Headers, footers and
other document features no longer select a larger paragraph capacity. Three
independent exact-height controls preserve two Word pages in both modes with no
static story, a header or a footer. The composed report also keeps Word's break
between review sections 8 and 9. Default report output remains byte-identical.
Visual gates remain unmet; matching pagination alone does not establish fidelity.

Table fragment capacity and its existing keep/widow fallbacks remain separately
qualified. A broader prototype exposed disagreement between measured membership
and painted continuation lines when combining printed capacity with the older
fragment baseline path. That prototype is retained as rejected development
evidence; table baseline/membership ownership needs its own bounded correction.
Following body paragraphs still use paragraph capacity after a table fallback.


The paragraph-frame artifact passes 46 selected methods in one invocation and
three portable rendering methods on Ubuntu. All ten portfolio mode outcomes have
matching page totals, expected prose and no ignored-equation diagnostic. The
held-out memo recovers its second AllMarkup page. Four default DOCX PDFs and both
decks remain byte-identical. Four integration PDFs match after transferring the
17 qualified assets without a root build or repeated suite. All eight DOCX raster
gates still fail. Fourteen of sixteen comparisons reuse exact candidate/reference
PDF and threshold identities; only two need new raster comparisons.


A subsequent content-preservation regression reproduces missing table words on
the delivered artifact: explicit widow control paints 117 of 120 synthetic words.
Independent 200-word public controls paint 198 in review output on the parent.
Automatic rows now publish their printed first-baseline geometry before fragment
membership, using the same natural-cell rule for admission and painting. The
later table translation does not scale that inset again. Keep/widow page-capacity
fallbacks, authored heights and rich-cell fallbacks remain explicit. Focused
coverage checks all words and baselines inside fragment bounds across lengths and
widow settings; emitted text alone does not establish unclipped visual parity.
The independent controls retain numerical visual and pagination failures.


The table-fragment artifact passes 43 exact methods in one invocation; the three
portable renderers and the word-preservation regression also pass in Ubuntu.
All ten fixed portfolio PDFs retain their accepted bytes. Four independent table
control PDFs match their compared prototype identities. Seventeen assets and four
root integration PDFs transfer without rebuild or repeated suite. Native controls
retain pagination/visual failures and three unresolved Office clip matches per
review source; candidate clips are classified. This fixes content loss within the
bounded natural-cell path and leaves broader frame/fidelity work explicit.


Natural automatic text-only body-table rows now receive the same printed capacity as body
paragraphs. The containing frame supplies this height once; fragment projection
scales baseline geometry separately. Headers, footers and unrelated document
stories no longer choose a larger table frame. Authored row heights, repeated
headers, explicit row/paragraph keeps and rich cells retain staged fallbacks.

Widow-protected splitting can choose one earlier whole-line boundary and recheck
every cell, retaining at least two continuation lines. This bounded retry honors
cancellation; invalid boundaries retain the whole-row fallback. The shared line
measurements determine both admission and painting. Two-cell checks ensure an
unprotected first cell cannot bypass the next cell's widow rule or lose content.
Pinned long-table controls also cover plain headers and footers in both modes.


The held-out memo rejected an initial broader frame candidate: it moved a
mathematical row to the next page, whereas Word keeps that row with its heading.
Text-metric eligibility now explicitly excludes equations as well as pictures
and textboxes; mixed-object rows retain their established frame and painting.
This semantic ownership boundary applies to every source and is not keyed to the
memo, a font, a size or a fitted threshold. Wider object-height ownership remains
unqualified. Held-out comparisons remain acceptance checks, not fitting inputs.


A transformation regression also retains all prose when unsupported mathematics
falls back. Font-baseline eligibility stays separate from frame-capacity
eligibility: an ignored object cannot switch off the published prose baseline.
The regression catches 117/120 words in an intermediate candidate and passes after
restoring the shared line-metric owner. Neither rejected candidate is delivered.


The accepted natural text-table frame passes 54 selected methods in one invocation
and five portable checks under Ubuntu. All eight long-table mode outcomes preserve
200 words and match Word's four-page totals; other numerical gates remain unmet.
All ten fixed portfolio PDFs retain accepted bytes, including the held-out memo.
Eighteen independent comparisons reuse PDF/reference/threshold identities. Four
root smoke PDFs match after transferring 17 assets without rebuild or suite replay.
The two rejected intermediate revisions remain retained as development evidence.


Equation-only bookmark navigation now consumes measured paragraph-frame origins
and tops across body, cell, header and footer stories. Those paragraphs do not
manufacture text segments for navigation. The same relative frame origin survives
alignment, translation and static-story scaling; existing text destinations keep
their qualified viewport rule. Mixed-inline/multiple-equation source annotation
binding remains a separate migration.

A compound public fixture has four links and four equation-only story targets,
with an explicit review comment. Word exports all four destinations in both modes;
the parent drops them. Candidate text and graphics inspection reports match the
parent byte-for-byte in both modes. Default destination differences are below
one point; review-mode vertical differences reach about 3.4pt, reflecting retained
layout residuals. No fitted destination correction or visual-threshold change is
introduced. A normal-math variant exercises all four targets with the pinned
portable font, independently of installed Windows fonts.


Navigation qualification passes 33 selected methods in one invocation and all
four actual portable renderers under Ubuntu. Each compound mode now exports four
links, matching Word's target count. The corresponding PNGs, text operations and
graphics operations retain parent bytes. All ten fixed portfolio PDFs are exact;
17 assets and four integration PDFs transfer without rebuild or suite replay.
Coordinate residuals and failed visual gates remain explicit; mixed-inline source
annotation binding is still open.
