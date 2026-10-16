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


Standalone body equations now appear as `Equation` in layout inspection, with
source block and column ownership, structural box bounds, explicit baseline and
occupied line height. Previously they appeared as `Unknown` with zero geometry.
These measurements describe layout frames, not glyph ink. The inspection-only
change passes 14 selected methods in one invocation and the pinned-font regression
on Ubuntu, where both DOCX modes are exercised. Four integration PDFs retain their
previous bytes after transferring 17 qualified assets without a root rebuild or
suite replay. No visual-fidelity improvement is claimed for this telemetry fix.


Bookmarks at formulas in the ordered mixed-content path now retain an inline
source boundary. Navigation waits for that formula's actual placed line, including
lines with no prose segments. The shared paragraph frame supplies the origin in
body, cell and static stories; existing run-relative fallback remains available.
A new control combines two formulas in each of those four stories, an inline image
and an explicit body line break. Word exports eight destinations in both modes;
the parent exports none. The candidate restores all eight without changing the
control's raster, text operations or graphics operations.

This is navigation recovery, not equation-geometry parity. Word centers adjacent
formulas as a display group with larger fraction typography; the candidate still
uses left-aligned inline math. Target differences reach 218.5pt
horizontally in Final and 165.9pt in review mode; vertical differences reach about
8pt. Both numerical visual gates still fail. Investigate the common math-only
paragraph layout before subpoint refinements; do not fit navigation offsets to
Word while painting formulas elsewhere.

The accepted batch passes 60 selected methods in one invocation and six portable
rendering/inspection checks on Ubuntu. All ten fixed portfolio PDFs are exact,
so their independent comparisons reuse verified identities. Seventeen assets and
four integration PDFs transfer without a root rebuild or suite replay. A new
metadata test initially omitted the separate static-story source enumerator; its
correction changes no production code. C1 remains open.


Fitting math-only rows with multiple adjacent formulas now share display-style
measurement and center the complete group in body, cell and static containers.
The same alignment helper consumes the measured group width in all three paths;
source equations remain immutable. Mixed prose, pictures and authored hard-break
rows retain their existing inline path in this component. Navigation retains
container context at line start and uses the placed position of a later formula.

Independent controls place two, three and two formulas in left-, center- and
right-aligned paragraphs. Word centers each math-only group. Final foreground
recall improves from 0.276 to 0.755; review recall improves from 0.988 to 0.994.
Other numerical visual gates still fail. In the compound navigation control,
maximum horizontal destination difference falls from 218.5pt to 0.54pt; remaining
vertical differences reach about 8pt. The isolated formula after an image/break
still uses inline typography and left alignment; its shared row ownership is the
next rendering problem, without fitting a paragraph-specific correction.

Qualification passes 71 selected methods in one invocation and seven portable
rendering/inspection checks on Ubuntu. All ten portfolio PDFs remain exact, with
14 independent comparisons reused by source/reference/PDF/threshold identity.
Four smoke PDFs match after transferring 17 assets without root rebuild or suite
replay. C1 remains incomplete and all eight DOCX portfolio gates remain failed.


Authored breaks now delimit shared source rows for mathematical style. A formula
on a row without another visible inline owner uses display measurement, including
when another row of the paragraph contains an image. Prepared boxes retain that
style, so wrapping and container placement consume the same decision. Soft
wrapping stays unchanged. Invisible prose cannot change row ownership; its
metamorphic geometry regression fails on the parent and passes after aligning the
source visibility rule. Paragraphs without prepared math bypass row classification.

Independent image/break variants cover body, cell, header and footer, separately
in both modes. All eight destinations survive in each outcome. Final target
differences are at most 0.52pt horizontally and 0.93pt vertically; review residuals
reach 0.54pt and 3.39pt. All eight numerical visual gates still fail. Foreground
recall improves in every control, but review cell MAE increases about 0.004;
retain that nonmonotonic residual rather than claiming universal metric gains.

The frozen batch passes 73 selected methods in one invocation and eight portable
rendering/inspection checks on Ubuntu, including both modes in all four containers.
All ten fixed portfolio PDFs remain exact. Eighteen independent comparisons reuse
verified identities. Seventeen assets and four integration PDFs transfer without
root rebuild or suite replay. C1 remains open; legacy single-inline marker paths
and broader composition/geometry concerns remain to be addressed in bounded slices.


Single-equation prose paragraphs now keep one qualified layout owner whether or
not their unchanged-format text carries a hyperlink. Native literal/fraction/
script regressions and a pinned-font compound-letter regression demonstrate that
annotations no longer change measurement or placement. Body and cell bookmarks
consume actual placed math geometry, including the body's separate equation item,
with source/block identity retained instead of resolving to following prose.
Independent two-story targets differ from Word by at most 1.45pt X and 0.94pt Y.

Blanket migration to the shared whole-object wrapper was evaluated and rejected.
Word breaks the overfull literal at an operator; whole-formula relocation adds a
line and reduces recall from 0.375 to 0.329. Both old relocation fallback assertions
failed. That prototype is retained separately and is not delivered. Operator
breaking remains a distinct capability required before retiring the legacy marker
path; the accepted annotation fix preserves both fallbacks.

Qualification passes 82 selected methods in one invocation and nine portable
rendering/inspection checks on Ubuntu. The eight DOCX portfolio outputs retain
page totals and show small recall/MAE improvements; both decks retain bytes.
All eight DOCX numerical gates still fail. Ten final PDFs match the compared
prototype, so independent comparisons reuse their verified identities. Seventeen
assets and four smoke PDFs transfer without root rebuild or suite replay; the
smoke includes exact parent output from the overfull fallback control.


`tools/AuditCompositionContent.py` screens complete public source prose and simple
literal/fraction expressions against the measured candidate and cached Office
PDFs. It pins frozen inputs/library, PDF/comparison identities and inspector
outputs. Referenced headers/footers repeat per reference page; only referenced
notes and mode-appropriate comments contribute. Complex source features remain
explicitly unresolved. This is a token-presence screen, not reading-order, image,
ink, clipping or geometry proof; numerical visual gates remain independent.

The first audit reuses frozen `1845342c` without conversions or library builds.
All ten source-prose/simple-math screens pass. Four review-mode Office token
screens flag the generated comment caption: Word prints `Commented [1]:`, while
the candidate substitutes `Commented [Public reviewer1]:` when initials are absent.
This is a confirmed caption difference, not missing body prose. It remains open.
The same audit functions identify exactly two missing `AAA` tokens in retained
198/200-word table output and no source-prose deficit after the accepted correction.


Word-compatible comment captions now use explicit initials and encounter-order
numbers, without substituting author names. A five-case Office control confirms
`[1]`, `[2]`, `[RV3]`, `[4]`, `[QA5]` for absent/empty/explicit initials, empty
authors and nonsequential source IDs. Ordinary author titles, palette and thread
rules stay on their existing paths. The portable regression fails on the parent.

Twenty-three selected title/wrapping/thread/palette/rendering methods pass in one
invocation; ten portable methods pass on Ubuntu. The four Final DOCX PDFs and both
decks retain bytes. Four review outputs improve MAE on the affected pages, while
all eight DOCX visual gates still fail. The complete source/Office token and simple
math screen now passes all ten outcomes: 4,234 source prose tokens, no token/math
deficits or unresolved source features. This does not establish image/ink/clipping
or reading-order fidelity. Seventeen assets/four smoke PDFs transfer without root
rebuild or suite replay. C1 remains open for geometry and composed-content limits.


Autofit column measurement now consumes prepared equation advances. Its minimum
protects each whole box; its maximum reuses the ordered inline measurement used
for wrapping, including prose, pictures, formulas and authored breaks. Fixed,
preferred-width and span fallbacks retain their existing paths. Cancellation flows
through the new measurement work; no font/size/document-specific rule is added.

An independent four-row control exposed a 10.7pt equation column and four omitted
fractions. All four now paint; added-prose and duplicated-formula transformations
also preserve their four/eight expressions and increase preferred width. Native
foreground recall improves from 0.275 to 0.851; other visual gates still fail. An
unreferenced template header continues to report an ignored equation, separately
from the visible body loss fixed here. Fixed-grid spacing controls show ordinary
full-row heights are close to Word; the held-out page-boundary discrepancy remains.

Thirty-two selected methods pass in one invocation and eleven portable checks
pass on Ubuntu. All ten portfolio PDFs remain exact and their content audit has
no deficits. Twelve independent comparisons reuse verified identities. Seventeen
assets/four smoke PDFs transfer without root rebuild or suite replay. The next
rendering investigation is independently controlled trailing-spacing admission
at a page boundary; the held-out memo remains an acceptance check.

Natural mathematical last rows now distinguish measured content from trailing
paragraph after-spacing at a page boundary. If all content fits atomically,
placement may discard only the overflowing unpainted tail. Existing measurements
are reused. Authored heights, keeps, merges, positioned content and shifted runs
retain their fallbacks; no font, size or whole-document qualification is added.

Independent 0/6/24pt spacing controls keep the equation on page one in Word and
the candidate. CantSplit and insufficient-space controls keep it on page two in
both modes. The held-out memo's Final equation row now stays on page one. Its
four source cell clips match, with no unresolved cells and a maximum bounds delta
of 1.333pt; the raster also shows the complete fraction inside the cell.

Forty-one selected methods pass in one invocation; twelve pinned-font rendering
methods pass on Ubuntu. Nine portfolio PDFs retain bytes, and fifteen independent
comparisons reuse verified identities. A page-local token screen detects the old
memo displacement despite matching page totals and passes all ten new outcomes.
It supplements global prose/simple-math presence checks, without proving ink,
image, clipping or reading-order fidelity. All eight DOCX visual gates still fail;
both decks pass. Memo Final page-two MAE improves from 1.172 to 0.429 and recall
from 0.541 to 0.840. Seventeen assets/four exact smoke PDFs transfer without root
rebuild or suite replay. C1 remains open for significant geometry and composition
limits.

Ordinary list numbering no longer excludes supported composed equations. Label
measurement and text-state planning are shared with ordinary paragraph layout;
the first measured line owns its label segments, while containers place those
segments in their own frames. Authored indentation/suffix rules are reused.
Numbered rows retain inline style and existing unsupported rich-content fallbacks.

A numbering-only transformation previously removed four equations from three
mixed paragraphs in both modes, while preserving prose. The candidate retains
four equations and three labels across body/cell/header/footer controls. Native
unchanged-format run splitting and pinned-font source transformations also retain
content. The body control's recall improves 0.141→0.671 and MAE 1.464→0.821;
all eight independent numbered-story visual gates still fail. Raster review shows
the labels, fractions, literals and illustration together.

Fifty-two selected methods pass in one invocation, including numbering emission,
indentation/suffix/continuation and table-math neighbors. Thirteen portable methods
pass on Ubuntu. All ten portfolio PDFs retain bytes; eighteen independent Office
comparisons reuse verified identities. Global/page-local prose and simple-math
screens pass. Seventeen assets/four exact smoke PDFs transfer without rebuilding
or repeating the suite. The portfolio DOCX gates remain failed; decks remain passed.

Geometry triage now preserves PDF page ownership in both nearest-position and
indexed text comparisons. Page-moved content previously passed both modes. A
second adversarial control exposed rounding-edge splits of collinear runs; anchored
0.1pt baseline bands now keep those runs in X order without chaining distinct
lines. Twenty-two adversarial cases pass. Optional structured reports retain input,
tool/helper hashes, matching options, tolerances, pages, text and measured deltas.
Existing raster thresholds are unchanged.

Cached portfolio inspection required no conversion or library build. All 334
Final merged pairs match normalized text and page, with maximum start differences
of 0.877pt X and 0.981pt Y. Review outcomes retain 28/33 unresolved operation pairs
per document, including differing mathematical operation segmentation; these are
not inferred content omissions. Exact review pairs reach 3.152pt Y. Global and
page-local content screens remain separate evidence; neither proves complete ink
or clipping fidelity. Sub-point Final refinements remain lower priority than
confirmed composition loss and the overfull operator-breaking contract.

Top-level inline literals now expose binary/relation break opportunities from
their existing measured glyph order. The ordered paragraph path wraps those
pieces, drops inter-piece glue at a line boundary and preserves source equation
identity. Structured boxes, display rows, nested and unary-only literals retain
their qualified fallbacks. No font/size/document-shape rule is added.

Word 16 keeps `α+` on the first line and `β=γ` on the next for default, before,
after and repeat settings in the independent inline control. This observed inline
behavior differs from the general default described in [Microsoft's brkBin
documentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.math.breakbinary?view=openxml-2.10.0);
the implementation claim is limited to the validated inline scope. All five
glyphs now paint in a two-line source paragraph, including with bookmarks and run
splitting. The portable regression fails on the previous library.

Forty selected methods pass in one invocation; fourteen portable methods pass on
Ubuntu. The overfull control's recall improves 0.375→0.480 and MAE 0.290→0.280.
All eight setting/mode controls still fail foreground/MAE/SSIM gates. Its body
auto line height remains 16.968pt versus Word's 14.64pt; font/measurement ownership
needs separate evidence before changing it. Eight fitting literal/fraction/script
outcomes retain their comparison results. All ten portfolio rasters retain bytes
at the existing 144dpi qualification, while eight DOCX PDF byte streams change
with mathematical boxing. Content/page-local screens pass, but DOCX visual gates
remain failed. Seventeen assets/four exact smoke PDFs transfer without root build
or suite replay.

An ink-only inline-frame prototype worsened review portfolio MAE by up to 1.904
and recall by 0.0434. That measurement rule was rejected and reverted before
delivery; the accepted operator change retains the prior frames. The rejected
artifact and comparisons remain available as negative evidence.

Broken-literal content is now independently exercised in a cell, header and
footer, in both modes. All five glyphs and following prose survive; native
bookmarks/run splitting and pinned-font placement regressions pass. The production
library and its source tree remain identical to the accepted operator artifact:
only two new methods run in one invocation, plus the new Ubuntu method. The
earlier forty-method/fourteen-Ubuntu qualification and portfolio outputs are
reused by identity. Seventeen assets/four exact smoke PDFs transfer without root
rebuild, suite replay or repeated portfolio conversions.

The six new Office controls still fail their original visual gates. Final-mode
second-line differences are 2.334pt in the cell and 2.452pt in the header; the
footer's first line is 7.322pt high. Reader fallback currently supplies 278/240
when no line spacing resolves, while these sources use a single-line pitch in
Word. That discrepancy is confirmed; paragraph-mark metric ownership is a
hypothesis to test, not an accepted cause. Preserve older independent unstyled
controls while isolating document defaults, style inheritance and mark formatting.

Document paragraph defaults now retain their source ownership even when `pPr` is
empty. With no resolved line token, a declared document default supplies single
spacing; an absent owner retains implicit Normal 278/240. A uniquely declared
default paragraph style binds through the existing cascade, so its explicit line
setting overrides the fallback. Authored style IDs and cascade provenance remain
separate from inherited formatting. Two identical fallback constants and a
redundant qualification helper are removed. No font/size/document-shape guard is
added or indiscriminately removed.

The seven-source ownership matrix separates absent, run-only, empty/nonempty
paragraph defaults and explicit Normal inheritance. Its fourteen mode outcomes
preserve explicit Normal PDF bytes while affected recall improves 0.222→0.773 and
MAE 9.173→3.981. The source controls still fail their original visual gates. Plain
cell/header/footer controls improve recall from 0.377/0.286/0.353 to
0.836/0.850/0.831; MAE and changed-pixel gates pass, with foreground/SSIM failures
remaining. These controls disprove a global fallback replacement and do not
establish paragraph-mark metrics as the cause.

This shared-reader milestone qualifies 1,494 DOCX/neighbor methods once in one
invocation and sixteen portable methods on Ubuntu. An earlier artifact passed
1,493/1,494 but changed authored style identity; that defect was corrected before
delivery and the failed artifact/report retained. All ten portfolio PDFs and their
comparisons remain exact; full/global/page-local content screens pass. Seventeen
assets/four exact smoke PDFs transfer without a root build or suite replay. The
broader selection is a deliberate shared-contract milestone, not the selection
policy for routine local changes. Existing limits, cancellation, safe managed code
and visual thresholds remain intact.

PDF text comparisons now transform advances, character spacing and font height
consistently with effective coordinates. Page-local exact decoded matching reserves
content before geometric fallback, and decoded comparison is case-sensitive.
Twenty-eight adversarial controls pass, including scaled/sheared/rotated text,
adjacent columns and case changes. Existing tolerances and raster gates remain.
Cached inspection pairs all 676 normalized portfolio operations with no conversion
or library build: Final maxima are 0.877pt X / 0.981pt Y; review maxima are
0.484pt X / 3.166pt Y. These supersede the earlier unresolved segmentation counts.
Raw whitespace differences remain explicit: review output loses an authored space
after a comment reference in all four documents. This confirmed composition defect
is the next bounded correction; normalized matching does not establish exact text
or visible ink fidelity.

Authored whitespace now survives comment composition. The old review-only filter
suppressed every whitespace-only emission part in any commented paragraph; it
retained positioned gaps but damaged PDF text extraction. No generated comment
placeholder exists in this path. Removing the filter and its plumbing preserves
single/multiple/nonbreaking separators, same-run references, a following footnote
and an explicit break. Font cmap aliases may extract NBSP as ordinary space.

Twenty-two selected methods pass once in one invocation, and two affected portable
methods pass on Ubuntu. Six portfolio PDFs retain bytes; all seven pages of the
four changed review outcomes retain raster bytes at 144dpi. The two independent
Office source controls retain their visual metrics. All ten content screens pass;
eight DOCX numerical gates still fail and both decks pass. Seventeen assets/four
exact smoke PDFs transfer without root build or suite replay.

Freezing now rejects differing library copies in the library, tests and CLI output
directories. A test-only rebuild exposed a stale CLI copy; the coherent snapshot
retains the already-passed library/test hashes, so qualification is reused. Its
small portfolio outputs were rechecked once and all twelve comparisons reused by
PDF/reference/threshold identity. This extra check resolves the concrete coherence
failure rather than adding routine qualification work.

Printed inline mathematics now separates baseline placement from its qualified
occupied frame. The ordinary body path compared scaled mathematical extents with
an unscaled local text inset; the ordered path omitted the first-body baseline
anchor transition used by ordinary text. Both now place the local baseline in the
same physical scale and retain the source frame's minimum occupied height. Label
leading, glyph drawing and container fallbacks remain. No font/size/document-shape
qualification is added. Earlier experiments that also reduced occupied height
regressed three review documents and were rejected.

Ten source controls (literal/fraction/root, 12/18-point leading anchors, single or
multiple formulas) have twenty independent Word mode references. Mathematical
baseline gaps pass the 0.4pt placement bound; all twenty preserve simple math text.
A pinned-font transformation test proves that adding a fitting second formula
preserves its baseline and following advance. It fails on the previous library
with a 1.360pt review displacement. Twenty-three selected methods pass once in one
invocation, plus two affected Ubuntu methods.

Six portfolio PDFs retain bytes. Each review document's page-one MAE improves by
about 0.348; all 676 normalized page-local text pairs resolve, and maximum review
start Y difference falls from 3.166pt to 2.223pt. All ten source/Office content
screens pass. The numbered body neighbor keeps its visual metrics. All eight
DOCX numerical gates still fail; both decks pass. Seventeen assets/four exact
smoke PDFs transfer without root build or suite replay. Cached comparisons reuse
checked PDF/reference/threshold identities.

A broader following-line assertion revealed an existing literal automatic advance
discrepancy (2.219pt in one review control). It is retained as an unresolved
measurement issue; the accepted baseline correction preserves occupied height.
Current accepted-build checks also supersede older broken-literal story residuals:
Final cell/header/footer prose baselines differ from Word by at most 0.364pt after
the reader correction. These findings do not establish complete ink/clipping or
reading-order fidelity.

Fitting inline math/text rows now use the ordinary text frame when their
mathematical bounds fit above and below its physical baseline. They no longer
retain a second unscaled font frame. Taller/display math and image-containing
rows preserve their qualified occupied extents; no font/size/document-shape
qualification is added. The row retains its chosen text baseline separately from
object bounds and label leading.

The independent Word regression fails on the prior accepted library by 2.219pt.
Across twenty mode controls, the largest following-paragraph gap error is now
0.199pt (previous literal cases reached 2.297pt). All preserve their mathematical
text. Literal review MAE improves about 0.258–0.261; SSIM failures remain. The native
assertion now checks both the mathematical baseline and following advance. A
pinned-font ordinary-text substitution exercises the same line-box contract on
Windows/Ubuntu; it already passed with the prior library, so it is integration
coverage rather than a claimed reproduction of the native-font defect.

Twenty-eight selected methods pass once in one invocation, plus three Ubuntu
methods. All ten portfolio PDFs retain bytes and independent comparisons; global
and page-local content screens pass. All eight portfolio DOCX numerical gates
still fail and both decks pass. Seventeen assets/four exact smoke PDFs transfer
without root build or suite replay. The compound-document allocation baseline
was separately refreshed at the preceding 0fa503ac artifact; it does not establish
a repeated-layout cost or a speedup. Superseded compiled snapshots were archived
and hash-verified, retiring another 133.8MB while preserving inputs and evidence.

Isolated cell equations now scale the transition from the nominal cell baseline,
sharing the ordinary text rule. Models without a recorded nominal cell anchor retain the source-size fallback;
parsed empty runs normally carry the paragraph font metrics. Static stories with already physical
tops retain their mapping. No font/size/document-shape qualification is added.

Six new source controls cover literal/fraction/root content, alone in a cell or
between text paragraphs, in both presets. Review literal/fraction glyph-baseline
errors fall from about 2.6pt to at most 0.137pt; root/fraction retains about 0.59pt.
Final PDFs retain bytes, and all six review comparisons improve. Four independent
native baseline controls pass; the prior library fails by 2.682pt. Twenty-three
selected methods pass once in one invocation and three portable methods pass on
Ubuntu. The portable cell controls use normal-text fractions; the pinned face does
not cover the original mathematical italic Greek fixture.

Six portfolio PDFs retain bytes; all review comparisons improve and fractional
baseline differences are at most 0.704pt. All ten content screens and normalized
review pairs resolve. Eight DOCX numerical gates still fail and both decks pass.
Seventeen assets/four exact smoke PDFs transfer without root qualification build
or suite replay. One accidental root prototype build was corrected by restoring
all seventeen qualified assets by hash. The record distinguishes that incident
from the frozen qualification.

Visual inspection also confirms an existing composition defect in the new
zero-after-spacing controls: a table top border crosses preceding descenders,
where Word leaves a gap. The table frame's eligibility and coordinate ownership
remain the next bounded task. The baseline correction improves the equation and
preserves frame geometry; it does not claim to resolve that overlap or establish
complete clipping fidelity. Source-backed geometry, independent references and
original numerical failures remain retained.

Natural multi-paragraph cells now retain the same printed table origin as
single-paragraph cells. The old one-paragraph guard disabled the shared geometry
projection, making borders cross preceding descenders. Every paragraph must still
use natural line spacing; fixed paragraph spacing and the existing richer-content,
authored-row, alignment and fragment fallbacks remain. Page-admission eligibility
is separate and unchanged. Empty resolved runs already carry equation-only font
metrics; tracing rejected the earlier missing-owner hypothesis.

Independent controls place the table after plain text, fitting mixed mathematics
and an image, in both presets. Top-border error falls from about 2.6pt to at most
0.223pt. A portable transformation proves that adding cell paragraphs increases
occupied height without shifting the table top. Both it and the native Word gap
check fail on the previous library. Fourteen selected methods pass in one
invocation, plus three Ubuntu methods. Ten affected control outcomes retain their
mathematical text; original visual failures remain explicit.

The fixed portfolio contains only one-paragraph cells, where the predicate remains
equivalent. Its prior complete qualification is retained, with no full-portfolio
reconversion or new full-corpus claim. Letter and held-out memo review smoke PDFs
retain bytes. Seventeen assets/four exact integration PDFs transfer without root
build or suite replay. The generic measurement tool now calls custom-manifest
results pinned document/mode outcomes, distinguishing them from full portfolio
qualification. The remaining C1 task is the finite document severity and visible
clipping/overlap review, retaining independent thresholds and unresolved evidence.


The finite C1 portfolio severity review is complete. All sixteen candidate/Office
page pairs were inspected, using comparison-chain report hashes and exact PDF
identities. No missing visible content or serious pagination, overlap or clipping
defect was observed in these documents. Global and page-local prose/simple-math
screens remain clear, and both source PNG colors retain visible patches in all
eight DOCX outcomes. This is bounded visual evidence, not a proof of arbitrary
reading order, glyph ink bounds or unsupported combinations.

| Document | Final / review pages | Content / serious placement | Residual outcome |
|---|---|---|---|
| Report | 2 / 2 | No omission or serious defect observed | Typography/table spacing; review pane and connector geometry |
| Letter | 1 / 1 | No omission or serious defect observed | Typography/table spacing; review pane and connector geometry |
| Specification | 2 / 2 | No omission or serious defect observed | Typography/table spacing; review pane and connector geometry |
| Held-out technical memo | 2 / 2 | No omission or serious defect observed | Table/baseline spacing; review pane and connector geometry |
| Briefing deck | 1 | No omission or serious defect observed | Numerical gate passes |
| Held-out deck | 1 | No omission or serious defect observed | Numerical gate passes |

All eight DOCX numerical visual gates remain failed under their original
thresholds; neither manual review nor the bounded reliability checkpoint relabels
them. Review pane background extent is visibly different even where content is
clear. The held-out documents remain acceptance checks, outside derivation.
Detailed per-outcome severity, page raster hashes and comparison provenance are
retained in `artifacts/composition-c1/c1-portfolio-severity.json`. The source PNG
patch refresh is `portfolio-image-patches-current-frame.json`. Latest production
2ad7208a retains the 07b20914 full portfolio through the documented one-paragraph
predicate equivalence and exact integration smokes; no corpus reconversion was
needed for this review. CI 38820298013 (root 66d66fc1) is green on both platforms.

C1's bounded mixed-content reliability checkpoint is accepted with these explicit
visual residuals. Wider rich/related/floating stories and OMML remain on the
roadmap. The next maintenance step reuses unchanged immutable fixture snapshots;
subsequent rendering work returns to supported-content composition rather than
fitting the held-out portfolio or lowering visual thresholds.


Qualification freezes now share a read-only public fixture snapshot keyed by its
Git Cases tree across artifacts and linked worktrees. Every freeze still checks
all source/snapshot SHA values and file counts, then links its runner's Cases
directory to that snapshot. Changed input trees remain separate. Seven focused
checks pass on Windows and Ubuntu, including dirty inputs, corrupt snapshot and
mixed-library rejection, and linked-worktree reuse. CI runs the same small tool
checks; no library suite is added.

Applied artifact `frozen-shared-inputs` verifies 1,227 fixtures and preserves the
previously qualified 2ad7208a library bytes. One runner integration method and one
CLI conversion through the fixture link pass, the latter reproducing the accepted
PDF exactly. No library rebuild or qualification replay was needed. Evidence:
`artifacts/composition-c1/shared-fixture-evidence.json`. Historical fixture copies
and the previously rejected retirement targets remain untouched.


Ordinary default/left tabs now coexist with supported equations through the
ordered measurement path. The existing shared tab-stop resolver consumes preceding
text/math/image advances; body, cells, direct headers/footers, keep estimates and
autofit receive the document's default interval. Tab positions scale once. Right,
center, decimal, bar/leader and out-of-frame stop semantics keep explicit fallbacks.
Visible tabs prevent sole-formula display promotion. No font/size/document-ID
qualification is added.

Eight public source controls exercise an 18pt default interval and authored 72pt
left stops in four containers, separately in both presets. The prior library
omits both formulas in all sixteen outcomes; native and portable regressions
reproduce it. The accepted build retains both formulas, prose and one-page
placement. Independent Word relative field offsets differ by at most 0.052pt.
Review cell baselines remain about 2.5pt high; reviewed body/cell/header/footer
crops show visible unclipped formula content. All sixteen original numerical
control gates remain failed; this change establishes composition, not exact
typography.

Frozen 32f59eae / root b42436cb passes 23 selected methods in one invocation and
three Ubuntu methods. Source ordering, wrapping, scaled stops, cancellation and
explicit fallback controls pass. A read-only fixture workflow issue initially
blocked seven mutation neighbors: the helper copied the source attribute onto its
temporary DOCX. Only the temporary copy is now writable; the final artifact reuses
the verified 1,237-file shared snapshot. Failed preliminary evidence is retained.
All sixteen accepted control PDFs match the prototype and reuse pinned Office
comparisons. The portfolio has no authored inline tabs, so its complete previous
qualification is retained; letter and held-out memo review smokes retain bytes.
Seventeen assets/four integration PDFs transfer without a root rebuild or suite
replay. Evidence: `artifacts/composition-c1/math-left-tabs-evidence.json`.


A follow-up footprint check found 6.089GB, principally historical runtime
archives. Their obsolete compiled blobs and 86 Git-export duplicates were
retired after every exported file was checked against reachable Git history
(recording checkout CRLF normalization for text). Four exports lacking reachable
commits were retained. Compact archives retain original path maps and explicit
retirement metadata; all remaining blobs were SHA-verified. Current/rollback
runtimes, source inputs, PDF/raster evidence, uncommitted patches and Git refs
remain. No previously rejected retirement target was retried. The workspace is
4.606GB / 257,880 files with no scan errors; proof and compaction maps are under
`artifacts/workspace-maintenance/`, and the complete check is
`artifacts/workspace-footprint-after-archive-compaction.json`.


Equation/comment composition now keeps ordered source boundaries and measured
visible extents. Comment ranges attach as metadata; prose, formula-only, between-
formula and mixed selections preserve equations in body paragraphs and cells.
Highlights and connectors use placed source geometry, including the formula's
visible content rather than its padded layout advance. Adjacent prose owns the
inter-object whitespace. The line owns selection height independently of range
endpoints. Legacy annotation geometry remains on its existing path. Note/field/
revision and broader story combinations retain staged fallbacks.

Eight public sources have independent Word exports in both presets. Twelve
initial parent outcomes omitted both equations; native and portable regressions
fail on that prior library. The accepted base retains mathematical/prose content
in sixteen outcomes and passes eight independent selection-width checks, with
maximum horizontal error 0.080pt. Bookmarks and unchanged-format run splitting
retain source ranges. The sole-formula policy also retains display ownership.
Eight review crops show visible equations, prose and comment captions; cell
baselines, selection height, connectors and balloon/pane cosmetics remain
approximate. All sixteen original numerical visual gates remain failed.

Base artifact 535a5d8e / root c5130a51 passes fifty selected methods in one
invocation and four Ubuntu methods. Follow-up 6ff2f99e / root 9478690f corrects
line-owned highlight height; its regression fails on the base artifact. Eight
affected methods pass once in one invocation, plus one Ubuntu method. Only the
four prose/between source documents are reconverted for that follow-up (eight
mode outcomes); all four Final PDFs and their comparisons retain identity. Other
base qualification remains retained, rather than reporting a new complete run.
The 1,247-input snapshot is reused. Seventeen assets/four exact PDF smokes transfer
without a root rebuild or suite replay.

The fixed portfolio has no paragraphs combining comments and equations; retain
its prior full qualification and severity review, with exact letter and held-out
memo review integration PDFs. The held-out document remains outside derivation.
Evidence: `artifacts/composition-c1/math-comments-evidence.json`.


Continuous automatic footnote references now retain explicit resolved marker
owners in the ordered paragraph source. Equations and note glyphs coexist before
or after each other in body paragraphs and cells, with or without comment parts.
Shared authored runs and starts 1/9/12/103 retain ownership. Custom/unowned labels,
page/section/unknown restart requests and equation/endnote combinations retain
staged fallback. Plain-text correction eligibility now excludes equations, a
semantic ownership rule instead of another font/size/document-shape exception.
The prior predicate demonstrably accepted equation paragraphs; no separate
visible regression from that misclassification is asserted.

Ordered placement also scales source vertical run offsets once, preserving
superscript marker raises in physical review frames. Sixteen independent Word
mode controls retain both formulas, note text and destinations; relative marker
raises differ by at most 0.040pt. An earlier prototype fails this assertion by
about one point. Review cell baselines still differ by about 2.5pt, and all sixteen
original numerical control gates remain failed. Four reviewed review crops retain
visible equations and note content. The previous library omits both equations in
all sixteen mode probes; native/portable content regressions fail there.

Base 951e1370 / root eb790419 passes 54 selected methods in one invocation and
five Ubuntu methods. A rendered shared-run control exposed an additional link
ownership defect: an authored digit identical to the automatic label acquired a
second link. Follow-up 5ec7ac65 / root 41f224c8 retains the marker role through
formatting/emission and links only that owner. Its regression fails on the base;
14 focused methods and two Ubuntu methods pass. The controls cover paragraphs
with and without equations, plus existing page numbering, custom marks and
endnote behavior.

The base sixteen source/mode outputs and Office evidence are reused: canonical
reference runs contain no colliding authored digit. No full control/corpus replay
was needed for the role follow-up. Seventeen assets/four exact integration PDFs
transfer, including two canonical footnote controls and the portfolio letter/
held-out memo. The 1,258-input snapshot is reused. No root rebuild or suite replay
is performed. Portfolio paragraphs do not combine equations and footnote
references, and their tables/static stories/comments already exclude plain-text
corrections; prior complete qualification remains. Evidence:
`artifacts/composition-c1/math-footnotes-evidence.json`.

C2 adds a separately pinned, held-out composed annotation report. The existing
qualified library preserves all prose and equations in both DOCX presets, but
moves a complete review section to page 2 where Word retains it on page 1.
The checkpoint remains open. Independent small note-admission controls must
reproduce the cause before changing layout; the acceptance report stays outside
derivation. Reuse committed supplemental manifests with
`AuditCompositionContent.py --input-root . --manifest <public-manifest>` and the
existing frozen runtime. Input Git identity, source/PDF hashes and public-only
boundaries remain checked. Legacy ten-outcome compatibility and three negative
input controls pass without new conversions or library builds. Evidence is in
`artifacts/composition-c1/supplemental-audit-tool-evidence.json` and
`composed-report-content-audit.json`. CI through `be411f37` is green on both OSes.

An independent atomic-footnote reduction confirms unused future continuation
reserve can cause a spurious second page. `d2c5dde6` reserves only the complete
head for a new one-line note; split notes and continued tails keep their existing
reserve. Thirty-nine affected methods and five Ubuntu checks pass. The source-only
control-width follow-up `99991423` reuses all build assets and repeats only two
fixture-dependent methods (one on Ubuntu); all three new regressions fail with
the parent library. Four Word controls now match one-page membership and retain
all content, while their numerical gates still fail (note/separator Y about 6pt).
All ten original portfolio PDFs are exact, so Office comparisons and page reviews
are reused. Transfer verifies 17 assets and four PDFs, without root build/replay.
C2 Final now matches Word's page-local prose and its two page pairs show no serious
placement defect. Review remains unchanged and still moves section 7 to page 2;
both modes retain all equations/prose, but both numerical gates fail. C2 stays
open for an independent scaled-admission reduction. Evidence:
`artifacts/composition-c1/atomic-note-evidence.json` and
`composed-report-atomic-note-content.json`. Workspace is 4.941GB at the latest scan.

C2 now passes prose, simple-math and page-local content screens in both presets;
all four held-out page pairs are reviewed. Table-fraction/border clearance,
annotation geometry and typography remain residuals; both numerical gates fail.
`c433142b` converts note reserve once to printed paragraph points and uses shared
resolved script sizes for existing exact-line eligibility. Four valid standard-page
Word controls retain eighteen body lines and their note on one page; the parent
fails all three new regressions. Custom short-sheet review references were rejected
for out-of-page Word ink. The final frozen identity qualifies 44 affected methods
in one invocation and five Ubuntu methods. A test-only revision-metadata rebuild
required that final qualification; retain the qualified DLL for future test-only
builds with `--no-dependencies`. Four original review portfolio PDFs remain exact.
Transfer/remaining PDF identity smokes use no root build or suite replay.
`39d7beab` repairs the sole Windows CI failure at `fcc0ad96`: its synthetic column
test now retains the two-column frame and verifies all four note destinations on
the admitted page. Ubuntu was green; the next push is under CI. Evidence:
`artifacts/composition-c1/scaled-note-evidence.json`. C2 is accepted for bounded
composition with the numerical/geometry residuals; release preparation stays deferred.

`256ca0a4` corrects the shared review table-fraction clearance. Ordered cell frames
measure printed text struts and receive the row-origin translation without a
second nominal-inset conversion. Eight independent controls vary note-marker size
11/12 and script alignment; maximum review border-relative error falls from
2.652pt to 0.149pt. Four Final PDFs are exact. Twenty-eight focused methods in
one invocation and five Ubuntu checks pass; the parent fails both new regressions.
C2 retains all source/page-local content; its Final PDF and review page-2 raster
are exact, while the corrected review page 1 is reviewed. Both numerical C2 gates
and all eight control gates remain failed. Eight Office prose-token screens remain
unresolved because normal-text math words use Office's prose font; source prose,
simple math, direct glyph geometry and reviewed crops retain the content. The
fixed portfolio excludes changed ordered cells; prior Office reviews plus exact
letter/memo smokes remain accepted without full-corpus replay. Transfer checks
17 assets/four PDFs without root build or suite replay. Evidence:
`artifacts/composition-c1/ordered-cell-frame-evidence.json`. CI through `ed9566a9`
is green on Windows and Ubuntu. Workspace was 5.069GB at the last scan; automatic
review rejected retirement of 48 verified compiled duplicates (blocked by policy),
so no files were removed and those targets will not be retried.

`17b0e6cb` bounds future fixture growth by sharing verified read-only files between
published input trees. Checkout files remain independent; corruption is rejected
both when reusing a tree and when seeding a new one. Ten tool checks pass on
Windows and Ubuntu. A real snapshot shares 1,271 files / 12.09MB, copies 7.2KB of
byte-distinct metadata and rechecks all 1,274 input hashes after a temporary-OOXML
mutation test. All 17 compiled assets and one CLI PDF are exact, without library
build or suite replay. The footprint command retains its PowerShell interface and
uses Python standard-library file IDs to count hard links once; both-OS tests also
reject directory loops. Latest size: 5.159GB unique content / 5.171GB logical file
lengths, near the loose cap. Historical/denied retirement targets are untouched.
Evidence: `artifacts/composition-c1/cross-tree-inputs-evidence.json`. CI through
`3b7b8865` is green on both OSes; resume annotation geometry with existing references.

`1765e222` maps review-pane vertical bounds through the shared printed paper
frame, independently of body margins, and retires two fitted inset constants.
The same Y mapping preserves floating-textbox emission. Nine independent Word
controls cover six margin pairs, Letter/A4/landscape and mirrored pages: maximum
error is 0.438pt, down from a 377.571pt body-margin error. Twenty-six affected
methods in one invocation and five Ubuntu checks pass; both parent regressions
fail. The shared paint change justifies one fixed-corpus run: all ten content/
page-local screens pass; text operations and non-pane graphics remain exact;
seven review pages change only gray paper pixels, and six Final/deck PDFs are
exact. Eight DOCX gates still fail and two decks pass. C2 content/pagination remain
qualified in both presets, with both numerical gates failed and connector/range
geometry still residual. All nine new control gates remain failed. Transfer uses
17 checked assets/four PDFs, without root build or suite replay. `-DocxPreset Review`
selects only relevant DOCX outcomes for subsequent control batches. Evidence:
`artifacts/composition-c1/review-paper-frame-evidence.json`. CI at `c71ef4c8`
passed Ubuntu and Windows rendering but exposed a Windows long hard-link path;
`367656f1` uses extended paths, passing ten tool checks on both OSes and a
277-character shared-input path without renderer rebuild/requalification.

`3b79ab77` makes comment selections consume the occupied inline frame, including
its baseline strut, with visible bounds allowed to enlarge it. Eight independent
Word controls covering body/cell and prose/formula/mixed ranges reduce the height
deficit to at most 0.119pt; baseline-relative top/bottom error stays below 0.187pt.
Nineteen capability methods pass in one invocation, four portable Ubuntu checks
pass, and both new regressions fail with the previous library. All eight control
text-operation files remain exact; four paired crops are reviewed. No font/size/
document eligibility branch is added. All eight numerical control gates still
fail. C2 content and page-local screens pass in both presets; Final PDF and both
text-operation files are exact, with only five review selection operations changed.
The image/math selection crop is reviewed; both numerical C2 gates remain failed.
Connector and balloon attachment remain explicit residuals. Eight older Office
prose screens for normal-text equation words remain unresolved. Frozen `0d02535c`
qualifies once; root verifies 17 assets/four PDFs without rebuilding or replaying
the suite. Evidence: `artifacts/composition-c1/comment-occupied-frame-evidence.json`.
CI through `881f8871` is green on Windows and Ubuntu.

`9a14f5ac` shares measured selection depth with comment connectors while retaining
independent balloon placement. Eight Word controls reduce baseline-relative
connector error from 2.22pt to at most 0.112pt. Twenty-five methods pass in one
invocation and four portable Ubuntu checks pass; both new connector regressions
fail with the preceding library. All eight text-operation files, including balloon
text, stay exact; only two connector strokes change per control. A first prototype
coupled the two coordinates and displaced balloon titles; an independent title
guard demonstrates the rejected regression. Only corrected frozen `8a9db635`
transfers to root. Four control pairs are inspected; all eight numerical gates
still fail. C2 passes both content/page-local screens, retains exact Final PDF and
both text-operation files, and changes only two review connector strokes. Both
numerical C2 gates remain failed. Connector X/dash, balloon baseline and typography
remain approximate. Transfer checks 17 assets/four PDFs without root rebuild or
suite replay. Evidence: `artifacts/composition-c1/comment-connector-placement-evidence.json`.
CI through `6400abec` is green on both OSes.

`1ba76589` retains supported equations beside materialized cached fields. Shared
semantic source ownership replaces the blanket field exclusion; complex field
results packed with their control markers now enter the ordered source view even
after a formula. Eight controls / sixteen Word comparisons cover simple/packed
REF results in body, cells, headers and footers. Both fractions and source prose
remain in all outcomes, field advances agree within 0.144pt, and printed baselines
within 0.375pt. All eight simple/packed PDF pairs are exact. Twenty-nine methods
pass in one invocation and five portable Ubuntu checks pass; four new regressions
fail with the preceding library. Four paired crops are reviewed. All sixteen
numerical visual gates remain failed. The generic audit also retains sixteen
unresolved field-source/math-token classifications, despite zero source-prose and
simple-math deficits; these are not relabeled as passes. PAGE/NUMPAGES, unresolved
fields and revision owners retain their fallback. This adds cached-result display,
not field evaluation. Frozen `4609b4bb` transfers with 17 assets/four exact PDFs,
including accepted C2 review and held-out memo review, without root rebuild or
suite/corpus replay. The reference command adds an optional DOCX preset selector;
verified cached review references are reused without export. Evidence:
`artifacts/composition-c1/math-fields-evidence.json`. CI through `4ed093b6` is green
on both OSes. The latest workspace scan reports 5.365GB with four long-path fixture
traversal errors; those errors remain explicit pending the bounded tool correction.

The workspace counter now uses extended Windows paths for traversal while retaining
its ordinary reported root. A regression reproduces the missing long-path byte;
Windows wrapper and direct Ubuntu Python checks retain unique/logical byte, hard
link and loop behavior. The complete scan reports 5.379GB with zero errors. The
Ubuntu PowerShell wrapper remains a CI check; local Ubuntu has no PowerShell.
No library build or rendering suite is replayed. Evidence:
`artifacts/composition-c1/footprint-longpath-evidence.json`.

The bounded C3 cached-field checkpoint reuses frozen `4609b4bb` for a mechanical
held-out report variation: visible prose, six equations, one image and all package
parts except main XML remain unchanged. Simple/packed cached results coexist with
math, the image/comment range and an internal link. Both presets retain two pages
and pass all source/Office prose, simple-math and page-local screens. Existing four
link targets survive; review link end differs by 0.114pt. Merged text geometry has
zero failures (88 Final / 90 review lines), and graphics operations are exact
against accepted C2. Candidate and Office second-page rasters are exact; changed
first-page regions are inspected. Both numerical gates remain failed. Source
commit `43339a3d` and tool `b583a8d7` pin the independent input/audit work; all 17
runtime assets are unchanged, with no build or rendering-suite/corpus replay.
Twelve source-audit adversarial controls pass on Windows/Ubuntu; cached fields
resolve their old classification failures while dynamic/malformed/cross-paragraph
fields remain unresolved. Ten C1 audit screens still pass. Sixteen fraction probes
retain rate/time token deficits and all their visual failures. Initial acceptance
phrases were corrected to the actual baseline note/text; document bytes and both
candidate PDFs remain exact, without Office re-export. Evidence:
`artifacts/composition-c1/composed-fields-evidence.json`. CI through `373c99ad` is
green on both OSes; workspace traversal now reports 5.379GB with zero errors.

`934c79a1` supports PAGE/NUMPAGES beside equations in page-owned direct headers and
footers. Four Word controls / eight outcomes cover twelve pages and the 9/10 digit
boundary; every one of the 96 page outcomes retains the actual page values and
both fractions. Forty-seven methods pass in one invocation, six Ubuntu checks pass,
and two rendering regressions fail with the preceding library. Baseline and leading
field error stay within 0.375/0.120pt. Word's simple-field digits use Calibri while
cached source formatting retains Arial: count-field advance residual is 1.115pt,
and cumulative end offset reaches 2.396pt. Four pairs are reviewed; all eight control
gates remain failed. Complex-field cache ownership now ends at its end marker,
retaining following authored spaces/numbers; simple placeholders publish source
items. Shared substitution replaces three copies. Body/cell and related-story
dynamic math retain fallback; source placeholders cannot qualify without actual
page inputs. Ten fixed portfolio PDFs are exact, reusing Office comparisons and
ten passing content/page-local screens; eight DOCX gates still fail and decks pass.
Frozen `89d9d93a` transfers 17 assets/four exact PDF smokes, including accepted C3
review and held-out memo, with no root build or suite replay. The generic prose
audit deliberately leaves dynamic source fields unresolved; independent per-page
checks qualify this slice. Evidence: `artifacts/composition-c1/static-page-fields-evidence.json`.
CI through `cce36176` is green on Windows and Ubuntu.

`72ae0ebd` shares owned automatic-reference rules across footnotes and endnotes,
retaining restarted/custom/unowned fallbacks. Four Word controls / eight outcomes
keep both fractions, note text and the destination in body/cell, before/after
combinations. Twenty-nine methods pass in one invocation and five Ubuntu checks
pass; three regressions fail with the preceding library. Single/repeated formulas,
bookmarks, run splitting, wider Roman starts and adjoining marker-run text are
covered. Marker raise/font ratio agree within 0.040pt/0.000187; X residual is
1.243pt. All eight visual gates remain failed, as do the generic rate/time token
screens; source prose and simple-math deficits are zero. Four pairs reveal a
separate significant defect: endnote story gap is excessive by 8.163pt in Final
table cases and about 6.213pt in review body/table cases. Notes remain visible;
broader composition acceptance waits for semantic spacing work. No fitting branch
is added. Frozen `3b79f164` transfers 17 assets/four exact PDFs, including C3 review
and held-out memo Final, without root build or suite/corpus replay. Evidence:
`artifacts/composition-c1/math-endnotes-evidence.json`. CI through `500cdacf` is
green on Windows and Ubuntu.

The document-endnote start boundary now belongs to its trailing body paragraph,
with numeric after-spacing and empty-separator flow resolved locally. Earlier
tables and review print scaling no longer add 8.163pt / 6.213pt of excess gap.
Whole-note typography, continuation, mixed-footnote and between-note guards remain.
Frozen 2b320cd8 / root 4b811739 passes 42 affected methods in one invocation and
six Ubuntu checks; both new regressions fail with the preceding library. Eight
independent Word control outcomes agree within 0.163pt on the story-baseline gap.
Body text operations are exact and only note/separator baselines move; four visual
pairs were reviewed. All eight numerical visual gates and generic rate/time token
screens remain failed; source prose/simple-math deficits are zero. The checked
17-asset/four-PDF transfer preserves C3 review and held-out memo Final output,
without rebuilding or replaying a root suite. Evidence:
`artifacts/composition-c1/endnote-boundary-evidence.json`. A held-out complete
endnote report is the next acceptance step; reduced controls alone do not
establish general related-story composition.

The held-out composed-field report also passes the bounded endnote-composition
checkpoint after mechanically replacing its footnote with an endnote. Visible
prose, six equations, image bytes and unrelated package parts are preserved.
Both presets retain two pages, all source/Office global and page-local prose and
simple math, and four navigation targets; the note destination moves to page two.
Four page pairs were reviewed without a visible omission or serious pagination,
overlap or clipping defect. Note-baseline residual is 0.256pt Final / 0.023pt review.
Both numerical visual gates remain failed; typography and annotation residuals
remain explicit. Source 0b8ceeec reuses frozen 2b320cd8 with all 17 asset hashes
rechecked, no rebuild, suite replay or new fitting branch. The shared authoring
transformation also regenerates the four reduced controls byte for byte.
Evidence: `artifacts/composition-c1/composed-endnotes-evidence.json`.

Direct normal footnote/endnote paragraphs now retain supported mixed equations.
Preparation and diagnostics preserve their note-part source identity; ordered
measurement supplies occupied line boxes, and the common printed note map scales
math and annotation extents together. After-spacing belongs to the final line.
Frozen c57f2dd2 (compiled fd6abcfe) / root c24c1249 passes 56 affected methods in
one invocation and six Ubuntu checks; three regressions fail with the previous
library. Bookmarks, run splitting, repeated formulas, wrapping and real multipage
continuations retain source content and destinations. Eight independent Word
outcomes retain all four body/note fractions and prose, with first-note baseline/X
errors below 0.367pt / 0.062pt. Five crops were reviewed. All eight numerical gates
and generic rate/time token screens remain failed; source prose/simple-math
screens have zero deficits. A significant review-endnote width defect remains:
the wrapped note adds a line, shifting numerators 24.645pt horizontally / 11.439pt
vertically. Broader note-content placement acceptance waits for that correction.
Images, note tables and special-separator equations retain staged fallbacks.
The 17-asset/four-PDF transfer preserves held-out endnote-report review and memo
Final without rebuilding or replaying a suite. Evidence:
`artifacts/composition-c1/note-content-evidence.json`.

Normal footnotes and endnotes now share their owning section's full design text
width before the review print transform. This removes the extra wrapped endnote
line and 24.645pt X / 11.439pt Y formula displacement. Frozen 86b6dd92 / root
eb81f193 passes 60 affected methods in one invocation and four Ubuntu checks;
both new regressions fail with the preceding library. Eight Word outcomes retain
all four fractions and source prose/simple math; note numerator residuals are at
most 0.704pt X / 0.381pt Y. Six identical PDFs reuse comparisons, and both changed
review crops are reviewed. All eight numerical gates and generic math-token
screens remain failed. The complete C7 report changes only two note clip widths:
text operations and both raster pages remain exact. Its affected review outcome
passes source/page-local screens. The checked 17-asset/four-PDF transfer uses the
verified new PDF identity without a root build, suite or corpus replay. Audit
selection now honors measured Both/Final/Review modes; 14 adversarial controls
retain missing/extra/duplicate rejection on Windows/Ubuntu. Broader composed
note-math acceptance remains the next step. Evidence:
`artifacts/composition-c1/note-width-evidence.json`.

The held-out composed endnote report also retains two supported fractions inside
its note, alongside all six original equations, prose, image, cached fields and
annotations. Only the endnote package part changes. Both presets retain two
pages, pass global/page-local prose and simple-math presence screens, and preserve
four navigation targets with the note destination on page two. All four page
pairs were reviewed without visible omission or serious pagination, overlap or
clipping defects. Note-baseline residual reaches 0.302pt; note-numerator horizontal
residual reaches 1.682pt. Both numerical visual gates remain failed. Source
7ede0fb3 reuses frozen 86b6dd92 with 17 asset hashes checked, no library rebuild,
suite replay or new fitting branch. The authoring command's original mode still
regenerates the prior C7 source byte for byte. Evidence:
`artifacts/composition-c1/composed-note-math-evidence.json`.

Direct normal notes now compose inline images and equations through shared
measured frames. Note-image positions and dimensions join the text/math printed
map; owned inline frames also retain local endnote start-spacing ownership.
Existing whole-flow, table/floating and rich-story fallbacks remain. Frozen
65b9c4fe / root 2cc71652 passes 67 affected methods in one invocation and five
Ubuntu checks; three regressions fail with the previous library. Four Word
control documents in both presets cover inline/wrapped notes and two image sizes.
All four fractions, PNG, prose and destination survive; source prose/simple-math
deficits are zero. Eight exact prototype PDFs reuse comparisons, and four source/
mode identities reuse Word references. Four crops were reviewed. All eight
numerical gates and generic rate/time screens remain failed. PNG raster-bound
residuals reach 0.5pt top/size and 1pt horizontal prose advance. Native dimension
checks account for the two-edge 144 DPI sampling uncertainty; visual thresholds
are unchanged. The 17-asset/four-PDF transfer preserves complete note-math report
review and held-out memo Final without a root rebuild or suite/corpus replay.
Complete-document note-image acceptance remains the next finite step. Evidence:
`artifacts/composition-c1/note-images-evidence.json`.

The finite normal-note composition checkpoint is also accepted on the complete
held-out report with a note image beside its equations. Source 258b3328 adds only
the note drawing and its relationship, reusing the original PNG bytes. Eight
source equations, two image placements, prose and four navigation targets survive
both presets on two pages, with zero global/page-local content deficits. Both
first-page rasters are exact for candidate and Word, reusing their prior reviews;
the two new second-page pairs were reviewed without visible omission or serious
pagination, overlap or clipping defects. Note-image dimensions agree in the
sampled raster; top/X residuals reach 0.5pt / 1.5pt. Both numerical visual gates
remain failed. All 17 assets of frozen 65b9c4fe were rechecked with no library
build, suite replay or fitting branch. Further note variants are outside this
finite acceptance step. Evidence:
`artifacts/composition-c1/composed-note-image-math-evidence.json`.

Shared inline counter measurement now resolves PAGE from its actual page number
and NUMPAGES from its actual total, without requiring the unrelated input. Mixed
counter paragraphs still require both; source ownership and missing/noncanonical
fallbacks remain. Simple/packed body/cell source controls cover values 9/10/12 and
resolved advances. This is a dependency prerequisite; body/table page-aware math
admission remains deferred. Frozen 8fe300f6 / root 6190893d passes 26 capability/
neighbor methods in one invocation and five Ubuntu checks. The measurement
regression fails with the preceding library through existing APIs; the new
resolver-API unit is excluded from that parent check. The 17-asset/four-PDF
transfer preserves the complete note-image/math report in both presets, the
twelve-page static-header counter control and held-out memo Final. No root build,
suite/corpus replay or Office re-export; prior comparisons remain reusable with
unchanged failed numerical gates. Evidence:
`artifacts/composition-c1/counter-inputs-evidence.json`.

Reusable inline measurement and source-driven wrapping are now separate. Frozen
9550187c / root 3527ff81 passes 23 exact capability/neighbor methods in one
invocation and ten Ubuntu checks. A source cursor preserves authored breaks,
tabs, first-line numbering and wrapping frames; split counters retain the value
bound when their first digits were placed. Rebinding updates only changed counter
fragments and keeps canonical source ownership. All ten fixed-portfolio PDFs are
byte-identical to prior qualified outputs, reusing independent Office comparisons
and their explicit visual failures. Transfer verifies 17 assets/four exact PDFs
without a root rebuild or suite replay. This establishes the contract for actual
page-aware admission; body/table PAGE-math support and measured performance gains
are not claimed. Evidence: `artifacts/composition-c1/inline-cursors-evidence.json`.

Simple table cells now preserve supported equations beside PAGE counters. The
shared preparation policy distinguishes PAGE dependencies from an unavailable
total count; measurement and painting use actual placement inputs. Existing cell
height calculation and dynamic memo keys already carry those inputs, so no new
layout flag or cache policy was added. Frozen 2086a381 / root 05dff46f passes 25
capability/neighbor methods in one invocation, six Ubuntu checks and two genuine
parent-library rendering regressions. The independent Word control preserves four
fractions and actual values across pages nine and ten in both presets. Relevant
page pairs were reviewed through exact preset raster reuse, with no visible
omission or serious overlap/clipping. Baseline error is 0.029pt; counter/equation X
residuals reach 1.041pt / 2.349pt, including Word's Calibri field digits versus
canonical Arial. Both visual gates and the generic dynamic-field source screens
remain failed/unresolved; no simple-math deficit is observed. Transfer verifies
17 assets/four exact PDFs without a root build or suite replay. Body PAGE and
unknown total-count composition remain deferred. Evidence:
`artifacts/composition-c1/cell-page-counters-evidence.json`.

Body PAGE/math composition now binds counters after line admission. A source
cursor retries unchanged when pagination moves a line, including past an unfit
continuation column. A confirmed trailing-whitespace cursor defect was reproduced
and corrected. Fixed paragraphs reuse their measured wrap; counter indices/tab
stops are retained and rebinding occurs on page changes. No speedup is claimed.
Frozen 645d5508 / root c1f41899 passes 33 exact capability/neighbor checks in one
invocation, fifteen Ubuntu checks and four genuine parent behavior regressions.
Simple/packed source tests retain bookmarks and run splits. The independent Word
simple-field control places one paragraph across pages nine and ten with two
fractions and correct values 9 then 10. Relevant page pairs were reviewed through
exact preset raster reuse, with no visible omission or serious overlap/clipping.
Both visual gates remain failed: baseline residual reaches 0.105pt, authored
leading-space trimming shifts the second prose start 3.384pt, and field/numerator
X residuals reach 3.456pt/2.148pt. Generic dynamic-field source screens remain
unresolved, with no simple-math deficit. Ten fixed-portfolio PDFs are exact;
transfer verifies 17 assets/four exact PDFs without a root rebuild or repeated
suite/corpus. Complete compound-document acceptance is next; unknown total-count
composition remains staged. Evidence:
`artifacts/composition-c1/body-page-counters-evidence.json`.

The finite PAGE composition checkpoint is accepted on the complete held-out
report (source 7a8208ea), using frozen 645d5508 without a library rebuild or suite
replay. All 17 assets were rechecked. Existing cached field words remain, with
three PAGE fields added beside body math/images, inside an internal hyperlink and
in the math/note table cell. Eight source equations, two image placements and four
logical navigation targets survive both presets on two pages; the wrapped link
uses two rectangles, giving five painted navigation rectangles. The note target
remains on page two. All four page pairs were reviewed against independent Word;
no visible omission or serious pagination/overlap/clipping defect was observed.
Global/page-local prose and simple-math deficits are zero. Generic source screens
retain their explicit unresolved dynamic-field evaluation; both numerical visual
gates remain failed. Typography, spacing and annotation geometry remain cosmetic
residuals; PNG top/X residuals reach 1pt/1.5pt with half-point raster sampling.
Unknown total-count composition remains staged. Evidence:
`artifacts/composition-c1/composed-page-counter-evidence.json`.

Shared ordered wrapping now preserves authored leading spaces and the closing
paragraph line after a trailing hard break. The paragraph-end cursor represents
the source mark explicitly; soft-wrap whitespace remains discarded. A combined
body/cell/header control corrects accumulated occupied-height errors up to 29.175pt.
Eighteen independent Word anchor comparisons have maximum X/Y errors of
0.118pt/0.139pt. The existing cross-page PAGE control's former 3.384pt start offset
falls to 0.05pt. Frozen a08dcfa6 / root 323c0539 passes 38 capability/neighbor
methods in one invocation, sixteen Ubuntu checks and four genuine parent behavior
regressions. All eighteen equations and prose remain visible in the reviewed
control; both preset page pairs share exact raster identities. Both numerical
visual gates and generic rate/time whole-word screens remain failed; source prose
and simple-math deficits are zero. Ten fixed-portfolio PDFs plus both complete
PAGE-report PDFs remain exact. Transfer verifies 17 assets/four exact PDFs with
no root build or repeated suite. Evidence:
`artifacts/composition-c1/inline-whitespace-evidence.json`.

Parent regressions now use `tools/NewParentRegressionRuntime.ps1`, with independent
owned input copies and rejection of unowned/linked destinations. Thirteen Windows
fixture-sharing/ownership checks pass and run in existing CI on both platforms;
local Ubuntu lacks pwsh, while its .NET checks pass. This follows a confirmed
agent setup error: a legacy Cases junction received six unpinned additions and
one shared input lost its read-only flag. Existing pinned bytes remain exact and
protection is restored. Automatic review blocked removal of the additions and old
junction; both remain, with no retry. The historical 444f5a55 input set is excluded
from strict immutable identity claims. All 1,337 current frozen inputs retain
pinned hashes/protection and the six current parent inputs have independent file
ownership. Ownership repair/rejection evidence remains beside the delivery report.

The measured font-maintenance batch reduces warm portfolio allocation to
28.75–33.83MB for Final DOCX and 32.34–37.36MB for review, with unchanged PDFs.
Counted kerning source batches replace the expanded pair list, retaining bounded
source contributions, last-wins lookup values and cancellation. All 715,743 map
entries across the three actual faces match. Frozen f35fa653 / root 993543d6 passes
31 Windows methods in one invocation and 17 Ubuntu checks (four installed-font
checks skip); transfer verifies 17 assets/four PDFs. The ten allocation-probe
whole/staged PDFs provide corpus identity, without an extra CLI sweep. A generated
out-of-range class-matrix regression confirms an older managed bounds exception;
wide offsets now preserve the valid prefix. This is not evidence of corruption.
Matched sustained-warm timings overlap, so no speedup is claimed. Office output
and existing failed visual gates remain unchanged. See
`docs/OperatingEnvelope.md` and `artifacts/composition-c1/kerning-stream-evidence.json`.

Simple header/footer table cells now share body-cell math preparation and consume
prepared boxes and actual PAGE/NUMPAGES inputs in grid/height/painting. A confirmed
unchanged-content transformation omitted every header equation before this change.
Static table rows also join painted-equation diagnostic enumeration. Existing
simple-cell exclusions still apply to merged, fit-text, vertical, nested/rich and
note-table owners; no font/size/document guard was added.

The independent mixed/plain controls also exposed a footer table 38.664pt too low.
The complete measured block now shares bottom anchoring and existing row/image
translation; table top edges participate in the footer keep-out. Sixteen Word
anchors have maximum X/Y errors 0.024pt/0.140pt. All eight source equations, four
image placements and page fields survive both presets; the twelve-page portable
control verifies actual counters across 9/10. All four page pairs were reviewed
through exact preset raster reuse, with no visible omission or serious
pagination/overlap/clipping. Four numerical visual gates remain failed, and the
generic dynamic-field screens remain explicitly unresolved; prose/simple-math
deficits are zero.

Frozen f684dc29 / root 8bf5354e passes 32 exact Windows methods in one invocation,
15 Ubuntu methods and five genuine parent behavior regressions. Eight affected
fixed DOCX PDFs retain bytes; PPTX production is unchanged and the held-out deck
passes transfer smoke. Transfer verifies 17 assets/four PDFs with no root rebuild
or suite replay. Evidence: `artifacts/composition-c1/static-cell-evidence.json`.

The complete held-out static-cell report preserves every original package part
except its header/footer stories, adding mixed footer math/image/page fields.
It exposed two consequential integration defects beyond the simple controls:
a row's bottom Y was used as its top keep-out, and scaled review tables painted
below the page. The boundary now uses Y + Height. Simple static tables join the
existing complete-frame transform for text, math, images, clips, cells and borders;
printed paragraph admission converts the already resolved footer/note reserve
once. Shared simple-cell eligibility and the inherited physical-coordinate marker
are reused; no new flag, font/size/document guard or quota was added.

The accepted report retains ten source equations, three source image placements
and five navigation rectangles on two pages in each preset. All four page pairs
were reviewed against independent Word: the footer remains visible, section seven
starts on page two, and no serious omission/overlap/clipping is observed. Twelve
static prose anchors have maximum X/Y errors 0.024pt/0.422pt. Both numerical visual
gates still fail. Generic source screens remain unresolved: cached digit-one gaps
match Word, and whole-word rate/time extraction gaps remain; all simple expressions
are present. Typography, spacing, note-image alignment and annotation geometry
remain cosmetic residuals.

Frozen d0cc21fe / root 5b1af0af passes 29 exact Windows methods in one invocation;
22 Ubuntu checks pass and seven Office-font checks skip. The portable complete
report normalizes mathematical letters to the pinned face's supported ordinary
alphabet; the Windows Office control uses the original document. Two regressions
fail against the preceding artifact. Eight fixed DOCX PDFs and four prior static
controls remain exact. Transfer verifies 17 assets/four PDFs without a root rebuild
or suite replay. An intermediate keep-out artifact remains as the regression
parent; final promotion followed review of all presets. Evidence:
`artifacts/composition-c1/static-table-print-evidence.json` and
`artifacts/composition-c1/composed-static-cell-evidence.json`.

Ordered inline painting now rejoins adjacent word atoms within the same source
owner and identical font/spacing state. Wrapping retains word atoms and source
extents; painting retains glyph-run continuity with linear string assembly. This
improves all eight portfolio first-page Office comparisons (MAE down 0.04146 Final /
0.01897 review), with later-page metrics unchanged, and all four complete-report
page pairs. Eight classic failure sets remain unchanged; small cosmetic wrapping
metric tradeoffs are retained. All existing numerical thresholds remain.

A broader single-equation eligibility prototype is rejected: it discarded 0.348pt
of script occupied leading and added a failed visual gate. The original eligibility
and script/fraction PDFs remain unchanged; no script/font/size exception is added.
The next investigation is the shared prepared-frame versus ink-fit contract.

Frozen 94f6f425 / root 8994b85a passes 36 Windows methods in one invocation and
12 Ubuntu checks; 24 native/Office-font controls skip without their prerequisites.
Two missing test prerequisites reproduced parent failures and are now explicit.
Text-show anchor extraction accounts for leading spaces, preserving the original
Office values/tolerances. The test-only correction retains qualified library
SHA-256 44A76742D209D729F993E9BC692E9D53EAA81313178F942A29C9AF3499BBA354;
all completed conversion/comparison work is reused.

Ten portfolio outcomes retain zero prose/simple-math presence deficits; eight DOCX
gates fail and both decks pass. The complete report retains two pages in both
presets, all ten source equations, three source image placements and five navigation
rectangles; all four page pairs were reviewed without serious omission/overlap/
clipping. Its generic dynamic-field/rate-time screens remain unresolved, and both
numerical gates fail. Transfer verifies 17 assets/four PDFs with no root build or
suite replay. No speedup is claimed. Evidence:
`artifacts/composition-c1/ordered-glyph-run-evidence.json`.

### Prepared occupied frames (C26)

Ordered wrapping no longer substitutes ink bounds for a prepared equation frame
when the ink fits the ordinary text strut. It unions the complete occupied frame
with that strut once, preserving leading without a font, size or script exception.
Two new regressions fail on the parent and pass on frozen 23d10821 / root d72e938d.
Eight independent Word baseline gaps are within 0.15pt in both presets; all four
new script page pairs were reviewed without observed omission, overlap or clipping.

All ten fixed portfolio and both complete-report PDFs retain their accepted bytes.
The eight classic controls keep their failed gate sets. One wrapped literal
improves MAE by 0.03090; another loses 0.04248 foreground recall and adds 0.02014
MAE despite a closer formula baseline. Its following-prose line-height mismatch
remains. Script comparisons also retain cosmetic failures. Thresholds and
single-equation eligibility are unchanged; marker migration remains deferred.

Qualification runs 33 Windows methods once and 8 Ubuntu checks (25 font-specific
skips). Transfer verifies 17 assets and four exact PDFs without a root rebuild or
suite replay. The preceding test-only CI repair uses visible comment/tab text
anchors, preserving Office bounds and tolerances; Windows and Ubuntu CI are green
through 1ba9d797. Evidence: `artifacts/composition-c1/inline-occupied-frame-evidence.json`
and `artifacts/composition-c1/glyph-run-anchor-repair-evidence.json`.

### Ordinary-word emergency wrapping (C28)

Default/on paragraphs now emergency-break ordinary words wider than their text
frame, using the existing bounded Unicode-safe search. Protected spans retain
their existing fallback. Independent Word controls cover default/on/off in both
presets: default/on recall improves from 0.377 to 0.828 and MAE from 0.278 to 0.094;
explicit-off metrics remain unchanged. Relative line gaps are within 0.15pt.
Candidate `formu/la` versus Word `form/ula` remains a character-break residual.
All six visual gates still fail recall/SSIM; thresholds remain unchanged.

All ten fixed portfolio and both complete-report PDFs retain accepted bytes.
A formerly omitted narrow-column script now survives, as in Word, but sits one
line too high because of differing text breaks. This significant geometry gap
remains open; broad single-equation migration is rejected and remains deferred.
No observed omission, overlap or clipping in reviewed new page pairs; identical
rasters reuse reviews across presets/settings.

Frozen a6f1c557 / root d4bb5b02 passes 81 selected Windows methods once and 54 Ubuntu
checks (27 font-specific skips). Two new regressions fail on the accepted parent.
The test-only correction retains the qualified production DLL and tests the
actual whole-formula fallback boundary at an 8pt frame. Transfer checks 17 assets
and four exact PDFs without a root build or suite replay. Evidence:
`artifacts/composition-c1/emergency-word-wrap-evidence.json`.
