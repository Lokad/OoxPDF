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
