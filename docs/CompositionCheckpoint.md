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
