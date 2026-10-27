# 0.2.0-preview qualification

The bounded PPTX slide-membership/reference-metadata milestone is qualified. Publication, performance optimization and expansion remain deferred.
The library was built once at `db235b9b59b20a843341b1b258884386f30af305`.
Final harness source: `1020e018a664457113373731ef73f95930d7b0a4`.
Library SHA-256: `81F4A70CBD3F63C8A1871EC2CD13153BEDCCF09D0501C569A0E1DD983894542A`.
[Windows/Ubuntu CI](https://github.com/Lokad/OoxPDF/actions/runs/40480830373)
passes at the final harness source. CI builds are independent of the frozen runtime.

| Gate | Result |
|---|---|
| Release/affected checks | One committed linked-worktree library build. Eleven exact affected methods pass on both hosts. A read-only temporary-fixture repair rebuilds only the test harness; only its two failed methods repeat per host. All 21 other runtime assets and 2,608 public inputs are identical. |
| Windows milestone | **3,819 pass, one inherited failure, five skips**; 3,825 unique methods/3,827 attempts in three invocations. No prior passing method regresses. |
| Ubuntu milestone | **1,224 pass, seven inherited failures, 77 skips**; 1,308 unique methods/1,310 attempts in three invocations. Inherited failures retain their outcomes/messages. |
| Independent PowerPoint controls | Two visible-list controls retain byte-identical PDFs. Both all-hidden controls fail without PDF, matching PowerPoint's refusal to export them. Four pinned inputs; no new Office export/rasterization. |
| Reference-worker mechanics | 13 supervision, eight exception-stage, ten timeout-stage and four raster-forwarding checks pass. Current raw metadata survives failure/timeout; no-status runs clear stale metadata. Unrelated evidence and process ownership remain intact. No Office/build. |
| Representative documents | Nine sources/16 outcomes/32 pages retain exact previous PDFs, diagnostics and semantic results. Each outcome converted once. |
| Independent visual gates | Reused by PDF identity: two strict PPTX passes/fourteen strict DOCX failures. Thresholds and residuals remain unchanged; 715 validated manifests/12 families are integrity evidence, not a fidelity pass count. |
| API/runtime | 247 public declarations unchanged; IL-only, no P/Invoke/imported COM, explicit AllowUnsafeBlocks=false; assembly/file version 0.2.0.0. |
| Package/transfer | Same package and consumer source pass Windows/Ubuntu. All 24 runtime assets and two integration PDFs match after transfer; no root library build or suite replay. |

## Conversion behavior and bounded maintenance

An explicit PPTX slide list now remains authoritative after hidden slides are
filtered. Relationships alone cannot make an unlisted slide printable. An absent
list still permits existing relationship-based recovery; normal ordering, hidden
slide links and presentation-part discovery retain their behavior. Explicit empty,
unresolved and all-hidden lists produce the existing no-page failure. Cancellation
and failed conversion preserve file publication. No blank-page policy, new option,
public API declaration or rendering guard is introduced.

The reference supervisor clears its generated worker-status snapshot at the start
of a run and copies the current raw status before owned worker cleanup, including
failure and timeout paths when status exists. The supervisor status remains the
run outcome; for example, raw success can accompany supervisor PDF-verification
failure. Unrelated reference output is preserved. This fixes evidence loss/staleness;
no false cache import or fidelity pass was demonstrated.

The initial two fixture-mutation checks copied immutable inputs with their read-only
flag. Only their temporary copies are made writable. Original failure reports and
both frozen harnesses remain; the tracked report join verifies unchanged production
source, runtime assets, inputs and exact affected coverage. No passing check repeats.

## Tooling follow-up: baseline gates and build intent

The review-36 tooling fixes at `a262db9c` reuse the qualified library, all 24
runtime assets, source/input tree and package unchanged. There is no new local
library/test/tool build, broad suite, portfolio or package qualification.
Eighty-three focused line-start controls pass. The existing family-binding command
now covers the real case/inspector boundary, ten inspection capabilities and twenty
stubbed execution calls. The official PPTX whitespace case's automatic gates pass with actual
prebuilt tools; its candidate PDF matches the focused comparison and both inspection
records say BuildRequested=false. Neighboring text/markup and manifest checks pass.

An explicit maxTextLineStartYDelta now enforces Y separately from line grouping;
X-only cases retain their meaning. Ten affected line-start gates report six passes
and four failures. Calibri wrapping and compact before-spacing newly fail their
existing 0.25pt limits (about 0.26-0.28pt baseline differences). Font-table alternate
wrapping retains X failures about 1.21pt/64.86pt; the boundary-invariance PPTX retains
two operation-count mismatches. These results preserve thresholds and do not claim
overall visual fidelity. Nine Office references are reused; one uncached X-only
neighbor needed a PowerPoint 16 export. No renderer change or fitted correction.
Evidence: artifacts/preview-qualification-gates-20261027/qualification.json.
[Tooling CI](https://github.com/Lokad/OoxPDF/actions/runs/40486283787) passes on
Windows and Ubuntu. Publication remains deferred.

## Inherited failures and limitations

**DocxWordCompatibleBalloonBodyWrapsLikeOfficeTitles** retains its local failure:
`InvalidOperationException: The second body should wrap to a second row like Office; observed Height=12.610000000000001.`
CI may skip its Office-font precondition. It is not counted as a local pass.

The three previously recorded Linux chart-link failures remain:

- `PptxChartTitleExternalHyperlinkEmitsUriAnnotation`
- `PptxChartAxisTitleExternalHyperlinkEmitsUriAnnotation`
- `PptxChartDataLabelCustomTextHyperlinkEmitsUriAnnotation`

Each observes `InvalidOperationException: Sequence contains no elements` with the
default Windows font catalog absent. The preceding qualification already reproduced
them on its parent runtime. That evidence is retained without another parent build
or method replay. They remain test portability debt; no assertion relaxation,
skip substitution or production font policy was added. Existing platform skips
remain separate from failures.

Fourteen strict DOCX visual failures still include significant typography/geometry.
SVG Review Y is 1.436035pt against 1pt; the raw Word footnote annotation discrepancy
is 46.794pt while candidate picture/link agree within .000016pt. Two field/extraction
flags and the earlier CMYK histogram .311828 versus .98 remain unresolved.
No Office parity or general color-management claim follows from this maintenance.

Qualified chart/font/size/document-shape guards, partial rendering and deferred
rollback-accounting/enum-helper concerns remain recorded in the roadmap. No live
rendering guard was replaced or broad rendering rewrite undertaken.

Out-of-range chart dates retain numeric fallback rather than exact Office tick
text. The preceding 72 DPI screen remains needs-review (MAE 1.3170 > 1.2,
recall .89705 < .90, changed ratio .02740 > .025); it was separate from the
manifest's 144 DPI gate. That completed change is not reopened here.

The newly exercised Linux method
`PptxTextFrameVerticalClipDropsLinesWithBaselineOutsideClip` fails with
`InvalidOperationException: Expected first clipped line to remain visible.`
It reproduces identically on the prior qualified artifact. Default Windows-font
glyph inspection is unavailable on this host. The failure remains recorded; it
was not changed into a skip. Other inherited native-font skips remain distinct from failures. The pinned-font
clipping regression remains covered; the new visibility regressions pass on both hosts.

The narrow clip probe still fails its 72 DPI exploratory foreground-recall limit:
.8948106592 against .90 (MAE .03772714; changed-pixel ratio .00073110). The explicit
horizontal overpaint is removed; remaining glyph/underline differences are retained
as needs-review. This screen is separate from the source manifest's 144 DPI gate.
Candidate overflow-line link annotations also differ from both Office controls;
the existing annotation policy is preserved pending independent qualification.

Three newly exercised Ubuntu inspection methods also reproduce on the preceding
qualified artifact without its native Windows font catalog:

- `PptxTextGlyphRunInspectionReportsInterGlyphAdjustmentAggregates` — `InvalidOperationException: Sequence contains no matching element`.
- `PptxNumberedAutofitFramesEmitOfficeCharacterSpacingTextState` — `InvalidOperationException: Expected 3, got 0.`
- `PptxNumberedNoAutofitFramesEmitOfficeCharacterSpacingTextState` — `InvalidOperationException: Expected 3, got 0.`

They bring the explicit Ubuntu inventory to seven failures. Three exact parent
methods were run, with no parent build; candidate failures were not replaced by
skips. Existing Arial/Cambria preconditions remain separate. Rendering through
explicit portable fonts and both package consumers still passes.

The restored-note 144 DPI screen retains foreground recall .9771410174 < .99.
MAE .01545325, changed ratio .00044409 and SSIM .99406542 pass their limits.
Both canonical control PDFs are identical to their pre-fix versions, so this is
existing typography/geometry residual, not an ID-fix regression. Keep needs-review;
no threshold change or formal manifest/agent-rating pass is claimed.

## Evidence and workflow

Ignored evidence: artifacts/preview-slide-membership-20261027/qualification.json.
It binds both frozen harnesses, original/repair/exact reports, independent slide
controls, worker checks, unchanged portfolio, same-package consumers, transfer and
CI. The original failed harness attempts remain. The
[preceding qualification](https://github.com/Lokad/OoxPDF/blob/e3917998/docs/PreviewQualification.md)
records selected-story diagnostics and shared XML ownership; that cleanup stays complete.

Safe managed dependency-free execution, cancellation, meaningful bounds, coarse
host admission/refunds, bounded caches/staging, deterministic cleanup, atomic
publication and PDF validity remain. The section-history header probe remains
unresolved; no schema/normalization or layout guard is replaced by this milestone.

The last full workspace check remains 6.778 GB/935,126 files before recent batches
(2026-10-26), above the loose 5 GB target. This milestone is about 55 MB plus one
23.8 MB shared input snapshot. Both actual harness artifacts and useful evidence
remain. No full rescan, duplicate input snapshot or blanket deletion was needed.
