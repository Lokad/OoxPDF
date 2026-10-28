# 0.2.0-preview qualification

The bounded locked-field/workbook cleanup is qualified. Publication, performance
optimization and feature expansion remain deferred. One successful linked-worktree
Release library compilation; frozen source: `466d374db1579e14a12f97e1b6b7f04bde524c53`.
Library SHA-256: `427DE0A6301B357AFCFF7AD0A99F5F8650189C57B72C2A127EBC36D8ACC4BA49`.
[Windows/Ubuntu CI](https://github.com/Lokad/OoxPDF/actions/runs/40656921323) passes.
CI builds are separate from this qualified immutable runtime.

| Gate | Result |
|---|---|
| Release/affected checks | 31 affected methods pass on Windows, 30 on Ubuntu. All four new portable methods pass both hosts. Exact affected and broader sets are disjoint; no qualification method repeats. Two existing nullable warnings remain in unrelated test source. |
| Windows milestone | **3,845 pass, one inherited failure, five skips**; 3,851 unique methods in two runs. No prior pass regresses. |
| Ubuntu milestone | **1,261 pass, eight inherited failures, 85 skips**; 1,354 unique methods in two runs. Native-font geometry qualifies on Windows; new rendering tests use pinned portable fonts. No failure becomes a skip. |
| Locked fields | Simple/complex PAGE and NUMPAGES retain locked values 99/88 while unlocked counters remain 1. Empty results, split instructions, stored formatting, independent nested locks and body/cell/header/footer placement pass. Two serialized PDFs match independent Word text. The same main regression fails on the previous library with one exact negative attempt and no parent build. |
| Workbook cleanup | Removes unused calculation/filter types and readers, formula-attribute dictionaries, retained sheet metadata and test-only accessors. Mixed tests keep cached content, styles/dates, scoped names, table ranges and visibility checks. Cached-chart and unlocked header-field PDFs/diagnostics are unchanged. No performance gain was measured. |
| Representative documents | Nine sources/16 outcomes/32 pages retain exact prior PDFs, diagnostics and semantics. Reuse two strict PPTX passes/fourteen strict DOCX failures. All 16 convert once; no Office portfolio/raster replay or threshold change. |
| API/runtime | 247 public declarations unchanged; IL-only, no P/Invoke/imported COM, explicit AllowUnsafeBlocks=false; assembly/file version 0.2.0.0 and informational version 0.2.0-preview plus the frozen revision. No new caller migration change. |
| Package/transfer | One package and the same consumer source pass Windows/Ubuntu, including admission/refunds, version and caller-document checks. All 24 assets and two integration PDFs match after transfer. No root build or broad-suite replay. One new 2,616-input tree includes the two control documents and Word gold. |

Locked fields reuse stored-result handling; nested fields keep independent locks.
An empty locked result remains empty. General unlocked REF evaluation remains
outside the converter's stored-result scope: an exploratory Office control
evaluated it, and that probe is retained outside this gate. The final nested
control locks both the outer REF and inner PAGE. The review's Word reference is
reused; two small additional edge exports establish the empty/nested contract.
These are content checks, not new typography or pixel-parity claims.

The first portfolio attempt stopped after one conversion because the wrong
inspector was supplied. Its PDF/diagnostics/log were retained; fifteen additional
conversions formed a pinned raw batch, and the existing reuse workflow inspected
all sixteen using the pinned PDF inspector. No conversion was repeated. The
first conversion's elapsed time is unavailable; no performance claim uses it.

The preceding [qualification record](https://github.com/Lokad/OoxPDF/blob/53c611ffe507955b94f9fe1414053a71fc275840/docs/PreviewQualification.md)
retains source-numbering and paragraph-identity evidence. That cleanup and all
earlier accepted fixes remain complete. Chart-color and document-shape guards
were not changed. Manifest validation covers 715 cases; it is not an Office
fidelity pass count.

## Retained review-36 line-gate residuals

These earlier line-start screens were not replayed for this field/workbook correction.
Calibri wrapping and compact before-spacing retain their 0.25pt Y failures
(about 0.26-0.28pt). Font-table alternate wrapping retains X failures about
1.21pt/64.86pt; the boundary-invariance PPTX retains two operation-count mismatches.
No thresholds change, overall visual pass or cosmetic reclassification is claimed.
The [preceding tooling qualification](https://github.com/Lokad/OoxPDF/blob/86fbdfb5/docs/PreviewQualification.md)
retains its inputs, independent references and raw results.

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

The previously selected Ubuntu `PptxSceneBuilderBuildsResolvedNodeLists` also fails
with `InvalidOperationException: Expected Hello, got .` before the changed style
assertions: the glyph inspector cannot resolve its default Windows fonts. One
exact method on the preceding qualified artifact reproduces that first failure,
without a parent build. This brings the current Ubuntu inventory to **eight**
inherited failures. Its Windows check passes, as do all new portable checks.
The failure remains recorded and is not replaced with a skip.

## Evidence and workflow

Ignored evidence: artifacts/preview-locked-fields-20261028/qualification.json.
It binds the frozen source/runtime, disjoint host reports, pinned Word controls,
parent negative, unchanged portfolio/neighbors, managed/API checks, package
consumers, transfer, manifest validation and CI. Failed setup evidence is retained.

Package: artifacts/nuget/preview-020-locked-fields-20261028/Lokad.OoxPdf.0.2.0-preview.nupkg.
SHA-256: `118B788DAE312CCE4410C015A5D76431371297A01774981B52689246A761AFE3`.
Safe managed dependency-free execution, cancellation, meaningful bounds, coarse
host admission/refunds, bounded caches/staging, deterministic cleanup, atomic
publication and PDF validity remain. Publication/performance/expansion deferred.

The last full footprint remains 6.778 GB/935,126 files before recent batches,
above the loose 5 GB target. footprint.json measures this milestone only. A new
input tree is pinned once for its added controls; no full rescan or under-cap
workspace claim is made.
