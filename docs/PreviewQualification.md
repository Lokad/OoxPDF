# 0.2.0-preview qualification

The bounded source-numbering/paragraph-identity milestone is qualified.
Publication, performance optimization and feature expansion remain deferred.
Library build: `aa199ac843b50f2f88cf5237cb94c28830b83146`.
Frozen source after test compilation repair: `7c61a2b3335cfa426aaf4ad267993d49973dac8b`.
Library SHA-256: `CC4AC58A64E3F65133CD7BC42B3E98F7ED4C63EBBD3B69C308B1E68420DBB4FE`.
[Windows/Ubuntu CI](https://github.com/Lokad/OoxPDF/actions/runs/40541741903)
passes. CI builds are separate from this qualified immutable runtime.

| Gate | Result |
|---|---|
| Release/affected checks | One linked-worktree library build. Four test compilation errors corrected before the first freeze; subsequent test/CLI/inspector builds use that DLL without building references. 24 affected methods pass on Windows, 16 on Ubuntu; all eight new portable methods pass. No qualification method repeats. |
| Windows milestone | **3,841 pass, one inherited failure, five skips**; 3,847 unique methods in two disjoint runs. No prior pass regresses. |
| Ubuntu milestone | **1,251 pass, eight inherited failures, 85 skips**; 1,344 unique methods in two disjoint runs. Default-Windows-font neighbors qualify on Windows; the new cases use pinned portable fonts. No failure becomes a skip. |
| Slide numbering | Four independent PowerPoint controls agree in full PDF text, including hidden-slide gaps and firstSlideNum=10. Ordinary control stays byte-identical. Fields in shapes/master/layout/table cells and output diagnostic/navigation indexes pass. Existing order/absence policies and cancellation remain. |
| Paragraph identity | Four portable tests preserve colliding/equal-value distinct owners, deduplicate wrapped same-owner lines, scope comment IDs correctly and separate body/textbox comments. The same collision regression fails on the preceding library (Expected 2, got 1), without a parent build. |
| Representative documents | Nine sources/16 outcomes/32 pages retain exact prior PDFs, diagnostics and semantics; each outcome converts once. Reuse two strict PPTX passes/fourteen strict DOCX failures. Ordinary field and dense-review neighbors are identical. No Office export/raster replay or threshold change. |
| API/runtime | 247 public declarations unchanged; IL-only, no P/Invoke/imported COM, explicit AllowUnsafeBlocks=false; assembly/file version 0.2.0.0. No new caller migration change. |
| Package/transfer | Same package and consumer source pass Windows/Ubuntu. Updated release notes are included. All 24 assets and two integration PDFs match after transfer; no root build or broad-suite replay. Existing 2,613-input snapshot reused. |

The internal hash-collision counterexample is now covered through the full
renderer inspector. It does not establish a natural public-file allocation
sequence or failure frequency. The original placement probe and parent negative
control remain evidence. Parent setup initially omitted its pinned font; that
failed attempt is retained, followed by the correct one-method reproduction with
an independent font copy. Production code has no test hook or forced-hash path.

Source numbering is distinct from output indexing. Hidden slides keep their
place in the authored number sequence, while output diagnostics and links keep
their prior page identities. UInt32 origins are stored/evaluated with sufficient
range; absent/invalid origins keep a default of one. No general field evaluator,
new rendering guard or layout rewrite is introduced.

The preceding [qualification record](https://github.com/Lokad/OoxPDF/blob/18172127cf2d13abc597d1a8d3d83dbcb0f2221f/docs/PreviewQualification.md)
retains admitted-content/style cleanup results and prior runtime identities.
That cleanup and its independent Office controls remain complete.

## Retained review-36 line-gate residuals

These earlier line-start screens were not replayed for this chart ownership correction.
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

The newly selected Ubuntu `PptxSceneBuilderBuildsResolvedNodeLists` also fails
with `InvalidOperationException: Expected Hello, got .` before the changed style
assertions: the glyph inspector cannot resolve its default Windows fonts. One
exact method on the preceding qualified artifact reproduces that first failure,
without a parent build. This brings the current Ubuntu inventory to **eight**
inherited failures. Its Windows check passes, as do all new portable checks.
The failure remains recorded and is not replaced with a skip.

## Evidence and workflow

Ignored evidence: artifacts/preview-source-identity-20261027/qualification.json.
It binds the build/frozen identity, exact disjoint method reports, four pinned
Office numbering controls, parent collision test, unchanged portfolio/neighbors,
managed/API checks, same-package consumers, transfer, catalog and CI. No library
or positive-check replay followed the pre-freeze test compilation correction.

Package: artifacts/nuget/preview-020-source-identity-20261027/Lokad.OoxPdf.0.2.0-preview.nupkg.
SHA-256: `7608CDE2AADE9A5EAD7AA9F084C53D41D02A25A4399A9914E553B6E4BEB9CBC4`.
Safe managed dependency-free execution, cancellation, meaningful bounds, coarse
host admission/refunds, bounded caches/staging, deterministic cleanup, atomic
publication and PDF validity remain. Publication/performance/expansion deferred.

The last full footprint remains 6.778 GB/935,126 files before recent batches,
above the loose 5 GB target. footprint.json measures this milestone only; the
shared input snapshot is reused. No full rescan or blanket deletion was needed.
