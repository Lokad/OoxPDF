# 0.2.0-preview qualification

The bounded missing-picture recovery/reference-harness milestone is qualified.
Publication, performance optimization and expansion remain deferred. Library
build source: `fe248aa1b31ab4999410894e3569251799ee4775`; final test-only source:
`f1ec324e08d41f7dcb7a4b8be885a65774703aae`. Library SHA-256:
`391136E123FB6F2FC7921ABE5ECB415D9BD90A2D9CBD7227FAFD8E3B9E61642B`.
[Windows/Ubuntu CI](https://github.com/Lokad/OoxPDF/actions/runs/40351135252)
passes at the final harness source. CI builds are independent of the frozen runtime.

| Gate | Result |
|---|---|
| Release/affected checks | One committed linked-worktree library build; 28 exact neighboring methods. Accepted results: Windows 28 pass; Ubuntu 27 pass/one installed-font skip. Four new portable recovery methods pass both hosts. |
| Windows milestone | **3,810 pass, one inherited failure, five skips**; 3,816 unique methods/3,817 attempts in three invocations. |
| Ubuntu milestone | **1,203 pass, seven inherited failures, 74 skips**; 1,284 unique methods/1,285 attempts. Twenty-seven methods are newly covered relative to the preceding Ubuntu milestone. |
| Office controls | Missing-picture following text retains the present/corrupt control position; strict exit is now 3 with IMAGE_MISSING_PART. Present/corrupt PDFs are byte-identical to the preceding runtime. The inherited 2.365051 pt Office baseline residual remains. Three pinned Word Final references reused; no new export or visual rasterization. |
| Test ownership | Full reference workflow passes using the frozen CLI. All six registered test-cache paths are absent afterward. Live cleanup blocks also pass isolated success/failure and refusal probes while preserving an independent reference. The ineffective timeout parameter is removed; fixed supervision deadlines remain. |
| Representative documents | Nine sources/16 outcomes/32 pages retain exact prior PDFs, diagnostics and semantic results. Each outcome converted once. |
| Independent visual gates | Reused by identity: two strict PPTX passes/fourteen strict DOCX failures. Thresholds and residuals remain; manifest validation is not fidelity evidence. |
| API/runtime | 247 public declarations unchanged; IL-only, no P/Invoke/imported COM, explicit AllowUnsafeBlocks=false; assembly/file version 0.2.0.0. |
| Package/transfer | Same dependency-free package and consumer source pass Windows/Ubuntu. All 24 runtime assets and two integration PDFs match after transfer; no root library build or broad suite replay. |

## Conversion behavior and bounded maintenance

An unresolved supported DOCX picture retains its authored frame through the
existing paragraph/container layout. Its unavailable bytes are represented
explicitly; painting reports IMAGE_MISSING_PART rather than silently dropping
both paint and space. Existing failure caching suppresses repeated asset errors;
unpainted pictures emit no picture link. SVG previews and non-image textboxes
retain their existing paths. This does not reconstruct missing picture pixels.
Caller guidance records the intentional strict-exit behavior change.

The first container check incorrectly expected prose links to disappear with a
picture. It now preserves those independent links while the dedicated test checks
picture-link omission. One failed method per platform is rerun on a test-only
rebuild; no successful method is replayed and production remains identical. The
tracked join verifies 21 unchanged runtime assets and all 2,599 public inputs.
Initial failed reports are retained. All inherited failure details match exactly.

The reference-workflow test owns fresh exact cache identities and removes only
those paths in finally, after stopping its publishers. Existing references cannot
be claimed. A frozen CliPath avoids replaying the library build gate; the original
default build-gate check remains available. No cache framework, quota system,
rendering guard or public API was added or removed.

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

Ignored evidence: artifacts/preview-image-recovery-20261026/qualification.json.
It binds the immutable runtime, exact test sets/joins, original failed attempt,
Office controls, portfolio identity, package consumers, cleanup probes, transfer
and CI. Public fixtures retain the independently exported review inputs unchanged.
The [preceding qualification](https://github.com/Lokad/OoxPDF/blob/219b0676/docs/PreviewQualification.md)
records optional-part ownership/workbook cleanup and its evidence.

Safe managed dependency-free execution, cancellation, meaningful bounds, coarse
host admission/refunds, bounded caches/staging, cleanup, atomic publication and
PDF validity remain. The preserved roadmap stays deferred. The latest full
workspace check is 6.778 GB/935,126 files before recent batches (2026-10-26), above
the loose 5 GB target. This milestone is about 65 MB, plus its shared input
snapshot and 0.4 MB of useful harness evidence. No full rescan or blanket deletion.
