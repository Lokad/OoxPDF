# 0.2.0-preview qualification

The bounded bullet-choice/visible-diagnostic milestone is qualified. Publication,
performance optimization and feature expansion remain deferred. One linked-worktree
Release library compilation at `6b694b869dfdd5e0b0eef03efd82064ee12c3c80`; final test-only repair/frozen source
`4bd471cf599378abfa523c06cac6a644b406eeb5`. Library SHA-256:
`C467FF01A677F717410810717F8F8527848A4466787834602F0EA96B38B0E188`.
[Windows/Ubuntu CI](https://github.com/Lokad/OoxPDF/actions/runs/40668569967) passes.
CI builds are separate from this qualified immutable runtime.

| Gate | Result |
|---|---|
| Release/affected checks | 36 affected methods pass on each host; all six new portable methods pass both. One failed compound-test setup was repaired with identical library/runtime/inputs. Keep 35 initial passes per host; no positive method replays. Broader sets are disjoint. |
| Windows milestone | **3,851 pass, one inherited failure, five skips**; 3,857 unique methods, 3,858 attempts across initial/one-method repair/broader runs. No prior pass regresses. |
| Ubuntu milestone | **1,273 pass, eight inherited failures, 85 skips**; 1,366 unique methods, 1,367 attempts. Pinned portable checks pass; native-font failures remain failures. |
| Bullet choices | Eight independent PowerPoint controls agree in nominal size, font family and RGB. Origins stay within the existing one-point criterion; the .024pt PowerPoint follow-text size residual is retained. Shape/table/placeholder and numbering neighbors pass. No fitted rule or style framework. |
| Visible diagnostics | Hidden video/audio/OLE/unknown-frame owners and groups do not warn; visible unsupported content retains warnings. Required background/orphan/global/compatibility diagnostics, cancellation and atomic publication pass. Four review PDFs retain exact paint; only hidden strict exits/diagnostics change. |
| Parent negatives | The two new regressions fail on the preceding library (Expected 24, got 12; Expected 0, got 1) in one exact two-method run. Candidate methods pass. No parent build or missing-input retry. |
| Affected visuals | Existing bullet-style and bullet-size-percent cases pass unchanged 144 DPI numeric, page/dimension/diagnostic and minimum-rating gates; reviewed at 4/5. Their PDFs remain identical to the parent. Cached Office PDFs are reused; missing raster derivatives are generated offline. |
| Representative documents | Nine sources/16 outcomes/32 pages retain exact prior PDFs, diagnostics and semantics. Each outcome converts once; reuse two strict PPTX passes/fourteen strict DOCX failures. No portfolio Office/raster replay or threshold change. |
| API/runtime | 247 public declarations unchanged; IL-only, no P/Invoke/imported COM, explicit AllowUnsafeBlocks=false; assembly/file version 0.2.0.0, informational version 0.2.0-preview plus the library build revision. No new caller migration change. |
| Package/transfer | One package and the same consumer source pass Windows/Ubuntu, including admission/refunds, metadata and caller docs. All 24 assets and two smoke PDFs match after transfer. No root build or broad-suite replay. One new 2,618-input tree includes the public PPTX control and PowerPoint gold. |

The initial compound bullet test passed an eight-shape document to a helper
whose contract requires one selected shape. That precondition failed on both
hosts before the assertions. The test now uses the existing active per-node
rendering path. Only that method was rebuilt/rerun, with BuildProjectReferences=false
and an unchanged library DLL. The join tool checks all 21 non-test runtime
assets and 2,618 inputs; original failed attempts and the 35 retained passes
remain evidence. The two existing nullable warnings are unrelated test code.

The retained .024pt follow-text size detail is separate from the corrected
12pt choice error. No positional correction or broader pixel-parity claim was
introduced. Hidden-media diagnostics now follow the same owner/ancestor policy
as scene admission, while global/part-level warnings retain their distinct scope.

The preceding [qualification record](https://github.com/Lokad/OoxPDF/blob/45371dd7fad764c1d3826f9d185f7eb2fc15112d/docs/PreviewQualification.md)
retains locked-field/workbook cleanup evidence and earlier runtime identities.
Those fixes remain complete. Chart-color and document-shape guards are unchanged.
715 manifests validate; that count is separate from Office visual passes.

## Retained review-36 line-gate residuals

These earlier line-start screens were not replayed for this bullet/visibility correction.
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

Ignored evidence: artifacts/preview-bullet-media-20261028/qualification.json.
It binds the initial/final harness identities and joined reports, exact broader
sets, independent PowerPoint controls, parent negatives, unchanged portfolio,
affected visual gates, managed/API checks, consumers, transfer, catalog and CI.

Package: artifacts/nuget/preview-020-bullet-media-20261028/Lokad.OoxPdf.0.2.0-preview.nupkg.
SHA-256: `AA306EB595584DD995B802C4CC3694326AA89A7DABE3BD6ADC7FD7343DFF8C08`.
Safe managed dependency-free execution, cancellation, meaningful bounds, coarse
host admission/refunds, bounded caches/staging, deterministic cleanup, atomic
publication and PDF validity remain. Publication/performance/expansion deferred.

The last full footprint remains 6.778 GB/935,126 files before recent batches,
above the loose 5 GB target. footprint.json measures this milestone only; its
new input tree is pinned once and reused for the harness repair. No full rescan
or under-cap workspace claim is made.
