# 0.2.0-preview qualification

The bounded master-theme/hidden-picture milestone is qualified. Publication,
performance optimization and broader expansion remain deferred. One linked-worktree
Release library compilation at `1da2ed01b938b341413f7f94beb429640a6ff787`; final test-only repair/frozen source
`22f3a24c00ee407bd8f7ba084d0b6ba077ce2968`. Library SHA-256:
`EAC6A31D9F48340C4AF6B558A3FAE507D143D665B9ABE1E0F361D40F7F4CAC0C`.
[Windows/Ubuntu CI](https://github.com/Lokad/OoxPDF/actions/runs/40676688524) passes. CI builds are separate from the qualified runtime.

| Gate | Result |
|---|---|
| Release/affected checks | All eight new portable methods pass each host. Windows 35 affected passes; Ubuntu 34 passes/one missing-Cambria-Math skip. One invalid optional-model assertion was removed with identical library/runtime/inputs; only that failed method reruns. Broader sets are disjoint. |
| Windows milestone | **3,859 pass, one inherited failure, five skips**; 3,865 unique methods, 3,866 attempts across three invocations. No prior pass regresses. |
| Ubuntu milestone | **1,285 pass, eight inherited failures, 86 skips**; 1,379 unique methods, 1,380 attempts. Native-font failures remain failures. Newly selected DocxInlineCompositionHiddenRunPreservesMathGeometry requires installed Cambria Math and passes on Windows. |
| Master themes | Two independent PowerPoint controls agree in accent fill and rectangle geometry. Scene/inspection/paint share the actual master-owned theme, with one model per theme part; global scene theme removed. Portable background/font/shared-master/fallback neighbors pass; no new Office multi-theme chart/font parity claim. |
| Hidden inline pictures | Three accepted Word controls agree in visible image count. Hidden picture and link are suppressed; paragraph/source ownership survives. Following text is -.027pt from Word. Ordinary/docPr-hidden control PDFs stay identical, including their existing -2.286pt offset. Style inheritance/false overrides, computed fields, revisions, cells/static stories, cancellation and publication pass. |
| Parent negatives | Both regressions fail on the preceding library: missing expected blue fill; expected zero images, got one. One two-method run, no parent build or missing-input retry. |
| Affected visuals | Corporate-theme and DOCX-image cases pass unchanged 144 DPI metric, page/dimension/diagnostic and rating gates (4/5, 5/5). Cached Office references reused; no new Office export or threshold change. |
| Representative documents | Nine sources/16 outcomes/32 pages retain exact prior PDFs, diagnostics and semantics. Convert each once; reuse two strict PPTX passes/fourteen strict DOCX failures. No portfolio Office/raster replay. |
| API/runtime | 247 public declarations unchanged; IL-only, no P/Invoke/imported COM, explicit AllowUnsafeBlocks=false. Assembly/file 0.2.0.0; informational 0.2.0-preview plus library build revision. No caller migration change. |
| Package/transfer | One package and identical consumer source pass Windows/Ubuntu, including admission/refunds, metadata and caller docs. 24 assets/two smoke PDFs match after transfer. No root build or broad replay. One new 2,625-input snapshot reused for the harness repair. |

The failed source-index assertion assumed every paragraph had the optional ordered
inline representation. The source-index, computed-field and rendering assertions
remain; no production behavior changed for the repair. BuildProjectReferences=false
and the report join verify 21 unchanged runtime assets and all 2,625 public inputs.
Original failed attempts remain; successful methods are not repeated. Two existing
nullable warnings remain unrelated test code.

The candidate's following-text shift is 50.541pt versus Word's 52.800pt. Retain the
relative 2.259pt residual and the unchanged visible-control offset; do not fit a
spacing correction. The earlier .024pt PowerPoint follow-text size detail remains.
Common inline image admission follows its run; independent anchored/VML/equation
and multi-theme chart/font Office comparisons remain outside this milestone.

The [preceding qualification record](https://github.com/Lokad/OoxPDF/blob/a17988092a5804e2b8343975b66dc2f1d189c23d/docs/PreviewQualification.md)
retains bullet-choice/visible-diagnostic evidence and earlier identities. Earlier
cleanup stays complete; chart-color and document-shape guards are unchanged.
715 manifests validate; that count is separate from Office visual passes.

## Retained review-36 line-gate residuals

These earlier line-start screens were not replayed for this theme/hidden-picture correction.
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

Ignored evidence: artifacts/preview-theme-visibility-20261028/qualification.json.
It binds initial/final harness identities, exact joined/disjoint test sets, pinned
Office controls, parent negatives, unchanged portfolio, visual gates, managed/API
checks, consumers, transfer, catalog and CI. No new validation framework.

Package: artifacts\nuget\preview-020-theme-visibility-20261028\Lokad.OoxPdf.0.2.0-preview.nupkg.
SHA-256: `C6F508638CB8238D6F40E7044833FE5A5B32D263E4F9AEC70EAE68395B7A6017`.
Safe managed dependency-free execution, cancellation, meaningful bounds, coarse
host admission/refunds, bounded caches/staging, deterministic cleanup, atomic
publication and PDF validity remain. Publication/performance/expansion deferred.

Last full footprint, 2026-10-28: 8.162 GB/1,013,013 files; zero errors, before a
bounded removal of 15 obsolete bin/obj directories from two clean inactive
worktrees (248 files, 96.6 MB logical bytes). Inputs, probe evidence, Git and
qualified artifacts remain. Hard links prevent inferring physical savings from
that logical count; the workspace remains above the loose 5 GB target.
footprint.json measures this milestone only; the shared input tree is pinned once
and reused for the repair. Preserve useful inputs/evidence during maintenance.
