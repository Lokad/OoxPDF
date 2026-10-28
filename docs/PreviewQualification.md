# 0.2.0-preview qualification

## Qualification join checkpoint

Tool source `cba5e93870267201268b1a0d729cb5b3d1f5335c` binds the final artifact's
test assembly to the final supplied report-producing harness and rejects empty
reports before publication. Identical copies remain accepted; earlier passing
methods retain their original harness provenance without being rerun.

Windows metadata checks pass three valid joins and thirteen refusal scenarios;
the preceding helper fails the new unrun-final-harness regression. Four retained
Windows/Ubuntu reports rejoin with identical results and provenance; only the
tool fingerprint changes. These are metadata checks, not new console-test runs.
[Current-source CI](https://github.com/Lokad/OoxPDF/actions/runs/40697316633)
passes both-platform script and broader checks on separate builds.

Evidence: artifacts/preview-tested-harness-20261028/qualification.json. The frozen
library/package below stay unchanged; this slice performs no local library build,
console-test run, package check, conversion, Office export or rasterization.

## Frozen library checkpoint

The bounded diagnostic-provenance/dead-state milestone is qualified. Publication,
performance optimization and feature expansion remain deferred. One linked-worktree
Release library compilation/frozen source: `bc85828093870587b65c3d938894deaf6ba5042e`.
Library SHA-256: `7B2EDB6A081DA18EC1F62B9585FA73E9C44C2D0559B62D9AC3DF3F0D8BA8CA13`.
[Windows/Ubuntu CI](https://github.com/Lokad/OoxPDF/actions/runs/40687441247) passes, including both hosts' PowerShell orchestration.
CI builds are separate from this immutable runtime.

| Gate | Result |
|---|---|
| Diagnostic reuse | Saved diagnostic arrays and cached sidecars must agree before reuse work; copied/source bytes remain pinned through inspection. Fourteen invalid-cache scenarios reject before output/work; four late changes reject before completed report publication. Matching empty/nonempty arrays and equivalent JSON remain reusable. No replacement evidence framework or new report schema. |
| Independent negative/control | The preceding script fails the cleared-warning regression. One real reinspection on the preceding qualified runtime retains all four warnings; cleared/added review copies reject before output. No conversion or library build for this script correction. Original accepted evidence stays unchanged. |
| Internal scene cleanup | Remove the unused Document member and constructor argument only. The builder's live input, slide themes, models and useful inspectors remain. No public API change, rendering rule or measured memory/performance claim. |
| Focused frozen checks | 47 exact scene/theme/inspection methods in one invocation per host: **47 Windows pass; 46 Ubuntu pass/one inherited failure; zero skips**. Every selected outcome/detail matches the preceding artifact. No new test merely asserting property removal. |
| Broader current-source gate | Both-platform CI passes its established broader, package, portable and script checks. Local full DOCX/broad suites were not repeated for this mechanical change. Local Ubuntu has no pwsh; the Linux script gate runs in CI, while local Linux .NET checks still run on the frozen DLL. |
| Independent rendering evidence | Two PPTX/DOCX smoke PDFs are byte-identical to the preceding candidate, with strict exits 0 and no diagnostics. Reuse their pinned Office-control comparisons by identity. No new Office export, portfolio conversion batch, raster replay or threshold change. Other retained visual failures remain failures. |
| API/runtime | 247 public declarations unchanged; IL-only, no P/Invoke/imported COM, explicit AllowUnsafeBlocks=false. Assembly/file 0.2.0.0; informational 0.2.0-preview plus the build revision. No caller migration change. |
| Package/transfer | One package and the same consumer source pass Windows/Ubuntu, including admission/refunds, metadata and caller docs. 24 assets/two smoke PDFs match after transfer. No root build or suite replay. Existing 2,625-input snapshot reused; no added input tree. |

The first fake late-copy test needed absolute-path normalization in its shim;
that attempt remains alongside the corrected result. JoinCompositionPortfolio
continues preserving original row classifications and pinned copies; revised
reinspection verifies the same diagnostic association for joined inputs.

The [preceding full qualification](https://github.com/Lokad/OoxPDF/blob/67622308f01c545dc322d5a51fda6718ea6caf95/docs/PreviewQualification.md)
retains **3,859 Windows passes/one inherited failure/five skips** and **1,285
Ubuntu passes/eight inherited failures/86 skips** on its own frozen runtime.
These are historical counts, not a claim that those full local sets ran on this
DLL. Its 16 portfolio outcomes/32 pages, two strict PPTX passes/fourteen strict
DOCX visual failures and earlier rendering residuals remain recorded. Earlier
cleanup and theme/hidden-picture corrections stay complete; chart-color and
document-shape guards are unchanged. No new Office-fidelity pass count.

## Retained review-36 line-gate residuals

These earlier line-start screens were not replayed for this maintenance milestone.
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

Ignored evidence: artifacts/preview-diagnostic-provenance-20261028/qualification.json.
It binds script regressions/refusals, the retained original failed attempt,
one real cached reinspection, exact scene sets and inherited result, immutable
runtime, smoke identities, managed/API checks, consumers, transfer and CI.

Package: artifacts/nuget/preview-020-diagnostic-provenance-20261028/Lokad.OoxPdf.0.2.0-preview.nupkg.
SHA-256: `44DB2CA434A586D589B41E9BA4750C487152E05515EA65FB08504B77165B333E`.
Safe managed dependency-free execution, meaningful input bounds, cancellation,
coarse host admission/refunds, bounded caches/staging, deterministic cleanup,
atomic publication and PDF validity remain. Publication stays deferred.

The last full workspace measurement remains 8.162 GB/1,013,013 files before the
preceding bounded obsolete-build cleanup, above the loose 5 GB target. This slice
reuses the shared input tree and performs no full rescan or cleanup; footprint.json
measures this milestone only. Preserve useful inputs/evidence and Git history.
