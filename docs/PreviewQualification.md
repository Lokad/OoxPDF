# 0.2.0-preview qualification

The bounded table-paragraph-default/image-identity milestone is qualified. Publication, performance optimization and expansion remain deferred.
The library was built once at `79a6e2c715976946b73df60bef28570c377b7b68`.
Final harness source: `48a7dc064cf37408270ee25d2ba01e17f6bd9d15`; catalog repair: `4a8d2841`.
Library SHA-256: `1828BCE4B5E87637A8A35A1A1B98C751C52A6941E0EEE896614F7073C94B23B7`.
[Windows/Ubuntu CI](https://github.com/Lokad/OoxPDF/actions/runs/40493687709)
passes at the catalog repair source. CI builds are independent of the frozen runtime.
The superseded source run failed only a stale renamed-test catalog entry; complete
catalog selection now resolves against the unchanged frozen runner.

| Gate | Result |
|---|---|
| Release/affected checks | One committed linked-worktree library build. Twenty-two affected methods: Windows 22 pass; Ubuntu 14 pass/eight Windows-font skips. Two harness-only assertion repairs repeat only the failed painting method per host. All 21 other runtime assets and 2,613 public inputs remain identical. |
| Windows milestone | **3,823 pass, one inherited failure, five skips**; 3,829 unique methods/3,831 attempts in four invocations. No prior passing method regresses. |
| Ubuntu milestone | **1,228 pass, seven inherited failures, 85 skips**; 1,320 unique methods/1,322 attempts. Eight additional skips are explicit native-font preconditions; no inherited failure becomes a skip. |
| Independent PowerPoint controls | All five valid controls agree on bold/color. Three paragraph-default mismatches are corrected. Inherited-style/direct-run PDFs remain byte-identical. Reused Office PDFs; no new Office export/rasterization. The malformed initial run control remains excluded. |
| Neighboring tables | Three fixtures retain exact parent PDFs/diagnostics, including explicit wrapping and font fragmentation. No parent build or raster replay. |
| Image identity | One production/five test selector arguments removed. Existing complete identities, content guard, stable deduplication, CMYK distinctions and cancellation remain; no replacement hook or public API change. |
| Representative documents | Nine sources/16 outcomes/32 pages retain exact prior PDFs, diagnostics and semantic results. Each current outcome converted once. |
| Independent visual gates | Reused by PDF identity: two strict PPTX passes/fourteen strict DOCX failures. Thresholds/residuals remain. Integrity validation of 715 manifests/12 families remains separate from fidelity evidence. |
| API/runtime | 247 public declarations unchanged; IL-only, no P/Invoke/imported COM, explicit AllowUnsafeBlocks=false; assembly/file version 0.2.0.0. |
| Package/transfer | Same package and consumer source pass Windows/Ubuntu. All 24 runtime assets and two integration PDFs match after transfer; no root library build or suite replay. |

## Conversion behavior and bounded cleanup

Authored table paragraph defaults now suppress only the lower-priority table
fallback for bold and fill. Scene parsing and rendering share this rule through
existing raw/cascade ownership. Explicit false clears bold; direct run formatting
still wins. Inherited theme/master defaults cannot suppress styled headers.
No font/size/style-GUID-specific correction or renderer rewrite is introduced.

Image deduplication computes its existing ResourceKey directly. Full digests,
byte equality and cancellation remain. The stale-identity test mutates an already
exposed model array; it demonstrates guard behavior, not a live corruption path.
No ownership layer, cache redesign or measured performance claim follows.

The initial portable painting assertion counted background color, then unrelated
body bold. Both failed reports/artifacts remain. The final test scopes its checks
to the header text object. Only that failed method repeats twice per host; passing
methods and broader remainders run once. Library bytes never change during those
harness repairs. A remaining catalog entry for the intentionally renamed identity
test was corrected without rebuilding the library or runner.

## Retained review-36 line-gate residuals

These earlier screens were not replayed for this bounded table correction.
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

## Evidence and workflow

Ignored evidence: artifacts/preview-table-defaults-20261027/qualification.json.
It binds the original and repaired harnesses, exact method reports, five valid
Office controls, neighboring table identity, unchanged portfolio, same-package
consumers, transfer, catalog repair and CI. The preceding library qualification
remains in Git history and artifacts/preview-slide-membership-20261027/.

Safe managed dependency-free execution, cancellation, meaningful bounds, coarse
host admission/refunds, bounded caches/staging, deterministic cleanup, atomic
publication and PDF validity remain. Unresolved layout/history/chart guards are
preserved. No publication, performance work or broader feature expansion.

The last full footprint remains 6.778 GB/935,126 files before recent batches,
above the loose 5 GB target. This milestone is about 82 MB plus one shared 24 MB
input snapshot. Used failed/qualified artifacts and evidence remain; no duplicate
input snapshot, full rescan or blanket deletion was needed.
