# 0.2.0-preview qualification

The bounded workbook/page-ownership milestone is qualified. Publication,
performance optimization and expansion remain deferred.
The library and harness were built once at
`5bcd061e2258e5a9b1f2295bfd58ddf850d83cb9`.
Library SHA-256: `5178EB96E5B3EA4CD99AB3E7551E564D3371A60C013F1D0251C6C32FF8C1E685`.
[Windows/Ubuntu CI](https://github.com/Lokad/OoxPDF/actions/runs/40517500300)
passes at the tool-only caller follow-up source. Both jobs pass the chart
classifier and capability-catalog gates. CI builds are independent of the frozen runtime.

| Gate | Result |
|---|---|
| Release/affected checks | One committed linked-worktree build; 12 affected methods pass on Windows and Ubuntu without skips or repairs. No method repeats. |
| Windows milestone | **3,830 pass, one inherited failure, five skips**; 3,836 unique methods in two disjoint invocations. No prior passing method regresses. |
| Ubuntu milestone | **1,237 pass, seven inherited failures, 85 skips**; 1,329 unique methods in two disjoint invocations. All four new portable methods pass; no inherited failure becomes a skip. |
| Workbook ownership | Two content-type scans and their two unused constants removed. Present/absent/orphan/referenced/alternate-target metadata controls, visibility/date/format semantics, cancellation and existing style/shared-string bounds pass. Referenced malformed metadata still fails; no broad malformed-workbook recovery. |
| Independent Office control | The exact Office-accepted orphan-style input now retains the entire chart, no diagnostic and the identical supported-control PDF. Both native controls agree with Office content tokens; pinned Office control/orphan page rasters are identical. A linked-metadata neighbor retains its previous PDF/diagnostics. No new Office export or rasterization. |
| Classifier ownership | Eighteen Windows classifier invocations pass page composition/reordering/filter/absent-own-plot controls. All six retained single-page outputs remain byte-identical. Both real PDF pages keep an axis pair, and the page-2 200-point move fails its classified gate. Existing PDF inspections are reused. |
| Region-summary caller | ByRegion keys now display PageNumber/RegionIndex, preserving distinct page-owned unassigned regions too. Three focused helper checks pass without replaying classifiers or rebuilding the library. The existing CI classifier step covers the same helper. |
| Representative documents | Nine sources/16 outcomes/32 pages retain exact prior PDFs, diagnostics and semantic results. Each current outcome converted once. Reuse two strict PPTX passes/fourteen strict DOCX failures; thresholds/residuals remain. These reused pixel gates do not use the changed classifiers. |
| API/runtime | 247 public declarations unchanged; IL-only, no P/Invoke/imported COM, explicit AllowUnsafeBlocks=false; assembly/file version 0.2.0.0. |
| Package/transfer | Same package/consumer source pass Windows and Ubuntu. All 24 assets and two integration PDFs match after transfer; no root library build or suite replay. Existing shared 2,613-input snapshot is reused. |

## Conversion behavior and bounded cleanup

Optional embedded-workbook styles and shared strings now follow that workbook's
relationships. Unrelated leftovers no longer activate metadata parsing or erase
an otherwise supported cached chart. Related style/string/date/visibility data,
its input bounds and cancellation stay active. The established workbook model,
range caches and useful inspectors remain; no new quota or ownership framework.

Chart graphics/text classifiers operate on one page at a time before joining
reports. Each page owns its plot boxes, evidence and region indexes. Existing
single-page heuristics and tolerances are unchanged. Region summaries qualify
keys by page/index; their default view remains a document-level kind rollup.

The source changes are in scoped commits. A tool-only caller follow-up at
`7f70882956a8d99cf05abc7739d774c13d5d6143` leaves all production/test source and
frozen runtime assets unchanged. Only its three new summary checks run locally;
the completed classifier/library qualification is reused. CI independently builds
and checks the final source. No harness repair or extra library build was needed.

The preceding [qualification record](https://github.com/Lokad/OoxPDF/blob/25f560fb/docs/PreviewQualification.md)
retains its runtime identities, inherited results and picture/table controls.
No new chart-color, document-shape or font-specific rendering guard is introduced.

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

## Evidence and workflow

Ignored evidence: artifacts/preview-chart-ownership-20261027/qualification.json.
It binds the frozen build, exact disjoint method reports, Office-backed chart
controls, unchanged neighbor/portfolio, classifier baselines and moved-page
negative control, same-package consumers, transfer, caller checks, catalog and CI.

Package: artifacts/nuget/preview-020-chart-ownership-20261027/Lokad.OoxPdf.0.2.0-preview.nupkg.
SHA-256: `61E7E1E4AFF78AFD4AE6C554466B2142865CFB78D81DE7AF60E94398D9E6175A`.
Safe managed dependency-free execution, cancellation, meaningful bounds, coarse
host admission/refunds, bounded caches/staging, deterministic cleanup, atomic
publication and PDF validity remain. Unresolved layout/history/chart guards are
preserved. Publication, performance work and broader expansion remain deferred.

The last full footprint remains 6.778 GB/935,126 files before recent batches,
above the loose 5 GB target. This milestone is about 50 MB and reuses the existing
public-input snapshot. Useful artifacts and evidence remain; no duplicate input
snapshot, full workspace rescan or blanket deletion.
