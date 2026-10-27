# 0.2.0-preview qualification

The bounded picture-recovery/graphics-gate milestone is qualified.
Publication, performance optimization and expansion remain deferred.
The library was built once at `304cf03e92f60fe901190b85bdfef4685416d575`.
Final harness source: `9e940f42b6df842a8abb3fc65ceb93f88bdfab73`.
Library SHA-256: `550407F3BB5338DD365A92E75BF975BA3B451A990D332D2E4DD5A34DE820B324`.
[Windows/Ubuntu CI](https://github.com/Lokad/OoxPDF/actions/runs/40506944038)
passes at the final harness source. Both jobs pass the graphics comparison and
capability-catalog gates. CI builds are independent of the frozen runtime.

| Gate | Result |
|---|---|
| Release/affected checks | One committed linked-worktree library build. Twenty-four affected methods: Windows 24 pass; Ubuntu 16 pass/eight existing native-font skips. One harness-only repair repeats only the failed stream-input method per host. All 21 other assets and 2,613 public inputs remain identical. |
| Windows milestone | **3,826 pass, one inherited failure, five skips**; 3,832 unique methods/3,833 attempts in three invocations. No prior passing method regresses. |
| Ubuntu milestone | **1,232 pass, seven inherited failures, 85 skips**; 1,324 unique methods/1,325 attempts. New recovery and source-frame checks pass; no inherited failure becomes a skip. |
| Picture recovery | Three frozen CLI controls retain exact prior PDFs. Present pictures remain strict exit 0; missing relationship and missing part now both report errors/exit 3 with distinct codes. Source part/slide metadata, original observer failure, cancellation and atomic publication checks pass. |
| Graphics comparison | Forty-five direct Windows controls pass, including page moves, page-local reordering, filtering, zero/nonfinite/missing geometry and explicit/optional hashes. Three real moved-page modes correctly fail using retained PDF inspections. Local Ubuntu has no PowerShell; its script gate is covered by CI. Existing chart classifier checks pass. |
| Table cleanup | One forwarding shim removed; the old table XML-versus-scene self-comparison becomes source-content/declared-frame assertions with a pinned font. Other independent XML agreement tests remain. Five table PDFs are unchanged and agree with existing independent PowerPoint bold/color controls. No new Office export/rasterization. |
| Representative documents | Nine sources/16 outcomes/32 pages retain exact previous PDFs, diagnostics and semantic results. Each current outcome converted once. Reuse two strict PPTX passes/fourteen strict DOCX failures; thresholds/residuals remain. Prior 715-manifest/12-family integrity validation is separate from fidelity and is reused with unchanged inputs. |
| API/runtime | 247 public declarations unchanged; IL-only, no P/Invoke/imported COM, explicit AllowUnsafeBlocks=false; assembly/file version 0.2.0.0. |
| Package/transfer | Same package/consumer source pass Windows and Ubuntu. All 24 assets and two integration PDFs match after transfer; no root library build or suite replay. |

## Conversion behavior and bounded cleanup

PPTX pictures without a resolvable embedded image relationship emit
`IMAGE_MISSING_RELATIONSHIP` with the source story part and slide. Missing target
parts retain `IMAGE_MISSING_PART`; graceful conversion and existing pixel output
remain. Caller diagnostics and CLI strict status now expose the omission.

Graphics matching stays within each PDF page and reports both page numbers.
Selected records require finite bounds and positive integer page numbers;
explicit text-hash matching requires both hashes. Zero geometry and disabled
optional hash checks remain valid. These are direct checks in the existing
comparator, without another PDF parser, schema framework or renderer rewrite.

The first stream regression omitted its required InputKind option. That failed
attempt remains in both platforms' reports. One test-only repair adds Pptx and
reruns only the failed method per host. Passing methods and disjoint broader
remainders run once; the library is never rebuilt. The superseded CI run was
cancelled after the harness repair to avoid duplicate CI work.

Prior table-default and image-identity cleanup remains complete. The last
[qualification record](https://github.com/Lokad/OoxPDF/blob/c54c67ba/docs/PreviewQualification.md)
retains its independent controls and original runtime/harness identities.
No visual thresholds or rendering guards are generalized by this milestone.

## Retained review-36 line-gate residuals

These earlier screens were not replayed for this bounded recovery/tooling correction.
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

- `PptxTextGlyphRunInspectionReportsInterGlyphAdjustmentAggregates` â€” `InvalidOperationException: Sequence contains no matching element`.
- `PptxNumberedAutofitFramesEmitOfficeCharacterSpacingTextState` â€” `InvalidOperationException: Expected 3, got 0.`
- `PptxNumberedNoAutofitFramesEmitOfficeCharacterSpacingTextState` â€” `InvalidOperationException: Expected 3, got 0.`

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

Ignored evidence: artifacts/preview-picture-gates-20261027/qualification.json.
It binds the original/repaired harnesses, exact method reports, three picture
controls, five independent Office table controls, real PDF comparison probes,
unchanged portfolio, same-package consumers, transfer, catalog and CI. Prior
qualified artifacts/references remain available; publication is deferred.

Package: artifacts/nuget/preview-020-picture-gates-20261027/Lokad.OoxPdf.0.2.0-preview.nupkg.
SHA-256: `FDE4156A75A31648C1E2ED27AAD6B10136D15C0B7A2E7DADD1AFCCD5E63D0A7A`.
Safe managed dependency-free execution, cancellation, meaningful bounds, coarse
host admission/refunds, bounded caches/staging, deterministic cleanup, atomic
publication and PDF validity remain. Unresolved layout/history/chart guards are
preserved. No performance work or broader feature expansion.

The last full footprint remains 6.778 GB/935,126 files before recent batches,
above the loose 5 GB target. This milestone is about 66 MB and reuses the existing
shared public-input snapshot. Failed/qualified artifacts and useful evidence
remain; no duplicate input snapshot or full workspace rescan.
