# 0.2.0-preview qualification

The bounded font-failure/diagnostic milestone is qualified. Publication,
performance optimization and broader expansion remain deferred. The accepted
library and harness source is `f0cc85784833d9ca58cae8255e9081b7624006ba`, SHA-256
`0A769DC8AAA40B01D0654BCFB8577723019CDCA94E6160D0247B4BB59A06345C`.
[Windows/Ubuntu CI passed](https://github.com/Lokad/OoxPDF/actions/runs/40314971817)
at the accepted source. CI builds are independent of the local frozen artifact.

| Gate | Result |
|---|---|
| Release/affected checks | One committed linked-worktree artifact; all 27 affected methods pass both platforms. Five new portable methods cover fatal limits, ordinary font-I/O recovery, cancellation and effective visibility. Two inherited nullable test warnings remain. |
| Windows milestone | **3,796 pass, one inherited failure, five skips**; 3,802 unique methods in two disjoint invocations. |
| Ubuntu milestone | **1,123 pass, seven inherited failures, 69 skips**; 1,199 unique methods. No prior pass regressed. |
| Original failure probes | Six public-API controls: PPTX/DOCX limit refusals and PPTX cancellation preserve the original exception and destination. Ordinary I/O recovery and successful controls retain identical PDFs. |
| Hidden-text diagnostics | Four original CLI probes: hidden missing glyphs now return strict exit 0; visible missing glyphs still warn/return 3. All four PDF hashes are unchanged, including the hidden case's existing font data. |
| Representative documents | Nine sources/16 outcomes/32 pages have exact prior PDFs, diagnostics and semantic results. Each outcome converted once; no new Office exports or rasterizations. |
| Independent portfolio gates | Reused by identity: two strict PPTX passes/fourteen strict DOCX failures. Existing visual thresholds and the retained note recall failure remain unchanged. |
| API/runtime | All 247 public declarations unchanged; IL-only, no P/Invoke/imported COM, explicit AllowUnsafeBlocks=false; assembly/file version 0.2.0.0. |
| Package/transfer | Same dependency-free package and consumer source pass Windows/Ubuntu. All 24 root assets, a corrected hidden-text strict DOCX and a representative PPTX match without root rebuild or suite replay. |

## Conversion behavior and bounded maintenance

PPTX font measurement excludes OoxPdfLimitExceededException from ordinary font-I/O
recovery, following the existing fatal-limit policy. A custom resolver's refusal
now aborts conversion rather than becoming a successful PDF publication. File and
stream tests preserve exception identity, destinations and caller-owned streams;
cancellation and ordinary I/O fallback retain their previous behavior.

DOCX missing-glyph reporting uses the same effective Hidden flag as inline layout.
It keeps warnings for visible missing characters and explicit vanish=false, while
hidden text retains source metadata without claiming that an invisible character
paints a question mark. Final/AllMarkup controls pass. Font preparation, source
inspectors, paragraph struts, bounded caches and recovery ownership remain; no
writer, quota, callback or resource-admission framework was added.

One successful Release build serves the affected set, disjoint broader remainders,
ten original review probes, package consumers and verified transfer. No failed
build, harness repair, parent build or accepted-method replay was required. The
original failed review outputs remain separate. The missing-glyph correction
changes diagnostics rather than PDF bytes; it does not optimize retained font data.

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

Ignored evidence: `artifacts/preview-font-recovery-20261026/qualification.json`;
artifact: `candidate-final/`. Package:
`artifacts/nuget/preview-020-font-recovery-20261026/`.
The exact method selections, isolated temporary directories and source/runtime
hashes are retained. Existing Office references, unchanged comparisons and the
715-case/12-family manifest validation are reused. The
[preceding qualification](https://github.com/Lokad/OoxPDF/blob/d2f5d6d8/docs/PreviewQualification.md)
retains the font-alias/font-slot controls, earlier maintenance and residual gates.

Safe managed dependency-free execution, cancellation, meaningful bounds, coarse
admission/refunds, bounded caches/staging, cleanup, atomic publication and PDF
validity remain. The preserved roadmap stays deferred. The latest full workspace
check remains **6.778 GB/935,126 files**, zero errors (2026-10-26), above the loose
5 GB target. This small batch is measured separately; no repeat full scan or
unverified deletion was needed. Preserve inputs and useful evidence.
