# 0.2.0 qualification

Release preparation sets package/informational version to **0.2.0** and retains
assembly/file version **0.2.0.0**. Renderer C# is unchanged from the qualified
preview below; its Office references, thresholds and residuals remain applicable.
The package includes relative documentation targets, and its smoke check verifies
all shipped Markdown links. Build/runtime identities, API and metadata checks,
Windows/Linux package consumers and smoke output identities are recorded in
artifacts/release-020-20261029/qualification.json. Publication is deferred.

## Retained preview renderer qualification

The bounded PPTX inheritance milestone is qualified. Publication, performance
optimization and broader feature expansion remain deferred. Production source:
`4093ae541acdbe4fb3bb9e55ece291b53a265c8d`. Final artifact/test-harness source:
`4be66e0c62a8fe92ede1852f81f0e4683110a472`; its assertion-only repair reused the same library.
Library SHA-256: `54B3C6167FEA79CA56656C2616AEC8DE0652F509D2B8F2AF8C58EDD222967999`.
[Windows/Ubuntu CI](https://github.com/Lokad/OoxPDF/actions/runs/40712648160) passes on separate builds.

| Gate | Result |
|---|---|
| Background ownership | Paint one effective supported solid background instead of layering master/layout/slide definitions. The transparent override now matches Office's light blue; owner fallback and opaque/default controls remain covered. No new background kind or fitted alpha/color rule. |
| Master artwork | Both layout and slide must allow master artwork. The shared admission decision reaches text/resources, painting, links and diagnostics; consumed themes/styles and compatibility warnings remain. Four independent Office permission controls agree. |
| Focused checks | **19 Windows passes; 18 Ubuntu passes/one declared installed-Arial inspection skip**. The initial Windows set had one assertion error: white is emitted as grayscale. Only that method reran on a repaired harness with identical library/runtime/input pins; initial passes and failure history remain. Both new Office regressions fail against the preceding DLL with no parent build. |
| Independent Office comparisons | Three retained PowerPoint 16 decks/six pages pass unchanged numerical thresholds: maximum MAE .103414 < 1.2; minimum foreground recall .956931 > .90; maximum changed ratio .002193 < .025. Labels and rectangle counts/bounds independently agree. These are bounded control results, not overall Office parity or a formal agent-rating count. |
| Raw extraction residual | Three original measurement rows retain substring flags and classifications: space-joining text operations changes master-theme to master - theme. Independent PDF extraction recovers every expected label from both candidate and Office PDFs. No raw row or threshold was rewritten. |
| API/runtime | 247 public declarations unchanged; IL-only, no P/Invoke/imported COM, explicit AllowUnsafeBlocks=false. Assembly/file 0.2.0.0; informational 0.2.0-preview plus production revision 4093ae54. No caller migration change. |
| Package and transfer | One package/same consumer source pass Windows and Ubuntu, including admission/refunds, metadata and caller docs. 24 transferred assets/two smoke PDFs match. The ordinary PPTX/DOCX smoke PDFs also match the preceding candidate, retaining their Office comparisons by identity. No root rebuild, new Office export or local full-suite replay. |
| Broader gate | Current-source CI passes the established broader, package, portable and script checks. It is separate from the frozen runtime. Historical inherited failures and approximate rendering remain listed below. |

One local library compilation and a test-only repair build; final frozen assets
retain identical production bytes. The three new public controls and their Office
expectations are hash-pinned. Raw attempted orchestration failures remain in the
evidence; a corrected parameter invocation completed the comparisons.

The [preceding full qualification](https://github.com/Lokad/OoxPDF/blob/67622308f01c545dc322d5a51fda6718ea6caf95/docs/PreviewQualification.md)
retains **3,859 Windows passes/one inherited failure/five skips** and **1,285
Ubuntu passes/eight inherited failures/86 skips** on its own frozen runtime.
Those historical local suites were not replayed on this DLL. Its 16 portfolio
outcomes/32 pages, two strict PPTX passes/fourteen strict DOCX visual failures,
and earlier rendering residuals remain recorded. No chart-color or document-shape
guard was changed. The [preceding tool/library checkpoint](https://github.com/Lokad/OoxPDF/blob/7dc5a087/docs/PreviewQualification.md)
retains the completed diagnostic, scene-state and final-harness corrections.

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

Ignored evidence: artifacts/preview-pptx-inheritance-20261028/qualification.json.
It binds the initial/repair/parent reports, exact method set and declared skip,
immutable runtime/input pins, independent controls and raw extraction residual,
managed/API checks, package consumers, transfer and CI. Earlier review reference
files remain unchanged. Public behavior notes are in the changelog/capabilities.

Package: artifacts/nuget/preview-020-pptx-inheritance-20261028/Lokad.OoxPdf.0.2.0-preview.nupkg.
SHA-256: `2C6CFB76082B48680C6F0218008A84A1318A4012B210771AD4B4D8A5E8CED60D`.
Safe managed dependency-free execution, meaningful input bounds, cancellation,
coarse host admission/refunds, bounded caches/staging, deterministic cleanup,
atomic publication and PDF validity remain. Publication stays deferred.

The last full workspace measure remains 8.162 GB/1,013,013 files before earlier
bounded cleanup, above the loose 5 GB target. This milestone's scoped measurement
is about 60 MB/273 files and excludes shared input links; it is not a full scan
or under-cap claim. Preserve useful inputs/evidence and Git history.
