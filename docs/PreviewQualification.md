# 0.2.0-preview qualification

The bounded active-content milestone is qualified.
Publication, performance optimization and broader feature expansion remain deferred.
The frozen library was built at `313b9a89e16b410eff438eaf51298564cc4a6144`, SHA-256
`BE1EA4A9EC08F009ADC0A0C41972190F0BA111F7B830DDE918103F152619DCC1`.
[Windows/Ubuntu CI passed](https://github.com/Lokad/OoxPDF/actions/runs/39986201510)
at the qualified source. CI builds are independent of this artifact.

| Gate | Result |
|---|---|
| Release/affected checks | One committed artifact. Windows: 24 exact methods pass. Ubuntu: 23 pass, one existing Arial-precondition skip. New section and visibility regressions pass on both hosts. Two inherited nullable test warnings remain. |
| Windows milestone | **3,780 pass, one inherited failure, five skips**; 3,786 unique methods in two disjoint invocations. |
| Ubuntu milestone | **1,041 pass, four inherited failures, 61 skips**; 1,106 unique methods in two disjoint invocations. No previous pass regressed. |
| Current DOCX sections | The historical-properties variant now produces the same portrait PDF as its current-properties control, matching retained Word 16 page geometry. Missing settings, paragraph sections and revision provenance remain covered in all markup modes. |
| Hidden PPTX objects | Hidden text, image/font resources and links are omitted; the hidden-shape 144 DPI raster matches the retained PowerPoint reference exactly. Absent/false/true flags, groups, pictures, table frames and all three scene sources are covered. Visible-control PDF remains identical. |
| Ellipsis guidance | Horizontal shape text still emits its marker. The guide distinguishes this from live diagnosed fallbacks; no diagnostic or public option was removed. |
| Representative documents | Nine sources/16 outcomes/32 pages retain exact PDFs, diagnostics and semantic results. Each outcome converted once. |
| Independent portfolio gates | Reused by input/PDF identity: two strict PPTX passes/fourteen strict DOCX failures. Thresholds are unchanged. Prior clipping/date/color residuals remain explicit below. |
| API/runtime | All 247 public declarations unchanged; IL-only, no P/Invoke/imported COM, explicit AllowUnsafeBlocks=false; assembly/file version 0.2.0.0. |
| Package/transfer | Same dependency-free package and consumer source pass Windows/Ubuntu. Packed guide bytes/local links match. All 24 root assets and two integration PDFs match, without root build or suite replay. |

## Conversion behavior and bounded maintenance

Section selection follows current body/paragraph ownership rather than accepting
the last sectPr anywhere in the XML tree. Revision snapshots remain inspectable;
formatting revision support is still approximate, including in Original view.
No new revision/layout framework or document-shape guard was introduced.

PPTX scene construction excludes explicitly hidden objects before their text,
images, charts or descendants enter resource preflight and painting. Existing
placeholder inheritance, hidden-slide exclusion and master-shape suppression are
preserved. Raw unsupported-feature diagnostic policy was not rewritten.

Five frozen probes reuse independent Word/PowerPoint 16 references. No Office
export or portfolio raster/comparison was repeated. The ambiguous historical-header
exploration proves no separate header defect and was excluded from this milestone.
No newly dead helper was verified, so this change adds no speculative deletion.

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

## Evidence and workflow

Ignored evidence: `artifacts/preview-active-content-20261023/qualification.json`;
artifact: `candidate-final/`; package:
`artifacts/nuget/preview-020-active-content-20261023/`.
One artifact served focused tests, disjoint broader remainders, five original
probes, sixteen portfolio outcomes, both consumers and verified transfer. No
qualified method replay, root rebuild or parent rebuild was needed.

The initial portfolio invocation incorrectly selected the DOCX inspector for a
PDF. That failed invocation is retained. Its successful first conversion was
reused with the PDF inspector; remaining outcomes were measured and joined with
existing commands. No conversion was repeated and no timing is claimed for that
reused outcome. This is an orchestration correction, not a library failure.

Unchanged 715-case/12-family manifest validation, existing Office comparisons and
parent evidence for the inherited failures were reused. The
[preceding qualification](https://github.com/Lokad/OoxPDF/blob/5beb7cb3/docs/PreviewQualification.md)
retains clipping/comparator work, earlier maintenance and unresolved visual gates.
Safe managed dependency-free execution, cancellation, meaningful bounds, coarse
host admission/refunds, bounded caches/staging, cleanup, atomic publication and
PDF validity remain. The preserved roadmap stays deferred.

The last full workspace inventory remains **5.943 GB/885,270 files**, above the
loose 5 GB target. No full walk or broad purge was repeated; the milestone records
its own footprint and reuses shared inputs/worktree.
