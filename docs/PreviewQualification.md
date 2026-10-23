# 0.2.0-preview qualification

The bounded chart-date recovery and reference-tool milestone is qualified.
Publication, performance optimization and broader feature expansion remain deferred.
The frozen library was built at `428416961ae3b28c52d47c68d8b79b81c9b7bf11`, SHA-256
`378538189110D1665A351C33882605436675A4D8A608E6962B2D9CD13B8FA789`.
[Windows/Ubuntu CI passed](https://github.com/Lokad/OoxPDF/actions/runs/39940644136)
at the qualified source revision.
CI builds are independent of the single artifact used for local qualification.

| Gate | Result |
|---|---|
| Release/affected checks | One committed linked-worktree artifact; build succeeds with two inherited nullable test warnings. All 23 number-format/axis/data-label/composite-chart methods pass on Windows and Ubuntu. |
| Windows milestone | **3,775 pass, one inherited failure, five skips**; 3,781 exact unique methods in two disjoint invocations. |
| Ubuntu milestone | **1,005 pass, three inherited failures, 50 skips**; 1,058 exact unique methods in two disjoint invocations. No previous pass regressed. |
| Chart composition probes | Three conversions; each retains twelve bars, four categories and three legend entries. The previously omitted date-formatted chart returns strict 0. Numeric and in-range date control PDFs are unchanged. |
| Reference tools | Thirty managed-stub checks cover supervision, exception stages, active timeout stages and rasterizer flag forwarding. Eight view/provenance and fifteen derivative checks pass. A real pinned-rasterizer integration reuses an Office PDF with no build or Office activation. |
| Representative documents | Nine sources/16 outcomes/32 pages; PDFs, diagnostics and semantic results match the preceding candidate. Each outcome converted once. |
| Independent portfolio gates | Reused through PDF/input identity: two strict PPTX passes/fourteen strict DOCX failures. No new Office export, portfolio raster or relaxed threshold. |
| API/runtime | All 247 public declarations unchanged; IL-only, no P/Invoke/imported COM, explicit AllowUnsafeBlocks=false; assembly/file version 0.2.0.0. |
| Package/transfer | Same dependency-free 0.2.0-preview package and consumer source pass Windows/Ubuntu. Packed guides match their recorded source and local links resolve. All 24 root assets and two integration PDFs match; no root build or suite replay. |

## Conversion behavior and bounded maintenance

Chart date formatting now checks the representable calendar range before converting
serial days or adding them to a date. An unrepresentable date uses the existing
numeric fallback for that label. The chart survives; ordinary dates, both date
systems, phantom-day/weekday behavior and numeric formatting retain their controls.
No new formatter framework, diagnostic or public option is introduced.

Independent PowerPoint 16 output confirms all twelve bars, four categories and
three legend entries on the retained omission probe. Exact out-of-range tick labels
remain approximate: Office repeats a date where the candidate uses numeric fallback.
The 72 DPI exploratory visual screen remains needs-review: MAE 1.3170 against 1.2,
foreground recall .89705 against .90 and changed-pixel ratio .02740 against .025.
These limits are unchanged; this screen is separate from the original manifest's
144 DPI strict qualification. Plot/title geometry also remains approximate. Visual
inspection found visible content and no clipping; it does not establish Office parity.

Reference workers write the active stage before activation, open, export, cleanup
and rasterization. Completed cleanup still preserves the original exception stage;
cleanup failures remain separately attributed. The supervisor continues to own only
its worker and retains the shared Office lock. No supervision framework was added.
Reference exports and cache misses forward `-SkipBuild` to the existing rasterizer;
default build freshness and PDFium pin checks remain. The binary pin is unchanged.

## Inherited failures and limitations

**DocxWordCompatibleBalloonBodyWrapsLikeOfficeTitles** retains its local failure:
`InvalidOperationException: The second body should wrap to a second row like Office; observed Height=12.610000000000001.`
CI may skip its Office-font precondition. It is not counted as a local pass.

The three inherited Linux failures remain:

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

## Evidence and workflow

Ignored evidence: `artifacts/preview-chart-date-recovery-20261023/qualification.json`;
artifact: `candidate-final/`; package:
`artifacts/nuget/preview-020-chart-date-recovery-20261023/`.
One artifact served affected tests, disjoint broader remainders, three chart probes,
sixteen representative outcomes, both package consumers and verified transfer.
Candidate methods were not rerun. Office references/comparisons and prior Linux
failure provenance were reused without parent builds or parent method replays.

The [preceding qualification](https://github.com/Lokad/OoxPDF/blob/8b9e70a3/docs/PreviewQualification.md)
retains BMP recovery, build isolation, font identity, PNG and earlier Office evidence.
Safe managed dependency-free execution, meaningful bounds, cancellation, coarse
host admission and refunds, bounded caches/staging, deterministic cleanup, atomic
publication and PDF validity remain. The preserved roadmap stays deferred.

The last full workspace inventory remains **5.943 GB/885,270 files**, above the
loose 5 GB target. No new full walk or purge was performed. Qualification records
the small milestone footprint separately; shared inputs and the worktree are reused.
