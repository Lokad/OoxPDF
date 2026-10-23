# 0.2.0-preview qualification

The bounded source-semantics and audit-comment cleanup milestone is qualified.
Publication, performance optimization and broader feature expansion remain deferred.
The frozen library was built at `f3cfe0c84399bcaecd116f07708d715e5844386b`, SHA-256
`E38DDD9C13C34DAD6F04EAD1C291C7AD37DEB1679FFF826897D0E66E94A69BA1`.
[Windows/Ubuntu CI passed](https://github.com/Lokad/OoxPDF/actions/runs/39959854800)
at the qualified source revision.
CI builds are independent of the one artifact used for local qualification.

| Gate | Result |
|---|---|
| Release/affected checks | One committed linked-worktree artifact; build succeeds with two inherited nullable test warnings. All 24 specific break/field/container and PPTX navigation methods pass on Windows and Ubuntu. |
| Windows milestone | **3,776 pass, one inherited failure, five skips**; 3,782 exact unique methods in two disjoint invocations. |
| Ubuntu milestone | **1,020 pass, three inherited failures, 50 skips**; 1,073 exact unique methods in two disjoint invocations. No previous pass regressed. |
| Source transformations | Seven conversions return strict 0 without diagnostics. Explicit/omitted DOCX break forms in Default/Review and mixed-case PPTX targets match their prior equivalent control PDFs exactly. |
| Independent Office evidence | Retained Word 16 references preserve the same two lines/origins; candidate deltas stay within the existing .1pt line-start/.75pt grouping tolerances. PowerPoint 16 retains the page-2 link for both target spellings. Input/reference hashes are verified. No new Office export, raster or pixel-parity claim. |
| Comment cleanup | 129 audit qualifiers removed across 14 files: 125 comments deleted, four semantic descriptions retained. Non-comment bytes remain identical in this separate commit; no live helper deleted. |
| Representative documents | Nine sources/16 outcomes/32 pages; PDFs, diagnostics and semantic results match the preceding candidate. Each outcome converted once. |
| Independent portfolio gates | Reused through PDF/input identity: two strict PPTX passes/fourteen strict DOCX failures. No new Office export, portfolio raster or relaxed threshold. |
| API/runtime | All 247 public declarations unchanged; IL-only, no P/Invoke/imported COM, explicit AllowUnsafeBlocks=false; assembly/file version 0.2.0.0. |
| Package/transfer | Same dependency-free 0.2.0-preview package and consumer source pass Windows/Ubuntu. Packed guides match their recorded source and local links resolve. All 24 root assets and two integration PDFs match; no root build or suite replay. |

## Conversion behavior and bounded maintenance

The shared DOCX text reader recognizes explicit `textWrapping` alongside its
omitted default. They retain a line break instead of joining adjacent words.
Page/column handling and existing floating/clear fallbacks remain. Regression
controls also split the source runs and add a content-control wrapper in both
markup modes. No layout correction or font/size/document-shape guard was added.

PPTX internal slide destinations use the same case-insensitive part identity
already used by package lookups. Relative, absolute and mixed-case target forms
select the same page; genuinely missing targets still diagnose and omit the link.
Relationship IDs, bookmarks and opaque font program IDs retain their distinct
identity contracts. No global normalization layer or public API was introduced.

Source cleanup removes review-era eligibility notes, including obsolete
single-use and future-refactoring claims. Useful semantic comments and live
helpers remain, including method-group consumers. The cleanup changes no runtime
mechanism, cache, ownership layer or resource contract. Four separate counter-family
symmetry comments remain outside the enumerated audit qualifiers.

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

Out-of-range chart dates retain numeric fallback rather than exact Office tick
text. The preceding 72 DPI screen remains needs-review (MAE 1.3170 > 1.2,
recall .89705 < .90, changed ratio .02740 > .025); it was separate from the
manifest's 144 DPI gate. That completed change is not reopened here.

## Evidence and workflow

Ignored evidence: `artifacts/preview-source-semantics-20261023/qualification.json`;
artifact: `candidate-final/`; package:
`artifacts/nuget/preview-020-source-semantics-20261023/`.
One artifact served affected tests, disjoint broader remainders, seven source
probes, sixteen representative outcomes, both package consumers and verified
transfer. Candidate methods were not rerun. Office references/comparisons, unchanged
715-case/12-family manifest validation and prior Linux failure provenance are
reused without parent builds or parent method replays.

The [preceding qualification](https://github.com/Lokad/OoxPDF/blob/2241af51/docs/PreviewQualification.md)
retains chart-date recovery, reference-tool checks, BMP recovery, build isolation,
font identity, PNG and earlier Office evidence. Safe managed dependency-free
execution, meaningful bounds, cancellation, coarse host admission/refunds,
bounded caches/staging, cleanup, atomic publication and PDF validity remain.
The preserved roadmap stays deferred.

The last full workspace inventory remains **5.943 GB/885,270 files**, above the
loose 5 GB target. No new full walk or purge was performed. Qualification records
the small milestone footprint separately; shared inputs and the worktree are reused.
