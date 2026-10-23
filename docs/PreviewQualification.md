# 0.2.0-preview qualification

The bounded BMP recovery, build-isolation and guidance milestone is qualified.
Publication, performance optimization and broader feature expansion remain deferred.
The frozen library was built at `e3d25ef364682c2da03323e0509d1457d1d06616`, SHA-256
`2A30744C98E4C3F357D5326C7300A9BE56BE0189855C07D4AD1DE311382B65C5`.
[Windows/Ubuntu CI passed](https://github.com/Lokad/OoxPDF/actions/runs/39931970399)
at the subsequent README-only revision `6fadc899`. Production, tests and tooling
are identical between those revisions. Package guides use the corrected README;
the qualified DLL was not rebuilt. CI builds are independent of this artifact.

| Gate | Result |
|---|---|
| Release/affected checks | One committed linked-worktree artifact; build succeeds with two inherited nullable warnings. All 18 BMP/shared-image affected methods pass on Windows and Ubuntu. |
| Windows milestone | **3,773 pass, one inherited failure, five skips**; 3,779 exact unique methods in two disjoint invocations. |
| Ubuntu milestone | **1,001 pass, three inherited failures, 50 skips**; 1,054 exact unique methods in two disjoint invocations. No previous pass regressed. |
| Public BMP probes | Six conversions: valid controls return 0; invalid offsets/header extents return strict 3 with one image Error and omitted image. Both previous aborts recover normally. Valid/recovery-control PDFs retain their identities. |
| Build isolation | Five stubbed checks pass locally on Windows and in both-platform CI: independent checkouts, shared checkout outputs, checkout aliases, failed-build old-output refusal and missing-output refusal. No real builds in those checks. |
| Representative documents | Nine sources/16 outcomes/32 pages; PDFs, diagnostics and semantic results match the preceding candidate. Each outcome converted once. |
| Independent portfolio gates | Reused through PDF/input identity: two strict PPTX passes/fourteen strict DOCX failures. No new Office export, portfolio raster or relaxed threshold. |
| API/runtime | All 247 public declarations unchanged; IL-only, no P/Invoke/imported COM, explicit AllowUnsafeBlocks=false; assembly/file version 0.2.0.0. |
| Package/transfer | Same dependency-free 0.2.0-preview package and consumer source pass Windows/Ubuntu. Packed guides match their recorded source and local links resolve. All 24 root assets and two integration PDFs match; no root build or suite replay. |

## Conversion behavior and bounded maintenance

BMP decoding now establishes that the declared DIB header fits in the available
bytes before adding the file-header offset. A negative pixel offset or oversized
header uses normal image-error handling rather than an unhandled indexing exception.
This is conversion error handling. It adds no codec, full BMP validator or broad
exception catch, and changes no public API.

Direct cases cover overflow-sized headers, negative/overlapping/out-of-range
offsets and incomplete headers. Supported bottom-up/top-down pixels, padded rows,
extended headers and existing 32-bit behavior retain their controls. DOCX/PPTX
image recovery, cancellation and file-publication neighbors pass. PNG and BMP
conversion fixtures share the renamed ImageRecoveryTests helper while retaining
the existing PNG method names and capability selections.

The build helper keys its existing mutex by the physical checkout path, resolving
directory links before constructing the key. Projects in one checkout continue
to serialize because their referenced bin/obj outputs may be shared. Independent
checkouts can proceed separately; aliases of one checkout share the lock. MSBuild
up-to-date checking and refusal of failed/missing output remain. No scheduler,
lock registry or replacement build framework was introduced. The stub checks
establish ownership behavior, not a real-build throughput measurement.

README now links to the existing pinned-PDFium instructions through an absolute
repository URL, also reachable from the package. The vendored binary/version/hash
remain unchanged. Font-download comments distinguish payload/concurrency bounds
from accumulator/copy construction headroom. Runtime accounting, cache ownership,
options and memory admission are unchanged; no allocation improvement is claimed.

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

Ignored evidence: `artifacts/preview-bmp-build-scope-20261023/qualification.json`;
artifact: `candidate-final/`; package:
`artifacts/nuget/preview-020-bmp-build-scope-20261023/`.
One artifact served affected tests, disjoint broader remainders, six public probes,
sixteen representative outcomes, both package consumers and verified transfer.
Candidate methods were not rerun. The README-only correction reused that DLL;
the superseded CI run was cancelled and the corrected revision passed.

The [preceding qualification](https://github.com/Lokad/OoxPDF/blob/90c72466/docs/PreviewQualification.md)
retains font-identity, PNG and earlier independent Office evidence. Safe managed
dependency-free execution, meaningful bounds, cancellation, coarse host admission
and refunds, bounded caches/staging, deterministic cleanup, atomic publication and
PDF validity remain. The preserved roadmap stays deferred.

The last full workspace inventory remains **5.943 GB/885,270 files**, above the
loose 5 GB target. No new full walk or purge was performed. Qualification records
the small milestone footprint separately; shared inputs and the worktree are reused.
