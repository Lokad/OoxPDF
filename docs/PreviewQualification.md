# 0.2.0-preview qualification — 2026-10-22

The compatibility milestone restores supported content inside declared
ProcessContent wrappers and corrects MustUnderstand prefix handling. Eight
misleading fixtures and caller-facing errata are corrected; the capability
inventory now states scope and residuals concisely. Publication, tags, releases,
performance optimization and broader feature expansion remain deferred.
See [scope](PreviewRelease.md), [migration](PublicApiChanges.md),
[capabilities](Capabilities.md) and [document review](PreviewVisualReview.md).

One committed Release artifact is frozen at
`b4a3ee64b1ab6c887528709c01961cb631a33081`: 24 assets and 2,582 verified shared
public inputs. The same DLL was qualified on Windows and Ubuntu, consumed through
the same package on both platforms, then transferred with two integration smokes.
There was no root library rebuild or broad suite replay.

Library SHA-256:
`3F71BC99A36F4562E12519F3A29B09E523B3E830B49CCE5ED83E16238BE7F07D`.
[Windows and Ubuntu CI passed](https://github.com/Lokad/OoxPDF/actions/runs/39771379327)
at the frozen source revision. CI builds are independent of the locally qualified DLL.

| Gate | Result |
|---|---|
| Release build | Passed; two existing nullable warnings in unchanged DOCX tests |
| Focused compatibility | 37 checks pass on each platform; scope, nested/wildcard wrappers, branch selection, cancellation, chart annotations and DOCX/PPTX composition |
| Reproduced omissions | Eleven pinned probes; all three omissions restored to exact unwrapped-control PDF bytes; ordinary control PDFs unchanged |
| Windows milestone | 3,754 pass, **one inherited failure**, five skips; 3,760 unique methods |
| Same frozen DLL on Ubuntu | 932 pass, zero failures, 47 skips; 979 unique methods |
| Qualification mechanics | Affected/broader sets disjoint initially. Four Windows/three Ubuntu CLI methods were repeated to repair slow-filter skips; all pass. Raw attempts retained; no broad replay |
| Local Linux limitation | Two case-distinct-path checks skip on the case-insensitive mounted filesystem; they passed in the earlier native-filesystem environment. They are not counted as passes here |
| Representative documents | Nine sources/16 outcomes/32 pages; PDFs, diagnostics and semantic outcomes exactly match the preceding candidate. Each outcome converted once |
| Independent Office gates | Reused through verified PDF/input identities: **two PPTX strict passes/fourteen DOCX strict failures**. No new Office export/raster or changed threshold |
| API/runtime | All 247 exported declarations unchanged; IL-only, no P/Invoke/imported COM, explicit AllowUnsafeBlocks=false; assembly/file version 0.2.0.0 |
| Package | Same dependency-free 0.2.0-preview package consumed on both platforms; loaded DLL hash, versions, admission release and conversion checked |
| Transfer | All 24 root assets verified; two integration PDFs match; retained candidates preserved |

## Behavior and caller guidance

Declared ignorable ProcessContent wrappers expose supported children through the
existing shared XML stage. Namespace declarations resolve before their carriers
are removed; unselected alternate branches do not participate. Ordinary extension
fallbacks and the supported namespace set remain unchanged. This adds no layout
framework, cache instrumentation or public API declaration.

OOXML_MUST_UNDERSTAND remains a warning. It now honors namespace-prefix lists;
Fallback is `Partial compatibility` rather than `Ignored`. Conversion continues
with partial handling, not a claim of full MCE conformance. The corrected README
returns identity with actual font-program bytes; different programs must have
different identities. The package now includes that example and the previously
clarified file-font stability precondition.

The PDF profile describes Form XObjects, Type1/WinAnsi fallbacks, ActualText and
compound glyph closure accurately. The diagnostic catalogue distinguishes painted
approximate equations from omitted unsupported equations. Capabilities.md shrank
from 9,169 to 1,588 words; historical detail remains linked, and no live chart-color,
balloon or typography guard was deleted.

## Inherited failures and residuals

**DocxWordCompatibleBalloonBodyWrapsLikeOfficeTitles** still fails with
`InvalidOperationException: The second body should wrap to a second row like Office; observed Height=12.610000000000001.`
The preceding/current runtime agree. It remains a failure; CI can skip its
Office-font precondition. No failure was replaced by a skip during report repairs.

The fourteen DOCX strict failures include significant geometry/typography, not
only cosmetic differences. SVG review Y remains 1.436035pt against 1pt; raw Word
footnote annotation mismatch remains 46.794pt while candidate picture/link agree
within .000016pt. Two field/extraction flags and historical nonportable font
failures remain explicit. Point-count chart colors and document-shape guards
require independent Office evidence before changes. Publication remains deferred.

## Evidence and candidate

Ignored evidence: `artifacts/preview-markup-compatibility-20261022/qualification.json`;
frozen runtime: `candidate-final/`; package:
`artifacts/nuget/preview-020-markup-compatibility-20261022/`.
Exact selections, original/repair reports, probe inputs/results, package identities
and transfer evidence are retained. A package preflight compared differing line
endings before Linux execution; it was corrected to compare the qualified worktree,
without rebuilding the package or replaying a successful conversion.

The [prior qualification record](https://github.com/Lokad/OoxPDF/blob/b4ae23938be7df60bb6b1be09985e73fc2e2a523/docs/PreviewQualification.md)
and original independent comparisons retain their historical identities.
PLAN.md keeps the rendering roadmap and standalone enum-helper deferral. Last full
workspace measurement remains **5.864 GB**, 871,935 files, 2026-10-21; no new full
scan or unverified purge was performed. Keep near the loose 5 GB target.
