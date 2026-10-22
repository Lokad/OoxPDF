# 0.2.0-preview qualification

The bounded font-identity, packaged-guidance and source-hygiene milestone is
qualified at source `affc955940876c5b05fdf4c418409aede5416367`. Publication,
performance optimization and broader feature expansion remain deferred.
The library SHA-256 is
`3AA0CECA044D6AC3F541EE46CAC6FA40670AEDC8877D3C8680856CC4680E98E4`.
[Windows/Ubuntu CI passed](https://github.com/Lokad/OoxPDF/actions/runs/39819749039)
at that source. CI builds are independent of the locally qualified artifact.

| Gate | Result |
|---|---|
| Release/focused checks | One committed linked-worktree artifact; build succeeds with two inherited nullable warnings. All 40 affected Windows checks pass; all 14 portable font/cache/fallback neighbors pass on both platforms. |
| Windows milestone | **3,770 pass, one inherited failure, five skips**; 3,776 exact unique methods in two disjoint invocations. |
| Ubuntu milestone | **996 pass, three inherited failures, 50 skips**; 1,049 exact unique methods in two disjoint invocations. Three newly exercised chart-link failures reproduce on the preceding qualified runtime. |
| Public font probe | Changing the custom primary's opaque ID no longer changes fallback, advance or PDF bytes; all three variants match the previous neutral control. Independent catalog programs no longer alias the primary. |
| Representative documents | Nine sources/16 outcomes/32 pages; PDFs, diagnostics and semantic outcomes match the preceding candidate. Each outcome converted once. |
| Independent portfolio gates | Reused through PDF/input identity: two strict PPTX passes/fourteen strict DOCX failures. No new Office export, portfolio raster or weakened threshold. |
| API/runtime | All 247 public declarations unchanged; IL-only, no P/Invoke/imported COM, explicit AllowUnsafeBlocks=false; assembly/file version 0.2.0.0. |
| Package | Same dependency-free 0.2.0-preview package and consumer source pass Windows/Ubuntu. All four packed guides match qualified source; eight local references resolve. The previous package's absent guide is refused before consumer execution. |
| Transfer | All 24 root assets and two integration PDFs match the qualified artifact. All 24 preceding frozen assets remain unchanged; no root build or suite replay. |

## Behavior and bounded cleanup

PPTX fallback now treats font IDs as ordinal identities belonging to their
resolver. A separately owned discovery catalog has its own program-cache
namespace; a primary resolver that also supplies its catalog retains shared
program reuse. Font-family matching remains case-insensitive.

Subsets use the resolved immutable font program and normalized codepoint-set hash
as their key, with exact set comparison on hits. This removes a redundant source
resolution argument and string-key construction while retaining repeated
face/style/collection/subset reuse. No public identity API, framework or global
counter was introduced. Existing cancellation and cache ownership remain.

Portable regressions use different glyph coverage to challenge case-distinct IDs
and equal IDs under independent owners. They check repeat reuse, shared-catalog
ownership and cancellation on cache hits. Ubuntu CI now retains this small exact
capability selection alongside existing portable rendering coverage.

The public custom-resolver probe previously selected Abadi with a neutral ID and
Aptos with either an intersecting case variant or exact ID. At 12pt its advance
changed from 7.5pt to 7.072266pt. All variants now select Abadi at 7.5pt and produce
the same PDF as the previous neutral control. The exact cross-owner program lookup
now maps U+0100 to glyph 10 rather than the primary program's zero glyph. These
are identity-contract results, not an Office font-preference or arbitrary-document
fidelity claim. The original probe retained its text; memory corruption and new
content loss were not demonstrated.

Repository-only documentation links in the packed README/preview guide use
explicit URLs. The existing package smoke checks local references in the four
authored guides, including the valid parent changelog link. No Markdown dependency
or renderer was added. A literal NUL in hyperlink source is expressed as the
equivalent C# escape, restoring normal source searches; its false unfinished-chart
comment is corrected. Hyperlink keys and live chart/layout guards are unchanged.

## Inherited failures and deferred concerns

**DocxWordCompatibleBalloonBodyWrapsLikeOfficeTitles** still fails locally:
`InvalidOperationException: The second body should wrap to a second row like Office; observed Height=12.610000000000001.`
It remains a failure; CI may skip its Office-font precondition.

The expanded Linux neighboring selection also exercises these older tests:

- `PptxChartTitleExternalHyperlinkEmitsUriAnnotation`
- `PptxChartAxisTitleExternalHyperlinkEmitsUriAnnotation`
- `PptxChartDataLabelCustomTextHyperlinkEmitsUriAnnotation`

Each fails with `InvalidOperationException: Sequence contains no elements` when
the default Windows font catalog is absent. All three fail identically with the
preceding frozen test assembly/library on the same Linux environment. They assume
ambient fonts for annotation emission and remain portability debt, not passes.
No assertion relaxation, skip substitution or production font policy was added.
The new synthetic-font controls pass. The additional native-font precondition skip
and the 49 previously retained platform skips remain separate from failures.

The fourteen strict DOCX visual failures include significant typography/geometry.
SVG Review Y remains 1.436035pt against 1pt; the raw Word footnote annotation
discrepancy is 46.794pt, while candidate picture/link agree within .000016pt.
Two field/extraction flags and historical font limitations remain. The earlier
CMYK histogram .311828 versus .98 stays needs-review; no ICC/YCCK parity claim
follows from this font maintenance.

Transparency-group rollback accounting still has only an internal counterexample;
no public document sequence was established. Partial MCE, enum-helper and qualified
chart/font/size/document-shape limitations remain. No live rendering guard was
demonstrated redundant or replaced. The previous cache/accounting/validation
cleanup remains closed.

## Evidence and workflow

Ignored evidence: `artifacts/preview-font-identity-20261022/qualification.json`;
artifact: `candidate-final/`; package:
`artifacts/nuget/preview-020-font-identity-20261022/`.
One artifact served focused checks, disjoint broader remainders, public probes,
document outcomes, package consumers and transfer. Candidate test methods were
not rerun. The three parent-only Linux checks used the existing preceding runtime,
without a build. Two spaced CI queries established success.

Exact selections, preflight corrections, raw failures and artifact identities are
retained in the ignored qualification record. No extra document conversion or
candidate test replay was needed to complete those corrections.

The [preceding qualification](https://github.com/Lokad/OoxPDF/blob/a7e15d0e/docs/PreviewQualification.md)
retains PNG semantics, independent PowerPoint comparisons and earlier history.
Safe managed dependency-free execution, meaningful bounds, cancellation, coarse
host admission/refunds, bounded caches/staging, deterministic cleanup, atomic
publication and PDF validity are preserved. The roadmap remains deferred.
The last full workspace inventory remains **5.943 GB/885,270 files**, above the
loose 5 GB target; no new full walk or purge was performed. The milestone's scoped
footprint is recorded in qualification.json; shared inputs/worktree are reused.
