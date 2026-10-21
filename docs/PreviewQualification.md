# 0.2.0-preview qualification - 2026-10-21

The fourth bounded review is qualified with the retained failures below. The preview
package is prepared; publication, tags, releases, renderer optimization and feature
expansion remain deferred. See [scope](PreviewRelease.md) and [migration](PublicApiChanges.md).

One committed Release artifact is frozen from `e76fbf74b632b9e187d01c7987aab5b36cea373e`.
Library SHA-256:
`F85CFD70F758376DC9FFFB5694E3930E8EDA1B8CF123C3C6EC83B37D95964EF4`.
Test assembly SHA-256:
`F4BF669669EB9CB2C893EA86FECA1C2E5803552B6E77C08D85F0C3202029B324`.
[Windows and Ubuntu CI passed](https://github.com/Lokad/OoxPDF/actions/runs/39565300451)
at `4ae2c6d2`; its subsequent tool/workflow repairs leave library/test source unchanged.
CI compiler outputs are independent of this locally qualified artifact.

| Gate | Result |
|---|---|
| Windows milestone | 3,731 pass, **1 inherited fail**, 3 skips across 3,735 unique methods: regular suite plus one affected slow native-font method; no method replay |
| Same frozen DLL on Ubuntu | Portable selection: 813 pass, 0 fail, 43 environmental skips across 856 exact methods, one invocation |
| Additional Ubuntu native-font probe | One existing slow glyph-inspection assertion fails without the native font; preceding runtime fails identically. Windows affected method passes. This result is separate from the portable selection |
| Affected checks | All 77 pass on Windows; results extracted from the milestone reports |
| Atomic CLI diagnostics | Windows: 4 pass, 2 symbolic-link privilege skips. Ubuntu: all 6 pass. Hard-link source/destination aliases, symbolic links, separate diagnostics and failed-publication cleanup are covered; preceding CLI fails the new hard-link regression |
| Raster qualification | 21 orchestration checks per platform cover provenance, complete reuse and stale/changed/missing/extra pages, DPI and tool mismatches. Real PDFium/VisualDiff exposes pagination in the two-page versus one-page control that previously falsely passed |
| CI artifact retention | Downloaded Windows candidate: all 24 runtime assets present and hash-verified, including the seven previously omitted DOCX-inspector files |
| Capability selections | Local math-annotation/display-group/review-frame sets: 20/22/18 methods, previously 642/636/621. Complete preceding sets remain under explicit milestone names. All 278 capabilities and 2,198 distinct selectors resolve |
| Package consumers | Windows and Ubuntu consume the same .nupkg; DLL identity, version, caller guides, embedded fonts and admission lifetime pass |
| Representative documents | Nine sources, 16 outcomes, 32 pages; every PDF, diagnostic list and recorded semantic result matches the preceding accepted runtime |
| Independent Office comparisons | Reused by proven input/options/PDF identity: two PPTX strict passes, fourteen DOCX strict failures; no portfolio export/raster replay or threshold changes |
| Root transfer | All 24 assets verified, two integration PDFs match; no root rebuild or suite replay |
| Safe managed | Explicit AllowUnsafeBlocks=false, IL-only, no P/Invoke/imported COM types, no NuGet runtime dependencies; assembly/file version 0.2.0.0 |

The Windows balloon assertion retains a 12.61pt single-row body where its Office-derived
expectation requires a second row; the preceding DLL fails identically. The additional
Ubuntu slow inspection probe reports missing Hello glyphs when its native font is
unavailable; the preceding runtime fails identically. This is a nonportable inspection
test, not a newly demonstrated conversion omission. Assertions and rendering guards
remain unchanged. Local failures and environmental skips are distinct from green CI.

## Changes and qualification scope

CLI diagnostics are serialized into an independently owned file and atomically replace
the destination entry, separating existing hard/symbolic links. Literal path aliases
still reject before early failures. Failed diagnostic publication removes only its own
staging file; ordinary argument, conversion and strict-exit behavior stays covered.

Portfolio raster reuse binds the current PDF, DPI, wrapper/rasterizer/native identity
and complete page hashes. It rejects stale/incomplete derivatives, leaves legacy raster
evidence untouched and preserves reuse of accepted identical comparison reports.
The real pagination control uses qualified synthetic PDFs, not a new Office reference.

Removed the test-only InspectTextRuns path, its source-run coalescer, converter and
PptxTextRunSnapshot type. Both test callers use live scene/color/glyph contracts;
independent frame/layout/flow/glyph inspectors remain. The retained PPTX inspector
also runs successfully against the qualified library without a library build. Removed two retired unsupported-
note diagnostic entries from the caller guide; current approximation IDs remain.
No public API removal or rendering rule is introduced by this milestone.

The three curated selections retain semantic neighbors and independent Office checks.
Their recorded method times change from approximately 142s each to 1.407/4.375/.640s
using the preceding report; these are selection-cost sums, not a renderer speedup or
new standalone timing measurement. Existing broad sets remain explicit milestones.
Unchanged exhaustive DOCX balloon sweeps were not repeated. The affected slow PPTX
method ran once outside the regular selection; its results are retained separately.

## Package and residuals

Final packages: `artifacts/nuget/preview-020-fourth-review-20261021/`.

- .nupkg SHA-256: `A7C19B0BB95CD23036DA4B5CB747976AAA017F0B099FFCE26792CBAD6F6BFC97`.
- .snupkg SHA-256: `D5FD9E82EFB4BF96B0715EFF22FAB2E785AC6CF51F3CDE8139D3912A533B7742`.

Packed without recompilation; all four bundled caller guides match the committed
checkout. Package/informational version stays `0.2.0-preview`. Prior runtime/package
and failed or superseded attempts remain retained.

Fourteen DOCX strict failures include significant geometry/typography. SVG Review
prose Y remains 1.436035pt against the original 1pt gate; raw linked-footnote Word URI
difference remains 46.794pt while native picture/link bounds agree within .000016pt.
Two cached-field/extraction flags remain explicit. The fixed portfolio shows no new
content/pagination regression; arbitrary-document fidelity is not established. Chart
point-count and font/size/document-shape guards retain unresolved behavior without
expansion or speculative replacement. All 715 visual manifests validate; that count
is not an Office visual-pass count.

Full identities, reports, deletion/capability evidence, retained failures and CI:
`artifacts/review-fourth-20261021/qualification.json`. Latest scan: **5.149 GB**, zero errors, before the small retained-inspector smoke.
Historical inputs, comparisons and runtimes are preserved.
Keep the loose 5 GB workspace target and preserve safe managed execution, input
bounds, bounded caches/staging, coarse admission/refunds, cancellation, deterministic
cleanup, atomic publication and PDF structural validity.
