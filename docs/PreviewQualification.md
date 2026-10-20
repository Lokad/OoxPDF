# 0.2.0-preview qualification — 2026-10-20

The second bounded review is locally qualified and the preview package is prepared.
Publication, tagging, release creation, performance optimization and broader feature
expansion remain deferred. See [scope](PreviewRelease.md) and [migration](PublicApiChanges.md).

One committed Release runtime is frozen from `0db17b3713cc5a64c8c880eacf606f588a9e7fd5`.
Library SHA-256:
`FDB4FFD5CF32349CE78A1FE904C842E1E455CF3334AD38C9DB0E2893B13B8742`.
Test assembly SHA-256:
`D83128EF6E340CC96D32BA6ADD0DC87484412DEB388CCA4BD216C96B49CD7BF2`.
[Candidate CI](https://github.com/Lokad/OoxPDF/actions/runs/39461816319) builds each
platform independently; its compiler outputs are separate from this frozen runtime.
Both Windows and Ubuntu CI jobs passed. Final documentation and catalog-loop
cleanup preserve the qualified library/test trees; catalog selections are verified
identical and missing selectors still reject. No broad suite was replayed.

| Gate | Result |
|---|---|
| Affected checks | All 136 pass: PDF/resources, decoders, observer outcomes and presentation-font neighbors; extracted from the single Windows broad run, no affected-suite replay |
| Windows broad | 3,722 pass, **1 inherited fail**, 11 skip; retained balloon assertion below |
| Same frozen DLL on Ubuntu | 799 pass, 0 fail, 44 environmental skips; one exact 843-method invocation, no library rebuild |
| Parent controls | Four new observer/publication regressions fail against the preceding DLL; its cancellation-token control passes |
| Qualification tools | Twenty fixture/snapshot checks and seven visual-family binding checks per platform; no real library execution in those tool checks |
| Capability catalog | All 273 capabilities and 2,178 distinct selectors resolve against one frozen catalog; missing selectors still reject; audit executes no tests |
| Final package consumers | Windows and Ubuntu consume the same final .nupkg; DLL identity, version, caller docs, embedded fonts and admission lifetime pass |
| Representative documents | Nine sources, sixteen outcomes, 32 pages; every PDF, diagnostic list and recorded semantic result matches the accepted parent |
| Office comparisons | Reused by input/options/PDF identity: two PPTX strict passes, fourteen DOCX strict failures; no Office export, new raster or threshold change |
| Transfer | All 17 assets verified and two integration PDFs match; no root build or suite replay |
| Safe managed | Explicit AllowUnsafeBlocks=false, IL-only, zero P/Invoke/imported COM types, no runtime package dependencies |

The local balloon assertion sees a 12.61pt single-row body where its Office-derived
expectation requires a second row. The preceding qualified DLL fails identically
on this machine. Its cause remains unresolved; the assertion, font guards and
rendering policy stay unchanged. Previous CI skipped it without Aptos. A separate
expensive balloon-spacing method passes on both current and preceding DLLs; live
stack inspection resolved a suspected stall. Its combinatorial test cost is not a
measured renderer regression or an optimization result. This local broad run has
one retained failure even though every affected check passes.

## Changes and compatibility

Caller diagnostic exceptions keep their original identity, type and stack through
recoverable PPTX paths. A private conversion-local marker prevents document-recovery
filters from swallowing observer failures; the public boundary restores the original.
No global observer state, deferred callbacks or extra output buffer is introduced.
Nested conversions, sync/async overloads and observer cancellation are covered.
Final-summary cancellation is checked before atomic file publication.

Removed the unused PdfLuminositySoftMask class, builder registration/index/rollback
bookkeeping, ExtGState field, writer planning/numbering/emission and two obsolete
mask-only tests. Live PNG alpha masks and SVG vector luminosity groups keep their
resource validation. Decoder tests now use live ReadOwned methods; the three legacy
Read wrappers and unused BMP result object are gone. A redundant presentation font
catalog field/eager fallback allocation is removed. No public API is removed here.
Earlier breaking changes remain in the caller migration guide.

Verbatim fixture checkout in .gitattributes makes shared input identity independent
of autocrlf without dropping byte verification. The stale clipping selector is
repaired. CheckVisualFamily accepts a frozen CLI and optional no-build mode, forwards
that binding to every case, pins runtime hashes and separates validation-tool builds.
Current equation, markup-triage and visual-validation guides replace contradictory
journals with scope/instructions and links to exact Git history. Shared font-cache
ownership remarks now describe the actual central bounded cache.

## Artifacts and residuals

Final packages: `artifacts/nuget/preview-020-second-review-20261020/`.

- .nupkg SHA-256: `B831958BE6E635FF429068DAFE55B154521A2D00A1591794F96DDBB4CD1898BA`.
- .snupkg SHA-256: `4FA6CB4FD783BCCBEC265CFDAEEF06EC68736F70D61843EA876672D8281ADFE1`.

Packaging uses committed caller guides at `759b3aed` and the qualified library/PDB
without recompilation. Package version stays `0.2.0-preview`, assembly/file version
`0.2.0.0`. All four bundled caller guides match the committed checkout. The preceding
qualified runtime/package remains preserved under `artifacts/review-follow-up-20261020/`
and `artifacts/nuget/preview-020-review-20261020/`.

Fourteen DOCX strict failures include significant geometry and typography. SVG
Review prose Y remains 1.436035pt against the unchanged 1pt gate; raw linked-footnote
Word URI difference remains 46.794pt while native picture/link bounds agree within
.000016pt. Two cached-field/extraction flags remain explicit. No new supported-scope
content or pagination regression appears in the fixed portfolio; this is not an
arbitrary-document ink/clipping guarantee. Live chart-color, exact-five-paragraph
balloon and font/document-shape guards lack a newly evidenced general replacement
and are preserved without expansion.

Full identities, deletions, controls, reports and residuals:
`artifacts/review-second-20261020/qualification.json`. Latest scan: **4.930 GB**,
zero errors, before the small final package/transfer addition. No historical
artifacts were removed; keep the loose 5 GB cap and infrequent checks. Preserve
safe managed execution, input bounds, coarse admission/refunds, bounded caches/page
staging, cancellation, deterministic cleanup, atomic publication and PDF validity.
