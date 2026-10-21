# 0.2.0-preview qualification - 2026-10-21

The fifth bounded review is qualified with the retained failures below. Publication,
tags, releases, performance optimization and broader feature expansion remain deferred.
See [scope](PreviewRelease.md) and [migration](PublicApiChanges.md).

One Release production DLL was built at `dfc5f018e57d7e853c484b31d75c03a60e0cfd4b`. A two-line test-fixture
repair at `b5a6a1a74fcbf2dc8a2695403c8972b907cba4cd` was compiled with project-reference builds disabled.
The library and all 21 runtime/support assets remain identical; only the reviewed
test harness changed. The final local artifact is frozen with 24 assets and 2,582
public inputs, reusing the same input snapshot.

Library SHA-256:
`4DE1338378D1039F89C88FBE8BE3112C29879F0AC5124C75B5623CA582CC6FE9`.
Final test assembly SHA-256:
`FE4C5E71EB598A84F3EBD715FD5F08D73CC2BF3824D9387C7E1125098C26C691`.
[Windows and Ubuntu CI passed](https://github.com/Lokad/OoxPDF/actions/runs/39574793428)
at `b5a6a1a7`. CI compiler outputs are independent of the locally qualified DLL.

| Gate | Result |
|---|---|
| Release build | Passed; two existing nullable warnings in unchanged DOCX test code. No library rebuild for the harness repair |
| Windows milestone | 3,732 pass, **1 inherited fail**, 5 skips across 3,738 unique methods; one corrected method rerun, no unchanged-method replay |
| Same frozen DLL on Ubuntu | 817 pass, 0 fail, 43 environmental skips across 860 unique methods; only the corrected method rerun |
| Exact affected capability | 19 methods: Windows 17 pass / 2 case-insensitive-filesystem skips; Ubuntu all 19 pass. Results reused from milestone reports plus reviewed repair |
| Publication regression | Relative and absolute destinations, synchronous/asynchronous summary callbacks and relative input with admission callbacks pass. Preceding DLL fails the new publication regression |
| Linux paths | Case-distinct input/output and PDF/diagnostic names pass on a case-sensitive filesystem; existing same-path checks remain |
| Raster wrapper | Five real PDFium checks pass: spaced tool/input/output paths, invalid PDF and zero-deadline timeout. No tool/library build or Office export |
| Transfer source checks | 28 checks per platform; dirty/committed inspector source and tools-parent targets refuse before transfer, harness-only edits remain permitted |
| CLI diagnostics | Windows 4 pass / 2 symbolic-link privilege skips; Ubuntu all 6 pass. Existing hard/symbolic-link publication and failed-publication cleanup remain covered |
| Capability catalogue | 279 capabilities / 2,205 distinct selectors resolve against 3,744 methods; the new path/publication set contains exactly 19 methods |
| Package consumers | Windows and Ubuntu consume the same .nupkg; DLL identity, versions, guides, embedded fonts and admission lifetime pass |
| Representative documents | Nine sources / 16 outcomes / 32 pages: every PDF, diagnostic list and semantic result matches the preceding runtime. Ten base outcomes and six supplemental outcomes converted once, then joined |
| Independent Office comparisons | Reused through proven identity: two PPTX strict passes / fourteen DOCX strict failures. No new Office exports, portfolio rasters or threshold changes |
| Root transfer | All 24 assets verified, two integration PDFs match; no root rebuild or suite replay |
| Safe managed | Explicit AllowUnsafeBlocks=false, IL-only, no P/Invoke/imported COM types; dependency-free package, assembly/file 0.2.0.0 |
| SDK cleanup | 28 unused headers/import-library/CMake/GN files removed, 610,054 bytes. Pinned native DLL, all licenses and version/provenance remain |

The initial Windows attempt also recorded a test-cleanup failure: File.Copy preserved
read-only attributes from frozen fixtures. The repair constructs an independently
owned writable input from bytes; it changes no library behavior. The original attempt
is retained. On Ubuntu an additional single-method attempt failed because its owned
/tmp target had disappeared between processes; restoring that target produced a pass.
That failure is also retained. The broad suites were not replayed. Ubuntu therefore
records 862 method attempts, Windows 3,739; unique-method counts above exclude repairs.

The inherited Windows balloon assertion still observes a 12.61pt single-row body where
its Office-derived expectation requires another row. Its preceding-runtime failure
matches exactly; no assertion or rendering guard changed. Earlier nonportable native-font
inspection failures remain historical evidence; that unchanged slow probe was not rerun
in this batch. Local retained failures/skips are separate from green CI.

## Changes and compatibility

File conversion captures absolute input/output paths before host callbacks. A callback
changing the process working directory cannot redirect input lookup or replace an
unrelated same-named destination. File-path comparison respects the platform; extension
and option matching stay case-insensitive. No public signature or option was removed.

The raster wrapper passes discrete arguments and preserves external timeout/kill and
process disposal. Transfer identity checks include the frozen DOCX inspector and its
parent build settings. Release guidance no longer hard-codes the obsolete 17-asset
count, and PDFium setup now describes the actual pinned DLL/managed-tool workflow.
Unused SDK development files were deleted without changing the native validation DLL.

Existing chart point-count regimes, five-paragraph balloons and font/size/document-shape
guards remain live, documented debt. No new independent Office evidence justified a
replacement. Useful inspectors, caches, input bounds, cancellation, coarse admission,
bounded staging, cleanup/publication and PDF structural validity remain.

Strict DOCX failures remain significant geometry/typography residuals, not uniformly
cosmetic: SVG Review baseline Y differs by 1.436035pt against the unchanged 1pt gate;
linked-footnote Review Word URI bounds differ by 46.794pt although candidate picture/link
bounds agree within .000016pt. Two field/extraction flags remain. No renderer policy,
Office threshold or supported-content scope changed in this batch.

## Retained candidate evidence

Ignored local evidence: `artifacts/preview-paths-20261021/qualification.json` and
`completion-audit.json`. The final artifact is `candidate-qualified/`; the initial
harness/runtime artifact and all failed attempts are retained separately. Reviewed
joins verify unchanged runtime assets, inputs and exact method coverage.

Package: `artifacts/nuget/preview-020-paths-20261021/`.
Package SHA-256:
`0F79B746D54E1E951DA537041AB252C670CACC366234D5FF226262423CBAB632`.
Symbol package SHA-256:
`A912B2C46DAFB8C512C6D3BD24813E6987297EE574B772C0E3EDBC963D9C30D4`.
The packaged DLL and symbol PDB match the frozen artifact. Four bundled caller guides
match the qualified checkout; Git line endings explain root-byte differences. The
package was not rebuilt or repacked for a test-only repair.

Latest workspace scan: **5.266 GB**, 848,179 files, zero errors,
2026-10-21. This counts unique file-content lengths, not allocated clusters. It is near
the loose 5 GB target. Historical artifacts, protected targets and useful evidence
were preserved; no broad historical purge occurred.
