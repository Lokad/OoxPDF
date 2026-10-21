# 0.2.0-preview qualification — 2026-10-21

The tenth bounded cleanup milestone is qualified with inherited rendering
failures explicit below. Publication, tags, releases, optimization and broader
feature expansion remain deferred. See [scope](PreviewRelease.md),
[migration](PublicApiChanges.md) and [document review](PreviewVisualReview.md).

The Release artifact is frozen at `0d1236c1911207e4cf832e274df37f1f3d168ea4`:
24 assets and 2,582 verified shared public inputs. One candidate was qualified;
there was no root library build or repeated broad suite. The same DLL and package
were used on Windows and Ubuntu.

Library SHA-256:
`F95922ED243D3789BCFA5DD2705F38333D7B2167980308C0BD3CE0B50DA7CDC1`.
[Windows and Ubuntu CI passed](https://github.com/Lokad/OoxPDF/actions/runs/39618218300)
at `0d1236c1`. CI builds are independent of the locally qualified DLL.

| Gate | Result |
|---|---|
| Release build | Passed; two existing nullable warnings in unchanged DOCX tests |
| Windows milestone | 3,740 pass, **one inherited failure**, five skips; 3,746 unique methods |
| Same frozen DLL on Ubuntu | 917 pass, zero failures, 45 environmental skips; 962 unique methods |
| Focused checks | All 98 Windows / 24 portable checks pass in one exact invocation per platform |
| Qualification reuse | Focused methods excluded from remaining milestone sets; two disjoint invocations per platform, no replay |
| Inert state | Removed the 39-field PPTX emission context, redundant forwarding profile and five-field legend result; actual geometry and authored/override spacing remain |
| Internal seams | Ten unused default-context overloads and three test-only forwarders removed; tests target live recolor, TTC and threaded-comment implementations |
| Argument cleanup | All 49 reviewed unread parameters across 33 methods eliminated, plus two immediate caller arguments; Office geometry conditions remain |
| API/callers | All 247 exported declarations unchanged. 25,673 method bodies screened; no unresolved method operands or authored private zero-caller candidates. This does not prove every internal/public path is necessary |
| Inspectors | Eleven useful inspector entry points and three useful color snapshot adapters retained; allocation tool builds/self-tests against the frozen DLL without rebuilding it |
| Representative documents | Nine sources / 16 outcomes / 32 pages; every PDF, diagnostic list and semantic result matches the preceding candidate. Each outcome converted once |
| Office comparisons | Reused by exact output identity: two PPTX strict passes / fourteen DOCX strict failures. No new Office export, portfolio raster or threshold change |
| Package consumers | Same .nupkg on Windows/Ubuntu; DLL/PDB, four guides, versions, embedded fonts and admission lifetime verified |
| Root transfer | All 24 assets verified; two integration PDFs match, without library rebuild or suite replay |
| Safe managed | Explicit unsafe=false, IL-only, no P/Invoke/imported COM; dependency-free runtime, assembly/file 0.2.0.0 |
| Materials/catalog | Historical audits link retained Git history and current contracts. All 715 manifests validate; 287 capability selections resolve. Manifest validity is not Office fidelity |

The inherited Windows balloon assertion observes a 12.61pt single-row body where
its Office-derived expectation requires another row. Its method and failure detail
match the previous candidate. Its assertion and rendering policy remain unchanged.
Historical nonportable native-font failures remain unrerun evidence. These local
failures and skips are separate from green CI.

## Changes and preserved behavior

PPTX glyph emission reads run font size and authored/override tracking directly.
The shared 600-DPI Office font-size rounding remains, along with independent
numbering/highlight spacing, synthetic faces, kerning, alignment and clip behavior.
Legend layout returns its computed width; real text measurement and calibrated
padding remain. Removing dead name-length storage does not replace chart geometry.

Unused arguments are removed from chart labels/titles, effects/custom paths,
baseline helpers, font-directory reads and DOCX story/cell helpers. Named/default
arguments were bound to exact source symbols. Source partitions, text-baseline and
image-overlap rules retain their distinct behavior. The old polar identity test
now checks the live radius and independent doughnut width clamp; legend padding
assertions remain. An inherited raw NUL in a hyperlink diagnostic key now uses the
same C# escape value, making the source searchable without changing the key.

Chart point-count regimes, five-paragraph balloons and font/size/document-shape
guards remain unchanged. Safe managed execution, meaningful input bounds,
cooperative cancellation, coarse admission/refunds, bounded caches/page staging,
deterministic cleanup, atomic publication and PDF structural validation remain.
No additional public API was removed and no performance improvement is claimed.

Fourteen DOCX strict gates retain typography/geometry failures. SVG Review Y
1.436035pt exceeds 1pt; raw Word footnote URI bounds differ by 46.794pt although
candidate picture/link bounds agree within .000016pt. Two field/extraction flags
remain. They were not relabeled cosmetic or passed.

## Evidence and candidate

Ignored evidence: `artifacts/preview-inert-state-20261021/qualification.json`;
frozen runtime: `candidate/`. Exact signature edits, deletion records, focused
reports and original disjoint milestone reports remain alongside it. Editing-tool
preparation is recorded separately from library qualification.

Package: `artifacts/nuget/preview-020-inert-state-20261021/`. Package/symbol hashes,
qualified guide identities and consumer summaries are in `package-identity.json`.
The preceding candidate and independent Office references remain retained.
Publication stays deferred.

Latest infrequent workspace scan: **5.771 GB**, 868,377 files,
zero errors. This counts unique content lengths, not allocated clusters. Keep
near the loose 5 GB target; useful inputs, evidence and Git history are preserved.
