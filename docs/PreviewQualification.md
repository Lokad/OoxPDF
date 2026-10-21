# 0.2.0-preview qualification — 2026-10-21

The eighth bounded maintenance milestone is qualified with the retained failures
below. Publication, tags, releases, performance optimization and broader rendering
expansion remain deferred. See [scope](PreviewRelease.md), [migration](PublicApiChanges.md)
and the retained [document review](PreviewVisualReview.md).

One Release production/test artifact was built and frozen at
`fdfdde4ad37482697ea345b44387fdbb2279e109`, with 24 assets and 2,582 shared public
inputs. Three scoped commits protect readable input-link targets, reject stale
served-reference pages, and remove wrapping counters and probe fuel. There was
no local root rebuild or repeated broad qualification after transfer.

Library SHA-256:
`E269F45E5AB9E3D4CF3842A8008D7ACEC9F8CE79A700F7F9054FC1B85527656F`.
[Windows and Ubuntu CI passed](https://github.com/Lokad/OoxPDF/actions/runs/39599672373)
at `fdfdde4a`. CI outputs are independent of the locally qualified DLL.

| Gate | Result |
|---|---|
| Release build | Passed; two existing nullable warnings in unchanged DOCX test code |
| Windows milestone | 3,740 pass, **1 inherited fail**, 5 skips across 3,746 unique methods |
| Same frozen DLL on Ubuntu | 864 pass, 0 fail, 45 environmental skips across 909 unique methods |
| Exact affected capabilities | 65 methods: 22 publication and 43 wrapping; each platform 63 pass / 2 explicitly declared environmental skips |
| Test reuse | Affected methods excluded from each platform's remaining milestone selection; two disjoint invocations, no method replay |
| Publication boundaries | Windows 13 pass / 10 symbolic-link privilege skips; Ubuntu all 23 pass. Input links/chains/linked target ancestors and same-entry input links refuse; separate output links/hardlinks remain allowed |
| Reference cache serving | Fifteen completion/recovery/provenance/freshness controls pass on each platform; caller files and trusted PDFs preserved |
| Independent wrapping correction | Synthetic mixed-width DOCX matches Word 16's six line lengths exactly: 700/165/165/165/165/40; all 1,400 characters preserved. Parent first line was only 395 characters. Maximum relative line-gap delta .067994pt, within the existing .15pt ordinary-word tolerance |
| Wrapping visual evidence | At 144 DPI, MAE improves .442024 to .297062 and SSIM .821860 to .900480. Remaining glyph/line geometry differences are explicit; no pixel-perfect claim or threshold change |
| Representative documents | Nine sources / 16 outcomes / 32 pages; every PDF, diagnostic list and semantic result matches the preceding runtime. Each outcome converted once |
| Independent portfolio comparisons | Reused through proven input/mode/PDF identity: two PPTX strict passes / fourteen DOCX strict failures; no new portfolio rasters or Office exports. The separate wrapping probe uses one new Word export |
| Package consumers | Windows and Ubuntu consume the same .nupkg; DLL/PDB identity, versions, four caller guides, embedded fonts and admission lifetime pass |
| Root transfer | All 24 assets verified, two integration PDFs match; no root rebuild or suite replay |
| Safe managed | Explicit unsafe=false, IL-only, no P/Invoke/imported COM types; dependency-free runtime, assembly/file 0.2.0.0 |
| Manifest/catalog integrity | 715 visual manifests validate; every capability selector resolves. This is not an Office fidelity count |

The Windows balloon assertion remains the same inherited failure: 12.61pt single-row
body where its Office-derived expectation requires another row. Its method name,
failure detail, assertion and rendering policy match the preceding qualification.
Earlier nonportable native-font inspection failures remain historical evidence;
that slow probe was not rerun. Local failures/skips are separate from green CI.

## Changes and compatibility

Input protection distinguishes the readable source from replaceable publication
entries. Final input symbolic links, chains and targets with linked ancestors
cannot be overwritten by PDF or CLI diagnostics publication. Input-entry checks
remain. Separate output file links and hardlinks can still be replaced atomically.
Existing directory/case alias rules, captured paths, callbacks and staging remain.
Concurrent external directory changes remain outside this path contract.

Cached-reference serving now requires a destination without PNGs, matching the
raster wrapper and direct validator. Reusing a longer reference run cannot leave
old pages beside a shorter cached PDF. Refusal preserves caller files and the
verified cache; use a fresh destination. The review's synthetic stale-output witness
does not show any previously accepted Office comparison to be invalid.

Deleted `WrapScanCounters`, both totals, increments, counting wrapper, parameter
plumbing and the single internal-total scaling test/helper. Useful opportunity
indexing, measured-character/allocation checks and semantic tests remain. Removed
the four 128-probe grow/shrink limits; cancellation replaces that work-only fuel.
The existing nonmonotonic-width search, Unicode boundaries, preferred breaks,
overwide fallback and source coverage remain. New tests cover long mixed-width
preferred/emergency growth, unfitting prefixes and cancellation during search.
No performance improvement is claimed and no replacement instrumentation was added.

Prior cleanup remains complete. There are no new public signature removals,
quotas, ownership layers or validation frameworks. Safe managed execution, useful
bounded caches, coarse admission, input bounds, bounded staging, deterministic
cleanup, atomic publication and PDF structural validity remain.

Live chart point-count, five-paragraph balloon and font/size/document-shape rules
remain documented debt; this milestone adds no such special cases. Strict DOCX
failures remain significant geometry/typography residuals: SVG Review Y 1.436035pt
exceeds 1pt; linked-footnote raw Word URI differs by 46.794pt although candidate
picture/link bounds agree within .000016pt. Two field/extraction flags remain.
None were relabeled cosmetic or passed.

## Retained candidate evidence

Ignored evidence: `artifacts/preview-input-wrap-20261021/qualification.json`;
frozen runtime under `candidate/`. Review and current/preceding negative witnesses
remain under `artifacts/review-eighth-20261021/`. The initial local capability-list
invocation rejected a shell array before tests; the corrected catalog-only check
passes, with no test replay. Both logs are retained. Earlier milestone evidence,
including the prior CI harness failure, remains unchanged.

Package: `artifacts/nuget/preview-020-input-wrap-20261021/`. Its DLL, symbol PDB and
four caller guides match the qualified checkout. Both consumer summaries pin the
same package; hashes and exact locations are in `package-identity.json`.
Publication remains deferred.

Latest infrequent workspace scan: **5.532 GB**, 861,208 files, zero errors.
This counts unique file-content lengths, not allocated clusters. It remains near
the loose 5 GB target; no historical purge or protected-target deletion occurred.
