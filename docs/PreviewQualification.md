# 0.2.0-preview qualification — 2026-10-21

The ninth bounded cleanup milestone is qualified with the inherited rendering
failures below. Publication, tags, releases, optimization and broader feature
expansion remain deferred. See [scope](PreviewRelease.md), [migration](PublicApiChanges.md)
and [document review](PreviewVisualReview.md).

The final Release artifact is frozen at
`4d5f7601a28fd625b70d0485c1b52c6e7739e85f`: 24 assets and 2,582 verified shared
public inputs. Both platforms' broad gates use this artifact once. Preparatory
snapshots and a failed reflection-based test are retained separately. Its repaired
test passed against the unchanged preparatory DLL; subsequent removal of two
orphan palette wrappers produced the final artifact before broad qualification.
No root build or broad-suite replay occurred.

Library SHA-256:
`44005AD1D84A9F9F9F230E66C0282B648ADBEB8710BD3F4FD54375EFDA63834D`.
[Windows and Ubuntu CI passed](https://github.com/Lokad/OoxPDF/actions/runs/39610347935)
at `4d5f7601`. CI outputs are independent of the locally qualified DLL.

| Gate | Result |
|---|---|
| Final Release build | Passed; two existing nullable warnings in unchanged DOCX tests |
| Windows milestone | 3,740 pass, **1 inherited fail**, 5 skips across 3,746 unique methods |
| Same frozen DLL on Ubuntu | 913 pass, 0 fail, 45 environmental skips across 958 unique methods |
| Focused final checks | Windows all 68 pass; Ubuntu all 58 pass. Exact inactive-helper, sparse/provenance, format, theme, cache, geometry and markup selections |
| Test reuse | Each platform's focused methods excluded from its remaining broad selection; two disjoint invocations, no method replay on the final artifact |
| Source cleanup | 45 source methods, 14 retired/dependent constants and two parse counters removed; live implementations and useful inspectors retained |
| Caller verification | Signature-level scan of 25,799 method bodies, no unresolved operands and no remaining authored private zero-caller candidates. This is not proof that every internal/public path is necessary |
| Public compatibility | All 247 exported declarations match the preceding candidate; no additional public API removal |
| Allocation tool compatibility | Probe built against the frozen DLL without a library build; self-test and live font-preparation stage smoke pass. Whole/staged PDF identities match; no performance improvement claimed |
| Representative documents | Nine sources / 16 outcomes / 32 pages; every PDF, diagnostic list and semantic result matches the preceding candidate. Each final outcome converted once |
| Independent Office evidence | Portfolio comparisons reused by identity: two PPTX strict passes / fourteen DOCX strict failures. Earlier mixed-width Word probe also retains exact candidate PDF identity. No new Office export, portfolio raster or threshold change |
| Package consumers | Same .nupkg on Windows and Ubuntu; DLL/PDB identity, versions, four guides, embedded fonts and admission lifetime pass |
| Root transfer | All 24 assets verified; two integration PDFs match, without a root build or repeated suite |
| Safe managed | Explicit unsafe=false, IL-only, no P/Invoke/imported COM; dependency-free runtime, assembly/file 0.2.0.0 |
| Materials/catalog | Checkpoint reduced from 9,570 to 33 lines with Git-history link. All 715 visual manifests validate and capability selectors resolve; manifest validity is not Office fidelity |

The inherited Windows balloon assertion still observes a 12.61pt single-row body
where its Office-derived expectation requires another row. Its method name and
failure detail match the preceding qualification. Its assertion and policy remain.
Historical nonportable native-font failures remain unrerun evidence. These local
failures/skips are separate from green CI.

## Changes and retained behavior

DOCX painting and emitted-end measurement now use authored runs directly instead
of cloning equal runs through retired zero-valued spacing/offset corrections.
Authored spacing, revision colors, terminal-space behavior, kerning and nonzero
balloon geometry remain. The zero-only balloon forwarding layer is removed.

Unused exact overloads, the obsolete PPTX word wrapper and default-context readers
are removed. Live explicit color-map/context implementations remain. Tests and the
allocation probe target live cores; source-block geometry uses explicit fixture
expectations, and sparse/provenance tests exercise compact values and the actual
rendering index. Old numeric dense/lookup paths and parse-count instrumentation
are removed without retaining copied algorithms or adding replacement counters.

Chart point-count regimes, five-paragraph balloons and font/size/document-shape
guards remain unchanged. They are live debt requiring independent Office evidence
before replacement. Safe managed execution, meaningful input bounds, cancellation,
coarse admission/refunds, useful bounded caches, deterministic cleanup, bounded
page staging, atomic publication and PDF structural validity remain.

Fourteen DOCX strict gates still fail on typography/geometry. SVG Review Y
1.436035pt exceeds 1pt; raw Word footnote URI bounds differ by 46.794pt although
candidate picture/link bounds agree within .000016pt. Two field/extraction flags
remain. None were relabeled cosmetic or passed.

## Retained evidence

Ignored evidence: `artifacts/preview-inactive-paths-20261021/qualification.json`;
final runtime under `candidate-final/`. Per-signature deletion records, source-
block characterization, original failed checks, corrected checks and preparatory
artifacts remain. `preparation-evidence.json` distinguishes them from the final
gates; an aggregation path correction reused existing reports without rerunning
tests. No broad suite ran against the preparatory artifacts.

Package: `artifacts/nuget/preview-020-inactive-paths-20261021/`. Package/symbol hashes,
qualified guide identities and both consumer summaries are in `package-identity.json`.
The previous candidate and its independent Office references remain retained.
Publication stays deferred.

Latest infrequent workspace scan: **5.658 GB**, 864,892 files, zero errors.
This counts unique content lengths, not allocated clusters. Keep near the loose
5 GB target; no historical purge or protected-target deletion occurred.
