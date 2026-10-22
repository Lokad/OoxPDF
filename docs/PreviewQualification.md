# 0.2.0-preview qualification — 2026-10-22

The result-integrity milestone repairs qualification joins and removes ten
verified unused internal methods. Publication, tags, releases, optimization and
broader feature expansion remain deferred. See [scope](PreviewRelease.md),
[migration](PublicApiChanges.md) and [document review](PreviewVisualReview.md).

One Release artifact is frozen at `446f6b9553f95159c068ea04cd9af2f34497844e`:
24 assets and 2,582 verified shared public inputs. The same DLL was qualified on
Windows and Ubuntu, consumed through the same package on both platforms, then
transferred with two integration smokes. There was no root library rebuild or
broad suite replay.

Library SHA-256:
`98AB5D538F957F4E1874525FDF009D0290F9A0480F274FB6477146D8E859693F`.
[Windows and Ubuntu CI passed](https://github.com/Lokad/OoxPDF/actions/runs/39751506245)
at the frozen source revision. The [tool-fixture follow-up](https://github.com/Lokad/OoxPDF/actions/runs/39753866703)
also passed at `cb1f8ee7`; library/test sources are unchanged. CI builds are
independent of the locally qualified DLL.

| Gate | Result |
|---|---|
| Release build | Passed; two existing nullable warnings in unchanged DOCX tests |
| Windows milestone | 3,743 pass, **one inherited failure**, five skips; 3,749 unique methods |
| Same frozen DLL on Ubuntu | 923 pass, zero failures, 45 expected environmental skips; 968 unique methods |
| Focused live-neighbor checks | Windows 33 pass; Ubuntu 19 pass/two expected font skips; one runner invocation per platform |
| Qualification reuse | Affected sets excluded from broad remainders; two runner invocations per platform, no repeated methods |
| Tool regressions | Both platforms pass portfolio conflict/ownership checks, test-repair failure/provenance checks and tiny-directory measurement/refusal checks |
| Internal cleanup/API | Ten exact unused signatures removed; all 247 exported declarations unchanged. Eleven useful inspectors, three color snapshots and the reflection-backed chart projection retained |
| Representative documents | Nine sources/16 outcomes/32 pages; PDFs, diagnostics and semantic results exactly match the preceding candidate. Each outcome converted once |
| Office comparisons | Reused by exact PDF identity: two PPTX strict passes/fourteen DOCX strict failures; no new Office export/raster or threshold change |
| Package consumers | Same .nupkg on Windows/Ubuntu; qualified DLL/PDB, four guides, versions, embedded fonts and admission lifetime verified |
| Root transfer | All 24 assets verified; two integration PDFs match without library rebuild or suite replay |
| Safe managed | Explicit unsafe=false, IL-only, no P/Invoke/imported COM; dependency-free runtime, assembly/file 0.2.0.0 |

Local milestone sets retain the prior explicit slow-test exclusions. The inherited
Windows failure remains `DocxWordCompatibleBalloonBodyWrapsLikeOfficeTitles`: a
12.610000000000001pt single-row body where its Office-derived expectation requires
another row. The preceding DLL fails with the same name/detail. The assertion and
rendering rule remain unchanged; the failure is not counted as a pass. CI can skip
its Office-font precondition. Historical nonportable native-font failures remain
unrerun evidence.

## Qualification behavior and preserved contracts

Portfolio joins require one authoritative report per document/mode outcome.
Repeating the same report selection is allowed; duplicate outcome rows, including
identical PDFs from distinct reports, are refused before publication. This prevents
source order from hiding a contradictory content blocker.

Test-report joins remember failures observed in any attempt. A later skip cannot
resolve a failed initial or repair check; a subsequent pass can. Initial-failure
provenance remains separate and original attempts remain retained. Neither defect
occurred in the preceding candidate's actual reports. Focused synthetic controls
cover both counterexamples without builds or document conversions.

Workspace measurement rejects unknown arguments; the supported report argument is
`-ReportPath`. Diagnostic JSON keeps its existing numeric severity format. The
preview scope guide now links completed maintenance history rather than repeating
it. These changes add no public API removal, rendering rule or performance claim.

Safe managed execution, input bounds, cancellation, coarse host admission/refunds,
bounded caches/page staging, deterministic cleanup, atomic publication and PDF
structural validation remain. Point-count chart colors, exact-five-paragraph
balloons and font/size/document-shape guards are unchanged. Older broad capability
keys and root-building private-case tools stay deferred until affected.

Fourteen DOCX strict gates retain typography/geometry failures. SVG Review Y
1.436035pt exceeds 1pt; raw Word footnote URI bounds differ by 46.794pt although
candidate picture/link bounds agree within .000016pt. Two field/extraction flags
remain. No thresholds or failure classifications changed.

## Transfer tooling follow-up

The current runtime is reused for the transfer-ownership follow-up. Root assets
are copied into owned staging files, checked, then published by replacing the
destination entry. Hard-link referents remain unchanged, linked asset directories
refuse during preflight, and failure removes only owned staging files.

Seven focused controls pass on Windows and Ubuntu: two successful ownership
cases and five publication, smoke or preflight refusals. The original transfer
tool fails the hard-link regression. A live transfer verifies all 24 existing
assets and two identical integration PDFs; no local build, broad suite, corpus,
Office export or raster was repeated. Both retained candidates still match their
48 pinned asset hashes.

Caller guidance now states that local font files remain unchanged while their
sources/resolver snapshots are active. Discovery invalidation applies to newly
created resolvers. This adds no runtime or public API change. The existing
candidate package keeps its originally qualified guide bytes; clarified repository
guidance will enter the next packaging milestone. No package was rebuilt.

Evidence: `artifacts/preview-transfer-ownership-20261022/reuse-evidence.json` and
its focused tool reports. [Tooling CI](https://github.com/Lokad/OoxPDF/actions/runs/39760494948)
uses `725a616f`; library/test sources remain identical to the frozen runtime.
The low-priority standalone enum-helper correction is deferred to a meaningful
API change.

## Evidence and candidate

Ignored evidence: `artifacts/preview-result-integrity-20261022/qualification.json`;
frozen runtime: `candidate-final/`. Exact selections, disjoint platform reports,
tool checks, package identities and transfer evidence remain alongside it.
An initial portfolio inspection used the wrong inspector; its successful first
PDF was retained and reinspected, with no repeated conversion. That outcome's
original cold-process timing is unavailable and is recorded as null.

Package: `artifacts/nuget/preview-020-result-integrity-20261022/`. DLL, symbol,
qualified guide and consumer identities are recorded in `package-identity.json`.
Publication remains deferred. The last full workspace scan remains the
2026-10-21 measurement: **5.864 GB**, 871,935 files, zero errors. No new full scan
or unverified purge was performed; keep near the loose 5 GB target.
