# 0.2.0-preview qualification — 2026-10-21

The eleventh bounded milestone fixes async file-path binding and diagnostic
overflow semantics, with inherited rendering failures explicit below.
Publication, tags, releases, optimization and broader feature expansion remain
deferred. See [scope](PreviewRelease.md), [migration](PublicApiChanges.md) and
[document review](PreviewVisualReview.md).

The Release artifact is frozen at `ec9df9af38459d21c66cfa3efe4e128219d1e88d`:
24 assets and 2,582 verified shared public inputs. One final artifact was qualified
on Windows and Ubuntu, then transferred without a root library build or suite
replay. The same package was consumed on both platforms.

Library SHA-256:
`813AEDB19E3173E25231D326C06B3A30EF9B5561E2C1AA3AC2C414671A69408F`.
[Windows and Ubuntu CI passed](https://github.com/Lokad/OoxPDF/actions/runs/39624808363)
at `ec9df9af`. CI builds are independent of the locally qualified DLL.

| Gate | Result |
|---|---|
| Release build | Passed; two existing nullable warnings in unchanged DOCX tests. An initial new-test overload ambiguity was corrected before freezing |
| Windows milestone | 3,743 pass, **one inherited failure**, five skips; 3,749 unique methods |
| Same frozen DLL on Ubuntu | 923 pass, zero failures, 45 environmental skips; 968 unique methods |
| Focused checks | Windows: 64 pass/two expected filesystem skips. Ubuntu: 54 pass after retrying only two NTFS skips on ext4 |
| Qualification reuse | Affected sets excluded from broad remainders. Windows used two runner invocations; Ubuntu used three, including the two-check environment repair. Original reports retained; no passing methods replayed |
| Regression controls | New queued-path and Info-only-overflow tests both fail against the previous qualified DLL and pass against this candidate |
| Internal cleanup | Four unread arguments and one forwarding overload removed; font validation parses, baseline selection, repeated headers and secondary-axis behavior retained |
| API/callers | All 247 exported declarations unchanged. Migration and diagnostic guidance describe the behavior corrections |
| Representative documents | Nine sources/16 outcomes/32 pages; PDFs, diagnostics and semantics exactly match the preceding candidate. Each outcome converted once |
| Office comparisons | Reused by exact output identity: two PPTX strict passes/fourteen DOCX strict failures. No new Office export/raster or threshold change |
| Package consumers | Same .nupkg on Windows/Ubuntu; DLL/PDB, four guides, versions, embedded fonts and admission lifetime verified |
| Root transfer | All 24 assets verified; two integration PDFs match, without library rebuild or suite replay |
| Safe managed | Explicit unsafe=false, IL-only, no P/Invoke/imported COM; dependency-free runtime, assembly/file 0.2.0.0 |
| Materials/catalog | One garbled apostrophe repaired. All 715 manifests validate; all 287 capability selections resolve. Manifest validity is not Office fidelity |

Local milestone selections retain the prior explicit slow-test exclusions.
The inherited Windows test is
`DocxWordCompatibleBalloonBodyWrapsLikeOfficeTitles`: it observes a
12.610000000000001pt single-row body where its Office-derived expectation requires
another row. The previous candidate fails with the same name/detail. Its assertion
and rendering policy remain unchanged; this is a failure, not a passing check.
CI can skip it when its Office-font preconditions are absent. Historical
nonportable native-font failures remain unrerun evidence.

## Changes and preserved behavior

Asynchronous file conversions bind relative paths to the invocation directory
before waiting for a worker. Input lookup and output publication remain bound
through later directory changes. Invalid arguments still fault the returned task;
pre-cancellation takes precedence. The regression controls scheduling inside a
separate test process; no library scheduler hook was added.

Info-only diagnostic overflow now emits an Info marker, agreeing with the
collector's warning/error flag and CLI strict policy. Collections that observed
warnings/errors retain the existing Error overflow marker, including when the
original entry was dropped. The review reproduced the Info inconsistency through
the public collector; it did not demonstrate Info-only overflow in a normal CLI
conversion. No new diagnostic counter or ownership layer was introduced.

Removed arguments belonged to secondary chart axes, font discovery, repeated
table headers and the hhea-ascender forwarding overload. Live geometry rules,
discarded malformed-font validation parses and useful inspectors remain.
Chart point-count regimes, five-paragraph balloons and font/size/document-shape
guards remain unchanged. No performance gain or rendering expansion is claimed.

Safe managed execution, meaningful input bounds, cancellation, coarse admission
and refunds, bounded caches/page staging, deterministic cleanup, atomic publication
and PDF structural validation remain. Older broad capability keys and root-building
private-case tools remain deferred until affected; they were not used for this batch.

Fourteen DOCX strict gates retain typography/geometry failures. SVG Review Y
1.436035pt exceeds 1pt; raw Word footnote URI bounds differ by 46.794pt although
candidate picture/link bounds agree within .000016pt. Two field/extraction flags
remain. These results were not relabeled cosmetic or passed.

## Evidence and candidate

Ignored evidence: `artifacts/preview-api-boundaries-20261021/qualification.json`;
frozen runtime: `candidate-final/`. Original focused/milestone reports, the
two-check Ubuntu environment repair, regression controls and preparation logs
remain alongside it. The initial build attempt was not frozen or qualified.

Package: `artifacts/nuget/preview-020-api-boundaries-20261021/`. Package/symbol
hashes, qualified guide identities and both consumer summaries are recorded in
`package-identity.json`. The previous candidate and independent Office references
remain retained. Publication stays deferred.

Last recorded infrequent workspace scan: **5.864 GB**, 871,935 files,
zero errors. This counts unique content lengths, not allocated clusters. Keep
near the loose 5 GB target; useful inputs, evidence and Git history are preserved.
