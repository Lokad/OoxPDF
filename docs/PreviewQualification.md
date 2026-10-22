# 0.2.0-preview qualification â€” 2026-10-22

This bounded milestone fixes malformed JPEG metadata publication and simplifies
qualification mechanics. Publication, tagging, performance optimization and
broader feature expansion remain deferred. See [scope](PreviewRelease.md),
[migration](PublicApiChanges.md), [capabilities](Capabilities.md) and
[document review](PreviewVisualReview.md).

One production Release build at `a09de4b974705aed9dc0858b328815f065a9585a`
was qualified on Windows and Ubuntu. A test-only stream-input correction at
`047aebed5954be7754d34207006a8b11e52a7ca0` reused the identical library and 21
runtime assets; only two failed Windows checks reran. The final artifact contains
24 assets and 2,582 shared public inputs. No root library rebuild or broad replay.

Library SHA-256:
`8FEDE62C27ED02FFEC97704DF45A8EC4BC497AA86C5D49458167D8952557B6E7`.
[Windows and Ubuntu CI passed](https://github.com/Lokad/OoxPDF/actions/runs/39780521926)
at the corrected test revision. CI builds are independent of the locally qualified DLL.

| Gate | Result |
|---|---|
| Release build | Passed; two inherited nullable warnings in unchanged DOCX tests. Initial missing test namespace corrected before freezing; failed compile log retained |
| Affected set | 49 methods: Windows 48 pass/one filesystem skip after two test-setup repairs; Ubuntu 46 pass/three filesystem/PDFium skips |
| JPEG publication | Four real encoded controls retain bytes: baseline RGB, gray, CMYK and progressive. Malformed dimensions, precision and component tables diagnose/omit through DOCX/PPTX, including recolor. Four strict CLI probes now exit 3 with no image object; valid control PDF unchanged |
| Windows milestone | **3,759 pass, one inherited failure, five skips**; 3,765 unique methods, 3,767 attempts across three invocations |
| Same library on Ubuntu | **950 pass, zero failures, 49 skips**; 999 unique methods across two disjoint invocations; selection includes 15 additional capability neighbors beyond the preceding selection plus five new methods |
| Harness and tools | Long stderr refusal completes with exit 2; four basic CLI checks are FAST. Five handwritten text cases pass with zero builds; missing CLI/inspector bindings refuse before conversion |
| Representative documents | Nine sources/16 outcomes/32 pages; PDFs, diagnostics and semantic outcomes match the preceding candidate exactly. Each outcome converted once |
| Independent Office gates | Reused by PDF/input identity: **two PPTX strict passes/fourteen DOCX strict failures**; no new Office exports, portfolio rasters or relaxed thresholds |
| API/runtime | All 247 public declarations unchanged; IL-only, no P/Invoke/imported COM, explicit AllowUnsafeBlocks=false; assembly/file version 0.2.0.0 |
| Package/transfer | Same dependency-free 0.2.0-preview package passes both consumers, including loaded DLL identity, versions and admission release. All 24 root assets verified; two integration PDFs match |

## Behavior and limits

JPEG passthrough checks positive frame dimensions, 8-bit precision and consistent
supported component metadata before writing a PDF image dictionary. It retains
encoded bytes and does not decode or fully validate the entropy stream. Malformed
or unsupported metadata follows the existing `IMAGE_UNSUPPORTED_FORMAT` error
and image-omission policy; strict CLI conversion returns 3. No public API or
resource quota changes were introduced. Safe managed code, cancellation, bounded
storage and atomic publication remain required.

Test child processes drain both output streams concurrently and have bounded
waits. Explicit prebuilt validation-tool bindings fail if missing. CI's text gate
uses its frozen CLI and already built inspector instead of six extra MSBuild
invocations; default interactive execution builds each required tool once.

## Inherited failures and residuals

**DocxWordCompatibleBalloonBodyWrapsLikeOfficeTitles** still fails:
`InvalidOperationException: The second body should wrap to a second row like Office; observed Height=12.610000000000001.`
It remains a failure; CI can skip the Office-font precondition. Two repaired
stream-input tests had to pass; their initial failures were retained, not skipped.
Two local Ubuntu case-distinct-path checks remain skipped on the mounted
case-insensitive filesystem. PDFium checks require Windows; these are not passes.

The fourteen DOCX strict failures include significant geometry/typography.
SVG Review Y is 1.436035pt against 1pt; raw Word footnote annotation mismatch is
46.794pt while candidate picture/link agree within .000016pt. Two field/extraction
flags and historical nonportable font failures remain explicit. Chart-color and
font/size/document-shape guards still require independent Office evidence before
changes. The warning-only partial MCE policy and standalone enum-helper deferral
remain documented; completed ownership/validation cleanup stays closed.

## Evidence and workspace

Ignored evidence: `artifacts/preview-image-boundaries-20261022/qualification.json`;
final runtime: `candidate-qualified/`; package:
`artifacts/nuget/preview-020-image-boundaries-20261022/`.
Original compile and test failures, repaired exact checks, joined reports,
JPEG probes, tool bindings, package identities and transfer evidence are retained.
The [preceding qualification](https://github.com/Lokad/OoxPDF/blob/2a8a99ed17cc1505ad2747a7bce54ed10c1a128d/docs/PreviewQualification.md)
retains compatibility and independent Office comparison history.

Three completed cleanup inventories were compressed with round-trip hash
verification, reclaiming **148,475,109 bytes**. Only those three originals were
removed; useful image/BIN/TAR evidence and prior qualified candidates remain.
Workspace measurement before this batch was 6.004 GB; the loose target is 5 GB.
Final measurement: **5.943 GB / 885,270 files**, no errors;
the report is retained. The workspace remains above the loose target; do not repeat already completed cleanup or
purge retained archives by their outdated binary filenames.
