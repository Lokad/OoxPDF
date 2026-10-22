# 0.2.0-preview qualification — 2026-10-22

This bounded milestone fixes PNG transparent-color interpretation and malformed
header recovery, preserves Office reference failure attribution, and removes
small stale guidance/comments. Publication, tagging, performance optimization
and broader feature expansion remain deferred. See [scope](PreviewRelease.md),
[migration](PublicApiChanges.md), [capabilities](Capabilities.md) and
[document review](PreviewVisualReview.md).

One committed Release artifact at `7cd9e45ceb89879162c9819c777fdc4824a5d5ac`
contains 24 runtime assets and 2,582 shared public inputs. It was qualified on
Windows/Ubuntu, consumed through the same package on both platforms, and
transferred with two integration smokes. No root rebuild, local method replay
or test-assembly repair was needed.

Library SHA-256:
`3EB2BCCBB4F213DB70FDC79E7E811941E9E27922AB884A05CD21F75EFCEE593C`.
[Windows and Ubuntu CI passed](https://github.com/Lokad/OoxPDF/actions/runs/39811906403)
at that source revision. CI builds are independent of the locally qualified DLL.

| Gate | Result |
|---|---|
| Release/affected checks | Build passes with two inherited nullable warnings; all 51 exact affected methods pass on both platforms, no skips |
| Windows milestone | **3,767 pass, one inherited failure, five skips**; 3,773 unique methods in two disjoint invocations |
| Same library on Ubuntu | **962 pass, zero failures, 49 skips**; 1,011 unique methods in two disjoint invocations, including four additional PNG neighbors beyond the prior selection |
| PNG semantics | Original-sample/alpha controls cover supported gray depths and RGB, ordinary/Adam7 data and equivalent RGBA output through DOCX/PPTX/crop/recolor |
| Pinned public probes | Eleven conversions; all four malformed headers return strict 3 with one image Error/no image object. Keyed/RGBA PDFs agree; four already-working control PDFs are unchanged |
| Independent PNG comparison | Two retained PowerPoint16 references: candidate/reference rasters are pixel-identical at 72 dpi; existing raw and support numerical thresholds pass |
| Representative documents | Nine sources/16 outcomes/32 pages; PDFs, diagnostics and semantic outcomes match the preceding candidate exactly; each outcome converted once |
| Independent portfolio gates | Reused through PDF/input identity: two PPTX strict passes/fourteen DOCX strict failures; no new Office exports or portfolio rasters |
| Reference tooling | Seven supervision/eight actual-worker failure-stage checks pass locally on Windows and in Windows/Ubuntu CI; no COM or build needed for synthetic controls |
| API/runtime | All 247 public declarations unchanged; IL-only, no P/Invoke/imported COM, explicit AllowUnsafeBlocks=false; assembly/file version 0.2.0.0 |
| Package/transfer | Same dependency-free 0.2.0-preview package passes both consumers; all 24 root assets/two integration PDFs match; all 192 checked retained/current/root assets retain their identities |

## Behavior and supported scope

RGB transparent-color keys now create the same alpha plane as equivalent RGBA
images. Packed grayscale keys compare original samples before expansion to RGB;
unused high bits in the key are masked to the supported sample depth. Existing
palette/full-alpha behavior remains covered. This adds no codec, color-management
framework, public API, quota or supported bit depth.

The retained truecolor probe previously painted a green opaque area; the one-bit
gray probe painted white. Both now reveal the gold slide background exactly like
PowerPoint. Native PDFium comparison at 72 dpi has MAE 0, SSIM 1, histogram
correlation 1 and foreground recall 1 for each probe. Input/reference PDF/PNG
identities are retained. These two synthetic probes establish their specific
outcomes; they do not establish general PNG conformance or Office fidelity.
No visual threshold was relaxed.

Fixed IHDR size and transparent-key shape are checked before field access.
Malformed headers now use the existing IMAGE_UNSUPPORTED_FORMAT Error/omission
policy instead of an unhandled indexing exception. Unsupported compression/filter
methods are explicit. Scanline recovery, cancellation, original observer exceptions
and existing-file preservation remain covered. There is no blanket exception catch
or complete PNG validator; sixteen-bit samples remain unsupported.

Successful Office cleanup restores the pending operation's failure stage. A real
cleanup failure remains attributed to cleanup. Actual-worker controls cover
activation, open, export and cleanup for Word and PowerPoint. Local WSL has no
PowerShell installed, so those tool checks use Ubuntu CI; local Ubuntu still
qualifies the frozen library and package. The pinned-PDFium recovery hint now agrees
with its version/hash instructions. Thirty-seven detached chart-comment/blank
lines were removed; no live chart or DOCX layout rule changed.

## Inherited failures and deferred concerns

**DocxWordCompatibleBalloonBodyWrapsLikeOfficeTitles** still fails:
`InvalidOperationException: The second body should wrap to a second row like Office; observed Height=12.610000000000001.`
It remains a failure; CI may skip its Office-font precondition. Local Ubuntu
filesystem/PDFium/font skips remain explicit and are not counted as passes.

The fourteen DOCX strict failures include significant geometry/typography.
SVG Review Y remains 1.436035pt against 1pt; raw Word footnote annotation mismatch
is 46.794pt while candidate picture/link agree within .000016pt. Two field/extraction
flags and historical font limitations remain. The preceding CMYK qualification
retains histogram correlation .311828 against .98: needs-review, without a complete
ICC/YCCK color-parity claim. This PNG batch does not resolve that difference.

Transparency-group rollback retains an output charge in an internal probe; no
public document sequence was established. That investigation remains deferred,
with runtime accounting unchanged. Existing partial MCE, enum-helper and qualified
chart/font/size/document-shape limitations remain documented. Completed ownership
and validation simplification stays closed.

## Evidence and workflow

Ignored evidence: `artifacts/preview-png-semantics-20261022/qualification.json`;
frozen runtime: `candidate-final/`; package:
`artifacts/nuget/preview-020-png-semantics-20261022/`.
Exact selections, public image/reference identities, raw native metrics, package
consumers, stage controls and transfer evidence are retained. One local catalogue
invocation failed native PowerShell array binding before execution; the corrected
in-process invocation validates the same frozen assembly. No library rebuild or
method rerun was required. Original error evidence remains available.

The [preceding qualification](https://github.com/Lokad/OoxPDF/blob/e817a2fe/docs/PreviewQualification.md)
retains the CMYK/scanline/metric milestone and earlier Office-comparison history.
The rendering roadmap remains preserved. Latest full workspace measurement stays
**5.943 GB/885,270 files**, above the loose 5 GB target; no new full walk or purge
was repeated. This milestone directory measures about **49 MB/3,860 files**;
reuse the worktree, fixture snapshot and immutable outputs.
