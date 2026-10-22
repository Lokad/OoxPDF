# 0.2.0-preview qualification — 2026-10-22

This bounded milestone corrects Adobe CMYK sample interpretation, restores
consistent PNG scanline recovery, and removes conversion/native dependencies from
six visual-metric checks. Publication, tagging, performance optimization and
broader feature expansion remain deferred. See [scope](PreviewRelease.md),
[migration](PublicApiChanges.md), [capabilities](Capabilities.md) and
[document review](PreviewVisualReview.md).

One committed Release artifact at `4d22dee69cd520ab13efda87b0ac14662221394e`
contains 24 runtime assets and 2,582 shared public inputs. The same library was
qualified on Windows/Ubuntu, consumed through the same package on both platforms,
and transferred with two integration smokes. No root rebuild, local method replay
or test-assembly repair was needed.

Library SHA-256:
`C15DE4B65F894D0A5DB0EEA4C09DE01E0EB8F9473B5208E43D7EBF62A34F61C8`.
[Windows and Ubuntu CI passed](https://github.com/Lokad/OoxPDF/actions/runs/39792238334)
at catalogue-only follow-up `5352c8b6`; compiled source is unchanged from the
frozen artifact. CI builds are independent of the locally qualified DLL.

| Gate | Result |
|---|---|
| Release/affected checks | Build passes with two inherited nullable warnings; all 44 exact affected methods pass on both platforms, no skips |
| Windows milestone | **3,764 pass, one inherited failure, five skips**; 3,770 unique methods in two disjoint invocations |
| Same library on Ubuntu | **955 pass, zero failures, 49 skips**; 1,004 unique methods in two disjoint invocations |
| CMYK interpretation | Actual Adobe CMYK painting, marker ordering and crop/recolor fallback checked; non-CMYK behavior and sample-mapping identity controls pass. Existing RGB control PDF unchanged |
| PNG recovery | Eight retained CLI control/truncation probes follow image policy; malformed payloads exit 3 with one image Error and no image object. Container/effect, cancellation and publication checks pass |
| Metric simplification | All six page/region recall and histogram premises pass portably using direct PNGs and prebuilt VisualDiff; removed 140-line script and its three builds/two DOCX conversions/two native rasters |
| Representative documents | Nine sources/16 outcomes/32 pages; PDFs, diagnostics and semantic outcomes match the preceding candidate exactly; each outcome converted once |
| Independent Office | Original portfolio retains two PPTX strict passes/fourteen DOCX strict failures through PDF/input identity; existing PowerPoint16 CMYK reference reused separately |
| API/runtime | All 247 public declarations unchanged; IL-only, no P/Invoke/imported COM, explicit AllowUnsafeBlocks=false; assembly/file version 0.2.0.0 |
| Package/transfer | Same dependency-free 0.2.0-preview package passes both consumers; all 24 root assets/two integration PDFs match; retained artifacts preserved |

## Behavior and remaining color differences

Adobe four-component CMYK/YCCK headers carry an inverted PDF Decode mapping.
The header reader recognizes Adobe metadata before or after SOF and stops before
entropy data; RGB/gray and untagged interpretation remain unchanged. Unsupported
four-component Adobe transforms diagnose explicitly. The mapping participates in
resource identity, including original-image crop/recolor fallback. This does not
add a JPEG codec, ICC framework, public API or quota. Transform2 metadata is tested;
no new YCCK pixel fixture or complete Office color-management parity is claimed.

The retained solid CMYK probe previously painted nearly black: (0,10,0). It now
paints (212,68,112), against PowerPoint's (205,70,103), using the same independent
PDF renderer. Sample channel residual is nine levels. Native PDFium comparison
has MAE 0.27, SSIM 0.9999896, foreground recall 1 and no pixels differing by 16
channel levels; it passes the existing PPTX raw pixel gates. Histogram correlation
is **0.311828 against the unchanged 0.98 support threshold**, so the support
classification remains **needs-review**. The gross inversion is fixed; remaining
color approximation is explicit. The 16-document portfolio totals exclude this
separate probe, and no visual threshold was relaxed.

Incomplete PNG filter/row data now raises InvalidDataException at the decoder and
uses `IMAGE_UNSUPPORTED_FORMAT` Error/omission consistently. This removes the
reproduced unhandled CLI termination and generic PPTX node warning for those
payloads. Observer failure/cancellation still preserves existing file output.
Three decoder tests now assert the recovery contract; one exact duplicate was
removed. Unexpected programming exceptions are not hidden by a blanket catch.

## Inherited failures and deferred investigation

**DocxWordCompatibleBalloonBodyWrapsLikeOfficeTitles** still fails:
`InvalidOperationException: The second body should wrap to a second row like Office; observed Height=12.610000000000001.`
It remains a failure; CI may skip its Office-font precondition. Local Ubuntu
filesystem/PDFium/font skips remain explicit and are not counted as passes.

The fourteen DOCX strict failures include significant geometry/typography.
SVG Review Y remains 1.436035pt against 1pt; raw Word footnote annotation mismatch
is 46.794pt while candidate picture/link agree within .000016pt. Two field/extraction
flags and historical font limitations remain. No chart/balloon/layout guard changed.

Transparency-group rollback demonstrably retains an output-content charge in an
internal probe. Bounded tracing found no public document example; runtime accounting
is unchanged and the concern remains deferred. No refund ledger, work fuel or
ownership framework was added. Existing partial MCE and enum-helper limitations
remain documented; completed simplification stays closed.

## Evidence and workflow

Ignored evidence: `artifacts/preview-image-paint-20261022/qualification.json`;
frozen runtime: `candidate-final/`; package:
`artifacts/nuget/preview-020-image-paint-20261022/`.
Exact selections, image/reference identities, raw color metrics, package consumers
and transfer evidence are retained. The first CI attempt found four stale PNG
selectors after test consolidation. A manifest-only follow-up validates the entire
capability catalogue on the same frozen assembly; no local method rerun or library
rebuild was required. Original CI failure evidence is retained.

The [preceding qualification](https://github.com/Lokad/OoxPDF/blob/6804a374/docs/PreviewQualification.md)
retains prior cleanup and Office-comparison history. The rendering roadmap is
preserved. Latest full workspace measurement remains **5.943 GB/885,270 files**,
above the loose 5 GB target. This new milestone measures about **49 MB/3,723 files**;
no repeated full walk or blind purge was performed. Reuse the worktree, fixture
snapshot and immutable outputs.
