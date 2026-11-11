# 0.3.0 development qualification

Updated 2026-11-11. **0.3.0 is not published and is not a completed release
candidate.** The published [0.2.0 qualification](PreviewQualification.md) remains
separate. Development targets .NET 10 and .NET 11.

The aligned-cell follow-up (`f5509b8ab`) keeps the shared printed frame for complete,
measured centered/bottom-aligned cells. Four Review cases previously let pictures
cover six caption letters; native picture-knockout checks reveal 512–514 hidden
dark pixels per case. All eight Final/Review outcomes now retain complete glyphs
and pictures without measured intersections. The regression fails on the parent;
91 Windows checks pass, and Ubuntu passes 74 with 17 declared environment skips.
All sixteen portfolio PDFs/diagnostics and four existing PNG/SVG alignment controls
remain exact. Nested and legacy-only alignment keep their qualified fallbacks.
New-source Word geometry/application view remain unqualified; no Office action.
The frozen .NET 11 library SHA-256 is
`303EC4FF9B7E47E0EACBA376152B0C7AEC27E3C7DC65ED8001F54A0BFDA576F2`.
CI43062947587 on `e7d591999` is green: Windows 4,800 passes/19 skips, Ubuntu
2,289/130, 69 .NET 10 checks per host, four package consumers and ten package
inspections per host. The prepared package below now uses that exact CI archive.

Six footer-cell outcomes retain 21,483 bound glyphs, all 24 source PNG paints,
repeated captions and the correct PAGE/NUMPAGES values without measured cuts or
picture/ink intersections. Body/note links remain unchanged; these sources do not
author footer hyperlinks. The new portable check passes both hosts and the parent,
so no additional footer defect is claimed. One harness-only build reuses the same
library; two test API compile mistakes and their failed attempts are retained.
Joined focused coverage is 92 Windows passes and 75 Ubuntu passes/17 declared
skips, without replaying the preceding 91 checks. This test addition follows the
broader CI run; no second full CI or new library qualification is claimed.

The preceding rendering slice (`044032aaf`) publishes the existing validated Review
table frame for complete rows on every page. A long table previously let pictures
cover caption glyph bottoms although text-count and clipping checks passed. Native
ink/raster checks now retain 216 body-row paints, twelve mixed headers and all
17,992 measured glyphs; eight causal outcomes have no picture/text intersections.
The new portable regression fails on the parent. Focused .NET 11 checks pass 95 on
Windows; Ubuntu passes 82 with 13 font skips (two indirect guards reviewed and
confirmed after the initial strict skip gate). Sixteen portfolio PDFs/diagnostics
and the authored six-name legend remain byte-identical; their two PPTX passes and
fourteen DOCX image failures are retained. No new Office/view parity is claimed.
The frozen library SHA-256 is
`DE83321FD2123FCB232056F935194EF84832D24C9C60443083510447D13867D2`.
Broader CI43041456284 on `03ef23fcc` is green: Windows 4,799 passes/19 skips,
Ubuntu 2,288/130, 69 .NET 10 checks per host, four package consumers and ten package
preflight checks per host. Production/test source matches the frozen slice;
CI builds have their own revision metadata. This does not complete the roadmap.
The package refresh below uses the current CI production payloads; the previous
prepared archive is retained as parent evidence. Publication stays deferred.

Additional kept-cell checks move complete prose/formula/picture rows to a fresh
page in both modes, with the following paragraph intact. Eight outcomes retain
29,110 glyphs without clip cuts/unknown traces and complete opaque pictures;
the portable regression passes on both hosts and fails on the parent. This extends
the existing fix's evidence without another library change or Office export.

The previous local rendering baseline pins production
`c1b59358b13c7c6b6492f8ab7cd20a7881bc1d2a`; harness
`7196e18eb13656c6b6f233cfd98db470b0818786` repairs only a test fixture.
The frozen .NET 11 library SHA-256 is
`840C42E925641050D35B2F697F0450E0DFA460577B94339F9A373F1312056C08`.
Documentation-only updates reuse that frozen library.

## Verified outcomes

| Gate | Result and scope |
|---|---|
| Selected comment-view diagnostics | Final/Original and document settings that hide comments no longer report intentional omission as unsupported. Visible SimpleMarkup/AllMarkup comments retain their approximation warning and related-part provenance. The CLI still publishes PDFs and returns 3 for genuine warnings. Eight causal PDFs and all sixteen portfolio PDFs are byte-identical; four causal/seven portfolio legacy warnings are removed. Both hosts pass 60 of 61 exact checks with one declared case-insensitive-filesystem skip. A single font-dependent CLI fixture failed on Linux; a test-only repair isolates strict exits with a textless anchor, while reader/rendering tests retain prose/body coverage. Only that method was rerun per host; prior reports remain retained and joined after runtime/input identity checks. Twenty-four transferred assets and two smoke PDFs match. No root build, rerasterization or Office action. Broader CI43015189979 is green; release scope remains open. |
| Words after fixed-line picture overhangs (preceding batch) | Four oversized-picture outcomes hid four/five letters of After. Captured ordinary content now retains legacy width-fitted picture paint inside the shared source cursor, allowing the following word to wrap. All eight causal outcomes retain 40 glyphs without measured cuts or unknowns; picture bounds, visible fractions and source colors remain unchanged. Four terminal-overhang PDFs remain exact. Both static paths share the original baseline clip, retaining the 22pt-line/17.6pt-clip assertion. Portable coverage includes two fonts, body/cells, both presets, source-equivalent split runs, whitespace, bookmarks and cancellation. Exact new-source Word geometry is unqualified. |
| Contained fixed-line pictures (preceding batch) | Four fitting exact16 picture/prose outcomes hid four letters of After despite complete PDF operations. Captured ordinary rows now reuse shared contained-picture wrapping with the container's physical baseline, preserving fixed advance. All four words and authored picture dimensions are restored; six required outcomes retain 30 glyphs without measured cuts or unknown traces. Both first-frame and oversized fallback PDFs remain identical. The rendered regression fails against the parent. Existing independent exact-frame Word gates passed unchanged; new-source Word geometry remains unqualified. |
| Minimum-spaced pictures (preceding batch) | Changing only automatic spacing to atLeast16/32 left a partial A and hid the other four letters of After in four cell/preset outcomes. The reader now captures ordinary picture source items through the existing publication helper; captured ordinary content shares automatic/minimum measurement. All four words and authored 240x18pt pictures are restored, with all source color bands and 30 glyphs clear of measured clips. Automatic and exact-height control PDFs remain identical. Portable coverage checks two fonts, body/cells, both presets, split runs, declared minimum frames and cancellation. Reconstructed minimum frames keep their qualified fallback; math/box/exact eligibility remains distinct. Two regressions fail against the parent. Exact new-source Word geometry is unqualified. |
| Leading pictures and following words | Two narrow-first-line cell controls clipped four letters of After despite complete source operations. Owned automatic-flow pictures and prose now share the existing ordered cursor; the word and full-size picture move into the continuation frame. All six causal outcomes retain the word, complete picture bounds and source colors; 30 actual-program glyphs have no clip cuts or unknowns, and the unindented PDF stays exact. Picture paint height stays separate from occupied descent. Three pinned Word controls preserve the original 0.1pt dimension/0.75pt baseline gates. Shared width-roundoff handling also removes an unnecessary cell line without changing the original hhea/image-baseline checks. At that milestone, declared, unowned and unfit rows retained their fallbacks; the fitting minimum/fixed rows are qualified separately above. No exact new-source Word/view or arbitrary overpaint claim. |
| Autofit paragraph insets | A 72pt first-line indent hid every glyph of a short label while text operations remained present. Natural cell widths now fund the same paragraph offsets used by placement, including positive right insets and measured inline content. Both missing labels recover the controls' 270 raster ink pixels; four control PDFs are unchanged and 30 actual-program glyphs have no measured cuts/unknown traces. Portable checks cover two fonts/presets, split runs, grid hints, left/first/right insets and cancellation; the regression fails on the parent. No all-content, compressed-width or exact new-source Word/view claim. |
| Terminal pictures | Fitting terminal/trailing-space pictures escaped width admission: body retained 69.98% of bounds and cells 39.78%. Both wrapping paths now share source-boundary admission, including final rows. Adjacent drawings use their combined advance at a shared offset, correcting a trial that retained only 8.49% of the second cell picture. Eight pictures in six native outputs retain complete bounds and all source colors; 636 actual-program glyphs have no measured clip/unknown traces. Portable regressions cover both presets, two pinned fonts, merged/split runs and adjacent pictures; parent/trial failures are retained. Unfit groups keep their existing fallback; arbitrary group splitting and exact new-source Word/view geometry remain unqualified. |
| Pictures after soft breaks | Two/three supported soft breaks caused an opaque inline picture to cover part of the following word, although the whole word remained extractable. Picture attachment now reuses wrapped source starts and per-line source offsets, removing duplicate painted-length reconstruction. Four native controls retain all source words/pages with no following-word overpaint; the unchanged no-break PDF is identical. The screen binds 44 glyphs without clip cuts or unknown traces; paired crops confirm restored visible words. Portable checks cover pinned Sans/Math, both presets and separate/merged runs; the regression fails against the parent. No new-source Word/application-view geometry claim. |
| Pictures inside text runs | Merging equally formatted text/picture/text runs moved PNG/SVG pictures about 35.6pt left and shifted preceding prose behind them by the picture advance. The reader now retains text child boundaries, and shared wrapping/placement uses recorded source run numbers and offsets. Both corrected merged-run PDFs exactly match their original controls. Two portable regressions cover pinned Sans/Math fonts, Final/Review, multiple drawings in one run and cancellation; both fail against the parent. The affected native screen measures 22 glyphs without cuts or unknown traces. No ordinary-layout migration or new layout layer. |
| Ordered picture rotation | Shared placement publishes the measured line frame, restoring six authored quarter-turns lost beside mathematics in body/header/footer content. Eight controls bind 546 glyphs without cuts or unknown text traces. The global transformed-origin change is reverted: five existing source-pinned Word PDF comparisons exposed ordinary PNG/SVG regressions, including a previously passing anchored picture. All five corrected PDFs are byte-identical to their earlier qualified outputs, restoring one pass and retaining four failures. A portable regression checks the Word anchor within the original 0.001pt gate with two pinned fonts and both modes. Source-fit dimensions, source-centered rotation, clips and invalid-frame fallback remain. |
| Actual legend ink (prior milestone) | Two caller-supplied pinned Math-font labels retain their full arrow glyphs beyond nominal font metrics. Clipping uses resolved glyph bounds at the existing PDF emission font size; known empty whitespace no longer aborts ink measurement. Text origins, programs, paths and colors stay exact. The two public-API outputs bind 126 filled glyphs with no legend cuts; one unchanged synthetic-bold title stroke per PDF remains unqualified. Bubble legend alignment survives inserting a nonpainting space. No exact new-source Office or application-view claim. |
| Side-legend ink | Four automatic left/right pie legends emitted all six names but cut 28 glyphs in their bottom labels. Their clips now include measured font ascent/descent within the chart frame; explicit authored viewports retain clipping. Twelve pie controls bind 804 glyphs with no cuts or unknown traces. Four pinned PowerPoint PDFs independently retain the complete labels. Text/key/plot origins, font programs, paths, colors and links remain unchanged. Whole-page pie comparisons remain two passes/ten failures; the four affected MAEs worsen slightly despite restored ink, and those results remain retained. |
| Picture bookmarks | Four ordinary body/cell Final/Review targets now reveal their placed pictures instead of following prose 33pt to the right. Source destinations and URI links share the existing physical picture traversal; ordered measured targets keep their path and URI decode/resource/clip guards remain. Page commands, actual font programs, image streams, native glyph positions and paths stay exact; the original 0.001pt gate remains. Portable target/gap checks cover body/cell/header/footer, wrapped/unwrapped content and nonpainting runs with/without math; the target regression fails on the parent. The prior sixteen static-picture targets remain protected. Exact combined-source Office geometry and unresolved container ownership remain unqualified. |
| Direct header/footer pictures | Ordinary mixed pictures bypassed shared wrapping even when their source could be measured. Four Final controls retained only 54.14% of the picture bounds. Dispatch now uses the shared inline path; all sixteen direct/boxed controls retain full pictures and clear source text. The screen binds 1,880 glyphs to actual programs/resources/GIDs with no paper/rectangle-clip cuts or unknown traces. Eight boxed PDFs are identical; all source/page/diagnostic identities remain. Independent static Word anchors retain their gates; new combined-source Word geometry is unqualified. |
| Fitting pictures in declared rows (prior milestone) | Fixed spacing bypassed wrapping and left a fitting picture's green center missing. Shared measurement now admits pictures contained above the declared baseline and grows minimum rows from their measured content. Twenty-four floating/inline automatic, fixed and minimum controls retain complete pictures. Eight inline outcomes keep every source text bound clear of the picture. The 28-outcome screen measures 4,046 actual-program-bound glyphs without paper/rectangle-clip cuts or unknown traces. Four oversized exact-slot fallbacks remain clipped and unqualified; authored floating `wrapNone` body-overlap residuals remain. New exact combined-source Word geometry is unqualified. |
| Nested picture/header equation | The preceding milestone restored native header math from 56.5% to 100% black-mark retention by correcting object-only origins. Its sixteen Final/Review probes retain all pictures and 55,880 actual-program-bound glyphs without paper/rectangle-clip cuts or unknown traces. Object-only frames preserve container displacement; text/math keeps independently tested Word anchors. |
| Leading blank/equation pages | The preceding milestone restored five complete footer labels (175 hidden glyphs) across three Review probes, including later note-only continuation pages. Static stories use their own measured page frame; standalone equations retain their measured body origin. Its six Final/Review outcomes contain 30,527 glyphs with no paper/rectangle-clip cuts or unknown traces. Portable regressions remain in the broader CI suite. |
| Note continuation composition | A leading blank page shifts the illustrated long-note report from four to five pages. All original text origins, sizes and bounds, picture/mask streams and link geometry remain unchanged; destinations move exactly one page. All 64 source continuation lines and both pictures survive in each mode. These are source-preserving transformation checks, not a new Word export. |
| Content-preserving transformations | Three held-out composed sources retain all native paint across six Final/Review outcomes after adding 149 paired paragraph bookmarks in body/cell/static/note stories (33,653 glyphs). A separate split of 179 equivalent text runs preserves actual-program/GID sequences, pictures, links, pages, paths/clips and diagnostics in six outcomes. Raw PDF/operation identity fails in all six split-run outcomes and remains recorded; maximum glyph-origin/bounds difference is 0.009308pt. A fresh actual-program ink screen measures 28,741 glyphs without cuts or unknowns. No new omission or pagination defect is demonstrated; no sub-point tuning or threshold change. No builds or Office actions. |
| Fitting mixed notes | Four sources/eight outcomes retain exact PDFs and all 64 continuation lines in their original order/pages. Twenty pictures retain both source colors and complete clips. Independent head/link/picture geometry keeps the original one-point gates. Native PDF/raster files are absent from the selected legacy folders, so no fresh whole-page comparison is claimed. |
| Earlier composition fixes | Table-first static frames, body-column printed-height admission, textbox picture/math neighbors and shared row/cell/box translation retain their bounded controls. Column balancing and general note-row fragmentation remain unsupported and diagnosed. |
| Static ownership repair | Broader CI exposed an unwanted ordered row for legacy images without source positions. The shared path now requires source ownership and keeps their block fallback. The previously failing ownership control and its wrapping neighbor pass; all 45 document outputs remain byte-identical, reusing their content and reference evidence. |
| Fixed-picture fallback | Broader CI caught changed placement of an oversized picture in an unmeasured exact-height row. The shared terminal-admission rule now retains that fallback; fitting declared pictures keep ordered measurement. All four fallback PDFs exactly match the preceding accepted outputs and the existing glyph/separation regression passes unchanged. Raw picture bounds alone had suggested NORTH-label overpaint, but native child clips preserve that label; no newly fixed text loss is claimed. The unfit picture remains approximate. |
| Preceding rendering checks | One exact 349-method selection per host: 347 Windows passes/two declared Aptos skips; 276 Ubuntu passes/73 predeclared font skips, with no failures or unexpected skips. Fifty-one neighbor PDFs remain identical; the sole changed SVG-footer PDF has identical glyph/path origins, pixel-identical 96-DPI output and identical fresh metrics against the pinned Word PDF. Original portfolio results remain two PPTX passes/fourteen DOCX failures; native Word pictures one/four. Twenty-four transferred assets/two smoke PDFs match. One accepted frozen artifact/shared fixture view; no root build or local full-DOCX replay. The broad parent test first fails on an inherited 0.072pt body-boundary overlap; the controlling content-loss proof is the pinned native parent screen (four clipped-word cases before, none afterward). |

The preceding autofit batch passed CI42976650266: 4,790 Windows/2,261 Ubuntu
checks, 19/129 skips, 69 .NET 10 checks per host and four package consumers.
The first leading-picture batch failed CI42986382781 on an existing page-start
assertion: the newly ordered body rows omitted before-spacing inspection metadata.
Its page-count/placement checks passed. The repair restores that metadata from the
shared spacing profile and retains the original assertions. All 36 rendering PDFs
are byte-identical; their content and independent Office evidence are reused.

The preceding leading-picture source passed CI42988198130:
**4,792 Windows/2,276 Ubuntu passes**, 19/130 skips, 69 .NET 10 checks per host
and four package consumers. Both terminal host logs were retained once. The earlier
failed run remains recorded. CI builds platform binaries separately; source coverage
does not imply binary identity across platforms.

The minimum-picture source passed CI42999234040 after a documentation-only
repair: **4,793 Windows/2,283 Ubuntu passes**, 19/130 skips, 69 .NET 10 checks
per host and four package consumers. The preceding failed CI42997335304 remains
recorded: its package preflight rejected a shipped relative CI hyperlink, which
resolved to an absent package file. The link was removed; the package check stayed
unchanged. Both terminal host logs were retained once. No local build or suite replay
was needed for the repair.

The contained fixed-line picture source passed CI43002444011:
**4,795 Windows/2,285 Ubuntu passes**, 19/130 skips, 69 .NET 10 checks per host
and all four package consumers. Both terminal host logs were retained once.
Platform binaries qualify source separately from the accepted local frozen artifact.
Status-only documentation updates reuse that qualification without local builds,
conversions or suite replay. The broader roadmap remains open.

The following-word/overhang source passed CI43008622492:
**4,796 Windows/2,286 Ubuntu passes**, 19/130 skips, 69 .NET 10 checks per host
and all four package consumers. Terminal host logs were retained once. The
broader roadmap and publication remain open.

The selected-comment-view source passed CI43015189979:
**4,798 Windows/2,287 Ubuntu passes**, 19/130 skips, 69 .NET 10
checks per host and all four package consumers. Terminal logs were retained once.
The subsequent composition checks reuse the same frozen production runtime;
documentation-only status updates preserve its qualification. The broader
roadmap and publication remain open.

Six additional page-admission probes reuse the frozen artifact: a fitting kept
cell stays on page one, an unfitting kept cell moves intact to page two, and a
body picture stays on page one while its following word resumes once on page two.
All 72 glyphs are clear of measured clips. No new-source Word geometry is claimed.

## Finite document checkpoint

The fixed portfolio contains nine sources and sixteen outcomes: both default
Final and Word-compatible Review for each DOCX, and one outcome per PPTX. Current
source, PDF, reference, mode and threshold identities have been reconciled. All
thirty DOCX Word/candidate page pairs have explicit pins. Four Final pages now
use pinned PDFium boards in place of less structured legacy thumbnail provenance;
the cached Word pages and existing candidate PDFs were reused. Shared template
ancestry limits the diversity of this cohort.

| Document | Final / Review pages | Reviewed content and pagination |
|---|---|---|
| Operations report | 2 / 2 | Heading, equations, illustration, table, items 1–3, sections 1–12, conclusion, note and footer retained. Sections 9–12 stay on page two. |
| Letter | 1 / 1 | Equations, illustration, acceptance items, sections 1–4, conclusion and note retained. |
| Specification | 2 / 2 | Sections 1–8 stay on page one; 9–20, conclusion and note on page two. |
| Held-out technical memo | 2 / 2 | Sections 1–8 and the equation table remain on page one; conclusion and note on page two. |
| Held-out static-cell report | 2 / 2 | Nested-cell/static-story equations, illustration, prose and transitions retained. |
| Held-out static-SVG report | 2 / 2 | Body illustration, header math and footer SVG retained; sections 9–12 stay on page two. |
| Held-out linked-footnote report | 4 / 4 | All 64 continuation lines retain order/pages; pictures, math and links remain. |
| Briefing and held-out decks | One page each | Pinned chart/table/text outcomes retain their two visual passes. |

This checkpoint reuses 56,513 actual-program glyph checks with zero measured cuts
or unknown traces. The separate authored-legend smoke remains identical to the
six-name/54-glyph audit. Its six origins still fail the one-point gate
(11.169/31.387pt), and MAE rose from 0.621337 to 0.713692 when content was restored.
The earlier whole-page pass had concealed missing names. Both portfolio PPTX
image gates pass; all fourteen DOCX gates still
fail, with typography, equation/table geometry and Review placement differences.
These are not blanket cosmetic classifications or threshold waivers. Two legacy
PPTX references have empty optional export logs; their source/mode/PDF identities
remain pinned. Application-view calibration remains unfinished.

Moving the existing mixed picture/equation paragraph into an ordinary table cell
in three portfolio sources adds six Final/Review controls. Native program/GID and
character sequences, picture pixels, owning pages, equations and diagnostics are
preserved; 30,327 glyphs have no measured cuts or unknowns. All eight picture frames
fit their rectangle clips. The generic screen retains eight alpha-mask abstentions;
a separate byte-level proof establishes that these unchanged Gray8 masks are fully
opaque, and all eight paired picture crops were reviewed. This qualifies these
source pictures without changing general mask handling or visual thresholds.

No new serious omission, clipping or wrong page was demonstrated in this bounded
review. The broader release gate remains open; candidate preparation is recorded
below and publication remains deferred. That evidence reconciliation reused the
qualified production runtime and closed CI evidence without a library build, full
suite, package qualification or Office action.

Candidate packaging must keep both framework assets and their portable symbols
bound to one committed build. Transferring a single-framework frozen runtime does
not refresh the other framework's root outputs. Reuse coherent qualified CI assets
or freeze both targets in an explicit worktree before packing. Verify every packed
library/PDB identity, current caller documentation and its local links. Cross-platform
package checks should consume the same selected package with each framework's
expected library hash, preserving the original qualified runtime.

## Development package payload

The current development package and symbols are prepared directly in
`artifacts/nuget/`, with checksums and a separate qualification record. The exact
same package passes Windows and Linux consumers on both frameworks. Windows CI
consumed these archive bytes; its consumer results are reused. Both Linux checks
consume that Windows package. There is no repack or library/root rebuild.

CI43062947587 / `e7d591999` supplies the .NET 10/.NET 11 libraries, portable symbols
and package. Its documentation remains the immutable CI snapshot; only later
development qualification notes differ in the current source. All caller/API
documents match, and 75 local links are checked. Status notes do not cause a repack.
The Windows artifact
digest is verified, and all 48 frozen runtime assets match their identities. Both
framework package entries match those DLL/PDB bytes. Assembly/file versions are
0.3.0.0; informational versions identify the same committed build.

The existing source-runtime portfolio/native evidence remains qualified; its
sixteen PDFs retain two PPTX passes/fourteen DOCX visual failures. Production/build
inputs match the CI package, and two CI-runtime identity smokes reproduce the
corrected deferred cell and six-name legend PDFs. No full portfolio or test-suite
replay is claimed for this metadata transfer. The preparation record keeps the
rendering and package runtime identities separate. The previous 03ef23fcc and
f852e431d packages remain parent evidence; published 0.2.0 is protected.

`TestPackageSmoke.ps1 -InspectOnly` checks all framework build identities and caller
links without packing, restoring or loading assemblies. Consumer checks cover
conversion, cancellation, admission refusal, ownership and output publication.
The selected framework still requires its qualified hash: matching versions alone
do not qualify unverified bytes. Preparation does not complete the remaining
rendering/calibration/resource roadmap or authorize publication.

## Independent content checks and limits

Ink checks bind each native paint span to its **actual embedded font program,
resource and glyph IDs**. Normalized font names can alias regular/bold subsets;
the earlier name-only 112,702-glyph portfolio claim is withdrawn. No corruption of
library font data was demonstrated. Corrected evidence checks 56,513 unchanged
portfolio glyphs and 66,612 fitting-note glyphs through verified PDF identities.

Painted glyph/picture checks supplement source-text presence, extraction and image
scores. These bounded checks do not establish arbitrary nonrectangular clipping,
image occlusion, reading order or application-view parity. Oblique/mask clips,
stroked text and unavailable/unsupported font programs remain unqualified. Native
Python validation runs on Windows; Linux library checks are separate.

A custom resolver mapping all text to pinned DejaVu Math exposes a small clipped
initial footnote A: 0.066pt in Final/0.050pt in Review. The same bearing overhang
exists in one-column controls. All six probes paint the complete source label;
the original 0.01pt ink gate still fails. This cosmetic residual is deferred.

## Remaining release gaps

- Fragmented/merged or unplaced picture-source targets retain unresolved ownership;
  exact combined-source Office destination geometry remains unqualified.
  Two Review run-split
  controls differ by 0.003479pt and fail the retained 0.001pt origin audit; picture
  and source-content checks pass. This small residual is deferred without tuning.
- The earlier claim that removing a synthetic 4.5pt/25% rotated-picture cut fixed
  a proven defect is withdrawn. The source-pinned Word inline-PNG PDF also clips
  25% of its transformed image; both anchored-PNG PDFs retain the full image.
  The exact composed-source clip remains unqualified. These supplemental Word
  PDFs use default Final and do not qualify the public AllMarkup gates or
  application view. Four original image-threshold failures remain unchanged.
- General note-row fragmentation/overflow, body/nested-table origins, floating
  table placement, custom Review canvas, numbering, annotations and typography
  remain open. Unfit fixed-line picture slots keep their clipped fallback.
- An unspaced picture/word boundary with pinned Sans flags 0.072pt of glyph-box/
  image overlap at the unchanged 0.01pt screen. Raster occlusion is unmeasured;
  the spaced portable source passes. Exact combined-source Word geometry and
  arbitrary paint occlusion remain unqualified. Earlier descent/paint trials
  failed protected gates and are not consumed.
- Words following captured ordinary fixed-picture overhangs now wrap without loss.
  Terminal, reconstructed and prepared overhang frames keep their approximate
  fallbacks. Intentional picture clipping and exact new-source Word geometry
  remain unresolved; this does not establish arbitrary full-picture visibility.
- The native SVG fixed-footer PDF changes only its rotation clip from 30pt
  (including paragraph after-spacing) to its measured 22pt slot. Pixels and fresh
  Word-PDF metrics are identical; the original structural-similarity/recall
  failure remains. No colored SVG pixels appear in either candidate or the
  pinned Word PDF. Application view is unavailable, so no stronger source-picture
  completeness or newly demonstrated library-omission conclusion is claimed.
- An unfit 300-label note row retains 279 complete labels, six partial and fifteen
  off-paper labels. The pinned Office PDFs also retain 279 complete/six partial,
  but omit the other fifteen paint operations. Both fitting 240-label controls
  remain complete. General fragmentation stays unresolved and warned.
- PPTX manual/side furniture, automatic ranges, dense labels and font/plot/palette
  geometry remain approximate. No new count-specific or document-shape patch is
  claimed. Original visual thresholds and failures remain unchanged.
- **Office application-view calibration is unfinished.** Office PDFs remain
  bootstrap references. Unavailable or rejected captures block dependent
  conclusions only; no new Office export was needed for these fixes.
- Coarse host-memory coverage remains partial. Safe managed execution, cancellation,
  input bounds, bounded caches/staging, deterministic ownership/refunds, atomic
  publication and PDF validity remain; no comprehensive heap ceiling is claimed.

The broader roadmap remains open; publication/tagging stay deferred. See
[capabilities](Capabilities.md), [diagnostics](Diagnostics.md) and
[caller migration](PublicApiChanges.md).
