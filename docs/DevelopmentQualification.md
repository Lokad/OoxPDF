# 0.3.0 development qualification

Updated 2026-11-11. **0.3.0 is not published and is not a completed release
candidate.** The published [0.2.0 qualification](PreviewQualification.md) remains
separate. Development targets .NET 10 and .NET 11.

This snapshot pins production
`f77a2f9191d235042c1bf90c4db0ce64b12af0b3` and the same harness revision.
The frozen .NET 11 library SHA-256 is
`37FAA83147392B1AB79D239EEBAFCE99DF66A259F94CBA6F20B6F01E7EB0C8F0`.
Documentation and harness updates do not rebuild that library.

## Verified outcomes

| Gate | Result and scope |
|---|---|
| Autofit paragraph insets | A 72pt first-line indent hid every glyph of a short label while text operations remained present. Natural cell widths now fund the same paragraph offsets used by placement, including positive right insets and measured inline content. Both missing labels recover the controls' 270 raster ink pixels; four control PDFs are unchanged and 30 actual-program glyphs have no measured cuts/unknown traces. Portable checks cover two fonts/presets, split runs, grid hints, left/first/right insets and cancellation; the regression fails on the parent. Two previously absent pictures also regain complete bounds/colors, but two extreme first-line picture controls still clip the following word. No all-content, compressed-width or exact new-source Word/view claim. |
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
| Fixed composed portfolio | Nine sources/sixteen outcomes retain their PDFs. Existing independent comparisons remain two PPTX image passes and fourteen DOCX image failures. The 64-page readability review is reused through PDF identity. Typography/geometry remain approximate; shared report ancestry limits document diversity. |
| Fitting mixed notes | Four sources/eight outcomes retain exact PDFs and all 64 continuation lines in their original order/pages. Twenty pictures retain both source colors and complete clips. Independent head/link/picture geometry keeps the original one-point gates. Native PDF/raster files are absent from the selected legacy folders, so no fresh whole-page comparison is claimed. |
| Authored legend | All six names/54 label glyphs remain visible; one focused conversion is byte-identical to the accepted correction. The image gate previously concealed missing names. Native origins, plot geometry and typography remain approximate. |
| Earlier composition fixes | Table-first static frames, body-column printed-height admission, textbox picture/math neighbors and shared row/cell/box translation retain their bounded controls. Column balancing and general note-row fragmentation remain unsupported and diagnosed. |
| Static ownership repair | Broader CI exposed an unwanted ordered row for legacy images without source positions. The shared path now requires source ownership and keeps their block fallback. The previously failing ownership control and its wrapping neighbor pass; all 45 document outputs remain byte-identical, reusing their content and reference evidence. |
| Fixed-picture fallback | Broader CI caught changed placement of an oversized picture in an unmeasured exact-height row. The shared terminal-admission rule now retains that fallback; fitting declared pictures keep ordered measurement. All four fallback PDFs exactly match the preceding accepted outputs and the existing glyph/separation regression passes unchanged. Raw picture bounds alone had suggested NORTH-label overpaint, but native child clips preserve that label; no newly fixed text loss is claimed. The unfit picture remains approximate. |
| Focused checks | 199 Windows passes/two Aptos skips; Ubuntu 168 passes/33 reviewed font-dependent skips in the same exact 201-method set. Four independent Word checks pass on Windows and require fonts unavailable on Ubuntu. The initial strict Windows gate's undeclared skips and Ubuntu's four additional font preconditions remain recorded; only the two Windows prerequisites were reselected, with no full replay. Thirty neighbor PDFs remain exact: 17 ordinary outputs including the portfolio/legend, five standalone pictures and eight static outcomes. Twenty-four transferred assets/two smoke PDFs match. One frozen production build/shared fixture view; no root build/local full-DOCX replay. |

The preceding picture-navigation correction at `b15e10859` passed CI42962826027: 4,787 Windows/
2,256 Ubuntu checks, 19/129 reviewed skips, 69 .NET 10 checks per host and four
package consumers. The earlier rotation correction also passed its broader gate
after repairing the note-ownership assertion exposed by CI42949586292.
CI42968585479 then failed the fixed-picture fallback neighbor on both hosts
(4,788/2,257 passes, 19/129 skips); terminal logs remain retained. The repaired
batch at `62391bc43` passed CI42970799951: **4,789 Windows/2,258 Ubuntu passes**,
19/129 reviewed skips, 69 .NET 10 checks per host and four package consumers.
Both terminal host logs were fetched once. The current autofit-inset batch at
`b33b55c3e` passed CI42976650266: **4,790 Windows/2,261 Ubuntu passes**,
19/129 reviewed skips, 69 .NET 10 checks per host and four package consumers.
This status-only update reuses that unchanged-source result. Earlier
CI42929150677 failed the legacy static-image ownership control on both hosts;
those reports remain retained, and the repair passed broader CI. Consult the
reports for the revision being consumed. CI builds platform binaries separately;
source coverage does not imply binary identity across platforms.

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
- Two extreme first-line picture controls retain complete pictures after the
  autofit repair, but still clip four glyphs of the following word. This supported
  composition remains wrong; exact combined-source Word geometry is unqualified.
  A broader picture-layout trial restored the words but failed an independent
  0.75pt following-paragraph gate and was rejected; its runtime is not consumed.
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
