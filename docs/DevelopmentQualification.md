# 0.3.0 development qualification

Updated 2026-11-11. **0.3.0 is not published and is not a completed release
candidate.** The published [0.2.0 qualification](PreviewQualification.md) remains
separate. Development targets .NET 10 and .NET 11.

This snapshot pins production and harness
`7d56d6dc81dbb88fe3c2b46079adb4223e1ca04e`. The frozen .NET 11 library
SHA-256 is `4A3A8CBC569F65297853A66886A4C512CCB7618E88AA44A3B63B4EBEBCB107B8`.
Documentation and harness updates do not rebuild that library.

## Verified outcomes

| Gate | Result and scope |
|---|---|
| Nested picture/header equation | A picture-only nested cell overpainted the header equation while extraction passed. The failing native raster now retains 100% of the control's black math marks, up from 56.5%. Object-only frames use their occupied top and preserve container displacement in existing source metadata. Sixteen Final/Review probes retain all pictures and 55,880 actual-program-bound glyphs, with no paper/rectangle-clip cuts or unknown traces. Mixed text/math Word anchors remain unchanged. New exact combined-source Office geometry is unqualified. |
| Leading blank/equation pages | The preceding milestone restored five complete footer labels (175 hidden glyphs) across three Review probes, including later note-only continuation pages. Static stories use their own measured page frame; standalone equations retain their measured body origin. Its six Final/Review outcomes contain 30,527 glyphs with no paper/rectangle-clip cuts or unknown traces. Portable regressions remain in the current selection. |
| Note continuation composition | A leading blank page shifts the illustrated long-note report from four to five pages. All original text origins, sizes and bounds, picture/mask streams and link geometry remain unchanged; destinations move exactly one page. All 64 source continuation lines and both pictures survive in each mode. These are source-preserving transformation checks, not a new Word export. |
| Fixed composed portfolio | Nine sources/sixteen outcomes retain their PDFs. Existing independent comparisons remain two PPTX image passes and fourteen DOCX image failures. The 64-page readability review is reused through PDF identity. Typography/geometry remain approximate; shared report ancestry limits document diversity. |
| Fitting mixed notes | Four sources/eight outcomes retain exact PDFs and all 64 continuation lines in their original order/pages. Twenty pictures retain both source colors and complete clips. Independent head/link/picture geometry keeps the original one-point gates. Native PDF/raster files are absent from the selected legacy folders, so no fresh whole-page comparison is claimed. |
| Authored legend | All six names/54 label glyphs remain visible; one focused conversion is byte-identical to the accepted correction. The image gate previously concealed missing names. Native origins, plot geometry and typography remain approximate. |
| Earlier composition fixes | Table-first static frames, body-column printed-height admission, textbox picture/math neighbors and shared row/cell/box translation retain their bounded controls. Column balancing and general note-row fragmentation remain unsupported and diagnosed. |
| Focused checks | 84 exact checks pass on Windows; 58 pass on Ubuntu with 26 reviewed installed-font skips. The new cross-container regression fails against the qualified parent; two neighboring controls pass. One accepted frozen runtime supplies all qualification outputs. Failed development trials and runner-setup attempts remain recorded separately. No root build or local full-DOCX replay. |

The preceding broader checkpoint `5a855a04e` passed CI42913236822: 4,773 Windows/
2,240 Ubuntu checks, 19/129 reviewed skips, 69 .NET 10 checks per host and four
package consumers. That historical gate is separate from the focused snapshot
above. Consult the CI reports for the revision being consumed. CI builds
platform binaries separately;
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

- General note-row fragmentation/overflow, body/nested-table origins, floating
  table placement, custom Review canvas, numbering, annotations and typography
  remain open. Unfit fixed-line picture slots keep their clipped fallback.
- Some native-matching unfragmented note tails still contain six partial and
  fifteen off-paper labels. Fitting controls do not qualify those overflow cases.
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
