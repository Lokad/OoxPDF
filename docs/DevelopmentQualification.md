# 0.3.0 development qualification

Updated 2026-11-10. **0.3.0 is not published and is not a completed release
candidate.** The published [0.2.0 qualification](PreviewQualification.md) remains
separate. Development targets .NET 10 and .NET 11.

This snapshot pins production `a68b0e702eaa41dce90a71f067b02b1e352d8380` and
harness `bb1375348529e7b382ef1f62cb1cbc1908ca3906`. The frozen .NET 11 library
SHA-256 is `E884050BE14772A7B9BE585312E6D94721D68E8EB32B3539CB850027E21ACB76`.
Documentation and harness updates do not rebuild that library.

## Verified outcomes

| Gate | Result and scope |
|---|---|
| Leading blank/equation pages | Restored five complete footer labels (175 hidden glyphs) across three Review probes, including later note-only continuation pages. Static stories use their own measured page frame; standalone equations retain their measured body origin. Six Final/Review outcomes contain 30,527 glyphs with no paper/rectangle-clip cuts or unknown traces. Body/header equation bounds remain separate. Exact combined-source Office geometry is unqualified. |
| Note continuation composition | A leading blank page shifts the illustrated long-note report from four to five pages. All original text origins, sizes and bounds, picture/mask streams and link geometry remain unchanged; destinations move exactly one page. All 64 source continuation lines and both pictures survive in each mode. These are source-preserving transformation checks, not a new Word export. |
| Fixed composed portfolio | Nine sources/sixteen outcomes retain their PDFs. Existing independent comparisons remain two PPTX image passes and fourteen DOCX image failures. The 64-page readability review is reused through PDF identity. Typography/geometry remain approximate; shared report ancestry limits document diversity. |
| Fitting mixed notes | Four sources/eight outcomes retain exact PDFs and all 64 continuation lines in their original order/pages. Twenty pictures retain both source colors and complete clips. Independent head/link/picture geometry keeps the original one-point gates. Native PDF/raster files are absent from the selected legacy folders, so no fresh whole-page comparison is claimed. |
| Authored legend | All six names/54 label glyphs remain visible; one focused conversion is byte-identical to the accepted correction. The image gate previously concealed missing names. Native origins, plot geometry and typography remain approximate. |
| Earlier composition fixes | Table-first static frames, body-column printed-height admission, textbox picture/math neighbors and shared row/cell/box translation retain their bounded controls. Column balancing and general note-row fragmentation remain unsupported and diagnosed. |
| Focused checks | 61 exact checks pass on Windows; 41 pass on Ubuntu with 20 declared installed-font skips. Both new portable regressions fail against the parent. One production build and one harness extension supply all checks and outputs; unchanged checks are joined by verified identity. No root build or local full-DOCX replay. |

The preceding broader checkpoint `00cc795af` passed CI42903561738: 4,771 Windows/
2,238 Ubuntu checks, 19/129 reviewed skips, 69 .NET 10 checks per host and four
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
