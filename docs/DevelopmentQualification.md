# 0.3.0 development qualification

Updated 2026-11-11. **0.3.0 is not published and is not a completed release
candidate.** The published [0.2.0 qualification](PreviewQualification.md) remains
separate. Development targets .NET 10 and .NET 11.

This snapshot pins production `7343c768bd197919333c7adbc0a192167d9312ef` and
harness `35c9e434f657e2bd42e29fc5a48eb7b9629d0ad9`. The frozen .NET 11 library
SHA-256 is `F885FD51FA851283AB57587BF6D9454DA232802654B9353211385A6A24D53E47`.
Documentation and harness updates do not rebuild that library.

## Verified outcomes

| Gate | Result and scope |
|---|---|
| Direct header/footer pictures | Ordinary mixed pictures bypassed shared wrapping even when their source could be measured. Four Final controls retained only 54.14% of the picture bounds. Dispatch now uses the shared inline path; all sixteen direct/boxed controls retain full pictures and clear source text. The screen binds 1,880 glyphs to actual programs/resources/GIDs with no paper/rectangle-clip cuts or unknown traces. Eight boxed PDFs are identical; all source/page/diagnostic identities remain. Independent static Word anchors retain their gates; new combined-source Word geometry is unqualified. |
| Fitting pictures in declared rows (prior milestone) | Fixed spacing bypassed wrapping and left a fitting picture's green center missing. Shared measurement now admits pictures contained above the declared baseline and grows minimum rows from their measured content. Twenty-four floating/inline automatic, fixed and minimum controls retain complete pictures. Eight inline outcomes keep every source text bound clear of the picture. The 28-outcome screen measures 4,046 actual-program-bound glyphs without paper/rectangle-clip cuts or unknown traces. Four oversized exact-slot fallbacks remain clipped and unqualified; authored floating `wrapNone` body-overlap residuals remain. New exact combined-source Word geometry is unqualified. |
| Nested picture/header equation | The preceding milestone restored native header math from 56.5% to 100% black-mark retention by correcting object-only origins. Its sixteen Final/Review probes retain all pictures and 55,880 actual-program-bound glyphs without paper/rectangle-clip cuts or unknown traces. Object-only frames preserve container displacement; text/math keeps independently tested Word anchors. |
| Leading blank/equation pages | The preceding milestone restored five complete footer labels (175 hidden glyphs) across three Review probes, including later note-only continuation pages. Static stories use their own measured page frame; standalone equations retain their measured body origin. Its six Final/Review outcomes contain 30,527 glyphs with no paper/rectangle-clip cuts or unknown traces. Portable regressions remain in the broader CI suite. |
| Note continuation composition | A leading blank page shifts the illustrated long-note report from four to five pages. All original text origins, sizes and bounds, picture/mask streams and link geometry remain unchanged; destinations move exactly one page. All 64 source continuation lines and both pictures survive in each mode. These are source-preserving transformation checks, not a new Word export. |
| Fixed composed portfolio | Nine sources/sixteen outcomes retain their PDFs. Existing independent comparisons remain two PPTX image passes and fourteen DOCX image failures. The 64-page readability review is reused through PDF identity. Typography/geometry remain approximate; shared report ancestry limits document diversity. |
| Fitting mixed notes | Four sources/eight outcomes retain exact PDFs and all 64 continuation lines in their original order/pages. Twenty pictures retain both source colors and complete clips. Independent head/link/picture geometry keeps the original one-point gates. Native PDF/raster files are absent from the selected legacy folders, so no fresh whole-page comparison is claimed. |
| Authored legend | All six names/54 label glyphs remain visible; one focused conversion is byte-identical to the accepted correction. The image gate previously concealed missing names. Native origins, plot geometry and typography remain approximate. |
| Earlier composition fixes | Table-first static frames, body-column printed-height admission, textbox picture/math neighbors and shared row/cell/box translation retain their bounded controls. Column balancing and general note-row fragmentation remain unsupported and diagnosed. |
| Focused checks | 103 exact checks pass on Windows; 79 pass on Ubuntu with 24 reviewed installed-font skips. The new static-picture regression fails against the qualified parent; two neighboring controls pass. One production build supplies all outputs; test-only repairs preserve its assets. A legacy route assertion was replaced by declared-slot and complete-picture invariants. No root build or local full-DOCX replay. |

The preceding broader checkpoint `48bf80f8d` passed CI42924107847: 4,775 Windows/
2,242 Ubuntu checks, 19/129 reviewed skips, 69 .NET 10 checks per host and four
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

- Body links to static picture bookmarks have no destinations in the sixteen
  combination probes. Exact Office behavior is unqualified. Two Review run-split
  controls differ by 0.003479pt and fail the retained 0.001pt origin audit; picture
  and source-content checks pass. This small residual is deferred without tuning.
- General note-row fragmentation/overflow, body/nested-table origins, floating
  table placement, custom Review canvas, numbering, annotations and typography
  remain open. Unfit fixed-line picture slots keep their clipped fallback.
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
