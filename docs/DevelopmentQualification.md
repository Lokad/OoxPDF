# 0.3.0 development qualification

Updated 2026-11-10. **0.3.0 is not published and is not a completed release
candidate.** The shipped [0.2.0 qualification](PreviewQualification.md) remains a
separate historical record. Development targets .NET 10 and .NET 11.

The rendering checkpoint below is source
`4131a48ae0c0468501e2924d87bb117c139029d5`. Its frozen .NET 11 library SHA-256 is
`38DC9BF837E2226D064C7F31A6F8D1CE58AD1947DCD2BE3513C446CC13F31ED3`.
Documentation updates do not rebuild that runtime.

## Verified outcomes

| Gate | Result and scope |
|---|---|
| Supported textbox neighbors | A fitting caption/picture retains shared wrapping beside an omitted unsupported equation. Final and Review recover picture bounds from 23.63% to 100%, including the green center. The equation warning remains; no unused math resource is published. Prepared-math frames, fixed-line picture admission and entirely unsupported paragraph fallbacks are preserved. Exact Office geometry for the new combined source remains unqualified. |
| Ordered picture composition | Ordinary automatic-flow textboxes retain pictures and caption ink across left, center, right and justified alignment and equivalent run splits. Shared floating row/cell/box translation and complete image projection preserve nested content. These bounded controls do not establish unrestricted floating layout. |
| Authored horizontal legend | All six names paint instead of two. Independent glyph-ink/clip checks supplement an image gate that previously concealed the omission. Origins still differ from the pinned Office PDF by about 11.17pt horizontally and 31.39pt vertically. |
| Oversized endnote table | Final/Review each recover 30 complete labels: 249 → 279, plus 72pt of a tall picture. Six labels remain partial and fifteen extend beyond paper, matching the native viewport. General row fragmentation remains unsupported and diagnosed. |
| Fixed composed portfolio | Nine sources/sixteen outcomes retain their PDF identities. Existing numerical results remain two PPTX passes and fourteen DOCX failures. Independent per-page character/picture checks and a 64-page-image readability review are reused through PDF identity; typography/geometry remain approximate. Shared report ancestry limits document diversity. |
| Focused checks and transfer | Twenty-six exact checks pass on Windows; twenty-five pass on Ubuntu with one declared Windows-font anchor skip. The picture regression fails on the parent; neighboring controls pass. 44/46 PDFs and all source/page/diagnostic identities match the parent cohort, including all fixed-line and separate-paragraph controls. One accepted frozen runtime supplies the checks, document outputs and byte-identical root transfer with two integration smokes; no root build or local full DOCX replay. |
| Prior broader gate | CI42856401794 records the preceding `207ecc530` rendering checkpoint and CI-only `61b64c900` repair: 4,759 Windows/2,225 Ubuntu passes, 19/129 reviewed skips, 69 .NET 10 checks per host and four package consumers. This is a historical broader gate, separate from the focused checkpoint above. CI builds platform binaries separately; no cross-platform binary-identity claim. |

Independent emitted-GID ink, effective rectangle clips and opaque-overpaint checks
supplement extraction and image scores. Current controlled labels remain complete;
the fixed portfolio's earlier 112,702-glyph screen is reused through PDF identity.
These bounded checks do not prove arbitrary nonrectangular clipping, image
occlusion, reading order or application-view parity. Original thresholds and
failures remain unchanged.

## Other retained evidence

At source `207ecc530`, four fitting mixed-note sources/eight outcomes preserve
all 64 continuation lines in order/on their independent Word pages. Twenty pictures
fit clips and show both source colors; note heads, targets and image frames retain
the original one-point numeric gates. A 66,612-glyph screen finds no clipping or
opaque-rectangle/controlled-picture overlap. Native PDF/raster files are unavailable
in the selected legacy folders; no fresh whole-page comparison is claimed.

Twelve earlier source-matched note/image comparisons retain their whole-image
failures. Eight neighboring reference input hashes are stale and four Review
picture outcomes lack references in those caches; they support no native fidelity
conclusion. No new Office export was needed.

At source `33609a92c`, the formerly failing
`DocxWordCompatibleBalloonBodyWrapsLikeOfficeTitles` and all eight named Ubuntu
inspection/chart-link failures in the
[historical record](PreviewQualification.md#inherited-failures-and-limitations)
actually passed with their preconditions satisfied. Their earlier failure reports
remain valid for their original runtime/environment; this is not a claim that the
current focused selection replays those nine checks.

## Remaining release gaps

- Supported content completeness is independent of whole-page image scores.
  General note-row fragmentation/footnote overflow, body/nested-table origins,
  custom Review print canvas, numbering, annotations and typography remain open.
  Unfit fixed-line picture slots retain their clipped fallback; the new neighbor
  correction qualifies ordinary automatic flow.
- PPTX manual/side furniture, automatic ranges, dense plot allocation, font geometry
  and palette/label rules remain approximate. No unvalidated count-specific or
  document-shape-specific replacement is claimed.
- **Office application-view calibration is unfinished.** Office-PDF comparisons
  remain bootstrap references. Unavailable or rejected captures block dependent
  conclusions only.
- Coarse host-memory coverage remains partial. Existing cancellation, bounded
  caches/staging, deterministic ownership/refunds, safe managed execution,
  publication guarantees and PDF validity remain; no comprehensive heap ceiling.

The broader rendering, resource and qualification roadmap remains open.
Publication/tagging stay deferred. See [capabilities](Capabilities.md),
[diagnostics](Diagnostics.md) and [caller migration](PublicApiChanges.md).
