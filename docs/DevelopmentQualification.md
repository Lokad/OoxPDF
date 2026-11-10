# 0.3.0 development qualification

Updated 2026-11-10. **0.3.0 is not published and is not a completed release
candidate.** The shipped [0.2.0 qualification](PreviewQualification.md) remains a
separate historical record. Development targets .NET 10 and .NET 11.

The current rendering checkpoint is source
`33609a92c80107272b5d7a4ab7075af242ddd4e5`. Its frozen .NET 11 library SHA-256 is
`237BF9E0B0E8FDE93C5DB005F464D25933B202474B113C084752BDB8B62B7D01`.
Documentation updates do not rebuild that runtime.

## Verified outcomes

| Gate | Result and scope |
|---|---|
| Authored horizontal legend | All six names paint instead of two. Independent glyph-ink and clip checks supplement the passing image gate, which previously concealed the omission. Absolute origins still differ from the pinned Office PDF by about 11.17 points horizontally and 31.39 vertically. |
| Oversized endnote table | Final and Review each recover 30 complete labels: 249 to 279. Six labels remain partial and fifteen extend beyond paper, matching the native viewport behavior. An additional 72 points of a tall picture now paint. Source text positions, cell clips and pagination stay unchanged. General row fragmentation remains unsupported and diagnosed. |
| Fixed composed portfolio | Nine public/synthetic sources, sixteen document/mode outcomes retain their PDF and diagnostic identities. Reports, letters, specifications and decks combine tables, numbering, links, equations, pictures, headers and notes. Their existing image results remain two PPTX passes and fourteen DOCX failures. |
| Independent content screens | Fourteen DOCX outputs and their native PDFs have no text ink outside the physical paper or their explicit rectangle clip stacks in the measured 112,702 glyphs. A later-opaque-rectangle screen finds no potential text overpaint in those PDFs. These bounded screens do not prove arbitrary image, nonrectangular clipping, reading-order or overpaint correctness. |
| Focused checks and transfer | Nineteen exact checks pass on Windows and Ubuntu with no skips. The new regression fails against the parent library; two neighboring controls pass. Twenty-four runtime assets transfer byte-identically and two integration PDFs match. No root build or full local DOCX replay. |
| Broader qualification | [CI 42814980035](https://github.com/Lokad/OoxPDF/actions/runs/42814980035) passes on Windows and Ubuntu: 4,754 and 2,208 console passes, with 19 and 128 declared skips. Each host passes 69 .NET 10 compatibility checks and two package consumers. CI uses separate platform builds of the same source; these are not claims of cross-platform binary identity. |

Twelve source-matched note/image comparisons retain their original whole-image
failures. Nine results are unchanged; the long-row error improves in both modes,
and the tall picture improves. Eight neighboring reference input hashes are stale;
four Review picture outcomes have no reference in the selected caches. They
support no native fidelity conclusion. No new Office export was needed.

## Historical failures rechecked

On the same frozen runtime, the formerly failing
`DocxWordCompatibleBalloonBodyWrapsLikeOfficeTitles` passes on Windows with its
font preconditions satisfied. All eight named Ubuntu inspection/chart-link
failures in the [historical record](PreviewQualification.md#inherited-failures-and-limitations)
also pass in one exact-method invocation. These are nine actual passes, with no
skip substitution, source change or rebuild during this recheck. Historical
failure reports remain valid for their earlier runtime and environment.

## Remaining release gaps

- Visible completeness is independent of whole-page image scores. Keep original
  thresholds and raw failures; a successful conversion does not establish fidelity.
- Note-row fragmentation and the beyond-paper tails above remain unsupported.
  Oversized native witnesses are endnotes; fitting footnote controls are preserved,
  but general footnote overflow fidelity is unqualified.
- DOCX body/nested-table origins, custom Review print canvas, numbering, annotations
  and other typography/layout differences remain. The fixed portfolio's page
  counts and content screens do not qualify every supported combination.
- PPTX manual/side furniture, automatic ranges, dense plot allocation, font geometry
  and composed layouts remain approximate. Four retained dense controls preserve
  native category counts and ellipsis counts; this does not resolve their plot or
  caption geometry. Some remaining chart-color and document-shape rules lack an
  independently validated general replacement.
- **Office application-view calibration is unfinished.** Office-PDF comparisons
  are bootstrap references, not calibrated visible-view truth. Unavailable or
  rejected captures block only dependent conclusions.
- Coarse host-memory coverage remains partial. Qualified owners, cancellation,
  bounded caches/staging, deterministic refunds, safe managed execution, publication
  guarantees and PDF validity are preserved; no comprehensive heap ceiling is promised.

The broader rendering, resource and qualification roadmap remains open. Publication
and tagging stay deferred. See [capabilities](Capabilities.md),
[diagnostics](Diagnostics.md) and [caller migration](PublicApiChanges.md) before
relying on a conversion.
