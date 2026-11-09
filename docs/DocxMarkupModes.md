# DOCX Markup Modes

DOCX conversion supports four review views through `OoxPdfOptions.DocxMarkupMode` and the CLI `--docx-markup` option. The default is `Final`, preserving the previous no-markup behavior.

## Modes

- `Final`: renders the final document text. Insertions and moved-to content are visible; deletions and moved-from content are hidden. Comment bodies, comment markers, change bars, and balloons are not printed.
- `Original`: renders the original document text. Deletions and moved-from content are visible; insertions and moved-to content are hidden. Markup indicators are not printed.
- `SimpleMarkup`: renders final document text, plus lightweight review indicators. Deletions and moved-from content remain hidden, changed paragraphs and table rows get page-margin change bars, comment anchors get compact body markers, and comment/revision balloons are suppressed.
- `AllMarkup`: renders insertions, deletions, moved-from content, and moved-to content with inline revision styling. Comment anchors get body markers, resolved comment stories are rendered as first-pass margin balloons with author/date metadata when available, resolved/open state, compact reply summaries, preview text, `[table]` and `[image]` fallbacks, and connectors. Changed paragraphs and table rows get first-pass tracked-change balloons with reviewer/date titles, short deleted-text previews when safe, nearby-change grouping, and connector offsets for same-anchor markup; formatting revisions include private-safe property-family labels such as formatted run, paragraph, table, row, cell, or section.

## Geometry

By default, markup modes keep the authored DOCX page media box and main text-column geometry. This compatibility path is `OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout` and is the default for both API and CLI conversion.

`OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin` is an opt-in all-markup geometry mode. It keeps the PDF media box deterministic, but increases the effective outside review margin before layout so the body text frame is narrower and comment/revision balloons have a Word-like review margin. The default reserve is on the right; documents with `w:mirrorMargins` mirror the body margins on even pages but keep the review lane on the right on all pages (Office balloons right on both pages). Page gutters from `w:pgMar/@w:gutter` are treated as authored inside-margin space before the outside review margin is reserved, so mirrored even pages apply the gutter to the right while keeping the right review lane. The mode is currently implemented as a deterministic body-frame shrink, not as page expansion. It has no effect unless `DocxMarkupMode` is `AllMarkup`.

`OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup` is the explicit Office-compatible all-markup profile. It is also opt-in and has no effect outside `AllMarkup`. It uses a reserved review margin with print scaling, text placement and balloon sizing compared against Word PDF exports; application-view calibration remains open. Mirrored pages keep equal scaled wrap widths and a fixed right balloon lane. Printed comment threads show the parent text without reply summaries, reply rows or separator strokes, matching Word 16 exports for resolved and open threads. Reply ownership and state remain available in the model and inspection; the other geometry profiles retain their compact reply summaries. Comment balloons and range markers print from the main body and its table cells; header/footer/note/textbox comments are omitted. Hidden-story-only comments retain nominal print scale and omit the gray lane. Surviving formatting balloons retain their lane. When no balloons survive, the authored body frame and nominal text placement are retained; unused comment parts and inline revisions do not reserve an empty margin. Document settings that hide comments also suppress their fitted text shift. Visible-balloon typography and positioning remain approximate. The profile remains approximate: see [DOCX markup fidelity triage](DocxMarkupFidelityTriage.md) for measured residuals.

Measured ordinary-body comment fills follow the authored or natural line box;
glyph brackets retain the line's natural font height. Increasing exact line
spacing can enlarge the fill without stretching the brackets. Balloons use
measured font boxes where available; object and unmeasured cases remain approximate.

Print fitting shrinks pages when the body and review lane need more room;
pages that already fit retain 100% scale. The gray lane follows the placed
balloons at unit scale as well as on reduced pages.

Ordinary minimum-spaced body paragraphs retain their measured line baselines when
they wrap, align horizontally or continue onto another page, in Final and unscaled Word-compatible
Review. Their backgrounds leave only the first source line's extra leading unshaded.
Explicit before/after point spacing preserves these baselines. Connected shaded
paragraphs also fill their shared gap. Note-numbering, kerning and reduced-print
alignment and spacing retain their separate limits.

Formatting captions use the available regular review UI face, while comment
paragraphs retain their authored faces. Fitting footer formatting balloons share
the occupied footer frame and printed body boundary; unused right-margin space
remains available to them. Unmeasured/font-fallback cases and custom page-height
Review placement remain approximate. A retained custom-height probe has a large
page-wide vertical shift, while its checked Final origins agree with Word.

An explicit point anchor before paragraph content prints no comment balloon in
Word-compatible geometry and reserves no review lane. Points after content keep
their balloons. Comment metadata remains available for inspection; incomplete
source ranges, fields and note-marker anchors retain their approximate fallback.

Independent Word 16 references for this profile use mixed `RevisionsMode=2`,
with comments, formatting changes and insertions/deletions explicitly enabled.
Inline mode 1 suppresses formatting balloons and is equivalent only where the
input has no surviving formatting changes. Export tools verify the observed view.
Formatting balloons anchored in direct header/footer paragraphs fit the page
like body balloons. Supported source-owned run revisions retain shared static
frame mapping, so adjacent equations stay visible. Other story/effect combinations
retain their previous scale behavior.

Dense overflow is still represented by compact margin continuation boxes when the reserved or existing margin cannot fit every balloon on the page.

## API

```csharp
OoxPdfConverter.Convert(
    "input.docx",
    "output.pdf",
    new OoxPdfOptions
    {
        DocxMarkupMode = OoxPdfDocxMarkupMode.AllMarkup,
        DocxMarkupGeometryMode = OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
    });
```

## CLI

```powershell
dotnet src/Lokad.OoxPdf.Cli/bin/Debug/net11.0/Lokad.OoxPdf.Cli.dll convert input.docx output.pdf --docx-markup all --docx-markup-geometry word-compatible
```

Accepted markup CLI values are `final`, `original`, `simple`, `simple-markup`, `all`, and `all-markup`. Accepted geometry CLI values are `preserve`, `preserve-layout`, `preserve-document-layout`, `reserve`, `reserve-margin`, `markup-margin`, `reserve-markup-margin`, `word`, `word-compatible`, `word-compatible-all-markup`, `office`, `office-compatible`, and `office-compatible-all-markup`. Both DOCX markup options are rejected for PPTX input.

## Diagnostics

Explicit markup modes emit approximation diagnostics instead of broad unsupported markup diagnostics when the corresponding feature has first-pass support:

- `DOCX_APPROXIMATED_COMMENTS`
- `DOCX_APPROXIMATED_FORMATTING_REVISIONS`
- `DOCX_APPROXIMATED_TRACKED_CHANGES`

`Final` remains the compatibility default. If markup exists but is not requested for printing, unsupported diagnostics can still be emitted to report ignored review content.

`markup-summary.json` includes private-safe comment anchor accounting for DOCX inspection runs. Comment bodies are counted as visibly anchored, hidden by the selected revision-filtering mode, orphaned when no package anchor is present, or unsupported when the comment story has no usable id. It also reports the expected visible comment-story anchors for `AllMarkup`, how many rendered comment balloon candidates represent them, and how many visible story anchors did not produce a rendered balloon candidate. Its geometry summary carries the document mirror-margin setting, gutter tokens, reserve-margin measurements, and per-page body-frame/markup-lane rectangles. Its balloon counters distinguish raw markup signals from rendered placements and grouped candidates: `Raw*SignalCount` fields count source/layout review signals, `Rendered*PlacementCount` fields count emitted balloon rectangles, and `Rendered*CandidateCount` fields count the comment/revision candidates represented by those rectangles. Balloon side buckets, page-local markup side summaries, connector side-consistency counts, and clamped connector-anchor counts make left/right lane selection, mirrored-lane stem regressions, and page-edge anchor adjustments visible without inspecting individual balloon rectangles. Cached-reference gates include `connector-side-inconsistent-count` as a zero-tolerance markup-geometry gate.


Anchor totals describe model visibility. A visible source anchor without a
rendered balloon can reflect profile suppression, including static-story comments
and points before paragraph content; the raw counter alone does not establish
lost comment metadata or a rendering defect.

## Note navigation

Explicit automatic marks inside referenced footnote/endnote stories follow the visible main-story numbering and retain their authored run styles. Visible automatic markers in body paragraphs and table cells create internal PDF links to the first placed matching note, including destinations on a later page. Missing or unmeasurable source/target content creates no link. Plain document endnotes that fit above bottom footnotes can remain on the same page at unit print scale; overflow, scaling and complex content retain continuation behavior.

Supported explicit note starts/formats from a single section apply to visible body numbering while the authored settings snapshot remains unchanged. Plain document endnotes at nominal scale follow the last body paragraph after-spacing; a blank separator does not add its own latent padding. Multiple sections, page/section restarts, document-setting precedence, custom markers and back-links remain incomplete. Complex/scaled note flow retains its earlier placement. Some link rectangles still differ from Office; complete parity remains unqualified.

Word-compatible revision bars use the authored left margin and the printed-page
scale. They remain centered in a narrow margin, with the distance from body text
capped at 36pt; their 0.72pt nominal thickness scales with the page. Independent
Word controls cover 36–144pt margins, right-margin changes, landscape and a
comment-fitted page. Vertical bar extents remain approximate.

In the Word-compatible profile, source-owned insertion/deletion authors occupy
comment palette slots before comment authors, independent of the body anchor
order. Author identity preserves name casing; repeated exact names share a slot.
Comment IDs retain their qualified numeric ordering and the twenty-color cycle.
Balloon fills/outlines and body range washes/brackets share the same slots. The
legacy revision-color hash keeps its existing normalization. Registration of
nontext/property/move-only authors and inline multi-author ink remains approximate.
