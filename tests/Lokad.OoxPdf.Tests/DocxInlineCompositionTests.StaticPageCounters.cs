using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxActiveMarginStoriesKeepPageLocalNotePaints()
    {
        foreach (string fixture in new[] { "math-page-notes-arial-11-zero", "cell-page-notes-arial-11-zero" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode, CachedMathGeometry(mode));
            var control = renderer.RenderPagesAsync(ReadInactiveNoteTemplate(fixture, mode), null, default).ToBlockingEnumerable().ToArray();
            foreach (string story in new[] { "header", "footer" })
            foreach (string variant in new[] { "default", "first", "even" })
            {
                var source = ReadInactiveNoteTemplate(fixture, mode, story, variant, enabled: true);
                var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
                TestAssert.Equal(2, pages.Length);
                for (int index = 0; index < pages.Length; index++)
                {
                    string text = DocxPageFootnoteNumberingTests.DecodeText(pages[index]);
                    bool caption = variant == "default" || (variant == "first" ? index == 0 : index == 1);
                    TestAssert.Equal(caption, text.Contains("UNPRINTED STATIC TEMPLATE", StringComparison.Ordinal));
                    TestAssert.Equal(DocxPageFootnoteNumberingTests.DecodeText(control[index]).Trim(),
                        text.Replace("UNPRINTED STATIC TEMPLATE", "", StringComparison.Ordinal).Trim());
                    TestAssert.Equal(control[index].Images.Count, pages[index].Images.Count);
                    var mathFonts = pages[index].Fonts.Where(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal))
                        .Select(font => font.Font).ToHashSet();
                    string markers = string.Join('|', DocxCommentsTests.ReadEmbeddedGlyphTextShows(pages[index], true)
                        .Where(show => show.Size < 9d && !mathFonts.Contains(show.Font) &&
                            show.Text.Length != 0 && show.Text.All(char.IsAsciiDigit)).Select(show => show.Text));
                    TestAssert.Equal("09|10|09|10", markers);
                }
            }
        }
    }

    public static void DocxStaticPageCountersRequireCompleteStableOwners()
    {
        var source = ReadInactiveNoteTemplate("math-page-notes-arial-11-zero", OoxPdfDocxMarkupMode.Final,
            "header", "first", enabled: true);
        var equations = source.Paragraphs.SelectMany(paragraph => paragraph.Equations).ToDictionary(equation => equation,
            equation => new DocxPreparedMathParagraph(equation, new(18d, 9d, 4d, []), 12d, 16d));
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        DocxLayout Layout(DocxDocument document) => engine.Create(document, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: equations);
        var first = Layout(source);
        TestAssert.True(first.Pages[0].StaticTextLines.Count > 0, "The control has measured active header paint.");
        TestAssert.True(!DocxPageFootnoteNumbering.IsEligibleDocument(source), "Typography retains its existing static-story scope.");
        TestAssert.True(DocxPageFootnoteNumbering.TryCreateDocument(source, first, out var numbered, out var owners,
            allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences), "Complete measured owners determine counters.");
        var final = Layout(numbered);
        TestAssert.True(DocxPageFootnoteNumbering.HasStableNumberingPages(numbered, final, owners,
            allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences), "A bounded retry keeps every marker/note page.");
        TestAssert.True(!DocxPageFootnoteNumbering.HasStablePages(numbered, final, owners,
            allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences), "Font purposes retain stricter admission.");
        TestAssert.True(!DocxPageFootnoteNumbering.TryCreateDocument(numbered, final, out var unchanged, out _,
            allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences) && ReferenceEquals(numbered, unchanged),
            "Correct counters need no additional retry.");

        var reference = source.Paragraphs.SelectMany(paragraph => paragraph.InlineReferences).First();
        var note = first.Pages[0].PlacedRelatedStories.Single(item => item.StoryLayout.Story.Id == reference.Id);
        var line = first.Pages[0].Items.OfType<DocxTextLineLayout>().Single(item => item.Segments.Any(segment =>
            ReferenceEquals(segment.StyleRun, reference.DisplayRun)));
        var marker = line.Segments.Single(segment => ReferenceEquals(segment.StyleRun, reference.DisplayRun));
        foreach (string variant in new[] { "missing", "duplicate", "partial", "moved", "marker", "columns" })
        {
            var pages = first.Pages.ToArray();
            pages[0] = pages[0] with
            {
                PlacedRelatedStories = variant switch
                {
                    "missing" or "moved" => pages[0].PlacedRelatedStories.Where(item => !ReferenceEquals(item, note)).ToArray(),
                    "duplicate" => [.. pages[0].PlacedRelatedStories, note],
                    "partial" => pages[0].PlacedRelatedStories.Select(item => ReferenceEquals(item, note) ? note with { ContentTopOffset = 1d } : item).ToArray(),
                    _ => pages[0].PlacedRelatedStories
                },
                Items = variant == "marker" ? pages[0].Items.Select(item => ReferenceEquals(item, line)
                    ? line with { Segments = [.. line.Segments, marker] } : item).ToArray() : pages[0].Items,
                ColumnFrames = variant == "columns" ? [pages[0].ColumnFrames[0], pages[0].ColumnFrames[0]] : pages[0].ColumnFrames
            };
            if (variant == "moved") pages[1] = pages[1] with { PlacedRelatedStories = [.. pages[1].PlacedRelatedStories, note] };
            TestAssert.True(!DocxPageFootnoteNumbering.TryCreateDocument(source, first with { Pages = pages }, out var fallback, out _,
                allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences) && ReferenceEquals(source, fallback),
                "Incomplete or ambiguous owners reject numbering atomically: " + variant);
        }
        var moved = final.Pages.ToArray();
        var finalNote = moved[0].PlacedRelatedStories.Single(item => item.StoryLayout.Story.Id == reference.Id);
        moved[0] = moved[0] with { PlacedRelatedStories = moved[0].PlacedRelatedStories.Where(item => !ReferenceEquals(item, finalNote)).ToArray() };
        moved[1] = moved[1] with { PlacedRelatedStories = [.. moved[1].PlacedRelatedStories, finalNote] };
        TestAssert.True(!DocxPageFootnoteNumbering.HasStableNumberingPages(numbered, final with { Pages = moved }, owners,
            allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences), "Moved retry owners preserve the original layout.");
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxPageFootnoteNumbering.TryCreateDocument(source, first, out _, out _,
            allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences, cancellationToken: cancellation.Token));
        TestAssert.Throws<OperationCanceledException>(() => DocxPageFootnoteNumbering.HasStableNumberingPages(numbered, final, owners,
            allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences, cancellationToken: cancellation.Token));
    }
}
