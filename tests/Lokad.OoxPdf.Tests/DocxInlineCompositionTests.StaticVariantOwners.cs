using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxDisabledStaticPropertyOwnersKeepTheBodyPrintScale()
    {
        var context = DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup,
            OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        foreach (string story in new[] { "header", "footer" })
        foreach (string variant in new[] { "first", "even" })
        {
            var source = ReadStaticVariantPropertySource(story, variant, OoxPdfDocxMarkupMode.AllMarkup, history: true, enabled: false);
            TestAssert.Equal("default", DocxLayoutEngine.GetStaticStoryVariant(source.PageSettings, 1));
            TestAssert.Equal("default", DocxLayoutEngine.GetStaticStoryVariant(source.PageSettings, 2));
            TestAssert.True(DocxRenderer.ResolveWordCompatiblePrintScale(source, context) == 1d,
                story + "/" + variant + ": an unreachable formatting owner cannot shrink the printed body.");
        }
    }

    public static void DocxDisabledStaticPropertyHistoryKeepsBodyPaintInBothViews()
    {
        foreach (string story in new[] { "header", "footer" })
        foreach (string variant in new[] { "first", "even" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode, CachedMathGeometry(mode));
            var control = ReadStaticVariantPropertySource(story, variant, mode, history: false, enabled: false);
            var source = ReadStaticVariantPropertySource(story, variant, mode, history: true, enabled: false);
            var expected = renderer.RenderPagesAsync(control, null, default).ToBlockingEnumerable().Single();
            var actual = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().Single();
            TestAssert.Equal(0, renderer.InspectMarkupBalloons(source).Count);
            TestAssert.True(expected.Content == actual.Content,
                story + "/" + variant + "/" + mode + ": history in an unselected story must preserve the complete body paint.");
            string text = DocxPageFootnoteNumberingTests.DecodeText(actual);
            TestAssert.Contains("Before public math table.", text);
            TestAssert.Contains("After public math table.", text);
            TestAssert.True(!text.Contains("Before inline", StringComparison.Ordinal), "The disabled story stays unprinted.");
            using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(source, null, cancellation.Token)
                .ToBlockingEnumerable().ToArray());
        }
    }

    public static void DocxEnabledStaticPropertyOwnersKeepVisibleCaptionsAndMath()
    {
        var context = DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup,
            OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        foreach (string story in new[] { "header", "footer" })
        foreach (string variant in new[] { "first", "even" })
        {
            var source = ReadStaticVariantPropertySource(story, variant, OoxPdfDocxMarkupMode.AllMarkup, history: true, enabled: true);
            int selectedPage = variant == "first" ? 0 : 1;
            if (selectedPage == 1) source = source with { BodyElements = source.BodyElements.Concat<DocxBodyElement>(
                [new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null),
                new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Second page body.", 12d, 12d))]).ToArray() };
            TestAssert.True(DocxRenderer.ResolveWordCompatiblePrintScale(source, context) < 1d,
                "An enabled source variant retains its printed caption lane.");
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true),
                OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            TestAssert.Equal(selectedPage, renderer.InspectMarkupBalloons(source).Single().PageIndex);
            var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(selectedPage + 1, pages.Length);
            string text = DocxPageFootnoteNumberingTests.DecodeText(pages[selectedPage]);
            TestAssert.Contains("Before inline", text);
            TestAssert.Contains("after equation.", text);
            TestAssert.Contains("Not Bold", text);
            var page = pages[selectedPage];
            var math = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).Where(show => page.Fonts.Any(font =>
                font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(font.Font, show.Font))).ToArray();
            TestAssert.Equal(4, math.Length);
        }
    }

    public static void DocxInheritedFirstStoryActivationKeepsItsRevisionLane()
    {
        var context = DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup,
            OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        foreach (string story in new[] { "header", "footer" })
        {
            var source = ReadInheritedFirstStory(story, OoxPdfDocxMarkupMode.AllMarkup, formattedTable: true);
            TestAssert.True(DocxRenderer.ResolveWordCompatiblePrintScale(source, context) < 1d,
                "Filtering raw section dictionaries must not lose a later activated inherited owner.");
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true),
                OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            TestAssert.Equal(1, renderer.InspectMarkupBalloons(source).Single().PageIndex);
            var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(2, pages.Length);
            TestAssert.Contains("series.", DocxPageFootnoteNumberingTests.DecodeText(pages[1]));
        }
    }

    private static DocxDocument ReadInheritedFirstStory(string story, OoxPdfDocxMarkupMode mode, bool formattedTable)
    {
        var source = ReadSectionScopedStory(story, mode, sectioned: true, formattedTable: formattedTable);
        var first = source.BodyElements.OfType<DocxSectionBreakElement>().Single();
        var earlier = first with { PageSettings = StaticVariantSettings(first.PageSettings, story, "first", enabled: false) };
        var final = StaticVariantSettings(source.PageSettings, story, "first", enabled: true);
        // The final section activates its predecessor's first story without
        // defining another reference. This exercises layout's inheritance.
        final = final with { HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(),
            FooterBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>() };
        return WithStaticVariantSettings(source, final) with
        { BodyElements = source.BodyElements.Select(element => ReferenceEquals(element, first) ? earlier : element).ToArray() };
    }

    private static DocxDocument ReadStaticVariantPropertySource(string story, string variant, OoxPdfDocxMarkupMode mode,
        bool history, bool enabled)
    {
        var source = ReadContainerPropertySource(story, mode, history);
        return WithStaticVariantSettings(source, StaticVariantSettings(source.PageSettings, story, variant, enabled));
    }

    private static DocxPageSettings StaticVariantSettings(DocxPageSettings settings, string story, string variant, bool enabled)
    {
        settings = settings with { TitlePage = variant == "first" && enabled, EvenAndOddHeaders = variant == "even" && enabled };
        return story == "header" ? settings with
        {
            HeaderBodyElementsByType = settings.HeaderBodyElementsByType.ToDictionary(_ => variant, pair => pair.Value, StringComparer.OrdinalIgnoreCase),
            HeaderFloatingDrawingsByType = settings.HeaderFloatingDrawingsByType.ToDictionary(_ => variant, pair => pair.Value, StringComparer.OrdinalIgnoreCase)
        } : settings with
        {
            FooterBodyElementsByType = settings.FooterBodyElementsByType.ToDictionary(_ => variant, pair => pair.Value, StringComparer.OrdinalIgnoreCase),
            FooterFloatingDrawingsByType = settings.FooterFloatingDrawingsByType.ToDictionary(_ => variant, pair => pair.Value, StringComparer.OrdinalIgnoreCase)
        };
    }

    private static DocxDocument WithStaticVariantSettings(DocxDocument source, DocxPageSettings settings) => source with
    {
        PageSettings = settings, HeaderParagraphs = [], FooterParagraphs = [],
        HeaderBodyElementsByType = settings.HeaderBodyElementsByType, FooterBodyElementsByType = settings.FooterBodyElementsByType,
        HeaderFloatingDrawingsByType = settings.HeaderFloatingDrawingsByType, FooterFloatingDrawingsByType = settings.FooterFloatingDrawingsByType
    };
}
