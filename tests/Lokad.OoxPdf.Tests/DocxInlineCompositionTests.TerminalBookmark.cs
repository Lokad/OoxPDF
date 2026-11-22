using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxTerminalBookmarksSurviveUnpaintedWhitespaceBeforePageBreak()
    {
        foreach (string fixture in new[] { "CompositionProbes/math-inner-bookmark-outer", "Composition/composed-static-cell-report" })
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            string name = fixture.StartsWith("Composition/", StringComparison.Ordinal) ? "ComposedCalculation" : "MathSource";
            var document = ReadMutation(fixture, mode, false, false, (xml, w, _) =>
            {
                var start = xml.Descendants(w + "bookmarkStart").Single(value => (string?)value.Attribute(w + "name") == name);
                start.AddAfterSelf(new XElement(w + "r", new XElement(w + "br", new XAttribute(w + "type", "page"))));
            });
            var owner = document.Paragraphs.Single(paragraph => paragraph.BookmarkAnchors.Any(value => value.Name == name));
            TestAssert.True(owner.Equations.Count == 0 && owner.Runs.Any(run => !string.IsNullOrWhiteSpace(run.Text)),
                "The source bookmark remains at the text fragment before the explicit break.");
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.AllMarkup ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                    : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var pages = renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.True(pages.Length >= 2 && pages.SelectMany(page => page.Annotations)
                .Any(link => link.Destination is { PageIndex: 0, Left: not null, Top: not null }),
                "The internal link must retain a concrete destination at the preceding painted source endpoint: " + fixture + "/" + mode);
        }
    }

    public static void DocxTerminalPictureAndMathBookmarksKeepTheirPlacedLinks()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool cell in new[] { false, true })
        foreach (bool picture in new[] { false, true })
        foreach (bool hiddenTail in new[] { false, true })
        {
            var document = ReadTerminalBookmarkDocument(mode, cell, picture ? "picture" : "math", hiddenTail ? false : null);
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.AllMarkup ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                    : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var pages = renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(2, pages.Length);
            var destinations = pages[0].Annotations.OrderBy(link => link.X).Select(link => link.Destination!.Value).ToArray();
            TestAssert.Equal(2, destinations.Length);
            TestAssert.Equal(1, destinations[0].PageIndex);
            TestAssert.Equal(1, destinations[1].PageIndex);
            TestAssert.True(destinations[0].Left is double start && destinations[1].Left is double end && end > start + 1d,
                $"The terminal bookmark uses the end of its placed object: {mode}/{cell}/{picture}/{hiddenTail}.");
            TestAssert.Equal(destinations[0].Top, destinations[1].Top);
            TestAssert.Equal(picture ? 1 : 0, pages[1].Images.Count);
            string paint = DocxPageFootnoteNumberingTests.DecodeText(pages[1]);
            TestAssert.Contains("Terminal owner", paint);
            TestAssert.Contains("Following text", paint);
            TestAssert.DoesNotContain("HiddenTail", paint);
        }
    }

    public static void DocxTerminalTextBookmarksKeepLinksAcrossHiddenSuffixes()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool cell in new[] { false, true })
        foreach (string content in new[] { "text", "mixed" })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.AllMarkup ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                    : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var control = Render(null);
            TestAssert.Equal(2, control.Length);
            var expected = control[0].Annotations.OrderBy(link => link.X).ToArray();
            TestAssert.Equal(2, expected.Length);
            TestAssert.Equal(1, expected[1].Destination!.Value.PageIndex);
            TestAssert.Contains("plain ending", DocxPageFootnoteNumberingTests.DecodeText(control[1]));
            (bool? HiddenBefore, char? BreakMarker)[] suffixes =
                [(true, null), (false, null), (null, '\u00AD'), (null, '\u200B')];
            foreach (var (hiddenBeforeBookmark, breakMarker) in suffixes)
            {
                var pages = Render(hiddenBeforeBookmark, breakMarker);
                TestAssert.Equal(control.Length, pages.Length);
                TestAssert.True(control.Select(DocxPageFootnoteNumberingTests.DecodeText)
                    .SequenceEqual(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)),
                    "The unpainted suffix must preserve decoded content on each page.");
                var actual = pages[0].Annotations.OrderBy(link => link.X).ToArray();
                TestAssert.True(expected.SequenceEqual(actual),
                    $"The terminal text boundary must retain both links: {mode}/{cell}/{content}/{hiddenBeforeBookmark}/{(int?)breakMarker}.");
            }

            PdfPage[] Render(bool? hiddenBeforeBookmark, char? breakMarker = null)
            {
                return renderer.RenderPagesAsync(
                    ReadTerminalBookmarkDocument(mode, cell, content, hiddenBeforeBookmark, AddBreakMarker),
                    null, default).ToBlockingEnumerable().ToArray();

                void AddBreakMarker(XElement anchor)
                {
                    XNamespace w = anchor.Name.Namespace;
                    if (breakMarker == '\u00AD')
                        anchor.Parent!.Elements(w + "r").Last().Add(new XElement(w + "softHyphen"));
                    else if (breakMarker is { } marker)
                        anchor.AddBeforeSelf(new XElement(w + "r", new XElement(w + "t", marker.ToString())));
                }
            }
        }
    }

    public static void DocxTerminalInlineObjectsKeepBookmarkLinksAcrossUnpaintedSuffixes()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool cell in new[] { false, true })
        foreach (string content in new[] { "picture", "math", "box" })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.AllMarkup ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                    : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var control = Render(null);
            TestAssert.Equal(2, control.Length);
            var links = control[0].Annotations.OrderBy(link => link.X).ToArray();
            TestAssert.Equal(2, links.Length);
            TestAssert.True(links.All(link => link.Destination is { PageIndex: 1 }),
                $"Both {content} links must land on the second page: {mode}/{cell}.");
            TestAssert.True(links[1].Destination!.Value.Left > links[0].Destination!.Value.Left,
                $"The end of the placed {content} must follow its start: {mode}/{cell}.");

            (bool Before, string? Text)[] suffixes = [(true, null), (false, "\u200B"), (true, "  "), (false, "  ")];
            foreach (var (before, text) in suffixes)
            {
                var pages = RenderSuffix(before, text);
                TestAssert.True(control.Select(DocxPageFootnoteNumberingTests.DecodeText)
                    .SequenceEqual(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)),
                    $"The unpainted suffix must preserve decoded page text: {mode}/{cell}/{content}/{before}/{text?.Length}.");
                TestAssert.True(links.SequenceEqual(pages[0].Annotations.OrderBy(link => link.X)),
                    $"The unpainted suffix must preserve the placed {content} endpoint: {mode}/{cell}/{before}/{text?.Length}.");
            }

            var beforeNoBreak = RenderSuffix(true, "\u00A0\u00A0")[0].Annotations.OrderBy(link => link.X).ToArray();
            var afterNoBreak = RenderSuffix(false, "\u00A0\u00A0")[0].Annotations.OrderBy(link => link.X).ToArray();
            TestAssert.True(beforeNoBreak.Length == 2 && afterNoBreak.Length == 2 &&
                beforeNoBreak[1].Destination is { PageIndex: 1, Left: double beforeX } &&
                afterNoBreak[1].Destination is { PageIndex: 1, Left: double afterX } && beforeX > afterX,
                $"A no-break space must retain its authored advance on the target page: {mode}/{cell}/{content}.");

            PdfPage[] RenderSuffix(bool before, string? text) => Render(anchor =>
            {
                XNamespace w = anchor.Name.Namespace;
                var suffix = new XElement(w + "r", text is null ? new XElement(w + "softHyphen") :
                    new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), text));
                if (before) anchor.AddBeforeSelf(suffix);
                else anchor.ElementsAfterSelf().First().AddAfterSelf(suffix);
            });

            PdfPage[] Render(Action<XElement>? addSuffix) => renderer.RenderPagesAsync(
                ReadTerminalBookmarkDocument(mode, cell, content, null, addSuffix),
                null, default).ToBlockingEnumerable().ToArray();
        }
    }

    public static void DocxTerminalTabBookmarksFollowBothPlacedEdges()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool cell in new[] { false, true })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.AllMarkup ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                    : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var control = Render(null);
            var clean = EndLink(control);
            var beforePages = RenderTab(4320, afterTab: false);
            var nearStart = EndLink(beforePages);
            var farStart = EndLink(RenderTab(5040, afterTab: false));
            var afterPages = RenderTab(4320, afterTab: true);
            var nearEnd = EndLink(afterPages);
            var farEnd = EndLink(RenderTab(5040, afterTab: true));
            TestAssert.Equal(nearStart, farStart);
            TestAssert.Equal(nearStart.Top, nearEnd.Top);
            TestAssert.True(nearEnd.Left > nearStart.Left &&
                Math.Abs(farEnd.Left!.Value - nearEnd.Left!.Value - 36d) < .001d,
                $"Tab endpoints must follow the two authored stops on page two: {mode}/{cell}.");
            TestAssert.Equal(clean, EndLink(RenderTab(4320, afterTab: true, hidden: true)));
            TestAssert.Equal(clean, EndLink(RenderTab(4320, afterTab: false, hidden: true)));
            TestAssert.True(afterPages.Select(page => page.Content).SequenceEqual(beforePages.Select(page => page.Content)),
                "Moving the bookmark across the tab must preserve the page's paint commands.");
            TestAssert.Contains("Following text", DocxPageFootnoteNumberingTests.DecodeText(afterPages[1]));

            PdfPage[] Render(Action<XElement>? mutate) => renderer.RenderPagesAsync(
                ReadTerminalBookmarkDocument(mode, cell, "math", null, mutate), null, default).ToBlockingEnumerable().ToArray();

            PdfPage[] RenderTab(int position, bool afterTab, bool hidden = false) => Render(anchor =>
            {
                XNamespace w = anchor.Name.Namespace;
                var owner = anchor.Parent!;
                var properties = owner.Element(w + "pPr");
                if (properties is null) owner.AddFirst(properties = new XElement(w + "pPr"));
                properties.Add(new XElement(w + "tabs", new XElement(w + "tab",
                    new XAttribute(w + "val", "left"), new XAttribute(w + "pos", position))));
                var tab = new XElement(w + "r", new XElement(w + "tab"));
                if (hidden) tab.AddFirst(new XElement(w + "rPr", new XElement(w + "vanish")));
                if (afterTab) anchor.AddBeforeSelf(tab);
                else anchor.ElementsAfterSelf().First().AddAfterSelf(tab);
            });

            PdfLinkDestination EndLink(PdfPage[] pages)
            {
                TestAssert.Equal(2, pages.Length);
                var links = pages[0].Annotations.OrderBy(link => link.X).ToArray();
                TestAssert.Equal(2, links.Length);
                TestAssert.True(links[1].Destination is { PageIndex: 1, Left: not null, Top: not null },
                    $"The terminal link must have a complete destination on page two: {mode}/{cell}.");
                return links[1].Destination!.Value;
            }
        }
    }

    private static DocxDocument ReadTerminalBookmarkDocument(OoxPdfDocxMarkupMode mode, bool cell, string contentKind,
        bool? hiddenBeforeBookmark, Action<XElement>? addSuffix = null) => ReadMutation(
        contentKind == "box" ? "InlineBoxMath/inline-box-body-equation-inline-box-valid" : "CompositionProbes/body",
        mode, false, false, (xml, w, m) =>
    {
        var body = xml.Root!.Element(w + "body")!;
        var original = body.Elements(w + "p").First(paragraph => contentKind != "box" ||
            paragraph.Descendants(w + "txbxContent").Any());
        var section = body.Element(w + "sectPr")!;
        section.Elements().Where(e => e.Name == w + "headerReference" || e.Name == w + "footerReference").Remove();
        if (contentKind == "box")
            body.Elements().Where(element => element != original && element != section).Remove();

        XElement Run(string value) => new(w + "r", new XElement(w + "t",
            new XAttribute(XNamespace.Xml + "space", "preserve"), value));
        XElement Link(string name, string label) => new(w + "hyperlink", new XAttribute(w + "anchor", name), Run(label));
        XElement Start(string id, string name) => new(w + "bookmarkStart", new XAttribute(w + "id", id), new XAttribute(w + "name", name));
        XElement End(string id) => new(w + "bookmarkEnd", new XAttribute(w + "id", id));
        XElement HiddenRun() => new(w + "r", new XElement(w + "rPr", new XElement(w + "vanish")), new XElement(w + "t", "HiddenTail"));

        var links = new XElement(w + "p", Link("ObjectStart", "Jump to start"), Run(" / "), Link("ObjectEnd", "Jump to end"));
        var content = contentKind switch
        {
            "picture" => new XElement(original.Element(w + "r")!),
            "text" => Run("plain ending"),
            "box" => new XElement(original.Elements(w + "r").Single(run => run.Descendants(w + "txbxContent").Any())),
            _ => new XElement(original.Element(m + "oMath")!)
        };
        var owner = new XElement(w + "p", Run("Terminal owner "));
        if (!cell) owner.AddFirst(new XElement(w + "pPr", new XElement(w + "pageBreakBefore")));
        if (contentKind == "box") owner.Add(new XElement(original.Element(m + "oMath")!));
        owner.Add(Start("701", "ObjectStart"), content, End("701"));
        if (contentKind == "mixed") owner.Add(Run(" plain ending"));
        if (hiddenBeforeBookmark is true) owner.Add(HiddenRun());
        var anchor = Start("702", "ObjectEnd");
        owner.Add(anchor, End("702"));
        addSuffix?.Invoke(anchor);
        if (hiddenBeforeBookmark is false) owner.Add(HiddenRun());

        var following = new XElement(w + "p", Run("Following text"));
        if (cell)
        {
            var pageBreak = new XElement(w + "p", new XElement(w + "r", new XElement(w + "br", new XAttribute(w + "type", "page"))));
            var table = new XElement(w + "tbl", new XElement(w + "tblGrid", new XElement(w + "gridCol", new XAttribute(w + "w", "8640"))),
                new XElement(w + "tr", new XElement(w + "tc", owner)));
            original.ReplaceWith(links, pageBreak, table, following);
        }
        else
        {
            original.ReplaceWith(links, owner, following);
        }
    });
}
