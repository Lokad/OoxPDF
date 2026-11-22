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
            foreach (bool hiddenBeforeBookmark in new[] { true, false })
            {
                var pages = Render(hiddenBeforeBookmark);
                TestAssert.True(control.Select(page => page.Content).SequenceEqual(pages.Select(page => page.Content)),
                    "A hidden suffix must preserve both pages' visible content.");
                var actual = pages[0].Annotations.OrderBy(link => link.X).ToArray();
                TestAssert.Equal(2, actual.Length);
                TestAssert.True(expected.SequenceEqual(actual),
                    $"The terminal text boundary must keep its whole link: {mode}/{cell}/{content}/{hiddenBeforeBookmark}.");
            }

            PdfPage[] Render(bool? hiddenBeforeBookmark) => renderer.RenderPagesAsync(
                ReadTerminalBookmarkDocument(mode, cell, content, hiddenBeforeBookmark), null, default).ToBlockingEnumerable().ToArray();
        }
    }

    private static DocxDocument ReadTerminalBookmarkDocument(OoxPdfDocxMarkupMode mode, bool cell, string contentKind,
        bool? hiddenBeforeBookmark) => ReadMutation("CompositionProbes/body", mode, false, false, (xml, w, m) =>
    {
        var body = xml.Root!.Element(w + "body")!;
        var original = body.Element(w + "p")!;
        var section = body.Element(w + "sectPr")!;
        section.Elements().Where(e => e.Name == w + "headerReference" || e.Name == w + "footerReference").Remove();

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
            _ => new XElement(original.Element(m + "oMath")!)
        };
        var owner = new XElement(w + "p", Run("Terminal owner "), Start("701", "ObjectStart"), content, End("701"));
        if (contentKind == "mixed") owner.Add(Run(" plain ending"));
        if (hiddenBeforeBookmark is true) owner.Add(HiddenRun());
        owner.Add(Start("702", "ObjectEnd"), End("702"));
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
            owner.AddFirst(new XElement(w + "pPr", new XElement(w + "pageBreakBefore")));
            original.ReplaceWith(links, owner, following);
        }
    });
}
