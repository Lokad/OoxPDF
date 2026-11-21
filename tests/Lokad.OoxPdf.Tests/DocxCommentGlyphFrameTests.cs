using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxCommentGlyphFrameTests
{
    public static void DocxCommentGlyphFramesMatchIndependentWordPaths()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent comment glyph frames require installed Windows fonts.");
        using var sources = ReadJson("sources.json");
        using var gold = ReadJson("office.json");
        foreach (var source in sources.RootElement.EnumerateArray())
        {
            var expected = gold.RootElement.GetProperty("Rows").EnumerateArray()
                .Single(row => row.GetProperty("Id").GetString() == source.GetProperty("Id").GetString());
            TestAssert.Equal(source.GetProperty("Sha256").GetString(), expected.GetProperty("InputSha256").GetString());
            var page = Render(Read(source, default), null);
            var actual = ReadGeometry(page);
            TestAssert.True(Math.Abs(actual.StrokeWidth - expected.GetProperty("StrokeWidth").GetDouble()) < .001d,
                "Comment brackets retain the physical Office stroke across print scales.");
            for (int path = 0; path < actual.Paths.Length; path++)
                for (int axis = 0; axis < actual.Paths[path].Length; axis++)
                    TestAssert.True(Math.Abs(actual.Paths[path][axis] - expected.GetProperty("Paths")[path][axis].GetDouble()) < 1d,
                        "All bracket corners and stems retain the absolute Word threshold: " + source.GetProperty("Id").GetString() + "/" + path + "/" + axis);
            AssertContent(page, source);
        }
    }

    public static void DocxCommentGlyphFramesKeepBookmarksRunSplitsAndCancellationNeutral()
    {
        using var sources = ReadJson("sources.json");
        var resolver = new PortableRenderingTests.PinnedResolver();
        foreach (var source in sources.RootElement.EnumerateArray().Where(row => row.GetProperty("PortableProbe").GetBoolean()))
        {
            var original = ReadGeometry(Render(Read(source, default), resolver));
            TestAssert.True(Math.Abs(original.StrokeWidth - .14d) < .001d,
                "Measured portable prose brackets retain a physical stroke: " + source.GetProperty("Id").GetString());
            foreach ((bool bookmark, bool split) in new[] { (true, false), (false, true), (true, true) })
            {
                var document = DocxInlineCompositionTests.ReadMutation(Fixture(source), OoxPdfDocxMarkupMode.AllMarkup, false, split,
                    (xml, w, _) =>
                    {
                        if (!bookmark) return;
                        var paragraph = xml.Descendants(w + "p").Single(p => p.Elements(w + "commentRangeStart").Any());
                        var start = new XElement(w + "bookmarkStart", new XAttribute(w + "id", "8230"), new XAttribute(w + "name", "PublicCommentGlyphFrame"));
                        if (paragraph.Element(w + "pPr") is { } properties) properties.AddAfterSelf(start); else paragraph.AddFirst(start);
                        paragraph.Add(new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "8230")));
                    });
                var page = Render(document, resolver);
                var transformed = ReadGeometry(page);
                AssertPathsEqual(original.Paths, transformed.Paths);
                TestAssert.Equal(original.StrokeWidth, transformed.StrokeWidth);
                AssertContent(page, source);
            }
        }
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => Read(sources.RootElement[0], canceled.Token));
        var prepared = Read(sources.RootElement[0], default);
        TestAssert.Throws<OperationCanceledException>(() => Render(prepared, resolver, canceled.Token));
    }

    public static void DocxCommentGlyphFramesSeparateNaturalBracketsFromAuthoredSpacing()
    {
        using var sources = ReadJson("sources.json");
        var source = sources.RootElement.EnumerateArray().Single(row => row.GetProperty("Id").GetString() == "arial12-auto");
        var resolver = new PortableRenderingTests.PinnedResolver();
        double[][]? originalShape = null;
        double originalFill = 0d;
        foreach ((string rule, int twips) in new[] { ("auto", 240), ("atLeast", 1200), ("exact", 640), ("exact", 1200) })
        {
            var document = DocxInlineCompositionTests.ReadMutation(Fixture(source), OoxPdfDocxMarkupMode.AllMarkup, false, false,
                (xml, w, _) =>
                {
                    var paragraph = xml.Descendants(w + "p").Single(p => p.Elements(w + "commentRangeStart").Any());
                    var spacing = paragraph.Element(w + "pPr")!.Element(w + "spacing")!;
                    spacing.SetAttributeValue(w + "lineRule", rule);
                    spacing.SetAttributeValue(w + "line", twips);
                });
            var actual = ReadGeometry(Render(document, resolver));
            double top = actual.Paths[1][3];
            double[][] shape = actual.Paths.Select(path => new[] { path[0], path[1] - top, path[2], path[3] - top }).ToArray();
            if (originalShape is null) { originalShape = shape; originalFill = actual.FillHeight; }
            else AssertPathsEqual(originalShape, shape);
            if (rule == "atLeast")
                TestAssert.True(Math.Abs(actual.FillHeight - originalFill) < .002d, "Extra minimum leading does not enlarge the natural selection strut.");
            if (rule == "exact")
            {
                double scale = DocxRenderer.ResolveWordCompatiblePrintScale(document,
                    DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup));
                TestAssert.True(Math.Abs(actual.FillHeight - twips / 20d * scale) < .002d,
                    "Exact spacing still owns the selection height while glyph brackets retain their natural shape.");
            }
        }
    }

    public static void DocxCommentGlyphFramesRetainUnmeasuredFontFallback()
    {
        using var sources = ReadJson("sources.json");
        var source = sources.RootElement.EnumerateArray().Single(row => row.GetProperty("Id").GetString() == "arial10-auto");
        // This math font's global Windows descent exceeds its ordinary text
        // strut. That existing unmeasured selection fallback must stay explicit.
        var page = Render(Read(source, default), new PortableRenderingTests.PinnedResolver(mathFont: true));
        TestAssert.True(Math.Abs(ReadGeometry(page).StrokeWidth - .475d) < .001d,
            "An unmeasured prose selection retains its established marker path.");
        AssertContent(page, source);
    }

    public static void DocxCommentGlyphFramesCarryOneSelectionAcrossAuthoredBreaks()
    {
        var resolver = new PortableRenderingTests.PinnedResolver();
        var renderer = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup,
            OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        foreach ((string Break, int Count, int[] Fills) sample in new[]
        {
            ("textWrapping", 1, new[] { 2 }),
            ("page", 1, new[] { 1, 1 }),
            ("column", 1, new[] { 2 }),
            ("page", 2, new[] { 1, 1, 1 }),
        })
        {
            DocxDocument source = Source(sample.Break, sample.Count, close: true);
            var paragraphs = DocxBlockTraversal.EnumerateBodyParagraphs(source).ToArray();
            TestAssert.Equal(1, paragraphs.Sum(paragraph => paragraph.InlineReferences.Count(r => r.Kind == DocxRelatedStoryKind.Comment)));
            if (sample.Break != "textWrapping")
            {
                TestAssert.Equal(sample.Count + 1, paragraphs.Length);
                var ranges = paragraphs[0].SplitComments?.Ranges;
                TestAssert.NotNull(ranges);
                TestAssert.True(paragraphs.All(p => ReferenceEquals(ranges, p.SplitComments?.Ranges)),
                    "All covered fragments share the authored comment range.");
            }
            var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.True(pages.Select(page => DocxTests.CountWordCompatibleCommentRangeFills(page.Content)).SequenceEqual(sample.Fills),
                "Each selected text fragment retains its source fill across " + sample.Break);
            string text = Regex.Replace(string.Concat(pages.SelectMany(page =>
                DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Select(show => show.Text))), @"\s+", "");
            foreach (string word in new[] { "OutsideLeft", "SelectedLeft", "SelectedRight", "OutsideRight" })
                TestAssert.Contains(word, text);
            TestAssert.Equal(1, Regex.Matches(text, "Table-cellcommentbody\\.").Count);
        }

        DocxDocument unclosed = Source("page", 1, close: false);
        TestAssert.True(DocxBlockTraversal.EnumerateBodyParagraphs(unclosed).All(p => p.SplitComments is null),
            "Unclosed ranges retain their previous fallback.");
        DocxDocument hidden = Source("page", 1, close: true);
        hidden = hidden with { Settings = hidden.Settings with
        {
            RevisionViewSettings = hidden.Settings.RevisionViewSettings with { ShowMarkup = false }
        } };
        TestAssert.Equal(0, renderer.RenderPagesAsync(hidden, null, default).ToBlockingEnumerable()
            .Sum(page => DocxTests.CountWordCompatibleCommentRangeFills(page.Content)));

        static DocxDocument Source(string kind, int count, bool close) =>
            DocxInlineCompositionTests.ReadMutation("docx-markup-comment-table", OoxPdfDocxMarkupMode.AllMarkup, false, false,
                (xml, w, _) =>
                {
                    XElement body = xml.Root!.Element(w + "body")!;
                    var run = new XElement(w + "r", new XElement(w + "t", "SelectedLeft"));
                    for (int index = 0; index < count; index++)
                    {
                        run.Add(new XElement(w + "br", new XAttribute(w + "type", kind)),
                            new XElement(w + "t", index == count - 1 ? "SelectedRight" : "SelectedMiddle"));
                    }
                    body.ReplaceNodes(new XElement(w + "p",
                        Text("OutsideLeft "),
                        new XElement(w + "commentRangeStart", new XAttribute(w + "id", "1")), run,
                        close ? new XElement(w + "commentRangeEnd", new XAttribute(w + "id", "1")) : null,
                        new XElement(w + "r", new XElement(w + "commentReference", new XAttribute(w + "id", "1"))),
                        Text(" OutsideRight")),
                        new XElement(w + "sectPr", new XElement(w + "pgSz",
                            new XAttribute(w + "w", "12240"), new XAttribute(w + "h", "15840")),
                            kind == "column" ? new XElement(w + "cols", new XAttribute(w + "num", "2")) : null));

                    XElement Text(string value) => new(w + "r", new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), value));
                });
    }

    public static void DocxCommentGlyphFramesCarryClosedRangesAcrossBodyParagraphs()
    {
        var resolver = new PortableRenderingTests.PinnedResolver();
        var review = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup,
            OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        foreach ((string Kind, int[]? Fills) sample in new (string, int[]?)[]
        {
            ("adjacent", new[] { 2 }),
            ("wrapped-page", new[] { 2, 1 }),
            ("unclosed", null),
            ("early-reference", null),
            ("table", null),
        })
        {
            DocxDocument source = Source(sample.Kind, OoxPdfDocxMarkupMode.AllMarkup);
            var paragraphs = DocxBlockTraversal.EnumerateDirectParagraphs(source.BodyElements).ToArray();
            TestAssert.Equal(1, paragraphs.Sum(p => p.InlineReferences.Count(r => r.Kind == DocxRelatedStoryKind.Comment)));
            TestAssert.True(paragraphs.All(p => p.CommentRanges.All(range =>
                range.StartSourceRunIndex is null || range.EndSourceRunIndex is not null || range.ReferenceSourceRunIndex is not null)),
                "Unmatched starts are consumed before layout.");
            if (sample.Fills is not null)
            {
                var ranges = TestAssert.NotNull(paragraphs[0].SplitComments?.Ranges);
                foreach (var paragraph in paragraphs)
                {
                    TestAssert.True(paragraph.SplitComments is { } fragment && ReferenceEquals(ranges, fragment.Ranges),
                        "The closed interval covers every selected body paragraph.");
                    var selected = new List<DocxSplitCommentRange>();
                    ranges.VisitContainingPart(paragraph.SplitComments!.Value.Part, selected.Add, default);
                    TestAssert.Equal("1", selected.Single().End.Id);
                }
            }
            else
            {
                TestAssert.True(paragraphs.All(p => p.SplitComments is null),
                    "Unclosed, out-of-order and interrupted sources retain their established fallback.");
            }

            var pages = review.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            if (sample.Fills is { } fills)
                TestAssert.True(pages.Select(page => DocxTests.CountWordCompatibleCommentRangeFills(page.Content)).SequenceEqual(fills),
                    "Body selection paints on its owning pages for " + sample.Kind);
            string text = ReadText(pages);
            foreach (string word in new[] { "OutsideLeft", "SelectedLeft", "SelectedRight", "OutsideRight" })
                TestAssert.Contains(word, text);
            TestAssert.Equal(1, Regex.Matches(text, "Table-cellcommentbody\\.").Count);
            if (sample.Kind == "wrapped-page")
                TestAssert.Contains("SelectedMiddle", ReadText([pages[0]]));
        }

        var final = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.Final,
            OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        DocxDocument finalSource = Source("adjacent", OoxPdfDocxMarkupMode.Final);
        PdfPage finalPage = final.RenderPagesAsync(finalSource, null, default).ToBlockingEnumerable().Single();
        TestAssert.Equal(0, DocxTests.CountWordCompatibleCommentRangeFills(finalPage.Content));
        TestAssert.Contains("SelectedRight", ReadText([finalPage]));
        TestAssert.DoesNotContain("Table-cellcommentbody", ReadText([finalPage]));

        static string ReadText(IEnumerable<PdfPage> pages) => Regex.Replace(string.Concat(pages.SelectMany(page =>
            DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Select(show => show.Text))), @"\s+", "");

        static DocxDocument Source(string kind, OoxPdfDocxMarkupMode mode) =>
            DocxInlineCompositionTests.ReadMutation("docx-markup-comment-table", mode, false, false,
                (xml, w, _) =>
                {
                    XElement body = xml.Root!.Element(w + "body")!;
                    XElement? originalTable = kind == "table" ? body.Element(w + "tbl") : null;
                    originalTable?.Descendants(w + "commentReference").Remove();
                    originalTable?.Descendants(w + "commentRangeStart").Remove();
                    originalTable?.Descendants(w + "commentRangeEnd").Remove();
                    var start = new XElement(w + "p", Text("OutsideLeft "), Marker("commentRangeStart"), Text("SelectedLeft"));
                    var end = new XElement(w + "p",
                        kind == "wrapped-page" ? new XElement(w + "pPr", new XElement(w + "pageBreakBefore")) : null,
                        kind == "early-reference" ? Reference() : null,
                        Text("SelectedRight"), kind == "unclosed" ? null : Marker("commentRangeEnd"),
                        kind == "early-reference" ? null : Reference(), Text(" OutsideRight"));
                    body.ReplaceNodes(start,
                        kind == "wrapped-page" ? new XElement(w + "p", Text("SelectedMiddle")) : null,
                        originalTable,
                        kind == "wrapped-page" ? new XElement(w + "sdt", new XElement(w + "sdtContent", end)) : end,
                        new XElement(w + "sectPr", new XElement(w + "pgSz",
                            new XAttribute(w + "w", "12240"), new XAttribute(w + "h", "15840"))));

                    XElement Text(string value) => new(w + "r",
                        new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), value));
                    XElement Marker(string name) => new(w + name, new XAttribute(w + "id", "1"));
                    XElement Reference() => new(w + "r", Marker("commentReference"));
                });
    }

    public static void DocxCommentGlyphFramesPlaceBodyLevelRangeBoundaries()
    {
        var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), OoxPdfDocxMarkupMode.AllMarkup,
            OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        var ordinary = RenderSource(false, false, wrappedPage: false);
        var paged = RenderSource(false, false, wrappedPage: true);
        foreach ((bool Start, bool End, bool WrappedPage) sample in new[]
        {
            (true, false, false),
            (false, true, false),
            (true, true, false),
            (true, true, true),
        })
        {
            var expected = sample.WrappedPage ? paged : ordinary;
            var actual = RenderSource(sample.Start, sample.End, sample.WrappedPage);
            TestAssert.Equal(expected.Length, actual.Length);
            for (int page = 0; page < actual.Length; page++)
            {
                TestAssert.Equal(expected[page].Content, actual[page].Content);
            }
        }

        PdfPage[] inlineDelayed = RenderSource(false, false, wrappedPage: false, singleParagraph: true);
        PdfPage[] blockDelayed = RenderSource(true, true, wrappedPage: false, singleParagraph: true);
        TestAssert.Equal(blockDelayed.Single().Content, inlineDelayed.Single().Content);
        TestAssert.Contains("OutsideAfter", string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(inlineDelayed.Single())
            .Select(show => show.Text)));

        PdfPage[] RenderSource(bool outsideStart, bool outsideEnd, bool wrappedPage, bool singleParagraph = false)
        {
            DocxDocument source = DocxInlineCompositionTests.ReadMutation("docx-markup-comment-table",
                OoxPdfDocxMarkupMode.AllMarkup, false, false, (xml, w, _) =>
                {
                    var begin = Marker("commentRangeStart");
                    var end = Marker("commentRangeEnd");
                    var first = new XElement(w + "p", outsideStart ? null : begin, Text("ChosenOpening"),
                        singleParagraph && !outsideEnd ? end : null);
                    var last = singleParagraph ? null : new XElement(w + "p",
                        wrappedPage ? new XElement(w + "pPr", new XElement(w + "pageBreakBefore")) : null,
                        Text("ChosenClosing"), outsideEnd ? null : end);
                    var blocks = new object?[] { outsideStart ? begin : null, first, last, outsideEnd ? end : null };
                    xml.Root!.Element(w + "body")!.ReplaceNodes(
                        new XElement(w + "p", Text("OutsideBefore")),
                        wrappedPage ? new XElement(w + "sdt", new XElement(w + "sdtContent", blocks)) : blocks,
                        new XElement(w + "p", Text("OutsideAfter"), new XElement(w + "r", Marker("commentReference"))),
                        new XElement(w + "sectPr", new XElement(w + "pgSz",
                            new XAttribute(w + "w", "12240"), new XAttribute(w + "h", "15840"))));

                    XElement Text(string value) => new(w + "r", new XElement(w + "t", value));
                    XElement Marker(string name) => new(w + name, new XAttribute(w + "id", "1"));
                });
            var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(wrappedPage ? 2 : 1, pages.Length);
            TestAssert.True(pages.Sum(page => DocxTests.CountWordCompatibleCommentRangeFills(page.Content)) >= (singleParagraph ? 2 : 3),
                "The selection fills and the delayed reference mark keep their own paragraphs.");
            return pages;
        }
    }

    public static void DocxCommentGlyphFramesCarryClosedRangesWithinTableCells()
    {
        var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), OoxPdfDocxMarkupMode.AllMarkup,
            OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        foreach (string variant in new[] { "paragraphs", "wrapped", "cell-boundaries", "separate-cells" })
        {
            DocxDocument source = DocxInlineCompositionTests.ReadMutation("docx-markup-comment-table",
                OoxPdfDocxMarkupMode.AllMarkup, false, false, (xml, w, _) =>
                {
                    bool blockMarkers = variant == "cell-boundaries";
                    bool separateCells = variant == "separate-cells";
                    XElement start = Marker("commentRangeStart");
                    XElement end = Marker("commentRangeEnd");
                    var first = new XElement(w + "p", blockMarkers ? null : Text("OutsideLeft "),
                        blockMarkers ? null : start, Text("SelectedLeft"));
                    var second = new XElement(w + "p", Text("SelectedRight"), blockMarkers ? null : end,
                        blockMarkers ? null : new XElement(w + "r", Marker("commentReference")));
                    var after = new XElement(w + "p", Text("OutsideRight"),
                        blockMarkers ? new XElement(w + "r", Marker("commentReference")) : null);
                    var cells = separateCells
                        ? new[] { new XElement(w + "tc", first), new XElement(w + "tc", second, after) }
                        : [new XElement(w + "tc", blockMarkers ? start : null, first,
                            variant == "wrapped" ? new XElement(w + "sdt", new XElement(w + "sdtContent", second)) : second,
                            blockMarkers ? end : null, after)];
                    xml.Root!.Element(w + "body")!.ReplaceNodes(
                        new XElement(w + "tbl",
                            new XElement(w + "tblGrid", cells.Select(_ => new XElement(w + "gridCol", new XAttribute(w + "w", "4000")))),
                            new XElement(w + "tr", cells)),
                        new XElement(w + "p", Text("Following body")),
                        new XElement(w + "sectPr", new XElement(w + "pgSz",
                            new XAttribute(w + "w", "12240"), new XAttribute(w + "h", "15840"))));

                    XElement Text(string value) => new(w + "r",
                        new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), value));
                    XElement Marker(string name) => new(w + name, new XAttribute(w + "id", "1"));
                });
            var paragraphs = DocxBlockTraversal.EnumerateTableParagraphs(source.Tables.Single()).ToArray();
            TestAssert.Equal(1, paragraphs.Sum(p => p.InlineReferences.Count(r => r.Kind == DocxRelatedStoryKind.Comment)));
            if (variant == "separate-cells")
            {
                TestAssert.True(paragraphs.All(p => p.SplitComments is null), "Separate cells retain their own range scope.");
            }
            else
            {
                var ranges = TestAssert.NotNull(paragraphs[0].SplitComments?.Ranges);
                TestAssert.True(ReferenceEquals(ranges, paragraphs[1].SplitComments?.Ranges),
                    "Both selected paragraphs share the cell's closed range.");
                TestAssert.True(paragraphs[2].SplitComments is null, "The following paragraph is outside the selection.");
            }

            PdfPage page = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().Single();
            if (variant != "separate-cells")
            {
                TestAssert.True(DocxTests.CountWordCompatibleCommentRangeFills(page.Content) >= 2,
                    "Both selected cell paragraphs paint their fills for " + variant);
            }
            string text = Regex.Replace(string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page)
                .Select(show => show.Text)), @"\s+", "");
            foreach (string witness in new[] { "SelectedLeft", "SelectedRight", "OutsideRight", "Followingbody" })
                TestAssert.Contains(witness, text);
            TestAssert.Equal(1, Regex.Matches(text, "Table-cellcommentbody\\.").Count);
        }
    }

    public static void DocxCommentGlyphFramesKeepSelectionsAtConsumedParagraphMarks()
    {
        const OoxPdfDocxMarkupMode mode = OoxPdfDocxMarkupMode.AllMarkup;
        var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode,
            OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        foreach (bool cell in new[] { false, true })
        foreach (bool startBoundary in new[] { false, true })
        foreach (string? hardBreak in cell ? new string?[] { null } : new string?[] { null, "page", "column" })
        {
            PdfPage[] inline = RenderVariant(blockMarker: false);
            PdfPage[] block = RenderVariant(blockMarker: true);
            TestAssert.Equal(inline.Length, block.Length);
            for (int pageIndex = 0; pageIndex < inline.Length; pageIndex++)
                TestAssert.Equal(inline[pageIndex].Content, block[pageIndex].Content);

            PdfPage[] RenderVariant(bool blockMarker)
            {
                DocxDocument source = DocxInlineCompositionTests.ReadMutation("Composition/report", mode, false, false,
                    (xml, _, _) =>
                    {
                        var start = Marker("commentRangeStart");
                        var end = Marker("commentRangeEnd");
                        var first = new XElement(w + "p", new XElement(w + "pPr", new XElement(w + "rPr",
                            new XElement(w + "del", new XAttribute(w + "id", "31"), new XAttribute(w + "author", "Reviewer")))),
                            Text("OutsideLeft "), startBoundary ? null : start, startBoundary ? null : Text("SelectedLeft "));
                        var boundary = startBoundary ? start : end;
                        var second = new XElement(w + "p", blockMarker ? null : boundary,
                            startBoundary ? Text("SelectedRight") : null, startBoundary ? end : null,
                            new XElement(w + "r", Marker("commentReference")), Text(" OutsideRight"),
                            hardBreak is null ? null : new XElement(w + "r", new XElement(w + "br", new XAttribute(w + "type", hardBreak))),
                            hardBreak is null ? null : Text("AfterHardBreak"));
                        object[] flow = blockMarker ? [first, boundary, second] : [first, second];
                        xml.Root!.Element(w + "body")!.ReplaceNodes(
                            cell ? new XElement(w + "tbl",
                                new XElement(w + "tblGrid", new XElement(w + "gridCol", new XAttribute(w + "w", "7000"))),
                                new XElement(w + "tr", new XElement(w + "tc", flow))) : flow,
                            new XElement(w + "p", Text("Following body")),
                            new XElement(w + "sectPr", new XElement(w + "pgSz",
                                new XAttribute(w + "w", "12240"), new XAttribute(w + "h", "15840")),
                                hardBreak == "column" ? new XElement(w + "cols", new XAttribute(w + "num", "2")) : null));
                    }, packageChange: zip =>
                    {
                        const string part = "word/settings.xml";
                        var entry = zip.GetEntry(part)!;
                        XDocument settings;
                        using (var stream = entry.Open()) settings = XDocument.Load(stream);
                        settings.Root!.Element(w + "revisionView")?.Remove();
                        settings.Root.AddFirst(new XElement(w + "revisionView", new XAttribute(w + "markup", "1"),
                            new XAttribute(w + "comments", "1"), new XAttribute(w + "insDel", "0")));
                        entry.Delete();
                        using var output = zip.CreateEntry(part).Open();
                        settings.Save(output);
                    });
                string selected = startBoundary ? "SelectedRight" : "SelectedLeft";
                var selectedFlow = cell ? DocxBlockTraversal.EnumerateTableParagraphs(source.Tables.Single()) :
                    DocxBlockTraversal.EnumerateBodyParagraphs(source).Take(1);
                DocxParagraph merged = selectedFlow.Single();
                TestAssert.Equal(1, merged.InlineReferences.Count(reference => reference.Kind == DocxRelatedStoryKind.Comment));
                TestAssert.True(string.Concat(merged.Runs.Select(run => run.Text)).Contains(selected, StringComparison.Ordinal),
                    "The selected word remains in the merged paragraph.");
                PdfPage[] pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
                TestAssert.Equal(hardBreak == "page" ? 2 : 1, pages.Length);
                var fills = Regex.Matches(pages[0].Content, @"0\.973 0\.863 0\.867 rg\s+[-\d.]+ [-\d.]+ (?<w>[-\d.]+) [-\d.]+ re f");
                TestAssert.True(fills.Any(fill => Number(fill, "w") > 20d), "The selection covers the word at the consumed mark.");
                string text = Regex.Replace(string.Concat(pages.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page))
                    .Select(show => show.Text)), @"\s+", "");
                foreach (string word in new[] { "OutsideLeft", selected, "OutsideRight", "Followingbody" })
                    TestAssert.Contains(word, text);
                if (hardBreak is not null) TestAssert.Contains("AfterHardBreak", text);
                TestAssert.Equal(1, Regex.Matches(text, "Checkthecalculationalongsidetheillustration\\.").Count);
                return pages;
            }
        }

        XElement Text(string value) => new(w + "r", new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), value));
        XElement Marker(string kind) => new(w + kind, new XAttribute(w + "id", "0"));
    }

    private static (double[][] Paths, double StrokeWidth, double FillHeight) ReadGeometry(PdfPage page)
    {
        var fills = Regex.Matches(page.Content, @"0\.973 0\.863 0\.867 rg\s+(?<x>[-\d.]+) (?<y>[-\d.]+) (?<w>[\d.]+) (?<h>[\d.]+) re f");
        TestAssert.True(fills.Count > 0, "Selected prose retains its filled source rectangle.");
        var fill = fills.OrderBy(match => Number(match, "x")).First();
        string afterFill = page.Content[(fill.Index + fill.Length)..];
        var paths = Regex.Matches(afterFill, @"(?m)^(?<x>[-\d.]+) (?<y>[-\d.]+) m\s+(?<endX>[-\d.]+) (?<endY>[-\d.]+) l S")
            .Take(6).Select(match => new[] { Number(match, "x"), page.Height - Number(match, "y"), Number(match, "endX"), page.Height - Number(match, "endY") }).ToArray();
        TestAssert.Equal(6, paths.Length);
        var width = Regex.Match(afterFill, @"(?m)^(?<w>[\d.]+) w\r?$");
        TestAssert.True(width.Success, "Source brackets declare their stroke before painting.");
        return (paths, Number(width, "w"), Number(fill, "h"));
    }

    private static void AssertPathsEqual(double[][] expected, double[][] actual)
    {
        TestAssert.Equal(expected.Length, actual.Length);
        for (int path = 0; path < expected.Length; path++)
            for (int axis = 0; axis < expected[path].Length; axis++)
                TestAssert.True(Math.Abs(expected[path][axis] - actual[path][axis]) < .002d,
                    "Non-rendering metadata, run splits and authored spacing preserve the source-owned bracket shape.");
    }

    private static void AssertContent(PdfPage page, JsonElement source)
    {
        string text = Regex.Replace(string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).Select(show => show.Text)), @"\s+", "");
        foreach (var expected in source.GetProperty("ExpectedText").EnumerateArray())
            TestAssert.Contains(Regex.Replace(expected.GetString()!, @"\s+", ""), text);
    }

    private static PdfPage Render(DocxDocument document, IFontResolver? resolver, CancellationToken token = default) =>
        new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .RenderPagesAsync(document, null, token).ToBlockingEnumerable().Single();

    private static DocxDocument Read(JsonElement source, CancellationToken token)
    {
        string path = Path.GetFullPath(Path.Combine(CasesDirectory, "..", Fixture(source) + ".docx"));
        using var stream = File.OpenRead(path);
        TestAssert.Equal(source.GetProperty("Sha256").GetString(), Convert.ToHexString(SHA256.HashData(stream)));
        stream.Position = 0;
        using var package = OoxPackage.Open(stream, token);
        return new DocxReader().Read(package, null, token, markupMode: OoxPdfDocxMarkupMode.AllMarkup);
    }

    private static string Fixture(JsonElement source)
    {
        const string prefix = "tests/Lokad.OoxPdf.Tests/Cases/";
        string input = source.GetProperty("Input").GetString()!;
        TestAssert.True(input.StartsWith(prefix, StringComparison.Ordinal), "Comment frame inputs remain public.");
        return input[prefix.Length..^5];
    }

    private static double Number(Match match, string name) => double.Parse(match.Groups[name].Value, CultureInfo.InvariantCulture);
    private static string CasesDirectory => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CommentGlyphFrames"));
    private static JsonDocument ReadJson(string name) => JsonDocument.Parse(File.ReadAllText(Path.Combine(CasesDirectory, name)));
}
