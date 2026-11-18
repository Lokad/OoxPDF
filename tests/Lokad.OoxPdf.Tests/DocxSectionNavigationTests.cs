using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxSectionNavigationTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public static void DocxSectionNavigationKeepsIndependentWordControlContent()
    {
        foreach (var mode in Modes)
        foreach (string id in new[] { "single-title", "two-title", "two-default", "bookmark-before", "bookmark-after", "bookmark-empty" })
        {
            var document = DocxTests.ReadDocx(Case(id), mode);
            PdfPage[] pages = Renderer(mode).RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(id == "single-title" ? 1 : 2, pages.Length);
            if (id.StartsWith("bookmark-", StringComparison.Ordinal))
            {
                var link = pages[0].Annotations.Single();
                TestAssert.True(link.Destination is { PageIndex: 1 }, "All three independently exported bookmark controls must retain their incoming GoTo.");
                TestAssert.Contains("TARGET PARAGRAPH", Text(pages[1]));
            }
            else
            {
                string header = id == "two-default" ? "SECOND SECTION DEFAULT" : "SECOND SECTION FIRST";
                TestAssert.Contains(header, Text(pages[^1]));
                if (pages.Length == 2) TestAssert.Contains("FIRST SECTION DEFAULT", Text(pages[0]));
                TestAssert.Contains("SECTION TWO", Text(pages[^1]));
            }
        }
    }

    public static void DocxSectionNavigationSelectsFirstThenDefaultAndEvenStories()
    {
        foreach (var mode in Modes)
        WithMutation("two-title", (parts, xml) => {
            XElement body = xml.Root!.Element(W+"body")!;
            XElement final = body.Element(W+"sectPr")!;
            final.AddBeforeSelf(Break(), Paragraph("CONTINUATION"), Break(), Paragraph("EVEN CONTINUATION"));
            AddFooter(parts, final, "default", "FOOTER DEFAULT");
            AddFooter(parts, final, "first", "FOOTER FIRST");
            var settings = XDocument.Parse(Encoding.UTF8.GetString(parts["word/settings.xml"]));
            settings.Root!.Elements(W+"evenAndOddHeaders").Remove();
            settings.Root.Add(new XElement(W+"evenAndOddHeaders"));
            parts["word/settings.xml"] = Encoding.UTF8.GetBytes(settings.ToString());
        }, path => {
            var document = DocxTests.ReadDocx(path, mode);
            var renderer = Renderer(mode);
            PdfPage[] pages = renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(4, pages.Length);
            TestAssert.Contains("SECOND SECTION FIRST", Text(pages[1]));
            TestAssert.Contains("FOOTER FIRST", Text(pages[1]));
            TestAssert.Contains("SECOND SECTION DEFAULT", Text(pages[2]));
            TestAssert.Contains("FOOTER DEFAULT", Text(pages[2]));
            TestAssert.True(!Text(pages[3]).Contains("SECOND SECTION") && !Text(pages[3]).Contains("FOOTER"),
                "An undefined even variant remains empty; the first variant applies only at the owning section's first page.");
        });
    }

    public static void DocxSectionNavigationDoesNotBackfillAnUndefinedFirstHeader()
    {
        foreach (var mode in Modes)
        WithMutation("two-title", (_, xml) => {
            xml.Root!.Element(W+"body")!.Element(W+"sectPr")!.Elements(W+"headerReference")
                .Where(e => (string?)e.Attribute(W+"type") == "first").Remove();
        }, path => {
            PdfPage[] pages = Renderer(mode).RenderPagesAsync(DocxTests.ReadDocx(path, mode), null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(2, pages.Length);
            TestAssert.Contains("FIRST SECTION DEFAULT", Text(pages[0]));
            TestAssert.True(!Text(pages[1]).Contains("SECOND SECTION"), "An undefined first variant must not select the default header.");
            TestAssert.Contains("SECTION TWO", Text(pages[1]));
        });
    }

    public static void DocxParagraphEndBookmarksFollowTheFinalWrappedPage()
    {
        foreach (var mode in Modes)
        WithMutation("bookmark-after", (_, xml) => {
            XElement body = xml.Root!.Element(W+"body")!;
            body.Elements(W+"p").Last().Descendants(W+"t").Single().Value = string.Join(" ", Enumerable.Repeat("TARGET PARAGRAPH", 90));
            XElement section = body.Element(W+"sectPr")!;
            section.Element(W+"pgSz")!.SetAttributeValue(W+"h", "6000");
            section.Element(W+"pgSz")!.SetAttributeValue(W+"w", "6000");
        }, path => CheckFinalTextDestination(path, mode, expectMultiplePages: true));
    }

    public static void DocxParagraphEndBookmarksIgnoreHiddenTrailingText()
    {
        foreach (var mode in Modes)
        WithMutation("bookmark-after", (_, xml) => {
            var anchor = xml.Descendants(W+"bookmarkStart").Single();
            anchor.AddBeforeSelf(new XElement(W+"r", new XElement(W+"rPr", new XElement(W+"vanish")), new XElement(W+"t", "HIDDEN TRAILING TEXT")));
        }, path => CheckFinalTextDestination(path, mode, expectMultiplePages: false));
    }

    public static void DocxParagraphEndBookmarksResolveUnpaintedTrailingWhitespace()
    {
        foreach (var mode in Modes)
        foreach (bool separateRun in new[] { false, true })
        foreach (bool wrapped in new[] { false, true })
        WithMutation("bookmark-after", (_, xml) => {
            var anchor = xml.Descendants(W+"bookmarkStart").Single();
            var text = anchor.Parent!.Descendants(W+"t").Single();
            if (wrapped)
            {
                text.Value = string.Join(" ", Enumerable.Repeat("TARGET PARAGRAPH", 90));
                var size = xml.Descendants(W+"sectPr").Single().Element(W+"pgSz")!;
                size.SetAttributeValue(W+"h", "6000");
                size.SetAttributeValue(W+"w", "6000");
            }
            if (separateRun)
                anchor.AddBeforeSelf(new XElement(W+"r", new XElement(W+"t",
                    new XAttribute(XNamespace.Xml+"space", "preserve"), "   ")));
            else
            {
                text.SetAttributeValue(XNamespace.Xml+"space", "preserve");
                text.Value += "   ";
            }
        }, path => CheckFinalTextDestination(path, mode, expectMultiplePages: wrapped, exactEndpoint: false));
    }

    public static void DocxParagraphEndBookmarksPreserveMissingTargetFallback()
    {
        WithMutation("bookmark-after", (_, xml) => xml.Descendants(W+"bookmarkStart").Remove(), path => {
            PdfPage[] pages = Renderer(OoxPdfDocxMarkupMode.Final).RenderPagesAsync(DocxTests.ReadDocx(path, OoxPdfDocxMarkupMode.Final), null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(0, pages.Sum(page => page.Annotations.Count));
            TestAssert.Contains("TARGET PARAGRAPH", Text(pages[1]));
        });
    }

    public static void DocxBodyBookmarksFollowTheirPlacedParagraph()
    {
        foreach (var mode in Modes)
        {
            var control = RenderWithBookmarkOutsideParagraph(false, mode);
            var candidate = RenderWithBookmarkOutsideParagraph(true, mode);
            TestAssert.Equal(2, candidate.Length);
            TestAssert.Contains("TARGET PARAGRAPH", Text(candidate[1]));
            TestAssert.True(control.Select(Text).SequenceEqual(candidate.Select(Text)),
                "Moving the markers preserves the text on each page.");
            var destination = candidate[0].Annotations.Single().Destination;
            TestAssert.True(destination is { PageIndex: 1 }, "The body bookmark follows the target across pageBreakBefore.");
            TestAssert.Equal(control[0].Annotations.Single().Destination, destination);
        }

        static PdfPage[] RenderWithBookmarkOutsideParagraph(bool outside, OoxPdfDocxMarkupMode mode)
        {
            PdfPage[]? pages = null;
            WithMutation("bookmark-before", (_, xml) => {
                (XElement start, XElement end, XElement target) = BookmarkSource(xml);
                target.ElementsBeforeSelf(W + "p").Last().Remove();
                target.Element(W + "pPr")?.Remove();
                target.AddFirst(new XElement(W + "pPr", new XElement(W + "pageBreakBefore")));
                if (outside) MoveBookmarkOutside(start, end, target);
            }, path => pages = Render(path, mode));
            return pages!;
        }
    }

    public static void DocxBlockBookmarksFollowPlacedCellAndStoryParagraphs()
    {
        foreach (var mode in Modes)
        foreach (string container in new[] { "nested-cell", "header", "footnote" })
        {
            PdfPage[] inline = RenderTarget(container, blockMarkers: false, mode);
            PdfPage[] block = RenderTarget(container, blockMarkers: true, mode);
            TestAssert.True(inline.Select(Text).SequenceEqual(block.Select(Text)),
                "The same paragraph content stays on the same pages.");
            var destination = block[0].Annotations.Single().Destination!.Value;
            TestAssert.Equal(inline[0].Annotations.Single().Destination, block[0].Annotations.Single().Destination);
            TestAssert.Contains("TARGET PARAGRAPH", Text(block[destination.PageIndex]));
        }
    }

    public static void DocxBlockBookmarksFollowPaintedTableRowsAcrossStories()
    {
        foreach (var mode in Modes)
        foreach (string container in new[] { "nested-cell", "header", "footer", "footnote", "endnote" })
        {
            PdfPage[] inline = RenderTarget(container, blockMarkers: false, mode, tableTarget: true);
            PdfPage[] block = RenderTarget(container, blockMarkers: true, mode, tableTarget: true);
            TestAssert.True(inline.Select(page => page.Content).SequenceEqual(block.Select(page => page.Content)),
                $"Moving the {container} table bookmark must retain the same page paint in {mode}.");
            var links = block[0].Annotations.Where(annotation => annotation.IsDestination).ToArray();
            TestAssert.Equal(1, links.Length);
            PdfLinkDestination destination = links[0].Destination!.Value;
            PdfPage targetPage = block[destination.PageIndex];
            TestAssert.Contains("TARGET PARAGRAPH", Text(targetPage));
            var fill = DocxParagraphShadingTests.ReadRectangles(targetPage.Content).Single(rect => rect.Color == "CC9944");
            TestAssert.True(Math.Abs(destination.Left!.Value - fill.X) < .001d &&
                Math.Abs(destination.Top!.Value - fill.Y - fill.Height) < .001d,
                $"The {container} destination tracks the painted upper-left row corner in {mode}.");
            if (mode == OoxPdfDocxMarkupMode.AllMarkup)
                TestAssert.True(block.Select(Text).Any(text => text.Contains("REVIEW COMMENT", StringComparison.Ordinal)),
                    "The Review control must exercise a printed balloon.");
        }
    }

    public static void DocxFloatingBoxTableRangeBookmarksFollowPaintedRows() =>
        CheckFloatingBoxTableRangeBookmarks(["body", "header", "footer"]);

    public static void DocxFloatingNoteBoxesKeepInkAndRangeDestinationsInPrintedFrame() =>
        CheckFloatingBoxTableRangeBookmarks(["footnote", "endnote"]);

    public static void DocxParagraphRelativeStaticBoxesFollowTheirSelectedStory()
    {
        string fixture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
            "TableGridOrigins", "Containers", "floating-table-ordered.docx"));
        var resolver = new PortableRenderingTests.PinnedResolver();
        XNamespace wp = OoxNamespaces.WordprocessingDrawingNamespace;
        foreach (string owner in new[] { "header", "footer" })
        foreach (var mode in Modes)
        {
            int offset = owner == "header" ? 15 : -62;
            string context = owner + "/" + mode;
            WithDocumentMutation(fixture, (parts, xml) =>
            {
                var body = xml.Descendants(W + "body").Single();
                var final = body.Element(W + "sectPr")!;
                var first = new XElement(final);
                first.AddFirst(new XElement(W + "type", new XAttribute(W + "val", "nextPage")));
                var anchor = body.Descendants(wp + "anchor").Single();
                anchor.Element(wp + "positionV")!.Element(wp + "posOffset")!.SetValue(offset * 12700);
                var box = anchor.Descendants(W + "txbxContent").Single();
                MoveBookmarkOutside(box.Descendants(W + "bookmarkStart").Single(),
                    box.Descendants(W + "bookmarkEnd").Single(), box.Element(W + "tbl")!);
                var drawingRun = anchor.Ancestors(W + "r").First();
                var nestedDrawing = new XElement(drawingRun);
                nestedDrawing.Descendants(W + "bookmarkStart").Remove();
                nestedDrawing.Descendants(W + "bookmarkEnd").Remove();
                nestedDrawing.Descendants(wp + "posOffset").Last().SetValue(100 * 12700);
                foreach (var text in nestedDrawing.Descendants(W + "t")) text.Value = "NESTED DRAWING LABEL";
                box.Add(new XElement(W + "p", nestedDrawing));
                var paragraph = Paragraph("STATIC DRAWING OWNER");
                paragraph.AddFirst(new XElement(W + "pPr",
                    new XElement(W + "spacing", new XAttribute(W + "before", "0"), new XAttribute(W + "after", "0")),
                    new XElement(W + "shd", new XAttribute(W + "fill", "DDEECC"))));
                paragraph.Add(new XElement(drawingRun));
                drawingRun.Ancestors(W + "p").First().Remove();

                XName root = W + (owner == "header" ? "hdr" : "ftr");
                string firstPart = "first-drawing-" + owner;
                string followingPart = "following-" + owner;
                first.AddFirst(new XElement(W + (owner + "Reference"), new XAttribute(W + "type", "first"),
                    new XAttribute(R + "id", firstPart)));
                first.Add(new XElement(W + "titlePg"));
                AddStoryPart(parts, firstPart, owner, new XElement(root, Paragraph("STATIC INTRO"), paragraph));
                final.AddFirst(new XElement(W + (owner + "Reference"), new XAttribute(W + "type", "default"),
                    new XAttribute(R + "id", followingPart)));
                AddStoryPart(parts, followingPart, owner, new XElement(root, Paragraph("FOLLOWING STATIC STORY")));
                final.AddBeforeSelf(Break(), Paragraph("FIRST SECTION CONTINUES"),
                    new XElement(W + "p", new XElement(W + "pPr", first)), Paragraph("FOLLOWING SECTION"));
            }, path =>
            {
                var geometry = mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                    : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
                var document = DocxTests.ReadDocx(path, mode);
                var renderer = new DocxRenderer(resolver, mode, geometry);
                PdfPage[] pages = renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
                TestAssert.Equal(3, pages.Length);
                DocxInlineCompositionTests.AssertSourceLabelsPaint(pages[0], mode,
                    ["BOX TABLE NORTH", "BOX TABLE SOUTH"], context);
                TestAssert.DoesNotContain("NESTED DRAWING LABEL", Text(pages[0]));
                TestAssert.Equal(1, pages[0].Annotations.Count(link => link.Destination is { PageIndex: 0 }));
                TestAssert.DoesNotContain("BOX TABLE NORTH", Text(pages[1]));
                TestAssert.DoesNotContain("BOX TABLE NORTH", Text(pages[2]));
                TestAssert.Contains("FIRST SECTION CONTINUES", Text(pages[1]));
                TestAssert.Contains("FOLLOWING STATIC STORY", Text(pages[2]));

                double scale = DocxRenderer.ResolveWordCompatiblePrintScale(document, DocxMarkupContext.FromMode(mode, geometry));
                var outline = DocxInlineCompositionTests.ReadFloatingOutline(pages[0], 60d * scale, 216d * scale, .02d).Single().Bounds;
                var fill = DocxParagraphShadingTests.ReadRectangles(pages[0].Content).Single(rect => rect.Color == "DDEECC");
                TestAssert.True(Math.Abs(outline[1] + outline[3] - fill.Y - fill.Height + offset * scale) < .003d,
                    "The floating outline must keep its authored offset from the printed owner: " + context);
            });
        }
    }

    public static void DocxStaticCellFloatingBoxesFollowTheirParagraphAndClip()
    {
        string fixture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
            "TableGridOrigins", "Containers", "floating-table-ordered.docx"));
        var resolver = new PortableRenderingTests.PinnedResolver();
        XNamespace wp = OoxNamespaces.WordprocessingDrawingNamespace;
        foreach (var (story, height, exact, wrap) in new[]
        {
            ("header", 0, false, false), ("footer", 0, false, false),
            ("header", 2400, false, true), ("header", 300, true, true)
        })
        foreach (var mode in Modes)
        {
            int offset = story == "footer" ? -62 : 30;
            string context = $"{story}/{height}/{mode}";
            WithDocumentMutation(fixture, (parts, xml) =>
            {
                var body = xml.Descendants(W + "body").Single();
                var anchor = body.Descendants(wp + "anchor").Single();
                anchor.Element(wp + "positionV")!.Element(wp + "posOffset")!.SetValue(offset * 12700);
                if (wrap) anchor.Element(wp + "wrapNone")!.ReplaceWith(new XElement(wp + "wrapSquare", new XAttribute("wrapText", "bothSides")));
                var box = anchor.Descendants(W + "txbxContent").Single();
                var boxTable = box.Element(W + "tbl")!;
                MoveBookmarkOutside(box.Descendants(W + "bookmarkStart").Single(),
                    box.Descendants(W + "bookmarkEnd").Single(), boxTable);
                var target = boxTable.Element(W + "tr")!.Element(W + "tc")!;
                var targetProperties = target.Element(W + "tcPr");
                if (targetProperties is null) target.AddFirst(targetProperties = new XElement(W + "tcPr"));
                targetProperties.Element(W + "shd")?.Remove();
                targetProperties.Add(new XElement(W + "shd", new XAttribute(W + "fill", "CC9944")));

                var paragraph = Paragraph("STATIC CELL OWNER");
                paragraph.AddFirst(new XElement(W + "pPr",
                    new XElement(W + "spacing", new XAttribute(W + "before", "0"), new XAttribute(W + "after", "0"))));
                paragraph.Add(new XElement(anchor.Ancestors(W + "r").First()));
                anchor.Ancestors(W + "p").First().Remove();
                var cell = new XElement(W + "tc", new XElement(W + "tcPr",
                    new XElement(W + "tcW", new XAttribute(W + "w", "6000"), new XAttribute(W + "type", "dxa"))), paragraph);
                if (!exact) paragraph.AddBeforeSelf(Paragraph("EARLIER CELL PARAGRAPH"));
                var row = new XElement(W + "tr", cell);
                if (height != 0) row.AddFirst(new XElement(W + "trPr", new XElement(W + "trHeight",
                    new XAttribute(W + "val", height), new XAttribute(W + "hRule", exact ? "exact" : "atLeast"))));
                var table = new XElement(W + "tbl", new XElement(W + "tblPr"),
                    new XElement(W + "tblGrid", new XElement(W + "gridCol", new XAttribute(W + "w", "6000"))), row);
                string name = "cell-anchor-" + story;
                body.Element(W + "sectPr")!.AddFirst(new XElement(W + (story + "Reference"),
                    new XAttribute(W + "type", "default"), new XAttribute(R + "id", name)));
                AddStoryPart(parts, name, story, new XElement(W + (story == "header" ? "hdr" : "ftr"), table));
            }, path =>
            {
                var geometry = mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                    : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
                var document = DocxTests.ReadDocx(path, mode);
                var page = new DocxRenderer(resolver, mode, geometry)
                    .RenderPagesAsync(document, null, default).ToBlockingEnumerable().Single();
                if (!exact)
                {
                    DocxInlineCompositionTests.AssertSourceLabelsPaint(page, mode,
                        ["BOX TABLE NORTH", "BOX TABLE SOUTH", "STATIC CELL OWNER"], context);
                    var target = DocxParagraphShadingTests.ReadRectangles(page.Content).Single(rect => rect.Color == "CC9944");
                    TestAssert.True(page.Annotations.Any(link => link.Destination is { PageIndex: 0, Left: double x, Top: double y } &&
                        Math.Abs(x - target.X) < .002d && Math.Abs(y - target.Y - target.Height) < .002d),
                        "The link reaches the painted row of the cell-owned box: " + context);
                }
                else
                {
                    TestAssert.Equal(0, page.Annotations.Count(link => link.Destination is not null));
                    double scale = DocxRenderer.ResolveWordCompatiblePrintScale(document, DocxMarkupContext.FromMode(mode, geometry));
                    var (bounds, clip) = DocxInlineCompositionTests.ReadFloatingOutline(page, 60d * scale, 216d * scale, .02d).Single();
                    TestAssert.True(clip.Height <= 15d * scale + .01d && bounds[1] + bounds[3] < clip.Y,
                        "The exact-height cell clips the entire drawing below its viewport: " + context);
                }
            });
        }
    }

    private static void CheckFloatingBoxTableRangeBookmarks(string[] owners)
    {
        string fixture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
            "TableGridOrigins", "Containers", "floating-table-ordered.docx"));
        var resolver = new PortableRenderingTests.PinnedResolver();
        XNamespace wp = OoxNamespaces.WordprocessingDrawingNamespace;
        foreach (string owner in owners)
        foreach (var mode in Modes)
        {
            string context = owner + "/" + mode;
            WithDocumentMutation(fixture, (parts, xml) =>
            {
                var body = xml.Descendants(W + "body").Single();
                var box = body.Descendants(W + "txbxContent").Single();
                var table = box.Element(W + "tbl")!;
                MoveBookmarkOutside(box.Descendants(W + "bookmarkStart").Single(),
                    box.Descendants(W + "bookmarkEnd").Single(), table);
                var firstCell = table.Element(W + "tr")!.Element(W + "tc")!;
                if (firstCell.Element(W + "tcPr") is null) firstCell.AddFirst(new XElement(W + "tcPr"));
                var properties = firstCell.Element(W + "tcPr")!;
                properties.Element(W + "shd")?.Remove();
                properties.Add(new XElement(W + "shd", new XAttribute(W + "fill", "CC9944")));
                if (owner == "body") return;

                var anchor = box.Ancestors(wp + "anchor").Single();
                var paragraph = anchor.Ancestors(W + "p").First();
                string name = "bookmark-floating-" + owner;
                if (owner is "header" or "footer")
                {
                    var position = anchor.Element(wp + "positionV")!;
                    position.SetAttributeValue("relativeFrom", "page");
                    position.Element(wp + "posOffset")!.SetValue((owner == "header" ? 30 : 700) * 12700);
                    body.Element(W + "sectPr")!.AddFirst(new XElement(W + (owner + "Reference"),
                        new XAttribute(W + "type", "default"), new XAttribute(R + "id", name)));
                    paragraph.Remove();
                    AddStoryPart(parts, name, owner, new XElement(W + (owner == "header" ? "hdr" : "ftr"), paragraph));
                }
                else
                {
                    paragraph.AddBeforeSelf(new XElement(W + "p", new XElement(W + "r",
                        new XElement(W + (owner + "Reference"), new XAttribute(W + "id", "1")))));
                    paragraph.Remove();
                    AddStoryPart(parts, name, owner + "s", new XElement(W + (owner + "s"),
                        new XElement(W + owner, new XAttribute(W + "id", "1"), paragraph,
                            Enumerable.Range(1, 11).Select(index => Paragraph("Additional related story line " + index)))));
                }
            }, path =>
            {
                var renderer = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final
                    ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                var page = renderer.RenderPagesAsync(DocxTests.ReadDocx(path, mode), null, default).ToBlockingEnumerable().Single();
                DocxInlineCompositionTests.AssertSourceLabelsPaint(page, mode, ["BOX TABLE NORTH", "BOX TABLE SOUTH"], context);
                var fill = DocxParagraphShadingTests.ReadRectangles(page.Content).Single(rect => rect.Color == "CC9944");
                TestAssert.True(page.Annotations.Any(link => link.Destination is { PageIndex: 0, Left: double x, Top: double y } &&
                    Math.Abs(x - fill.X) < .001d && Math.Abs(y - fill.Y - fill.Height) < .001d),
                    "The range link must reach the painted first-row corner: " + context);
            });
        }
    }

    private static PdfPage[] RenderTarget(string container, bool blockMarkers, OoxPdfDocxMarkupMode mode, bool tableTarget = false)
    {
        PdfPage[] pages = [];
        WithMutation("bookmark-before", (parts, xml) =>
        {
            XElement target = BookmarkSource(xml).Target;
            XElement section = target.ElementsAfterSelf(W + "sectPr").Single();
            if (tableTarget) AddReviewComment(parts, target.Parent!.Elements(W + "p").First());
            var content = new XElement(target);
            var start = content.Element(W + "bookmarkStart")!;
            var end = content.Element(W + "bookmarkEnd")!;
            if (tableTarget)
            {
                content = Table(content);
                content.Descendants(W + "tc").Single().AddFirst(new XElement(W + "tcPr",
                    new XElement(W + "shd", new XAttribute(W + "fill", "CC9944"))));
            }
            switch (container)
            {
                case "nested-cell":
                    target.ReplaceWith(Table(tableTarget ? content : Table(content), Paragraph(string.Empty)));
                    break;
                case "header":
                case "footer":
                {
                    var part = new XElement(W + (container == "header" ? "hdr" : "ftr"), content);
                    section.Elements(W + (container + "Reference")).Remove();
                    section.Element(W + "titlePg")?.Remove();
                    string name = "bookmark-" + container;
                    section.AddFirst(new XElement(W + (container + "Reference"),
                        new XAttribute(W + "type", "default"), new XAttribute(R + "id", name)));
                    target.ReplaceWith(Paragraph("SECOND PAGE"));
                    if (blockMarkers) MoveBookmarkOutside(start, end, content);
                    AddStoryPart(parts, name, container, part);
                    return;
                }
                case "footnote":
                case "endnote":
                {
                    target.ReplaceWith(new XElement(W + "p", new XElement(W + "r",
                        new XElement(W + (container + "Reference"), new XAttribute(W + "id", "1")))));
                    var note = new XElement(W + container, new XAttribute(W + "id", "1"), content);
                    if (blockMarkers) MoveBookmarkOutside(start, end, content);
                    AddStoryPart(parts, "bookmark-" + container, container + "s", new XElement(W + (container + "s"), note));
                    return;
                }
            }

            if (blockMarkers) MoveBookmarkOutside(start, end, content);
        }, path => pages = Render(path, mode));
        return pages;

        static XElement Table(params XElement[] content) => new(W + "tbl",
            new XElement(W + "tblPr"),
            new XElement(W + "tblGrid", new XElement(W + "gridCol", new XAttribute(W + "w", "4800"))),
            new XElement(W + "tr", new XElement(W + "tc", content)));

        static void AddReviewComment(Dictionary<string, byte[]> parts, XElement first)
        {
            XElement anchor = Paragraph("COMMENT ANCHOR");
            anchor.AddFirst(new XElement(W + "commentRangeStart", new XAttribute(W + "id", "7")));
            anchor.Add(new XElement(W + "commentRangeEnd", new XAttribute(W + "id", "7")),
                new XElement(W + "r", new XElement(W + "commentReference", new XAttribute(W + "id", "7"))));
            first.AddAfterSelf(anchor);
            AddStoryPart(parts, "bookmark-comment", "comments", new XElement(W + "comments",
                new XElement(W + "comment", new XAttribute(W + "id", "7"), new XAttribute(W + "author", "Public"),
                    Paragraph("REVIEW COMMENT"))));
        }
    }

    public static void DocxBodyBookmarksRespectClosedRevisionAndTableBoundaries()
    {
        WithMutation("bookmark-before", (_, xml) => {
            (XElement start, XElement end, XElement target) = BookmarkSource(xml);
            MoveBookmarkOutside(start, end, target);
            end.SetAttributeValue(W + "id", "00");
            var insertion = new XElement(W + "ins", new XAttribute(W + "id", "90"), new XAttribute(W + "author", "Public"));
            target.AddBeforeSelf(insertion);
            target.Remove();
            insertion.Add(target);
            end.AddAfterSelf(Paragraph("FOLLOWING PARAGRAPH"));
        }, path => {
            var final = Render(path, OoxPdfDocxMarkupMode.Final);
            TestAssert.Equal(2, final.Length);
            TestAssert.True(final[0].Annotations.Single().Destination is { PageIndex: 1 },
                "Final resolves the bookmark on the visible inserted paragraph.");
            var original = Render(path, OoxPdfDocxMarkupMode.Original);
            TestAssert.Equal(2, original.Length);
            TestAssert.Contains("FOLLOWING PARAGRAPH", Text(original[1]));
            TestAssert.DoesNotContain("TARGET PARAGRAPH", Text(original[1]));
            TestAssert.Equal(0, original[0].Annotations.Count);
        });

        foreach (var mode in Modes)
        WithMutation("bookmark-before", (_, xml) => {
            (XElement start, XElement end, XElement target) = BookmarkSource(xml);
            target.ElementsBeforeSelf(W + "p").Last().Remove();
            MoveBookmarkOutside(start, end, target);
            start.AddBeforeSelf(Table("FIRST ROW"));
            start.AddAfterSelf(Table("SECOND ROW"));
            end.Remove();
            target.AddBeforeSelf(end);
            XElement section = xml.Descendants(W + "sectPr").Single();
            section.Element(W + "pgSz")!.SetAttributeValue(W + "w", "10000");
            section.Element(W + "pgSz")!.SetAttributeValue(W + "h", "8000");
            foreach (string side in new[] { "top", "bottom", "left", "right" })
                section.Element(W + "pgMar")!.SetAttributeValue(W + side, "800");
        }, path => {
            var document = DocxTests.ReadDocx(path, mode);
            TestAssert.Equal(2, document.Tables.Single().Rows.Count);
            var pages = Renderer(mode).RenderPagesAsync(document, null, default)
                .ToBlockingEnumerable().ToArray();
            TestAssert.Equal(2, pages.Length);
            TestAssert.Contains("FIRST ROW", Text(pages[0]));
            TestAssert.DoesNotContain("SECOND ROW", Text(pages[0]));
            TestAssert.Contains("SECOND ROW", Text(pages[1]));
            var destination = pages[0].Annotations.Single().Destination;
            TestAssert.True(destination is { PageIndex: 1, Left: > 0d, Top: > 0d } && destination.Value.Top <= pages[1].Height,
                "The joined result retains the start of the second authored table on its own page.");
        });

        static XElement Table(string caption) => new(W + "tbl",
            new XElement(W + "tblPr", new XElement(W + "tblLayout", new XAttribute(W + "type", "fixed"))),
            new XElement(W + "tblGrid", new XElement(W + "gridCol", new XAttribute(W + "w", "3000"))),
            new XElement(W + "tr", new XElement(W + "trPr", new XElement(W + "cantSplit"),
                new XElement(W + "trHeight", new XAttribute(W + "val", "4400"), new XAttribute(W + "hRule", "exact"))),
                new XElement(W + "tc", Paragraph(caption))));
    }

    private static (XElement Start, XElement End, XElement Target) BookmarkSource(XDocument xml)
    {
        XElement start = xml.Descendants(W + "bookmarkStart").Single();
        return (start, xml.Descendants(W + "bookmarkEnd").Single(), start.Parent!);
    }

    private static void MoveBookmarkOutside(XElement start, XElement end, XElement target)
    {
        start.Remove();
        end.Remove();
        target.AddBeforeSelf(start);
        target.AddAfterSelf(end);
    }

    private static PdfPage[] Render(string path, OoxPdfDocxMarkupMode mode) =>
        Renderer(mode).RenderPagesAsync(DocxTests.ReadDocx(path, mode), null, default).ToBlockingEnumerable().ToArray();

    private static void CheckFinalTextDestination(string path, OoxPdfDocxMarkupMode mode, bool expectMultiplePages,
        bool exactEndpoint = true)
    {
        var document = DocxTests.ReadDocx(path, mode);
        var renderer = Renderer(mode);
        PdfPage[] pages = renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
        if (expectMultiplePages) TestAssert.True(pages.Length > 2, "The target must wrap across pages.");
        PdfLinkDestination destination = pages[0].Annotations.Single().Destination!.Value;
        TestAssert.Equal(pages.Length-1, destination.PageIndex);
        var lastLine = renderer.InspectTextEmission(document).Lines.Last(line => line.ContainerStoryKind == "Body" &&
            line.Segments.Any(segment => !segment.IsTerminalLineSpace && segment.TextLength > 0));
        var lastSegment = lastLine.Segments.Last(segment => !segment.IsTerminalLineSpace && segment.TextLength > 0);
        if (exactEndpoint)
            TestAssert.True(Math.Abs(destination.Left!.Value - (lastSegment.X+lastSegment.Width-3d*lastSegment.FontSize/14d)) < .02d,
                "The end bookmark must use the final emitted endpoint and existing viewport context.");
        else
        {
            // Authored whitespace can end before layout-owned terminal spaces.
            // These transformed inputs qualify the final painted line; the
            // original controls above retain their exact endpoint assertions.
            double context = 3d*lastSegment.FontSize/14d;
            double left = lastLine.Segments.Where(segment => !segment.IsTerminalLineSpace && segment.TextLength > 0)
                .Min(segment => segment.X);
            TestAssert.True(destination.Left >= left-context && destination.Left <= lastSegment.X+lastSegment.Width-context,
                "The whitespace boundary must retain a destination within its final painted source line.");
        }
        TestAssert.True(destination.Top > lastSegment.BaselineY && destination.Top < lastSegment.BaselineY+lastSegment.PdfFontSize*2d,
            "The destination must belong to the final line, not an earlier wrapped fragment.");
    }

    private static readonly OoxPdfDocxMarkupMode[] Modes = [OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup];
    private static string Case(string id) => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "CompositionProbes", "section-nav-"+id+".docx"));
    private static string Text(PdfPage page) => DocxPageFootnoteNumberingTests.DecodeText(page);
    private static DocxRenderer Renderer(OoxPdfDocxMarkupMode mode) => new(new TestFaceFontResolver(), mode,
        mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
    private static XElement Paragraph(string text) => new(W+"p", new XElement(W+"r", new XElement(W+"t", text)));
    private static XElement Break() => new(W+"p", new XElement(W+"r", new XElement(W+"br", new XAttribute(W+"type", "page"))));

    private static void AddFooter(Dictionary<string,byte[]> parts, XElement section, string type, string text)
    {
        string name = "footer-"+type;
        section.AddFirst(new XElement(W+"footerReference", new XAttribute(W+"type", type), new XAttribute(R+"id", name)));
        AddStoryPart(parts, name, "footer", new XElement(W+"ftr", Paragraph(text)));
    }

    private static void AddStoryPart(Dictionary<string, byte[]> parts, string name, string kind, XElement story)
    {
        parts["word/"+name+".xml"] = Encoding.UTF8.GetBytes(story.ToString());
        var rels = XDocument.Parse(Encoding.UTF8.GetString(parts["word/_rels/document.xml.rels"]));
        rels.Root!.Add(new XElement(rels.Root.Name.Namespace+"Relationship", new XAttribute("Id", name), new XAttribute("Target", name+".xml"), new XAttribute("Type", R.NamespaceName+"/"+kind)));
        parts["word/_rels/document.xml.rels"] = Encoding.UTF8.GetBytes(rels.ToString());
        var types = XDocument.Parse(Encoding.UTF8.GetString(parts["[Content_Types].xml"]));
        types.Root!.Add(new XElement(types.Root.Name.Namespace+"Override", new XAttribute("PartName", "/word/"+name+".xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml."+kind+"+xml")));
        parts["[Content_Types].xml"] = Encoding.UTF8.GetBytes(types.ToString());
    }

    private static void WithMutation(string id, Action<Dictionary<string,byte[]>,XDocument> mutate, Action<string> check) =>
        WithDocumentMutation(Case(id), mutate, check);

    private static void WithDocumentMutation(string path, Action<Dictionary<string,byte[]>,XDocument> mutate, Action<string> check)
    {
        using var archive = ZipFile.OpenRead(path);
        var parts = archive.Entries.ToDictionary(entry => entry.FullName, entry => {
            using var source = entry.Open(); using var bytes = new MemoryStream(); source.CopyTo(bytes); return bytes.ToArray();
        });
        var xml = XDocument.Parse(Encoding.UTF8.GetString(parts["word/document.xml"]));
        mutate(parts, xml);
        parts["word/document.xml"] = Encoding.UTF8.GetBytes(xml.ToString());
        string mutated = TestFixtures.WriteTempPackage(".docx", parts);
        try { check(mutated); } finally { File.Delete(mutated); }
    }
}
