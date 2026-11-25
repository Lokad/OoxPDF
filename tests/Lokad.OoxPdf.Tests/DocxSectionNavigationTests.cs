using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
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

    public static void DocxIncomingSectionParityKeepsBlankPagesAndPhysicalTargets()
    {
        foreach (var mode in Modes)
        foreach (var (outgoing, incoming, extraBodyPage, pageCount, blankPage) in new[]
        {
            ("nextPage", "oddPage", false, 3, 1),
            ("oddPage", "nextPage", false, 2, -1),
            ("nextPage", "evenPage", true, 4, 2),
            ("evenPage", "nextPage", true, 3, -1),
            ("nextPage", "evenPage", false, 2, -1)
        })
        WithMutation("two-title", (parts, xml) =>
        {
            var body = xml.Root!.Element(W + "body")!;
            var paragraphs = body.Elements(W + "p").ToArray();
            var first = paragraphs[0].Element(W + "pPr")!.Element(W + "sectPr")!;
            var following = body.Element(W + "sectPr")!;
            first.Element(W + "type")!.SetAttributeValue(W + "val", outgoing);
            following.Element(W + "type")!.SetAttributeValue(W + "val", incoming);
            AddFooter(parts, first, "default", "PUBLIC PARITY FOOTER");
            LinkToSection(paragraphs[0], paragraphs[1], "PublicParityTarget");
            if (extraBodyPage)
            {
                first.Remove();
                var extra = Paragraph("PUBLIC EXTRA BODY PAGE");
                extra.AddFirst(new XElement(W + "pPr", first));
                paragraphs[0].AddAfterSelf(Break(), extra);
            }
        }, path =>
        {
            var pages = Renderer(mode).RenderPagesAsync(DocxTests.ReadDocx(path, mode), null, default)
                .ToBlockingEnumerable().ToArray();
            TestAssert.Equal(pageCount, pages.Length);
            TestAssert.Contains("PUBLIC PARITY FOOTER", Text(pages[0]));
            TestAssert.Contains("SECTION TWO", Text(pages[^1]));
            TestAssert.Contains("SECOND SECTION FIRST", Text(pages[^1]));
            TestAssert.Equal(pageCount - 1, pages[0].Annotations.Single().Destination?.PageIndex);
            if (extraBodyPage) TestAssert.Contains("PUBLIC EXTRA BODY PAGE", Text(pages[1]));
            if (blankPage >= 0)
            {
                var blank = pages[blankPage];
                TestAssert.Equal(string.Empty, blank.Content);
                TestAssert.Equal(0, blank.Images.Count);
                TestAssert.Equal(0, blank.Annotations.Count);
            }
        });
    }

    public static void DocxIncomingNextPageKeepsPaperAndNumberingAfterContinuousSection()
    {
        foreach (var mode in Modes)
        foreach (bool explicitType in new[] { true, false })
        WithMutation("two-title", (parts, xml) =>
        {
            var body = xml.Root!.Element(W + "body")!;
            var paragraphs = body.Elements(W + "p").ToArray();
            var first = paragraphs[0].Element(W + "pPr")!.Element(W + "sectPr")!;
            var following = body.Element(W + "sectPr")!;
            first.Element(W + "type")!.SetAttributeValue(W + "val", "continuous");
            if (explicitType)
                following.Element(W + "type")!.SetAttributeValue(W + "val", "nextPage");
            else
                following.Element(W + "type")!.Remove();
            following.Element(W + "pgSz")!.SetAttributeValue(W + "w", "10800");
            following.Add(new XElement(W + "pgNumType", new XAttribute(W + "start", "40")));
            paragraphs[1].Add(RunText(" BODY PAGE "), Field("PAGE"), RunText(" END"));
            LinkToSection(paragraphs[0], paragraphs[1], "PublicNextPageTarget");
            following.AddFirst(new XElement(W + "footerReference", new XAttribute(W + "type", "first"),
                new XAttribute(R + "id", "first-page-counter")));
            AddStoryPart(parts, "first-page-counter", "footer",
                new XElement(W + "ftr", new XElement(W + "p", RunText("FOOTER "), Field("PAGE"),
                    RunText(" OF "), Field("NUMPAGES"))));
        }, path =>
        {
            var pages = Renderer(mode).RenderPagesAsync(DocxTests.ReadDocx(path, mode), null, default)
                .ToBlockingEnumerable().ToArray();
            TestAssert.Equal(2, pages.Length);
            TestAssert.Equal(612d, pages[0].Width);
            TestAssert.Equal(540d, pages[1].Width);
            TestAssert.Contains("FIRST SECTION DEFAULT", Text(pages[0]));
            string following = Text(pages[1]);
            TestAssert.Contains("SECOND SECTION FIRST", following);
            TestAssert.Contains("BODY PAGE 40 END", following);
            TestAssert.Contains("FOOTER 40 OF 2", following);
            TestAssert.Equal(1, pages[0].Annotations.Single().Destination?.PageIndex);
        });

        static XElement RunText(string text) => new(W + "r",
            new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), text));
        static XElement Field(string instruction) =>
            new(W + "fldSimple", new XAttribute(W + "instr", instruction), RunText("900"));
    }

    private static void LinkToSection(XElement from, XElement to, string anchor)
    {
        from.Add(new XElement(W + "hyperlink", new XAttribute(W + "anchor", anchor),
            new XElement(W + "r", new XElement(W + "t", "Follow"))));
        to.AddFirst(new XElement(W + "bookmarkStart", new XAttribute(W + "id", "9"), new XAttribute(W + "name", anchor)));
        to.Add(new XElement(W + "bookmarkEnd", new XAttribute(W + "id", "9")));
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

    public static void DocxFloatingBoxHyperlinksFollowThePrintedClip() => CheckTableAndBoxLinkClips(tableDepth: 0);

    public static void DocxFloatingBoxTableHyperlinksFollowCellAndAncestorClips()
    {
        CheckTableAndBoxLinkClips(tableDepth: 1);
        CheckTableAndBoxLinkClips(tableDepth: 2);
    }

    public static void DocxBodyTableHyperlinksFollowCellAndAncestorClips()
    {
        CheckTableAndBoxLinkClips(tableDepth: 1, "body");
        CheckTableAndBoxLinkClips(tableDepth: 2, "body");
    }

    public static void DocxStoryTableHyperlinksFollowPrintedCellClips()
    {
        foreach (string story in new[] { "header", "footer", "footnote", "endnote" })
            CheckTableAndBoxLinkClips(tableDepth: 1, story);
    }

    public static void DocxInlineBoxTableHyperlinksFollowPrintedCellClips() => CheckTableAndBoxLinkClips(tableDepth: 1, "inline");

    private static void CheckTableAndBoxLinkClips(int tableDepth, string placement = "floating")
    {
        const string prose = "PUBLIC PROSE TARGET";
        const string proseUri = "https://example.invalid/box-link/prose";
        const string mathUri = "https://example.invalid/box-link/math";
        string fixture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
            "TableGridOrigins", "Containers", "floating-table-ordered.docx"));
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        XNamespace wp = OoxNamespaces.WordprocessingDrawingNamespace;
        XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        bool inNote = placement is "footnote" or "endnote";
        foreach (var mode in Modes)
        {
            string context = $"{mode}, table depth {tableDepth}, {placement}";
            double scale = 1d;
            double fullHeight = placement == "floating" ? 100d : 140d;
            PdfPage tall = Render(fullHeight);
            var (baseline, tallClip) = FindProseFrame(tall);
            var fullLink = tall.Annotations.Single(link => link.Uri == proseUri);
            TestAssert.True(tall.Annotations.Count(link => link.Uri == mathUri) == 1, "The fitting math link is visible: " + context);

            // Cut the source clip through the uppercase linked glyphs. The
            // positive case should keep an actionable rectangle inside it.
            double partialHeight = fullHeight - (baseline + 4d * scale - tallClip.Y) / (inNote ? 1d : scale);
            TestAssert.True(partialHeight > 16d && partialHeight < fullHeight, context);
            PdfPage partial = Render(partialHeight);
            var (partialBaseline, partialClip) = FindProseFrame(partial);
            var partialLink = partial.Annotations.Single(link => link.Uri == proseUri);
            TestAssert.True(partialBaseline < partialClip.Y && partialLink.Height < fullLink.Height &&
                Math.Abs(partialLink.Y - partialClip.Y) < .002d &&
                partialLink.Y + partialLink.Height <= partialClip.Y + partialClip.Height + .002d, context);
            TestAssert.Equal(0, partial.Annotations.Count(link => link.Uri == mathUri));

            PdfPage shortBox = Render(16d);
            var (hiddenBaseline, shortClip) = FindProseFrame(shortBox);
            TestAssert.True(shortClip.Y > hiddenBaseline + 12d * scale, context);
            TestAssert.Equal(0, shortBox.Annotations.Count(link => link.Uri is proseUri or mathUri));

            PdfPage Render(double height)
            {
                PdfPage? result = null;
                WithDocumentMutation(fixture, (parts, document) =>
                {
                    var anchor = document.Descendants(wp + "anchor").Single();
                    anchor.Element(wp + "extent")!.SetAttributeValue("cy", (long)Math.Round((tableDepth == 0 ? height : 145d) * 12700));
                    anchor.Element(wp + "positionH")!.Element(wp + "posOffset")!.SetValue(170 * 12700);
                    var fraction = new XElement(m + "oMath", new XElement(m + "f",
                        new XElement(m + "num", Link("rIdBoxMath", new XElement(m + "r", new XElement(m + "t", "2")))),
                        new XElement(m + "den", new XElement(m + "r", new XElement(m + "t", "7")))));
                    XElement[] content =
                    [
                        Paragraph("PUBLIC BOX CAPTION"), Paragraph("PUBLIC BOX FOLLOWING"), Paragraph("PUBLIC SPACER LINE"),
                        new XElement(W + "p", Link("rIdBoxProse", new XElement(W + "r", new XElement(W + "t", prose)))),
                        new XElement(W + "p", fraction)
                    ];
                    if (tableDepth == 2) content = [ExactHeightTable(fullHeight, content), new XElement(W + "p")];
                    if (tableDepth != 0) content = [ExactHeightTable(height, content), Paragraph("PUBLIC AFTER TABLE")];
                    string relsName = PlaceContent(parts, document, anchor, content);
                    XNamespace relNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
                    var rels = parts.TryGetValue(relsName, out var existing)
                        ? XDocument.Parse(Encoding.UTF8.GetString(existing)) : new XDocument(new XElement(relNamespace + "Relationships"));
                    foreach (var (id, uri) in new[] { ("rIdBoxProse", proseUri), ("rIdBoxMath", mathUri) })
                        rels.Root!.Add(new XElement(rels.Root.Name.Namespace + "Relationship", new XAttribute("Id", id),
                            new XAttribute("Type", R.NamespaceName + "/hyperlink"), new XAttribute("Target", uri),
                            new XAttribute("TargetMode", "External")));
                    parts[relsName] = Encoding.UTF8.GetBytes(rels.ToString());
                }, path =>
                {
                    var story = DocxTests.ReadDocx(path, mode);
                    var geometry = mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                        : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
                    scale = DocxRenderer.ResolveWordCompatiblePrintScale(story, DocxMarkupContext.FromMode(mode, geometry));
                    if (mode == OoxPdfDocxMarkupMode.AllMarkup) TestAssert.True(scale < 1d, "Review uses its scaled printed frame.");
                    result = new DocxRenderer(resolver, mode, geometry)
                        .RenderPagesAsync(story, null, default).ToBlockingEnumerable().Single();
                });
                return result ?? throw new InvalidOperationException("The mutated document was not rendered.");
            }
        }

        string PlaceContent(Dictionary<string, byte[]> parts, XDocument document, XElement anchor, XElement[] content)
        {
            var paragraph = anchor.Ancestors(W + "p").First();
            switch (placement)
            {
                case "floating":
                case "inline":
                    if (placement == "inline")
                    {
                        anchor.Name = wp + "inline";
                        anchor.Elements().Where(element => element.Name.LocalName is "simplePos" or "positionH" or "positionV" or "wrapNone").Remove();
                    }
                    anchor.Descendants(W + "txbxContent").Single().ReplaceNodes(content);
                    return "word/_rels/document.xml.rels";
                case "body":
                    paragraph.ReplaceWith(content);
                    return "word/_rels/document.xml.rels";
                default:
                    string name = "link-clip-" + placement;
                    if (inNote)
                    {
                        paragraph.ReplaceWith(new XElement(W + "p", new XElement(W + "r",
                            new XElement(W + (placement + "Reference"), new XAttribute(W + "id", 1)))));
                        AddStoryPart(parts, name, placement + "s", new XElement(W + (placement + "s"),
                            new XElement(W + placement, new XAttribute(W + "id", 1), content)));
                    }
                    else
                    {
                        paragraph.Remove();
                        var section = document.Descendants(W + "sectPr").Single();
                        section.AddFirst(new XElement(W + (placement + "Reference"),
                            new XAttribute(W + "type", "default"), new XAttribute(R + "id", name)));
                        AddStoryPart(parts, name, placement, new XElement(W + (placement == "header" ? "hdr" : "ftr"), content));
                    }
                    return "word/_rels/" + name + ".xml.rels";
            }
        }

        static XElement Link(string id, XElement run) => new(W + "hyperlink", new XAttribute(R + "id", id), run);

        static XElement ExactHeightTable(double height, XElement[] content) => new(W + "tbl",
            new XElement(W + "tblPr", new XElement(W + "tblW", new XAttribute(W + "w", 3000), new XAttribute(W + "type", "dxa"))),
            new XElement(W + "tblGrid", new XElement(W + "gridCol", new XAttribute(W + "w", 3000))),
            new XElement(W + "tr", new XElement(W + "trPr", new XElement(W + "trHeight",
                new XAttribute(W + "val", (long)Math.Round(height * 20)), new XAttribute(W + "hRule", "exact"))),
                new XElement(W + "tc", new XElement(W + "tcPr", new XElement(W + "tcW", new XAttribute(W + "w", 3000),
                    new XAttribute(W + "type", "dxa")), new XElement(W + "tcMar", new[] { "top", "bottom", "left", "right" }
                    .Select(side => new XElement(W + side, new XAttribute(W + "w", 0), new XAttribute(W + "type", "dxa"))))), content)));

        static (double Baseline, PdfRectangle Clip) FindProseFrame(PdfPage page) =>
            DocxNoteClippingTests.ReadPhysicalPaintFrames(page).SelectMany(frame =>
                DocxCommentsTests.ReadEmbeddedGlyphTextShows(new PdfPage(page.Width, page.Height, frame.Block, page.Fonts))
                    .Where(show => show.Text.Trim() == prose).Select(show =>
                        (show.Y * frame.Transform.M22 + frame.Transform.M32, frame.Clip))).Single();
    }

    public static void DocxSplitNotesKeepFloatingBoxesWithTheirSourceLines()
    {
        string fixture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
            "TableGridOrigins", "Containers", "floating-table-ordered.docx"));
        var resolver = new PortableRenderingTests.PinnedResolver();
        XNamespace wp = OoxNamespaces.WordprocessingDrawingNamespace;
        int[] anchors = [4, 51, 84];
        foreach (string owner in new[] { "footnote", "endnote" })
        WithDocumentMutation(fixture, (parts, document) =>
        {
            var sourceDrawing = document.Descendants(wp + "anchor").Single().Parent!;
            var sourceParagraph = sourceDrawing.Ancestors(W + "p").First();
            sourceParagraph.ReplaceWith(new XElement(W + "p", new XElement(W + "r",
                new XElement(W + (owner + "Reference"), new XAttribute(W + "id", 1)))));
            var paragraphs = Enumerable.Range(0, 115).Select(index =>
            {
                var paragraph = Paragraph($"Note source line {index:000}.");
                if (!anchors.Contains(index)) return paragraph;
                var drawing = new XElement(sourceDrawing);
                var anchor = drawing.Element(wp + "anchor")!;
                anchor.Element(wp + "docPr")!.SetAttributeValue("id", 200 + index);
                anchor.Element(wp + "extent")!.SetAttributeValue("cy", 40 * 12700);
                anchor.Element(wp + "positionH")!.Element(wp + "posOffset")!.SetValue(170 * 12700);
                anchor.Element(wp + "positionV")!.Element(wp + "posOffset")!.SetValue(0);
                anchor.Descendants(W + "txbxContent").Single().ReplaceNodes(Paragraph($"Box for line {index:000}."));
                paragraph.Add(new XElement(W + "r", drawing));
                return paragraph;
            });
            AddStoryPart(parts, "split-note-" + owner, owner + "s", new XElement(W + (owner + "s"),
                new XElement(W + owner, new XAttribute(W + "id", 1), paragraphs)));
        }, path =>
        {
            foreach (var mode in Modes)
            {
                var geometry = mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                    : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
                var pages = new DocxRenderer(resolver, mode, geometry)
                    .RenderPagesAsync(DocxTests.ReadDocx(path, mode), null, default).ToBlockingEnumerable().ToArray();
                var shows = pages.SelectMany((page, index) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                    .Select(show => (Page: index, show.Text, show.X, show.Y))).ToArray();
                var lines = shows.Where(show => show.Text.TrimStart().StartsWith("Note source line ", StringComparison.Ordinal)).ToArray();
                TestAssert.True(lines.Select(line => line.Text.Trim()).SequenceEqual(
                    Enumerable.Range(0, 115).Select(index => $"Note source line {index:000}.")),
                    $"The {owner} lines retain their authored order in {mode}.");
                int previousPage = -1;
                foreach (int anchor in anchors)
                {
                    var line = lines[anchor];
                    var boxes = shows.Where(show => show.Text.Trim() == $"Box for line {anchor:000}.").ToArray();
                    TestAssert.Equal(1, boxes.Length);
                    var box = boxes[0];
                    TestAssert.True(line.Page > previousPage && box.Page == line.Page,
                        $"The {owner} box for line {anchor} belongs to its continued source page in {mode}.");
                    TestAssert.True(box.X > line.X + 100d && Math.Abs(box.Y - line.Y) < 7d,
                        $"The {owner} box for line {anchor} retains its horizontal offset and paragraph top in {mode}.");
                    previousPage = line.Page;
                }
            }
        });
    }

    public static void DocxFloatingBoxMathSurvivesBodyAndStoryPlacement()
    {
        string fixture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
            "TableGridOrigins", "Containers", "floating-table-ordered.docx"));
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        XNamespace wp = OoxNamespaces.WordprocessingDrawingNamespace;
        XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        double finalSize = 0d;
        foreach (string owner in new[] { "body", "header", "footer", "footnote", "endnote" })
        foreach (var mode in Modes)
        {
            string label = owner + "/" + mode;
            string partName = "floating-math-" + owner;
            WithDocumentMutation(fixture, (parts, document) =>
            {
                var body = document.Descendants(W + "body").Single();
                var drawing = body.Descendants(wp + "anchor").Single();
                drawing.Element(wp + "extent")!.SetAttributeValue("cy", 100 * 12700);
                var box = drawing.Descendants(W + "txbxContent").Single();
                var fraction = new XElement(m + "oMath", new XElement(m + "f",
                    new XElement(m + "num", new XElement(m + "r", new XElement(m + "t", "2"))),
                    new XElement(m + "den", new XElement(m + "r", new XElement(m + "t", "7")))));
                var caption = Paragraph("PUBLIC BOX CAPTION");
                box.ReplaceNodes(caption, new XElement(W + "p", fraction), Paragraph("PUBLIC BOX FOLLOWING"));

                if (owner == "body") return;
                var paragraph = drawing.Ancestors(W + "p").First();
                if (owner is "header" or "footer")
                {
                    var vertical = drawing.Element(wp + "positionV")!;
                    vertical.SetAttributeValue("relativeFrom", "page");
                    vertical.Element(wp + "posOffset")!.SetValue((owner == "header" ? 24 : 710) * 12700);
                    body.Element(W + "sectPr")!.AddFirst(new XElement(W + (owner + "Reference"),
                        new XAttribute(W + "type", "default"), new XAttribute(R + "id", partName)));
                    paragraph.Remove();
                    AddStoryPart(parts, partName, owner, new XElement(W + (owner == "header" ? "hdr" : "ftr"), paragraph));
                }
                else
                {
                    paragraph.AddBeforeSelf(new XElement(W + "p", new XElement(W + "r",
                        new XElement(W + (owner + "Reference"), new XAttribute(W + "id", 1)))));
                    paragraph.Remove();
                    AddStoryPart(parts, partName, owner + "s", new XElement(W + (owner + "s"),
                        new XElement(W + owner, new XAttribute(W + "id", 1), paragraph,
                            Enumerable.Range(1, 8).Select(index => Paragraph("Additional note line " + index)))));
                }
            }, path =>
            {
                var source = DocxTests.ReadDocx(path, mode);
                var geometry = mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                    : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
                var diagnostics = new List<OoxPdfDiagnostic>();
                var page = new DocxRenderer(resolver, mode, geometry)
                    .RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().Single();
                DocxInlineCompositionTests.AssertSourceLabelsPaint(page, mode, ["PUBLIC BOX CAPTION", "PUBLIC BOX FOLLOWING"], label);

                var mathFonts = page.Fonts.Where(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal))
                    .Select(font => font.Font).ToHashSet(ReferenceEqualityComparer.Instance);
                var symbols = DocxNoteClippingTests.ReadPhysicalPaintFrames(page).SelectMany(frame =>
                    DocxCommentsTests.ReadEmbeddedGlyphTextShows(new PdfPage(page.Width, page.Height, frame.Block, page.Fonts))
                        .Where(show => mathFonts.Contains(show.Font)).Select(show => (Frame: frame, Show: show))).ToArray();
                TestAssert.Equal("27", string.Concat(symbols.Select(symbol => symbol.Show.Text)).Normalize(NormalizationForm.FormKC));
                double printedSize = symbols[0].Frame.Transform.M11 * symbols[0].Show.Size;
                if (mode == OoxPdfDocxMarkupMode.Final) finalSize = printedSize;
                else
                {
                    double scale = DocxRenderer.ResolveWordCompatiblePrintScale(source, DocxMarkupContext.FromMode(mode, geometry));
                    TestAssert.True(scale < 1d && Math.Abs(printedSize - finalSize * scale) < .02d,
                        "The floating equation scales once with its printed box in " + label);
                }
                foreach (var (frame, show) in symbols)
                {
                    var font = show.Font.Font;
                    double unit = show.Size / font.UnitsPerEm;
                    TestAssert.True(font.TryGetGlyphBounds(font.MapCodePoint(show.Text[0]), out var ink), label);
                    var mtx = frame.Transform;
                    double left = (show.X + ink.XMin * unit) * mtx.M11 + mtx.M31;
                    double right = (show.X + ink.XMax * unit) * mtx.M11 + mtx.M31;
                    double bottom = (show.Y + ink.YMin * unit) * mtx.M22 + mtx.M32;
                    double top = (show.Y + ink.YMax * unit) * mtx.M22 + mtx.M32;
                    TestAssert.True(left >= frame.Clip.X - .01d && right <= frame.Clip.X + frame.Clip.Width + .01d &&
                        bottom >= frame.Clip.Y - .01d && top <= frame.Clip.Y + frame.Clip.Height + .01d,
                        "Both fraction glyphs survive the box and story clips in " + label);
                }
                TestAssert.Equal(0, diagnostics.Count(item => item.Id == "DOCX_UNSUPPORTED_EQUATION"));
                TestAssert.Equal(owner == "body" ? "/word/document.xml" : "/word/" + partName + ".xml",
                    diagnostics.Single(item => item.Id == "DOCX_APPROXIMATED_EQUATION").PartName);
            });
        }
    }

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
        foreach (var (story, height, exact, wrap, nestedAnchor) in new[]
        {
            ("header", 0, false, false, false), ("footer", 0, false, false, false),
            ("header", 2400, false, true, false), ("header", 300, true, true, false),
            ("header", 2400, false, true, true), ("footer", 2400, false, true, true)
        })
        foreach (var mode in Modes)
        {
            int offset = story == "footer" ? -62 : 30;
            string context = $"{story}/{height}/{mode}, nested anchor = {nestedAnchor}";
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

                var paragraph = nestedAnchor ? new XElement(W + "p") : Paragraph("STATIC CELL OWNER");
                paragraph.AddFirst(new XElement(W + "pPr",
                    new XElement(W + "spacing", new XAttribute(W + "before", "0"), new XAttribute(W + "after", "0"))));
                paragraph.Add(new XElement(anchor.Ancestors(W + "r").First()));
                anchor.Ancestors(W + "p").First().Remove();
                var cell = new XElement(W + "tc", new XElement(W + "tcPr",
                    new XElement(W + "tcW", new XAttribute(W + "w", "6000"), new XAttribute(W + "type", "dxa"))), paragraph);
                if (!exact) paragraph.AddBeforeSelf(Paragraph("EARLIER CELL PARAGRAPH"));
                if (nestedAnchor)
                    paragraph.AddBeforeSelf(new XElement(W + "tbl", new XElement(W + "tblPr"),
                        new XElement(W + "tblGrid", new XElement(W + "gridCol", new XAttribute(W + "w", "4800"))),
                        new XElement(W + "tr", new XElement(W + "tc", Paragraph("NESTED CELL CONTENT")))));
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
                        ["BOX TABLE NORTH", "BOX TABLE SOUTH", nestedAnchor ? "NESTED CELL CONTENT" : "STATIC CELL OWNER"], context);
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

    public static void DocxStaticCellFloatingBoxLinksKeepVisibleFragments()
    {
        const string label = "PUBLIC LINK";
        const string suffix = "Neighboring text";
        const string uri = "https://example.invalid/visible-cell-box";
        const string name = "floating-cell-link-header";
        string fixture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
            "TableGridOrigins", "Containers", "floating-table-ordered.docx"));
        var resolver = new PortableRenderingTests.PinnedResolver();
        XNamespace wp = OoxNamespaces.WordprocessingDrawingNamespace;
        XNamespace relationships = "http://schemas.openxmlformats.org/package/2006/relationships";

        foreach (var mode in Modes)
        {
            var geometry = mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
            var renderer = new DocxRenderer(resolver, mode, geometry);
            PdfPage fitting = Render(300, 85);
            var fittingLabel = PortableRenderingTests.ReadPaintedSourceLabel(fitting, label);
            var fittingSuffix = PortableRenderingTests.ReadPaintedSourceLabel(fitting, suffix);
            TestAssert.Equal(1, fitting.Annotations.Count(link => link.Uri == uri));

            PdfPage narrow = Render(102, 85);
            var narrowLabel = PortableRenderingTests.ReadPaintedSourceLabel(narrow, label);
            TestAssert.True(fittingLabel.SequenceEqual(narrowLabel), "The link's glyphs keep their ink and positions in the narrow cell.");
            PdfRectangle clip = ReadLabelClip(narrow);
            TestAssert.True(fittingSuffix.Max(glyph => glyph.Ink.X + glyph.Ink.Width) > clip.X + clip.Width,
                "The neighboring unlinked prose crosses the narrow cell clip.");
            var link = narrow.Annotations.SingleOrDefault(annotation => annotation.Uri == uri);
            TestAssert.True(link.Uri == uri, "The cell's clip must not remove a hyperlink over fully visible text.");
            TestAssert.True(link.X >= clip.X - .002d && link.Y >= clip.Y - .002d &&
                link.X + link.Width <= clip.X + clip.Width + .002d && link.Y + link.Height <= clip.Y + clip.Height + .002d,
                "The restored hyperlink stays inside the printed viewport.");

            PdfPage shortCell = Render(300, 36);
            TestAssert.True(ReadLabelClip(shortCell).Y > fittingLabel.Max(glyph => glyph.Ink.Y + glyph.Ink.Height),
                "The short cell clips all of the linked ink.");
            TestAssert.Equal(0, shortCell.Annotations.Count(annotation => annotation.Uri == uri));

            PdfPage merged = Render(300, 85, merge: true);
            var mergedLabel = PortableRenderingTests.ReadPaintedSourceLabel(merged, label);
            var mergedSuffix = PortableRenderingTests.ReadPaintedSourceLabel(merged, suffix);
            TestAssert.Equal(fittingLabel.Count, mergedLabel.Count);
            TestAssert.Equal(fittingSuffix.Count, mergedSuffix.Count);
            var mergedLink = merged.Annotations.Single(annotation => annotation.Uri == uri);
            TestAssert.True(mergedLabel.All(glyph => glyph.Ink.X >= mergedLink.X - .002d &&
                glyph.Ink.Y >= mergedLink.Y - .002d &&
                glyph.Ink.X + glyph.Ink.Width <= mergedLink.X + mergedLink.Width + .002d &&
                glyph.Ink.Y + glyph.Ink.Height <= mergedLink.Y + mergedLink.Height + .002d),
                "The restored link covers its visible label inside the merged cell.");
            if (mode == OoxPdfDocxMarkupMode.Final)
            {
                TestAssert.True(fittingLabel.SequenceEqual(mergedLabel) && fittingSuffix.SequenceEqual(mergedSuffix),
                    "Final keeps the ordinary-cell printed glyph geometry.");
                TestAssert.Equal(fitting.Annotations.Single(annotation => annotation.Uri == uri), mergedLink);
                TestAssert.Equal(ReadLabelClip(fitting), ReadLabelClip(merged));
            }

            PdfPage shortMerge = Render(300, 36, merge: true);
            TestAssert.True(ReadLabelClip(shortMerge).Y > mergedLabel.Max(glyph => glyph.Ink.Y + glyph.Ink.Height),
                "The short merge clips the linked ink.");
            TestAssert.Equal(0, shortMerge.Annotations.Count(annotation => annotation.Uri == uri));

            PdfPage Render(int width, int height, bool merge = false)
            {
                PdfPage? result = null;
                WithDocumentMutation(fixture, (parts, document) =>
                {
                    var anchor = document.Descendants(wp + "anchor").Single();
                    anchor.SetAttributeValue("layoutInCell", "1");
                    anchor.Element(wp + "positionV")!.Element(wp + "posOffset")!.SetValue(30 * 12700);
                    anchor.Element(wp + "wrapNone")!.ReplaceWith(new XElement(wp + "wrapSquare", new XAttribute("wrapText", "bothSides")));
                    anchor.Descendants(W + "txbxContent").Single().ReplaceNodes(
                        new XElement(W + "p",
                            new XElement(W + "pPr", new XElement(W + "spacing", new XAttribute(W + "before", "0"),
                                new XAttribute(W + "after", "0"), new XAttribute(W + "line", "300"), new XAttribute(W + "lineRule", "exact"))),
                            new XElement(W + "hyperlink", new XAttribute(R + "id", "box-uri"),
                                new XElement(W + "r", new XElement(W + "rPr", new XElement(W + "sz", new XAttribute(W + "val", "24"))),
                                    new XElement(W + "t", label))),
                            new XElement(W + "r", new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), " " + suffix))));

                    var owner = Paragraph("CELL OWNER");
                    owner.AddFirst(new XElement(W + "pPr", new XElement(W + "spacing",
                        new XAttribute(W + "before", "0"), new XAttribute(W + "after", "0"))));
                    owner.Add(new XElement(anchor.Ancestors(W + "r").First()));
                    anchor.Ancestors(W + "p").First().Remove();
                    int cellWidth = width * 20;
                    int rowCount = merge ? 3 : 1;
                    int rowHeight = height * 20 / rowCount;
                    var rows = Enumerable.Range(0, rowCount).Select(row => new XElement(W + "tr",
                        new XElement(W + "trPr", new XElement(W + "trHeight",
                            new XAttribute(W + "val", row == rowCount - 1 ? height * 20 - row * rowHeight : rowHeight),
                            new XAttribute(W + "hRule", "exact"))),
                        new XElement(W + "tc", new XElement(W + "tcPr",
                            new XElement(W + "tcW", new XAttribute(W + "w", cellWidth), new XAttribute(W + "type", "dxa")),
                            merge ? new XElement(W + "vMerge", row == 0 ? new XAttribute(W + "val", "restart") : null) : null),
                            row == 0 ? owner : new XElement(W + "p"))));
                    var table = new XElement(W + "tbl",
                        new XElement(W + "tblPr", new XElement(W + "tblLayout", new XAttribute(W + "type", "fixed"))),
                        new XElement(W + "tblGrid", new XElement(W + "gridCol", new XAttribute(W + "w", cellWidth))),
                        rows);
                    document.Descendants(W + "sectPr").Single().AddFirst(new XElement(W + "headerReference",
                        new XAttribute(W + "type", "default"), new XAttribute(R + "id", name)));
                    AddStoryPart(parts, name, "header", new XElement(W + "hdr", table));
                    parts[$"word/_rels/{name}.xml.rels"] = Encoding.UTF8.GetBytes(new XElement(relationships + "Relationships",
                        new XElement(relationships + "Relationship", new XAttribute("Id", "box-uri"),
                            new XAttribute("Type", R.NamespaceName + "/hyperlink"), new XAttribute("Target", uri),
                            new XAttribute("TargetMode", "External"))).ToString());
                }, path => result = renderer.RenderPagesAsync(DocxTests.ReadDocx(path, mode), null, default).ToBlockingEnumerable().Single());
                return result ?? throw new InvalidOperationException("The header fixture must render one page.");
            }
        }

        static PdfRectangle ReadLabelClip(PdfPage page) => DocxNoteClippingTests.ReadPhysicalPaintFrames(page)
            .Single(frame => DocxCommentsTests.ReadEmbeddedGlyphTextShows(new PdfPage(page.Width, page.Height, frame.Block, page.Fonts))
                .Any(show => show.Text.Trim() == label)).Clip;
    }

    public static void DocxBoxedPicturesLinkWithinTheirPaintedViewports()
    {
        const string uri = "https://example.invalid/boxed-picture";
        const int emuPerPoint = 12700;
        string fixture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
            "InlineBoxMath", "inline-box-mixed-source-report.docx"));
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        XNamespace wp = OoxNamespaces.WordprocessingDrawingNamespace;
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        foreach (var mode in Modes)
        {
            CheckBox(false, false, 72d, mode, clipped: false);
            CheckBox(false, true, 72d, mode, clipped: false);
            double partialHeight = CheckBox(true, false, 72d, mode, clipped: false);
            CheckBox(true, false, partialHeight, mode, clipped: true);
        }

        double CheckBox(bool floating, bool cell, double height, OoxPdfDocxMarkupMode mode, bool clipped)
        {
            double partialHeight = 0d;
            string context = $"floating={floating}, cell={cell}, height={height:F3}, {mode}";
            WithDocumentMutation(fixture, (parts, xml) =>
            {
                var body = xml.Descendants(W + "body").Single();
                var box = body.Descendants(W + "txbxContent").Single();
                var inner = box.Descendants(wp + "inline").Single(drawing => drawing.Descendants(a + "blip").Any());
                var props = inner.Element(wp + "docPr");
                if (props is null)
                {
                    props = new XElement(wp + "docPr", new XAttribute("id", "9101"), new XAttribute("name", "Public linked picture"));
                    inner.Element(wp + "extent")!.AddAfterSelf(props);
                }
                props.Add(new XElement(a + "hlinkClick", new XAttribute(R + "id", "rIdPublicBoxPicture")));
                var relationships = XDocument.Parse(Encoding.UTF8.GetString(parts["word/_rels/document.xml.rels"]));
                relationships.Root!.Add(new XElement(relationships.Root.Name.Namespace + "Relationship",
                    new XAttribute("Id", "rIdPublicBoxPicture"), new XAttribute("Type", R.NamespaceName + "/hyperlink"),
                    new XAttribute("Target", uri), new XAttribute("TargetMode", "External")));
                parts["word/_rels/document.xml.rels"] = Encoding.UTF8.GetBytes(relationships.ToString());

                var outer = box.Ancestors(wp + "inline").First();
                var extent = outer.Element(wp + "extent")!;
                extent.SetAttributeValue("cy", (long)Math.Round(height * emuPerPoint));
                if (floating)
                {
                    outer.Name = wp + "anchor";
                    outer.SetAttributeValue("simplePos", "0");
                    outer.SetAttributeValue("relativeHeight", "1");
                    outer.SetAttributeValue("layoutInCell", "1");
                    outer.AddFirst(new XElement(wp + "simplePos", new XAttribute("x", "0"), new XAttribute("y", "0")),
                        Position("H", 80), Position("V", 155));
                    extent.AddAfterSelf(new XElement(wp + "wrapNone"));
                }
                if (cell)
                {
                    var owner = outer.Ancestors(W + "p").First();
                    owner.ReplaceWith(new XElement(W + "tbl", new XElement(W + "tblPr"),
                        new XElement(W + "tblGrid", new XElement(W + "gridCol", new XAttribute(W + "w", "9000"))),
                        new XElement(W + "tr", new XElement(W + "tc", new XElement(owner)))));
                }

                var comment = Paragraph("PUBLIC PICTURE REVIEW");
                comment.AddFirst(new XElement(W + "commentRangeStart", new XAttribute(W + "id", "7")));
                comment.Add(new XElement(W + "commentRangeEnd", new XAttribute(W + "id", "7")),
                    new XElement(W + "r", new XElement(W + "commentReference", new XAttribute(W + "id", "7"))));
                body.Element(W + "sectPr")!.AddBeforeSelf(comment);
                const string commentsPart = "word/comments.xml";
                var comments = XDocument.Parse(Encoding.UTF8.GetString(parts[commentsPart]));
                comments.Root!.Add(new XElement(W + "comment", new XAttribute(W + "id", "7"),
                    new XAttribute(W + "author", "Public"), Paragraph("REVIEW COMMENT")));
                parts[commentsPart] = Encoding.UTF8.GetBytes(comments.ToString());

                XElement Position(string axis, int offset) => new(wp + "position" + axis, new XAttribute("relativeFrom", "page"),
                    new XElement(wp + "posOffset", offset * emuPerPoint));
            }, path =>
            {
                var geometry = mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                    : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
                var source = DocxTests.ReadDocx(path, mode);
                TestAssert.True(DocxBlockTraversal.EnumerateResourceImages(source, default).Any(image => image.ClickTarget == uri),
                    "The reader retained the authored picture target.");
                double scale = DocxRenderer.ResolveWordCompatiblePrintScale(source, DocxMarkupContext.FromMode(mode, geometry));
                if (mode == OoxPdfDocxMarkupMode.AllMarkup)
                    TestAssert.True(scale < 1d, "The Review control has a printed markup scale: " + context);
                var page = new DocxRenderer(resolver, mode, geometry)
                    .RenderPagesAsync(source, null, default).ToBlockingEnumerable().Single();
                string imagePaint = "/" + page.Images.Single().ResourceName + " Do";
                var paint = DocxNoteClippingTests.ReadPhysicalPaintFrames(page).Single(frame => frame.Block == imagePaint);
                var frame = new PdfRectangle(paint.Transform.M31, paint.Transform.M32, paint.Transform.M11, paint.Transform.M22);
                var links = page.Annotations.Where(link => link.Uri == uri).ToArray();
                if (!clipped)
                {
                    TestAssert.True(frame.X >= paint.Clip.X - .002d && frame.Y >= paint.Clip.Y - .002d &&
                        frame.X + frame.Width <= paint.Clip.X + paint.Clip.Width + .002d &&
                        frame.Y + frame.Height <= paint.Clip.Y + paint.Clip.Height + .002d,
                        "The complete picture is inside the active PDF clip: " + context);
                    TestAssert.True(links.Length == 1, "The complete boxed picture has one URI annotation: " + context);
                    var link = links[0];
                    TestAssert.True(Math.Abs(link.X - frame.X) < .002d && Math.Abs(link.Y - frame.Y) < .002d &&
                        Math.Abs(link.Width - frame.Width) < .002d && Math.Abs(link.Height - frame.Height) < .002d,
                        "The link rectangle matches the emitted image placement: " + context);
                    if (floating)
                    {
                        // Set the authored bottom edge through the picture's center.
                        // The tall control includes supported equation line heights.
                        partialHeight = height - (frame.Y + frame.Height / 2d - paint.Clip.Y) / scale;
                        TestAssert.True(partialHeight > 0d && partialHeight < height, context);
                    }
                }
                else
                {
                    TestAssert.True(frame.Y < paint.Clip.Y - .002d && frame.Y + frame.Height > paint.Clip.Y,
                        "The short box clips part of the picture: " + context);
                    TestAssert.Equal(0, links.Length);
                }
            });
            return partialHeight;
        }
    }

    public static void DocxRotatedStoryPictureLinksFollowPaintedViewports()
    {
        const string uri = "https://example.invalid/rotated-story-picture";
        const int size = 18 * 12700;
        string fixture = DocxGlyphlessAnchorTests.CreateDocument();
        var resolver = new PortableRenderingTests.PinnedResolver();
        XNamespace wp = OoxNamespaces.WordprocessingDrawingNamespace;
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XNamespace pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";
        foreach (var (owner, clipped, floating) in new[]
        {
            ("header", false, false), ("footer", false, false), ("footnote", false, false), ("endnote", false, false),
            ("header", true, false), ("footnote", true, false),
            ("footnote", false, true), ("endnote", false, true), ("footnote", true, true)
        })
        foreach (var mode in Modes)
        {
            int horizontalOffset = clipped ? 453 : 275;
            string context = $"{owner}, clipped={clipped}, floating={floating}, {mode}";
            WithDocumentMutation(fixture, (parts, document) =>
            {
                var body = document.Descendants(W + "body").Single();
                var section = body.Element(W + "sectPr")!;
                var graphic = new XElement(body.Descendants(a + "graphic").Single());
                graphic.Descendants(pic + "pic").Single().Add(new XElement(pic + "spPr", new XElement(a + "xfrm",
                    new XAttribute("rot", 90 * 60000),
                    new XElement(a + "off", new XAttribute("x", 0), new XAttribute("y", 0)),
                    new XElement(a + "ext", new XAttribute("cx", size), new XAttribute("cy", size)))));
                var drawing = new XElement(wp + (floating ? "anchor" : "inline"),
                    new XElement(wp + "extent", new XAttribute("cx", size), new XAttribute("cy", size)),
                    new XElement(wp + "docPr", new XAttribute("id", 1), new XAttribute("name", "Linked story picture"),
                        new XElement(a + "hlinkClick", new XAttribute(R + "id", "rIdPictureLink"))), graphic);
                if (floating)
                {
                    drawing.SetAttributeValue("behindDoc", "0");
                    drawing.AddFirst(new XElement(wp + "positionH", new XAttribute("relativeFrom", "margin"),
                            new XElement(wp + "posOffset", horizontalOffset * 12700)),
                        new XElement(wp + "positionV", new XAttribute("relativeFrom", "paragraph"),
                            new XElement(wp + "posOffset", 10 * 12700)));
                    drawing.Element(wp + "extent")!.AddAfterSelf(new XElement(wp + "wrapNone"));
                }
                var paragraph = Paragraph("Story picture");
                paragraph.AddFirst(
                    new XElement(W + "pPr", new XElement(W + "spacing",
                        new XAttribute(W + "line", clipped && !floating ? 300 : 960),
                        new XAttribute(W + "lineRule", clipped && !floating ? "exact" : "atLeast"))),
                    new XElement(W + "r", new XElement(W + "drawing", drawing)));
                var anchor = Paragraph("PUBLIC REVIEW ANCHOR");
                anchor.AddFirst(new XElement(W + "commentRangeStart", new XAttribute(W + "id", 7)));
                anchor.Add(new XElement(W + "commentRangeEnd", new XAttribute(W + "id", 7)),
                    new XElement(W + "r", new XElement(W + "commentReference", new XAttribute(W + "id", 7))));
                body.Element(W + "p")!.ReplaceWith(anchor);
                AddStoryPart(parts, "picture-review-comments", "comments", new XElement(W + "comments",
                    new XElement(W + "comment", new XAttribute(W + "id", 7), new XAttribute(W + "author", "Public"),
                        Paragraph("REVIEW COMMENT"))));

                string name = "rotated-picture-" + owner;
                if (owner is "header" or "footer")
                {
                    section.AddFirst(new XElement(W + (owner + "Reference"),
                        new XAttribute(W + "type", "default"), new XAttribute(R + "id", name)));
                    AddStoryPart(parts, name, owner, new XElement(W + (owner == "header" ? "hdr" : "ftr"), paragraph));
                }
                else
                {
                    anchor.Add(new XElement(W + "r", new XElement(W + (owner + "Reference"), new XAttribute(W + "id", 1))));
                    AddStoryPart(parts, name, owner + "s", new XElement(W + (owner + "s"),
                        new XElement(W + owner, new XAttribute(W + "id", 1), paragraph)));
                }
                var relationships = XDocument.Parse(Encoding.UTF8.GetString(parts["word/_rels/document.xml.rels"]));
                XNamespace relationshipsNamespace = relationships.Root!.Name.Namespace;
                var image = relationships.Root.Elements().Single(element => (string?)element.Attribute("Id") == "rIdImage1");
                var storyRelationships = new XElement(relationshipsNamespace + "Relationships", new XElement(image),
                    new XElement(relationshipsNamespace + "Relationship",
                        new XAttribute("Id", "rIdPictureLink"), new XAttribute("Type", R.NamespaceName + "/hyperlink"),
                        new XAttribute("Target", uri), new XAttribute("TargetMode", "External")));
                parts[$"word/_rels/{name}.xml.rels"] = Encoding.UTF8.GetBytes(storyRelationships.ToString());
            }, path =>
            {
                var geometry = mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                    : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
                var source = DocxTests.ReadDocx(path, mode);
                double scale = DocxRenderer.ResolveWordCompatiblePrintScale(source, DocxMarkupContext.FromMode(mode, geometry));
                if (mode == OoxPdfDocxMarkupMode.AllMarkup)
                    TestAssert.True(scale < 1d, "The Review control uses the printed markup scale: " + context);
                var page = new DocxRenderer(resolver, mode, geometry)
                    .RenderPagesAsync(source, null, default).ToBlockingEnumerable().Single();
                TestAssert.True(page.Images.Count == 1, "The selected story paints one picture: " + context);
                string imagePaint = "/" + page.Images[0].ResourceName + " Do";
                var paint = DocxNoteClippingTests.ReadPhysicalPaintFrames(page).Single(frame => frame.Block == imagePaint);
                var m = paint.Transform;
                TestAssert.True(m.M12 < -8d && m.M21 > 8d && Math.Abs(m.M11) < .002d && Math.Abs(m.M22) < .002d,
                    "The picture paints with the authored quarter-turn: " + context);
                var bounds = new PdfRectangle(m.M31, m.M32 + m.M12, m.M21, -m.M12);
                if (floating)
                {
                    TestAssert.True(Math.Abs(bounds.X - (72d + horizontalOffset) * scale) < .003d &&
                        Math.Abs(bounds.Width - 18d * scale) < .003d && Math.Abs(bounds.Height - 18d * scale) < .003d,
                        "The margin-relative picture position and extent use the printed note frame: " + context);
                    TestAssert.True(bounds.X < paint.Clip.X + paint.Clip.Width &&
                        bounds.Y + bounds.Height > paint.Clip.Y && bounds.Y < paint.Clip.Y + paint.Clip.Height,
                        "The note retains visible picture area, including the clipped control: " + context);
                    TestAssert.Contains("Story picture", Text(page));
                }
                bool contained = paint.Clip.X <= bounds.X + .002d && paint.Clip.Y <= bounds.Y + .002d &&
                    paint.Clip.X + paint.Clip.Width >= bounds.X + bounds.Width - .002d &&
                    paint.Clip.Y + paint.Clip.Height >= bounds.Y + bounds.Height - .002d;
                TestAssert.True(contained != clipped && paint.Clip.Width > 0d && paint.Clip.Height > 0d,
                    "The picture is fully visible or retains partial ink as authored: " + context);
                var links = page.Annotations.Where(link => link.Uri == uri).ToArray();
                TestAssert.True(links.Length == (clipped ? 0 : 1), "The story picture has the expected URI count: " + context);
                if (!clipped)
                {
                    var link = links[0];
                    TestAssert.True(Math.Abs(link.X - bounds.X) < .002d && Math.Abs(link.Y - bounds.Y) < .002d &&
                        Math.Abs(link.Width - bounds.Width) < .002d && Math.Abs(link.Height - bounds.Height) < .002d,
                        "The URI rectangle matches the rotated painted bounds: " + context);
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

    public static void DocxBodyBookmarksRespectClosedRevisionBoundaries()
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
    }

    public static void DocxTableRowRangeBookmarksKeepTheirPaintedDestination()
    {
        foreach (var mode in Modes)
        {
            PdfPage[] expected = RenderVariant("joined-tables", mode);
            TestAssert.Equal(2, expected.Length);
            TestAssert.Contains("FIRST ROW", Text(expected[0]));
            TestAssert.DoesNotContain("SECOND ROW", Text(expected[0]));
            TestAssert.Contains("SECOND ROW", Text(expected[1]));
            PdfLinkAnnotation link = expected[0].Annotations.Single();
            TestAssert.True(link.Destination is { PageIndex: 1, Left: > 0d, Top: > 0d } destination &&
                destination.Top <= expected[1].Height, "The target row starts on page two.");

            foreach (string variant in new[] { "row-range", "wrapped-row-range", "closed-range" })
            {
                PdfPage[] pages = RenderVariant(variant, mode);
                TestAssert.Equal(expected.Length, pages.Length);
                for (int pageIndex = 0; pageIndex < pages.Length; pageIndex++)
                    TestAssert.Equal(expected[pageIndex].Content, pages[pageIndex].Content);
                if (variant == "closed-range")
                    TestAssert.Equal(0, pages[0].Annotations.Count);
                else
                    TestAssert.Equal(link, pages[0].Annotations.Single());
            }
        }

        var final = RenderVariant("deleted-row", OoxPdfDocxMarkupMode.Final);
        TestAssert.Equal(0, final[0].Annotations.Count);
        TestAssert.DoesNotContain("SECOND ROW", string.Concat(final.Select(Text)));

        static PdfPage[] RenderVariant(string variant, OoxPdfDocxMarkupMode mode)
        {
            PdfPage[] pages = [];
            WithMutation("bookmark-before", (_, xml) => {
                (XElement start, XElement end, XElement target) = BookmarkSource(xml);
                target.ElementsBeforeSelf(W + "p").Last().Remove();
                start.Remove();
                end.Remove();
                XElement first = Row("FIRST ROW");
                XElement second = Row("SECOND ROW");
                if (variant == "joined-tables")
                    target.AddBeforeSelf(Table(first), start, Table(second), end);
                else if (variant == "wrapped-row-range")
                    target.AddBeforeSelf(Table(first, new XElement(W + "sdt",
                        new XElement(W + "sdtContent", start, second, end))));
                else if (variant == "closed-range")
                    target.AddBeforeSelf(Table(first, start, end, second));
                else
                {
                    if (variant == "deleted-row")
                        second.Element(W + "trPr")!.Add(new XElement(W + "del",
                            new XAttribute(W + "id", "99"), new XAttribute(W + "author", "Public")));
                    target.AddBeforeSelf(Table(first, start, second, end));
                }
                XElement section = xml.Descendants(W + "sectPr").Single();
                section.Element(W + "pgSz")!.SetAttributeValue(W + "w", "10000");
                section.Element(W + "pgSz")!.SetAttributeValue(W + "h", "8000");
                foreach (string side in new[] { "top", "bottom", "left", "right" })
                    section.Element(W + "pgMar")!.SetAttributeValue(W + side, "800");
            }, path => {
                var document = DocxTests.ReadDocx(path, mode);
                TestAssert.Equal(variant == "deleted-row" ? 1 : 2, document.Tables.Single().Rows.Count);
                pages = Renderer(mode).RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
            });
            return pages;
        }

        static XElement Row(string caption) => new(W + "tr",
            new XElement(W + "trPr", new XElement(W + "cantSplit"),
                new XElement(W + "trHeight", new XAttribute(W + "val", "4400"), new XAttribute(W + "hRule", "exact"))),
            new XElement(W + "tc", Paragraph(caption)));

        static XElement Table(params XElement[] children) => new(W + "tbl",
            new XElement(W + "tblPr", new XElement(W + "tblLayout", new XAttribute(W + "type", "fixed"))),
            new XElement(W + "tblGrid", new XElement(W + "gridCol", new XAttribute(W + "w", "3000"))), children);
    }

    public static void DocxNestedRowBookmarksFollowThePaintedMergeRestart()
    {
        foreach (var mode in Modes)
        {
            PdfPage plain = RenderTable(mode, merge: false).Single();
            PdfPage merged = RenderTable(mode, merge: true).Single();
            const string expectedText = "JUMP NESTED NEIGHBOR 1 NEIGHBOR 2 NEIGHBOR 3";
            TestAssert.Equal(expectedText, string.Join(' ', Text(plain).Split(' ', StringSplitOptions.RemoveEmptyEntries)));
            TestAssert.Equal(expectedText, string.Join(' ', Text(merged).Split(' ', StringSplitOptions.RemoveEmptyEntries)));
            PdfLinkAnnotation expected = plain.Annotations.Single();
            TestAssert.True(expected.Destination is { PageIndex: 0, Left: > 0d, Top: > 0d } target &&
                target.Top <= plain.Height, "The unmerged control targets the visible nested row on page one.");
            TestAssert.Equal(expected, merged.Annotations.Single());

            var split = RenderTable(mode, merge: true, pageHeight: 2300);
            TestAssert.Contains("JUMP", Text(split[0]));
            TestAssert.True(split.Skip(1).Any(page => Text(page).Contains("NEIGHBOR 3", StringComparison.Ordinal)),
                "The merge reaches a later physical page.");
            TestAssert.Equal(0, split.Sum(page => page.Annotations.Count));

            PdfPage clipped = RenderTable(mode, merge: true, clipNestedRow: true).Single();
            TestAssert.Contains("BEFORE TARGET", Text(clipped));
            TestAssert.DoesNotContain("NESTED", Text(clipped));
            TestAssert.Equal(0, clipped.Annotations.Count);
        }

        static PdfPage[] RenderTable(OoxPdfDocxMarkupMode mode, bool merge, int pageHeight = 6000, bool clipNestedRow = false)
        {
            var nested = Table(2000, 1,
                clipNestedRow ? Row(3600, new XElement(W + "tc", Paragraph("BEFORE TARGET"))) : null,
                new XElement(W + "bookmarkStart", new XAttribute(W + "id", 14), new XAttribute(W + "name", "NestedDestination")),
                new XElement(W + "tr", new XElement(W + "tc", Paragraph("NESTED"))),
                new XElement(W + "bookmarkEnd", new XAttribute(W + "id", 14)));
            var rows = Enumerable.Range(1, 3).Select(number => Row(960,
                new XElement(W + "tc",
                    merge ? new XElement(W + "tcPr", new XElement(W + "vMerge",
                        number == 1 ? new XAttribute(W + "val", "restart") : null)) : null,
                    number == 1 ? nested : null,
                    new XElement(W + "p")),
                new XElement(W + "tc", Paragraph($"NEIGHBOR {number}"))));
            var document = new XElement(W + "document", new XElement(W + "body",
                new XElement(W + "p", new XElement(W + "hyperlink", new XAttribute(W + "anchor", "NestedDestination"),
                    new XElement(W + "r", new XElement(W + "t", "JUMP")))),
                Table(3000, 2, rows),
                new XElement(W + "sectPr",
                    new XElement(W + "pgSz", new XAttribute(W + "w", 8640), new XAttribute(W + "h", pageHeight)),
                    new XElement(W + "pgMar", new XAttribute(W + "top", 480), new XAttribute(W + "bottom", 480),
                        new XAttribute(W + "left", 720), new XAttribute(W + "right", 720)))));
            using var input = TestFixtures.CreateZipPackage(new Dictionary<string, string>
            {
                ["[Content_Types].xml"] = TestFixtures.DocxContentTypes,
                ["_rels/.rels"] = TestFixtures.DocxPackageRelationships,
                ["word/document.xml"] = document.ToString(SaveOptions.DisableFormatting)
            });
            using var package = OoxPackage.Open(input, default);
            var source = new DocxReader().Read(package, null, default, markupMode: mode);
            return Renderer(mode).RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
        }

        static XElement Row(int height, params XElement[] cells) => new(W + "tr",
            new XElement(W + "trPr", new XElement(W + "cantSplit"),
                new XElement(W + "trHeight", new XAttribute(W + "val", height), new XAttribute(W + "hRule", "exact"))),
            cells);

        static XElement Table(int width, int columns, params object?[] children) => new(W + "tbl",
            new XElement(W + "tblPr", new XElement(W + "tblLayout", new XAttribute(W + "type", "fixed"))),
            new XElement(W + "tblGrid", Enumerable.Range(0, columns)
                .Select(_ => new XElement(W + "gridCol", new XAttribute(W + "w", width)))), children);
    }

    public static void DocxCellBoundaryBookmarksKeepThePaintedCellDestination()
    {
        foreach (var mode in Modes)
        {
            foreach (bool nested in new[] { false, true })
            {
                PdfPage[] control = RenderCell("inside", nested, mode);
                TestAssert.Equal(2, control.Length);
                TestAssert.Contains("NEIGHBOR CELL", Text(control[1]));
                TestAssert.Contains("TARGET PARAGRAPH", Text(control[1]));
                var neighbor = DocxParagraphShadingTests.ReadRectangles(control[1].Content).Single(rect => rect.Color == "CC9944");
                PdfLinkAnnotation link = control[0].Annotations.Single();
                TestAssert.True(link.Destination is { PageIndex: 1, Left: { } x } && x > neighbor.X + neighbor.Width / 2d,
                    "The destination is in the second cell on page two.");

                foreach (string placement in new[] { "around-cell", "wrapped-cell" })
                {
                    PdfPage[] pages = RenderCell(placement, nested, mode);
                    TestAssert.True(control.Select(page => page.Content).SequenceEqual(pages.Select(page => page.Content)),
                        "The bookmark placement preserves the painted pages.");
                    TestAssert.Equal(link, pages[0].Annotations.Single());
                }
            }

            PdfPage[] closed = RenderCell("closed-range", nested: false, mode);
            TestAssert.Equal(2, closed.Length);
            TestAssert.Contains("TARGET PARAGRAPH", Text(closed[1]));
            TestAssert.Equal(0, closed[0].Annotations.Count);
        }

        static PdfPage[] RenderCell(string placement, bool nested, OoxPdfDocxMarkupMode mode)
        {
            PdfPage[] pages = [];
            WithMutation("bookmark-before", (_, xml) => {
                (XElement start, XElement end, XElement target) = BookmarkSource(xml);
                start.Remove();
                end.Remove();
                var block = new XElement(target);
                if (nested)
                    block = Table(new XElement(W + "tr", new XElement(W + "tc", block)), "1600");
                var cell = new XElement(W + "tc", block);
                if (nested) cell.Add(Paragraph(string.Empty));
                var row = new XElement(W + "tr", new XElement(W + "tc",
                    new XElement(W + "tcPr", new XElement(W + "shd", new XAttribute(W + "fill", "CC9944"))),
                    Paragraph("NEIGHBOR CELL")));
                switch (placement)
                {
                    case "inside":
                        block.AddBeforeSelf(start);
                        block.AddAfterSelf(end);
                        row.Add(cell);
                        break;
                    case "wrapped-cell":
                        row.Add(new XElement(W + "sdt", new XElement(W + "sdtContent", start, cell, end)));
                        break;
                    case "closed-range":
                        row.Add(start, end, cell);
                        break;
                    default:
                        row.Add(start, cell, end);
                        break;
                }
                target.ReplaceWith(Table(row, "2000", "4400"));
            }, path => pages = Render(path, mode));
            return pages;
        }

        static XElement Table(XElement row, params string[] widths) => new(W + "tbl",
            new XElement(W + "tblPr", new XElement(W + "tblLayout", new XAttribute(W + "type", "fixed"))),
            new XElement(W + "tblGrid", widths.Select(width => new XElement(W + "gridCol", new XAttribute(W + "w", width)))), row);
    }

    public static void DocxColumnBookmarksFollowTheAuthoredTableGrid()
    {
        foreach (var mode in Modes)
        foreach (bool spanned in new[] { false, true })
        {
            PdfPage[] control = RenderColumn("inside", spanned, mode);
            TestAssert.Equal(2, control.Length);
            TestAssert.Contains("FIRST CELL", Text(control[1]));
            TestAssert.Contains("TARGET PARAGRAPH", Text(control[1]));
            PdfLinkAnnotation expected = control[0].Annotations.Single();
            TestAssert.True(expected.Destination is { PageIndex: 1, Left: > 100d },
                "The control bookmark points to the second cell on page two.");
            foreach (string variant in new[] { "column", "wrapped", "closed", "unoccupied",
                         "row-boundary", "row-boundary-closed", "row-boundary-closed-in-cell",
                         "row-boundary-closed-before-cell", "inside-row", "inside-row-closed",
                         "between-cell-start", "between-cell-end", "end-after-paragraph", "end-in-paragraph",
                         "inline-start", "trailing-start", "inline-closed", "trailing-closed" })
            {
                PdfPage[] pages = RenderColumn(variant, spanned, mode);
                TestAssert.True(control.Select(page => page.Content).SequenceEqual(pages.Select(page => page.Content)),
                    "A column bookmark preserves the painted page content.");
                if (variant is "closed" or "unoccupied" or "row-boundary-closed" or "row-boundary-closed-in-cell"
                    or "row-boundary-closed-before-cell" or "inside-row-closed" or "between-cell-end"
                    or "end-after-paragraph" or "end-in-paragraph" or "inline-closed" or "trailing-closed")
                    TestAssert.Equal(0, pages[0].Annotations.Count);
                else
                    TestAssert.Equal(expected, pages[0].Annotations.Single());
            }

            var ordinary = RenderColumn("inside-first", spanned, mode)[0].Annotations.Single();
            TestAssert.True(ordinary.Destination?.Left < expected.Destination?.Left,
                "The ordinary inline target is inside the first cell.");
            TestAssert.Equal(ordinary, RenderColumn("inside-first-column", spanned, mode)[0].Annotations.Single());

            if (!spanned)
            {
                var pages = RenderColumn("between-cell-multiple", spanned, mode);
                var blockStarts = RenderColumn("cell-block-multiple", spanned, mode);
                TestAssert.Contains("THIRD CELL", Text(pages[1]));
                TestAssert.True(pages.Select(page => page.Content).SequenceEqual(blockStarts.Select(page => page.Content)),
                    "Moving the starts into their cells keeps the same painted content.");
                var links = pages[0].Annotations;
                TestAssert.Equal(3, links.Count);
                TestAssert.True(links[2].Destination is { PageIndex: 1, Left: > 200d },
                    "The ordinary bookmark points to the third cell.");
                TestAssert.Equal(links[2].Destination, links[0].Destination);
                TestAssert.Equal(links[2].Destination, links[1].Destination);
                TestAssert.True(links.SequenceEqual(blockStarts[0].Annotations),
                    "Both cell-block starts and the ordinary bookmark reach the same authored cell.");
            }
        }

        static PdfPage[] RenderColumn(string variant, bool spanned, OoxPdfDocxMarkupMode mode)
        {
            PdfPage[] pages = [];
            WithMutation("bookmark-before", (_, xml) => {
                (XElement start, XElement end, XElement target) = BookmarkSource(xml);
                start.Remove();
                end.Remove();
                var first = new XElement(W + "tc", Paragraph("FIRST CELL"));
                var second = new XElement(W + "tc", new XElement(target));
                var row = new XElement(W + "tr");
                bool multipleStarts = variant is "between-cell-multiple" or "cell-block-multiple";
                int gridSize = multipleStarts ? 3 : spanned ? 5 : 2;
                if (spanned)
                {
                    row.Add(new XElement(W + "trPr", new XElement(W + "gridBefore", new XAttribute(W + "val", 1))));
                    foreach (var cell in new[] { first, second })
                        cell.AddFirst(new XElement(W + "tcPr", new XElement(W + "gridSpan", new XAttribute(W + "val", 2))));
                }
                row.Add(first, second);
                var table = new XElement(W + "tbl",
                    new XElement(W + "tblPr", new XElement(W + "tblLayout", new XAttribute(W + "type", "fixed"))),
                    new XElement(W + "tblGrid", Enumerable.Range(0, gridSize)
                        .Select(_ => new XElement(W + "gridCol", new XAttribute(W + "w", 1600)))), row);
                if (multipleStarts)
                {
                    var third = Paragraph("THIRD CELL");
                    third.AddFirst(new XElement(W + "bookmarkStart",
                        new XAttribute(W + "id", 12), new XAttribute(W + "name", "OrdinaryThird")));
                    third.Add(new XElement(W + "bookmarkEnd", new XAttribute(W + "id", 12)));
                    row.Add(new XElement(W + "tc", third));
                    start.SetAttributeValue(W + "colFirst", 2);
                    start.SetAttributeValue(W + "colLast", 2);
                    var nextParagraph = Paragraph("More context in the first cell");
                    first.Add(nextParagraph);
                    var another = new XElement(start);
                    another.SetAttributeValue(W + "id", 11);
                    another.SetAttributeValue(W + "name", "LaterThird");
                    if (variant == "cell-block-multiple")
                    {
                        nextParagraph.AddBeforeSelf(start);
                        second.Element(W + "p")!.AddBeforeSelf(another);
                    }
                    else
                    {
                        first.AddBeforeSelf(start);
                        second.AddBeforeSelf(another);
                    }
                    row.Add(end, new XElement(W + "bookmarkEnd", new XAttribute(W + "id", 11)));
                    var sourceLink = xml.Descendants(W + "hyperlink").Single();
                    var laterLink = new XElement(sourceLink);
                    var ordinaryLink = new XElement(sourceLink);
                    laterLink.SetAttributeValue(W + "anchor", "LaterThird");
                    ordinaryLink.SetAttributeValue(W + "anchor", "OrdinaryThird");
                    sourceLink.AddAfterSelf(laterLink, ordinaryLink);
                }
                else if (variant == "inside")
                {
                    XElement paragraph = second.Element(W + "p")!;
                    paragraph.Element(W + "r")!.AddBeforeSelf(start);
                    paragraph.Add(end);
                }
                else if (variant is "inside-first" or "inside-first-column")
                {
                    if (variant == "inside-first-column")
                    {
                        start.SetAttributeValue(W + "colFirst", spanned ? 2 : 0);
                        start.SetAttributeValue(W + "colLast", spanned ? 2 : 0);
                    }
                    first.Element(W + "p")!.Add(start, end);
                }
                else
                {
                    int column = variant == "unoccupied" ? gridSize : spanned ? 4 : 1;
                    start.SetAttributeValue(W + "colFirst", column);
                    start.SetAttributeValue(W + "colLast", column);
                    if (variant is "inline-start" or "trailing-start" or "inline-closed" or "trailing-closed")
                    {
                        XElement owner = variant.StartsWith("inline", StringComparison.Ordinal)
                            ? first.Element(W + "p")! : first;
                        owner.Add(start);
                        if (variant.EndsWith("closed", StringComparison.Ordinal)) owner.Add(end);
                        else row.AddAfterSelf(end);
                    }
                    else if (variant == "between-cell-start")
                    {
                        second.AddBeforeSelf(start);
                        row.AddAfterSelf(end);
                    }
                    else if (variant is "between-cell-end" or "end-after-paragraph" or "end-in-paragraph")
                    {
                        row.AddBeforeSelf(start);
                        if (variant == "between-cell-end") second.AddBeforeSelf(end);
                        else if (variant == "end-after-paragraph") first.Add(end);
                        else first.Element(W + "p")!.Add(end);
                    }
                    else if (variant.StartsWith("row-boundary", StringComparison.Ordinal))
                    {
                        row.AddBeforeSelf(start);
                        if (variant == "row-boundary-closed") start.AddAfterSelf(end);
                        else if (variant == "row-boundary-closed-before-cell") first.AddBeforeSelf(end);
                        else if (variant == "row-boundary-closed-in-cell") first.Element(W + "p")!.AddBeforeSelf(end);
                        else row.AddAfterSelf(end);
                    }
                    else if (variant is "inside-row" or "inside-row-closed")
                    {
                        first.AddBeforeSelf(start);
                        if (variant == "inside-row-closed") start.AddAfterSelf(end);
                        else row.AddAfterSelf(end);
                    }
                    else
                    {
                        first.Element(W + "p")!.AddBeforeSelf(variant == "wrapped"
                            ? new XElement(W + "sdt", new XElement(W + "sdtContent", start)) : start);
                        if (variant == "closed") start.AddAfterSelf(end);
                        else row.Add(end);
                    }
                }
                target.ReplaceWith(table);
            }, path => pages = Render(path, mode));
            return pages;
        }
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
