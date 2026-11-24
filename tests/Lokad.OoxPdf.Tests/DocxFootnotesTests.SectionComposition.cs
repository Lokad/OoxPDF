using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxFootnotesTests
{
    public static void DocxSectionStartsKeepPictureAndTableContentOnOwnedPages() =>
        CheckComposedSectionStarts(["picture", "table"]);

    public static void DocxSectionStartsKeepCommentAndMixedContentOnOwnedPages() =>
        CheckComposedSectionStarts(["comment", "mixed"]);

    public static void DocxNextColumnSectionStartsKeepTextInItsColumnAndPage()
    {
        XNamespace w = NoteWord;
        const string first = "FIRST SECTION WORDS";
        const string second = "SECOND SECTION WORDS";
        const string third = "THIRD SECTION WORDS";
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var samePage = Render(afterColumns: 2, lastColumn: false);
            TestAssert.Equal(1, samePage.Length);
            var firstGlyphs = Find(samePage[0], first);
            var secondGlyphs = Find(samePage[0], second);
            TestAssert.True(secondGlyphs.X - firstGlyphs.X > 150d, "The following section starts in the next column.");
            TestAssert.True(Math.Abs(secondGlyphs.Y - firstGlyphs.Y) < 0.01d, "The next column starts at the shared top.");
            TestAssert.Equal(1, samePage[0].Images.Count);

            var nextPage = Render(2, false, "nextPage", paperWidth: 540);
            TestAssert.Equal(2, nextPage.Length);
            TestAssert.Equal(secondGlyphs.X, Find(nextPage[0], second).X);
            TestAssert.Equal(firstGlyphs.X, Find(nextPage[1], third).X);

            foreach (var (afterColumns, lastColumn) in new[] { (2, true), (1, false) })
            {
                var paged = Render(afterColumns, lastColumn);
                TestAssert.Equal(2, paged.Length);
                TestAssert.Equal(lastColumn ? secondGlyphs.X : firstGlyphs.X, Find(paged[0], first).X);
                TestAssert.Equal(firstGlyphs.X, Find(paged[1], second).X);
                TestAssert.Equal(1, paged.Sum(page => page.Images.Count));
            }

            var continuous = Render(2, false, "continuous");
            TestAssert.Equal(1, continuous.Length);
            TestAssert.Equal(secondGlyphs.X, Find(continuous[0], third).X);
            TestAssert.Equal(0, continuous[0].Annotations.Single().Destination?.PageIndex);

            var differentPaper = Render(2, false, "continuous", paperWidth: 540);
            TestAssert.Equal(2, differentPaper.Length);
            TestAssert.Equal(612d, differentPaper[0].Width);
            TestAssert.Equal(540d, differentPaper[1].Width);
            TestAssert.Equal(1, differentPaper[0].Images.Count);
            TestAssert.Equal(secondGlyphs, Find(differentPaper[0], second));
            TestAssert.Equal(Find(nextPage[1], third), Find(differentPaper[1], third));
            TestAssert.Equal(1, differentPaper[0].Annotations.Single().Destination?.PageIndex);

            PdfPage[] Render(int afterColumns, bool lastColumn, string? thirdStart = null, int paperWidth = 612)
            {
                DocxDocument source = ReadComposedSections("nextPage", "nextColumn", "picture", mode, editBody: body =>
                {
                    XElement closingSection = body.Descendants(w + "sectPr").First();
                    XElement finalSection = body.Element(w + "sectPr")!;
                    closingSection.Add(SectionColumns(2));
                    finalSection.Add(SectionColumns(afterColumns));
                    if (lastColumn)
                    {
                        body.Descendants(w + "r").First().AddFirst(new XElement(w + "br", new XAttribute(w + "type", "column")));
                    }
                    if (thirdStart is not null)
                    {
                        const string target = "PublicThirdSection";
                        body.Elements(w + "p").First().Add(new XElement(w + "hyperlink", new XAttribute(w + "anchor", target),
                            new XElement(w + "r", new XElement(w + "t", "Follow"))));
                        body.Elements(w + "p").Last().AddFirst(new XElement(w + "pPr", new XElement(finalSection)));
                        finalSection.Element(w + "type")!.SetAttributeValue(w + "val", thirdStart);
                        finalSection.Element(w + "pgSz")!.SetAttributeValue(w + "w", paperWidth * 20);
                        finalSection.AddBeforeSelf(new XElement(w + "p",
                            new XElement(w + "bookmarkStart", new XAttribute(w + "id", "1"), new XAttribute(w + "name", target)),
                            new XElement(w + "r", new XElement(w + "t", third)),
                            new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "1"))));
                    }
                });
                return SectionCompositionRenderer(mode).RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            }
        }

        static (string Text, double Size, double X, double Y) Find(PdfPage page, string label) =>
            PhysicalSectionGlyphs(page).Single(glyphs => glyphs.Text == label);
    }

    public static void DocxNextColumnFallbackRestartsOwnedPageNumber()
    {
        XNamespace w = NoteWord;
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (var (columns, exhausted) in new[] { (2, true), (1, false) })
        {
            var source = ReadComposedSections("nextPage", "nextColumn", "picture", mode, editBody: body =>
            {
                var first = body.Elements(w + "p").First();
                var second = body.Elements(w + "p").Last();
                var final = body.Element(w + "sectPr")!;
                first.Descendants(w + "sectPr").Single().Add(SectionColumns(2));
                final.Add(new XElement(w + "pgNumType", new XAttribute(w + "start", 40)), SectionColumns(columns));
                if (exhausted)
                    first.Descendants(w + "r").First().AddFirst(new XElement(w + "br", new XAttribute(w + "type", "column")));
                const string target = "NumberedSection";
                first.Add(new XElement(w + "hyperlink", new XAttribute(w + "anchor", target), Run("Follow")));
                second.AddFirst(new XElement(w + "bookmarkStart", new XAttribute(w + "id", 1), new XAttribute(w + "name", target)));
                second.Add(Run(" BODY PAGE "), Field("PAGE"), Run(" OF "), Field("NUMPAGES"),
                    new XElement(w + "bookmarkEnd", new XAttribute(w + "id", 1)));
            });
            var pages = SectionCompositionRenderer(mode).RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(2, pages.Length);
            TestAssert.Equal(1, pages[0].Images.Count);
            TestAssert.Equal(1, pages[0].Annotations.Single().Destination?.PageIndex);
            var text = DocxPageFootnoteNumberingTests.DecodeText(pages[1]);
            TestAssert.True(Regex.Replace(text, @"\s+", "").Contains("BODYPAGE40OF2", StringComparison.Ordinal),
                $"Expected BODY PAGE 40 OF 2 on the owned page, got '{text}'.");
        }

        XElement Run(string text) => new(w + "r", new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), text));
        XElement Field(string instruction) => new(w + "fldSimple", new XAttribute(w + "instr", instruction), Run("900"));
    }

    public static void DocxSharedSectionPagesAdoptIncomingGridAndDefaultStories()
    {
        XNamespace w = NoteWord;
        const string shared = "SHARED SECTION BODY";
        const string headerText = "CONTINUOUS SECTION HEADER";
        const string footerText = "CONTINUOUS FOOTER ";
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadComposedSections("nextPage", "nextColumn", "picture", mode, editBody: body =>
            {
                var first = body.Elements(w + "p").First();
                var second = body.Elements(w + "p").Last();
                var final = body.Element(w + "sectPr")!;
                first.Descendants(w + "sectPr").Single().Add(SectionColumns(2));
                final.Add(SectionColumns(2));
                second.AddFirst(new XElement(w + "pPr", new XElement(final)));
                final.Element(w + "type")!.SetAttributeValue(w + "val", "continuous");
                final.Element(w + "cols")!.SetAttributeValue(w + "num", 1);
                final.Element(w + "pgMar")!.SetAttributeValue(w + "left", 1920);
                final.AddBeforeSelf(Paragraph(shared));
                for (int page = 2; page <= 3; page++)
                {
                    var continuation = Paragraph($"FINAL SECTION PAGE {page}");
                    continuation.AddFirst(new XElement(w + "pPr", new XElement(w + "pageBreakBefore")));
                    if (page == 3)
                    {
                        continuation.Element(w + "r")!.AddBeforeSelf(new XElement(w + "bookmarkStart",
                            new XAttribute(w + "id", 1), new XAttribute(w + "name", "FinalBody")));
                        continuation.Add(new XElement(w + "bookmarkEnd", new XAttribute(w + "id", 1)));
                    }
                    final.AddBeforeSelf(continuation);
                }
                first.Add(new XElement(w + "hyperlink", new XAttribute(w + "anchor", "FinalBody"),
                    new XElement(w + "r", new XElement(w + "t", "Follow"))));
            });
            var header = DocxTests.CreateDocxLayoutParagraph(headerText, 12d, 14d);
            var footer = DocxTests.CreateDocxLayoutParagraph(footerText, 12d, 14d);
            var run = footer.Runs.Single();
            footer = footer with { Runs = [run, run with { Text = "{PAGE}", FieldKind = DocxFieldKind.Page },
                run with { Text = " OF " }, run with { Text = "{NUMPAGES}", FieldKind = DocxFieldKind.NumPages }] };
            var settings = source.PageSettings with
            {
                HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>> { ["default"] = [header] },
                FooterParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>> { ["default"] = [footer] }
            };
            var document = source with { PageSettings = settings, FinalSectionBreak = source.FinalSectionBreak! with { PageSettings = settings } };
            var pages = SectionCompositionRenderer(mode).RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(3, pages.Length);
            TestAssert.Equal(1, pages[0].Images.Count);
            var sharedPage = PhysicalSectionGlyphs(pages[0]);
            TestAssert.Equal(sharedPage.Single(g => g.Text == "SECOND SECTION WORDS").X, sharedPage.Single(g => g.Text == shared).X);
            TestAssert.DoesNotContain(headerText, DocxPageFootnoteNumberingTests.DecodeText(pages[0]));
            TestAssert.Equal(2, pages[0].Annotations.Single().Destination?.PageIndex);
            for (int index = 1; index < pages.Length; index++)
            {
                string text = DocxPageFootnoteNumberingTests.DecodeText(pages[index]);
                TestAssert.Contains(headerText, text);
                TestAssert.Contains($"{footerText}{index + 1} OF 3", text);
                var body = PhysicalSectionGlyphs(pages[index]).Single(g => g.Text == $"FINAL SECTION PAGE {index + 1}");
                TestAssert.True(Math.Abs(body.X - 96d) < .01d, "Later pages use the incoming section's authored left margin.");
            }
        }

        XElement Paragraph(string text) => new(w + "p", new XElement(w + "r", new XElement(w + "t", text)));
    }

    private static XElement SectionColumns(int count) => new(NoteWord + "cols",
        new XAttribute(NoteWord + "num", count), new XAttribute(NoteWord + "space", 720));

    private static void CheckComposedSectionStarts(string[] contents)
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (var (first, following) in new (string?, string?)[]
            { ("nextPage", "continuous"), ("continuous", "nextPage"), (null, "continuous"), ("continuous", null) })
        foreach (string content in contents)
        {
            var renderer = SectionCompositionRenderer(mode);
            var source = ReadComposedSections(first, following, content, mode);
            var control = ReadComposedSections(following, following, content, mode);
            var expected = renderer.RenderPagesAsync(control, null, default).ToBlockingEnumerable().ToArray();
            var actual = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(following == "continuous" ? 1 : 2, expected.Length);
            TestAssert.True(actual.Length == expected.Length,
                $"{content}/{mode}: the following section owns the boundary, including a source-matched first-section control.");
            for (int page = 0; page < actual.Length; page++)
            {
                var a = PhysicalSectionGlyphs(actual[page]); var b = PhysicalSectionGlyphs(expected[page]);
                TestAssert.Equal(b.Length, a.Length);
                for (int glyph = 0; glyph < a.Length; glyph++) TestAssert.Equal(b[glyph], a[glyph]);
                TestAssert.Equal(expected[page].Annotations.Count, actual[page].Annotations.Count);
                TestAssert.Equal(Regex.Matches(expected[page].Content, @"/Im\d+ Do").Count,
                    Regex.Matches(actual[page].Content, @"/Im\d+ Do").Count);
            }
            string text = string.Concat(actual.SelectMany(PhysicalSectionGlyphs).Select(x => x.Text));
            TestAssert.Contains("FIRST SECTION WORDS", text); TestAssert.Contains("SECOND SECTION WORDS", text);
            if (content == "table") TestAssert.Contains("TABLE CELL WORDS", text);
            if (content is "picture" or "mixed")
                TestAssert.Equal(1, actual.Sum(page => Regex.Matches(page.Content, @"/Im\d+ Do").Count));
            if (mode == OoxPdfDocxMarkupMode.Final)
                TestAssert.True(!text.Contains("A public comment.", StringComparison.Ordinal), "Final comments remain unpainted.");
        }
    }

    public static void DocxSectionStartOwnershipKeepsNoteAndFrameAdmissionFallbacks()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string content in new[] { "plain", "picture", "comment" })
        {
            var source = ReadComposedSections("continuous", "continuous", content, mode, notes: true);
            var pages = SectionCompositionRenderer(mode).RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(content == "plain" ? 2 : 1, pages.Length);
            TestAssert.Contains("PUBLIC NOTE", string.Concat(pages.SelectMany(PhysicalSectionGlyphs).Select(x => x.Text)));
            TestAssert.Equal(1, pages.SelectMany(page => page.Annotations).Count(link => link.IsDestination));
        }
        foreach (string fallback in new[] { "margins", "columns", "header", "mirror", "unknown-final" })
        {
            var source = ReadComposedSections("nextPage", "continuous", "picture", OoxPdfDocxMarkupMode.Final,
                fallback: fallback);
            if (fallback == "mirror") source = source with { Settings = source.Settings with { MirrorMargins = true } };
            if (fallback == "unknown-final") source = source with { FinalSectionBreak = null };
            var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            TestAssert.True(engine.Create(source, new DocxTests.FamilyWidthTextMeasurer(), default).Pages.Count == 2,
                "The existing frame fallback remains explicit: " + fallback);
        }
        foreach (double scale in new[] { .5d, .75d, 1d })
        foreach (string following in new[] { "continuous", "nextPage" })
        {
            var source = ReadComposedSections(following == "continuous" ? "nextPage" : "continuous", following,
                "picture", OoxPdfDocxMarkupMode.AllMarkup);
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, wordCompatiblePrintScale: scale)
                .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default);
            TestAssert.Equal(following == "continuous" ? 1 : 2, layout.Pages.Count);
        }
    }

    private static DocxRenderer SectionCompositionRenderer(OoxPdfDocxMarkupMode mode) => new(new TestFaceFontResolver(), mode,
        mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
            : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);

    private static DocxDocument ReadComposedSections(string? first, string? following, string content,
        OoxPdfDocxMarkupMode mode, bool notes = false, string? fallback = null, Action<XElement>? editBody = null)
    {
        XNamespace w = NoteWord;
        XNamespace rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        XNamespace pr = "http://schemas.openxmlformats.org/package/2006/relationships";
        XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
        var types = new XElement(ct + "Types",
            new XElement(ct + "Default", new XAttribute("Extension", "rels"), new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")),
            new XElement(ct + "Default", new XAttribute("Extension", "xml"), new XAttribute("ContentType", "application/xml")),
            new XElement(ct + "Override", new XAttribute("PartName", "/word/document.xml"),
                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")));
        var relationships = new XElement(pr + "Relationships");
        var parts = new Dictionary<string, byte[]>();
        void AddPart(string name, string kind, string id, byte[] data)
        {
            parts["word/" + name] = data;
            relationships.Add(new XElement(pr + "Relationship", new XAttribute("Id", id),
                new XAttribute("Type", rel.NamespaceName + "/" + kind), new XAttribute("Target", name)));
            if (kind != "image") types.Add(new XElement(ct + "Override", new XAttribute("PartName", "/word/" + name),
                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml." + kind + "+xml")));
        }
        XElement Section(string? kind) => new(w + "sectPr",
            kind is null ? null : new XElement(w + "type", new XAttribute(w + "val", kind)),
            new XElement(w + "pgSz", new XAttribute(w + "w", "12240"), new XAttribute(w + "h", "15840")),
            new XElement(w + "pgMar", new XAttribute(w + "top", "1440"), new XAttribute(w + "bottom", "1440"),
                new XAttribute(w + "left", "1440"), new XAttribute(w + "right", "1440")));
        XElement Paragraph(string text) => new(w + "p", new XElement(w + "r", new XElement(w + "t", text)));
        var firstSection = Section(first); var finalSection = Section(following);
        var firstParagraph = Paragraph("FIRST SECTION WORDS");
        firstParagraph.AddFirst(new XElement(w + "pPr", firstSection));
        var body = new XElement(w + "body", firstParagraph, Paragraph("SECOND SECTION WORDS"), finalSection);
        if (content is "picture" or "mixed")
        {
            AddPart("media/picture.png", "image", "rPicture", TestFixtures.CreateTinyStoredRgbPng());
            types.Add(new XElement(ct + "Default", new XAttribute("Extension", "png"), new XAttribute("ContentType", "image/png")));
            firstParagraph.Add(XElement.Parse($$"""
                <w:r xmlns:w="{{w}}" xmlns:r="{{rel}}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:drawing><wp:inline><wp:extent cx="12700" cy="12700"/><wp:docPr id="1" name="Public picture"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="Public picture"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rPicture"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="12700" cy="12700"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>
                """));
        }
        if (content is "comment" or "mixed")
        {
            firstParagraph.Element(w + "r")!.AddBeforeSelf(new XElement(w + "commentRangeStart", new XAttribute(w + "id", "0")));
            firstParagraph.Add(new XElement(w + "commentRangeEnd", new XAttribute(w + "id", "0")),
                new XElement(w + "r", new XElement(w + "commentReference", new XAttribute(w + "id", "0"))));
            AddPart("comments.xml", "comments", "rComment", Encoding.UTF8.GetBytes(new XElement(w + "comments",
                new XElement(w + "comment", new XAttribute(w + "id", "0"), new XAttribute(w + "author", "Public Reviewer"),
                    Paragraph("A public comment."))).ToString()));
        }
        if (content == "table")
        {
            firstSection.Remove();
            firstParagraph.AddAfterSelf(new XElement(w + "tbl", new XElement(w + "tr", new XElement(w + "tc", Paragraph("TABLE CELL WORDS")))),
                new XElement(w + "p", new XElement(w + "pPr", firstSection)));
        }
        if (notes)
        {
            firstParagraph.Add(new XElement(w + "r", new XElement(w + "footnoteReference", new XAttribute(w + "id", "1"))));
            AddPart("footnotes.xml", "footnotes", "rNote", Encoding.UTF8.GetBytes(new XElement(w + "footnotes",
                new XElement(w + "footnote", new XAttribute(w + "id", "1"), Paragraph("PUBLIC NOTE"))).ToString()));
        }
        if (fallback == "margins") finalSection.Element(w + "pgMar")!.SetAttributeValue(w + "left", "1800");
        if (fallback == "columns") finalSection.Add(new XElement(w + "cols", new XAttribute(w + "num", "2")));
        if (fallback == "header")
        {
            firstSection.AddFirst(new XElement(w + "headerReference", new XAttribute(w + "type", "default"), new XAttribute(rel + "id", "rHeader")));
            AddPart("header.xml", "header", "rHeader", Encoding.UTF8.GetBytes(new XElement(w + "hdr", Paragraph("PUBLIC HEADER")).ToString()));
        }
        editBody?.Invoke(body);
        parts["[Content_Types].xml"] = Encoding.UTF8.GetBytes(types.ToString());
        parts["_rels/.rels"] = Encoding.UTF8.GetBytes(new XElement(pr + "Relationships", new XElement(pr + "Relationship",
            new XAttribute("Id", "rDocument"), new XAttribute("Type", rel.NamespaceName + "/officeDocument"), new XAttribute("Target", "word/document.xml"))).ToString());
        parts["word/_rels/document.xml.rels"] = Encoding.UTF8.GetBytes(relationships.ToString());
        parts["word/document.xml"] = Encoding.UTF8.GetBytes(new XElement(w + "document", body).ToString());
        string input = TestFixtures.WriteTempPackage(".docx", parts);
        try { return DocxTests.ReadDocx(input, mode); } finally { File.Delete(input); }
    }
}
