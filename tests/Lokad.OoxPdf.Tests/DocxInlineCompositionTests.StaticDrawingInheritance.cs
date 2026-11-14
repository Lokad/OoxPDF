using System.IO.Compression;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInheritedStaticPicturesKeepContentInBothViews()
    {
        foreach (string story in new[] { "header", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode, CachedMathGeometry(mode));
            var control = renderer.RenderPagesAsync(ReadInheritedStaticPicture(story, mode, "explicit"), null, default)
                .ToBlockingEnumerable().ToArray();
            var pages = renderer.RenderPagesAsync(ReadInheritedStaticPicture(story, mode), null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(3, pages.Length); TestAssert.Equal(control.Length, pages.Length);
            for (int index = 0; index < pages.Length; index++)
            {
                TestAssert.Equal(1, CountInheritedStaticPictures(control[index]));
                TestAssert.Equal(1, CountInheritedStaticPictures(pages[index]));
                TestAssert.Equal(DocxPageFootnoteNumberingTests.DecodeText(control[index]), DocxPageFootnoteNumberingTests.DecodeText(pages[index]));
            }
        }
    }

    public static void DocxExplicitStaticReplacementClearsInheritedPictures()
    {
        foreach (string story in new[] { "header", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadInheritedStaticPicture(story, mode, "replacement");
            var drawings = story == "header" ? source.PageSettings.HeaderFloatingDrawingsByType : source.PageSettings.FooterFloatingDrawingsByType;
            TestAssert.True(drawings.ContainsKey("default") && drawings["default"].Count == 0,
                "An explicit story owns an empty drawing list, not an absent inheritance key.");
            var pages = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode, CachedMathGeometry(mode))
                .RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(3, pages.Length); TestAssert.Equal(1, CountInheritedStaticPictures(pages[0]));
            foreach (var page in pages.Skip(1))
            {
                TestAssert.Equal(0, CountInheritedStaticPictures(page));
                TestAssert.Contains("REPLACEMENT " + story.ToUpperInvariant(), DocxPageFootnoteNumberingTests.DecodeText(page));
            }
        }
    }

    public static void DocxInheritedStaticDrawingVariantsStayIndependent()
    {
        foreach (string story in new[] { "header", "footer" })
        foreach (string variant in new[] { "first", "even" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ActivateInheritedDrawingVariant(ReadInheritedStaticPicture(story, mode), story, variant);
            var pages = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode, CachedMathGeometry(mode))
                .RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(3, pages.Length);
            TestAssert.Equal(0, CountInheritedStaticPictures(pages[0]));
            TestAssert.Equal(1, CountInheritedStaticPictures(pages[1]));
            TestAssert.Equal(0, CountInheritedStaticPictures(pages[2]));
        }
    }

    public static void DocxInheritedStaticTextBoxesKeepActualFontsAndRefusals()
    {
        foreach (string story in new[] { "header", "footer" })
        {
            var source = ActivateInheritedDrawingVariant(ReadInheritedStaticPicture(story, OoxPdfDocxMarkupMode.Final), story, "first");
            var earlier = source.BodyElements.OfType<DocxSectionBreakElement>().Single();
            var settings = earlier.PageSettings;
            var drawings = story == "header" ? settings.HeaderFloatingDrawingsByType : settings.FooterFloatingDrawingsByType;
            var paragraph = DocxTests.CreateDocxLayoutParagraph("INHERITED BOX", 12d, 12d);
            paragraph = paragraph with { Runs = paragraph.Runs.Select(run => run with
            { FontFamily = "Consolas", Fonts = new DocxRunFonts("Consolas", null, null, null, null, null, null, null) }).ToArray() };
            var box = drawings["first"].Single() with { Image = null, ImageRelationshipId = null,
                ExtentCxValue = "1828800", ExtentCyValue = "457200", TextBoxBodyElements = [new DocxParagraphElement(paragraph)] };
            var boxes = new Dictionary<string, IReadOnlyList<DocxFloatingDrawing>>(StringComparer.OrdinalIgnoreCase) { ["first"] = [box] };
            settings = story == "header" ? settings with { HeaderFloatingDrawingsByType = boxes } : settings with { FooterFloatingDrawingsByType = boxes };
            var replacement = earlier with { PageSettings = settings };
            source = source with { BodyElements = source.BodyElements.Select(item => ReferenceEquals(item, earlier) ? replacement : item).ToArray() };
            var renderer = new DocxRenderer(new SectionBoxResolver(), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(3, pages.Length);
            string painted = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(pages[1], true)
                .Where(show => show.Font.Font.FamilyName == "DejaVu Math TeX Gyre").Select(show => show.Text));
            TestAssert.Contains("INHERITED BOX", painted);
            TestAssert.DoesNotContain("INHERITED BOX", DocxPageFootnoteNumberingTests.DecodeText(pages[0]));
            TestAssert.DoesNotContain("INHERITED BOX", DocxPageFootnoteNumberingTests.DecodeText(pages[2]));
            var refused = new StaticFontRequestResolver("Consolas");
            TestAssert.Throws<InvalidOperationException>(() => new DocxRenderer(refused, OoxPdfDocxMarkupMode.Final,
                OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray());
            TestAssert.True(refused.Requests.Contains("Consolas"), "Only the later enabled inherited textbox requires this font.");
            using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(source, null, cancellation.Token).ToBlockingEnumerable().ToArray());
        }
    }

    private static int CountInheritedStaticPictures(PdfPage page) => DocxNoteClippingTests.ReadPhysicalPaintFrames(page).Count(frame =>
        frame.Block.StartsWith("/", StringComparison.Ordinal) && Math.Abs(frame.Transform.M11 - 20d) < .001d &&
        Math.Abs(frame.Transform.M22 - 20d) < .001d && Math.Abs(frame.Transform.M12) < .001d && Math.Abs(frame.Transform.M21) < .001d &&
        Math.Abs(frame.Transform.M31 - 420d) < .001d && Math.Abs(frame.Transform.M32 - (page.Height - 44d)) < .001d);

    private static DocxDocument ActivateInheritedDrawingVariant(DocxDocument source, string story, string variant)
    {
        var earlier = source.BodyElements.OfType<DocxSectionBreakElement>().Single();
        var replacement = earlier with { PageSettings = StaticVariantSettings(earlier.PageSettings, story, variant, enabled: false) };
        return WithStaticVariantSettings(source, StaticVariantSettings(source.PageSettings, story, variant, enabled: true)) with
        { BodyElements = source.BodyElements.Select(item => ReferenceEquals(item, earlier) ? replacement : item).ToArray() };
    }

    private static DocxDocument ReadInheritedStaticPicture(string story, OoxPdfDocxMarkupMode mode, string binding = "inherited")
    {
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes",
            "static-" + story + "-picture-inherited.docx"));
        using var memory = new MemoryStream(); memory.Write(File.ReadAllBytes(path));
        if (binding != "inherited")
        {
            using var zip = new ZipArchive(memory, ZipArchiveMode.Update, leaveOpen: true);
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            var document = Read("word/document.xml");
            document.Root!.Element(w + "body")!.Elements(w + "sectPr").Single().Add(new XElement(w + story + "Reference",
                new XAttribute(w + "type", "default"), new XAttribute(r + "id", binding == "explicit" ? story : "replacement")));
            Write("word/document.xml", document);
            if (binding == "replacement")
            {
                string name = "replacement-" + story + ".xml";
                var relationships = Read("word/_rels/document.xml.rels");
                relationships.Root!.Add(new XElement(relationships.Root.Name.Namespace + "Relationship", new XAttribute("Id", "replacement"),
                    new XAttribute("Type", r.NamespaceName + "/" + story), new XAttribute("Target", name)));
                Write("word/_rels/document.xml.rels", relationships);
                var types = Read("[Content_Types].xml");
                types.Root!.Add(new XElement(types.Root.Name.Namespace + "Override", new XAttribute("PartName", "/word/" + name),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml." + story + "+xml")));
                Write("[Content_Types].xml", types);
                Write("word/" + name, new XDocument(new XElement(w + (story == "header" ? "hdr" : "ftr"), new XElement(w + "p",
                    new XElement(w + "r", new XElement(w + "t", "REPLACEMENT " + story.ToUpperInvariant()))))));
            }
            XDocument Read(string name) { using var stream = zip.GetEntry(name)!.Open(); return XDocument.Load(stream); }
            void Write(string name, XDocument xml) { zip.GetEntry(name)?.Delete(); using var stream = zip.CreateEntry(name).Open(); xml.Save(stream); }
        }
        memory.Position = 0;
        return new DocxReader().Read(OoxPackage.Open(memory, default), null, default, mode);
    }
}
