using System.IO.Compression;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxSectionScopedInlineBoxesKeepTheirActualFontSources()
    {
        foreach (string story in new[] { "header", "footer" })
        {
            var source = ReadSectionScopedBox(story, OoxPdfDocxMarkupMode.Final, sectioned: true);
            var section = source.BodyElements.OfType<DocxSectionBreakElement>().Single();
            var parent = DocxBlockTraversal.EnumerateStaticStoryParagraphs(section.PageSettings)
                .Single(paragraph => paragraph.InlineTextBoxes.Count != 0);
            var child = DocxBlockTraversal.EnumerateBodyParagraphs(parent.InlineTextBoxes.Single().BodyElements).Single();
            var run = child.Runs.Single();
            TestAssert.Equal("Consolas", run.FontFamily);
            TestAssert.True(DocxBlockTraversal.EnumerateSourceParagraphs(source).Any(paragraph => ReferenceEquals(child, paragraph)),
                "An earlier section's child story retains its actual source owner: " + story);
            var plan = DocxFontPlan.Create(source, new SectionBoxResolver(), default);
            TestAssert.True(plan.Runs.Any(item => ReferenceEquals(run, item.Run) && item.ResolvedFamily == "DejaVu Math TeX Gyre"),
                "The child font is prepared even when a later static story replaces the document's variant key: " + story);
        }
    }

    public static void DocxSectionScopedInlineBoxesKeepVisibleLabelsInBothViews()
    {
        foreach (string story in new[] { "header", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool split in new[] { false, true })
        {
            var renderer = new DocxRenderer(new SectionBoxResolver(), mode, CachedMathGeometry(mode));
            var control = renderer.RenderPagesAsync(ReadSectionScopedBox(story, mode, sectioned: false, split),
                null, default).ToBlockingEnumerable().Single();
            var pages = renderer.RenderPagesAsync(ReadSectionScopedBox(story, mode, sectioned: true, split),
                null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(2, pages.Length);
            string label = "SECTION " + story.ToUpperInvariant() + " BOX LABEL";
            foreach (var page in new[] { control, pages[0] })
            {
                AssertSourceLabelsPaint(page, mode, [label], story + "/split=" + split);
                string actualFaceText = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true)
                    .Where(show => show.Font.Font.FamilyName == "DejaVu Math TeX Gyre").Select(show => show.Text));
                TestAssert.Contains(label, actualFaceText);
            }
            string lastPage = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(pages[1], true).Select(show => show.Text));
            TestAssert.Contains("SECOND " + story.ToUpperInvariant() + " PLAIN", lastPage);
            TestAssert.True(!lastPage.Contains(label, StringComparison.Ordinal), "Section-local content must not leak into the final section.");
        }
    }

    private static DocxDocument ReadSectionScopedBox(string story, OoxPdfDocxMarkupMode mode, bool sectioned, bool split = false)
    {
        string seed = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "InlineBoxMath",
            "inline-box-math-" + story + ".docx"));
        using var memory = new MemoryStream();
        memory.Write(File.ReadAllBytes(seed));
        using (var zip = new ZipArchive(memory, ZipArchiveMode.Update, leaveOpen: true))
        {
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            string part = "word/inline-box-" + story + ".xml";
            var staticXml = Read(part);
            string label = "SECTION " + story.ToUpperInvariant() + " BOX LABEL";
            var run = new XElement(w + "r", new XElement(w + "rPr",
                new XElement(w + "rFonts", new XAttribute(w + "ascii", "Consolas"), new XAttribute(w + "hAnsi", "Consolas")),
                new XElement(w + "sz", new XAttribute(w + "val", "24"))), new XElement(w + "t", label));
            var paragraph = new XElement(w + "p", run);
            if (split)
            {
                var second = new XElement(run);
                run.Element(w + "t")!.Value = "SECTION ";
                second.Element(w + "t")!.Value = label[8..];
                run.Element(w + "t")!.SetAttributeValue(XNamespace.Xml + "space", "preserve");
                paragraph.Add(second);
            }
            staticXml.Descendants(w + "txbxContent").Single().ReplaceNodes(paragraph);
            Write(part, staticXml);
            if (sectioned)
            {
                var document = Read("word/document.xml");
                var body = document.Root!.Element(w + "body")!;
                var final = body.Element(w + "sectPr")!;
                final.AddBeforeSelf(new XElement(w + "p", new XElement(w + "pPr", new XElement(final)),
                    new XElement(w + "r", new XElement(w + "t", "End first section."))),
                    new XElement(w + "p", new XElement(w + "r", new XElement(w + "t", "Second section body."))));
                final.Element(w + story + "Reference")!.SetAttributeValue(r + "id", "rSecondStory");
                Write("word/document.xml", document);
                var relationships = Read("word/_rels/document.xml.rels");
                relationships.Root!.Add(new XElement(relationships.Root.Name.Namespace + "Relationship",
                    new XAttribute("Id", "rSecondStory"),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" + story),
                    new XAttribute("Target", "second-" + story + ".xml")));
                Write("word/_rels/document.xml.rels", relationships);
                string secondPart = "word/second-" + story + ".xml";
                Write(secondPart, new XDocument(new XElement(w + (story == "header" ? "hdr" : "ftr"),
                    new XElement(w + "p", new XElement(w + "r", new XElement(w + "t", "SECOND " + story.ToUpperInvariant() + " PLAIN"))))));
                var types = Read("[Content_Types].xml");
                types.Root!.Add(new XElement(types.Root.Name.Namespace + "Override", new XAttribute("PartName", "/" + secondPart),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml." + story + "+xml")));
                Write("[Content_Types].xml", types);
            }

            XDocument Read(string name)
            {
                using var stream = zip.GetEntry(name)!.Open();
                return XDocument.Load(stream);
            }
            void Write(string name, XDocument xml)
            {
                zip.GetEntry(name)?.Delete();
                using var stream = zip.CreateEntry(name).Open();
                xml.Save(stream);
            }
        }
        memory.Position = 0;
        using var package = OoxPackage.Open(memory, default);
        return new DocxReader().Read(package, null, default, markupMode: mode);
    }

    private sealed class SectionBoxResolver : IFontResolver
    {
        private readonly PortableRenderingTests.PinnedResolver body = new();
        private readonly PortableRenderingTests.PinnedResolver box = new(mathFont: true);
        public FontFaceResolution Resolve(FontRequest request) =>
            (request.FamilyName == "Consolas" ? box : body).Resolve(request);
    }
}
