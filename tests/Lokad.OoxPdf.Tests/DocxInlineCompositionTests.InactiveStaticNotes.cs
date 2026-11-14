using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInactiveStaticTemplatesKeepMixedNoteCounterPaints()
    {
        foreach (string fixture in new[] { "math-page-notes-arial-11-zero", "cell-page-notes-arial-11-zero" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode, CachedMathGeometry(mode));
            var control = renderer.RenderPagesAsync(ReadInactiveNoteTemplate(fixture, mode), null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(2, control.Length);
            foreach (string story in new[] { "header", "footer" })
            foreach (string variant in new[] { "first", "even" })
            {
                var source = ReadInactiveNoteTemplate(fixture, mode, story, variant);
                var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
                TestAssert.Equal(control.Length, pages.Length);
                for (int index = 0; index < pages.Length; index++)
                {
                    string Markers(PdfPage page)
                    {
                        var mathFonts = page.Fonts.Where(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal))
                            .Select(font => font.Font).ToHashSet();
                        return string.Join('|', DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true)
                            .Where(show => show.Size < 9d && !mathFonts.Contains(show.Font) &&
                                show.Text.Length != 0 && show.Text.All(char.IsAsciiDigit)).Select(show => show.Text));
                    }
                    TestAssert.Equal("09|10|09|10", Markers(control[index]));
                    TestAssert.Equal(Markers(control[index]), Markers(pages[index]));
                    TestAssert.Equal(DocxPageFootnoteNumberingTests.DecodeText(control[index]), DocxPageFootnoteNumberingTests.DecodeText(pages[index]));
                }
            }
        }
    }

    public static void DocxInactiveStaticTemplatesKeepFinalEndnotePageFrames()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode, CachedMathGeometry(mode));
            var control = renderer.RenderPagesAsync(ReadInactiveNoteTemplate("endnote-empty-final-section", mode), null, default)
                .ToBlockingEnumerable().ToArray();
            foreach (string story in new[] { "header", "footer" })
            foreach (string variant in new[] { "first", "even" })
            foreach (bool rich in new[] { false, true })
            {
                var source = ReadInactiveNoteTemplate("endnote-empty-final-section", mode, story, variant, rich: rich);
                var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
                TestAssert.Equal(control.Length, pages.Length);
                for (int index = 0; index < pages.Length; index++)
                    TestAssert.Equal(DocxPageFootnoteNumberingTests.DecodeText(control[index]), DocxPageFootnoteNumberingTests.DecodeText(pages[index]));
                TestAssert.Equal(80, pages.Sum(page => Regex.Matches(DocxPageFootnoteNumberingTests.DecodeText(page), @"TailNote\d{3}").Count));
            }
        }
    }

    public static void DocxSelectableStaticOwnersRetainNoteFrameFallbacks()
    {
        foreach (string story in new[] { "header", "footer" })
        foreach (string variant in new[] { "default", "first", "even", "unknown" })
        {
            var source = ReadInactiveNoteTemplate("endnote-empty-final-section", OoxPdfDocxMarkupMode.Final, story, variant,
                enabled: variant != "unknown", rich: true);
            TestAssert.True(!DocxPlainBodyFrame.HasNoSelectableStaticStories(source),
                "Active or unknown static references retain their measured/fallback frame: " + story + "/" + variant);
            TestAssert.True(ReferenceEquals(source, DocxFinalSectionParagraph.Materialize(source, default)),
                "This repair does not change active static-story or unknown-source final-section placement.");
        }
        var disabled = ReadInactiveNoteTemplate("endnote-empty-final-section", OoxPdfDocxMarkupMode.Final, "header", "first");
        TestAssert.True(ReferenceEquals(disabled, DocxFinalSectionParagraph.Materialize(disabled, default, .8d)),
            "Scaled implicit-final-section geometry retains its existing fallback.");
    }

    public static void DocxStaticNoteEligibilityKeepsInheritedOwnersAndCancellation()
    {
        foreach (string story in new[] { "header", "footer" })
        {
            var source = ReadInactiveNoteTemplate("endnote-empty-final-section", OoxPdfDocxMarkupMode.Final, story, "first");
            var section = source.BodyElements.OfType<DocxSectionBreakElement>().Single();
            var current = source.PageSettings;
            var owner = section.PageSettings with
            {
                HeaderParagraphsByType = current.HeaderParagraphsByType, HeaderBodyElementsByType = current.HeaderBodyElementsByType,
                FooterParagraphsByType = current.FooterParagraphsByType, FooterBodyElementsByType = current.FooterBodyElementsByType
            };
            DocxDocument WithSection(DocxPageSettings settings, DocxDocument document) => document with
            { BodyElements = document.BodyElements.Select(item => ReferenceEquals(item, section) ? section with { PageSettings = settings } : item).ToArray() };
            TestAssert.True(!DocxPlainBodyFrame.HasNoSelectableStaticStories(WithSection(owner with { TitlePage = true }, source)),
                "An earlier active story is not hidden by the final section's disabled title page.");
            var final = current with { TitlePage = true,
                HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(),
                HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(),
                FooterParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(),
                FooterBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>() };
            var inherited = WithSection(owner with { TitlePage = false }, WithStaticVariantSettings(source, final));
            TestAssert.True(!DocxPlainBodyFrame.HasNoSelectableStaticStories(inherited),
                "A later section can activate an inherited first story without defining another reference.");
            using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxPlainBodyFrame.HasNoSelectableStaticStories(source, cancellation.Token));
            TestAssert.Throws<OperationCanceledException>(() => DocxFinalSectionParagraph.Materialize(source, cancellation.Token));
        }
    }

    private static DocxDocument ReadInactiveNoteTemplate(string fixture, OoxPdfDocxMarkupMode mode,
        string? story = null, string variant = "first", bool enabled = false, bool rich = false)
    {
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes", fixture + ".docx"));
        using var memory = new MemoryStream(); memory.Write(File.ReadAllBytes(path));
        if (story is not null)
        {
            using var zip = new ZipArchive(memory, ZipArchiveMode.Update, leaveOpen: true);
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
            var xml = Read("word/document.xml"); var final = xml.Root!.Element(w + "body")!.Elements(w + "sectPr").Single();
            final.AddFirst(new XElement(w + story + "Reference", new XAttribute(w + "type", variant), new XAttribute(r + "id", "inactiveTemplate")));
            if (enabled && variant == "first") final.Add(new XElement(w + "titlePg"));
            if (enabled && variant == "even")
            {
                var settings = Read("word/settings.xml"); settings.Root!.Add(new XElement(w + "evenAndOddHeaders")); Write("word/settings.xml", settings);
            }
            Write("word/document.xml", xml);
            var rels = Read("word/_rels/document.xml.rels"); string name = "inactive-" + story + ".xml";
            rels.Root!.Add(new XElement(rels.Root.Name.Namespace + "Relationship", new XAttribute("Id", "inactiveTemplate"),
                new XAttribute("Type", r.NamespaceName + "/" + story), new XAttribute("Target", name))); Write("word/_rels/document.xml.rels", rels);
            var types = Read("[Content_Types].xml"); types.Root!.Add(new XElement(types.Root.Name.Namespace + "Override",
                new XAttribute("PartName", "/word/" + name), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml." + story + "+xml")));
            Write("[Content_Types].xml", types);
            var root = new XElement(w + (story == "header" ? "hdr" : "ftr"), new XElement(w + "p", new XElement(w + "r", new XElement(w + "t", "UNPRINTED STATIC TEMPLATE"))));
            if (rich) root.Add(new XElement(w + "p", new XElement(m + "oMath", new XElement(m + "r", new XElement(m + "t", "z")))));
            Write("word/" + name, new XDocument(root));
            XDocument Read(string entry) { using var input = zip.GetEntry(entry)!.Open(); return XDocument.Load(input); }
            void Write(string entry, XDocument document) { zip.GetEntry(entry)?.Delete(); using var output = zip.CreateEntry(entry).Open(); document.Save(output); }
        }
        memory.Position = 0;
        return new DocxReader().Read(OoxPackage.Open(memory, default), null, default, mode);
    }
}
