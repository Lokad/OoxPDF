using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxFontSlotTests
{
    public static void DocxSymbolsKeepFontCodesAndSourceOwners()
    {
        var document = Read("<w:p><w:bookmarkStart w:id=\"1\" w:name=\"target\"/>" +
            "<w:hyperlink w:anchor=\"target\"><w:ins w:id=\"2\" w:author=\"Review\"><w:r>" +
            Properties.Replace("</w:rPr>", "<w:caps/><w:smallCaps/><w:b/></w:rPr>") +
            "<w:t>X</w:t>" + Symbol("0061") + "<w:t>Z</w:t></w:r></w:ins></w:hyperlink>" +
            "<w:bookmarkEnd w:id=\"1\"/></w:p>", "body");
        var p = document.Paragraphs.Single();
        TestAssert.Equal("X|a|Z", string.Join("|", p.Runs.Select(run => run.Text)));
        TestAssert.Equal("0|1|2", string.Join("|", p.Runs.Select(run => run.SourceTextOffsetInRun)));
        TestAssert.Equal(1, p.Runs.Select(run => run.SourceRunIndex).Distinct().Count());
        TestAssert.True(p.Runs.All(run => run.Bold && run.Revision is not null), "Symbol fragments retain formatting and revision ownership.");
        TestAssert.True(!p.Runs[1].AllCaps && !p.Runs[1].SmallCaps, "A font-relative code must not change case.");
        TestAssert.Equal("Slot A|Slot H|Slot A", string.Join("|", p.Runs.Select(run => DocxFontResolver.ResolveTypeface(run, document.FontCatalog).Primary)));
        TestAssert.Equal(1, p.Hyperlinks.Count);
        TestAssert.Equal(1, p.Hyperlinks[0].SourceRunCount);
        TestAssert.Equal(3, p.Hyperlinks[0].TextRunCount);
        TestAssert.Equal(1, p.BookmarkAnchors.Count);
    }

    public static void DocxSymbolsPaintTheSameGlyphsAcrossSupportedContainers()
    {
        var resolver = new SlotResolver();
        foreach (string container in new[] { "body", "cell", "header", "footer" })
        foreach (var (code, text) in new[] { ("0061", "a"), ("F041", "\uF041") })
        {
            string head = "<w:r>" + Properties + "<w:t xml:space=\"preserve\">BEFORE </w:t>";
            string tail = "<w:t xml:space=\"preserve\"> AFTER</w:t></w:r>";
            var symbol = Read("<w:p>" + head + Symbol(code) + tail + "</w:p>", container);
            var literal = Read("<w:p>" + head + "</w:r><w:r>" + Properties.Replace("Slot A", "Slot H").Replace("Slot E", "Slot H").Replace("Slot C", "Slot H") +
                "<w:t>" + text + "</w:t></w:r><w:r>" + Properties + tail + "</w:p>", container);
            PdfPage actual = SymbolPage(symbol, resolver), expected = SymbolPage(literal, resolver);
            TestAssert.Equal(expected.Content, actual.Content);
            TestAssert.Equal("BEFORE " + text + " AFTER", SymbolPaintText(actual).Trim());
            TestAssert.True(actual.Fonts.Any(resource => resource.Font.Font.FamilyName == "Slot H"), "The element's own font must reach paint in " + container);
            PdfContentValidator.ValidatePage(actual, 0, default);
        }
    }

    public static void DocxSymbolsRetainHiddenRevisionAndNavigationSelection()
    {
        string paragraph = "<w:p><w:bookmarkStart w:id=\"1\" w:name=\"target\"/><w:hyperlink w:anchor=\"target\">" +
            "<w:ins w:id=\"2\" w:author=\"Reviewer\"><w:r>" + Symbol("0061") + "</w:r></w:ins>" +
            "<w:del w:id=\"3\" w:author=\"Reviewer\"><w:r>" + Symbol("0042") + "</w:r></w:del>" +
            "<w:r><w:rPr><w:vanish/></w:rPr>" + Symbol("0043") + "</w:r></w:hyperlink><w:bookmarkEnd w:id=\"1\"/></w:p>";
        foreach (var (mode, expected) in new[] { (OoxPdfDocxMarkupMode.Final, "a"), (OoxPdfDocxMarkupMode.Original, "B"), (OoxPdfDocxMarkupMode.AllMarkup, "aB") })
        {
            var document = Read(paragraph, "body", mode);
            PdfPage page = SymbolPage(document, new SlotResolver(), mode);
            TestAssert.Equal(expected, SymbolPaintText(page).Trim());
            TestAssert.Equal(1, page.Annotations.Count);
            TestAssert.True(page.Annotations[0].Destination is not null, "The symbol retains its visible internal-link owner.");
            PdfContentValidator.ValidatePage(page, 0, default);
        }
    }

    public static void DocxSymbolsKeepSimpleAndComplexCachedFieldResults()
    {
        var document = Read("<w:p><w:fldSimple w:instr=\" DATE \"><w:r>" + Symbol("0041") + "</w:r></w:fldSimple>" +
            "<w:r><w:fldChar w:fldCharType=\"begin\"/><w:instrText> DATE </w:instrText>" +
            "<w:fldChar w:fldCharType=\"separate\"/>" + Symbol("0042") + "<w:fldChar w:fldCharType=\"end\"/></w:r>" +
            "<w:r><w:fldChar w:fldCharType=\"begin\" w:fldLock=\"1\"/><w:instrText> PAGE </w:instrText>" +
            "<w:fldChar w:fldCharType=\"separate\"/>" + Symbol("0043") + "<w:fldChar w:fldCharType=\"end\"/></w:r></w:p>", "body");
        var paragraph = document.Paragraphs.Single();
        TestAssert.Equal("ABC", string.Concat(paragraph.Runs.Select(run => run.Text)));
        TestAssert.Equal(3, paragraph.FieldReferences.Count);
        TestAssert.True(paragraph.FieldReferences.All(field => field.HasCachedResult && field.RendersCachedResult && !field.UsesPlaceholder), "Both field readers must recognize symbol content as a stored result.");
        TestAssert.Equal("ABC", SymbolPaintText(SymbolPage(document, new SlotResolver())).Trim());
    }

    public static void DocxSymbolsReplaceMalformedPayloadsAndDiagnoseOwningStories()
    {
        foreach (string payload in new[] { "w:font=\"Slot H\"", "w:char=\"0041\"", "w:font=\"Slot H\" w:char=\"ZZZZ\"", "w:font=\"Slot H\" w:char=\"D800\"", "w:font=\"Slot H\" w:char=\"0000\"", "w:font=\"Slot H\" w:char=\"00AD\"", "w:font=\"Slot H\" w:char=\"1F600\"" })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            var document = Read("<w:p><w:r><w:t>L</w:t><w:sym " + payload + "/><w:t>R</w:t></w:r></w:p>", "body", diagnostics: diagnostics.Add);
            TestAssert.Equal("L\uFFFDR", string.Concat(document.Paragraphs.Single().Runs.Select(run => run.Text)));
            var warning = diagnostics.Single(d => d.Id == "DOCX_UNSUPPORTED_SYMBOL");
            TestAssert.Equal("/word/document.xml", warning.PartName!);
            TestAssert.Equal("Replacement character", warning.Fallback!);
            TestAssert.Equal("L\uFFFDR", SymbolPaintText(SymbolPage(document, new PortableRenderingTests.PinnedResolver())).Trim());
        }
        foreach (string container in new[] { "cell", "header", "footer" })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            Read("<w:p><w:r><w:sym w:char=\"0041\"/></w:r></w:p>", container, diagnostics: diagnostics.Add);
            var warning = diagnostics.Single(d => d.Id == "DOCX_UNSUPPORTED_SYMBOL");
            TestAssert.Equal(container == "cell" ? "/word/document.xml" : "/word/" + container + "1.xml", warning.PartName!);
        }
    }

    private static string Symbol(string code) => "<w:sym w:font=\"Slot H\" w:char=\"" + code + "\"/>";
    private static PdfPage SymbolPage(DocxDocument document, Lokad.OoxPdf.Fonts.IFontResolver resolver,
        OoxPdfDocxMarkupMode mode = OoxPdfDocxMarkupMode.Final) =>
        new DocxRenderer(resolver, mode, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .RenderPagesAsync(document, null, default).ToBlockingEnumerable().Single();
    private static string SymbolPaintText(PdfPage page) => string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Select(show => show.Text));
}
