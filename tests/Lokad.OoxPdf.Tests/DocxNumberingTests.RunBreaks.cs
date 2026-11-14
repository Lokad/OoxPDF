using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxNumberingTests
{
    public static void DocxNumberedRunBreaksConsumeOneSourceCounterEvent()
    {
        foreach (string kind in new[] { "page", "column" })
        foreach (int count in new[] { 1, 2 })
        foreach (bool trailing in new[] { false, true })
        foreach (bool leading in new[] { false, true })
        foreach (bool insertedMark in new[] { false, true })
        foreach (OoxPdfDocxMarkupMode mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            var source = ReadNumberedRunBreaks(kind, count, trailing, leading, insertedMark, mode);
            var next = source.Paragraphs.Single(p => string.Concat(p.Runs.Select(run => run.Text)) == "NEXT");
            var last = source.Paragraphs.Single(p => string.Concat(p.Runs.Select(run => run.Text)) == "LAST");
            TestAssert.Equal("2.", TestAssert.NotNull(next.ListLabel).Text);
            TestAssert.Equal("3.", TestAssert.NotNull(last.ListLabel).Text);
            bool review = DocxMarkupContext.FromMode(mode).AppliesInlineRevisionStyle;
            TestAssert.Equal(insertedMark && review ? "1." : null, next.ListLabel?.NumberingChange?.PreviousText);
            TestAssert.Equal(insertedMark && review ? "2." : null, last.ListLabel?.NumberingChange?.PreviousText);
            TestAssert.True(source.Paragraphs.Where(p => string.Concat(p.Runs.Select(run => run.Text)).StartsWith("AFTER", StringComparison.Ordinal))
                .All(p => p.ListLabel is null), "Continuation fragments do not acquire independent labels.");
        }
    }

    public static void DocxNumberedRunBreaksPaintConsecutiveLabels()
    {
        foreach (string kind in new[] { "page", "column" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadNumberedRunBreaks(kind, 2, trailing: true, leading: false, insertedMark: false, mode);
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText));
            foreach (string label in new[] { "1.", "2.", "3.", "BEFORE", "AFTER1", "NEXT", "LAST" })
                TestAssert.Contains(label, text);
            TestAssert.True(!text.Contains("4.", StringComparison.Ordinal) && !text.Contains("5.", StringComparison.Ordinal),
                "Unpainted continuation labels cannot advance visible list items.");
            for (int page = 0; page < pages.Length; page++) PdfContentValidator.ValidatePage(pages[page], page, default);
            using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(source, null, cancellation.Token)
                .ToBlockingEnumerable().ToArray());
        }
    }

    private static DocxDocument ReadNumberedRunBreaks(string kind, int count, bool trailing, bool leading,
        bool insertedMark, OoxPdfDocxMarkupMode mode)
    {
        string number = "<w:numPr><w:numId w:val=\"1\"/></w:numPr>";
        string mark = insertedMark ? "<w:rPr><w:ins w:id=\"31\" w:author=\"Public reviewer\"/></w:rPr>" : "";
        string breakRun = "<w:r><w:br w:type=\"" + kind + "\"/></w:r>";
        string middle = string.Concat(Enumerable.Range(1, count).Select(index => breakRun +
            (trailing && index == count ? "" : "<w:r><w:t>AFTER" + index + "</w:t></w:r>")));
        string body = "<w:p><w:pPr>" + number + mark + "</w:pPr>" + (leading ? breakRun : "") +
            "<w:r><w:t>BEFORE</w:t></w:r>" + middle + "</w:p>" +
            "<w:p><w:pPr>" + number + "</w:pPr><w:r><w:t>NEXT</w:t></w:r></w:p>" +
            "<w:p><w:pPr>" + number + "</w:pPr><w:r><w:t>LAST</w:t></w:r></w:p>";
        using var input = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>
                """,
            ["_rels/.rels"] = """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="numbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>
                """,
            ["word/numbering.xml"] = """
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:suff w:val="space"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>
                """,
            ["word/document.xml"] = "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
                body + "<w:sectPr><w:cols w:num=\"2\"/></w:sectPr></w:body></w:document>"
        });
        return new DocxReader().Read(OoxPackage.Open(input, default), null, default, mode);
    }
}
