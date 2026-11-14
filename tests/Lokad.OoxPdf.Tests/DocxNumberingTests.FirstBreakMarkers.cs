using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxNumberingTests
{
    public static void DocxHardBreaksRetainParagraphOwnedFirstMarkers()
    {
        foreach (string kind in new[] { "page", "column" })
        foreach (bool pure in new[] { false, true })
        foreach (int count in new[] { 1, 2 })
        foreach (string carrier in new[] { "runs", "control", "hyperlink" })
        foreach (bool inserted in new[] { false, true })
        foreach (OoxPdfDocxMarkupMode mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            var source = ReadFirstMarkerBreak(kind, pure, count, carrier, "decimal", inserted, mode);
            var numbered = source.BodyElements.OfType<DocxParagraphElement>()
                .Where(element => element.Paragraph.ListLabel is not null).ToArray();
            TestAssert.True(numbered.Select(element => element.Paragraph.ListLabel!.Text).SequenceEqual(new[] { "1.", "2.", "3." }),
                "A hard break cannot compact away a paragraph-owned marker or consume another numbering event.");
            TestAssert.True(string.Concat(numbered[0].Paragraph.Runs.Select(run => run.Text)).Length == 0,
                "The existing first source line owns the marker before the leading break, including a break-only paragraph.");
            int markerIndex = source.BodyElements.ToList().IndexOf(numbered[0]);
            int breakIndex = source.BodyElements.ToList().FindIndex(element => element is DocxPageBreakElement or DocxManualBreakElement);
            TestAssert.True(markerIndex < breakIndex, "Marker placement follows source first-line order.");
            int events = source.BodyElements.OfType<DocxPageBreakElement>().Sum(element => element.BreakCount) +
                source.BodyElements.OfType<DocxManualBreakElement>().Sum(element => element.BreakCount);
            TestAssert.Equal(count, events);
            bool review = DocxMarkupContext.FromMode(mode).AppliesInlineRevisionStyle;
            TestAssert.Equal(inserted && review ? "1." : null, numbered[1].Paragraph.ListLabel?.NumberingChange?.PreviousText);
            TestAssert.Equal(inserted && review ? "2." : null, numbered[2].Paragraph.ListLabel?.NumberingChange?.PreviousText);
        }
    }

    public static void DocxHardBreakFirstMarkersPaintOnTheirSourceLine()
    {
        foreach (string kind in new[] { "page", "column" })
        foreach (bool pure in new[] { false, true })
        foreach (int count in new[] { 1, 2 })
        foreach (string family in new[] { "decimal", "bullet" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            (double X, double Y, double Size) Marker(string breakKind)
            {
                var source = ReadFirstMarkerBreak(breakKind, pure, count, "runs", family, false, mode);
                var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
                string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText));
                foreach (string word in pure ? new[] { "BEFORE", "NEXT", "LAST" } : new[] { "BEFORE", "AFTER", "NEXT", "LAST" })
                    TestAssert.Contains(word, text);
                string marker = family == "bullet" ? "\u2022" : "1.";
                if (family == "decimal") foreach (string label in new[] { "1.", "2.", "3." }) TestAssert.Contains(label, text);
                else TestAssert.Equal(3, text.Count(character => character == '\u2022'));
                var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(pages[0], usePageCoordinates: true, includeTextTransforms: true);
                TestAssert.True(shows.Any(show => show.Text == marker), "The source first marker must paint on its initial frame.");
                var first = shows.First(show => show.Text == marker);
                for (int page = 0; page < pages.Length; page++) PdfContentValidator.ValidatePage(pages[page], page, default);
                return (first.X, first.Y, first.Size);
            }
            var control = Marker("soft"); var actual = Marker(kind);
            TestAssert.True(Math.Abs(control.X - actual.X) < .001d && Math.Abs(control.Y - actual.Y) < .001d && control.Size == actual.Size,
                "A generated marker retains the measured source first line when only the following text restart changes: " + kind + "/" + pure + "/" + family);
            using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(
                ReadFirstMarkerBreak(kind, pure, count, "runs", family, false, mode), null, cancellation.Token)
                .ToBlockingEnumerable().ToArray());
        }
    }

    public static void DocxHardBreakMarkerAdmissionKeepsEmptyAndUnknownFallbacks()
    {
        foreach (string kind in new[] { "page", "column" })
        foreach (string family in new[] { "empty", "whitespace", "unknown", "disabled" })
        foreach (OoxPdfDocxMarkupMode mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            var source = ReadFirstMarkerBreak(kind, true, 2, "runs", family, false, mode);
            TestAssert.Equal(3, source.BodyElements.OfType<DocxParagraphElement>().Count());
            TestAssert.Equal(1, source.BodyElements.Count(element => element is DocxPageBreakElement or DocxManualBreakElement));
            var compact = source.BodyElements.Single(element => element is DocxPageBreakElement or DocxManualBreakElement);
            TestAssert.True(compact is DocxPageBreakElement { BreakCount: 2, BreakParagraph: not null } or
                DocxManualBreakElement { BreakCount: 2, BreakParagraph: not null },
                "A paragraph without marker ink keeps its existing break/mark frame instead of creating another row.");
        }
    }

    private static DocxDocument ReadFirstMarkerBreak(string kind, bool pure, int count, string carrier, string family,
        bool inserted, OoxPdfDocxMarkupMode mode, bool inCell = false)
    {
        string breaks = string.Concat(Enumerable.Repeat(string.Concat(kind.Split('-').Select(value =>
            "<w:r><w:br" + (value == "soft" ? "" : " w:type=\"" + value + "\"") + "/></w:r>")), count));
        breaks = carrier switch
        {
            "control" => "<w:sdt><w:sdtContent>" + breaks + "</w:sdtContent></w:sdt>",
            "hyperlink" => "<w:hyperlink>" + breaks + "</w:hyperlink>",
            _ => breaks
        };
        string numberId = family == "unknown" ? "999" : family == "disabled" ? "0" : "1";
        string properties = "<w:numPr><w:numId w:val=\"" + numberId + "\"/></w:numPr>" +
            (inserted ? "<w:rPr><w:ins w:id=\"31\" w:author=\"Public reviewer\"/></w:rPr>" : "");
        string following = "<w:p><w:pPr><w:numPr><w:numId w:val=\"1\"/></w:numPr></w:pPr><w:r><w:t>";
        string body = "<w:p><w:r><w:t>BEFORE</w:t></w:r></w:p><w:p><w:pPr>" + properties + "</w:pPr>" +
            breaks + (pure ? "" : "<w:r><w:t>AFTER</w:t></w:r>") + "</w:p>" +
            following + "NEXT</w:t></w:r></w:p>" + following + "LAST</w:t></w:r></w:p>";
        if (inCell) body = "<w:tbl><w:tblPr><w:tblW w:w=\"5000\" w:type=\"dxa\"/></w:tblPr>" +
            "<w:tblGrid><w:gridCol w:w=\"5000\"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w=\"5000\" w:type=\"dxa\"/></w:tcPr>" +
            body + "</w:tc></w:tr></w:tbl>";
        string format = family is "bullet" or "empty" or "whitespace" ? "bullet" : "decimal";
        string pattern = family == "empty" ? "" : family == "whitespace" ? " " : family == "bullet" ? "\u2022" : "%1.";
        using var input = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>
                """,
            ["_rels/.rels"] = """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="numbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>
                """,
            ["word/styles.xml"] = """
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="DejaVu Sans" w:hAnsi="DejaVu Sans"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>
                """,
            ["word/numbering.xml"] = "<w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:abstractNum w:abstractNumId=\"7\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"" + format + "\"/><w:lvlText w:val=\"" + pattern + "\"/><w:pPr><w:ind w:left=\"720\" w:hanging=\"360\"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId=\"1\"><w:abstractNumId w:val=\"7\"/></w:num></w:numbering>",
            ["word/document.xml"] = "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
                body + "<w:sectPr><w:cols w:num=\"2\"/></w:sectPr></w:body></w:document>"
        });
        return new DocxReader().Read(OoxPackage.Open(input, default), null, default, mode);
    }
}
