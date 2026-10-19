using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxPrintedParagraphAdmissionPreservesMeasuredFooterReserve()
    {
        foreach (double scale in new[] { .5d, .75d })
        foreach (double reserve in new[] { 0d, 10d, 20d })
        {
            var paragraphs = Enumerable.Range(0, 40).Select(_ => DocxTests.CreateDocxLayoutParagraph("Line", 10d, 10d)).ToArray();
            var document = new DocxDocument(300d, 200d, 20d, 20d, 20d, 20d, DocxPageSettings.Empty, [], [], [],
                paragraphs.Select(paragraph => (DocxBodyElement)new DocxParagraphElement(paragraph)).ToArray(), paragraphs, []);
            var measurer = new DocxTests.FamilyWidthTextMeasurer();
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, scale)
                .Create(document, new DocxTests.ScaledLayoutTextMeasurer(measurer, scale, scale), default, measurer,
                    footerFrameBottomDisplacementByPage: new Dictionary<int, double> { [0] = reserve });
            // The160pt source body becomes80/120pt. A measured footer's10/20pt
            // reserve is already printed geometry, independent of that scale.
            double lowerEdge = 180d - 160d * scale + reserve;
            TestAssert.True(layout.Pages.SelectMany(page => page.Items.OfType<DocxTextLineLayout>())
                .All(line => line.BaselineY >= lowerEdge), "Printed paragraph baselines must clear the measured footer reserve.");
        }
    }

    public static void DocxComposedInlineFrameReportRetainsIndependentWordPagesAndContent()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent composed report geometry uses installed Windows fonts.");
        CheckComposedInlineFrameReport(null, independentWordGeometry: true);
    }

    internal static void CheckComposedInlineFrameReport(IFontResolver? resolver, bool independentWordGeometry)
    {
        string directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "Composition"));
        using var gold = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "composed-inline-frame-office.json")));
        string input = Path.Combine(directory, "composed-inline-frame-report.docx");
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var source = DocxTests.ReadDocx(input, mode);
            var diagnostics = new List<Diagnostics.OoxPdfDiagnostic>();
            var pages = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(source, diagnostics.Add, default).ToArray();
            var shows = pages.SelectMany((page, index) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page,
                usePageCoordinates: true).Select(show => (Page: index + 1, Height: page.Height, Show: show))).ToArray();
            var lines = shows.GroupBy(item => (item.Page, item.Show.Y)).Select(line =>
                (line.Key.Page, Y: line.First().Height - line.Key.Y, X: line.Min(item => item.Show.X),
                    Text: Regex.Replace(string.Concat(line.OrderBy(item => item.Show.X).Select(item => item.Show.Text))
                        .Normalize(NormalizationForm.FormKC), @"\s+", " ").Trim())).ToArray();
            string prose = string.Join(" ", lines.Select(line => line.Text));
            foreach (var text in row.GetProperty("ExpectedText").EnumerateArray())
                TestAssert.Contains(text.GetString()!, prose);
            TestAssert.Equal(0, diagnostics.Count(message => message.Id == "DOCX_UNSUPPORTED_EQUATION" && message.Fallback == "Ignored"));
            var boxes = lines.Where(line => line.Text.StartsWith("Inline box content.", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(2 * pages.Length + 1, boxes.Length);
            TestAssert.Equal(2 * pages.Length + 3, pages.Sum(page => page.Images.Count));
            TestAssert.True(boxes.All(line => line.Y > 0d && line.Y < pages[line.Page - 1].Height),
                "Every repeated static box and the body box must retain visible on-page text origins.");
            foreach (var anchor in row.GetProperty("WordBodyParagraphStarts").EnumerateArray())
            {
                var actual = lines.Single(line => line.Text.StartsWith(anchor.GetProperty("Text").GetString()!, StringComparison.Ordinal));
                if (independentWordGeometry) TestAssert.Equal(anchor.GetProperty("Page").GetInt32(), actual.Page);
            }
            if (!independentWordGeometry) continue;
            TestAssert.Equal(row.GetProperty("WordPages").GetInt32(), pages.Length);
            foreach (var anchor in row.GetProperty("WordProseAnchorGroups")[0].GetProperty("Occurrences").EnumerateArray())
                TestAssert.Equal(1, boxes.Count(line => line.Page == anchor.GetProperty("Page").GetInt32() &&
                    Math.Abs(line.X - anchor.GetProperty("X").GetDouble()) < 1d && Math.Abs(line.Y - anchor.GetProperty("Y").GetDouble()) < 1d));
            var nativeMath = shows.Where(item => item.Show.Font.Font.FamilyName.Replace(" ", string.Empty, StringComparison.Ordinal)
                .Contains("CambriaMath", StringComparison.OrdinalIgnoreCase)).SelectMany(item => item.Show.Text
                    .Normalize(NormalizationForm.FormKC).Where(character => !char.IsWhiteSpace(character))
                    .Select(character => (item.Page, Text: character.ToString()))).OrderBy(item => item.Page).ThenBy(item => item.Text, StringComparer.Ordinal).ToArray();
            var wordMath = row.GetProperty("WordMathAnchors").EnumerateArray().Select(anchor =>
                (Page: anchor.GetProperty("Page").GetInt32(), Text: anchor.GetProperty("Text").GetString()!))
                .OrderBy(item => item.Page).ThenBy(item => item.Text, StringComparer.Ordinal).ToArray();
            TestAssert.True(nativeMath.SequenceEqual(wordMath), "Per-page mathematical content retains the independent Word glyph multiset.");
        }
    }
}
