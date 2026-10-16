using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxPageFootnoteNumberingTests
{
    public static void DocxMeasuredFootnoteOwnersAlreadyCorrectAvoidsResourceRetry()
    {
        string input = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes",
            "footnote-page-owners-single-page-unchanged.docx"));
        var source = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
        TestAssert.Equal("09,10,11,12", string.Join(',', source.Paragraphs.SelectMany(p => p.InlineReferences).Select(reference => reference.DisplayText)));
        var layout = CreateLayout(source);
        TestAssert.Equal(1, layout.Pages.Count);
        TestAssert.True(!DocxPageFootnoteNumbering.TryCreateDocument(source, layout, out var unchanged, out _,
            allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences),
            "An already-correct complete page must not request a font/layout retry.");
        TestAssert.True(ReferenceEquals(source, unchanged), "No-op numbering retains the original source document.");
    }

    public static void DocxMeasuredFootnoteOwnersMatchIndependentWordAcrossTypography()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent note controls use installed Windows fonts.");
        CheckMeasuredFootnoteOwners(null);
    }

    internal static void CheckMeasuredFootnoteOwners(IFontResolver? resolver)
    {
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var golden = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "footnote-page-owners-office.json")));
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string input = Path.Combine(folder, row.GetProperty("Id").GetString()! + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            var source = DocxTests.ReadDocx(input, mode);
            var owners = source.Paragraphs.SelectMany(p => p.Runs).ToArray();
            double baselineSize = owners.Where(run => run.VerticalAlignment == DocxRunVerticalAlignment.Baseline && run.Text.Length != 0).Min(run => run.FontSize);
            var renderer = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var pages = renderer.RenderBlankPages(source, null, default).ToArray();
            TestAssert.Equal(2, pages.Length);
            for (int index = 0; index < pages.Length; index++)
            {
                var page = pages[index];
                var markerText = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                    .Where(show => show.Size < baselineSize * .8d && show.Text.Length != 0 && show.Text.All(char.IsAsciiDigit))
                    .Select(show => show.Text).ToArray();
                var expected = row.GetProperty("Markers").EnumerateArray()
                    .Where(marker => marker.GetProperty("PageNumber").GetInt32() == index + 1)
                    .Select(marker => marker.GetProperty("DecodedText").GetString()).ToArray();
                TestAssert.Equal(string.Join(',', expected), string.Join(',', markerText));
                string text = DecodeText(page);
                TestAssert.Contains("105", text); TestAssert.Contains("108", text);
                TestAssert.Equal(2, page.Annotations.Count);
                TestAssert.Equal(2, page.Annotations.Select(annotation => annotation.Destination).Distinct().Count());
                TestAssert.True(page.Annotations.All(annotation => annotation.Destination!.Value.PageIndex == index &&
                    annotation.Width > 0d && annotation.Height > 0d), "Each marker still targets its complete same-page note.");
            }
            TestAssert.Equal("105,106,107,108", string.Join(',', source.Paragraphs.SelectMany(p => p.InlineReferences).Select(reference => reference.DisplayText)));
            TestAssert.True(owners.Zip(source.Paragraphs.SelectMany(p => p.Runs)).All(pair => ReferenceEquals(pair.First, pair.Second)),
                "Numbering must preserve original parsed runs and authored digits.");
        }
    }
}
