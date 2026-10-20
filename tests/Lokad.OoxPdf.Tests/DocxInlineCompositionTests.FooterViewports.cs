using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxFooterTableOutlineViewportMatchesIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent footer viewport origins use installed Windows fonts.");
        CheckFooterTableOutlineViewport(null);
    }

    internal static void CheckFooterTableOutlineViewport(IFontResolver? resolver)
    {
        using var gold = ReadInlineBoxMathGold("footer-table-viewports-office.json");
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        foreach (bool transformed in new[] { false, true })
        {
            string id = row.GetProperty("Id").GetString()!;
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var document = ReadMutation("InlineBoxMath/" + id, mode,
                transformed && row.GetProperty("WordMathAnchors").GetArrayLength() != 0, transformed,
                partOverride: row.GetProperty("MathPart").GetString());
            var renderer = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var page = renderer.RenderBlankPages(document, null, default).Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).ToArray();
            string content = string.Concat(shows.Select(show => show.Text)).Normalize(NormalizationForm.FormKC);
            foreach (var text in row.GetProperty("ExpectedText").EnumerateArray())
                TestAssert.True(content.Contains(text.GetString()!, StringComparison.Ordinal),
                    id + ": the outline viewport must preserve parent and child source text.");
            TestAssert.Equal(row.GetProperty("WordMathAnchors").GetArrayLength(), shows.Count(show =>
                page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(font.Font, show.Font))));
            var clip = row.GetProperty("WordOutlineClip");
            double expectedBottom = page.Height - clip.GetProperty("IntersectedCenterlineBounds")[3].GetDouble();
            double expectedWidth = clip.GetProperty("SourceRect")[2].GetDouble() - clip.GetProperty("SourceRect")[0].GetDouble();
            double fullHeight = clip.GetProperty("SourceRect")[3].GetDouble() - clip.GetProperty("SourceRect")[1].GetDouble();
            bool found = Regex.Matches(page.Content, @"[-\d.]+ ([-\d.]+) ([-\d.]+) ([-\d.]+) re\s+W\s+n").Any(match =>
                Math.Abs(double.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture) - expectedBottom) < 1d &&
                Math.Abs(double.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture) - expectedWidth) < 1d &&
                double.Parse(match.Groups[3].Value, CultureInfo.InvariantCulture) < fullHeight - 1d);
            TestAssert.True(found, id + ": the painted outline retains Word's independent body-frame viewport.");
            using var cancelled = new CancellationTokenSource();
            cancelled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderBlankPages(document, null, cancelled.Token).ToArray());
        }
    }
}
