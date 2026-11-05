using System.Text.Json;
using System.Xml.Linq;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxColumnIntervalTests
{
    private static string NumericSideCases => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "NumericSideSpacing"));

    public static void PptxNumericSideBandsFollowIndependentFontLineGeometry()
    {
        using var oracle = JsonDocument.Parse(File.ReadAllText(Path.Combine(NumericSideCases, "office.json")));
        foreach (var row in oracle.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var page = RenderAdmission(row, caseDirectory: NumericSideCases);
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                includeTextTransforms: true, fontResourceFilter: name => name.StartsWith("CVA", StringComparison.Ordinal)).ToArray();
            TestAssert.True(shows.Select(show => show.Text).SequenceEqual(new[] { "7", "32", "57", "82", "107" }),
                row.GetProperty("Id").GetString() + ": changing the numeric face preserves the complete authored numeric interval: " +
                string.Join(",", shows.Select(show => show.Text)));
            var font = shows[0].Font.Font;
            var os2 = font.Os2;
            double units = (os2.SelectionFlags & 0x80) != 0
                ? os2.TypographicAscender - os2.TypographicDescender + os2.TypographicLineGap
                : os2.WindowsAscender + os2.WindowsDescender;
            double selectedLine = units * shows[0].Size / font.UnitsPerEm;
            double nativeLine = row.GetProperty("NativeNumericLineHeight").GetDouble();
            double padding = row.GetProperty("OuterPadding").GetDouble();
            double sourceLeft = row.GetProperty("SourceLeft").GetDouble();
            double nativeBand = row.GetProperty("NativeFrame")[0].GetDouble() - sourceLeft;
            double nativeAdvance = row.GetProperty("NativeNumericAdvance").GetDouble();
            double nativeGap = nativeBand - nativeAdvance - padding;
            // The pinned portable face changes advance and line height. Scale
            // the independently observed native gap, without embedding the
            // renderer's calibrated coefficient in this expectation.
            double expected = padding + shows.Max(show => show.Font.MeasureTextPoints(show.Text, show.Size)) +
                nativeGap * selectedLine / nativeLine;
            TestAssert.True(Math.Abs(AdmissionBounds(page).Left - sourceLeft - expected) <= 1d,
                row.GetProperty("Id").GetString() + ": the numeric side band follows the selected face's line geometry.");
        }
    }

    public static void PptxNumericSideGapsShareBothLabelSidesAndPreserveCaptions()
    {
        using var oracle = JsonDocument.Parse(File.ReadAllText(Path.Combine(NumericSideCases, "office.json")));
        foreach (var row in oracle.RootElement.GetProperty("Rows").EnumerateArray().Where(row =>
            row.GetProperty("Id").GetString() is "numeric-arial36" or "numeric-calibri24" or "numeric-segoe24"))
        {
            var left = RenderAdmission(row, caseDirectory: NumericSideCases);
            var right = RenderAdmission(row, caseDirectory: NumericSideCases, editChart: chart =>
            {
                XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
                chart.Descendants(c + "valAx").Single().Element(c + "tickLblPos")!.SetAttributeValue("val", "high");
            });
            double lowBand = AdmissionBounds(left).Left - row.GetProperty("SourceLeft").GetDouble();
            double highBand = row.GetProperty("SourceRight").GetDouble() - AdmissionBounds(right).Right;
            NearHorizontal(lowBand, highBand);
            TestAssert.True(AdmissionNumericShows(left).Select(show => show.Text)
                .SequenceEqual(AdmissionNumericShows(right).Select(show => show.Text)),
                "Mirroring the numeric strip preserves every authored caption.");
        }
    }
}
