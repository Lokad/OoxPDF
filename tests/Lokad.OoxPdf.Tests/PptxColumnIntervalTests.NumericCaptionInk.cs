using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxColumnIntervalTests
{
    public static void PptxVerticalNumericCaptionsKeepFormattedGlyphsInsideTheirClips()
    {
        using var oracle = LogicalCaptionOracle();
        foreach (bool right in new[] { false, true })
        {
            var page = RenderAdmission(LogicalCaptionRow(oracle, "reference-numeric-format"),
                caseDirectory: LogicalCaptionCases, editChart: chart =>
                {
                    XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
                    chart.Descendants(c + "valAx").Single().Element(c + "tickLblPos")!.SetAttributeValue("val", right ? "high" : "low");
                });
            var shows = AdmissionNumericShows(page);
            TestAssert.True(shows.Select(show => show.Text).SequenceEqual(new[] { "7.00", "32.00", "57.00", "82.00", "107.00" }),
                "Formatted axis values survive on either source-selected side.");
            const string number = @"([-\d.]+)";
            var clips = Regex.Matches(page.Content, number + " " + number + " " + number + " " + number +
                @" re W\* n\s+BT\s+0 g\s+/CVA\d+ ").ToArray();
            TestAssert.Equal(shows.Length, clips.Length);
            for (int index = 0; index < shows.Length; index++)
            {
                var show = shows[index];
                var clip = clips[index];
                double left = double.Parse(clip.Groups[1].Value, CultureInfo.InvariantCulture);
                double edge = left + double.Parse(clip.Groups[3].Value, CultureInfo.InvariantCulture);
                double cursor = show.X;
                var font = show.Font.Font;
                double scale = show.Size / font.UnitsPerEm;
                foreach (Rune rune in show.Text.EnumerateRunes())
                {
                    ushort glyph = font.MapCodePoint(rune.Value);
                    TestAssert.True(font.TryGetGlyphBounds(glyph, out var bounds), "Expected pinned numeric glyph bounds.");
                    TestAssert.True(cursor + bounds.XMin * scale >= left - .005d &&
                        cursor + bounds.XMax * scale <= edge + .005d,
                        "Every formatted numeric glyph remains visible inside its own clip: " + show.Text);
                    cursor += font.GetAdvanceWidth(glyph) * scale;
                }
            }
        }
    }
}
