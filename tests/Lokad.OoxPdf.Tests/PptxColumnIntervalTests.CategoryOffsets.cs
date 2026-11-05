using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxColumnIntervalTests
{
    internal static void PptxCategoryOffsetsShareMeasuredUnitsAcrossPreparedContainers()
    {
        using var oracle = OuterOracle();
        var row = OuterRow(oracle, "outer-width35-left12-y03");
        // Independent Office inner/outer/automatic controls distinguish frame
        // movement from glyph movement. Use the pinned face's line metric so
        // this relation also holds with portable font substitution.
        foreach (string target in new[] { "outer", "inner", "automatic" })
        foreach (bool high in new[] { false, true })
        {
            PdfPage Render(int offset) => RenderCaption(row, 18d, high, 6d, offset, "ordinary", target: target);
            var baseline = Render(100);
            var reference = Shows(baseline);
            TestAssert.Equal(8, reference.Length);
            var font = reference[0].Font.Font;
            var os2 = font.Os2;
            double line = 18d / font.UnitsPerEm * ((os2.SelectionFlags & 0x80) != 0
                ? os2.TypographicAscender - os2.TypographicDescender + os2.TypographicLineGap
                : os2.WindowsAscender + os2.WindowsDescender);
            var original = AdmissionBounds(baseline);
            foreach (int offset in new[] { 0, 50, 200 })
            {
                var page = Render(offset);
                var labels = Shows(page);
                var box = AdmissionBounds(page);
                double increment = line * .25d * (offset / 100d - 1d);
                TestAssert.True(labels.Select(show => show.Text).SequenceEqual(reference.Select(show => show.Text)),
                    "Offset changes retain every source caption.");
                if (target == "inner")
                {
                    TestAssert.True(new[] { box.Left - original.Left, box.Top - original.Top,
                        box.Right - original.Right, box.Bottom - original.Bottom }.All(delta => Math.Abs(delta) < .002d),
                        "An admitted inner rectangle retains its authored geometry.");
                    double shift = high ? increment : -increment;
                    TestAssert.True(labels.Zip(reference).All(pair => Math.Abs(pair.First.Y - pair.Second.Y - shift) < .002d),
                        "Inner captions move by the normal-line offset increment, not a scaled baseline.");
                }
                else
                {
                    double shift = high ? box.Top - original.Top : original.Bottom - box.Bottom;
                    TestAssert.True(Math.Abs(shift - increment) < .002d,
                        "Outer and automatic frames reserve the same measured offset increment.");
                    TestAssert.True(labels.Zip(reference).All(pair => Math.Abs(pair.First.Y - pair.Second.Y) < .002d),
                        "Outer/automatic admission keeps the painted baseline at its container edge.");
                }
                TestAssert.True(AdmissionNumericShows(page).Select(show => show.Text)
                    .SequenceEqual(AdmissionNumericShows(baseline).Select(show => show.Text)),
                    "Offset changes preserve the numeric interval.");
            }
        }

        static (PdfEmbeddedFont Font, string Text, double Size, double X, double Y)[] Shows(PdfPage page) =>
            DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).ToArray();
    }
}
