using System.Text.RegularExpressions;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PdfWriterTests
{
    public static void AlphaResourceDictionariesShareOnlyEqualEmittedPairs()
    {
        var graphics = new PdfGraphicsBuilder();
        for (int i = 1; i <= 10; i++)
        {
            graphics.SetAlpha(.5d + i / 100000d, 1d);
            graphics.FillRectangle(i * 5d, 10d, 4d, 10d);
        }
        TestAssert.Equal(1, graphics.ExtGStates.Count);
        string pdf = WritePdfText([new PdfPage(100d, 100d, graphics.ToString(), [], [], graphics.ExtGStates)]);
        TestAssert.Equal(1, Regex.Matches(pdf, @"/ca 0\.5 /CA 1\b").Count);
        TestAssert.Equal(1, Regex.Matches(pdf, @"/GS50000F100000S <<").Count);
        graphics.SetAlpha(.5005d, 1d);
        graphics.SetAlpha(.5d, .99949d);
        pdf = WritePdfText([new PdfPage(100d, 100d, graphics.ToString(), [], [], graphics.ExtGStates)]);
        TestAssert.Equal(3, graphics.ExtGStates.Count);
        TestAssert.Equal(1, Regex.Matches(pdf, @"/ca 0\.501 /CA 1\b").Count);
        TestAssert.Equal(1, Regex.Matches(pdf, @"/ca 0\.5 /CA 0\.999\b").Count);

        // Both inputs round to .50050 at five decimals, but their actual PDF
        // values straddle the three-decimal boundary. They must stay distinct.
        var boundary = new PdfGraphicsBuilder();
        boundary.SetAlpha(.5004999d, 1d);
        boundary.SetAlpha(.5005001d, 1d);
        TestAssert.Equal(2, boundary.ExtGStates.Count);
        pdf = WritePdfText([new PdfPage(100d, 100d, boundary.ToString(), [], [], boundary.ExtGStates)]);
        TestAssert.Equal(1, Regex.Matches(pdf, @"/ca 0\.5 /CA 1\b").Count);
        TestAssert.Equal(1, Regex.Matches(pdf, @"/ca 0\.501 /CA 1\b").Count);
    }
}
