using System.Globalization;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PdfWriterTests
{
    public static void PdfTransformsKeepDistantRotationPivotsAndLocalGeometry()
    {
        foreach (double degrees in new[] { -60d, -45d, 45d, 60d })
        foreach (double center in new[] { 500d, 2500d })
        {
            double radians = degrees * Math.PI / 180d;
            double a = Math.Cos(radians), b = Math.Sin(radians), c = -b, d = a;
            var graphics = new PdfGraphicsBuilder();
            graphics.Transform(a, b, c, d, center - a * center - c * center, center - b * center - d * center);
            string content = graphics.ToString();
            double[] matrix = content.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)
                .Take(6).Select(value => double.Parse(value, CultureInfo.InvariantCulture)).ToArray();
            double x = matrix[0] * center + matrix[2] * center + matrix[4];
            double y = matrix[1] * center + matrix[3] * center + matrix[5];
            TestAssert.True(Math.Abs(x - center) < .003 && Math.Abs(y - center) < .003,
                "Serializing a rotation must preserve its physical pivot.");
            TestAssert.True(Math.Abs(matrix[0] * 100d + matrix[2] * 20d - (a * 100d + c * 20d)) < .001 &&
                Math.Abs(matrix[1] * 100d + matrix[3] * 20d - (b * 100d + d * 20d)) < .001,
                "The same transform preserves geometry relative to the pivot.");
            PdfContentValidator.ValidatePage(new PdfPage(3000, 3000, content), 0, default);
        }
    }

    public static void PdfTransformNumbersKeepFiniteInvariantPlainDecimalEncoding()
    {
        CultureInfo original = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
            var graphics = new PdfGraphicsBuilder();
            graphics.Transform(.0001234, -.0001234, .1234567, 1d, 12.34567, -12.34567);
            TestAssert.Equal("0.000123 -0.000123 0.123457 1 12.346 -12.346 cm", graphics.ToString().Trim());
            TestAssert.Equal("12.346", PdfDocumentWriter.FormatNumber(12.34567));
            foreach (double invalid in new[] { double.NaN, double.PositiveInfinity, double.NegativeInfinity })
            for (int index = 0; index < 6; index++)
            {
                double[] values = [1d, 0d, 0d, 1d, 0d, 0d];
                values[index] = invalid;
                var rejected = new PdfGraphicsBuilder();
                TestAssert.Throws<ArgumentOutOfRangeException>(() =>
                    rejected.Transform(values[0], values[1], values[2], values[3], values[4], values[5]));
            }
        }
        finally { CultureInfo.CurrentCulture = original; }
    }
}
