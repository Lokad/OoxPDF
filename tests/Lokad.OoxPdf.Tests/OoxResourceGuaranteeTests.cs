using System.Reflection;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class OoxResourceGuaranteeTests
{
    public static void JpegScanReservesBeforeAllocating()
    {
        byte[] jpeg = new byte[] { 0xff, 0xd8, 0xff, 0xc0, 0, 11, 8, 4, 0, 4, 0, 1, 1, 0x11, 0, 0xff, 0xda, 0, 8, 1, 1, 0, 0, 63, 0, 0xff, 0xd9 };
        using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxLiveImageBytesPerConversion = 0 }))
        {
            TestAssert.Throws<OoxPdfLimitExceededException>(() => JpegImage.Read(jpeg));
            TestAssert.Equal(0, scope.Budget.PeakLiveImageBytes);
        }

        using (OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxLiveImageBytesPerConversion = 256L * 1024L * 1024L }))
        {
            TestAssert.Throws<InvalidDataException>(() => JpegImage.Read(jpeg));
        }
    }

    public static void PptxImageVariantsRespectImageBudget()
    {
        foreach (string name in new[] { "pptx-ladder-07-image-crop.pptx", "pptx-ladder-07-jpeg-duotone-recolor-diagnostic.pptx" })
        {
            string input = FindCase(name);
            string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
            TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions
            {
                InputKind = OoxPdfInputKind.Pptx,
                ConversionLimits = new OoxConversionLimits { MaxImagesDecodedPerConversion = 0 },
            }));
            TestAssert.True(!File.Exists(output) || new FileInfo(output).Length == 0, "Budget failure must not publish image output for " + name);
        }
    }

    public static void DocxFontsRespectFontBudget()
    {
        string input = FindCase("docx-tables.docx");
        var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            ReportResourceUsage = true,
            DiagnosticSink = diagnostics.Add,
        });
        string summary = diagnostics.Single(d => d.Id == "CONVERSION_RESOURCE_SUMMARY").Message;
        TestAssert.Contains("fontWork=", summary);
        int fontWork = ParseCounter(summary, "fontWork=");
        TestAssert.True(fontWork > 0, "DOCX tables conversion must report font work, got: " + summary);

        string failing = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, failing, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            ConversionLimits = new OoxConversionLimits { MaxFontWorkPerConversion = 0 },
        }));
    }

    public static void ChartDensePointsRespectRangeBudget()
    {
        Type vectorType = typeof(PptxRenderer).GetNestedType("ChartIndexedNumberVector", BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected vector type.");
        Type pointType = typeof(PptxRenderer).GetNestedType("ChartIndexedNumberPoint", BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected point type.");
        ConstructorInfo ctor = vectorType.GetConstructors().Single();
        object?[] args = ctor.GetParameters().Select(p => p.ParameterType.IsValueType ? Activator.CreateInstance(p.ParameterType) : null).ToArray();
        args[0] = Array.CreateInstance(pointType, 0);
        args[1] = 10000;
        object vector = ctor.Invoke(args) ?? throw new InvalidOperationException("Expected vector instance.");
        System.Reflection.MethodInfo dense = vectorType.GetMethod("DensePoints") ?? throw new InvalidOperationException("Expected DensePoints.");

        using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxChartRangeCellsPerConversion = 0 }))
        {
            try
            {
                dense.Invoke(vector, null);
            }
            catch (TargetInvocationException ex) when (ex.InnerException is OoxPdfLimitExceededException)
            {
                TestAssert.Equal(0, scope.Budget.ChartRangeCells);
                return;
            }

            throw new InvalidOperationException("Expected dense expansion to hit the range budget.");
        }
    }

    private static string FindCase(string name)
    {
        string[] candidates = new[]
        {
            Path.Combine("tests", "Lokad.OoxPdf.Tests", "Cases", name),
            Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", name),
        };
        foreach (string candidate in candidates)
        {
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        throw new InvalidOperationException("Case file not found: " + name);
    }

    private static int ParseCounter(string summary, string prefix)
    {
        int start = summary.IndexOf(prefix, StringComparison.Ordinal);
        if (start < 0)
        {
            return 0;
        }

        start += prefix.Length;
        int end = start;
        while (end < summary.Length && char.IsDigit(summary[end]))
        {
            end++;
        }

        return int.TryParse(summary.Substring(start, end - start), out int value) ? value : 0;
    }
}
