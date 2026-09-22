using System.Reflection;
using System.Text;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Ooxml;
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

    public static void XmlLoadsRespectAggregateNodeBudget()
    {
        // R04: individually legal XML parses accumulate against one conversion quota
        // instead of each resetting it. Ten small documents (5 objects each) exceed
        // a 10-object aggregate while passing every per-document cap.
        byte[] xml = Encoding.UTF8.GetBytes("<?xml version=\"1.0\" encoding=\"UTF-8\"?><a><b x=\"1\"/><b x=\"2\"/></a>");
        using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxXmlNodesPerConversion = 10 }))
        {
            TestAssert.Throws<OoxPdfLimitExceededException>(() =>
            {
                for (int i = 0; i < 10; i++)
                {
                    using var stream = new MemoryStream(xml, writable: false);
                    SafeXml.Load(stream, CancellationToken.None);
                }
            });
            TestAssert.True(scope.Budget.XmlNodes > 0, "Attempted parses must charge before failing.");
        }

        using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxXmlNodesPerConversion = 1_000 }))
        {
            using var stream = new MemoryStream(xml, writable: false);
            SafeXml.Load(stream, CancellationToken.None);
            // Root plus two empty elements plus two attributes: 5 objects.
            TestAssert.Equal(5, scope.Budget.XmlNodes);
        }
    }

    public static void ChartWorkbooksRespectAggregateCellBudget()
    {
        // R04: cells accumulate across embedded workbooks against one conversion quota.
        byte[] xlsx = MinimalSingleCellWorkbook();
        var resource = new PptxScenePackageResource("/xl/embed.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", xlsx);
        var external = PptxSceneChartExternalData.Defined("rId9", "/xl/embed.xlsx", resource, null, string.Empty);
        MethodInfo getOrCreate = typeof(PptxRenderer).GetMethods(BindingFlags.NonPublic | BindingFlags.Static)
            .First(method => method.Name == "GetOrCreateChartWorkbook");
        object? Invoke(Dictionary<string, PptxRenderer.ChartWorkbookData?>? shared)
        {
            try
            {
                return getOrCreate.Invoke(null, [shared, external, CancellationToken.None]);
            }
            catch (TargetInvocationException ex) when (ex.InnerException is not null)
            {
                throw ex.InnerException;
            }
        }

        using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxWorkbookCellsPerConversion = 0 }))
        {
            TestAssert.Throws<OoxPdfLimitExceededException>(() => Invoke(new Dictionary<string, PptxRenderer.ChartWorkbookData?>(StringComparer.Ordinal)));
            TestAssert.Equal(0, scope.Budget.WorkbookCells);
        }

        using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(null))
        {
            Invoke(new Dictionary<string, PptxRenderer.ChartWorkbookData?>(StringComparer.Ordinal));
            TestAssert.Equal(1, scope.Budget.WorkbookCells);
        }
    }

    private static byte[] MinimalSingleCellWorkbook()
    {
        const string spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        const string officeRels = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const string packageRels = "http://schemas.openxmlformats.org/package/2006/relationships";
        using MemoryStream packageStream = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] =
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
                "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
                "</Types>",
            ["xl/workbook.xml"] =
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<workbook xmlns=\"" + spreadsheet + "\" xmlns:r=\"" + officeRels + "\">" +
                "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>",
            ["xl/_rels/workbook.xml.rels"] =
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<Relationships xmlns=\"" + packageRels + "\">" +
                "<Relationship Id=\"rId1\" Type=\"" + officeRels + "/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>",
            ["xl/worksheets/sheet1.xml"] =
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<worksheet xmlns=\"" + spreadsheet + "\"><sheetData><row r=\"1\"><c r=\"A1\"><v>5</v></c></row></sheetData></worksheet>",
        });
        return packageStream.ToArray();
    }

    public static void PdfPagesRespectPageBudget()
    {
        // R06: page serialization charges against the conversion budget (previously
        // the writer ran outside every budget). A zero page budget trips even a
        // one-page document without publishing output.
        string input = FindCase("docx-tables.docx");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            ConversionLimits = new OoxConversionLimits { MaxPagesPerConversion = 0 },
        }));
        TestAssert.True(!File.Exists(output), "Budget failure must not publish a partial PDF.");

        string ample = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, ample, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            ConversionLimits = new OoxConversionLimits { MaxPagesPerConversion = 100 },
        });
        TestAssert.True(new FileInfo(ample).Length > 0, "Ample page budget must convert.");
    }

    public static void PdfContentRespectContentBudget()
    {
        // R06: encoded content bytes charge up front during serialization. A zero
        // content budget trips conversions that previously sailed through.
        string input = FindCase("docx-tables.docx");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            ConversionLimits = new OoxConversionLimits { MaxPdfContentBytesPerConversion = 0 },
        }));
        TestAssert.True(!File.Exists(output), "Budget failure must not publish a partial PDF.");
    }

    public static void PdfOutputRespectOutputBudget()
    {
        // R06: measured output bytes charge while the scope is still open, after
        // writing but before atomic publication.
        string input = FindCase("docx-tables.docx");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            ConversionLimits = new OoxConversionLimits { MaxOutputBytesPerConversion = 0 },
        }));
        TestAssert.True(!File.Exists(output), "Budget failure must not publish a partial PDF.");
    }

    public static void ConversionLimitsRejectSerializationNegativeCaps()
    {
        // R06 (plus the R04 Validate gap): negative aggregate caps fail fast at
        // option validation instead of tripping spuriously mid-conversion.
        foreach (Func<OoxConversionLimits> capped in new Func<OoxConversionLimits>[]
        {
            () => new OoxConversionLimits { MaxXmlNodesPerConversion = -1 },
            () => new OoxConversionLimits { MaxWorkbookCellsPerConversion = -1 },
            () => new OoxConversionLimits { MaxPagesPerConversion = -1 },
            () => new OoxConversionLimits { MaxPdfContentBytesPerConversion = -1 },
            () => new OoxConversionLimits { MaxOutputBytesPerConversion = -1 },
            () => new OoxConversionLimits { MaxSceneNodesPerConversion = -1 },
            () => new OoxConversionLimits { MaxNestedPackageBytesPerConversion = -1 },
            () => new OoxConversionLimits { MaxWorkbookModelsPerConversion = -1 },
        })
        {
            string input = FindCase("docx-tables.docx");
            string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
            TestAssert.Throws<ArgumentOutOfRangeException>(() => OoxPdfConverter.Convert(
                input,
                output,
                new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx, ConversionLimits = capped() }));
        }
    }

    public static void SceneNodesRespectNodeBudget()
    {
        // R04: scene nodes charge as they materialize across slides, masters, and
        // layouts. A zero node budget trips even a single-picture deck.
        string input = FindCase("pptx-ladder-07-image-crop.pptx");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Pptx,
            ConversionLimits = new OoxConversionLimits { MaxSceneNodesPerConversion = 0 },
        }));
        TestAssert.True(!File.Exists(output), "Budget failure must not publish a partial PDF.");
    }

    public static void NestedPackagesRespectAggregateByteBudget()
    {
        // R04: nested package bytes accumulate across embedded workbooks, each
        // already capped individually. A zero aggregate budget trips the first open.
        byte[] xlsx = MinimalSingleCellWorkbook();
        var resource = new PptxScenePackageResource("/xl/embed.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", xlsx);
        var external = PptxSceneChartExternalData.Defined("rId9", "/xl/embed.xlsx", resource, null, string.Empty);
        MethodInfo getOrCreate = typeof(PptxRenderer).GetMethods(BindingFlags.NonPublic | BindingFlags.Static)
            .First(method => method.Name == "GetOrCreateChartWorkbook");
        object? Invoke(Dictionary<string, PptxRenderer.ChartWorkbookData?>? shared)
        {
            try
            {
                return getOrCreate.Invoke(null, [shared, external, CancellationToken.None]);
            }
            catch (TargetInvocationException ex) when (ex.InnerException is not null)
            {
                throw ex.InnerException;
            }
        }

        using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxNestedPackageBytesPerConversion = 0 }))
        {
            TestAssert.Throws<OoxPdfLimitExceededException>(() => Invoke(new Dictionary<string, PptxRenderer.ChartWorkbookData?>(StringComparer.Ordinal)));
            TestAssert.Equal(0, scope.Budget.NestedPackageBytes);
        }
    }

    public static void ChartWorkbooksRespectModelBudget()
    {
        // R04: retained workbook models accumulate across embedded workbooks while
        // cached models bypass the read without recharging.
        byte[] xlsx = MinimalSingleCellWorkbook();
        var resource = new PptxScenePackageResource("/xl/embed.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", xlsx);
        var external = PptxSceneChartExternalData.Defined("rId9", "/xl/embed.xlsx", resource, null, string.Empty);
        MethodInfo getOrCreate = typeof(PptxRenderer).GetMethods(BindingFlags.NonPublic | BindingFlags.Static)
            .First(method => method.Name == "GetOrCreateChartWorkbook");
        object? Invoke(Dictionary<string, PptxRenderer.ChartWorkbookData?>? shared)
        {
            try
            {
                return getOrCreate.Invoke(null, [shared, external, CancellationToken.None]);
            }
            catch (TargetInvocationException ex) when (ex.InnerException is not null)
            {
                throw ex.InnerException;
            }
        }

        using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxWorkbookModelsPerConversion = 0 }))
        {
            TestAssert.Throws<OoxPdfLimitExceededException>(() => Invoke(new Dictionary<string, PptxRenderer.ChartWorkbookData?>(StringComparer.Ordinal)));
            TestAssert.Equal(0, scope.Budget.WorkbookModels);
        }

        using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(null))
        {
            var cache = new Dictionary<string, PptxRenderer.ChartWorkbookData?>(StringComparer.Ordinal);
            Invoke(cache);
            TestAssert.Equal(1, scope.Budget.WorkbookModels);
            Invoke(cache);
            TestAssert.Equal(1, scope.Budget.WorkbookModels);
        }
    }

    public static void XmlAggregateChargesScaleLinearly()
    {
        // R22: aggregate accounting must scale with input size (N/2N/4N documents
        // charge 5N objects), proving per-document quotas neither reset nor compound
        // superlinearly across a conversion.
        byte[] xml = Encoding.UTF8.GetBytes("<?xml version=\"1.0\" encoding=\"UTF-8\"?><a><b x=\"1\"/><b x=\"2\"/></a>");
        foreach (int documents in new[] { 2, 4, 8 })
        {
            using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxXmlNodesPerConversion = 1_000_000 }))
            {
                for (int i = 0; i < documents; i++)
                {
                    using var stream = new MemoryStream(xml, writable: false);
                    SafeXml.Load(stream, CancellationToken.None);
                }

                TestAssert.Equal(5 * documents, scope.Budget.XmlNodes);
            }
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
