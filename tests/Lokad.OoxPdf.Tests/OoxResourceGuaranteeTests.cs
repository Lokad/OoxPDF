using System.Reflection;
using System.Text;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
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
        // R06.1: output bytes admit before every write while the scope is still
        // open, so a zero budget trips before publication.
        string input = FindCase("docx-tables.docx");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            ConversionLimits = new OoxConversionLimits { MaxOutputBytesPerConversion = 0 },
        }));
        TestAssert.True(!File.Exists(output), "Budget failure must not publish a partial PDF.");
    }

    public static void PdfOutputZeroBudgetWritesNothingToCountingDestination()
    {
        // R06.1: each output chunk admits against the budget before its write, so a
        // zero budget trips before the destination sees any byte (previously the
        // whole PDF was written, then charged).
        string input = FindCase("docx-tables.docx");
        using var inner = new MemoryStream();
        using var counting = new CountingWriteStream(inner);
        using FileStream inputStream = File.OpenRead(input);
        TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(
            inputStream,
            counting,
            new OoxPdfOptions
            {
                InputKind = OoxPdfInputKind.Docx,
                ConversionLimits = new OoxConversionLimits { MaxOutputBytesPerConversion = 0 },
            }));
        TestAssert.Equal(0L, counting.TotalWritten);
        TestAssert.Equal(0L, inner.Length);
    }

    public static void PdfOutputExactBudgetSucceedsOneByteShortFailsAcrossPaths()
    {
        // R06.1: exact/one-byte-too-small boundaries on file, seekable-stream, and
        // forward-only-stream paths. The one-byte-short file case trips at the tail
        // (after fonts/images/xref) and preserves a pre-existing destination.
        string input = FindCase("docx-tables.docx");
        string ample = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, ample, new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx });
        long size = new FileInfo(ample).Length;
        TestAssert.True(size > 0, "Ample conversion must produce output.");

        string exact = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, exact, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            ConversionLimits = new OoxConversionLimits { MaxOutputBytesPerConversion = size },
        });
        TestAssert.Equal(size, new FileInfo(exact).Length);

        string sentinel = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        byte[] sentinelBytes = new byte[] { 1, 2, 3, 4 };
        File.WriteAllBytes(sentinel, sentinelBytes);
        TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, sentinel, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            ConversionLimits = new OoxConversionLimits { MaxOutputBytesPerConversion = size - 1 },
        }));
        TestAssert.True(File.ReadAllBytes(sentinel).SequenceEqual(sentinelBytes), "Budget failure must preserve the pre-existing destination.");

        using (var output = new MemoryStream())
        {
            using FileStream seekableInput = File.OpenRead(input);
            OoxPdfConverter.Convert(seekableInput, output, new OoxPdfOptions
            {
                InputKind = OoxPdfInputKind.Docx,
                ConversionLimits = new OoxConversionLimits { MaxOutputBytesPerConversion = size },
            });
            TestAssert.Equal(size, output.Length);
        }

        using (var output = new MemoryStream())
        {
            using FileStream seekableInput = File.OpenRead(input);
            TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(seekableInput, output, new OoxPdfOptions
            {
                InputKind = OoxPdfInputKind.Docx,
                ConversionLimits = new OoxConversionLimits { MaxOutputBytesPerConversion = size - 1 },
            }));
        }

        // Forward-only destinations never expose Position/Length; the writer keeps
        // its own count and the exact budget still succeeds.
        using var forwardInner = new MemoryStream();
        using var forwardOnly = new ForwardOnlyWriteStream(forwardInner);
        using FileStream forwardInput = File.OpenRead(input);
        OoxPdfConverter.Convert(forwardInput, forwardOnly, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            ConversionLimits = new OoxConversionLimits { MaxOutputBytesPerConversion = size },
        });
        TestAssert.Equal(size, forwardInner.Length);
    }

    public static void PdfOutputCountingDestinationNeverCrossesBudget()
    {
        // R06.1: a counting destination proves no write crosses the allowed total,
        // including late font/image/xref chunks on an image deck; the destination
        // keeps its allowed prefix after the later failure.
        string input = FindCase("pptx-ladder-07-image-crop.pptx");
        string ample = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, ample, new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx });
        long size = new FileInfo(ample).Length;
        TestAssert.True(size > 0, "Ample conversion must produce output.");

        using var inner = new MemoryStream();
        using var counting = new CountingWriteStream(inner);
        using FileStream inputStream = File.OpenRead(input);
        TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(
            inputStream,
            counting,
            new OoxPdfOptions
            {
                InputKind = OoxPdfInputKind.Pptx,
                ConversionLimits = new OoxConversionLimits { MaxOutputBytesPerConversion = size - 1 },
            }));
        TestAssert.True(counting.TotalWritten <= size - 1, "No write may cross the allowed total, wrote " + counting.TotalWritten + " of " + size + ".");
        TestAssert.True(counting.TotalWritten > 0, "Earlier chunks stay admitted; the failure happens at the tail.");
    }

    public static void PdfOutputThrowingDestinationPropagatesWithoutWrapping()
    {
        // R06.1: admission precedes the write, so a destination failure surfaces as
        // the destination exception itself, never wrapped or mistaken for a budget trip.
        string input = FindCase("docx-tables.docx");
        using FileStream inputStream = File.OpenRead(input);
        using var throwing = new ThrowingWriteStream(new IOException("probe-destination-boom"));
        IOException thrown = TestAssert.Throws<IOException>(() => OoxPdfConverter.Convert(
            inputStream,
            throwing,
            new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx }));
        TestAssert.Contains("probe-destination-boom", thrown.Message);
    }

    public static void PdfObjectWriterAdmitsEveryWriteMethodBeforeWriting()
    {
        // R06.1: direct writer regressions exercise all write methods. A zero budget
        // trips before any byte; an ample budget matches the unscoped byte sequence
        // and the admitted total equals the written total.
        byte[] content = Encoding.ASCII.GetBytes("BT /F1 12 Tf (x) Tj ET");
        byte[] payload = new byte[] { 1, 2, 3, 4, 5 };

        using var ampleInner = new MemoryStream();
        using var ampleCounting = new CountingWriteStream(ampleInner);
        using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(new OoxConversionLimits()))
        {
            var writer = new PdfObjectWriter(ampleCounting, CancellationToken.None);
            writer.WriteHeader();
            writer.WriteObject(1, "<< /Type /Catalog >>\n");
            writer.WriteContentStreamObject(2, content);
            writer.WriteStreamObject(3, "/Filter /FlateDecode", payload);
            writer.WriteAscii("trailer\n");
            TestAssert.Equal(ampleCounting.TotalWritten, scope.Budget.PdfOutputBytes);
        }

        byte[] admitted = ampleInner.ToArray();
        TestAssert.True(admitted.Length > 0, "Ample budget must write the full sequence.");
        using (var plain = new MemoryStream())
        {
            var writer = new PdfObjectWriter(plain, CancellationToken.None);
            writer.WriteHeader();
            writer.WriteObject(1, "<< /Type /Catalog >>\n");
            writer.WriteContentStreamObject(2, content);
            writer.WriteStreamObject(3, "/Filter /FlateDecode", payload);
            writer.WriteAscii("trailer\n");
            TestAssert.True(plain.ToArray().SequenceEqual(admitted), "Admission must not alter emitted bytes.");
        }

        void Attempt(Action<PdfObjectWriter> write)
        {
            using var attemptInner = new MemoryStream();
            using var attemptCounting = new CountingWriteStream(attemptInner);
            using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(
                new OoxConversionLimits { MaxOutputBytesPerConversion = 0 }))
            {
                var writer = new PdfObjectWriter(attemptCounting, CancellationToken.None);
                TestAssert.Throws<OoxPdfLimitExceededException>(() => write(writer));
            }

            TestAssert.Equal(0L, attemptCounting.TotalWritten);
        }

        Attempt(writer => writer.WriteHeader());
        Attempt(writer => writer.WriteObject(1, "<< /Type /Catalog >>\n"));
        Attempt(writer => writer.WriteContentStreamObject(2, content));
        Attempt(writer => writer.WriteStreamObject(3, "/Filter /FlateDecode", payload));
        Attempt(writer => writer.WriteAscii("trailer\n"));

        // Cancellation still precedes admission and writing.
        using var cancelledInner = new MemoryStream();
        using var cancelledCounting = new CountingWriteStream(cancelledInner);
        using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(new OoxConversionLimits()))
        {
            var writer = new PdfObjectWriter(cancelledCounting, new CancellationToken(canceled: true));
            TestAssert.Throws<OperationCanceledException>(() => writer.WriteAscii("x"));
            TestAssert.Equal(0L, scope.Budget.PdfOutputBytes);
        }

        TestAssert.Equal(0L, cancelledCounting.TotalWritten);
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
            () => new OoxConversionLimits { MaxPdfFontBytesPerConversion = -1 },
            () => new OoxConversionLimits { MaxPdfImageBytesPerConversion = -1 },
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

    public static void PdfPagesAndContentScaleLinearly()
    {
        // R06: page and content accounting must scale with page count: identical
        // pages charge identical content bytes (no reset, no superlinear compounding),
        // and conversion allocations grow at most linearly across deck sizes.
        var pageCounts = new List<long>();
        var contentBytes = new List<long>();
        var allocatedBytes = new List<long>();
        foreach (int slides in new[] { 1, 2, 4 })
        {
            string input = WriteSlideDeck(slides);
            string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
            var diagnostics = new List<OoxPdfDiagnostic>();
            long before = GC.GetAllocatedBytesForCurrentThread();
            OoxPdfConverter.Convert(input, output, new OoxPdfOptions { ReportResourceUsage = true, DiagnosticSink = diagnostics.Add });
            allocatedBytes.Add(GC.GetAllocatedBytesForCurrentThread() - before);
            string message = diagnostics.Single(diagnostic => diagnostic.Id == "CONVERSION_RESOURCE_SUMMARY").Message;
            pageCounts.Add(ParseSummaryCounter(message, "pdfPages="));
            contentBytes.Add(ParseSummaryCounter(message, "pdfContentBytes="));
        }

        TestAssert.Equal(1, pageCounts[0]);
        TestAssert.Equal(2, pageCounts[1]);
        TestAssert.Equal(4, pageCounts[2]);
        TestAssert.Equal(2 * contentBytes[0], contentBytes[1]);
        TestAssert.Equal(2 * contentBytes[1], contentBytes[2]);
        TestAssert.True(allocatedBytes[2] <= 2 * allocatedBytes[1], "Expected conversion allocations to grow at most linearly with page count.");
    }

    private static long ParseSummaryCounter(string message, string marker)
    {
        int start = message.IndexOf(marker, StringComparison.Ordinal) + marker.Length;
        int end = message.IndexOf(";", start, StringComparison.Ordinal);
        if (end < 0)
        {
            end = message.IndexOf(".", start, StringComparison.Ordinal);
        }
        return long.Parse(message.Substring(start, end - start), System.Globalization.CultureInfo.InvariantCulture);
    }
    private static string WriteSlideDeck(int slides)
    {
        var sldIds = new StringBuilder();
        var slideRels = new StringBuilder();
        var slideOverrides = new StringBuilder();
        for (int i = 1; i <= slides; i++)
        {
            sldIds.Append($"<p:sldId id='{255 + i}' r:id='rId{i}'/>");
            slideRels.Append($"<Relationship Id='rId{i}' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide' Target='slides/slide{i}.xml'/>");
            slideOverrides.Append($"<Override PartName='/ppt/slides/slide{i}.xml' ContentType='application/vnd.openxmlformats-officedocument.presentationml.slide+xml'/>");
        }
        var presentation = new StringBuilder();
        presentation.Append("<?xml version='1.0' encoding='UTF-8'?>");
        presentation.Append("<p:presentation xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>");
        presentation.Append("<p:sldSz cx='9144000' cy='6858000'/>");
        presentation.Append("<p:sldIdLst>" + sldIds.ToString() + "</p:sldIdLst>");
        presentation.Append("</p:presentation>");
        var presRels = new StringBuilder();
        presRels.Append("<?xml version='1.0' encoding='UTF-8'?>");
        presRels.Append("<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>" + slideRels.ToString() + "</Relationships>");
        var types = new StringBuilder();
        types.Append("<?xml version='1.0' encoding='UTF-8'?>");
        types.Append("<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>");
        types.Append("<Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/>");
        types.Append("<Default Extension='xml' ContentType='application/xml'/>");
        types.Append("<Override PartName='/ppt/presentation.xml' ContentType='application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'/>");
        types.Append(slideOverrides.ToString());
        types.Append("</Types>");
        string slideXml = "<?xml version='1.0' encoding='UTF-8'?>"
            + "<p:sld xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main' xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'>"
            + "<p:cSld><p:spTree><p:sp>"
            + "<p:nvSpPr><p:cNvPr id='2' name='Deck'/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>"
            + "<p:spPr><a:xfrm><a:off x='914400' y='914400'/><a:ext cx='2743200' cy='457200'/></a:xfrm><a:prstGeom prst='rect'/></p:spPr>"
            + "<p:txBody><a:bodyPr/><a:lstStyle/>"
            + "<a:p><a:r><a:rPr sz='2400'><a:latin typeface='Arial'/>"
            + "</a:rPr><a:t>Page</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>";
        var parts = new Dictionary<string, string>();
        parts["[Content_Types].xml"] = types.ToString();
        parts["_rels/.rels"] = PptxTests.PackageRelationship();
        parts["ppt/presentation.xml"] = presentation.ToString();
        parts["ppt/_rels/presentation.xml.rels"] = presRels.ToString();
        for (int i = 1; i <= slides; i++)
        {
            parts["ppt/slides/slide" + i + ".xml"] = slideXml;
        }

        return TestFixtures.WriteTempPackage(".pptx", parts);
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

    public static void PdfFontsRespectFontByteBudget()
    {
        // R06: retained font resources accumulate across embedded fonts. A zero font
        // byte budget trips a font-embedding conversion without publishing output.
        string input = FindCase("docx-tables.docx");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            ConversionLimits = new OoxConversionLimits { MaxPdfFontBytesPerConversion = 0 },
        }));
        TestAssert.True(!File.Exists(output), "Budget failure must not publish a partial PDF.");
    }

    public static void PdfImagesRespectImageByteBudget()
    {
        // R06: retained image resources (including soft masks) accumulate across
        // embedded images. A zero image byte budget trips an image conversion.
        string input = FindCase("pptx-ladder-07-image-crop.pptx");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Pptx,
            ConversionLimits = new OoxConversionLimits { MaxPdfImageBytesPerConversion = 0 },
        }));
        TestAssert.True(!File.Exists(output), "Budget failure must not publish a partial PDF.");
    }

    public static void PdfPageBudgetAdmitsDuringRenderBeforeLaterSlides()
    {
        // R06.2: per-slide page admission trips during rendering, before the last
        // slide (and its image) materializes. With a generous page budget the same
        // deck trips the image budget instead, proving the trailing slide decodes.
        string input = WriteImageSlideDeck(3);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfLimitExceededException imageTrip = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Pptx,
            ConversionLimits = new OoxConversionLimits { MaxImagesDecodedPerConversion = 0 },
        }));
        TestAssert.Contains("image decode budget", imageTrip.Message);
        OoxPdfLimitExceededException pageTrip = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Pptx,
            ConversionLimits = new OoxConversionLimits
            {
                MaxPagesPerConversion = 2,
                MaxImagesDecodedPerConversion = 0,
            },
        }));
        TestAssert.Contains("page budget", pageTrip.Message);
        TestAssert.True(!File.Exists(output), "Budget failure must not publish a partial PDF.");
    }

    public static void PdfPageBudgetAdmitsDuringRenderBeforeLaterDocxPages()
    {
        // R06.2: per-page DOCX admission trips during page emission, before the
        // trailing page (and its image) produces. The image-budget run proves the
        // trailing image decodes when pages are ample.
        string input = WritePagedDocxWithTrailingImage();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfLimitExceededException imageTrip = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            ConversionLimits = new OoxConversionLimits { MaxImagesDecodedPerConversion = 0 },
        }));
        TestAssert.Contains("image decode budget", imageTrip.Message);
        OoxPdfLimitExceededException pageTrip = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            ConversionLimits = new OoxConversionLimits
            {
                MaxPagesPerConversion = 2,
                MaxImagesDecodedPerConversion = 0,
            },
        }));
        TestAssert.Contains("page budget", pageTrip.Message);
        TestAssert.True(!File.Exists(output), "Budget failure must not publish a partial PDF.");
    }

    public static void PdfContentBudgetAdmitsDuringRenderBeforeLaterPages()
    {
        // R06.2: per-page content admission trips on the first emitted page, before
        // the trailing image decodes (previously content charged in the writer after
        // the whole document rendered).
        string input = WritePagedDocxWithTrailingImage();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfLimitExceededException contentTrip = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            ConversionLimits = new OoxConversionLimits
            {
                MaxPdfContentBytesPerConversion = 0,
                MaxImagesDecodedPerConversion = 0,
            },
        }));
        TestAssert.Contains("content byte budget", contentTrip.Message);
        TestAssert.True(!File.Exists(output), "Budget failure must not publish a partial PDF.");
    }

    private sealed class CountingWriteStream(MemoryStream inner) : Stream
    {
        public long TotalWritten { get; private set; }

        public override bool CanRead => false;

        public override bool CanSeek => false;

        public override bool CanWrite => true;

        public override long Length => inner.Length;

        public override long Position { get => inner.Position; set => throw new NotSupportedException(); }

        public override void Flush() => inner.Flush();

        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count)
        {
            inner.Write(buffer, offset, count);
            TotalWritten += count;
        }

        public override void Write(ReadOnlySpan<byte> buffer)
        {
            inner.Write(buffer);
            TotalWritten += buffer.Length;
        }

        public override void WriteByte(byte value)
        {
            inner.WriteByte(value);
            TotalWritten += 1;
        }
    }

    private sealed class ForwardOnlyWriteStream(MemoryStream inner) : Stream
    {
        public override bool CanRead => false;

        public override bool CanSeek => false;

        public override bool CanWrite => inner.CanWrite;

        public override long Length => throw new NotSupportedException();

        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }

        public override void Flush() => inner.Flush();

        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) => inner.Write(buffer, offset, count);

        public override void Write(ReadOnlySpan<byte> buffer) => inner.Write(buffer);

        protected override void Dispose(bool disposing)
        {
        }
    }

    private sealed class ThrowingWriteStream(Exception failure) : Stream
    {
        public override bool CanRead => false;

        public override bool CanSeek => false;

        public override bool CanWrite => true;

        public override long Length => throw new NotSupportedException();

        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }

        public override void Flush()
        {
        }

        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) => throw failure;

        public override void Write(ReadOnlySpan<byte> buffer) => throw failure;
    }

    private static string WriteImageSlideDeck(int slides)
    {
        // R06.2: text slides with a picture on the trailing slide only, so a
        // last-slide producer tripwire (image decode) distinguishes render-time
        // admission from end-of-render charges. Single quotes keep the XML readable.
        var sldIds = new StringBuilder();
        var slideRels = new StringBuilder();
        var slideOverrides = new StringBuilder();
        for (int i = 1; i <= slides; i++)
        {
            sldIds.Append("<p:sldId id=\"" + (255 + i) + "\" r:id=\"rId" + i + "\"/>");
            slideRels.Append("<Relationship Id=\"rId" + i + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide" + i + ".xml\"/>");
            slideOverrides.Append("<Override PartName=\"/ppt/slides/slide" + i + ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>");
        }
        string presentation = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            + "<p:presentation xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
            + "<p:sldSz cx=\"9144000\" cy=\"6858000\"/>"
            + "<p:sldIdLst>" + sldIds.ToString() + "</p:sldIdLst>"
            + "</p:presentation>";
        string presRels = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" + slideRels.ToString() + "</Relationships>";
        string types = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
            + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
            + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
            + "<Default Extension=\"png\" ContentType=\"image/png\"/>"
            + "<Override PartName=\"/ppt/presentation.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml\"/>"
            + slideOverrides.ToString()
            + "</Types>";
        string slideXml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            + "<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
            + "<p:cSld><p:spTree><p:sp>"
            + "<p:nvSpPr><p:cNvPr id=\"2\" name=\"Deck\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>"
            + "<p:spPr><a:xfrm><a:off x=\"914400\" y=\"914400\"/><a:ext cx=\"2743200\" cy=\"457200\"/></a:xfrm><a:prstGeom prst=\"rect\"/></p:spPr>"
            + "<p:txBody><a:bodyPr/><a:lstStyle/>"
            + "<a:p><a:r><a:rPr sz=\"2400\"><a:latin typeface=\"Arial\"/>"
            + "</a:rPr><a:t>Page</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>";
        string imageSlideXml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            + "<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
            + "<p:cSld><p:spTree><p:pic>"
            + "<p:nvPicPr><p:cNvPr id=\"2\" name=\"DeckImage\"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>"
            + "<p:blipFill><a:blip r:embed=\"rIdImage1\"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>"
            + "<p:spPr><a:xfrm><a:off x=\"914400\" y=\"914400\"/><a:ext cx=\"2743200\" cy=\"1371600\"/></a:xfrm><a:prstGeom prst=\"rect\"/></p:spPr>"
            + "</p:pic></p:spTree></p:cSld></p:sld>";
        string imageSlideRels = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rIdImage1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image1.png\"/>"
            + "</Relationships>";
        var parts = new Dictionary<string, string>();
        parts["[Content_Types].xml"] = types;
        parts["_rels/.rels"] = PptxTests.PackageRelationship();
        parts["ppt/presentation.xml"] = presentation;
        parts["ppt/_rels/presentation.xml.rels"] = presRels;
        for (int i = 1; i < slides; i++)
        {
            parts["ppt/slides/slide" + i + ".xml"] = slideXml;
        }

        parts["ppt/slides/slide" + slides + ".xml"] = imageSlideXml;
        parts["ppt/slides/_rels/slide" + slides + ".xml.rels"] = imageSlideRels;
        string path = TestFixtures.WriteTempPackage(".pptx", parts);
        using (var zip = System.IO.Compression.ZipFile.Open(path, System.IO.Compression.ZipArchiveMode.Update))
        {
            var entry = zip.CreateEntry("ppt/media/image1.png");
            using var entryStream = entry.Open();
            byte[] png = TestFixtures.CreateRgbPng(2, 1, new byte[] { 255, 0, 0, 0, 0, 255 });
            entryStream.Write(png, 0, png.Length);
        }

        return path;
    }

    private static string WritePagedDocxWithTrailingImage()
    {
        // R06.2: three forced pages with an inline image on the last page only.
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, byte[]>()
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Default Extension=\"png\" ContentType=\"image/png\"/>"
                + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
                + "</Types>"),
            ["_rels/.rels"] = TestFixtures.Utf8("<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
                + "</Relationships>"),
            ["word/_rels/document.xml.rels"] = TestFixtures.Utf8("<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rIdImage1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image1.png\"/>"
                + "</Relationships>"),
            ["word/document.xml"] = TestFixtures.Utf8("<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""
                + " xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\""
                + " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""
                + " xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\""
                + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                + "<w:body>"
                + "<w:p><w:r><w:t>first</w:t></w:r><w:r><w:br w:type=\"page\"/></w:r></w:p>"
                + "<w:p><w:r><w:t>second</w:t></w:r><w:r><w:br w:type=\"page\"/></w:r></w:p>"
                + "<w:p><w:r><w:drawing><wp:inline><wp:extent cx=\"1828800\" cy=\"914400\"/>"
                + "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
                + "<pic:pic><pic:blipFill><a:blip r:embed=\"rIdImage1\"/></pic:blipFill></pic:pic>"
                + "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"
                + "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>"
                + "</w:body></w:document>"),
            ["word/media/image1.png"] = TestFixtures.CreateRgbPng(2, 1, new byte[] { 255, 0, 0, 0, 0, 255 }),
        });
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
