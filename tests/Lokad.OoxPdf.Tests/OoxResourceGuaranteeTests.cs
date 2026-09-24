using System.Reflection;
using System.Text;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
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
            () => new OoxConversionLimits { MaxRetainedImageBytesPerConversion = -1 },
            () => new OoxConversionLimits { MaxRetainedFontBytesPerConversion = -1 },
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
        // RV01: deterministic embeddable test face runs this on every host.
        string input = FindCase("docx-tables.docx");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            FontResolver = new TestFaceFontResolver(),
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
        // RV01: deterministic embeddable test face runs this on every host.
        string input = WritePagedDocxWithTrailingImage();
        var face = new TestFaceFontResolver();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfLimitExceededException imageTrip = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            FontResolver = face,
            ConversionLimits = new OoxConversionLimits { MaxImagesDecodedPerConversion = 0 },
        }));
        TestAssert.Contains("image decode budget", imageTrip.Message);
        OoxPdfLimitExceededException pageTrip = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            FontResolver = face,
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
        // RV01: deterministic embeddable test face runs this on every host.
        string input = WritePagedDocxWithTrailingImage();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfLimitExceededException contentTrip = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            FontResolver = new TestFaceFontResolver(),
            ConversionLimits = new OoxConversionLimits
            {
                MaxPdfContentBytesPerConversion = 0,
                MaxImagesDecodedPerConversion = 0,
            },
        }));
        TestAssert.Contains("content byte budget", contentTrip.Message);
        TestAssert.True(!File.Exists(output), "Budget failure must not publish a partial PDF.");
    }

    public static void RetainedImageBytesTripBeforeSerialization()
    {
        // R06.2: retained production charges land at the owning producer. A zero
        // retained-image budget trips during rendering, before any serialization.
        string input = FindCase("pptx-ladder-07-image-crop.pptx");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfLimitExceededException thrown = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Pptx,
            ConversionLimits = new OoxConversionLimits { MaxRetainedImageBytesPerConversion = 0 },
        }));
        TestAssert.Contains("retained image byte budget", thrown.Message);
        TestAssert.True(!File.Exists(output), "Budget failure must not publish a partial PDF.");
    }
    public static void EffectShadowsRespectRetainedImageBudget()
    {
        // RV10: shadow effect rasters retain encoded bytes plus the soft mask.
        // A zero retained-image budget trips during rendering, before serialization.
        string input = FindCase("pptx-ladder-12-shadow-diagnostic.pptx");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfLimitExceededException thrown = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Pptx,
            ConversionLimits = new OoxConversionLimits { MaxRetainedImageBytesPerConversion = 0 },
        }));
        TestAssert.Contains("retained image byte budget", thrown.Message);
        TestAssert.True(!File.Exists(output), "Budget failure must not publish a partial PDF.");
    }

    public static void EffectGlowsRespectRetainedImageBudget()
    {
        // RV10: same admission for glow rasters, through a synthetic rect shape
        // carrying a glow effect (no tracked fixture exists for glow).
        string input = WriteGlowSlideDeck();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfLimitExceededException thrown = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Pptx,
            ConversionLimits = new OoxConversionLimits { MaxRetainedImageBytesPerConversion = 0 },
        }));
        TestAssert.Contains("retained image byte budget", thrown.Message);
        TestAssert.True(!File.Exists(output), "Budget failure must not publish a partial PDF.");
    }

    public static void EffectShadowBudgetFailureKeepsStreamEmpty()
    {
        // RV10: pre-write stream failures write nothing (R20 outcome contract).
        string input = FindCase("pptx-ladder-12-shadow-diagnostic.pptx");
        using FileStream inputStream = File.OpenRead(input);
        using var output = new MemoryStream();
        TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(
            inputStream,
            output,
            new OoxPdfOptions
            {
                InputKind = OoxPdfInputKind.Pptx,
                ConversionLimits = new OoxConversionLimits { MaxRetainedImageBytesPerConversion = 0 },
            }));
        TestAssert.Equal(0, output.Length);
    }

    public static void RetainedFontBytesTripBeforeSerialization()
    {
        // R06.2: a zero retained-font budget trips at the first subset build.
        // RV01: deterministic embeddable test face runs this on every host.
        string input = FindCase("docx-tables.docx");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfLimitExceededException thrown = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            FontResolver = new TestFaceFontResolver(),
            ConversionLimits = new OoxConversionLimits { MaxRetainedFontBytesPerConversion = 0 },
        }));
        TestAssert.Contains("retained font byte budget", thrown.Message);
        TestAssert.True(!File.Exists(output), "Budget failure must not publish a partial PDF.");
    }
    public static void RepeatedImagesDoNotRechargeRetainedBytes()
    {
        // R06.2: the same image on every slide decodes once (cache hits create
        // nothing), so retained bytes equal the single-use deck.
        long singleBytes = RetainedImageBytesOf(WriteImageSlideDeck(3), OoxPdfInputKind.Pptx);
        long repeatedBytes = RetainedImageBytesOf(WriteImageSlideDeck(3, imageOnEverySlide: true), OoxPdfInputKind.Pptx);
        TestAssert.True(singleBytes > 0, "Image deck must retain image bytes.");
        TestAssert.Equal(singleBytes, repeatedBytes);
    }
    public static void DistinctDocxImagesChargePerProducedPage()
    {
        // RV17: DOCX charges per unique image resource, not per produced page:
        // the same image on two pages charges exactly the single-image deck,
        // while two distinct images charge the sum of their decks.
        long one = RetainedImageBytesOf(WritePagedDocxWithTrailingImage(), OoxPdfInputKind.Docx);
        long twoSame = RetainedImageBytesOf(WritePagedDocxWithTrailingImage(imageOnMiddlePage: true), OoxPdfInputKind.Docx);
        TestAssert.True(one > 0, "Image document must retain image bytes.");
        TestAssert.Equal(one, twoSame);
        long oneBlue = RetainedImageBytesOf(WritePagedDocxWithTrailingImage(trailingEmbed: "rIdImage2"), OoxPdfInputKind.Docx);
        long twoDistinct = RetainedImageBytesOf(WritePagedDocxWithTrailingImage(imageOnMiddlePage: true, distinctMiddleImage: true), OoxPdfInputKind.Docx);
        TestAssert.Equal(one + oneBlue, twoDistinct);
    }
    public static void RetainedBytesReportedOnFilePath()
    {
        // R06.2: retained fields ride the typed file-path summary alongside the
        // serialized writer-stage fields.
        // RV01: deterministic embeddable test face runs this on every host.
        string input = FindCase("docx-tables.docx");
        var diagnostics = new List<OoxPdfDiagnostic>();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            FontResolver = new TestFaceFontResolver(),
            ReportResourceUsage = true,
            DiagnosticSink = diagnostics.Add,
        });
        OoxPdfDiagnostic summary = diagnostics.Single(d => d.Id == "CONVERSION_RESOURCE_SUMMARY");
        TestAssert.True(ParseCounter(summary.Message, "retainedFontBytes=") > 0, "Retained font bytes must be reported, got: " + summary.Message);
        TestAssert.True(summary.Message.Contains("retainedImageBytes=", StringComparison.Ordinal), "Retained image bytes must be reported, got: " + summary.Message);
    }
    private static bool EmbedsFontBytes(string input, OoxPdfInputKind kind)
    {
        var diagnostics = new List<OoxPdfDiagnostic>();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = kind,
            ReportResourceUsage = true,
            DiagnosticSink = diagnostics.Add,
        });
        OoxPdfDiagnostic summary = diagnostics.Single(d => d.Id == "CONVERSION_RESOURCE_SUMMARY");
        return ParseCounter(summary.Message, "pdfFontBytes=") > 0;
    }

    private static long RetainedImageBytesOf(string input, OoxPdfInputKind kind)
    {
        var diagnostics = new List<OoxPdfDiagnostic>();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = kind,
            ReportResourceUsage = true,
            DiagnosticSink = diagnostics.Add,
        });
        OoxPdfDiagnostic summary = diagnostics.Single(d => d.Id == "CONVERSION_RESOURCE_SUMMARY");
        return ParseCounter(summary.Message, "retainedImageBytes=");
    }
    public static void StreamMidWriteFailureKeepsAllowedPrefix()
    {
        // R20: a destination that fails mid-write surfaces its own exception; the
        // caller keeps the allowed prefix and owns the stream.
        var failure = new IOException("probe-mid-write-boom");
        string input = FindCase("docx-tables.docx");
        using FileStream inputStream = File.OpenRead(input);
        using var inner = new MemoryStream();
        using var faulting = new FaultAfterWriteStream(inner, 100, failure);
        IOException thrown = TestAssert.Throws<IOException>(() => OoxPdfConverter.Convert(
            inputStream,
            faulting,
            new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx }));
        TestAssert.Contains("probe-mid-write-boom", thrown.Message);
        TestAssert.True(inner.Length <= 100, "Only the allowed prefix may exist.");
    }
    public static void PdfContentBudgetTripsBeforeSpillingDocx()
    {
        // RV11: per-page content admission lands before the page is handed to
        // the writer, so a zero content cap throws before any spill or encoding.
        string input = FindCase("docx-basic-paragraphs.docx");
        var diagnostics = new List<OoxPdfDiagnostic>();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfLimitExceededException thrown = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            ReportResourceUsage = true,
            DiagnosticSink = diagnostics.Add,
            ConversionLimits = new OoxConversionLimits
            {
                MaxPdfContentBytesPerConversion = 0,
                MaxResidentPageContentBytesPerConversion = 0,
            },
        }));
        TestAssert.Contains("content byte budget", thrown.Message);
        TestAssert.True(!diagnostics.Any(d => d.Id == "PDF_PAGE_CONTENT_SPILLED"), "Content admission must trip before spill work.");
        TestAssert.True(!File.Exists(output), "Budget failure must not publish a partial PDF.");
    }

    public static void PdfContentBudgetTripsBeforeSpillingPptx()
    {
        // RV11: same ordering on the PPTX path.
        string input = FindCase("pptx-ladder-02-plain-text.pptx");
        var diagnostics = new List<OoxPdfDiagnostic>();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfLimitExceededException thrown = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Pptx,
            ReportResourceUsage = true,
            DiagnosticSink = diagnostics.Add,
            ConversionLimits = new OoxConversionLimits
            {
                MaxPdfContentBytesPerConversion = 0,
                MaxResidentPageContentBytesPerConversion = 0,
            },
        }));
        TestAssert.Contains("content byte budget", thrown.Message);
        TestAssert.True(!diagnostics.Any(d => d.Id == "PDF_PAGE_CONTENT_SPILLED"), "Content admission must trip before spill work.");
        TestAssert.True(!File.Exists(output), "Budget failure must not publish a partial PDF.");
    }

    public static void TinyWindowMatchesDefaultBytesOnFileAndStream()
    {
        // R06.3: forced file escalation must emit byte-identical output on both
        // public paths; the window changes residency, never bytes.
        string input = FindCase("docx-tables.docx");
        string defaultPdf = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        string spilledPdf = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, defaultPdf, new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx });
        OoxPdfConverter.Convert(input, spilledPdf, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            ConversionLimits = new OoxConversionLimits { MaxResidentPageContentBytesPerConversion = 1 },
        });
        TestAssert.True(File.ReadAllBytes(defaultPdf).SequenceEqual(File.ReadAllBytes(spilledPdf)), "Window must not alter emitted bytes.");
        using FileStream defaultInput = File.OpenRead(input);
        using FileStream spilledInput = File.OpenRead(input);
        using var defaultOutput = new MemoryStream();
        using var spilledOutput = new MemoryStream();
        OoxPdfConverter.Convert(defaultInput, defaultOutput, new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx });
        OoxPdfConverter.Convert(spilledInput, spilledOutput, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            ConversionLimits = new OoxConversionLimits { MaxResidentPageContentBytesPerConversion = 1 },
        });
        TestAssert.True(defaultOutput.ToArray().SequenceEqual(spilledOutput.ToArray()), "Window must not alter stream bytes.");
    }

    public static void TinyWindowSpillsAndReports()
    {
        // R06.3: a one-byte window spills every content byte; the typed summary
        // reports exactly the serialized content total, and the spill diagnostic fires.
        // The default window spills nothing and stays silent.
        string input = FindCase("docx-tables.docx");
        var spilledDiagnostics = new List<OoxPdfDiagnostic>();
        string spilledPdf = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, spilledPdf, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            ReportResourceUsage = true,
            DiagnosticSink = spilledDiagnostics.Add,
            ConversionLimits = new OoxConversionLimits { MaxResidentPageContentBytesPerConversion = 1 },
        });
        OoxPdfDiagnostic spilledSummary = spilledDiagnostics.Single(d => d.Id == "CONVERSION_RESOURCE_SUMMARY");
        int contentBytes = ParseCounter(spilledSummary.Message, "pdfContentBytes=");
        TestAssert.True(contentBytes > 0, "Content must exist to spill.");
        TestAssert.Equal(contentBytes, ParseCounter(spilledSummary.Message, "pageContentSpilledBytes="));
        TestAssert.True(spilledDiagnostics.Any(d => d.Id == "PDF_PAGE_CONTENT_SPILLED"), "Escalation must notify.");
        var defaultDiagnostics = new List<OoxPdfDiagnostic>();
        string defaultPdf = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, defaultPdf, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            ReportResourceUsage = true,
            DiagnosticSink = defaultDiagnostics.Add,
        });
        OoxPdfDiagnostic defaultSummary = defaultDiagnostics.Single(d => d.Id == "CONVERSION_RESOURCE_SUMMARY");
        TestAssert.Equal(0, ParseCounter(defaultSummary.Message, "pageContentSpilledBytes="));
        TestAssert.True(defaultDiagnostics.All(d => d.Id != "PDF_PAGE_CONTENT_SPILLED"), "Default window must stay silent.");
    }

    public static void SpilledBytesScaleWithPages()
    {
        // R06.3: N/2N proof that spilled volume follows output size (identical slides
        // spill exactly proportional content), i.e. retention follows the window.
        long spilled2 = SpilledBytesOf(WriteSlideDeck(2));
        long spilled4 = SpilledBytesOf(WriteSlideDeck(4));
        TestAssert.True(spilled2 > 0, "Slides must spill content.");
        TestAssert.Equal(2 * spilled2, spilled4);
    }

    private static long SpilledBytesOf(string input)
    {
        var diagnostics = new List<OoxPdfDiagnostic>();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Pptx,
            ReportResourceUsage = true,
            DiagnosticSink = diagnostics.Add,
            ConversionLimits = new OoxConversionLimits { MaxResidentPageContentBytesPerConversion = 1 },
        });
        OoxPdfDiagnostic summary = diagnostics.Single(d => d.Id == "CONVERSION_RESOURCE_SUMMARY");
        return ParseCounter(summary.Message, "pageContentSpilledBytes=");
    }

    public static void StagingCleansTempOnFailure()
    {
        // R06.3: a zero output budget trips during emission after the spill, so the
        // owned spill temp must vanish with the staging; nothing is published.
        string input = FindCase("docx-tables.docx");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            ConversionLimits = new OoxConversionLimits
            {
                MaxResidentPageContentBytesPerConversion = 1,
                MaxOutputBytesPerConversion = 0,
            },
        }));
        TestAssert.True(!File.Exists(output), "Budget failure must not publish a partial PDF.");
    }

    public static void ForwardOnlyStreamWithTinyWindowMatchesDefault()
    {
        // R06.3: forward-only destinations never expose Position/Length; spilled
        // content is re-read from the owned temp, so exact bytes still succeed.
        string input = FindCase("docx-tables.docx");
        using var defaultInner = new MemoryStream();
        using var defaultForward = new ForwardOnlyWriteStream(defaultInner);
        using FileStream defaultInput = File.OpenRead(input);
        OoxPdfConverter.Convert(defaultInput, defaultForward, new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx });
        using var spilledInner = new MemoryStream();
        using var spilledForward = new ForwardOnlyWriteStream(spilledInner);
        using FileStream spilledInput = File.OpenRead(input);
        OoxPdfConverter.Convert(spilledInput, spilledForward, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            ConversionLimits = new OoxConversionLimits { MaxResidentPageContentBytesPerConversion = 1 },
        });
        TestAssert.Equal(defaultInner.Length, spilledInner.Length);
        TestAssert.True(defaultInner.ToArray().SequenceEqual(spilledInner.ToArray()), "Window must not alter forward-only bytes.");
    }

    public static void EmptyPageSequenceThrows()
    {
        // R06.3: staged emission still rejects empty documents like the writer always has.
        using var output = new MemoryStream();
        TestAssert.Throws<ArgumentException>(() => PdfDocumentWriter.WriteStaged(output, Array.Empty<PdfPage>(), new OoxConversionLimits(), diagnosticSink: null, CancellationToken.None));
    }

    public static void PageContentStagingRoundTripsAndCleansTemp()
    {
        // R06.3: store-level proof. Memory mode round-trips; past the window the
        // store escalates to a temp file preserving order, accounts spilled bytes,
        // and deletes the temp on dispose.
        using var staging = new PdfPageContentStaging(10, diagnosticSink: null);
        staging.AddPage(new byte[] { 1, 2, 3 }, CancellationToken.None);
        TestAssert.Equal(false, staging.IsFileBacked);
        staging.AddPage(new byte[] { 4, 5, 6, 7, 8, 9, 10, 11 }, CancellationToken.None);
        TestAssert.Equal(true, staging.IsFileBacked);
        string? path = staging.SpillPathForTests;
        TestAssert.True(path is not null, "Escalation must create a temp file.");
        TestAssert.True(staging.GetPageBytes(0, CancellationToken.None).SequenceEqual(new byte[] { 1, 2, 3 }), "First page must round-trip.");
        TestAssert.True(staging.GetPageBytes(1, CancellationToken.None).SequenceEqual(new byte[] { 4, 5, 6, 7, 8, 9, 10, 11 }), "Second page must round-trip.");
        TestAssert.Equal(2, staging.Count);
        TestAssert.Equal(11L, staging.SpilledBytes);
        staging.Dispose();
        TestAssert.True(!File.Exists(path), "Spill temp must be deleted on dispose.");
    }

    // RV11-P1: with a 1 KiB resident window, emitting N/2N/4N single-page
    // payloads must not transiently allocate whole pages: staging scratch
    // stays flat while payloads quadruple (output itself is discarded).
    public static void StagedEmissionScratchStaysFlatAtN2N4N()
    {
        foreach (int size in new[] { 262144, 524288, 1048576 })
        {
            var pages = new[] { new PdfPage(612, 792, new string((char)10, size)) };
            var limits = new OoxConversionLimits { MaxResidentPageContentBytesPerConversion = 1024 };
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            long before = GC.GetAllocatedBytesForCurrentThread();
            using var output = new DiscardingWriteStream();
            PdfDocumentWriter.WriteStaged(output, pages, limits, diagnosticSink: null, CancellationToken.None);
            long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
            TestAssert.True(output.TotalWritten >= size, "Emission must write the page payload.");
            TestAssert.True(allocated <= 393216L, string.Format("Staged emission scratch must stay flat, allocated {0} bytes for a {1}-byte page.", allocated, size));
        }
    }

    // RV11-P1: spilled pages must support bounded partial reads that
    // reassemble exactly, in memory and file modes, with cancellation.
    public static void SpilledPagePartialReadsReassemble()
    {
        byte[] first = new byte[] { 1, 2, 3 };
        byte[] second = new byte[300];
        for (int i = 0; i < second.Length; i++)
        {
            second[i] = (byte)(i % 251);
        }

        byte[] expected = first.Concat(second).ToArray();
        foreach (long window in new[] { 1L, 100000L })
        {
            using var staging = new PdfPageContentStaging(window, diagnosticSink: null);
            InvokeBeginPage(staging);
            InvokeAppendPageBytes(staging, first);
            InvokeAppendPageBytes(staging, second);
            InvokeEndPage(staging);
            TestAssert.Equal(expected.Length, InvokeGetPageLength(staging, 0));
            var reassembled = new List<byte>();
            byte[] buffer = new byte[7];
            int offset = 0;
            int read;
            while ((read = InvokeReadPageBytes(staging, buffer, offset, CancellationToken.None)) > 0)
            {
                reassembled.AddRange(buffer.Take(read));
                offset += read;
            }

            TestAssert.True(expected.SequenceEqual(reassembled), "Partial reads must reassemble the page exactly.");
            TestAssert.Equal(0, InvokeReadPageBytes(staging, buffer, expected.Length, CancellationToken.None));
            using var cancelled = new CancellationTokenSource();
            cancelled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => InvokeReadPageBytes(staging, buffer, 0, cancelled.Token));
        }
    }

    // RV11-P1: chunked production (odd splits, empty appends, empty pages)
    // must match whole-page appends byte for byte, in both backing modes.
    public static void ChunkedPageProductionMatchesWholePage()
    {
        byte[] payload = new byte[1000];
        for (int i = 0; i < payload.Length; i++)
        {
            payload[i] = (byte)(i % 251);
        }

        foreach (long window in new[] { 10L, 100000L })
        {
            using var whole = new PdfPageContentStaging(window, diagnosticSink: null);
            whole.AddPage(payload, CancellationToken.None);
            using var chunked = new PdfPageContentStaging(window, diagnosticSink: null);
            chunked.BeginPage(CancellationToken.None);
            chunked.AppendPageBytes(payload, 0, 0, CancellationToken.None);
            int offset = 0;
            foreach (int size in new[] { 1, 2, 3, 5, 8, 13, 21, 34, 55, 89 })
            {
                int take = Math.Min(size, payload.Length - offset);
                if (take == 0)
                {
                    break;
                }

                chunked.AppendPageBytes(payload, offset, take, CancellationToken.None);
                offset += take;
            }

            if (offset < payload.Length)
            {
                chunked.AppendPageBytes(payload, offset, payload.Length - offset, CancellationToken.None);
            }

            chunked.EndPage(CancellationToken.None);
            chunked.BeginPage(CancellationToken.None);
            chunked.EndPage(CancellationToken.None);
            whole.BeginPage(CancellationToken.None);
            whole.EndPage(CancellationToken.None);
            TestAssert.True(whole.GetPageBytes(0, CancellationToken.None).SequenceEqual(chunked.GetPageBytes(0, CancellationToken.None)), "Chunked production must match whole-page bytes.");
            TestAssert.Equal(0, chunked.GetPageLength(1, CancellationToken.None));
            TestAssert.Equal(whole.Count, chunked.Count);
        }
    }

    // RV11-P1: exact window boundaries. A page that fits stays resident; one
    // byte more escalates to the temp file.
    public static void WindowBoundariesStayMemoryOrSpill()
    {
        using var fitting = new PdfPageContentStaging(10, diagnosticSink: null);
        fitting.AddPage(new byte[10], CancellationToken.None);
        TestAssert.Equal(false, fitting.IsFileBacked);
        TestAssert.True(fitting.GetPageBytes(0, CancellationToken.None).SequenceEqual(new byte[10]), "Fitting page must round-trip.");
        using var overflowing = new PdfPageContentStaging(10, diagnosticSink: null);
        overflowing.AddPage(new byte[11], CancellationToken.None);
        TestAssert.Equal(true, overflowing.IsFileBacked);
        TestAssert.True(overflowing.GetPageBytes(0, CancellationToken.None).SequenceEqual(new byte[11]), "Spilled page must round-trip.");
    }

    // RV11-P1: unpaired production or reads during an open page fail loudly
    // instead of silently misaligning page indexes.
    public static void UnpairedStagingCallsFailLoudly()
    {
        using var staging = new PdfPageContentStaging(100000, diagnosticSink: null);
        TestAssert.Throws<InvalidOperationException>(() => staging.AppendPageBytes(new byte[1], 0, 1, CancellationToken.None));
        TestAssert.Throws<InvalidOperationException>(() => staging.EndPage(CancellationToken.None));
        staging.BeginPage(CancellationToken.None);
        TestAssert.Throws<InvalidOperationException>(() => staging.BeginPage(CancellationToken.None));
        TestAssert.Throws<InvalidOperationException>(() => staging.GetPageLength(0, CancellationToken.None));
        staging.EndPage(CancellationToken.None);
        TestAssert.Throws<InvalidOperationException>(() => staging.EndPage(CancellationToken.None));
    }

    // RV11-P1: a throwing spill observer propagates its own failure and the
    // owned temp still vanishes with the staging.
    public static void ThrowingSpillObserverPropagatesAndCleansTemp()
    {
        static void ThrowingSink(OoxPdfDiagnostic diagnostic) => throw new InvalidOperationException("observer boom");
        using var staging = new PdfPageContentStaging(10, ThrowingSink);
        staging.AddPage(new byte[] { 1, 2, 3 }, CancellationToken.None);
        InvalidOperationException thrown = TestAssert.Throws<InvalidOperationException>(() => staging.AddPage(new byte[100], CancellationToken.None));
        TestAssert.Equal("observer boom", thrown.Message);
        string? path = staging.SpillPathForTests;
        TestAssert.True(path is not null, "Escalation must have created a temp file.");
        staging.Dispose();
        TestAssert.True(!File.Exists(path), "Spill temp must be deleted on dispose.");
    }

    // RV11-P1: chunked ASCII encoding is byte-identical to whole-string
    // encoding, including non-ASCII replacements straddling chunk edges.
    public static void ChunkedAsciiEncodingMatchesWholeString()
    {
        char[] chars = new string((char)97, 200000).ToCharArray();
        chars[65534] = (char)233;
        chars[65535] = (char)233;
        chars[65536] = (char)233;
        chars[199999] = (char)200;
        string content = new string(chars);
        byte[] expected = Encoding.ASCII.GetBytes(content);
        using var staging = new PdfPageContentStaging(100000000, diagnosticSink: null);
        staging.BeginPage(CancellationToken.None);
        try
        {
            System.Reflection.MethodInfo encode = typeof(PdfDocumentWriter).GetMethod("AppendEncodedPageContent", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)
                ?? throw new InvalidOperationException("Expected chunked encoder.");
            encode.Invoke(null, [staging, content, CancellationToken.None]);
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }

        staging.EndPage(CancellationToken.None);
        TestAssert.True(expected.SequenceEqual(staging.GetPageBytes(0, CancellationToken.None)), "Chunked encoding must match whole-string encoding.");
    }

    private static void InvokeBeginPage(PdfPageContentStaging staging)
    {
        try
        {
            System.Reflection.MethodInfo begin = typeof(PdfPageContentStaging).GetMethod("BeginPage")
                ?? throw new InvalidOperationException("Expected chunked page production.");
            begin.Invoke(staging, [CancellationToken.None]);
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
    }

    private static void InvokeAppendPageBytes(PdfPageContentStaging staging, byte[] chunk)
    {
        try
        {
            System.Reflection.MethodInfo append = typeof(PdfPageContentStaging).GetMethod("AppendPageBytes")
                ?? throw new InvalidOperationException("Expected chunked page production.");
            append.Invoke(staging, [chunk, 0, chunk.Length, CancellationToken.None]);
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
    }

    private static void InvokeEndPage(PdfPageContentStaging staging)
    {
        try
        {
            System.Reflection.MethodInfo end = typeof(PdfPageContentStaging).GetMethod("EndPage")
                ?? throw new InvalidOperationException("Expected chunked page production.");
            end.Invoke(staging, [CancellationToken.None]);
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
    }

    private static int InvokeGetPageLength(PdfPageContentStaging staging, int pageIndex)
    {
        try
        {
            System.Reflection.MethodInfo length = typeof(PdfPageContentStaging).GetMethod("GetPageLength")
                ?? throw new InvalidOperationException("Expected chunked page reads.");
            return (int)length.Invoke(staging, [pageIndex, CancellationToken.None])!;
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
    }

    private static int InvokeReadPageBytes(PdfPageContentStaging staging, byte[] buffer, int sourceOffset, CancellationToken cancellationToken)
    {
        try
        {
            System.Reflection.MethodInfo read = typeof(PdfPageContentStaging).GetMethod("ReadPageBytes")
                ?? throw new InvalidOperationException("Expected chunked page reads.");
            return (int)read.Invoke(staging, [0, buffer, 0, sourceOffset, buffer.Length, cancellationToken])!;
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
    }

    private sealed class DiscardingWriteStream : Stream
    {
        public long TotalWritten { get; private set; }

        public override bool CanRead => false;

        public override bool CanSeek => false;

        public override bool CanWrite => true;

        public override long Length => TotalWritten;

        public override long Position { get => TotalWritten; set => throw new NotSupportedException(); }

        public override void Flush()
        {
        }

        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) => TotalWritten += count;

        public override void Write(ReadOnlySpan<byte> buffer) => TotalWritten += buffer.Length;
    }

    public static void RepaginationDoesNotRechargePages()
    {
        // R06.2: header displacement forces a second full layout pass, but page charges
        // apply once per emitted final page: an exact page budget succeeds. The model-level
        // check below proves the second pass engages (nonzero displacement); without it the
        // budget assertion would hold vacuously.
        string input = WriteDisplacingHeaderDocx();
        using FileStream probeStream = File.OpenRead(input);
        OoxPackage probePackage = OoxPackage.Open(probeStream, CancellationToken.None);
        DocxDocument probeDocument = new DocxReader().Read(probePackage, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxLayout probeLayout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(probeDocument, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxLayoutPage firstPage = probeLayout.Pages[0];
        double headerBottom = probeLayout.HeaderContentBottomByPage.TryGetValue(0, out double bottom) ? bottom : -1d;
        TestAssert.True(headerBottom < firstPage.Height - firstPage.MarginTop, "Fixture must displace headers into a second pass.");
        var diagnostics = new List<OoxPdfDiagnostic>();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            ReportResourceUsage = true,
            DiagnosticSink = diagnostics.Add,
            ConversionLimits = new OoxConversionLimits { MaxPagesPerConversion = 1 },
        });
        OoxPdfDiagnostic summary = diagnostics.Single(d => d.Id == "CONVERSION_RESOURCE_SUMMARY");
        TestAssert.Contains("pdfPages=1", summary.Message);
        TestAssert.True(new FileInfo(output).Length > 0, "Exact budget must publish.");
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

    private static string WriteImageSlideDeck(int slides, bool imageOnEverySlide = false)
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
        for (int i = 1; i <= slides; i++)
        {
            bool imageSlide = imageOnEverySlide || i == slides;
            parts["ppt/slides/slide" + i + ".xml"] = imageSlide ? imageSlideXml : slideXml;
            if (imageSlide)
            {
                parts["ppt/slides/_rels/slide" + i + ".xml.rels"] = imageSlideRels;
            }
        }
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

    private static string WriteGlowSlideDeck()
    {
        // RV10: single slide with a rect shape carrying a glow effect, so the
        // raster glow producer runs without any tracked glow fixture.
        string presentation = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            + "<p:presentation xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
            + "<p:sldSz cx=\"9144000\" cy=\"6858000\"/>"
            + "<p:sldIdLst><p:sldId id=\"256\" r:id=\"rId1\"/></p:sldIdLst>"
            + "</p:presentation>";
        string presRels = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide1.xml\"/>"
            + "</Relationships>";
        string types = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
            + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
            + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
            + "<Override PartName=\"/ppt/presentation.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml\"/>"
            + "<Override PartName=\"/ppt/slides/slide1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>"
            + "</Types>";
        string slideXml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            + "<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
            + "<p:cSld><p:spTree><p:sp>"
            + "<p:nvSpPr><p:cNvPr id=\"2\" name=\"Glow\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>"
            + "<p:spPr><a:xfrm><a:off x=\"914400\" y=\"914400\"/><a:ext cx=\"2743200\" cy=\"1371600\"/></a:xfrm>"
            + "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>"
            + "<a:effectLst><a:glow rad=\"254000\"><a:srgbClr val=\"FF0000\"><a:alpha val=\"75000\"/></a:srgbClr></a:glow></a:effectLst>"
            + "</p:spPr>"
            + "<p:txBody><a:bodyPr/><a:lstStyle/>"
            + "<a:p><a:r><a:rPr sz=\"1800\"><a:latin typeface=\"Arial\"/>"
            + "</a:rPr><a:t>Glow</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>";
        var parts = new Dictionary<string, string>();
        parts["[Content_Types].xml"] = types;
        parts["_rels/.rels"] = PptxTests.PackageRelationship();
        parts["ppt/presentation.xml"] = presentation;
        parts["ppt/_rels/presentation.xml.rels"] = presRels;
        parts["ppt/slides/slide1.xml"] = slideXml;
        return TestFixtures.WriteTempPackage(".pptx", parts);
    }

    private static string WritePagedDocxWithTrailingImage(bool imageOnMiddlePage = false, bool distinctMiddleImage = false, string trailingEmbed = "rIdImage1")
    {
        // R06.2/RV17: three forced pages with an inline image on the last page
        // (and optionally the middle page, same or distinct part, to pin
        // charge-on-new-resource semantics).
        string middleEmbed = distinctMiddleImage ? "rIdImage2" : "rIdImage1";
        string middleImage = imageOnMiddlePage
            ? "<w:p><w:r><w:drawing><wp:inline><wp:extent cx=\"1828800\" cy=\"914400\"/>"
                + "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
                + "<pic:pic><pic:blipFill><a:blip r:embed=\"" + middleEmbed + "\"/></pic:blipFill></pic:pic>"
                + "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"
            : string.Empty;
        string secondPage = "<w:p><w:r><w:t>second</w:t></w:r><w:r><w:br w:type=\"page\"/></w:r></w:p>" + middleImage;
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
                + "<Relationship Id=\"rIdImage2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image2.png\"/>"
                + "</Relationships>"),
            ["word/document.xml"] = TestFixtures.Utf8("<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""
                + " xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\""
                + " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""
                + " xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\""
                + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                + "<w:body>"
                + "<w:p><w:r><w:t>first</w:t></w:r><w:r><w:br w:type=\"page\"/></w:r></w:p>"
                + secondPage
                + "<w:p><w:r><w:drawing><wp:inline><wp:extent cx=\"1828800\" cy=\"914400\"/>"
                + "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
                + "<pic:pic><pic:blipFill><a:blip r:embed=\"" + trailingEmbed + "\"/></pic:blipFill></pic:pic>"
                + "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"
                + "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>"
                + "</w:body></w:document>"),
            ["word/media/image1.png"] = TestFixtures.CreateRgbPng(2, 1, new byte[] { 255, 0, 0, 0, 0, 255 }),
            ["word/media/image2.png"] = TestFixtures.CreateRgbPng(2, 1, new byte[] { 0, 255, 0, 0, 0, 255 }),
        });
    }

    private sealed class FaultAfterWriteStream(MemoryStream inner, long maxBytes, Exception failure) : Stream
    {
        private long written;
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() => inner.Flush();
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => Write(new ReadOnlySpan<byte>(buffer, offset, count));
        public override void Write(ReadOnlySpan<byte> buffer)
        {
            if (written + buffer.Length > maxBytes)
            {
                throw failure;
            }
            inner.Write(buffer);
            written += buffer.Length;
        }
    }
    private static string WriteDisplacingHeaderDocx()
    {
        var header = new System.Text.StringBuilder();
        for (int i = 1; i <= 8; i++)
        {
            header.Append("<w:p><w:r><w:t>Header line " + i + "</w:t></w:r></w:p>");
        }
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
                + "<Override PartName=\"/word/header1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/>"
                + "</Types>",
            ["_rels/.rels"] = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
                + "</Relationships>",
            ["word/_rels/document.xml.rels"] = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rIdHeader1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\"/>"
                + "</Relationships>",
            ["word/header1.xml"] = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
                + header.ToString()
                + "</w:hdr>",
            ["word/document.xml"] = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                + "<w:body>"
                + "<w:p><w:r><w:t>Body text</w:t></w:r></w:p>"
                + "<w:sectPr><w:headerReference w:type=\"default\" r:id=\"rIdHeader1\"/><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>"
                + "</w:body></w:document>",
        });
    }
    // RV17: one header image repeated across pages must decode (and charge)
    // once, not once per page.
    public static void RepeatedHeaderImageDecodesOnce()
    {
        string input = WriteHeaderImageDocx(6);
        var diagnostics = new List<OoxPdfDiagnostic>();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            ReportResourceUsage = true,
            DiagnosticSink = diagnostics.Add,
        });
        OoxPdfDiagnostic summary = diagnostics.Single(d => d.Id == "CONVERSION_RESOURCE_SUMMARY");
        TestAssert.Equal(1, ParseCounter(summary.Message, "imagesDecoded="));
        TestAssert.Contains("pdfPages=6", summary.Message);
        string pdfText = System.Text.Encoding.Latin1.GetString(File.ReadAllBytes(output));
        TestAssert.True(System.Text.RegularExpressions.Regex.Matches(pdfText, @"/Subtype /Image\b").Count >= 1, "Header image must still emit an image object.");
    }

    private static string WriteHeaderImageDocx(int pages)
    {
        byte[] png = TestFixtures.CreateRgbPng(8, 8, new byte[8 * 8 * 3]);
        var body = new System.Text.StringBuilder();
        for (int i = 0; i < pages; i++)
        {
            body.Append(i == 0
                ? "<w:p><w:r><w:t>Body page " + i + "</w:t></w:r></w:p>"
                : "<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Body page " + i + "</w:t></w:r></w:p>");
        }

        string documentTemplate = """
            <?xml version="1.0" encoding="UTF-8"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <w:body>
            [[BODY]]
                <w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
              </w:body>
            </w:document>
            """;

        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """),
            ["word/_rels/document.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
                </Relationships>
                """),
            ["word/_rels/header1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
                </Relationships>
                """),
            ["word/header1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                       xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:p><w:r><w:drawing>
                    <wp:inline>
                      <wp:extent cx="914400" cy="914400"/>
                      <a:graphic>
                        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                          <pic:pic><pic:blipFill><a:blip r:embed="rIdImage1"/></pic:blipFill></pic:pic>
                        </a:graphicData>
                      </a:graphic>
                    </wp:inline>
                  </w:drawing></w:r></w:p>
                </w:hdr>
                """),
            ["word/document.xml"] = TestFixtures.Utf8(documentTemplate.Replace("[[BODY]]", body.ToString())),
            ["word/media/image1.png"] = png,
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
