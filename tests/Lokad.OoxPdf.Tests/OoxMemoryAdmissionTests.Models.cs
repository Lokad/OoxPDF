using System.IO.Compression;
using System.Reflection;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    public static void ModelAdmissionDocxStorageOutlivesItsPackage()
    {
        var pool = new TestMemoryBudget();
        using var input = new MemoryStream(XmlAdmissionPublicInput(true));
        using (var operation = OoxConversionBudget.BeginScope(null, pool))
        {
            var package = OoxPackage.Open(input, default);
            package.GetRelationships("/", default);
            package.LoadXml(TestAssert.NotNull(package.GetPart("/word/document.xml")), default);
            long encodedAndXml = pool.Reserved;
            DocxDocument document = new DocxReader().Read(package, null, default, OoxPdfDocxMarkupMode.Final);
            TestAssert.True(pool.Reserved > encodedAndXml, "A materialized document has capacity beyond its encoded source and DOM.");
            TestAssert.True(pool.Peak > pool.Reserved, "Model construction headroom refunds while retained capacity remains.");
            package.Dispose();
            TestAssert.True(pool.Reserved > 0, "The conversion owns derived models independently of their package.");
            TestAssert.Equal(1025, document.Paragraphs.Count);
            long held = pool.Reserved;
            var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var first = engine.Create(document, (IDocxTextMeasurer?)null, default);
            long placed = pool.Reserved;
            TestAssert.True(placed > held, "Page frames have reusable capacity beyond source models.");
            var second = engine.Create(document, (IDocxTextMeasurer?)null, default);
            TestAssert.Equal(first.Pages.Count, second.Pages.Count);
            TestAssert.Equal(placed, pool.Reserved);
            TestAssert.True(input.CanRead, "Package disposal preserves caller input ownership.");
        }
        TestAssert.Equal(0L, pool.Reserved);
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        TestAssert.Equal(0, pool.CallerDisposals);
    }

    public static void ModelAdmissionRefusalAndCancellationPrecedeDocxConstruction()
    {
        foreach (bool cancel in new[] { false, true })
        {
            var pool = new TestMemoryBudget();
            using var cancellation = new CancellationTokenSource();
            using var input = new MemoryStream(XmlAdmissionPublicInput(true));
            using (var operation = OoxConversionBudget.BeginScope(null, pool))
            using (var package = OoxPackage.Open(input, default))
            {
                package.GetRelationships("/", default);
                package.LoadXml(TestAssert.NotNull(package.GetPart("/word/document.xml")), default);
                long prior = pool.Reserved;
                bool requested = false;
                pool.RefuseChange = (before, after) => {
                    if (after <= before) return false;
                    requested = true;
                    if (cancel) cancellation.Cancel();
                    return !cancel;
                };
                if (cancel)
                    TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(() =>
                        new DocxReader().Read(package, null, cancellation.Token, OoxPdfDocxMarkupMode.Final)).CancellationToken);
                else
                    TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, TestAssert.Throws<OoxPdfLimitExceededException>(() =>
                        new DocxReader().Read(package, null, default, OoxPdfDocxMarkupMode.Final)).Kind);
                TestAssert.True(requested, "Warm XML does not bypass derived-model admission.");
                TestAssert.Equal(prior, pool.Reserved);
                pool.RefuseChange = null;
                TestAssert.Equal(1025, new DocxReader().Read(package, null, default, OoxPdfDocxMarkupMode.Final).Paragraphs.Count);
            }
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        }
    }

    public static void ModelAdmissionWorkbookOutlivesItsNestedPackage()
    {
        using var input = TestFixtures.CreateZipPackage(new Dictionary<string, string> {
            ["[Content_Types].xml"] = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"xml\" ContentType=\"application/xml\"/><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/></Types>",
            ["xl/workbook.xml"] = "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"Data\" sheetId=\"1\" r:id=\"s1\"/></sheets></workbook>",
            ["xl/_rels/workbook.xml.rels"] = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"s1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>",
            ["xl/worksheets/sheet1.xml"] = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"1\"><c r=\"A1\"><v>42</v></c></row></sheetData></worksheet>"
        });
        var pool = new TestMemoryBudget();
        using (var operation = OoxConversionBudget.BeginScope(null, pool))
        {
            PptxRenderer.ChartWorkbookData data;
            using (var package = OoxPackage.Open(input, default))
            {
                var method = typeof(PptxRenderer).GetMethod("ReadWorkbookDataCore", BindingFlags.NonPublic | BindingFlags.Static)!;
                try { data = (PptxRenderer.ChartWorkbookData)TestAssert.NotNull(method.Invoke(null, [package, CancellationToken.None])); }
                catch (TargetInvocationException ex) { throw ex.InnerException ?? ex; }
            }
            TestAssert.True(pool.Reserved > 0, "Workbook records stay funded after nested package buffers refund.");
            TestAssert.Equal(1, data.ReadRangeCells("Data!A1:A1").Length);
        }
        TestAssert.Equal(0L, pool.Reserved);
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        TestAssert.Equal(0, pool.CallerDisposals);
    }

    public static void ModelAdmissionPublicRefusalPreservesDestinations()
    {
        string directory = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ooxpdf-model-refusal-" + Guid.NewGuid().ToString("N"))).FullName;
        try
        {
            foreach (bool docx in new[] { true, false })
            {
                byte[] bytes = XmlAdmissionPublicInput(docx);
                // Allow the measured encoded/XML intake peak plus a fixed staging
                // window. Refusal then challenges additional model capacity.
                var preflight = new TestMemoryBudget();
                using (var operation = OoxConversionBudget.BeginScope(null, preflight))
                using (var source = new MemoryStream(bytes))
                using (var package = OoxPackage.Open(source, default))
                    foreach (var part in package.Parts.Where(p => p.Name.EndsWith(".xml") || p.Name.EndsWith(".rels"))) package.LoadXml(part, default);
                long capacity = preflight.Peak + 256 * 1024;
                string sourcePath = Path.Combine(directory, docx ? "source.docx" : "source.pptx");
                File.WriteAllBytes(sourcePath, bytes);
                foreach (bool file in new[] { false, true })
                {
                    bool acquiredFont = false;
                    var pool = new TestMemoryBudget(capacity);
                    var options = new OoxPdfOptions { InputKind = docx ? OoxPdfInputKind.Docx : OoxPdfInputKind.Pptx,
                        MemoryBudget = pool, FontResolver = new ProgramResolver(new ObservedFontProgram(() => acquiredFont = true)) };
                    string target = Path.Combine(directory, "result.pdf");
                    File.WriteAllBytes(target, [1, 2, 3]);
                    using var source = new MemoryStream(bytes);
                    using var output = new MemoryStream();
                    Task conversion = file ? OoxPdfConverter.ConvertAsync(sourcePath, target, options) : OoxPdfConverter.ConvertAsync(source, output, options);
                    var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => conversion.GetAwaiter().GetResult());
                    TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
                    TestAssert.True(!acquiredFont, "Model capacity refuses before font preparation.");
                    TestAssert.Equal(0L, output.Length);
                    TestAssert.True(File.ReadAllBytes(target).SequenceEqual(new byte[] { 1, 2, 3 }), "Refusal preserves the destination.");
                    TestAssert.True(source.CanRead && output.CanWrite, "Caller streams remain open.");
                    TestAssert.Equal(0, Directory.GetFiles(directory, "*.tmp-*").Length);
                    TestAssert.Equal(0L, pool.Reserved);
                    TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
                }
            }
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    public static void ModelAdmissionRefundsBeforeStreamSummaryAndKeepsPdfBytes()
    {
        foreach (bool docx in new[] { true, false })
        {
            byte[] bytes = XmlAdmissionPublicInput(docx);
            var kind = docx ? OoxPdfInputKind.Docx : OoxPdfInputKind.Pptx;
            byte[] expected;
            using (var input = new MemoryStream(bytes))
            using (var output = new MemoryStream())
            {
                OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = kind, FixedCreationDate = DateTimeOffset.UnixEpoch,
                    FontResolver = new ProgramResolver(new ObservedFontProgram(() => { })) });
                expected = output.ToArray();
            }
            var pool = new TestMemoryBudget();
            long duringModel = 0;
            bool summarized = false;
            using var source = new MemoryStream(bytes);
            using var destination = new MemoryStream();
            OoxPdfConverter.Convert(source, destination, new OoxPdfOptions { InputKind = kind, MemoryBudget = pool,
                FixedCreationDate = DateTimeOffset.UnixEpoch, ReportResourceUsage = true,
                FontResolver = new ProgramResolver(new ObservedFontProgram(() => duringModel = Math.Max(duringModel, pool.Reserved))),
                DiagnosticSink = d => {
                    if (d.Id != "CONVERSION_RESOURCE_SUMMARY") return;
                    summarized = true;
                    TestAssert.True(pool.Reserved > 0 && pool.Reserved < duringModel, "Model capacity refunds after production while source/PDF storage remains funded.");
                    TestAssert.Equal(0L, destination.Length);
                }});
            TestAssert.True(summarized, "Observe the real staging/serialization boundary.");
            TestAssert.True(expected.AsSpan().SequenceEqual(destination.ToArray()), "Governed conversion preserves exact output.");
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        }
    }

    public static void ModelAdmissionPendingFontBorrowKeepsCapacityThroughCancellation()
    {
        using var cancellation = new CancellationTokenSource();
        var program = new PendingModelFontProgram();
        var pool = new TestMemoryBudget();
        using var input = new MemoryStream(XmlAdmissionPublicInput(true));
        using var output = new MemoryStream();
        Task conversion = OoxPdfConverter.ConvertAsync(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx,
            MemoryBudget = pool, FontResolver = new ProgramResolver(program) }, cancellation.Token);
        try
        {
            program.Entered.Task.WaitAsync(TimeSpan.FromSeconds(5)).GetAwaiter().GetResult();
            long held = pool.Reserved;
            TestAssert.True(held > 0, "Keep model/source capacity through awaited font acquisition.");
            cancellation.Cancel();
            Task winner = Task.WhenAny(conversion, Task.Delay(50)).GetAwaiter().GetResult();
            TestAssert.True(!ReferenceEquals(winner, conversion), "Cancellation cannot end an outstanding source borrow.");
            TestAssert.Equal(held, pool.Reserved);
        }
        finally { program.Resume.TrySetResult(); }
        TestAssert.Throws<OperationCanceledException>(() => conversion.GetAwaiter().GetResult());
        TestAssert.Equal(0L, pool.Reserved);
        TestAssert.Equal(0L, output.Length);
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        TestAssert.Equal(0, pool.CallerDisposals);
    }

    private sealed class PendingModelFontProgram : IFontProgramSource
    {
        public string StableId => "pending-model-source";
        public readonly TaskCompletionSource Entered = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public readonly TaskCompletionSource Resume = new(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly byte[] bytes = TestFontBuilder.CreateTestFont();
        public async ValueTask<IOoxPdfFontProgramLease> AcquireAsync(IMemoryBudget? operationBudget, CancellationToken ct = default)
        {
            Entered.TrySetResult();
            await Resume.Task.ConfigureAwait(false);
            return TestFontPrograms.Borrow(bytes);
        }
    }

    private static Dictionary<string, byte[]> ModelAdmissionParts(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read);
        return zip.Entries.ToDictionary(entry => entry.FullName, entry => {
            using var source = entry.Open(); using var copy = new MemoryStream(); source.CopyTo(copy); return copy.ToArray();
        });
    }
}
