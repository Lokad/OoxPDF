using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    public static void LayoutAdmissionRepeatedHeadersReuseCapacityAcrossPasses()
    {
        IDocxTextMeasurer measurer = LayoutAdmissionMeasurer(); // Caller-owned before the operation.
        var pool = new TestMemoryBudget();
        using var input = new MemoryStream(LayoutAdmissionInput(200));
        using (var operation = OoxConversionBudget.BeginScope(null, pool))
        {
            var package = OoxPackage.Open(input, default);
            var document = new DocxReader().Read(package, null, default, OoxPdfDocxMarkupMode.Final);
            long sourceModels = pool.Reserved;
            var first = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, measurer, default);
            TestAssert.Equal(200, first.Pages.Count);
            TestAssert.Equal(4000, first.Pages.Sum(page => page.StaticTextLines.Count));
            TestAssert.True(pool.Reserved > sourceModels + 3_000_000, "Repeated static frames need significant capacity beyond their single source model.");
            long held = pool.Reserved;
            pool.Capacity = held;
            // Keep the first candidate readable while a new engine repeats the
            // same placement. A private per-pass allowance would refuse here.
            var second = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, measurer, default);
            TestAssert.Equal(first.Pages.Count, second.Pages.Count);
            TestAssert.Equal(first.Pages[^1].StaticTextLines[0].BaselineY, second.Pages[^1].StaticTextLines[0].BaselineY);
            TestAssert.Equal(held, pool.Reserved);
            package.Dispose();
            TestAssert.True(pool.Reserved > 0, "Retained layouts outlive their package buffers.");
            TestAssert.Equal("Header line", first.Pages[^1].StaticTextLines[0].Text);
            TestAssert.Equal("Header line", second.Pages[^1].StaticTextLines[0].Text);
        }
        TestAssert.Equal(0L, pool.Reserved);
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        TestAssert.Equal(0, pool.CallerDisposals);
    }

    public static void LayoutAdmissionRefusalAndCancellationPrecedeFrameConstruction()
    {
        foreach (bool cancel in new[] { false, true })
        foreach (int refusedGrowth in new[] { 1, 3 })
        {
            IDocxTextMeasurer measurer = LayoutAdmissionMeasurer();
            var pool = new TestMemoryBudget();
            using var cancellation = new CancellationTokenSource();
            using var input = new MemoryStream(LayoutAdmissionInput(20));
            using (var operation = OoxConversionBudget.BeginScope(null, pool))
            using (var package = OoxPackage.Open(input, default))
            {
                var document = new DocxReader().Read(package, null, default, OoxPdfDocxMarkupMode.Final);
                int growths = 0;
                pool.RefuseChange = (before, after) => {
                    if (after <= before || ++growths != refusedGrowth) return false;
                    if (cancel) cancellation.Cancel();
                    return !cancel;
                };
                var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
                if (cancel)
                    TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(() =>
                        engine.Create(document, measurer, cancellation.Token)).CancellationToken);
                else
                    TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, TestAssert.Throws<OoxPdfLimitExceededException>(() =>
                        engine.Create(document, measurer, default)).Kind);
                TestAssert.Equal(refusedGrowth, growths);
                pool.RefuseChange = null;
                TestAssert.Equal(20, engine.Create(document, measurer, default).Pages.Count);
            }
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        }
    }

    public static void LayoutAdmissionPublicRefusalPreservesDestinations()
    {
        byte[] bytes = LayoutAdmissionInput(10);
        // Identify the real first placement request through the live engine,
        // independently of source/font/staging requests and scope ordinals.
        long placementRequest = LayoutAdmissionFirstRequest(bytes);
        string directory = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ooxpdf-layout-refusal-" + Guid.NewGuid().ToString("N"))).FullName;
        try
        {
            string path = Path.Combine(directory, "source.docx"), target = Path.Combine(directory, "result.pdf");
            File.WriteAllBytes(path, bytes);
            foreach (bool file in new[] { false, true })
            {
                bool reached = false, acquired = false;
                var pool = new TestMemoryBudget { RefuseChange = (before, after) => {
                    if (before != 0 || after != placementRequest) return false;
                    reached = true;
                    return true;
                }};
                var options = new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx, MemoryBudget = pool,
                    FontResolver = new ProgramResolver(new ObservedFontProgram(() => acquired = true)) };
                File.WriteAllBytes(target, [1, 2, 3]);
                using var input = new MemoryStream(bytes);
                using var output = new MemoryStream();
                Task conversion = file ? OoxPdfConverter.ConvertAsync(path, target, options) : OoxPdfConverter.ConvertAsync(input, output, options);
                TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, TestAssert.Throws<OoxPdfLimitExceededException>(() => conversion.GetAwaiter().GetResult()).Kind);
                TestAssert.True(reached && acquired, "Refuse actual layout admission after font preparation.");
                TestAssert.True(File.ReadAllBytes(target).SequenceEqual(new byte[] { 1, 2, 3 }), "Layout refusal preserves the destination.");
                TestAssert.Equal(0L, output.Length);
                TestAssert.True(input.CanRead && output.CanWrite, "Caller streams remain open.");
                TestAssert.Equal(0, Directory.GetFiles(directory, "*.tmp-*").Length);
                TestAssert.Equal(0L, pool.Reserved);
                TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
            }
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    public static void LayoutAdmissionRepeatedHeadersPreservePdfAndRefundBeforeSummary()
    {
        byte[] bytes = LayoutAdmissionInput(10, countField: true);
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            byte[] expected;
            using (var input = new MemoryStream(bytes))
            using (var output = new MemoryStream())
            {
                OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx,
                    DocxMarkupMode = mode, FixedCreationDate = DateTimeOffset.UnixEpoch,
                    FontResolver = new ProgramResolver(new ObservedFontProgram(() => { })) });
                expected = output.ToArray();
            }
            var pool = new TestMemoryBudget();
            bool summarized = false;
            using var source = new MemoryStream(bytes);
            using var destination = new MemoryStream();
            OoxPdfConverter.Convert(source, destination, new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx,
                DocxMarkupMode = mode, MemoryBudget = pool, FixedCreationDate = DateTimeOffset.UnixEpoch, ReportResourceUsage = true,
                FontResolver = new ProgramResolver(new ObservedFontProgram(() => { })),
                DiagnosticSink = diagnostic => {
                    if (diagnostic.Id != "CONVERSION_RESOURCE_SUMMARY") return;
                    summarized = true;
                    TestAssert.True(pool.Reserved > 0 && pool.Reserved < pool.Peak / 2, "Repeated-layout capacity refunds before serialization; source/PDF storage stays funded.");
                    TestAssert.Equal(0L, destination.Length);
                }});
            TestAssert.True(summarized, "Observe the live staging/serialization boundary.");
            TestAssert.True(expected.AsSpan().SequenceEqual(destination.ToArray()), "Admission preserves exact output through count-dependent retries.");
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        }
    }

    private static long LayoutAdmissionFirstRequest(byte[] bytes)
    {
        IDocxTextMeasurer measurer = LayoutAdmissionMeasurer();
        var pool = new TestMemoryBudget();
        long request = 0;
        using (var operation = OoxConversionBudget.BeginScope(null, pool))
        using (var input = new MemoryStream(bytes))
        using (var package = OoxPackage.Open(input, default))
        {
            var document = new DocxReader().Read(package, null, default, OoxPdfDocxMarkupMode.Final);
            pool.RefuseChange = (before, after) => { if (request == 0 && after > before) request = after; return false; };
            _ = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, measurer, default);
        }
        TestAssert.True(request > 0, "Placement must ask for capacity beyond admitted source models.");
        TestAssert.Equal(0L, pool.Reserved);
        return request;
    }

    private static IDocxTextMeasurer LayoutAdmissionMeasurer() => new DocxEmbeddedTextMeasurer(PdfEmbeddedFont.Create(
        OpenTypeFont.Load(TestFontBuilder.CreateTestFont(), 0, default), Enumerable.Range(32, 95), default));

    private static byte[] LayoutAdmissionInput(int pages, bool countField = false)
    {
        const string w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        string body = string.Concat(Enumerable.Range(0, pages).Select(index =>
            "<w:p><w:r><w:t>Page A</w:t></w:r>" + (countField && index == 0 ? "<w:fldSimple w:instr=\"NUMPAGES\"><w:r><w:t>1</w:t></w:r></w:fldSimple>" : "") +
            (index + 1 < pages ? "<w:r><w:br w:type=\"page\"/></w:r>" : "") + "</w:p>"));
        using var input = TestFixtures.CreateZipPackage(new Dictionary<string, string> {
            ["[Content_Types].xml"] = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>",
            ["word/document.xml"] = $"<w:document xmlns:w=\"{w}\" xmlns:r=\"{r}\"><w:body>{body}<w:sectPr><w:headerReference w:type=\"default\" r:id=\"h1\"/><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"7200\" w:bottom=\"1440\" w:left=\"1440\" w:right=\"1440\" w:header=\"720\"/></w:sectPr></w:body></w:document>",
            ["word/_rels/document.xml.rels"] = $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"h1\" Type=\"{r}/header\" Target=\"header1.xml\"/></Relationships>",
            ["word/header1.xml"] = $"<w:hdr xmlns:w=\"{w}\">" + string.Concat(Enumerable.Repeat("<w:p><w:r><w:t>Header line</w:t></w:r></w:p>", 20)) + "</w:hdr>"
        });
        return input.ToArray();
    }
}
