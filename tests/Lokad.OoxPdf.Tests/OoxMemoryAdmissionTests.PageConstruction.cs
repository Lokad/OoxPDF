using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    public static void PptxSparseChartConstructionAdmissionFollowsSlotExpansion()
    {
        var compact = ObserveSparseChartConstruction(2);
        var sparse = ObserveSparseChartConstruction(100_000);
        Console.WriteLine($"Chart slot grant/allocation: compact {compact.Request}/{compact.Allocated}, sparse {sparse.Request}/{sparse.Allocated} bytes.");
        TestAssert.True(sparse.Allocated >= 8L * 100_000, "Exercise significant dense construction from the sparse source.");
        TestAssert.True(sparse.Request >= 8L * 100_000 && sparse.Request > compact.Request,
            $"Page capacity must cover significant slot expansion: compact grant/allocation {compact.Request}/{compact.Allocated}, sparse {sparse.Request}/{sparse.Allocated} bytes.");
    }

    private static (long Request, long Allocated) ObserveSparseChartConstruction(int slots, int frames = 1)
    {
        using var input = SparseChartConstructionPackage(slots, frames);
        using var package = OoxPackage.Open(input, default);
        var document = new PptxReader().Read(package, default);
        var budget = new PageMemoryBudget();
        long request, allocated;
        using (var operation = OoxConversionBudget.BeginScope(null, budget))
        {
            var reader = new PptxRenderer(new TestFaceFontResolver()).RenderPagesAsync(document, package, null, default).GetAsyncEnumerator();
            try
            {
                long before = GC.GetTotalAllocatedBytes(precise: true);
                TestAssert.True(reader.MoveNextAsync().AsTask().GetAwaiter().GetResult(), "Render the source chart through the real page producer.");
                allocated = GC.GetTotalAllocatedBytes(precise: true) - before;
                request = budget.LargestLiveGrant().Bytes;
                TestAssert.True(reader.Current.Content.Length > 0, "Exercise the chart graphics path.");
            }
            finally { reader.DisposeAsync().AsTask().GetAwaiter().GetResult(); }
        }
        TestAssert.Equal(0L, budget.Pool.Reserved);
        TestAssert.Equal(budget.Pool.ScopesCreated, budget.Pool.ScopesDisposed);
        return (request, allocated);
    }

    private static MemoryStream SparseChartConstructionPackage(int slots, int frames = 1)
    {
        string chart = $$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea>
              <c:lineChart><c:grouping val="standard"/><c:ser><c:idx val="0"/><c:order val="0"/>
                <c:marker><c:symbol val="none"/></c:marker>
                <c:cat><c:strLit><c:ptCount val="{{slots}}"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="{{slots-1}}"><c:v>Z</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:ptCount val="{{slots}}"/><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="{{slots-1}}"><c:v>4</c:v></c:pt></c:numLit></c:val><c:smooth val="0"/>
              </c:ser><c:axId val="1"/><c:axId val="2"/></c:lineChart>
              <c:catAx><c:axId val="1"/><c:delete val="1"/><c:crossAx val="2"/></c:catAx>
              <c:valAx><c:axId val="2"/><c:scaling><c:min val="0"/><c:max val="5"/></c:scaling><c:delete val="1"/><c:crossAx val="1"/></c:valAx>
            </c:plotArea><c:dispBlanksAs val="span"/></c:chart></c:chartSpace>
            """;
        MemoryStream original = PptxTests.CreateSingleChartPackageStream(chart);
        if (frames == 1) return original;
        using (original)
        using (var archive = new ZipArchive(original, ZipArchiveMode.Read, leaveOpen: true))
        using (var source = archive.GetEntry("ppt/slides/slide1.xml")!.Open())
        {
            var slide = XDocument.Load(source);
            XNamespace presentation = "http://schemas.openxmlformats.org/presentationml/2006/main";
            var tree = slide.Descendants(presentation + "spTree").Single();
            var frame = tree.Elements(presentation + "graphicFrame").Single();
            for (int i = 1; i < frames; i++)
            {
                var copy = new XElement(frame);
                copy.Descendants(presentation + "cNvPr").FirstOrDefault()?.SetAttributeValue("id", 100 + i);
                tree.Add(copy);
            }
            return PptxTests.CreateSingleChartPackageStream(chart,
                new Dictionary<string, byte[]> { ["ppt/slides/slide1.xml"] = TestFixtures.Utf8(slide.ToString()) });
        }
    }

    public static void PptxSparseChartFramesShareOneConstructionWindow()
    {
        long one = ObserveSparseChartConstruction(100_000).Request;
        long two = ObserveSparseChartConstruction(100_000, frames: 2).Request;
        TestAssert.True(two >= one && two < one + one / 2,
            $"Sequential chart frames share construction headroom: {one} versus {two} bytes.");
    }

    public static void PptxSparseChartRefusalAndCancellationPreservePublication()
    {
        long request = ObserveSparseChartConstruction(100_000).Request;
        using var sourcePackage = SparseChartConstructionPackage(100_000);
        byte[] bytes = sourcePackage.ToArray();
        string inputPath = Path.Combine(Path.GetTempPath(), "sparse-source-" + Guid.NewGuid().ToString("N") + ".pptx");
        File.WriteAllBytes(inputPath, bytes);
        foreach (bool cancel in new[] { false, true })
        foreach (bool file in new[] { false, true })
        {
            using var cancellation = new CancellationTokenSource();
            var budget = new PageMemoryBudget { MatchBytes = request, Cancel = cancel ? cancellation : null };
            string target = Path.Combine(Path.GetTempPath(), "sparse-chart-" + Guid.NewGuid().ToString("N") + ".pdf");
            File.WriteAllBytes(target, [7, 8, 9]);
            using var input = new MemoryStream(bytes);
            using var output = new MemoryStream();
            var options = new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx, FontResolver = new TestFaceFontResolver(), MemoryBudget = budget };
            void Convert()
            {
                if (file) OoxPdfConverter.Convert(inputPath, target, options, cancellation.Token);
                else OoxPdfConverter.Convert(input, output, options, cancellation.Token);
            }
            if (cancel) TestAssert.Throws<OperationCanceledException>(Convert);
            else
            {
                OoxPdfLimitExceededException? refused = null;
                try { Convert(); } catch (OoxPdfLimitExceededException ex) { refused = ex; }
                TestAssert.True(refused is { Kind: OoxPdfLimitKind.MemoryAdmission, Scope: OoxPdfLimitScope.Conversion }, "Refusal stays structured and terminal.");
            }
            TestAssert.True(budget.Reached, "Select the observed chart construction request.");
            TestAssert.Equal(1, budget.MatchCount);
            TestAssert.Equal(0L, output.Length);
            TestAssert.True(File.ReadAllBytes(target).SequenceEqual(new byte[] { 7, 8, 9 }), "Preserve the existing destination.");
            TestAssert.Equal(0L, budget.Pool.Reserved);
            TestAssert.Equal(budget.Pool.ScopesCreated, budget.Pool.ScopesDisposed);
            TestAssert.Equal(0, budget.Pool.CallerDisposals);
        }
    }

    public static void PageMemoryConstructionFollowsPageAdvanceAndEarlyDisposal()
    {
        foreach (var kind in new[] { OoxPdfInputKind.Docx, OoxPdfInputKind.Pptx })
        foreach (bool early in new[] { false, true })
        {
            byte[] bytes = PageMemoryInput(kind, 3);
            using var source = new MemoryStream(bytes); using var package = OoxPackage.Open(source, default);
            var resolver = new ProgramResolver(new ObservedFontProgram(() => { }));
            var pages = PageMemoryPages(package, kind, resolver, 3, default);
            var budget = new PageMemoryBudget();
            using (var operation = OoxConversionBudget.BeginScope(null, budget))
            {
                var reader = pages.GetAsyncEnumerator();
                IMemoryReservation? previous = null;
                try
                {
                    int seen = 0;
                    while (reader.MoveNextAsync().AsTask().GetAwaiter().GetResult())
                    {
                        if (previous is not null) TestAssert.Equal(0L, previous.Bytes);
                        previous = budget.LargestLiveGrant();
                        TestAssert.True(previous.Bytes > 0, "Keep page construction funded while its yielded content is borrowed.");
                        if (kind == OoxPdfInputKind.Pptx)
                            TestAssert.True(previous.Bytes <= 256 * 1024, "A background-only slide must not fund unused master/layout definitions.");
                        TestAssert.True(reader.Current.Content.Length > 0, "Exercise the main graphics producer.");
                        if (++seen == 1) budget.Pool.Capacity = budget.Pool.Peak;
                        if (early) break;
                    }
                    TestAssert.Equal(early ? 1 : 3, seen);
                }
                finally { reader.DisposeAsync().AsTask().GetAwaiter().GetResult(); }
                TestAssert.True(previous is not null, "Observe the real page lease.");
                TestAssert.Equal(0L, previous!.Bytes);
            }
            TestAssert.Equal(0L, budget.Pool.Reserved);
            TestAssert.Equal(budget.Pool.ScopesCreated, budget.Pool.ScopesDisposed);
        }
    }

    public static void PageMemoryConstructionLivesAcrossPendingStagingIo()
    {
        foreach (var kind in new[] { OoxPdfInputKind.Docx, OoxPdfInputKind.Pptx })
        {
            byte[] bytes = File.ReadAllBytes(Case(kind == OoxPdfInputKind.Docx ? "docx-svg-inline.docx" : "pptx-svg-extensions.pptx"));
            using var source = new MemoryStream(bytes); using var package = OoxPackage.Open(source, default);
            var resolver = new ProgramResolver(new ObservedFontProgram(() => { }));
            var pages = PageMemoryPages(package, kind, resolver, 1, default);
            var budget = new PageMemoryBudget();
            using (var operation = OoxConversionBudget.BeginScope(null, budget))
            {
                var storage = new PageMemoryStorage();
                Task<PdfStagedDocument> production = PdfDocumentWriter.ProduceStagedPagesAsync(pages,
                    new OoxConversionLimits { MaxResidentPageContentBytesPerConversion = 0 }, null, default, storage).AsTask();
                IMemoryReservation? page = null;
                try
                {
                    TestAssert.True(storage.Started.Task.Wait(TimeSpan.FromSeconds(5)), "Observe pending staging I/O.");
                    TestAssert.True(!production.IsCompleted, "The staging reader is still borrowing page content.");
                    page = budget.LargestLiveGrant();
                    TestAssert.True(page.Bytes > 0, "Do not refund construction before the pending reader finishes.");
                }
                finally { storage.Proceed.TrySetResult(); }
                using var staged = production.GetAwaiter().GetResult();
                TestAssert.True(page is not null, "Observe a page construction grant.");
                TestAssert.Equal(0L, page!.Bytes);
                TestAssert.True(budget.Pool.Reserved > 0, "Independent staged storage remains funded after page construction ends.");
            }
            TestAssert.Equal(0L, budget.Pool.Reserved);
        }
    }

    public static void PageMemoryPublicRefusalAndCancellationPreserveDestinations()
    {
        string directory = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ooxpdf-page-construction-" + Guid.NewGuid().ToString("N"))).FullName;
        try
        {
            foreach (var kind in new[] { OoxPdfInputKind.Docx, OoxPdfInputKind.Pptx })
            {
                byte[] bytes = PageMemoryInput(kind, 1);
                var resolver = new ProgramResolver(new ObservedFontProgram(() => { }));
                long request;
                var control = new PageMemoryBudget();
                using (var source = new MemoryStream(bytes))
                using (var package = OoxPackage.Open(source, default))
                {
                    var pages = PageMemoryPages(package, kind, resolver, 1, default);
                    using var operation = OoxConversionBudget.BeginScope(null, control);
                    var reader = pages.GetAsyncEnumerator();
                    try
                    {
                        TestAssert.True(reader.MoveNextAsync().AsTask().GetAwaiter().GetResult(), "Prepare a real page grant for selective refusal.");
                        var page = control.LargestLiveGrant(); request = page.Bytes;
                        reader.DisposeAsync().AsTask().GetAwaiter().GetResult();
                        TestAssert.Equal(0L, page.Bytes);
                    }
                    finally { reader.DisposeAsync().AsTask().GetAwaiter().GetResult(); }
                }
                foreach (bool file in new[] { false, true })
                foreach (bool cancel in new[] { false, true })
                {
                    using var cancellation = new CancellationTokenSource();
                    var budget = new PageMemoryBudget { MatchBytes = request, Cancel = cancel ? cancellation : null };
                    string path = Path.Combine(directory, "input" + (kind == OoxPdfInputKind.Docx ? ".docx" : ".pptx"));
                    string target = Path.Combine(directory, "result.pdf");
                    File.WriteAllBytes(path, bytes); File.WriteAllBytes(target, [7, 8, 9]);
                    using var source = new MemoryStream(bytes); using var output = new MemoryStream();
                    var options = new OoxPdfOptions { InputKind = kind, FontResolver = resolver, MemoryBudget = budget };
                    void Convert() => (file ? OoxPdfConverter.ConvertAsync(path, target, options, cancellation.Token) :
                        OoxPdfConverter.ConvertAsync(source, output, options, cancellation.Token)).GetAwaiter().GetResult();
                    if (cancel) TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(Convert).CancellationToken);
                    else TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, TestAssert.Throws<OoxPdfLimitExceededException>(Convert).Kind);
                    TestAssert.True(budget.Reached, "Select the observed page request, not a package/model/font grant.");
                    TestAssert.Equal(0L, output.Length);
                    TestAssert.True(File.ReadAllBytes(target).SequenceEqual(new byte[] { 7, 8, 9 }), "Preserve the destination.");
                    TestAssert.True(source.CanRead && output.CanWrite, "Caller streams stay open.");
                    TestAssert.Equal(0, Directory.GetFiles(directory, "*.tmp-*").Length);
                    TestAssert.Equal(0L, budget.Pool.Reserved);
                    TestAssert.Equal(budget.Pool.ScopesCreated, budget.Pool.ScopesDisposed);
                }
            }
        }
        finally { Directory.Delete(directory, true); }
    }

    private static byte[] PageMemoryInput(OoxPdfInputKind kind, int pages)
    {
        if (kind == OoxPdfInputKind.Pptx) return File.ReadAllBytes(Case("pptx-ladder-01-solid-background.pptx"));
        var parts = ModelAdmissionParts(File.ReadAllBytes(Case("docx-ladder-00-blank.docx")));
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var xml = XDocument.Parse(Encoding.UTF8.GetString(parts["word/document.xml"]));
        var body = xml.Root!.Element(w + "body")!;
        var section = body.Element(w + "sectPr"); section?.Remove(); body.RemoveNodes();
        for (int page = 0; page < pages; page++)
        {
            if (page > 0) body.Add(new XElement(w + "p", new XElement(w + "r", new XElement(w + "br", new XAttribute(w + "type", "page")))));
            body.Add(new XElement(w + "p", new XElement(w + "r", new XElement(w + "t", "AAA"))));
        }
        if (section is not null) body.Add(section);
        parts["word/document.xml"] = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
        using var package = TestFixtures.CreateZipPackage(parts); return package.ToArray();
    }

    private static IAsyncEnumerable<PdfPage> PageMemoryPages(OoxPackage package, OoxPdfInputKind kind, IFontResolver resolver, int repeats, CancellationToken token)
    {
        // Models are caller-owned in these isolated producer tests; model scopes
        // still stay distinct when PPTX builds its scene during enumeration.
        if (kind == OoxPdfInputKind.Docx)
        {
            var document = new DocxReader().Read(package, null, token, OoxPdfDocxMarkupMode.Final);
            return new DocxRenderer(resolver, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderPagesAsync(document, null, token);
        }
        var deck = new PptxReader().Read(package, token, null);
        deck = deck with { Slides = Enumerable.Range(0, repeats).Select(index => new PptxSlide(deck.Slides[0].PartName, index)).ToArray() };
        return new PptxRenderer(resolver).RenderPagesAsync(deck, package, null, token);
    }

    // Caller instrumentation observes direct root reservations; nested model
    // scopes delegate unchanged. No production hooks or per-object ledger.
    private sealed class PageMemoryBudget : IMemoryBudget
    {
        public readonly TestMemoryBudget Pool = new();
        private readonly List<IMemoryReservation> direct = [];
        public long MatchBytes;
        public CancellationTokenSource? Cancel;
        public bool Reached;
        public int MatchCount;
        public IMemoryReservation LargestLiveGrant() => direct.Where(item => item.Bytes > 0).MaxBy(item => item.Bytes)
            ?? throw new InvalidOperationException("The producer must hold a page construction reservation.");
        public bool TryReserve(long bytes, out IMemoryReservation? reservation) => Pool.TryReserve(bytes, out reservation);
        public IMemoryScope CreateScope() => new Root(this, Pool.CreateScope());
        private sealed class Root(PageMemoryBudget owner, IMemoryScope inner) : IMemoryScope
        {
            public IMemoryScope CreateScope() => inner.CreateScope();
            public bool TryReserve(long bytes, out IMemoryReservation? reservation)
            {
                bool selected = bytes > 0 && bytes == owner.MatchBytes;
                if (selected)
                {
                    owner.Reached = true;
                    owner.MatchCount++;
                    if (owner.Cancel is null) { reservation = null; return false; }
                }
                bool accepted = inner.TryReserve(bytes, out reservation);
                if (accepted && bytes > 0 && reservation is not null) owner.direct.Add(reservation);
                if (selected) owner.Cancel!.Cancel();
                return accepted;
            }
            public void Dispose() => inner.Dispose();
        }
    }

    private sealed class PageMemoryStorage : IOoxPdfTemporaryStorage
    {
        public readonly TaskCompletionSource Started = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public readonly TaskCompletionSource Proceed = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public ValueTask<Stream> CreateAsync(IMemoryBudget? operationBudget, CancellationToken cancellationToken = default)
            => ValueTask.FromResult<Stream>(new GateStream(this));
        private sealed class GateStream(PageMemoryStorage owner) : MemoryStream
        {
            public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default)
            {
                owner.Started.TrySetResult();
                await owner.Proceed.Task.WaitAsync(cancellationToken);
                await base.WriteAsync(buffer, cancellationToken);
            }
        }
    }
}
