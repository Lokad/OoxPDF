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
