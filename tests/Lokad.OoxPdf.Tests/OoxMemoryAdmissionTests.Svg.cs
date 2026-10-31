using System.Text;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    private const string SvgMemorySource = "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 100 100\"><defs><linearGradient id=\"g\"><stop offset=\"0\" stop-color=\"red\" stop-opacity=\"0\"/><stop offset=\"1\" stop-color=\"blue\"/></linearGradient></defs><audit-probe/><g opacity=\"0.5\"><rect width=\"80\" height=\"60\" fill=\"url(#g)\"/></g></svg>";

    public static void SvgMemoryScratchLivesThroughPaintAndReusesCapacity()
    {
        byte[] bytes = Encoding.UTF8.GetBytes(SvgMemorySource);
        var baseline = new PdfGraphicsBuilder();
        TestAssert.True(PptxRenderer.TryRenderSvgImage(baseline, 0, 0, 100, 100, bytes, _ => { }, 0, null, default), "Render the ungoverned control.");
        string expected = baseline.ToString();
        var pool = new TestMemoryBudget();
        using (var operation = OoxConversionBudget.BeginScope(null, pool))
        {
            for (int picture = 0; picture < 20; picture++)
            {
                var graphics = new PdfGraphicsBuilder();
                bool observed = false;
                TestAssert.True(PptxRenderer.TryRenderSvgImage(graphics, 0, 0, 100, 100, bytes, _ => {
                    observed = true;
                    TestAssert.True(pool.Reserved > 0, "Keep the DOM/gradient working grant while its diagnostic reader is active.");
                }, 0, null, default), "Render while the working grant is owned.");
                TestAssert.True(observed, "The existing unsupported-element diagnostic observes the live parse.");
                TestAssert.Equal(0L, pool.Reserved); // This is render-local scratch, not retained PDF funding.
                TestAssert.Equal(0, graphics.StateDepth);
                TestAssert.Equal(expected, graphics.ToString());
                TestAssert.Equal(baseline.Groups.Count, graphics.Groups.Count);
                TestAssert.True(graphics.Groups.Count > 0, "Retained transparency resources remain readable after scratch refunds.");
                TestAssert.True(graphics.Groups[0].Group.Content.Length > 0, "Group content owns its retained snapshot.");
                pool.Capacity = pool.Peak; // Repeated pictures cannot accumulate work-only charges.
            }
        }
        TestAssert.Equal(0L, pool.Reserved);
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        TestAssert.Equal(0, pool.CallerDisposals);
    }

    public static void SvgMemoryRefusalAndCancellationPrecedeParsing()
    {
        foreach (bool cancel in new[] { false, true })
        {
            using var cancellation = new CancellationTokenSource();
            bool requested = false, diagnosed = false;
            var pool = new TestMemoryBudget { RefuseChange = (_, after) => {
                if (after == 0) return false;
                requested = true;
                if (cancel) cancellation.Cancel();
                return !cancel;
            }};
            var graphics = new PdfGraphicsBuilder();
            using (var operation = OoxConversionBudget.BeginScope(null, pool))
            {
                if (cancel)
                    TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(() =>
                        PptxRenderer.TryRenderSvgImage(graphics, 0, 0, 10, 10, "<svg"u8.ToArray(), _ => diagnosed = true, 0, null, cancellation.Token)).CancellationToken);
                else
                    TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, TestAssert.Throws<OoxPdfLimitExceededException>(() =>
                        PptxRenderer.TryRenderSvgImage(graphics, 0, 0, 10, 10, "<svg"u8.ToArray(), _ => diagnosed = true, 0, null, default)).Kind);
                TestAssert.True(requested && !diagnosed, "Refusal/cancellation escapes before invalid SVG can become ignored content.");
                TestAssert.Equal(0L, pool.Reserved);
                TestAssert.Equal(0, graphics.StateDepth);
            }
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        }
    }

    public static void SvgMemoryParseAndObserverFailuresRefundScratch()
    {
        var pool = new TestMemoryBudget();
        using (var operation = OoxConversionBudget.BeginScope(null, pool))
        {
            bool observed = false;
            TestAssert.True(!PptxRenderer.TryRenderSvgImage(new PdfGraphicsBuilder(), 0, 0, 10, 10, "<svg"u8.ToArray(), _ => {
                observed = true;
                TestAssert.True(pool.Reserved > 0, "Malformed-content reporting still reads within scratch ownership.");
            }, 0, null, default), "Malformed SVG keeps its existing recovery result.");
            TestAssert.True(observed, "Malformed content keeps its existing diagnostic.");
            TestAssert.Equal(0L, pool.Reserved);
            var sentinel = new InvalidOperationException("SVG observer sentinel");
            TestAssert.Equal(sentinel, TestAssert.Throws<InvalidOperationException>(() =>
                PptxRenderer.TryRenderSvgImage(new PdfGraphicsBuilder(), 0, 0, 10, 10, Encoding.UTF8.GetBytes(SvgMemorySource), _ => {
                    TestAssert.True(pool.Reserved > 0, "The observer borrows still-live SVG models.");
                    throw sentinel;
                }, 0, null, default)));
            TestAssert.Equal(0L, pool.Reserved);
        }
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
    }

    public static void SvgMemoryPublicRefusalPreservesFileAndStreamDestinations()
    {
        byte[] noise = new byte[128 * 1024]; new Random(4811).NextBytes(noise);
        byte[] malformed = Encoding.UTF8.GetBytes(Convert.ToBase64String(noise) + "<svg");
        long request = SvgMemoryRequest(malformed);
        string directory = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ooxpdf-svg-admission-" + Guid.NewGuid().ToString("N"))).FullName;
        try
        {
            foreach (var (name, kind) in new[] { ("docx-svg-inline.docx", OoxPdfInputKind.Docx), ("pptx-svg-extensions.pptx", OoxPdfInputKind.Pptx) })
            {
                var parts = ModelAdmissionParts(File.ReadAllBytes(Case(name)));
                string[] images = parts.Keys.Where(key => key.EndsWith(".svg", StringComparison.OrdinalIgnoreCase)).ToArray();
                TestAssert.True(images.Length > 0, "The public input contains authored SVG pictures.");
                foreach (string image in images) parts[image] = malformed;
                using var package = TestFixtures.CreateZipPackage(parts);
                string sourcePath = Path.Combine(directory, name), target = Path.Combine(directory, "result.pdf");
                byte[] bytes = package.ToArray(); File.WriteAllBytes(sourcePath, bytes);
                foreach (bool file in new[] { false, true })
                {
                    bool reached = false, svgDiagnostic = false;
                    var pool = new TestMemoryBudget { RefuseChange = (before, after) => {
                        if (before != 0 || after != request) return false;
                        reached = true; return true;
                    }};
                    var options = new OoxPdfOptions { InputKind = kind, MemoryBudget = pool,
                        FontResolver = new ProgramResolver(new ObservedFontProgram(() => { })),
                        DiagnosticSink = diagnostic => svgDiagnostic |= diagnostic.Id.StartsWith("SVG_", StringComparison.Ordinal) };
                    File.WriteAllBytes(target, [1, 2, 3]);
                    using var input = new MemoryStream(bytes); using var output = new MemoryStream();
                    Task conversion = file ? OoxPdfConverter.ConvertAsync(sourcePath, target, options) : OoxPdfConverter.ConvertAsync(input, output, options);
                    var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => conversion.GetAwaiter().GetResult());
                    TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
                    TestAssert.Equal(OoxPdfLimitScope.Conversion, failure.Scope);
                    TestAssert.True(reached && !svgDiagnostic, "SVG refusal cannot become DOCX raster fallback or PPTX ignored-node recovery.");
                    TestAssert.Equal(0L, output.Length);
                    TestAssert.True(File.ReadAllBytes(target).SequenceEqual(new byte[] { 1, 2, 3 }), "Preserve the destination.");
                    TestAssert.True(input.CanRead && output.CanWrite, "Caller streams stay open.");
                    TestAssert.Equal(0, Directory.GetFiles(directory, "*.tmp-*").Length);
                    TestAssert.Equal(0L, pool.Reserved);
                    TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
                }
            }
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    private static long SvgMemoryRequest(byte[] bytes)
    {
        var pool = new TestMemoryBudget();
        using (var operation = OoxConversionBudget.BeginScope(null, pool))
            PptxRenderer.TryRenderSvgImage(new PdfGraphicsBuilder(), 0, 0, 10, 10, bytes, _ => { }, 0, null, default);
        TestAssert.True(pool.Peak > 0, "The shared SVG painter must admit its own working capacity.");
        TestAssert.Equal(0L, pool.Reserved);
        return pool.Peak;
    }
}
