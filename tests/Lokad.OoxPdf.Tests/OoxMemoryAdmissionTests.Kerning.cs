using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    public static void MemoryBudgetKerningMapRefusalRefundsThePrivateProgram()
    {
        byte[] bytes = TestFontBuilder.CreateTestFont();
        bool reachedMap = false;
        var pool = new TestMemoryBudget { RefuseChange = (before, after) =>
        {
            if (before != bytes.Length || after <= before) return false;
            reachedMap = true;
            return true;
        }};
        var face = ProgramResolution(new MemoryFontProgramSource("kerning-admission", bytes));
        using (var operation = OoxConversionBudget.BeginScope(null, pool))
        {
            var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => FontProgramLoader.Load(face, default));
            TestAssert.True(reachedMap, "Refuse parsed storage after the private SFNT copy is admitted.");
            TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
            TestAssert.Equal(OoxPdfLimitScope.Conversion, failure.Scope);
            TestAssert.Equal(0L, pool.Reserved);
        }
        OpenTypeFont font = TestAssert.NotNull(FontProgramLoader.Load(face, default));
        TestAssert.Equal((short)-50, font.GetKerning(34, 35));
        TestAssert.True(font.Bytes.Span.SequenceEqual(bytes), "Refusal leaves the caller's source usable and unchanged.");
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        TestAssert.Equal(0, pool.CallerDisposals);
    }

    public static void MemoryBudgetKerningSortRefusalAndCancellationRefundConstruction()
    {
        byte[] bytes = TestFontBuilder.CreateTestFont();
        var face = ProgramResolution(new MemoryFontProgramSource("kerning-sort", bytes));
        foreach (bool cancel in new[] { false, true })
        {
            bool mapAdmitted = false, reachedSort = false;
            using var cancellation = new CancellationTokenSource();
            var pool = new TestMemoryBudget { RefuseChange = (before, after) =>
            {
                if (before == bytes.Length && after > before) mapAdmitted = true;
                if (!mapAdmitted || before != 0 || after == 0) return false;
                reachedSort = true;
                if (cancel) cancellation.Cancel();
                return !cancel;
            }};
            using (var operation = OoxConversionBudget.BeginScope(null, pool))
            {
                if (cancel)
                {
                    var failure = TestAssert.Throws<OperationCanceledException>(() => FontProgramLoader.Load(face, cancellation.Token));
                    TestAssert.Equal(cancellation.Token, failure.CancellationToken);
                }
                else
                {
                    var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => FontProgramLoader.Load(face, cancellation.Token));
                    TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
                }
                TestAssert.True(reachedSort && pool.Peak > bytes.Length, "Sort scratch is admitted while the program and parsed arrays remain owned.");
                TestAssert.Equal(0L, pool.Reserved);
            }
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
            TestAssert.Equal(0, pool.CallerDisposals);
        }
    }

    public static void MemoryBudgetKerningMapRefusalPreservesPublicPublication()
    {
        string directory = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ooxpdf-kerning-admission-" + Guid.NewGuid().ToString("N"))).FullName;
        int programLength = TestFontBuilder.CreateTestFont().Length;
        try
        {
            foreach ((string relative, OoxPdfInputKind kind) in new[] {
                ("Composition/report.docx", OoxPdfInputKind.Docx),
                ("pptx-ladder-02-plain-text.pptx", OoxPdfInputKind.Pptx) })
            foreach (bool file in new[] { false, true })
            {
                bool acquired = false, reachedMap = false;
                var pool = new TestMemoryBudget { RefuseChange = (before, after) =>
                {
                    if (!acquired || before != programLength || after <= before) return false;
                    reachedMap = true;
                    return true;
                }};
                var diagnostics = new List<OoxPdfDiagnostic>();
                var options = new OoxPdfOptions { InputKind = kind, MemoryBudget = pool,
                    FontResolver = new ProgramResolver(new ObservedFontProgram(() => acquired = true)), DiagnosticSink = diagnostics.Add };
                string outputPath = Path.Combine(directory, "result.pdf");
                File.WriteAllBytes(outputPath, [7, 8, 9]);
                using var input = File.OpenRead(Case(relative));
                using var output = new MemoryStream();
                Task conversion = file ? OoxPdfConverter.ConvertAsync(Case(relative), outputPath, options)
                    : OoxPdfConverter.ConvertAsync(input, output, options);
                var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => conversion.GetAwaiter().GetResult());
                TestAssert.True(reachedMap, "The public route reaches retained kerning storage after program acquisition.");
                TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
                TestAssert.Equal(OoxPdfLimitScope.Conversion, failure.Scope);
                TestAssert.Equal(0L, output.Length);
                TestAssert.True(File.ReadAllBytes(outputPath).SequenceEqual(new byte[] { 7, 8, 9 }), "Refusal preserves the previous file.");
                TestAssert.Equal(0, Directory.GetFiles(directory, "*.tmp-*").Length);
                TestAssert.True(input.CanRead && output.CanWrite, "Caller streams remain open.");
                TestAssert.True(diagnostics.All(d => d.Id is not "FONT_NO_USABLE_FACE" and not "PPTX_NODE_RENDER_FAILED"), "Refusal is terminal rather than font or node recovery.");
                TestAssert.Equal(0L, pool.Reserved);
                TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
                TestAssert.Equal(0, pool.CallerDisposals);
            }
        }
        finally { Directory.Delete(directory, recursive: true); }
    }
}
