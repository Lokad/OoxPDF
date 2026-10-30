using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    public static void MemoryBudgetFontProgramsFundPrivateCopiesAndSelectedCollectionFaces()
    {
        byte[] collection = TestFontBuilder.CreateCollection("TestFontA", "TestFontB");
        var padded = new byte[4 * 1024 * 1024];
        collection.CopyTo(padded, 0);
        foreach ((byte[] bytes, int face) in new[] { (TestFontBuilder.CreateTestFont(), 0), (padded, 1) })
        {
            var resolution = ProgramResolution(new MemoryFontProgramSource("caller-owned-font", bytes), face);
            OpenTypeFont baseline = TestAssert.NotNull(FontProgramLoader.Load(resolution, default));
            var pool = new TestMemoryBudget(2L * baseline.Bytes.Length);
            using (var conversion = OoxConversionBudget.BeginScope(null, pool))
            {
                OpenTypeFont first = TestAssert.NotNull(FontProgramLoader.Load(resolution, default));
                TestAssert.Equal((long)first.Bytes.Length, pool.Reserved);
                TestAssert.True(first.Bytes.Span.SequenceEqual(baseline.Bytes.Span), "Selected SFNT bytes must remain exact.");
                TestAssert.True(FontProgramLoader.Load(ProgramResolution(new MemoryFontProgramSource("malformed", new byte[3])), default) is null,
                    "Ordinary malformed-font recovery is retained.");
                TestAssert.Equal((long)first.Bytes.Length, pool.Reserved);
                OpenTypeFont second = TestAssert.NotNull(FontProgramLoader.Load(resolution, default));
                TestAssert.Equal(2L * first.Bytes.Length, pool.Reserved);
                TestAssert.True(second.Bytes.Span.SequenceEqual(baseline.Bytes.Span), "Independent copies preserve source data.");
                var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => FontProgramLoader.Load(resolution, default));
                TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
                TestAssert.Equal(2L * first.Bytes.Length, pool.Reserved);
            }
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
            TestAssert.Equal(0, pool.CallerDisposals);
        }
    }

    public static void MemoryBudgetFontProgramRefusalIsTerminalAcrossPublicConverters()
    {
        string directory = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ooxpdf-font-admission-" + Guid.NewGuid().ToString("N"))).FullName;
        try
        {
            foreach ((string relative, OoxPdfInputKind kind) in new[] {
                ("Composition/report.docx", OoxPdfInputKind.Docx),
                ("pptx-ladder-02-plain-text.pptx", OoxPdfInputKind.Pptx) })
            foreach (bool file in new[] { false, true })
            {
                bool acquired = false;
                var source = new ObservedFontProgram(() => acquired = true);
                var pool = new TestMemoryBudget { RefuseChange = (_, after) => acquired && after > 0 };
                var diagnostics = new List<OoxPdfDiagnostic>();
                var options = new OoxPdfOptions { InputKind = kind, FontResolver = new ProgramResolver(source),
                    MemoryBudget = pool, DiagnosticSink = diagnostics.Add };
                OoxPdfLimitExceededException failure;
                if (file)
                {
                    string output = Path.Combine(directory, "result.pdf");
                    File.WriteAllBytes(output, [7, 8, 9]);
                    failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(Case(relative), output, options));
                    TestAssert.True(File.ReadAllBytes(output).SequenceEqual(new byte[] { 7, 8, 9 }), "Late font refusal preserves an existing destination.");
                    TestAssert.Equal(0, Directory.GetFiles(directory, "*.tmp-*").Length);
                }
                else
                {
                    using var input = File.OpenRead(Case(relative));
                    using var output = new MemoryStream();
                    failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output, options));
                    TestAssert.Equal(0L, output.Length);
                    TestAssert.True(input.CanRead && output.CanWrite, "Caller streams stay open.");
                }
                TestAssert.True(acquired, "Refusal occurs at font preparation, after package intake and source acquisition.");
                TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
                TestAssert.Equal(OoxPdfLimitScope.Conversion, failure.Scope);
                TestAssert.True(diagnostics.All(d => d.Id is not "FONT_NO_USABLE_FACE" and not "PPTX_NODE_RENDER_FAILED"),
                    "Admission refusal must not become a font or node fallback.");
                TestAssert.Equal(0L, pool.Reserved);
                TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
                TestAssert.Equal(0, pool.CallerDisposals);
            }
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    public static void MemoryBudgetFontSubsetFundsConstructionAndRetainsOnlyTheResult()
    {
        // This parsed source is caller-owned before the conversion scope. Only
        // its newly constructed subset belongs to that scope.
        OpenTypeFont font = OpenTypeFont.Load(TestFontBuilder.CreateTestFont());
        int[] codePoints = [32, 65, 66, 0x301, 0x1F600];
        byte[] original = font.Bytes.ToArray();
        PdfEmbeddedFont baseline = PdfEmbeddedFont.Create(font, codePoints, default);
        var pool = new TestMemoryBudget();
        long constructionPeak;
        using (var conversion = OoxConversionBudget.BeginScope(null, pool))
        {
            PdfEmbeddedFont subset = PdfEmbeddedFont.Create(font, codePoints, default);
            TestAssert.True(subset.UsesSubsetFontProgram, "Exercise actual subset production.");
            TestAssert.True(subset.FontProgramBytes.Span.SequenceEqual(baseline.FontProgramBytes.Span), "Admission must preserve every subset byte.");
            TestAssert.Equal((long)subset.FontProgramBytes.Length, pool.Reserved);
            TestAssert.True(pool.Peak > pool.Reserved, "Construction funds temporary table buffers alongside the final SFNT.");
            constructionPeak = pool.Peak;
            long retained = pool.Reserved;
            using var output = new ObservedOutput(() => TestAssert.True(pool.Reserved >= retained,
                "Owned subset bytes remain funded through actual resource writes."));
            var page = new PdfPage(100, 100, "BT /F1 12 Tf <0001> Tj ET\n", [new PdfFontResource("F1", subset)]);
            PdfDocumentWriter.WriteStagedAsync(output, AsyncEnumerable.ToAsyncEnumerable<PdfPage>([page]), new OoxConversionLimits { MaxResidentPageContentBytesPerConversion = 0 }, null, default, temporaryStorage: new OoxPdfFileTemporaryStorage(Path.GetTempPath())).GetAwaiter().GetResult();
            TestAssert.True(output.Writes > 0, "Exercise real PDF resource emission.");
            TestAssert.Equal(retained, pool.Reserved);
            PdfEmbeddedFont whole = PdfEmbeddedFont.Create(font, [], default);
            TestAssert.True(!whole.UsesSubsetFontProgram, "The established empty-subset fallback remains.");
            TestAssert.Equal(retained, pool.Reserved);
        }
        TestAssert.Equal(0L, pool.Reserved);
        pool.Capacity = constructionPeak - 1;
        using (var conversion = OoxConversionBudget.BeginScope(null, pool))
        {
            var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => PdfEmbeddedFont.Create(font, codePoints, default));
            TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
            TestAssert.Equal(0L, pool.Reserved);
        }
        TestAssert.True(font.Bytes.Span.SequenceEqual(original), "Direct remapping must never mutate shared source glyphs.");
        TestAssert.Equal(0, pool.CallerDisposals);
    }

    public static void MemoryBudgetFontSubsetCancellationRefundsConstruction()
    {
        OpenTypeFont font = OpenTypeFont.Load(TestFontBuilder.CreateTestFont());
        using var cancellation = new CancellationTokenSource();
        var pool = new TestMemoryBudget { RefuseChange = (_, after) => { if (after > 0) cancellation.Cancel(); return false; } };
        using (var conversion = OoxConversionBudget.BeginScope(null, pool))
        {
            var failure = TestAssert.Throws<OperationCanceledException>(() => PdfEmbeddedFont.Create(font, [65, 66], cancellation.Token));
            TestAssert.Equal(cancellation.Token, failure.CancellationToken);
            TestAssert.True(pool.Peak > 0, "Cancel after construction admission has succeeded.");
            TestAssert.Equal(0L, pool.Reserved);
        }
        TestAssert.True(PdfEmbeddedFont.Create(font, [65, 66], default).UsesSubsetFontProgram, "Cancellation leaves the caller's immutable source usable.");
    }

    public static void MemoryBudgetFontSubsetPreservesPortableCompoundOutlines()
    {
        byte[] bytes = File.ReadAllBytes(Case("PortableFonts/DejaVuSans.ttf"));
        OpenTypeFont font = OpenTypeFont.Load(bytes);
        ushort originalGlyph = font.MapCodePoint('é');
        TestAssert.True(font.TryReadGlyphOutline(originalGlyph, out var expected) && expected.IsCompound,
            "The pinned portable face supplies a real compound glyph.");
        var pool = new TestMemoryBudget();
        using (var conversion = OoxConversionBudget.BeginScope(null, pool))
        {
            PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, [65, 'é', 0x1F600], default);
            OpenTypeFont subset = OpenTypeFont.Load(embedded.FontProgramBytes.ToArray());
            ushort glyph = subset.MapCodePoint('é');
            TestAssert.True(subset.TryReadGlyphOutline(glyph, out var actual), "Remapped component references remain readable.");
            TestAssert.Equal(expected.Bounds, actual.Bounds);
            TestAssert.Equal(expected.Contours.Count, actual.Contours.Count);
            for (int i = 0; i < expected.Contours.Count; i++)
                TestAssert.True(expected.Contours[i].Points.SequenceEqual(actual.Contours[i].Points), "Compound contours preserve all source points.");
            TestAssert.Equal(font.GetAdvanceWidth(originalGlyph), subset.GetAdvanceWidth(glyph));
            TestAssert.Equal((long)embedded.FontProgramBytes.Length, pool.Reserved);
        }
        TestAssert.Equal(0L, pool.Reserved);
        TestAssert.True(bytes.AsSpan().SequenceEqual(font.Bytes.Span), "Owned subset writes leave the original font unchanged.");
    }

    private static FontFaceResolution ProgramResolution(IFontProgramSource source, int face = 0) =>
        new("TestFont", "TestFont", new FontStyleKey(false, false, 400, face, false), source, IsFallback: false);

    private sealed class ProgramResolver(IFontProgramSource source) : IFontResolver
    {
        public FontFaceResolution Resolve(FontRequest request) => ProgramResolution(source);
    }

    private sealed class ObservedFontProgram(Action acquired) : IFontProgramSource
    {
        private readonly byte[] bytes = TestFontBuilder.CreateTestFont();
        public string StableId => "observed-font-admission";
        public ValueTask<IOoxPdfFontProgramLease> AcquireAsync(IMemoryBudget? operationBudget, CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            acquired();
            return ValueTask.FromResult(TestFontPrograms.Borrow(bytes));
        }
    }
}
