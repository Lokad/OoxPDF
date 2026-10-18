using System.Collections.Concurrent;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static class DocxFontCacheTests
{
    public static void DocxConversionLoadsEachFontProgramOnce()
    {
        byte[] fontBytes = TestFontBuilder.CreateTestFont();
        var source = new CountingProgramSource("counting-test", fontBytes);
        var resolver = new SingleSourceResolver(source);
        byte[] pdfBytes = ConvertMinimalDoc(resolver);

        TestAssert.True(pdfBytes.Length > 0, "Conversion should produce output.");
        TestAssert.Equal(1, source.RequestCount);
    }

    public static void DocxCoveredRunsSkipFallbackProgramLoads()
    {
        byte[] fontBytes = TestFontBuilder.CreateTestFont();
        var counters = new ConcurrentDictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var resolver = new PerFamilyCountingResolver(fontBytes, counters);
        byte[] pdfBytes = ConvertMinimalDoc(resolver);

        TestAssert.True(pdfBytes.Length > 0, "Conversion should produce output.");
        TestAssert.Equal(1, counters.Count);
        TestAssert.Equal("counting:TestFont", counters.Keys.Single());
        TestAssert.Equal(1, counters.Values.Single());
    }

    public static void DocxNumberingRetriesKeepParsedFontsOwnedByOneConversion()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = new CountingProgramSource("numbering-retry-test", TestFontBuilder.CreateTestFont());
            var resolver = new SingleSourceResolver(source);
            byte[] first = ConvertNumberedDoc(resolver, mode), second = ConvertNumberedDoc(resolver, mode);
            TestAssert.Equal(2, source.RequestCount);
            TestAssert.True(first.AsSpan().SequenceEqual(second), "Independent conversions retain deterministic note output without sharing parsed fonts.");
        }
    }

    public static void DocxConcurrentNumberingRetriesKeepIndependentParsedFontOwners()
    {
        var source = new CountingProgramSource("parallel-numbering-retry-test", TestFontBuilder.CreateTestFont());
        var resolver = new SingleSourceResolver(source);
        var tasks = Enumerable.Range(0, 4).Select(_ => Task.Run(() => ConvertNumberedDoc(resolver, OoxPdfDocxMarkupMode.Final))).ToArray();
        Task.WaitAll(tasks);
        TestAssert.Equal(tasks.Length, source.RequestCount);
        TestAssert.True(tasks.All(task => task.Result.AsSpan().SequenceEqual(tasks[0].Result)),
            "Concurrent conversions may share immutable source bytes, with their own parsed font/preparation ownership.");
    }

    private static byte[] ConvertNumberedDoc(IFontResolver resolver, OoxPdfDocxMarkupMode mode)
    {
        string input = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes",
            "spanning-page-notes-calibri-12-zero.docx"));
        using var source = File.OpenRead(input);
        using var output = new MemoryStream();
        OoxPdfConverter.Convert(source, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx, FontResolver = resolver, DocxMarkupMode = mode,
            DocxMarkupGeometryMode = mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
        });
        return output.ToArray();
    }

    private static byte[] ConvertMinimalDoc(IFontResolver resolver)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:rPr><w:rFonts w:ascii="TestFont" w:hAnsi="TestFont"/></w:rPr><w:t>Probe</w:t></w:r>
                      <w:r><w:rPr><w:b/><w:rFonts w:ascii="TestFont" w:hAnsi="TestFont"/></w:rPr><w:t>Bold</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using var inputStream = File.OpenRead(input);
        using var outputStream = new MemoryStream();
        OoxPdfConverter.Convert(inputStream, outputStream, new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx, FontResolver = resolver });
        return outputStream.ToArray();
    }

    private sealed class CountingProgramSource(string stableId, byte[] bytes) : IFontProgramSource
    {
        private int requestCount;

        public string StableId { get; } = stableId;

        public int RequestCount => requestCount;

        public ValueTask<ReadOnlyMemory<byte>> GetBytesAsync(CancellationToken ct)
        {
            System.Threading.Interlocked.Increment(ref requestCount);
            ct.ThrowIfCancellationRequested();
            return ValueTask.FromResult((ReadOnlyMemory<byte>)bytes);
        }
    }

    private sealed class SingleSourceResolver(CountingProgramSource source) : IFontResolver
    {
        public FontFaceResolution Resolve(FontRequest request)
        {
            return new FontFaceResolution(
                request.FamilyName,
                "TestFont",
                new FontStyleKey(request.Bold, request.Italic, request.Bold ? 700 : 400, 0, false),
                source,
                IsFallback: false);
        }
    }

    private sealed class PerFamilyCountingResolver(byte[] fontBytes, ConcurrentDictionary<string, int> counters) : IFontResolver
    {
        public FontFaceResolution Resolve(FontRequest request)
        {
            string family = string.IsNullOrWhiteSpace(request.FamilyName) ? "(default)" : request.FamilyName;
            var source = new CountingProgramSource("counting:" + family, fontBytes);
            return new FontFaceResolution(
                request.FamilyName,
                family,
                new FontStyleKey(request.Bold, request.Italic, request.Bold ? 700 : 400, 0, false),
                new FamilySource(counters, source),
                IsFallback: false);
        }

        private sealed class FamilySource(ConcurrentDictionary<string, int> counters, CountingProgramSource inner) : IFontProgramSource
        {
            public string StableId => inner.StableId;

            public ValueTask<ReadOnlyMemory<byte>> GetBytesAsync(CancellationToken ct)
            {
                counters.AddOrUpdate(inner.StableId, 1, (_, count) => count + 1);
                return inner.GetBytesAsync(ct);
            }
        }
    }
}
