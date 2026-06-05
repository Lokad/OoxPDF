using Lokad.OoxPdf;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static class PublicApiTests
{
    public static void PublicApiRejectsMissingInput()
    {
        string missingInput = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
        string output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pdf");

        FileNotFoundException ex = TestAssert.Throws<FileNotFoundException>(
            () => OoxPdfConverter.Convert(missingInput, output));

        TestAssert.Equal(missingInput, ex.FileName);
    }

    public static void AutoDetectsPptxExtension()
    {
        OoxPdfInputKind kind = OoxPdfConverter.DetectInputKind("deck.PPTX");

        TestAssert.Equal(OoxPdfInputKind.Pptx, kind);
    }

    public static void AutoDetectsDocxExtension()
    {
        OoxPdfInputKind kind = OoxPdfConverter.DetectInputKind("document.docx");

        TestAssert.Equal(OoxPdfInputKind.Docx, kind);
    }

    public static void DeterministicConversionProducesStableBytes()
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
                  <w:body><w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body>
                </w:document>
                """
        });
        string output1 = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        string output2 = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output1, new OoxPdfOptions { Deterministic = true });
        OoxPdfConverter.Convert(input, output2, new OoxPdfOptions { Deterministic = true });

        byte[] first = File.ReadAllBytes(output1);
        byte[] second = File.ReadAllBytes(output2);
        TestAssert.True(first.SequenceEqual(second), "Deterministic conversion should produce stable PDF bytes.");
    }

    public static void ConverterUsesCustomFontResolverForDocxText()
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
                    <w:p><w:r><w:t>resolver probe</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var resolver = new CountingFontResolver();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { FontResolver = resolver });

        TestAssert.True(resolver.ResolveCalls > 0, "DOCX conversion should use the supplied font resolver for text embedding.");
    }

    public static void ConvertWithCancelledTokenThrowsBeforeInputValidation()
    {
        string missingInput = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        string output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pdf");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        TestAssert.Throws<OperationCanceledException>(
            () => OoxPdfConverter.Convert(missingInput, output, cancellation.Token));

        TestAssert.True(!File.Exists(output), "Cancelled conversion should not create an output file.");
    }

    public static void ConvertAsyncWithCancelledTokenThrowsBeforeInputValidation()
    {
        string missingInput = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
        string output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pdf");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        TestAssert.Throws<OperationCanceledException>(
            () => OoxPdfConverter.ConvertAsync(missingInput, output, cancellation.Token).GetAwaiter().GetResult());

        TestAssert.True(!File.Exists(output), "Cancelled async conversion should not create an output file.");
    }

    public static void ConvertAsyncCompletesForDocx()
    {
        string input = WriteMinimalDocx("<w:p/>");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.ConvertAsync(input, output).GetAwaiter().GetResult();

        TestAssert.True(File.Exists(output), "Async conversion should create a PDF.");
        TestAssert.True(new FileInfo(output).Length > 0, "Async conversion should write PDF bytes.");
    }

    public static void ConverterPassesCancellationTokenToFontProgramSource()
    {
        string input = WriteMinimalDocx("<w:p><w:r><w:t>token probe</w:t></w:r></w:p>");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        using var cancellation = new CancellationTokenSource();
        var source = new ObservingFontProgramSource();

        OoxPdfConverter.Convert(
            input,
            output,
            new OoxPdfOptions { FontResolver = new ObservingFontResolver(source) },
            cancellation.Token);

        TestAssert.True(source.WasCalled, "DOCX conversion should load the custom font source.");
        TestAssert.True(source.ObservedCanBeCanceled, "DOCX conversion should pass the caller token to font loading.");
    }

    private static string WriteMinimalDocx(string bodyContent)
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
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
                """ + bodyContent + """
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
    }

    private sealed class ObservingFontResolver(IFontProgramSource source) : IFontResolver
    {
        public FontFaceResolution Resolve(FontRequest request)
        {
            return new FontFaceResolution(
                request.FamilyName,
                request.FamilyName,
                new FontStyleKey(request.Bold, request.Italic),
                source,
                IsFallback: false);
        }
    }

    private sealed class ObservingFontProgramSource : IFontProgramSource
    {
        public string StableId => "test:observed-font";

        public bool WasCalled { get; private set; }

        public bool ObservedCanBeCanceled { get; private set; }

        public ValueTask<ReadOnlyMemory<byte>> GetBytesAsync(CancellationToken ct = default)
        {
            WasCalled = true;
            ObservedCanBeCanceled = ct.CanBeCanceled;
            return ValueTask.FromResult(ReadOnlyMemory<byte>.Empty);
        }
    }
}
