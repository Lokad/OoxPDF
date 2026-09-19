using System.Text;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static class PublicApiTests
{
    public static void PublicApiPreservesOptionalRecordArguments()
    {
        // These public call forms compiled against NuGet 0.1.4. Keep named and
        // omitted arguments working when internal signatures are refactored.
        var request = new FontRequest(FamilyName: "Arial", Italic: true);
        TestAssert.Equal(false, request.Bold);
        TestAssert.Equal(true, request.Italic);
        TestAssert.Equal(new FontStyleKey(false, false, 400, 0, false), new FontStyleKey());
        TestAssert.Equal(new FontStyleKey(false, false, 400, 0, true), new FontStyleKey(HasMathTable: true));

        var diagnostic = new OoxPdfDiagnostic("TEST", OoxPdfSeverity.Warning, "Message", Feature: "font");
        TestAssert.Equal<string?>(null, diagnostic.PartName);
        TestAssert.Equal<int?>(null, diagnostic.SlideIndex);
        TestAssert.Equal<int?>(null, diagnostic.PageIndex);
        TestAssert.Equal("font", diagnostic.Feature);
        TestAssert.Equal<string?>(null, diagnostic.Fallback);
    }

    public static void PublicFontSourcesPreserveOptionalCancellationTokens()
    {
        byte[] bytes = [1, 2, 3];
        var memory = new MemoryFontProgramSource("compatibility", bytes);
        IFontProgramSource source = memory;
        TestAssert.True(memory.GetBytesAsync().GetAwaiter().GetResult().Span.SequenceEqual(bytes), "Concrete memory source must accept an omitted token.");
        TestAssert.True(source.GetBytesAsync().GetAwaiter().GetResult().Span.SequenceEqual(bytes), "Interface source must accept an omitted token.");

        string path = Path.GetTempFileName();
        try
        {
            File.WriteAllBytes(path, bytes);
            var file = new FileFontProgramSource(path);
            TestAssert.True(file.GetBytesAsync().GetAwaiter().GetResult().Span.SequenceEqual(bytes), "File source must accept an omitted token.");
        }
        finally
        {
            File.Delete(path);
        }

        using var client = new HttpClient();
        // Invalid input exercises the public call without contacting a server.
        var exception = TestAssert.Throws<ArgumentException>(() =>
            OoxPdfFontPackResolver.CreateHttpAsync("", new Uri("https://example.invalid/"), client).GetAwaiter().GetResult());
        TestAssert.Equal("packId", exception.ParamName);
    }

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

    public static void OptionsDefaultToFinalDocxMarkupMode()
    {
        var options = new OoxPdfOptions();

        TestAssert.Equal(OoxPdfDocxMarkupMode.Final, options.DocxMarkupMode);
        TestAssert.Equal(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, options.DocxMarkupGeometryMode);
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

    public static void ConvertStreamCompletesForDocx()
    {
        byte[] bytes = ReadMinimalDocxBytes("<w:p/>");
        using var input = new MemoryStream(bytes, writable: false);
        using var output = new MemoryStream();

        OoxPdfConverter.Convert(
            input,
            output,
            new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx });

        TestAssert.True(input.CanRead, "Stream conversion should leave the input stream open.");
        TestAssert.True(output.CanWrite, "Stream conversion should leave the output stream open.");
        AssertPdfHeader(output);
    }

    public static void ConvertAsyncStreamCompletesForDocx()
    {
        byte[] bytes = ReadMinimalDocxBytes("<w:p/>");
        using var input = new MemoryStream(bytes, writable: false);
        using var output = new MemoryStream();

        OoxPdfConverter.ConvertAsync(
            input,
            output,
            new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx })
            .GetAwaiter()
            .GetResult();

        AssertPdfHeader(output);
    }

    public static void ConvertStreamRequiresExplicitInputKind()
    {
        byte[] bytes = ReadMinimalDocxBytes("<w:p/>");
        using var input = new MemoryStream(bytes, writable: false);
        using var output = new MemoryStream();

        NotSupportedException ex = TestAssert.Throws<NotSupportedException>(
            () => OoxPdfConverter.Convert(input, output, new OoxPdfOptions()));

        TestAssert.Contains("InputKind", ex.Message);
        TestAssert.Equal(0L, output.Length);
    }

    public static void ConvertStreamWithCancelledTokenThrowsBeforeInputKindValidation()
    {
        using var input = new MemoryStream([]);
        using var output = new MemoryStream();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        TestAssert.Throws<OperationCanceledException>(
            () => OoxPdfConverter.Convert(input, output, new OoxPdfOptions(), cancellation.Token));

        TestAssert.Equal(0L, output.Length);
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

    public static void ConvertStreamWorksWithForwardOnlyOutputForDocx()
    {
        byte[] bytes = ReadMinimalDocxBytes("<w:p><w:r><w:t>forward only</w:t></w:r></w:p>");
        using var input = new MemoryStream(bytes, writable: false);
        using var inner = new MemoryStream();
        using var output = new ForwardOnlyWriteStream(inner);

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx });

        TestAssert.True(input.CanRead, "Stream conversion should leave the input stream open.");
        TestAssert.True(output.CanWrite, "Stream conversion should leave the output stream open.");
        ValidatePdfXref(inner.ToArray());
    }

    public static void ConvertAsyncStreamWorksWithForwardOnlyOutputForDocx()
    {
        byte[] bytes = ReadMinimalDocxBytes("<w:p><w:r><w:t>forward only async</w:t></w:r></w:p>");
        using var input = new MemoryStream(bytes, writable: false);
        using var inner = new MemoryStream();
        using var output = new ForwardOnlyWriteStream(inner);

        OoxPdfConverter.ConvertAsync(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx }).GetAwaiter().GetResult();

        ValidatePdfXref(inner.ToArray());
    }

    public static void ConvertStreamWorksWithForwardOnlyOutputForPptx()
    {
        byte[] bytes = ReadMinimalPptxBytes();
        using var input = new MemoryStream(bytes, writable: false);
        using var inner = new MemoryStream();
        using var output = new ForwardOnlyWriteStream(inner);

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx });

        ValidatePdfXref(inner.ToArray());
    }

    public static void ConvertAsyncStreamWorksWithForwardOnlyOutputForPptx()
    {
        byte[] bytes = ReadMinimalPptxBytes();
        using var input = new MemoryStream(bytes, writable: false);
        using var inner = new MemoryStream();
        using var output = new ForwardOnlyWriteStream(inner);

        OoxPdfConverter.ConvertAsync(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx }).GetAwaiter().GetResult();

        ValidatePdfXref(inner.ToArray());
    }

    public static void ConvertStreamWorksWithForwardOnlyInputForDocx()
    {
        byte[] bytes = ReadMinimalDocxBytes("<w:p><w:r><w:t>forward only input</w:t></w:r></w:p>");
        using var inner = new MemoryStream(bytes, writable: false);
        using var input = new ForwardOnlyReadStream(inner);
        using var output = new MemoryStream();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx });

        TestAssert.True(input.CanRead, "Stream conversion should leave the input stream open.");
        TestAssert.True(output.CanWrite, "Stream conversion should leave the output stream open.");
        AssertPdfHeader(output);
    }

    public static void ConvertAsyncStreamWorksWithForwardOnlyInputForDocx()
    {
        byte[] bytes = ReadMinimalDocxBytes("<w:p><w:r><w:t>forward only input async</w:t></w:r></w:p>");
        using var inner = new MemoryStream(bytes, writable: false);
        using var input = new ForwardOnlyReadStream(inner);
        using var output = new MemoryStream();

        OoxPdfConverter.ConvertAsync(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx }).GetAwaiter().GetResult();

        AssertPdfHeader(output);
    }

    public static void ConvertStreamWorksWithForwardOnlyInputForPptx()
    {
        byte[] bytes = ReadMinimalPptxBytes();
        using var inner = new MemoryStream(bytes, writable: false);
        using var input = new ForwardOnlyReadStream(inner);
        using var output = new MemoryStream();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx });

        ValidatePdfXref(output.ToArray());
    }

    public static void ConvertAsyncStreamWorksWithForwardOnlyInputForPptx()
    {
        byte[] bytes = ReadMinimalPptxBytes();
        using var inner = new MemoryStream(bytes, writable: false);
        using var input = new ForwardOnlyReadStream(inner);
        using var output = new MemoryStream();

        OoxPdfConverter.ConvertAsync(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx }).GetAwaiter().GetResult();

        ValidatePdfXref(output.ToArray());
    }

    public static void ConvertStreamRejectsIdenticalInputOutput()
    {
        byte[] bytes = ReadMinimalDocxBytes("<w:p/>");
        using var both = new MemoryStream(bytes);

        TestAssert.Throws<ArgumentException>(() => OoxPdfConverter.Convert(both, both, new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx }));
    }

    public static void ConvertStreamRejectsNonEmptyOutput()
    {
        byte[] bytes = ReadMinimalDocxBytes("<w:p/>");
        using var input = new MemoryStream(bytes, writable: false);
        using var output = new MemoryStream(new byte[] { 1, 2, 3 });

        TestAssert.Throws<ArgumentException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx }));
    }

    public static void ConvertStreamRejectsNonZeroOutputPosition()
    {
        byte[] bytes = ReadMinimalDocxBytes("<w:p/>");
        using var input = new MemoryStream(bytes, writable: false);
        using var output = new MemoryStream();
        output.WriteByte(0);
        output.Position = 1;

        TestAssert.Throws<ArgumentException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx }));
    }

    public static void ConvertRejectsSameInputOutputPath()
    {
        string input = WriteMinimalDocx("<w:p/>");

        TestAssert.Throws<ArgumentException>(() => OoxPdfConverter.Convert(input, input));
    }

    private static byte[] ReadMinimalPptxBytes()
    {
        string path = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """,
        });
        return File.ReadAllBytes(path);
    }

    private static void ValidatePdfXref(byte[] pdf)
    {
        TestAssert.True(pdf.Length > 20, "Forward-only conversion should write PDF bytes.");
        string text = System.Text.Encoding.ASCII.GetString(pdf);
        TestAssert.Contains("%PDF-1.7", text);
        TestAssert.Contains("xref", text);
        TestAssert.Contains("trailer", text);
        TestAssert.Contains("%%EOF", text);

        int xref = text.IndexOf("\nxref\n", StringComparison.Ordinal);
        TestAssert.True(xref >= 0, "Expected xref table.");
        string[] lines = text.Substring(xref).Split('\n');
        TestAssert.True(lines.Length >= 3, "Expected xref entries.");
        string[] header = lines[2].Trim().Split(' ');
        TestAssert.Equal(2, header.Length);
        int count = int.Parse(header[1], System.Globalization.CultureInfo.InvariantCulture);
        TestAssert.True(count >= 2, "Expected at least catalog and pages objects.");
        for (int i = 0; i < count; i++)
        {
            string entry = lines[3 + i].Trim();
            if (i == 0)
            {
                TestAssert.Equal("0000000000 65535 f", entry);
                continue;
            }

            string offsetText = entry.Split(' ')[0];
            int offset = int.Parse(offsetText, System.Globalization.CultureInfo.InvariantCulture);
            TestAssert.True(offset >= 0 && offset < pdf.Length, "Xref offset must point inside the PDF, got: " + offset);
            string atOffset = text.Substring(offset, Math.Min(20, text.Length - offset));
            TestAssert.True(atOffset.Contains(" 0 obj", StringComparison.Ordinal), "Xref offset must point at an object header, got: " + atOffset);
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

        protected override void Dispose(bool disposing)
        {
        }
    }
    private sealed class ForwardOnlyReadStream(MemoryStream inner) : Stream
    {
        public override bool CanRead => inner.CanRead;

        public override bool CanSeek => false;

        public override bool CanWrite => false;

        public override long Length => throw new NotSupportedException();

        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }

        public override void Flush() => inner.Flush();

        public override int Read(byte[] buffer, int offset, int count) => inner.Read(buffer, offset, count);

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing)
        {
        }
    }
    public static void ConvertRejectsUndefinedOptionEnums()
    {
        byte[] bytes = ReadMinimalDocxBytes("<w:p/>");
        using var input = new MemoryStream(bytes, writable: false);
        using var output = new MemoryStream();

        TestAssert.Throws<ArgumentOutOfRangeException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = (OoxPdfInputKind)99 }));
        TestAssert.Equal(0L, output.Length);
    }

    public static void ConvertRejectsUndefinedMarkupModes()
    {
        string input = WriteMinimalDocx("<w:p/>");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        TestAssert.Throws<ArgumentOutOfRangeException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DocxMarkupMode = (OoxPdfDocxMarkupMode)99 }));
        TestAssert.Throws<ArgumentOutOfRangeException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DocxMarkupGeometryMode = (OoxPdfDocxMarkupGeometryMode)99 }));
    }

    public static void ConvertHonorsFixedCreationDate()
    {
        string input = WriteMinimalDocx("<w:p><w:r><w:t>dated</w:t></w:r></w:p>");
        var date = new DateTimeOffset(2026, 9, 9, 12, 30, 0, TimeSpan.Zero);
        string datedOutput = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        string undatedOutput = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, datedOutput, new OoxPdfOptions { FixedCreationDate = date });
        OoxPdfConverter.Convert(input, undatedOutput);

        string dated = File.ReadAllText(datedOutput, Encoding.ASCII);
        string undated = File.ReadAllText(undatedOutput, Encoding.ASCII);
        TestAssert.Contains("/CreationDate (D:20260909123000+00", dated);
        TestAssert.DoesNotContain("/Info", undated);
    }

    public static void ConvertPreservesExistingDestinationWhenWriterRejectsPage()
    {
        string input = WriteZeroWidthDocx();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        File.WriteAllText(output, "SENTINEL");

        TestAssert.Throws<ArgumentOutOfRangeException>(() => OoxPdfConverter.Convert(input, output));
        TestAssert.Equal("SENTINEL", File.ReadAllText(output));
        TestAssert.Equal(0, Directory.GetFiles(Path.GetDirectoryName(output)!, Path.GetFileName(output) + ".tmp-*").Length);
    }

    public static void ConvertReplacesExistingDestinationOnSuccess()
    {
        string input = WriteMinimalDocx("<w:p><w:r><w:t>replaced</w:t></w:r></w:p>");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        File.WriteAllText(output, "SENTINEL");

        OoxPdfConverter.Convert(input, output);

        byte[] bytes = File.ReadAllBytes(output);
        TestAssert.True(bytes.Length > 5 && bytes[0] == 37, "Completed conversion must publish a real PDF over the destination.");
        TestAssert.Equal(0, Directory.GetFiles(Path.GetDirectoryName(output)!, Path.GetFileName(output) + ".tmp-*").Length);
    }

    public static void ConvertPreservesExistingDestinationOnUnreadableInput()
    {
        string input = Path.ChangeExtension(Path.GetTempFileName(), ".docx");
        File.WriteAllBytes(input, new byte[] { 0x50, 0x4B, 0x03, 0x04, 0x00 });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        File.WriteAllText(output, "SENTINEL");

        TestAssert.Throws<InvalidDataException>(() => OoxPdfConverter.Convert(input, output));
        TestAssert.Equal("SENTINEL", File.ReadAllText(output));
    }

    private static string WriteZeroWidthDocx()
    {
        string input = WriteMinimalDocx("<w:p><w:r><w:t>zero</w:t></w:r></w:p>");
        string xml;
        using (var archive = new System.IO.Compression.ZipArchive(File.Open(input, FileMode.Open, FileAccess.Read), System.IO.Compression.ZipArchiveMode.Read))
        using (var reader = new StreamReader(TestAssert.NotNull(archive.GetEntry("word/document.xml")).Open()))
        {
            xml = reader.ReadToEnd().Replace("12240", "00000");
        }

        using (var archive = new System.IO.Compression.ZipArchive(File.Open(input, FileMode.Open, FileAccess.ReadWrite), System.IO.Compression.ZipArchiveMode.Update))
        using (var writer = new StreamWriter(TestAssert.NotNull(archive.GetEntry("word/document.xml")).Open()))
        {
            writer.Write(xml);
        }

        return input;
    }
    private static byte[] ReadMinimalDocxBytes(string bodyContent)
    {
        return File.ReadAllBytes(WriteMinimalDocx(bodyContent));
    }

    private static void AssertPdfHeader(MemoryStream output)
    {
        byte[] bytes = output.ToArray();
        TestAssert.True(bytes.Length > 5, "Stream conversion should write PDF bytes.");
        TestAssert.Equal((byte)'%', bytes[0]);
        TestAssert.Equal((byte)'P', bytes[1]);
        TestAssert.Equal((byte)'D', bytes[2]);
        TestAssert.Equal((byte)'F', bytes[3]);
        TestAssert.Equal((byte)'-', bytes[4]);
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
                new FontStyleKey(request.Bold, request.Italic, 400, 0, false),
                source,
                IsFallback: false);
        }
    }

    private sealed class ObservingFontProgramSource : IFontProgramSource
    {
        public string StableId => "test:observed-font";

        public bool WasCalled { get; private set; }

        public bool ObservedCanBeCanceled { get; private set; }

        public ValueTask<ReadOnlyMemory<byte>> GetBytesAsync(CancellationToken ct)
        {
            WasCalled = true;
            ObservedCanBeCanceled = ct.CanBeCanceled;
            return ValueTask.FromResult(ReadOnlyMemory<byte>.Empty);
        }
    }
}
