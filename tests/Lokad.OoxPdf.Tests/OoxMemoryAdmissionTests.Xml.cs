using System.IO.Compression;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    public static void MemoryBudgetXmlCacheOwnsStorageThroughPackageLifetime()
    {
        using var input = XmlAdmissionPackage();
        var pool = new TestMemoryBudget();
        OoxPackage package;
        long funded;
        using (var conversion = OoxConversionBudget.BeginScope(null, pool))
        {
            package = OoxPackage.Open(input, default);
            long intake = pool.Reserved;
            var part = TestAssert.NotNull(package.GetPart("/data.xml"));
            var document = package.LoadXml(part, default);
            funded = pool.Reserved;
            TestAssert.True(funded > intake, "Parsed XML storage is admitted in addition to its encoded source.");
            TestAssert.True(pool.Peak > funded, "Parser construction headroom refunds while the cached DOM stays owned.");
            int requests = pool.Requests;
            TestAssert.True(ReferenceEquals(document, package.LoadXml(part, default)), "Reuse the admitted DOM.");
            TestAssert.Equal(requests, pool.Requests);
            TestAssert.Equal(funded, pool.Reserved);
            TestAssert.Equal(1024, document.Root!.Elements().Count());
        }
        try
        {
            TestAssert.True(funded == pool.Reserved, "An independently owned package scope survives its invoking scope.");
            TestAssert.Equal(1024, package.LoadXml(TestAssert.NotNull(package.GetPart("/data.xml")), default).Root!.Elements().Count());
        }
        finally { package.Dispose(); package.Dispose(); }
        TestAssert.Equal(0L, pool.Reserved);
        TestAssert.True(input.CanRead, "The caller owns the input.");
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        TestAssert.Equal(0, pool.CallerDisposals);
    }

    public static void MemoryBudgetXmlCacheRefusalCancellationAndMalformedInputRefundGrowth()
    {
        foreach (string mode in new[] { "refuse", "cancel", "malformed" })
        {
            using var input = XmlAdmissionPackage();
            using var cancellation = new CancellationTokenSource();
            var pool = new TestMemoryBudget();
            using (var conversion = OoxConversionBudget.BeginScope(null, pool))
            using (var package = OoxPackage.Open(input, default))
            {
                var stable = package.LoadXml(TestAssert.NotNull(package.GetPart("/stable.xml")), default);
                long prior = pool.Reserved;
                bool acquired = false;
                pool.RefuseChange = (before, after) =>
                {
                    if (after <= before) return false;
                    acquired = true;
                    if (mode == "cancel") cancellation.Cancel();
                    return mode == "refuse";
                };
                var part = TestAssert.NotNull(package.GetPart(mode == "malformed" ? "/bad.xml" : "/data.xml"));
                if (mode == "refuse")
                {
                    var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => package.LoadXml(part, default));
                    TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
                    TestAssert.Equal(OoxPdfLimitScope.Conversion, failure.Scope);
                }
                else if (mode == "cancel")
                {
                    var failure = TestAssert.Throws<OperationCanceledException>(() => package.LoadXml(part, cancellation.Token));
                    TestAssert.Equal(cancellation.Token, failure.CancellationToken);
                }
                else TestAssert.Throws<InvalidDataException>(() => package.LoadXml(part, default));
                TestAssert.True(acquired, "Admission occurs before XML parsing.");
                TestAssert.True(prior == pool.Reserved, "Failed parsing refunds only attempted XML growth.");
                pool.RefuseChange = null;
                int requests = pool.Requests;
                TestAssert.True(ReferenceEquals(stable, package.LoadXml(TestAssert.NotNull(package.GetPart("/stable.xml")), default)), "A failed attempt preserves prior funded cache entries.");
                TestAssert.Equal(requests, pool.Requests);
                if (mode != "malformed")
                {
                    TestAssert.Equal(1024, package.LoadXml(part, default).Root!.Elements().Count());
                    TestAssert.True(pool.Reserved > prior, "Refusal/cancellation did not publish an unfunded cache entry.");
                }
            }
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
            TestAssert.Equal(0, pool.CallerDisposals);
        }
    }

    public static void MemoryBudgetContentTypeParsingAdmitsBeforeReadingXml()
    {
        string types = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            string.Concat(Enumerable.Range(0, 512).Select(i => "<Default Extension=\"e" + i + "\" ContentType=\"application/xml\"/>")) + "</Types>";
        foreach (bool cancel in new[] { false, true })
        {
            using var input = TestFixtures.CreateZipPackage(new Dictionary<string, string> { ["[Content_Types].xml"] = types });
            using var cancellation = new CancellationTokenSource();
            bool acquired = false;
            var pool = new TestMemoryBudget { RefuseChange = (_, after) =>
            {
                if (after < 512 * 1024) return false;
                acquired = true;
                if (cancel) cancellation.Cancel();
                return !cancel;
            }};
            using (var conversion = OoxConversionBudget.BeginScope(null, pool))
            {
                if (cancel)
                    TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(() => OoxPackage.Open(input, cancellation.Token)).CancellationToken);
                else
                    TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPackage.Open(input, default)).Kind);
                TestAssert.True(acquired, "Significant parsed metadata is funded separately from bounded encoded input.");
                TestAssert.Equal(0L, pool.Reserved);
            }
            TestAssert.True(input.CanRead, "Metadata failure leaves the caller input open.");
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
            TestAssert.Equal(0, pool.CallerDisposals);
        }
    }

    public static void MemoryBudgetXmlCacheRefusalPreservesPublicPublication()
    {
        string directory = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ooxpdf-xml-admission-" + Guid.NewGuid().ToString("N"))).FullName;
        try
        {
            foreach (bool docx in new[] { true, false })
            {
                byte[] bytes = XmlAdmissionPublicInput(docx);
                string sourcePath = Path.Combine(directory, docx ? "source.docx" : "source.pptx");
                File.WriteAllBytes(sourcePath, bytes);
                // Validate the supported source independently of the refusal path.
                using (var source = new MemoryStream(bytes))
                using (var output = new MemoryStream())
                {
                    OoxPdfConverter.Convert(source, output, new OoxPdfOptions {
                        InputKind = docx ? OoxPdfInputKind.Docx : OoxPdfInputKind.Pptx,
                        FontResolver = new ProgramResolver(new ObservedFontProgram(() => { })) });
                    TestAssert.True(output.Length > 0, "The composed source converts without a refusing budget.");
                }
                foreach (bool file in new[] { false, true })
                {
                    bool fontUsed = false, refused = false;
                    var pool = new TestMemoryBudget { RefuseChange = (_, after) =>
                    {
                        if (fontUsed || after < 512 * 1024) return false;
                        refused = true;
                        return true;
                    }};
                    var diagnostics = new List<OoxPdfDiagnostic>();
                    var options = new OoxPdfOptions { InputKind = docx ? OoxPdfInputKind.Docx : OoxPdfInputKind.Pptx,
                        FontResolver = new ProgramResolver(new ObservedFontProgram(() => fontUsed = true)),
                        MemoryBudget = pool, DiagnosticSink = diagnostics.Add };
                    string target = Path.Combine(directory, "result.pdf");
                    File.WriteAllBytes(target, [7, 8, 9]);
                    using var source = new MemoryStream(bytes);
                    using var output = new MemoryStream();
                    Task conversion = file ? OoxPdfConverter.ConvertAsync(sourcePath, target, options) : OoxPdfConverter.ConvertAsync(source, output, options);
                    var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => conversion.GetAwaiter().GetResult());
                    TestAssert.True(refused && !fontUsed, "XML storage admission precedes font preparation.");
                    TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
                    TestAssert.Equal(OoxPdfLimitScope.Conversion, failure.Scope);
                    TestAssert.Equal(0L, output.Length);
                    TestAssert.True(File.ReadAllBytes(target).SequenceEqual(new byte[] { 7, 8, 9 }), "Refusal preserves the previous destination.");
                    TestAssert.True(source.CanRead && output.CanWrite, "Caller streams remain open after refusal.");
                    TestAssert.True(diagnostics.All(d => d.Id is not "PPTX_NODE_RENDER_FAILED" and not "CONVERSION_RESOURCE_SUMMARY"), "Refusal escapes recovery and does not report success.");
                    TestAssert.Equal(0, Directory.GetFiles(directory, "*.tmp-*").Length);
                    TestAssert.Equal(0L, pool.Reserved);
                    TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
                    TestAssert.Equal(0, pool.CallerDisposals);
                }
            }
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    private static MemoryStream XmlAdmissionPackage() => TestFixtures.CreateZipPackage(new Dictionary<string, string> {
        ["[Content_Types].xml"] = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"xml\" ContentType=\"application/xml\"/></Types>",
        ["stable.xml"] = "<r><n/></r>",
        ["data.xml"] = "<r>" + string.Concat(Enumerable.Repeat("<n a=\"1\"/>", 1024)) + "</r>",
        ["bad.xml"] = "<r>"
    });

    private static byte[] XmlAdmissionPublicInput(bool docx)
    {
        using var archive = ZipFile.OpenRead(Case(docx ? "docx-ladder-00-blank.docx" : "pptx-ladder-02-plain-text.pptx"));
        var entries = new Dictionary<string, byte[]>();
        foreach (var entry in archive.Entries)
        {
            using var source = entry.Open();
            using var copy = new MemoryStream();
            source.CopyTo(copy);
            entries[entry.FullName] = copy.ToArray();
        }
        string part = docx ? "word/document.xml" : "ppt/slides/slide1.xml";
        using var input = new MemoryStream(entries[part]);
        var document = XDocument.Load(input);
        XNamespace ns = docx ? "http://schemas.openxmlformats.org/wordprocessingml/2006/main" : "http://schemas.openxmlformats.org/drawingml/2006/main";
        var parent = docx ? document.Root!.Element(ns + "body")! : document.Descendants(ns + "p").First().Parent!;
        var paragraphs = Enumerable.Range(0, 1024).Select(_ => new XElement(ns + "p", new XElement(ns + "r", new XElement(ns + "t", "A"))));
        if (docx) parent.AddFirst(paragraphs); else parent.Add(paragraphs);
        using var xml = new MemoryStream();
        document.Save(xml, SaveOptions.DisableFormatting);
        entries[part] = xml.ToArray();
        using var result = TestFixtures.CreateZipPackage(entries);
        return result.ToArray();
    }
}
