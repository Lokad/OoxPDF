using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    public static void ModelAdmissionMathLiteralCapacityStaysWithSourceOwner()
    {
        var pool = new TestMemoryBudget();
        using var input = MathLiteralAdmissionInput(8192);
        using (var operation = OoxConversionBudget.BeginScope(null, pool))
        {
            using var package = OoxPackage.Open(input, default);
            package.GetRelationships("/", default);
            package.LoadXml(TestAssert.NotNull(package.GetPart("/word/document.xml")), default);
            long prior = pool.Reserved;
            var document = new DocxReader().Read(package, null, default, OoxPdfDocxMarkupMode.Final);
            // Retained boxes measured about 2.56 MB, with 6.19 MB allocated during
            // preparation. Assert headroom, independently of the estimator formula.
            TestAssert.True(pool.Reserved - prior > 6 * 1024 * 1024,
                "Long supported literals need significant preparation capacity beyond ordinary text.");
            TestAssert.True(pool.Peak > pool.Reserved, "Source construction capacity refunds after parsing.");
            package.Dispose();
            TestAssert.True(pool.Reserved > 6 * 1024 * 1024,
                "Prepared-content capacity belongs to the source model, independently of package buffers.");
            TestAssert.True(document.Paragraphs.Count > 0, "Keep the model live through its ownership check.");
        }
        TestAssert.Equal(0L, pool.Reserved);
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        TestAssert.Equal(0, pool.CallerDisposals);
    }

    public static void ModelAdmissionMathLiteralRefusalAndCancellationPrecedePreparation()
    {
        foreach (bool cancel in new[] { false, true })
        {
            var pool = new TestMemoryBudget();
            using var input = MathLiteralAdmissionInput(8192);
            using var cancellation = new CancellationTokenSource();
            using (var operation = OoxConversionBudget.BeginScope(null, pool))
            using (var package = OoxPackage.Open(input, default))
            {
                package.GetRelationships("/", default);
                package.LoadXml(TestAssert.NotNull(package.GetPart("/word/document.xml")), default);
                long prior = pool.Reserved;
                bool requested = false;
                pool.RefuseChange = (before, after) => {
                    if (after - before < 6 * 1024 * 1024) return false;
                    requested = true;
                    if (cancel) cancellation.Cancel();
                    return !cancel;
                };
                if (cancel)
                    TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(() =>
                        new DocxReader().Read(package, null, cancellation.Token, OoxPdfDocxMarkupMode.Final)).CancellationToken);
                else
                    TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, TestAssert.Throws<OoxPdfLimitExceededException>(() =>
                        new DocxReader().Read(package, null, default, OoxPdfDocxMarkupMode.Final)).Kind);
                TestAssert.True(requested, "Warm XML must still admit math preparation with the derived model.");
                TestAssert.Equal(prior, pool.Reserved);
                pool.RefuseChange = null;
            }
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        }
    }

    public static void ModelAdmissionMathLiteralConversionRefundsAndPreservesPdf()
    {
        var font = new MemoryFontProgramSource("caller-owned-math", File.ReadAllBytes(Case("PortableFonts/DejaVuMathTeXGyre.ttf")));
        var resolver = new ProgramResolver(font);
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var pool = new TestMemoryBudget();
            byte[] Convert(IMemoryBudget? budget)
            {
                using var input = MathLiteralAdmissionInput(512);
                using var output = new MemoryStream();
                var diagnostics = new List<OoxPdfDiagnostic>();
                OoxPdfConverter.Convert(input, output, new OoxPdfOptions {
                    InputKind = OoxPdfInputKind.Docx, MemoryBudget = budget, FontResolver = resolver,
                    DocxMarkupMode = mode, FixedCreationDate = DateTimeOffset.UnixEpoch, DiagnosticSink = diagnostics.Add });
                TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_EQUATION"), "Exercise the supported mathematical painting path.");
                TestAssert.True(!diagnostics.Any(d => d.Fallback == "Ignored" && d.Feature == "equation"), "Keep the literal equation.");
                return output.ToArray();
            }
            byte[] expected = Convert(null), actual = Convert(pool);
            TestAssert.True(expected.AsSpan().SequenceEqual(actual), "Host admission must preserve exact rendering.");
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
            TestAssert.Equal(0, pool.CallerDisposals);
        }
    }

    private static MemoryStream MathLiteralAdmissionInput(int repetitions)
    {
        var parts = ModelAdmissionParts(File.ReadAllBytes(Case("PortableFonts/math-literal-inline-break-default-upright.docx")));
        var xml = XDocument.Parse(Encoding.UTF8.GetString(parts["word/document.xml"]));
        XNamespace math = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        xml.Descendants(math + "t").First().Value = string.Concat(Enumerable.Repeat("α+", repetitions)) + "α";
        parts["word/document.xml"] = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
        return TestFixtures.CreateZipPackage(parts);
    }
}
