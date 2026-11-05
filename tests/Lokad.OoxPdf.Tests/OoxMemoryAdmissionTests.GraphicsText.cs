using System.Security.Cryptography;
using System.Text;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    public static void MemoryBudgetGraphicsTextRetainsBufferAndSnapshotWithoutChangingOperators()
    {
        string hex = new('0', 32 * 1024);
        var plain = new PdfGraphicsBuilder();
        foreach (string kind in new[] { "glyph", "positioned", "fallback", "actual" }) DrawAdmissionText(plain, kind, hex);
        var budget = new TestMemoryBudget();
        using (var scope = budget.CreateScope())
        using (var grant = OoxMemory.Reserve(scope, 64 * 1024)!)
        {
            var graphics = new PdfGraphicsBuilder(grant, 0);
            foreach (string kind in new[] { "glyph", "positioned", "fallback", "actual" }) DrawAdmissionText(graphics, kind, hex);
            string content = graphics.ToString();
            TestAssert.Equal(plain.ToString(), content);
            TestAssert.True(grant.Bytes >= 4L * content.Length,
                "The borrowed page grant covers its UTF16 builder and immutable snapshot while the consumer uses them.");
            TestAssert.Equal(grant.Bytes, budget.Reserved);
        }
        TestAssert.Equal(0L, budget.Reserved);
        TestAssert.Equal(budget.ScopesCreated, budget.ScopesDisposed);
        TestAssert.Equal(0, budget.CallerDisposals);
    }

    public static void MemoryBudgetGraphicsTextRefusalAndCancellationPrecedeOperators()
    {
        foreach (string kind in new[] { "glyph", "positioned", "fallback", "actual" })
        foreach (bool cancel in new[] { false, true })
        {
            var budget = new TestMemoryBudget();
            using (var scope = budget.CreateScope())
            using (var grant = OoxMemory.Reserve(scope, 64 * 1024)!)
            using (var cancellation = new CancellationTokenSource())
            {
                var graphics = new PdfGraphicsBuilder(grant, 0);
                int attempts = 0;
                budget.RefuseChange = (old, next) =>
                {
                    if (next <= old) return false;
                    attempts++;
                    TestAssert.Equal(0, graphics.ContentLength);
                    TestAssert.Equal(0, graphics.UsedFontResourceNames.Count);
                    if (cancel) { cancellation.Cancel(); throw new OperationCanceledException(cancellation.Token); }
                    return true;
                };
                if (cancel)
                {
                    var error = TestAssert.Throws<OperationCanceledException>(() => DrawAdmissionText(graphics, kind, "0001"));
                    TestAssert.Equal(cancellation.Token, error.CancellationToken);
                }
                else
                {
                    var error = TestAssert.Throws<OoxPdfLimitExceededException>(() => DrawAdmissionText(graphics, kind, "0001"));
                    TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, error.Kind);
                    TestAssert.Equal(OoxPdfLimitScope.Conversion, error.Scope);
                }
                TestAssert.Equal(1, attempts);
                TestAssert.Equal(0, graphics.ContentLength);
                TestAssert.Equal(64L * 1024, grant.Bytes);
            }
            TestAssert.Equal(0L, budget.Reserved);
            TestAssert.Equal(budget.ScopesCreated, budget.ScopesDisposed);
        }
    }

    public static void MemoryBudgetPptxFormattedLabelsFundTheirActualContent()
    {
        var governed = ObserveFormattedLabelAdmission(governed: true);
        var plain = ObserveFormattedLabelAdmission(governed: false);
        TestAssert.Equal(plain.Hash, governed.Hash);
        TestAssert.Equal(plain.Characters, governed.Characters);
    }

    private static (string Hash, int Characters) ObserveFormattedLabelAdmission(bool governed)
    {
        string values = string.Concat(Enumerable.Range(0, 128).Select(i => $"<c:pt idx=\"{i}\"><c:v>{i + 1}</c:v></c:pt>"));
        string format = "&quot;" + new string('X', 2048) + "&quot;0";
        using var input = PptxTests.CreateSingleChartPackageStream($$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea>
            <c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit><c:ptCount val="128"/>{{values}}</c:numLit></c:val></c:ser>
            <c:dLbls><c:numFmt formatCode="{{format}}" sourceLinked="0"/><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="600"><a:latin typeface="DejaVu Sans"/></a:defRPr></a:pPr></a:p></c:txPr><c:dLblPos val="outEnd"/><c:showVal val="1"/></c:dLbls>
            <c:axId val="10"/><c:axId val="20"/></c:barChart><c:catAx><c:axId val="10"/><c:delete val="1"/><c:crossAx val="20"/></c:catAx><c:valAx><c:axId val="20"/><c:delete val="1"/><c:crossAx val="10"/></c:valAx>
            </c:plotArea></c:chart></c:chartSpace>
            """);
        using var package = OoxPackage.Open(input, default);
        var document = new PptxReader().Read(package, default);
        var budget = new TestMemoryBudget();
        string hash; int characters;
        using (var operation = OoxConversionBudget.BeginScope(null, governed ? budget : null))
        {
            var iterator = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
                .RenderPagesAsync(document, package, null, default).GetAsyncEnumerator();
            try
            {
                TestAssert.True(iterator.MoveNextAsync().AsTask().GetAwaiter().GetResult(), "Render formatted labels.");
                string content = iterator.Current.Content;
                characters = content.Length;
                hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(content)));
                if (governed)
                {
                    TestAssert.True(budget.Reserved >= 4L * characters,
                        "A shared number-format literal must fund generated page text, not only its short source declaration.");
                }
            }
            finally { iterator.DisposeAsync().AsTask().GetAwaiter().GetResult(); }
        }
        TestAssert.Equal(0L, budget.Reserved);
        TestAssert.Equal(budget.ScopesCreated, budget.ScopesDisposed);
        return (hash, characters);
    }

    private static void DrawAdmissionText(PdfGraphicsBuilder graphics, string kind, string hex)
    {
        if (kind == "glyph") graphics.DrawGlyphText("F0", 12, 1, 2, 0, 0, 0, hex, false, 0, 0, 0, 0, 0, 0);
        else if (kind == "positioned") graphics.DrawGlyphPositionedText("F0", 12, 1, 2, 0, 0, 0, "[<" + hex + ">]", false, 0, 0, 0, 0, 0, 0, textRotationQuarterTurns: 1);
        else if (kind == "fallback") graphics.DrawFallbackText("F0", 12, 0, 0, 0, [new PdfFallbackGlyph(1, 2, 65)]);
        else { graphics.BeginActualText(0x1F642); graphics.EndActualText(); }
    }
}
