using System.Text.Json;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInspectionTests
{
    public static void DocxCombinedInspectionPreservesIndependentSnapshotViews()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.SimpleMarkup, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string source = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "Composition", "report.docx"));
            Compare(DocxTests.ReadDocx(source, mode), mode, mode == OoxPdfDocxMarkupMode.AllMarkup
                ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        }
        // Numbering can change the prepared document and the final font plan.
        var notes = DocxPageFootnoteNumberingTests.ReadFixture();
        TestAssert.True(DocxPageFootnoteNumbering.IsEligibleDocument(notes),
            "The note witness must exercise inspection of the final prepared font plan.");
        Compare(notes, OoxPdfDocxMarkupMode.AllMarkup,
            OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);

        static void Compare(DocxDocument document, OoxPdfDocxMarkupMode mode, OoxPdfDocxMarkupGeometryMode geometry)
        {
            var separate = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode, geometry);
            var layout = separate.InspectLayout(document);
            var fonts = separate.InspectFontPlan(document);
            var emission = separate.InspectTextEmission(document);
            var balloons = separate.InspectMarkupBalloons(document);
            var shared = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode, geometry)
                .InspectAsync(document).AsTask().GetAwaiter().GetResult();
            Same(layout, shared.Layout, "Layout");
            Same(fonts, shared.FontPlan, "Requested/final font plan");
            Same(emission, shared.TextEmission, "Text emission");
            Same(balloons, shared.MarkupBalloons, "Markup balloons");
        }
    }

    public static void DocxCombinedInspectionRetainsSourceFontPlanBeforeSubstitution()
    {
        var source = InspectionDocument("CffFamily");
        var renderer = new DocxRenderer(new DocxTests.CffPrimaryFontResolver(
            TestFontBuilder.CreateCffKindFont("CffFamily"), TestFontBuilder.CreateTestFont()),
            OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var expected = renderer.InspectFontPlan(source);
        var shared = renderer.InspectAsync(source).AsTask().GetAwaiter().GetResult();
        TestAssert.True(expected.PrimaryCount > 0 && expected.MissingCount == 0,
            "The source font plan records the requested CFF face, independently of embedding fallback.");
        Same(expected, shared.FontPlan, "Unsubstituted source font plan");
        TestAssert.True(shared.TextEmission.LineCount > 0, "Embedding fallback still supplies the emission view.");
    }

    public static void DocxCombinedInspectionPreparesLayoutOnceWithoutGlobalCaching()
    {
        var source = InspectionDocument();
        var singleResolver = new InspectionResolver();
        var separate = new DocxRenderer(singleResolver, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        separate.InspectLayout(source);
        separate.InspectFontPlan(source); // The distinct source-plan view remains necessary.
        var combinedResolver = new InspectionResolver();
        var combined = new DocxRenderer(combinedResolver, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        combined.InspectAsync(source).AsTask().GetAwaiter().GetResult();
        TestAssert.True(singleResolver.Requests.Count > 0 && singleResolver.Requests.SequenceEqual(combinedResolver.Requests),
            "Deriving all views must not repeat the layout/font acquisition path.");
        int first = combinedResolver.Requests.Count;
        combined.InspectAsync(source).AsTask().GetAwaiter().GetResult();
        TestAssert.Equal(first * 2, combinedResolver.Requests.Count);
    }

    public static void DocxCombinedInspectionCancellationAndRefusalRefundPreparation()
    {
        var source = InspectionDocument();
        using (var canceled = new CancellationTokenSource())
        {
            canceled.Cancel();
            var resolver = new InspectionResolver();
            var renderer = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            TestAssert.Throws<OperationCanceledException>(() => renderer.InspectAsync(source, canceled.Token).AsTask().GetAwaiter().GetResult());
            TestAssert.Equal(0, resolver.Requests.Count);
        }
        foreach (bool cancel in new[] { true, false })
        {
            using var cancellation = new CancellationTokenSource();
            bool admitted = false;
            var pool = new TestMemoryBudget { RefuseChange = (_, after) => {
                if (after <= 0) return false;
                admitted = true;
                if (cancel) cancellation.Cancel();
                return !cancel;
            }};
            using (OoxConversionBudget.BeginScope(null, pool))
            {
                var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), OoxPdfDocxMarkupMode.Final,
                    OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
                if (cancel)
                    TestAssert.Throws<OperationCanceledException>(() => renderer.InspectAsync(source, cancellation.Token).AsTask().GetAwaiter().GetResult());
                else
                    TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, TestAssert.Throws<OoxPdfLimitExceededException>(() =>
                        renderer.InspectAsync(source).AsTask().GetAwaiter().GetResult()).Kind);
            }
            TestAssert.True(admitted, "The failure follows a real shared-memory admission request.");
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
            TestAssert.Equal(0, pool.CallerDisposals);
        }
    }

    private static void Same<T>(T expected, T actual, string view) => TestAssert.True(
        JsonSerializer.Serialize(expected) == JsonSerializer.Serialize(actual), "Shared inspection preserves the independent " + view + " snapshot.");

    private static DocxDocument InspectionDocument(string family = "InspectionFamily") => DocxTests.CreateFontPlanDocument(
        new DocxTextRun("AV Alpha Beta", 12d, null, false, false, false, null, family)
        { Fonts = new DocxRunFonts(family, null, null, null, null, null, null, null) }, new DocxFontCatalog([], DocxThemeFonts.Empty));

    private sealed class InspectionResolver : IFontResolver
    {
        private readonly PortableRenderingTests.PinnedResolver inner = new();
        public List<FontRequest> Requests { get; } = [];
        public FontFaceResolution Resolve(FontRequest request)
        {
            Requests.Add(request);
            return inner.Resolve(request);
        }
    }
}
