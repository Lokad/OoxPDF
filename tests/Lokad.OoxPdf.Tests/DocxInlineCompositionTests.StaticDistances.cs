using System.Globalization;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxDefaultStaticDistancesKeepMissingAndExplicitPlacementEquivalent()
    {
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (string story in new[] { "header", "footer" })
        foreach (int margin in new[] { 36, 108 })
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            var geometry = mode == OoxPdfDocxMarkupMode.AllMarkup ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout;
            var renderer = new DocxRenderer(resolver, mode, geometry);
            var actual = renderer.RenderPagesAsync(ReadDefaultStaticDistance(story, margin, mode), null, default)
                .ToBlockingEnumerable().ToArray();
            var expected = renderer.RenderPagesAsync(ReadDefaultStaticDistance(story, margin, mode, distance: 36), null, default)
                .ToBlockingEnumerable().ToArray();
            TestAssert.Equal(expected.Length, actual.Length);
            for (int index = 0; index < expected.Length; index++)
            {
                TestAssert.Equal(expected[index].Content, actual[index].Content);
                TestAssert.Equal(expected[index].Fonts.Count, actual[index].Fonts.Count);
                TestAssert.Equal(expected[index].Images.Count, actual[index].Images.Count);
                TestAssert.True(expected[index].Annotations.SequenceEqual(actual[index].Annotations),
                    "An omitted distance has the same placed content and navigation as Word's explicit half-inch setting.");
            }
        }
    }

    public static void DocxDefaultStaticDistancesRetainRawAndExplicitSettingsAndCancellation()
    {
        foreach (string story in new[] { "header", "footer" })
        foreach (int margin in new[] { 36, 108 })
        {
            var implicitDocument = ReadDefaultStaticDistance(story, margin, OoxPdfDocxMarkupMode.Final);
            double? RawDistance(DocxDocument document) => story == "header"
                ? document.PageSettings.HeaderDistancePoints : document.PageSettings.FooterDistancePoints;
            TestAssert.True(RawDistance(implicitDocument) is null,
                "Layout resolution must not rewrite an absent source token into explicit inspection metadata.");
            foreach (int value in new[] { 0, 18, 36, 54 })
                TestAssert.Equal((double)value, RawDistance(ReadDefaultStaticDistance(story, margin, OoxPdfDocxMarkupMode.Final, value))!.Value);
        }
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        using var input = File.OpenRead(DefaultStaticDistanceCase("footer", 108));
        using var output = new MemoryStream();
        TestAssert.Throws<OperationCanceledException>(() => OoxPdfConverter.Convert(input, output,
            new OoxPdfOptions { FontResolver = new PortableRenderingTests.PinnedResolver(mathFont: true) }, canceled.Token));
        TestAssert.Equal(0L, output.Length);
    }

    public static void DocxDefaultStaticDistancesMatchIndependentWordMathAndBodyOrigins() =>
        CheckIndependentRateTimeOrigins("static-default-distances-office.json",
            ["static-default-header36.docx", "static-default-header108.docx",
                "static-default-footer36.docx", "static-default-footer108.docx"], checkBodyOrigin: true);

    private static string DefaultStaticDistanceCase(string story, int margin) => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "CompositionProbes", "static-default-" + story + margin + ".docx"));

    private static DocxDocument ReadDefaultStaticDistance(string story, int margin, OoxPdfDocxMarkupMode mode, int? distance = null) =>
        ReadMutation("CompositionProbes/static-default-" + story + margin, mode, false, false,
            distance is null ? null : (xml, w, _) => xml.Descendants(w + "pgMar").Single()
                .SetAttributeValue(w + story, (distance.Value * 20).ToString(CultureInfo.InvariantCulture)));
}
