using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
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

    public static void DocxDefaultStaticDistancesMatchIndependentWordMathAndBodyOrigins()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the independent default header/footer-distance origins.");
        string folder = Path.GetDirectoryName(DefaultStaticDistanceCase("header", 36))!;
        using var gold = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "static-default-distances-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string input = Path.Combine(folder, row.GetProperty("Input").GetString()!);
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            var mode = row.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var pages = new DocxRenderer(null, mode, CachedMathGeometry(mode)).RenderPagesAsync(DocxTests.ReadDocx(input, mode), null, default)
                .ToBlockingEnumerable().ToArray();
            TestAssert.Equal(1, pages.Length);
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(pages[0], true).ToArray();
            var math = shows.Where(show => pages[0].Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) &&
                ReferenceEquals(font.Font, show.Font))).ToArray();
            TestAssert.Equal("ratetimeratetime", string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
            TestAssert.Equal(16, math.Length);
            var anchors = row.GetProperty("MathOrigins").EnumerateArray().ToArray();
            for (int index = 0; index < anchors.Length; index++)
            {
                var actual = math[4 * index]; var expected = anchors[index];
                TestAssert.Equal(expected.GetProperty("Text").GetString(), actual.Text.Normalize(NormalizationForm.FormKC));
                TestAssert.True(Math.Abs(actual.X - expected.GetProperty("X").GetDouble()) < 1d &&
                    Math.Abs(actual.Y - expected.GetProperty("Y").GetDouble()) < 1d,
                    "Missing header/footer distances retain the existing one-point Word origin gate.");
            }
            var body = shows.First(show => show.Text.StartsWith("North", StringComparison.Ordinal));
            var expectedBody = row.GetProperty("BodyOrigin");
            TestAssert.True(Math.Abs(body.X - expectedBody.GetProperty("X").GetDouble()) < 1d &&
                Math.Abs(body.Y - expectedBody.GetProperty("Y").GetDouble()) < 1d,
                "The same source default also preserves Word's body admission beside the static story.");
        }
    }

    private static string DefaultStaticDistanceCase(string story, int margin) => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "CompositionProbes", "static-default-" + story + margin + ".docx"));

    private static DocxDocument ReadDefaultStaticDistance(string story, int margin, OoxPdfDocxMarkupMode mode, int? distance = null) =>
        ReadMutation("CompositionProbes/static-default-" + story + margin, mode, false, false,
            distance is null ? null : (xml, w, _) => xml.Descendants(w + "pgMar").Single()
                .SetAttributeValue(w + story, (distance.Value * 20).ToString(CultureInfo.InvariantCulture)));
}
