using System.Text.Json;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxColumnIntervalTests
{
    private static string LogicalCaptionCases => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "LogicalCaptionFrames"));
    private static JsonDocument LogicalCaptionOracle() => JsonDocument.Parse(File.ReadAllText(Path.Combine(LogicalCaptionCases, "office.json")));
    private static JsonElement LogicalCaptionRow(JsonDocument oracle, string id) => oracle.RootElement.GetProperty("Rows")
        .EnumerateArray().Single(row => row.GetProperty("Id").GetString() == id);
    private static PdfPage RenderLogicalCaption(JsonDocument oracle, string id) => RenderAdmission(
        LogicalCaptionRow(oracle, id), caseDirectory: LogicalCaptionCases);

    public static void PptxOuterCaptionReferenceIgnoresNumericAndDefaultTextRoles()
    {
        using var oracle = LogicalCaptionOracle();
        const string baseline = "outer-high45-no-data-labels";
        double right = AdmissionBounds(RenderLogicalCaption(oracle, baseline)).Right;
        double nativeRight = LogicalCaptionRow(oracle, baseline).GetProperty("NativeFrame")[2].GetDouble();
        foreach (string id in new[] { "outer-high45-no-legend", "outer-high45-literal", "reference-root12", "reference-root36",
            "reference-root-times24", "reference-numeric-format" })
        {
            TestAssert.True(Math.Abs(LogicalCaptionRow(oracle, id).GetProperty("NativeFrame")[2].GetDouble() - nativeRight) < .02d,
                "Independent native controls isolate the authored reference from unrelated text and numeric gutters.");
            NearHorizontal(right, AdmissionBounds(RenderLogicalCaption(oracle, id)).Right);
        }
        TestAssert.True(AdmissionBounds(RenderLogicalCaption(oracle, "reference-numeric-format")).Left >
            AdmissionBounds(RenderLogicalCaption(oracle, baseline)).Left,
            "Longer numeric captions still own their actual side band.");
    }

    public static void PptxOuterCaptionReferenceTracksSourcePitchAndOrdinalReversal()
    {
        using var oracle = LogicalCaptionOracle();
        void CheckMovement(string first, string second)
        {
            var a = AdmissionBounds(RenderLogicalCaption(oracle, first));
            var b = AdmissionBounds(RenderLogicalCaption(oracle, second));
            var na = LogicalCaptionRow(oracle, first).GetProperty("NativeFrame");
            var nb = LogicalCaptionRow(oracle, second).GetProperty("NativeFrame");
            TestAssert.True(Math.Abs((b.Left - a.Left) - (nb[0].GetDouble() - na[0].GetDouble())) < .02d,
                "The left reference follows independently exported source-pitch movement.");
            TestAssert.True(Math.Abs((b.Right - a.Right) - (nb[2].GetDouble() - na[2].GetDouble())) < .02d,
                "The right reference follows independently exported source-pitch movement.");
        }
        // Font substitution changes absolute text spans; these native source-
        // pitch movements remain independent of the selected face.
        CheckMovement("outer-high45-no-data-labels", "reference-count12");
        CheckMovement("logical-hidden-numeric", "logical-reversed-hidden8");
        CheckMovement("logical-reversed-hidden8", "logical-reversed-hidden12");
        NearHorizontal(576d, AdmissionBounds(RenderLogicalCaption(oracle, "reference-count4")).Right);
    }

    public static void PptxOuterCaptionReferenceKeepsEndpointContentAndWholeChartInk()
    {
        using var oracle = LogicalCaptionOracle();
        foreach (var row in oracle.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var page = RenderAdmission(row, caseDirectory: LogicalCaptionCases);
            var expected = row.GetProperty("ExpectedCaptions").EnumerateArray().Select(value => value.GetString()!).ToArray();
            string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).Select(show => show.Text));
            TestAssert.Equal(string.Concat(expected), text);
            PptxSourceUnitBoundsTests.CheckCaptionInk(page, row.GetProperty("ExpectedAngle").GetInt32(),
                expectedCount: expected.Length, frameRight: 792d, frameBottom: 72d, frameTop: 504d);
        }
        var baseline = AdmissionBounds(RenderLogicalCaption(oracle, "outer-high45-literal"));
        var middle = AdmissionBounds(RenderLogicalCaption(oracle, "outer-high45-middle-wide"));
        var last = AdmissionBounds(RenderLogicalCaption(oracle, "outer-high45-last-wide"));
        NearHorizontal(baseline.Right, middle.Right);
        NearHorizontal(last.Bottom - last.Top, middle.Bottom - middle.Top);
        TestAssert.True(last.Right < middle.Right,
            "A longer endpoint allocates horizontal space; the same interior caption owns vertical extent without becoming the endpoint.");
    }
}
