using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxColumnIntervalTests
{
    private static string LogicalCaptionCases => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "LogicalCaptionFrames"));
    private static JsonDocument LogicalCaptionOracle() => JsonDocument.Parse(File.ReadAllText(Path.Combine(LogicalCaptionCases, "office.json")));
    private static JsonElement LogicalCaptionRow(JsonDocument oracle, string id) => oracle.RootElement.GetProperty("Rows")
        .EnumerateArray().Single(row => row.GetProperty("Id").GetString() == id);

    private const string LogicalCaptionBaseline = "outer-high45-no-data-labels";
    private const string LogicalCaptionLiteralAlias = "outer-high45-literal";
    private static readonly string[] LogicalCaptionUnrelatedRoles =
    [
        "outer-high45-no-legend", "reference-root12", "reference-root36",
        "reference-root-times24", "reference-numeric-format"
    ];
    private static readonly (string First, string Second)[] LogicalCaptionPitchChanges =
    [
        (LogicalCaptionBaseline, "reference-count12"),
        ("logical-hidden-numeric", "logical-reversed-hidden8"),
        ("logical-reversed-hidden8", "logical-reversed-hidden12")
    ];

    private static HashSet<string> LogicalCaptionFrameCases()
    {
        var cases = new HashSet<string>(LogicalCaptionUnrelatedRoles, StringComparer.Ordinal)
        {
            "reference-count4", "outer-high45-middle-wide", "outer-high45-last-wide"
        };
        foreach (var (first, second) in LogicalCaptionPitchChanges)
        {
            cases.Add(first);
            cases.Add(second);
        }

        return cases;
    }

    public static void PptxLogicalCaptionFramesPreserveSourceGeometryAndInk()
    {
        using var oracle = LogicalCaptionOracle();
        var geometryCases = LogicalCaptionFrameCases();
        var bounds = new Dictionary<string, (double Left, double Top, double Right, double Bottom)>(StringComparer.Ordinal);
        foreach (var row in oracle.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            if (!geometryCases.Contains(id))
            {
                continue;
            }

            PdfPage page = RenderAdmission(row, caseDirectory: LogicalCaptionCases);
            CheckLogicalCaptionInk(page, row);
            bounds.Add(id, AdmissionBounds(page));
            if (id is LogicalCaptionBaseline or "reference-numeric-format" or "logical-reversed-hidden12")
            {
                CheckDeletedSeriesCaptionsAndLegend(page, id);
            }
        }

        var baseline = bounds[LogicalCaptionBaseline];
        double nativeRight = LogicalCaptionRow(oracle, LogicalCaptionBaseline).GetProperty("NativeFrame")[2].GetDouble();
        foreach (string id in LogicalCaptionUnrelatedRoles)
        {
            TestAssert.True(Math.Abs(LogicalCaptionRow(oracle, id).GetProperty("NativeFrame")[2].GetDouble() - nativeRight) < .02d,
                "Independent native controls isolate the authored reference from unrelated text and numeric gutters.");
            NearHorizontal(baseline.Right, bounds[id].Right);
        }
        TestAssert.True(bounds["reference-numeric-format"].Left > baseline.Left,
            "Longer numeric captions still own their actual side band.");

        // Font substitution changes absolute text spans. The native movement
        // caused by source pitch remains independent of the selected face.
        foreach (var (first, second) in LogicalCaptionPitchChanges)
        {
            var a = bounds[first];
            var b = bounds[second];
            var nativeA = LogicalCaptionRow(oracle, first).GetProperty("NativeFrame");
            var nativeB = LogicalCaptionRow(oracle, second).GetProperty("NativeFrame");
            TestAssert.True(Math.Abs((b.Left - a.Left) - (nativeB[0].GetDouble() - nativeA[0].GetDouble())) < .02d,
                "The left reference follows independently exported source-pitch movement.");
            TestAssert.True(Math.Abs((b.Right - a.Right) - (nativeB[2].GetDouble() - nativeA[2].GetDouble())) < .02d,
                "The right reference follows independently exported source-pitch movement.");
        }
        NearHorizontal(576d, bounds["reference-count4"].Right);

        var middle = bounds["outer-high45-middle-wide"];
        var last = bounds["outer-high45-last-wide"];
        NearHorizontal(baseline.Right, middle.Right);
        NearHorizontal(last.Bottom - last.Top, middle.Bottom - middle.Top);
        TestAssert.True(last.Right < middle.Right,
            "A longer endpoint allocates horizontal space; the same interior caption owns vertical extent without becoming the endpoint.");

        // Change just the category carrier. The formula uses the eight pinned
        // cached captions, so the assertion is source equivalence within Lokad.OoxPdf.
        var literal = LogicalCaptionRow(oracle, LogicalCaptionBaseline);
        PdfPage cached = RenderAdmission(literal, caseDirectory: LogicalCaptionCases, editChart: chart =>
        {
            XElement categories = chart.Descendants(CrossingC + "cat").Single();
            XElement literals = categories.Elements().Single();
            TestAssert.Equal(CrossingC + "strLit", literals.Name);
            literals.ReplaceWith(new XElement(CrossingC + "strRef",
                new XElement(CrossingC + "f", "Sheet1!$A$2:$A$9"),
                new XElement(CrossingC + "strCache", literals.Elements())));
        });
        CheckLogicalCaptionInk(cached, literal);
        TestAssert.True(baseline == AdmissionBounds(cached), "A cached reference preserves the literal categories' plot frame.");
    }

    public static void PptxLogicalCaptionVariantsKeepExpectedInk()
    {
        using var oracle = LogicalCaptionOracle();
        var geometryCases = LogicalCaptionFrameCases();
        foreach (var row in oracle.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            // The Office literal alias has the same authored bytes as LogicalCaptionBaseline.
            if (geometryCases.Contains(id) || id == LogicalCaptionLiteralAlias)
            {
                continue;
            }

            PdfPage page = RenderAdmission(row, caseDirectory: LogicalCaptionCases);
            CheckLogicalCaptionInk(page, row);
        }
    }

    private static void CheckDeletedSeriesCaptionsAndLegend(PdfPage page, string id)
    {
        TestAssert.True(!page.Fonts.Any(font => font.ResourceName.StartsWith("CBD", StringComparison.Ordinal)),
            "A deleted series with no point requests emits no data-label resources or paint.");
        var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false);
        TestAssert.Equal(1, shows.Count(show => show.Text.Contains("Regional", StringComparison.Ordinal)));
        if (id == "reference-numeric-format")
        {
            TestAssert.True(AdmissionNumericShows(page).Select(show => show.Text)
                .SequenceEqual(new[] { "7.00", "32.00", "57.00", "82.00", "107.00" }),
                "Suppressing series labels preserves complete formatted numeric-axis captions.");
        }
    }

    private static void CheckLogicalCaptionInk(PdfPage page, JsonElement row)
    {
        var expected = row.GetProperty("ExpectedCaptions").EnumerateArray().Select(value => value.GetString()!).ToArray();
        string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
            fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).Select(show => show.Text));
        TestAssert.Equal(string.Concat(expected), text);
        PptxSourceUnitBoundsTests.CheckCaptionInk(page, row.GetProperty("ExpectedAngle").GetInt32(),
            expectedCount: expected.Length, frameRight: 792d, frameBottom: 72d, frameTop: 504d);
    }
}
