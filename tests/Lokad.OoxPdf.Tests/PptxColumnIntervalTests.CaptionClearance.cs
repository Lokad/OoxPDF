using System.Text.Json;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxColumnIntervalTests
{
    private static string CaptionClearanceCases => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "OuterCaptionClearance"));

    public static void PptxOuterRotatedCaptionClearanceIsSideIndependent()
    {
        using var oracle = JsonDocument.Parse(File.ReadAllText(Path.Combine(CaptionClearanceCases, "office.json")));
        var rows = oracle.RootElement.GetProperty("Rows").EnumerateArray().ToDictionary(row => row.GetProperty("Id").GetString()!);
        foreach (var pair in oracle.RootElement.GetProperty("HeightPairs").EnumerateArray())
        {
            var high = rows[pair[0].GetString()!];
            var low = rows[pair[1].GetString()!];
            TestAssert.True(Math.Abs(high.GetProperty("NativeInsideHeight").GetDouble() -
                low.GetProperty("NativeInsideHeight").GetDouble()) < .002d,
                "Independent Office controls reserve equal heights when rotating mirrored outer captions.");
            var upper = AdmissionBounds(RenderAdmission(high, caseDirectory: CaptionClearanceCases));
            var lower = AdmissionBounds(RenderAdmission(low, caseDirectory: CaptionClearanceCases));
            TestAssert.True(Math.Abs(upper.Bottom - upper.Top - lower.Bottom + lower.Top) < .002d,
                high.GetProperty("Id").GetString() + ": the measured clearance survives side, direction, font size and explicit offset changes.");
        }
    }

    public static void PptxOuterRotatedCaptionOffsetUsesMeasuredLineHeight()
    {
        using var oracle = JsonDocument.Parse(File.ReadAllText(Path.Combine(CaptionClearanceCases, "office.json")));
        var rows = oracle.RootElement.GetProperty("Rows").EnumerateArray().ToDictionary(row => row.GetProperty("Id").GetString()!);
        foreach (string side in new[] { "high", "low" })
        {
            var ordinary = rows["outer-authored45-" + side];
            var offset = rows["outer-" + side + "45-offset200"];
            var before = AdmissionBounds(RenderAdmission(ordinary, caseDirectory: CaptionClearanceCases));
            var after = AdmissionBounds(RenderAdmission(offset, caseDirectory: CaptionClearanceCases));
            double native = ordinary.GetProperty("NativeInsideHeight").GetDouble() - offset.GetProperty("NativeInsideHeight").GetDouble();
            TestAssert.True(Math.Abs(before.Bottom - before.Top - after.Bottom + after.Top - native) <= 1d,
                "The native offset increment survives pinned font substitution without multiplying the full baseline reserve.");
        }
    }

    public static void PptxOuterRotatedCaptionClearanceKeepsCompleteInkAndAngles()
    {
        using var oracle = JsonDocument.Parse(File.ReadAllText(Path.Combine(CaptionClearanceCases, "office.json")));
        foreach (var row in oracle.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var page = RenderAdmission(row, caseDirectory: CaptionClearanceCases);
            CheckCategoryOrderPage(page, row);
            PptxSourceUnitBoundsTests.CheckCaptionInk(page, row.GetProperty("ExpectedRendererAngle").GetInt32(),
                expectedCount: 8, frameRight: 792d, frameBottom: 72d, frameTop: 504d);
        }
    }
}
