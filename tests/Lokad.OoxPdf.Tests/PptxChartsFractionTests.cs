using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxChartsFractionTests
{
    public static void PptxSyntheticFractionCurveVaryColors()
    {
        var rendererType = typeof(PptxRenderer);
        var method = rendererType.GetMethod(
"TryResolveFractionCurveVaryColorsFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected fraction-curve engine to remain inspectable.");
        var rgbType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.RgbColor");
        TestAssert.True(rgbType is not null, "Expected RgbColor to remain resolvable.");
        // Gate: below the 120-point validated range the engine stays out (ladder keeps those counts).
        foreach (int count in new[] { 0, 6, 60, 114, 119 })
        {
            object?[] pastArgs = [5, count, null];
            TestAssert.Equal(false, (bool)method!.Invoke(null, pastArgs)!);
        }
        // Era bases (Office-observed, constant within each era).
        int[] eraIdx = [0, 1, 2, 3, 4, 5];
        // Era 0.
        int[] era0R = [47, 121, 96, 79, 44, 157];
        int[] era0G = [79, 47, 117, 60, 108, 93];
        int[] era0B = [119, 45, 53, 101, 124, 41];
        int era0N = 120;
        for (int s = 0; s < eraIdx.Length; s++)
        {
            object?[] eraArgs = [eraIdx[s], era0N, null];
            TestAssert.Equal(true, (bool)method!.Invoke(null, eraArgs)!);
            TestAssert.Equal((byte)era0R[s], (byte)rgbType!.GetProperty("Red")!.GetValue(eraArgs[2])!);
            TestAssert.Equal((byte)era0G[s], (byte)rgbType!.GetProperty("Green")!.GetValue(eraArgs[2])!);
            TestAssert.Equal((byte)era0B[s], (byte)rgbType!.GetProperty("Blue")!.GetValue(eraArgs[2])!);
        }
        // Era 1.
        int[] era1R = [46, 119, 95, 78, 43, 155];
        int[] era1G = [78, 47, 116, 60, 106, 92];
        int[] era1B = [117, 45, 53, 100, 123, 40];
        int era1N = 132;
        for (int s = 0; s < eraIdx.Length; s++)
        {
            object?[] eraArgs = [eraIdx[s], era1N, null];
            TestAssert.Equal(true, (bool)method!.Invoke(null, eraArgs)!);
            TestAssert.Equal((byte)era1R[s], (byte)rgbType!.GetProperty("Red")!.GetValue(eraArgs[2])!);
            TestAssert.Equal((byte)era1G[s], (byte)rgbType!.GetProperty("Green")!.GetValue(eraArgs[2])!);
            TestAssert.Equal((byte)era1B[s], (byte)rgbType!.GetProperty("Blue")!.GetValue(eraArgs[2])!);
        }
        // Era 2.
        int[] era2R = [45, 117, 94, 77, 43, 152];
        int[] era2G = [77, 46, 114, 59, 105, 91];
        int[] era2B = [115, 44, 52, 98, 121, 40];
        int era2N = 162;
        for (int s = 0; s < eraIdx.Length; s++)
        {
            object?[] eraArgs = [eraIdx[s], era2N, null];
            TestAssert.Equal(true, (bool)method!.Invoke(null, eraArgs)!);
            TestAssert.Equal((byte)era2R[s], (byte)rgbType!.GetProperty("Red")!.GetValue(eraArgs[2])!);
            TestAssert.Equal((byte)era2G[s], (byte)rgbType!.GetProperty("Green")!.GetValue(eraArgs[2])!);
            TestAssert.Equal((byte)era2B[s], (byte)rgbType!.GetProperty("Blue")!.GetValue(eraArgs[2])!);
        }
        // Spot cells where the engine reproduces the Office render bit-exactly
        // (every literal below is the Office-observed fill; engine output matched before pinning).
        int[] spotCat = [0, 25, 44, 59, 80, 125, 3, 12, 21, 30, 39, 54, 63, 72, 81, 90, 99, 114, 123, 132, 141, 2, 16, 46, 56, 78, 119, 140, 154, 10, 22, 49, 58, 67, 76, 88, 127, 144, 160, 169, 4, 33, 53, 62, 78, 90, 105, 135];
        int[] spotCount = [132, 132, 132, 132, 132, 132, 150, 150, 150, 150, 150, 150, 150, 150, 150, 150, 150, 150, 150, 150, 150, 162, 162, 162, 162, 162, 162, 162, 162, 174, 174, 174, 174, 174, 174, 174, 174, 174, 174, 174, 180, 180, 180, 180, 180, 180, 180, 180];
        int[] spotR = [46, 151, 137, 233, 170, 251, 78, 53, 92, 62, 105, 71, 120, 78, 132, 120, 161, 166, 190, 192, 209, 94, 49, 62, 138, 78, 249, 208, 201, 46, 52, 164, 65, 178, 70, 74, 213, 182, 195, 232, 43, 98, 210, 138, 75, 79, 144, 179];
        int[] spotG = [78, 61, 165, 141, 196, 213, 60, 89, 71, 102, 82, 117, 94, 127, 106, 152, 145, 184, 180, 204, 203, 114, 118, 144, 167, 127, 185, 221, 223, 112, 123, 67, 151, 73, 162, 170, 159, 195, 220, 208, 105, 76, 127, 167, 123, 129, 123, 168];
        int[] spotB = [117, 59, 78, 66, 123, 196, 100, 132, 118, 151, 134, 172, 152, 186, 165, 198, 183, 214, 204, 225, 219, 52, 137, 167, 79, 186, 150, 188, 232, 129, 142, 64, 174, 71, 187, 195, 158, 220, 229, 208, 121, 124, 58, 79, 181, 189, 172, 196];
        for (int s = 0; s < spotCat.Length; s++)
        {
            object?[] spotArgs = [spotCat[s], spotCount[s], null];
            TestAssert.Equal(true, (bool)method!.Invoke(null, spotArgs)!);
            TestAssert.Equal((byte)spotR[s], (byte)rgbType!.GetProperty("Red")!.GetValue(spotArgs[2])!);
            TestAssert.Equal((byte)spotG[s], (byte)rgbType!.GetProperty("Green")!.GetValue(spotArgs[2])!);
            TestAssert.Equal((byte)spotB[s], (byte)rgbType!.GetProperty("Blue")!.GetValue(spotArgs[2])!);
        }
    }
}
