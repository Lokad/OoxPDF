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
        int[] spotCat = [0, 26, 42, 53, 75, 1, 15, 32, 49, 59, 81, 105, 126, 4, 14, 24, 34, 50, 60, 70, 80, 90, 100, 116, 126, 136, 147, 16, 51, 63, 79, 106, 140, 0, 20, 42, 52, 62, 72, 83, 113, 137, 159, 169, 5, 51, 69, 89];
        int[] spotCount = [126, 126, 126, 126, 126, 138, 138, 138, 138, 138, 138, 138, 138, 150, 150, 150, 150, 150, 150, 150, 150, 150, 150, 150, 150, 150, 150, 162, 162, 162, 162, 162, 162, 174, 174, 174, 174, 174, 174, 174, 174, 174, 174, 174, 180, 180, 180, 180];
        int[] spotR = [47, 123, 70, 229, 146, 119, 89, 125, 174, 231, 144, 179, 198, 43, 107, 58, 58, 137, 74, 72, 158, 120, 134, 196, 185, 192, 215, 49, 111, 117, 189, 129, 208, 45, 110, 64, 63, 139, 74, 239, 248, 250, 206, 232, 152, 108, 117, 241];
        int[] spotG = [79, 148, 115, 139, 125, 47, 68, 152, 72, 140, 123, 168, 208, 106, 130, 97, 137, 165, 121, 166, 189, 152, 191, 213, 198, 218, 210, 118, 86, 91, 79, 189, 221, 77, 134, 106, 146, 168, 122, 145, 171, 193, 199, 208, 91, 84, 91, 147];
        int[] spotB = [119, 69, 169, 64, 173, 45, 113, 71, 69, 65, 172, 196, 227, 123, 60, 144, 159, 78, 178, 191, 96, 198, 210, 169, 221, 228, 223, 137, 140, 149, 76, 209, 188, 115, 62, 157, 169, 79, 178, 67, 124, 164, 216, 208, 40, 137, 148, 68];
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
