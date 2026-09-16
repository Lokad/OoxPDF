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
        var colorMapType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.PptxColorMap");
        TestAssert.True(colorMapType is not null, "Expected PptxColorMap to remain resolvable.");
        object? colorMap = colorMapType!.GetProperty("Default")!.GetValue(null);
        var paletteType = typeof(System.Collections.Generic.List<>).MakeGenericType(rgbType!);
        object? palette = System.Activator.CreateInstance(paletteType);
        var addMethod = paletteType.GetMethod("Add");
        TestAssert.True(addMethod is not null, "Expected List.Add to remain resolvable.");
        int[][] stockAccents = [[79,129,189], [192,80,77], [155,187,89], [128,100,162], [75,172,198], [247,150,70]];
        foreach (int[] accent in stockAccents)
        {
            object? accentColor = System.Activator.CreateInstance(rgbType!, (byte)accent[0], (byte)accent[1], (byte)accent[2]);
            addMethod!.Invoke(palette, [accentColor]);
        }
        // Gate: below the 96-point validated range the engine stays out (ladder keeps those counts).
        foreach (int count in new[] { 0, 6, 60, 90 })
        {
            object?[] pastArgs = [null, null, colorMap, 5, count, null];
            TestAssert.Equal(false, (bool)method!.Invoke(null, pastArgs)!);
        }
        // Gate: the 96-point floor takes the engine.
        foreach (int floorCount in new[] { 96, 114, 119 })
        {
            object?[] floorArgs = [palette, null, colorMap, 5, floorCount, null];
            TestAssert.Equal(true, (bool)method!.Invoke(null, floorArgs)!);
        }
        // Era shade bases (engine-computed from stock accents; Office-verified within 2).
        int[] eraIdx = [0, 1, 2, 3, 4, 5];
        // Era 0 (N=100).
        int[] era0R = [47, 122, 97, 80, 45, 158];
        int[] era0G = [80, 48, 119, 61, 109, 94];
        int[] era0B = [120, 46, 54, 102, 126, 41];
        int era0N = 100;
        for (int s = 0; s < eraIdx.Length; s++)
        {
            object?[] eraArgs = [palette, null, colorMap, eraIdx[s], era0N, null];
            TestAssert.Equal(true, (bool)method!.Invoke(null, eraArgs)!);
            TestAssert.Equal((byte)era0R[s], (byte)rgbType!.GetProperty("Red")!.GetValue(eraArgs[5])!);
            TestAssert.Equal((byte)era0G[s], (byte)rgbType!.GetProperty("Green")!.GetValue(eraArgs[5])!);
            TestAssert.Equal((byte)era0B[s], (byte)rgbType!.GetProperty("Blue")!.GetValue(eraArgs[5])!);
        }
        // Era 1 (N=140).
        int[] era1R = [46, 119, 95, 78, 44, 155];
        int[] era1G = [78, 47, 116, 60, 106, 92];
        int[] era1B = [117, 45, 53, 100, 123, 40];
        int era1N = 140;
        for (int s = 0; s < eraIdx.Length; s++)
        {
            object?[] eraArgs = [palette, null, colorMap, eraIdx[s], era1N, null];
            TestAssert.Equal(true, (bool)method!.Invoke(null, eraArgs)!);
            TestAssert.Equal((byte)era1R[s], (byte)rgbType!.GetProperty("Red")!.GetValue(eraArgs[5])!);
            TestAssert.Equal((byte)era1G[s], (byte)rgbType!.GetProperty("Green")!.GetValue(eraArgs[5])!);
            TestAssert.Equal((byte)era1B[s], (byte)rgbType!.GetProperty("Blue")!.GetValue(eraArgs[5])!);
        }
        // Era 2 (N=170).
        int[] era2R = [45, 118, 94, 77, 43, 153];
        int[] era2G = [77, 46, 114, 59, 105, 91];
        int[] era2B = [116, 44, 52, 98, 121, 40];
        int era2N = 170;
        for (int s = 0; s < eraIdx.Length; s++)
        {
            object?[] eraArgs = [palette, null, colorMap, eraIdx[s], era2N, null];
            TestAssert.Equal(true, (bool)method!.Invoke(null, eraArgs)!);
            TestAssert.Equal((byte)era2R[s], (byte)rgbType!.GetProperty("Red")!.GetValue(eraArgs[5])!);
            TestAssert.Equal((byte)era2G[s], (byte)rgbType!.GetProperty("Green")!.GetValue(eraArgs[5])!);
            TestAssert.Equal((byte)era2B[s], (byte)rgbType!.GetProperty("Blue")!.GetValue(eraArgs[5])!);
        }
        // Era 3 (N=204).
        int[] era3R = [45, 116, 92, 75, 42, 150];
        int[] era3G = [76, 45, 113, 58, 103, 89];
        int[] era3B = [114, 43, 51, 97, 120, 39];
        int era3N = 204;
        for (int s = 0; s < eraIdx.Length; s++)
        {
            object?[] eraArgs = [palette, null, colorMap, eraIdx[s], era3N, null];
            TestAssert.Equal(true, (bool)method!.Invoke(null, eraArgs)!);
            TestAssert.Equal((byte)era3R[s], (byte)rgbType!.GetProperty("Red")!.GetValue(eraArgs[5])!);
            TestAssert.Equal((byte)era3G[s], (byte)rgbType!.GetProperty("Green")!.GetValue(eraArgs[5])!);
            TestAssert.Equal((byte)era3B[s], (byte)rgbType!.GetProperty("Blue")!.GetValue(eraArgs[5])!);
        }
        // Spot cells where the engine reproduces the Office render bit-exactly
        // (every literal below is the Office-observed fill; engine output matched before pinning).
        int[] spotCat = [0, 53, 52, 8, 64, 36, 8, 47, 28, 23, 22, 33, 41, 77, 45, 104, 20, 66, 31, 2, 100, 48, 130, 61, 131, 62, 36, 0, 112, 99, 12, 26, 42, 57, 2, 61, 7, 66, 19, 70, 9, 68, 4, 71, 2, 69, 50, 39];
        int[] spotCount = [96, 98, 101, 106, 108, 110, 112, 113, 116, 119, 122, 125, 128, 131, 133, 134, 136, 137, 139, 141, 142, 144, 145, 147, 148, 150, 152, 154, 155, 156, 158, 159, 160, 161, 163, 164, 166, 167, 169, 170, 172, 173, 175, 176, 178, 179, 192, 198];
        int[] spotR = [47, 247, 75, 106, 103, 70, 106, 231, 59, 191, 55, 107, 214, 248, 113, 198, 116, 79, 156, 95, 150, 70, 190, 182, 251, 145, 64, 46, 156, 156, 53, 118, 66, 115, 94, 177, 125, 74, 138, 70, 82, 145, 43, 229, 94, 118, 129, 99];
        int[] spotG = [80, 150, 172, 129, 180, 116, 129, 140, 139, 115, 131, 83, 129, 163, 88, 214, 140, 129, 64, 116, 197, 115, 217, 76, 209, 176, 106, 78, 200, 139, 89, 143, 109, 90, 114, 73, 49, 121, 55, 161, 63, 175, 105, 139, 114, 92, 156, 77];
        int[] spotB = [120, 70, 198, 59, 203, 170, 59, 65, 161, 52, 151, 136, 59, 107, 143, 172, 65, 189, 61, 53, 214, 169, 228, 73, 189, 83, 157, 117, 216, 180, 132, 66, 161, 146, 52, 70, 47, 178, 53, 185, 105, 83, 121, 64, 52, 150, 73, 126];
        for (int s = 0; s < spotCat.Length; s++)
        {
            object?[] spotArgs = [palette, null, colorMap, spotCat[s], spotCount[s], null];
            TestAssert.Equal(true, (bool)method!.Invoke(null, spotArgs)!);
            TestAssert.Equal((byte)spotR[s], (byte)rgbType!.GetProperty("Red")!.GetValue(spotArgs[5])!);
            TestAssert.Equal((byte)spotG[s], (byte)rgbType!.GetProperty("Green")!.GetValue(spotArgs[5])!);
            TestAssert.Equal((byte)spotB[s], (byte)rgbType!.GetProperty("Blue")!.GetValue(spotArgs[5])!);
        }
    }
}
