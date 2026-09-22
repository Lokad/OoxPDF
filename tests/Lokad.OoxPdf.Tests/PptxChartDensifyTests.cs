using System.Reflection;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxChartDensifyTests
{
    public static void PresenceChecksMatchDensifiedResults()
    {
        Type vectorType = NumberVectorType();
        Type pointType = NumberPointType();
        object empty = BuildNumberVector(vectorType, pointType, [], null);
        TestAssert.True(!InvokeHasAnyValue(vectorType, empty), "Empty vectors must report no value.");
        TestAssert.True(!InvokeHasAnyDenseSlot(vectorType, empty), "Empty vectors must report no dense slot.");
        TestAssert.Equal(0, InvokeDenseLength(vectorType, empty));

        object sparse = BuildNumberVector(vectorType, pointType, [(0, 5d), (2, null)], null);
        TestAssert.True(InvokeHasAnyValue(vectorType, sparse), "A sparse valued point must report a value.");
        TestAssert.True(InvokeHasAnyDenseSlot(vectorType, sparse), "A sparse vector must report dense slots.");
        TestAssert.Equal(3, InvokeDenseLength(vectorType, sparse));

        object allNull = BuildNumberVector(vectorType, pointType, [(0, null), (1, null)], 5);
        TestAssert.True(!InvokeHasAnyValue(vectorType, allNull), "All-null values must report no value.");
        TestAssert.True(InvokeHasAnyDenseSlot(vectorType, allNull), "A declared count still materializes slots.");
        TestAssert.Equal(5, InvokeDenseLength(vectorType, allNull));
    }

    public static void TextPresenceChecksMatchDensifiedResults()
    {
        Type vectorType = typeof(PptxRenderer).GetNestedType("ChartIndexedTextVector", BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("Expected text vector type.");
        Type pointType = typeof(PptxRenderer).GetNestedType("ChartIndexedTextPoint", BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("Expected text point type.");
        ConstructorInfo pointCtor = pointType.GetConstructors().Single();
        object textPoint = pointCtor.Invoke([0, Activator.CreateInstance(pointType.GetConstructors().Single().GetParameters()[1].ParameterType), "a", true, null]);
        Array points = Array.CreateInstance(pointType, 1);
        points.SetValue(textPoint, 0);
        object vector = BuildVector(vectorType, points, null);
        TestAssert.True(InvokeBool(vectorType, vector, "HasAnyText"), "A text point must report text.");
        TestAssert.True(InvokeBool(vectorType, vector, "HasAnyDenseSlot"), "A text vector must report dense slots.");
    }

    public static void PresenceChecksPreserveDenseCaps()
    {
        Type vectorType = NumberVectorType();
        Type pointType = NumberPointType();
        object huge = BuildNumberVector(vectorType, pointType, [], 200000);
        TestAssert.Throws<OoxPdfLimitExceededException>(() => InvokeHasAnyValueThrowing(vectorType, huge));
        TestAssert.Throws<OoxPdfLimitExceededException>(() => InvokeHasAnyDenseSlotThrowing(vectorType, huge));
    }

    public static void CountRenderableSeriesSkipsDensifyAllocation()
    {
        Type vectorType = NumberVectorType();
        Type pointType = NumberPointType();
        Array series = Array.CreateInstance(vectorType, 3);
        series.SetValue(BuildNumberVector(vectorType, pointType, [(0, 1d)], null), 0);
        series.SetValue(BuildNumberVector(vectorType, pointType, [], null), 1);
        series.SetValue(BuildNumberVector(vectorType, pointType, [], 10000), 2);
        MethodInfo count = typeof(PptxRenderer).GetMethod("CountRenderableSeries", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("Expected count helper.");
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        long before = GC.GetAllocatedBytesForCurrentThread();
        int renderable;
        try
        {
            renderable = (int)count.Invoke(null, [series])!;
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }

        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        TestAssert.Equal(1, renderable);
        TestAssert.True(allocated <= 262144L, $"Counting renderable series must not densify 10,000-slot vectors, allocated {allocated} bytes.");
    }

    public static void SubsetCacheChargesOncePerCodepointSet()
    {
        byte[] bytes = TestFontBuilder.CreateTestFont();
        OpenTypeFont font = OpenTypeFont.Load(bytes);
        var resolution = new FontFaceResolution(
            font.FamilyName,
            font.FamilyName,
            new FontStyleKey(),
            new MemoryFontProgramSource("memory:r15-subset", bytes),
            IsFallback: false);
        var resolver = new PresentationFontResolver(new CannedFontResolver(resolution));
        int[] codePoints = Enumerable.Range(65, 200).ToArray();
        using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxFontWorkPerConversion = 100 }))
        {
            PdfEmbeddedFont first = resolver.GetOrCreateSubset(resolution, font, codePoints, CancellationToken.None);
            PdfEmbeddedFont second = resolver.GetOrCreateSubset(resolution, font, codePoints, CancellationToken.None);
            TestAssert.True(ReferenceEquals(first, second), "Identical codepoint sets must share one subset.");
            TestAssert.Equal(1, scope.Budget.FontWork);
        }
    }

    public static void VisibleTextCheckMatchesDensifiedLabels()
    {
        // R15: the sparse visible-text probe must agree with the dense scan,
        // including whitespace-only, empty, declared-empty, and out-of-range cases.
        Type vectorType = typeof(PptxRenderer).GetNestedType("ChartIndexedTextVector", BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("Expected text vector type.");
        object real = BuildTextVector(vectorType, [(0, "a", true)], null);
        TestAssert.True(InvokeBool(vectorType, real, "HasAnyVisibleText"), "A real label must report visible text.");
        object whitespace = BuildTextVector(vectorType, [(0, "   ", true)], null);
        TestAssert.True(!InvokeBool(vectorType, whitespace, "HasAnyVisibleText"), "Whitespace-only labels must report no visible text.");
        object empty = BuildTextVector(vectorType, [], null);
        TestAssert.True(!InvokeBool(vectorType, empty, "HasAnyVisibleText"), "Empty vectors must report no visible text.");
        object blank = BuildTextVector(vectorType, [(0, "", false)], null);
        TestAssert.True(!InvokeBool(vectorType, blank, "HasAnyVisibleText"), "Textless points must report no visible text.");
        object negative = BuildTextVector(vectorType, [(-1, "a", true)], null);
        TestAssert.True(!InvokeBool(vectorType, negative, "HasAnyVisibleText"), "Negative indexes must report no visible text.");
        object declaredExtends = BuildTextVector(vectorType, [(5, "a", true)], 3);
        TestAssert.True(InvokeBool(vectorType, declaredExtends, "HasAnyVisibleText"), "Declared counts extend (never truncate) the inferred range.");
    }

    public static void RadarBuildSkipsValuelessVectorsWithoutCharging()
    {
        // R15: valueless series (even with declared slots) filter before densifying,
        // so only valued vectors materialize and charge.
        Type vectorType = NumberVectorType();
        Type pointType = NumberPointType();
        Array series = Array.CreateInstance(vectorType, 2);
        series.SetValue(BuildNumberVector(vectorType, pointType, [(0, null), (1, null)], 5), 0);
        series.SetValue(BuildNumberVector(vectorType, pointType, [(0, 1d)], null), 1);
        MethodInfo build = typeof(PptxRenderer).GetMethod("BuildRadarSeries", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("Expected radar builder.");
        using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxChartRangeCellsPerConversion = 5 }))
        {
            System.Collections.IList built;
            try
            {
                built = (System.Collections.IList)build.Invoke(null, [series])!;
            }
            catch (TargetInvocationException ex)
            {
                throw ex.InnerException ?? ex;
            }

            TestAssert.Equal(1, built.Count);
            TestAssert.Equal(1, scope.Budget.ChartRangeCells);
        }
    }

    private static Type NumberVectorType()
    {
        return typeof(PptxRenderer).GetNestedType("ChartIndexedNumberVector", BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("Expected number vector type.");
    }

    private static Type NumberPointType()
    {
        return typeof(PptxRenderer).GetNestedType("ChartIndexedNumberPoint", BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("Expected number point type.");
    }

    private static object BuildNumberVector(Type vectorType, Type pointType, (int Index, double? Value)[] entries, int? pointCount)
    {
        ConstructorInfo pointCtor = pointType.GetConstructors().Single();
        ParameterInfo[] pointParams = pointCtor.GetParameters();
        Array points = Array.CreateInstance(pointType, entries.Length);
        for (int i = 0; i < entries.Length; i++)
        {
            object?[] args = pointParams.Select(p => p.ParameterType.IsValueType ? Activator.CreateInstance(p.ParameterType) : null).ToArray();
            args[0] = entries[i].Index;
            args[2] = entries[i].Value;
            points.SetValue(pointCtor.Invoke(args), i);
        }

        return BuildVector(vectorType, points, pointCount);
    }

    private static object BuildVector(Type vectorType, Array points, int? pointCount)
    {
        ConstructorInfo ctor = vectorType.GetConstructors().Single();
        object?[] args = ctor.GetParameters().Select(p => p.ParameterType.IsValueType ? Activator.CreateInstance(p.ParameterType) : null).ToArray();
        args[0] = points;
        args[1] = pointCount;
        return ctor.Invoke(args) ?? throw new InvalidOperationException("Expected vector instance.");
    }

    private static object BuildTextVector(Type vectorType, (int Index, string Text, bool HasText)[] entries, int? pointCount)
    {
        Type pointType = typeof(PptxRenderer).GetNestedType("ChartIndexedTextPoint", BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("Expected text point type.");
        ConstructorInfo pointCtor = pointType.GetConstructors().Single();
        ParameterInfo[] pointParams = pointCtor.GetParameters();
        Array points = Array.CreateInstance(pointType, entries.Length);
        for (int i = 0; i < entries.Length; i++)
        {
            object?[] args = pointParams.Select(p => p.ParameterType.IsValueType ? Activator.CreateInstance(p.ParameterType) : null).ToArray();
            args[0] = entries[i].Index;
            args[2] = entries[i].Text;
            args[3] = entries[i].HasText;
            points.SetValue(pointCtor.Invoke(args), i);
        }

        return BuildVector(vectorType, points, pointCount);
    }

    private static bool InvokeHasAnyValue(Type vectorType, object vector)
    {
        return InvokeBool(vectorType, vector, "HasAnyValue");
    }

    private static bool InvokeHasAnyDenseSlot(Type vectorType, object vector)
    {
        return InvokeBool(vectorType, vector, "HasAnyDenseSlot");
    }

    private static bool InvokeBool(Type vectorType, object vector, string method)
    {
        try
        {
            return (bool)vectorType.GetMethod(method)!.Invoke(vector, null)!;
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
    }

    private static void InvokeHasAnyValueThrowing(Type vectorType, object vector)
    {
        InvokeBool(vectorType, vector, "HasAnyValue");
    }

    private static void InvokeHasAnyDenseSlotThrowing(Type vectorType, object vector)
    {
        InvokeBool(vectorType, vector, "HasAnyDenseSlot");
    }

    private static int InvokeDenseLength(Type vectorType, object vector)
    {
        try
        {
            return ((System.Collections.IList)vectorType.GetMethod("DensePoints")!.Invoke(vector, null)!).Count;
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
    }

    private sealed class CannedFontResolver(FontFaceResolution resolution) : IFontResolver
    {
        public FontFaceResolution Resolve(FontRequest request)
        {
            return resolution;
        }
    }
}
