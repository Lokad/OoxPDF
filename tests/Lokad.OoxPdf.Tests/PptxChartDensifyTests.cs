using System.Linq.Expressions;
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

    public static void SparseMaxDenseCountMatchesDensifiedMax()
    {
        // R15: the sparse max-dense-count probe must agree with densify-then-max,
        // including sparse, declared, empty, negative-index, and over-cap cases.
        Type vectorType = NumberVectorType();
        Type pointType = NumberPointType();
        Array series = Array.CreateInstance(vectorType, 5);
        series.SetValue(BuildNumberVector(vectorType, pointType, [], null), 0);
        series.SetValue(BuildNumberVector(vectorType, pointType, [(0, 5d), (2, null)], null), 1);
        series.SetValue(BuildNumberVector(vectorType, pointType, [(0, null), (1, null)], 5), 2);
        series.SetValue(BuildNumberVector(vectorType, pointType, [], 4), 3);
        series.SetValue(BuildNumberVector(vectorType, pointType, [(-1, 1d)], null), 4);
        TestAssert.Equal(5, InvokeMaxDensePointCount(series));
        TestAssert.Equal(InvokeDensifiedMaxCount(series), InvokeMaxDensePointCount(series));

        Array empty = Array.CreateInstance(vectorType, 0);
        TestAssert.Equal(0, InvokeMaxDensePointCount(empty));

        object huge = BuildNumberVector(vectorType, pointType, [], 200000);
        Array hugeSeries = Array.CreateInstance(vectorType, 1);
        hugeSeries.SetValue(huge, 0);
        TestAssert.Throws<OoxPdfLimitExceededException>(() => InvokeMaxDensePointCount(hugeSeries));
        TestAssert.Throws<OoxPdfLimitExceededException>(() => InvokeDensifiedMaxCount(hugeSeries));
    }

    public static void MaxDenseCountSkipsDensifyAllocation()
    {
        Type vectorType = NumberVectorType();
        Type pointType = NumberPointType();
        Array series = Array.CreateInstance(vectorType, 3);
        series.SetValue(BuildNumberVector(vectorType, pointType, [(0, 1d)], null), 0);
        series.SetValue(BuildNumberVector(vectorType, pointType, [], null), 1);
        series.SetValue(BuildNumberVector(vectorType, pointType, [], 10000), 2);
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        long before = GC.GetAllocatedBytesForCurrentThread();
        int max = InvokeMaxDensePointCount(series);
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        TestAssert.Equal(10000, max);
        TestAssert.True(allocated <= 262144L, "Counting dense slots must not densify 10,000-slot vectors, allocated " + allocated + " bytes.");
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

    public static void CategoryLabelMemoSharesDenseResultsWithinFrame()
    {
        // R15: one dense array per (source, visibility) per frame; clearing the
        // frame memo recomputes on next access.
        Type workbookType = typeof(PptxRenderer).GetNestedType("ChartWorkbookData", BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("Expected workbook type.");
        var sheets = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
        {
            ["Sheet1"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["A1"] = "5" },
        };
        object workbook = Activator.CreateInstance(workbookType, [sheets])
            ?? throw new InvalidOperationException("Expected workbook instance.");
        MethodInfo memo = workbookType.GetMethod("GetOrAddCategoryLabels", BindingFlags.NonPublic | BindingFlags.Instance)
            ?? throw new InvalidOperationException("Expected label memo.");
        Type factoryType = memo.GetParameters()[3].ParameterType;

        Type vectorType = typeof(PptxRenderer).GetNestedType("ChartIndexedTextVector", BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("Expected text vector type.");
        object vector = BuildTextVector(vectorType, [(0, "a", true)], null);
        MethodInfo dense = vectorType.GetMethod("DensePoints")
            ?? throw new InvalidOperationException("Expected densify method.");
        object denseA = dense.Invoke(vector, null) ?? throw new InvalidOperationException("Expected dense array.");
        object denseB = dense.Invoke(vector, null) ?? throw new InvalidOperationException("Expected dense array.");
        TestAssert.True(!ReferenceEquals(denseA, denseB), "DensePoints must allocate per call (precondition).");

        object first = InvokeMemo(memo, workbook, "k", null, false, ConstantFactory(factoryType, denseA));
        object second = InvokeMemo(memo, workbook, "k", null, false, ConstantFactory(factoryType, denseB));
        TestAssert.True(ReferenceEquals(first, second), "Same key must share one dense array.");
        TestAssert.True(ReferenceEquals(first, denseA), "First computation must win.");

        try
        {
            workbookType.GetMethod("ClearRangeMemo", BindingFlags.NonPublic | BindingFlags.Instance)?.Invoke(workbook, null);
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }

        object third = InvokeMemo(memo, workbook, "k", null, false, ConstantFactory(factoryType, denseB));
        TestAssert.True(ReferenceEquals(third, denseB), "Cleared memo must recompute on next access.");
    }

    public static void SeriesNameMemoSharesResultsWithinFrame()
    {
        // R15: one series-name list per (source, element) per frame; clearing the
        // frame memo recomputes on next access.
        Type workbookType = typeof(PptxRenderer).GetNestedType("ChartWorkbookData", BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("Expected workbook type.");
        var sheets = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
        {
            ["Sheet1"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["A1"] = "5" },
        };
        object workbook = Activator.CreateInstance(workbookType, [sheets])
            ?? throw new InvalidOperationException("Expected workbook instance.");
        MethodInfo memo = workbookType.GetMethod("GetOrAddSeriesNames", BindingFlags.NonPublic | BindingFlags.Instance)
            ?? throw new InvalidOperationException("Expected series-name memo.");
        IReadOnlyList<PptxRenderer.ChartSeriesNameRecord> arrayA = [];
        IReadOnlyList<PptxRenderer.ChartSeriesNameRecord> arrayB = [];
        Func<IReadOnlyList<PptxRenderer.ChartSeriesNameRecord>> factoryA = () => arrayA;
        Func<IReadOnlyList<PptxRenderer.ChartSeriesNameRecord>> factoryB = () => arrayB;
        object first = InvokeUntyped(memo, workbook, "k", null, factoryA);
        object second = InvokeUntyped(memo, workbook, "k", null, factoryB);
        TestAssert.True(ReferenceEquals(first, second), "Same key must share one name list.");
        TestAssert.True(ReferenceEquals(first, arrayA), "First computation must win.");
        try
        {
            workbookType.GetMethod("ClearRangeMemo", BindingFlags.NonPublic | BindingFlags.Instance)?.Invoke(workbook, null);
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }

        object third = InvokeUntyped(memo, workbook, "k", null, factoryB);
        TestAssert.True(ReferenceEquals(third, arrayB), "Cleared memo must recompute on next access.");
    }

    private static object InvokeUntyped(MethodInfo memo, object workbook, object? first, object? second, object factory)
    {
        try
        {
            return memo.Invoke(workbook, [first, second, factory])
                ?? throw new InvalidOperationException("Expected shared list.");
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
    }

    public static void SeriesVectorsMemoSharesResultsWithinFrame()
    {
        // R15: one series-vector list per (source, element, visibility) per frame;
        // clearing the frame memo recomputes on next access.
        Type workbookType = typeof(PptxRenderer).GetNestedType("ChartWorkbookData", BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("Expected workbook type.");
        var sheets = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
        {
            ["Sheet1"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["A1"] = "5" },
        };
        object workbook = Activator.CreateInstance(workbookType, [sheets])
            ?? throw new InvalidOperationException("Expected workbook instance.");
        MethodInfo memo = workbookType.GetMethod("GetOrAddSeriesVectors", BindingFlags.NonPublic | BindingFlags.Instance)
            ?? throw new InvalidOperationException("Expected series-vectors memo.");
        IReadOnlyList<PptxRenderer.ChartIndexedNumberVector> arrayA = [];
        IReadOnlyList<PptxRenderer.ChartIndexedNumberVector> arrayB = [];
        Func<IReadOnlyList<PptxRenderer.ChartIndexedNumberVector>> factoryA = () => arrayA;
        Func<IReadOnlyList<PptxRenderer.ChartIndexedNumberVector>> factoryB = () => arrayB;
        object first = InvokeUntyped4(memo, workbook, "k", null, false, factoryA);
        object second = InvokeUntyped4(memo, workbook, "k", null, false, factoryB);
        TestAssert.True(ReferenceEquals(first, second), "Same key must share one vector list.");
        TestAssert.True(ReferenceEquals(first, arrayA), "First computation must win.");
        try
        {
            workbookType.GetMethod("ClearRangeMemo", BindingFlags.NonPublic | BindingFlags.Instance)?.Invoke(workbook, null);
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }

        object third = InvokeUntyped4(memo, workbook, "k", null, false, factoryB);
        TestAssert.True(ReferenceEquals(third, arrayB), "Cleared memo must recompute on next access.");
    }

    public static void ValueExtentsMemoSharesResultsWithinFrame()
    {
        // R15: one extent result per (source, element, grouping, visibility) per
        // frame; clearing the frame memo recomputes on next access.
        Type workbookType = typeof(PptxRenderer).GetNestedType("ChartWorkbookData", BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("Expected workbook type.");
        var sheets = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
        {
            ["Sheet1"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["A1"] = "5" },
        };
        object workbook = Activator.CreateInstance(workbookType, [sheets])
            ?? throw new InvalidOperationException("Expected workbook instance.");
        MethodInfo memo = workbookType.GetMethod("GetOrAddValueExtents", BindingFlags.NonPublic | BindingFlags.Instance)
            ?? throw new InvalidOperationException("Expected extent memo.");
        PptxRenderer.ChartValueExtents extentsA = new(0d, 10d);
        PptxRenderer.ChartValueExtents extentsB = new(0d, 99d);
        Func<PptxRenderer.ChartValueExtents> factoryA = () => extentsA;
        Func<PptxRenderer.ChartValueExtents> factoryB = () => extentsB;
        object first = InvokeUntyped5(memo, workbook, "k", null, PptxSceneChartGrouping.Clustered, false, factoryA);
        TestAssert.Equal(extentsA, (PptxRenderer.ChartValueExtents)first);
        object second = InvokeUntyped5(memo, workbook, "k", null, PptxSceneChartGrouping.Clustered, false, factoryB);
        TestAssert.Equal(extentsA, (PptxRenderer.ChartValueExtents)second);
        object other = InvokeUntyped5(memo, workbook, "k", null, PptxSceneChartGrouping.Stacked, false, factoryB);
        TestAssert.Equal(extentsB, (PptxRenderer.ChartValueExtents)other);
        try
        {
            workbookType.GetMethod("ClearRangeMemo", BindingFlags.NonPublic | BindingFlags.Instance)?.Invoke(workbook, null);
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }

        object third = InvokeUntyped5(memo, workbook, "k", null, PptxSceneChartGrouping.Clustered, false, factoryB);
        TestAssert.Equal(extentsB, (PptxRenderer.ChartValueExtents)third);
    }

    public static void LineValueExtentsMemoSharesResultsWithinFrame()
    {
        // R15: one line extent result per (source, element, flags, visibility) per
        // frame; clearing the frame memo recomputes on next access.
        Type workbookType = typeof(PptxRenderer).GetNestedType("ChartWorkbookData", BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("Expected workbook type.");
        var sheets = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
        {
            ["Sheet1"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["A1"] = "5" },
        };
        object workbook = Activator.CreateInstance(workbookType, [sheets])
            ?? throw new InvalidOperationException("Expected workbook instance.");
        MethodInfo memo = workbookType.GetMethod("GetOrAddLineValueExtents", BindingFlags.NonPublic | BindingFlags.Instance)
            ?? throw new InvalidOperationException("Expected line extent memo.");
        PptxRenderer.ChartValueExtents extentsA = new(0d, 10d);
        PptxRenderer.ChartValueExtents extentsB = new(0d, 99d);
        Func<PptxRenderer.ChartValueExtents> factoryA = () => extentsA;
        Func<PptxRenderer.ChartValueExtents> factoryB = () => extentsB;
        object first = InvokeUntyped6(memo, workbook, "k", null, false, false, false, factoryA);
        TestAssert.Equal(extentsA, (PptxRenderer.ChartValueExtents)first);
        object second = InvokeUntyped6(memo, workbook, "k", null, false, false, false, factoryB);
        TestAssert.Equal(extentsA, (PptxRenderer.ChartValueExtents)second);
        object other = InvokeUntyped6(memo, workbook, "k", null, true, false, false, factoryB);
        TestAssert.Equal(extentsB, (PptxRenderer.ChartValueExtents)other);
        try
        {
            workbookType.GetMethod("ClearRangeMemo", BindingFlags.NonPublic | BindingFlags.Instance)?.Invoke(workbook, null);
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }

        object third = InvokeUntyped6(memo, workbook, "k", null, false, false, false, factoryB);
        TestAssert.Equal(extentsB, (PptxRenderer.ChartValueExtents)third);
    }

    private static object InvokeUntyped4(MethodInfo memo, object workbook, object? first, object? second, bool visibleOnly, object factory)
    {
        try
        {
            return memo.Invoke(workbook, [first, second, visibleOnly, factory])
                ?? throw new InvalidOperationException("Expected shared list.");
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
    }

    private static object InvokeUntyped5(MethodInfo memo, object workbook, object? source, object? chartElement, object grouping, bool visibleOnly, object factory)
    {
        try
        {
            return memo.Invoke(workbook, [source, chartElement, grouping, visibleOnly, factory])
                ?? throw new InvalidOperationException("Expected shared extents.");
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
    }

    private static object InvokeUntyped6(MethodInfo memo, object workbook, object? source, object? chartElement, object stacked, object percentStacked, bool visibleOnly, object factory)
    {
        try
        {
            return memo.Invoke(workbook, [source, chartElement, stacked, percentStacked, visibleOnly, factory])
                ?? throw new InvalidOperationException("Expected shared line extents.");
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
    }

    private static object InvokeMemo(MethodInfo memo, object workbook, object? source, object? chartElement, bool visibleOnly, object factory)
    {
        try
        {
            return memo.Invoke(workbook, [source, chartElement, visibleOnly, factory])
                ?? throw new InvalidOperationException("Expected shared dense array.");
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
    }

    private static object ConstantFactory(Type factoryType, object value)
    {
        Type returnType = factoryType.GetGenericArguments()[0];
        return Expression.Lambda(factoryType, Expression.Constant(value, returnType)).Compile();
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

    private static int InvokeMaxDensePointCount(Array series)
    {
        MethodInfo count = typeof(PptxRenderer).GetMethod("MaxDensePointCount", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("Expected max-dense-count helper.");
        try
        {
            return (int)count.Invoke(null, [series])!;
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
    }

    private static int InvokeDensifiedMaxCount(Array series)
    {
        MethodInfo densify = typeof(PptxRenderer).GetMethod("DensifyChartPointSeries", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("Expected densify helper.");
        try
        {
            var dense = (System.Collections.IList)densify.Invoke(null, [series])!;
            int max = 0;
            foreach (object? values in dense)
            {
                max = Math.Max(max, ((System.Collections.IList)values!).Count);
            }

            return max;
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
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

    public static void CategoryLabelLookupsScaleLinearly()
    {
        // RV14: N sequential label lookups must not rescan points per lookup.
        const int count = 2000;
        Type vectorType = typeof(PptxRenderer).GetNestedType("ChartIndexedTextVector", BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("Expected text vector type.");
        var inner = new List<PptxRenderer.ChartIndexedTextPoint>(count);
        for (int i = 0; i < count; i++)
        {
            inner.Add(new PptxRenderer.ChartIndexedTextPoint(i, default, "label" + i, true, default));
        }

        var counting = new CountingTextPoints(inner);
        object vector = BuildTextVectorWithPoints(vectorType, counting, null);
        MethodInfo lookup = typeof(PptxRenderer).GetMethod("GetIndexedCategoryLabel", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("Expected category label lookup.");
        for (int i = 0; i < count; i++)
        {
            try
            {
                lookup.Invoke(null, [vector, i]);
            }
            catch (TargetInvocationException ex)
            {
                throw ex.InnerException ?? ex;
            }
        }

        TestAssert.True(counting.Visits <= 4L * count, $"Sequential label lookups must scan linearly, visited {counting.Visits} points for {count} labels.");
    }

    public static void CategoryLabelLookupKeepsFirstMatchPrecedence()
    {
        // RV14: duplicates keep the first text; sparse, blank, and out-of-range
        // indices yield empty; out-of-order points resolve correctly.
        Type vectorType = typeof(PptxRenderer).GetNestedType("ChartIndexedTextVector", BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("Expected text vector type.");
        object vector = BuildTextVector(vectorType, [(2, "b", true), (0, "a", true), (2, "B2", true), (1, "", true), (3, "c", false)], null);
        MethodInfo lookup = typeof(PptxRenderer).GetMethod("GetIndexedCategoryLabel", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("Expected category label lookup.");
        TestAssert.Equal("a", InvokeCategoryLabel(lookup, vector, 0));
        TestAssert.Equal("", InvokeCategoryLabel(lookup, vector, 1));
        TestAssert.Equal("b", InvokeCategoryLabel(lookup, vector, 2));
        TestAssert.Equal("", InvokeCategoryLabel(lookup, vector, 3));
        TestAssert.Equal("", InvokeCategoryLabel(lookup, vector, 5));
    }

    public static void ValueOnlyLabelsSkipCategoryLookup()
    {
        // RV14: value-only labels must not touch category points at all.
        Type vectorType = typeof(PptxRenderer).GetNestedType("ChartIndexedTextVector", BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("Expected text vector type.");
        object vector = BuildTextVectorWithPoints(vectorType, new ThrowingTextPoints(), null);
        MethodInfo format = typeof(PptxRenderer).GetMethod("FormatCartesianDataLabel", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("Expected cartesian label formatter.");
        Type optionsType = typeof(PptxRenderer).GetNestedType("ChartDataLabelOptions", BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("Expected label options type.");
        object options = Activator.CreateInstance(optionsType)!;
        Type namesType = typeof(PptxRenderer).GetNestedType("ChartSeriesNameRecord", BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("Expected series name type.");
        Array names = Array.CreateInstance(namesType, 0);
        string label;
        try
        {
            label = (string)format.Invoke(null, [1.5, 0, 0, null, null, null, options, vector, names])!;
        }
        catch (TargetInvocationException ex)
        {
            throw new InvalidOperationException("Value-only labels must not enumerate category points.", ex.InnerException ?? ex);
        }

        TestAssert.Equal("", label);
    }

    private static string InvokeCategoryLabel(MethodInfo lookup, object vector, int index)
    {
        try
        {
            return (string)lookup.Invoke(null, [vector, index])!;
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
    }

    private static object BuildTextVectorWithPoints(Type vectorType, object points, int? pointCount)
    {
        ConstructorInfo ctor = vectorType.GetConstructors().Single();
        object?[] args = ctor.GetParameters().Select(p => p.ParameterType.IsValueType ? Activator.CreateInstance(p.ParameterType) : null).ToArray();
        args[0] = points;
        args[1] = pointCount;
        try
        {
            return ctor.Invoke(args) ?? throw new InvalidOperationException("Expected vector instance.");
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
    }

    private sealed class CountingTextPoints(IReadOnlyList<PptxRenderer.ChartIndexedTextPoint> inner) : IReadOnlyList<PptxRenderer.ChartIndexedTextPoint>
    {
        public long Visits;

        public PptxRenderer.ChartIndexedTextPoint this[int index]
        {
            get { Visits++; return inner[index]; }
        }

        public int Count
        {
            get { Visits++; return inner.Count; }
        }

        public IEnumerator<PptxRenderer.ChartIndexedTextPoint> GetEnumerator()
        {
            foreach (PptxRenderer.ChartIndexedTextPoint point in inner)
            {
                Visits++;
                yield return point;
            }
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }

    private sealed class ThrowingTextPoints : IReadOnlyList<PptxRenderer.ChartIndexedTextPoint>
    {
        public PptxRenderer.ChartIndexedTextPoint this[int index] => throw new InvalidOperationException("Category points must not be enumerated.");

        public int Count => throw new InvalidOperationException("Category points must not be enumerated.");

        public IEnumerator<PptxRenderer.ChartIndexedTextPoint> GetEnumerator() => throw new InvalidOperationException("Category points must not be enumerated.");

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }

    private sealed class CannedFontResolver(FontFaceResolution resolution) : IFontResolver
    {
        public FontFaceResolution Resolve(FontRequest request)
        {
            return resolution;
        }
    }
}
