using System.IO.Compression;
using System.Reflection;
using System.Xml.Linq;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Ooxml;
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
        TestAssert.Equal(0, InvokeDenseLength(vectorType, empty));

        object sparse = BuildNumberVector(vectorType, pointType, [(0, 5d), (2, null)], null);
        TestAssert.True(InvokeHasAnyValue(vectorType, sparse), "A sparse valued point must report a value.");
        TestAssert.Equal(3, InvokeDenseLength(vectorType, sparse));

        object allNull = BuildNumberVector(vectorType, pointType, [(0, null), (1, null)], 5);
        TestAssert.True(!InvokeHasAnyValue(vectorType, allNull), "All-null values must report no value.");
        TestAssert.Equal(5, InvokeDenseLength(vectorType, allNull));
    }

    public static void TextDenseSlotsRetainDeclaredAndInferredCounts()
    {
        Type vectorType = typeof(PptxRenderer).GetNestedType("ChartIndexedTextVector", BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("Expected text vector type.");
        MethodInfo dense = vectorType.GetMethod("DensePoints") ?? throw new InvalidOperationException("Expected dense text slots.");
        Array Slots((int Index, string Text, bool HasText)[] points, int? count)
        {
            try { return (Array)dense.Invoke(BuildTextVector(vectorType, points, count), null)!; }
            catch (TargetInvocationException ex) { throw ex.InnerException ?? ex; }
        }

        TestAssert.Equal(0, Slots([(-1, "a", true)], null).Length);
        TestAssert.Equal(4, Slots([], 4).Length);
        var sparse = Slots([(5, "a", true)], 3);
        TestAssert.Equal(6, sparse.Length);
        TestAssert.True(sparse.GetValue(0) is null && sparse.GetValue(5) is not null,
            "A smaller declared count must not truncate a later authored slot or fill the gaps before it.");
        TestAssert.Throws<OoxPdfLimitExceededException>(() => Slots([], -1));
        TestAssert.Throws<OoxPdfLimitExceededException>(() => Slots([], PptxRenderer.MaxChartDensePoints + 1));
        TestAssert.Throws<OoxPdfLimitExceededException>(() => Slots([(PptxRenderer.MaxChartDensePoints, "a", true)], null));
    }

    public static void SparseTextSlotsPreservePointsWithoutDuplicatingEmptyProvenance()
    {
        Type vectorType = typeof(PptxRenderer).GetNestedType("ChartIndexedTextVector", BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("Expected text vector type.");
        var denseMethod = vectorType.GetMethod("DensePoints") ?? throw new InvalidOperationException("Expected dense text slots.");
        denseMethod.Invoke(BuildTextVector(vectorType, [(0, "warm", true)], 2), null);
        object source = BuildTextVector(vectorType, [(0, "first", true), (99_999, "last", true), (0, "replacement", true)], 100_000);
        long before = GC.GetAllocatedBytesForCurrentThread();
        var slots = (Array)(denseMethod.Invoke(source, null) ?? throw new InvalidOperationException("Expected dense text slots."));
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        TestAssert.Equal(100_000, slots.Length);
        TestAssert.True(slots.GetValue(1) is null, "Retain the missing source slot.");
        string Text(int index) => (string)(slots.GetValue(index)!.GetType().GetProperty("Text")!.GetValue(slots.GetValue(index))!);
        TestAssert.Equal("replacement", Text(0));
        TestAssert.Equal("last", Text(99_999));
        TestAssert.True(allocated < 32L * slots.Length,
            $"Empty slots must not duplicate the large workbook provenance record: {allocated} bytes for {slots.Length} slots.");
    }

    public static void PresenceChecksPreserveDenseCaps()
    {
        Type vectorType = NumberVectorType();
        Type pointType = NumberPointType();
        object huge = BuildNumberVector(vectorType, pointType, [], 200000);
        TestAssert.Throws<OoxPdfLimitExceededException>(() => InvokeHasAnyValueThrowing(vectorType, huge));
        TestAssert.Throws<OoxPdfLimitExceededException>(() => ((PptxRenderer.ChartIndexedNumberVector)huge).DensePointCount());
    }

    public static void CountRenderableSeriesSkipsDensifyAllocation()
    {
        var series = CreateSparseCountProbe();
        MethodInfo count = typeof(PptxRenderer).GetMethod("CountRenderableSeries", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("Expected count helper.");
        AssertSparseCount(() => (int)count.Invoke(null, [series])!, 1, "Series presence");
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
        var series = CreateSparseCountProbe();
        AssertSparseCount(() => InvokeMaxDensePointCount(series), PptxRenderer.MaxChartDensePoints, "Slot count");
    }

    // A single dense array at the supported limit exceeds the allocation
    // tolerance. The sparse counts only need the declared slots and the one point.
    private static PptxRenderer.ChartIndexedNumberVector[] CreateSparseCountProbe() =>
    [
        new([new(0, default, 1d, "1", true, default)], null, null, null, default, [], false),
        new([], null, null, null, default, [], false),
        new([], PptxRenderer.MaxChartDensePoints, null, null, default, [], false)
    ];

    private static void AssertSparseCount(Func<int> count, int expected, string operation)
    {
        _ = count(); // Warm the reflection call before measuring.
        long before = GC.GetAllocatedBytesForCurrentThread();
        int result = count();
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        TestAssert.Equal(expected, result);
        TestAssert.True(allocated <= 256L * 1024L,
            $"{operation} allocated {allocated} bytes with a sparse {PptxRenderer.MaxChartDensePoints}-slot vector.");
    }

    public static void DuplicateCategoryCaptionsSharePlotReservationAndPaint()
    {
        string fixture = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "ComposedAxes", "solo-column.pptx");
        byte[] source = File.ReadAllBytes(fixture);
        foreach (bool visible in new[] { false, true })
        {
            PdfPage control = Render(visible, duplicate: false);
            PdfPage alternate = Render(visible, duplicate: true);
            TestAssert.Equal(control.Content, alternate.Content);
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(alternate, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal));
            string categories = string.Concat(shows.Select(show => show.Text));
            TestAssert.Equal(visible ? "W1W2W3W4W5W6W7W8" : string.Empty, categories);
        }

        PdfPage Render(bool visible, bool duplicate)
        {
            using var stream = new MemoryStream();
            stream.Write(source);
            stream.Position = 0;
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
            {
                var entry = archive.GetEntry("ppt/charts/chart1.xml")!;
                XDocument chart;
                using (var input = entry.Open()) chart = XDocument.Load(input);
                XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
                foreach (XElement point in chart.Descendants(c + "cat").Single().Descendants(c + "pt").ToArray())
                {
                    XElement value = point.Element(c + "v")!;
                    string caption = value.Value;
                    value.Value = visible ? caption : " ";
                    if (duplicate)
                    {
                        XElement predecessor = new(point);
                        predecessor.Element(c + "v")!.Value = visible ? " " : caption;
                        point.AddBeforeSelf(predecessor);
                    }
                }
                using var output = entry.Open();
                output.SetLength(0);
                chart.Save(output);
            }
            stream.Position = 0;
            using var package = OoxPackage.Open(stream, default);
            return new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(
                new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
        }
    }

    public static void RadarBuildKeepsEmptySeriesPositionsWithoutDensifying()
    {
        Type vectorType = NumberVectorType();
        Type pointType = NumberPointType();
        Array series = Array.CreateInstance(vectorType, 2);
        series.SetValue(BuildNumberVector(vectorType, pointType, [(0, null), (1, null)], 5), 0);
        series.SetValue(BuildNumberVector(vectorType, pointType, [(0, 1d)], null), 1);
        MethodInfo build = typeof(PptxRenderer).GetMethod("BuildRadarSeries", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("Expected radar builder.");
        {
            double?[][] built;
            try
            {
                built = (double?[][])build.Invoke(null, [series])!;
            }
            catch (TargetInvocationException ex)
            {
                throw ex.InnerException ?? ex;
            }

            TestAssert.Equal(2, built.Length);
            TestAssert.Equal(0, built[0].Length);
            TestAssert.Equal(1d, built[1].Single());
        }
    }

    public static void CategoryLabelMemoSharesDenseResultsWithinFrame()
    {
        var (workbook, plot, otherPlot) = CreateMemoContext();
        IReadOnlyList<PptxRenderer.ChartIndexedTextPoint?> firstResult =
            [new(0, PptxRenderer.ChartPointIndexSource.OoxmlIndex, "First", true, default)];
        IReadOnlyList<PptxRenderer.ChartIndexedTextPoint?> otherResult =
            [new(0, PptxRenderer.ChartPointIndexSource.OoxmlIndex, "Other", true, default)];
        TestAssert.True(!ReferenceEquals(firstResult, otherResult), "Factory results must be independently distinguishable.");
        int computations = 0;
        Func<IReadOnlyList<PptxRenderer.ChartIndexedTextPoint?>> firstFactory = () => { computations++; return firstResult; };
        Func<IReadOnlyList<PptxRenderer.ChartIndexedTextPoint?>> otherFactory = () => { computations++; return otherResult; };
        var first = workbook.GetOrAddCategoryLabels(plot, false, firstFactory);
        var repeat = workbook.GetOrAddCategoryLabels(plot, false, otherFactory);
        TestAssert.True(ReferenceEquals(first, repeat) && ReferenceEquals(first, firstResult), "The same plot and visibility share one result.");
        TestAssert.Equal("First", repeat.Single()!.Text);
        TestAssert.Equal(1, computations);
        TestAssert.True(ReferenceEquals(otherResult, workbook.GetOrAddCategoryLabels(plot, true, otherFactory)), "Visibility owns a distinct result.");
        TestAssert.True(ReferenceEquals(otherResult, workbook.GetOrAddCategoryLabels(otherPlot, false, otherFactory)), "Another plot owns a distinct result.");
        workbook.ClearRangeMemo();
        TestAssert.True(ReferenceEquals(otherResult, workbook.GetOrAddCategoryLabels(plot, false, otherFactory)), "Clearing recomputes the original key.");
        TestAssert.Equal(4, computations);
    }

    public static void SeriesNameMemoSharesResultsWithinFrame()
    {
        var (workbook, plot, otherPlot) = CreateMemoContext();
        IReadOnlyList<PptxRenderer.ChartSeriesNameRecord> firstResult = [new("First", PptxRenderer.ChartSeriesNameSource.Cache, default, [])];
        IReadOnlyList<PptxRenderer.ChartSeriesNameRecord> otherResult = [new("Other", PptxRenderer.ChartSeriesNameSource.Cache, default, [])];
        TestAssert.True(!ReferenceEquals(firstResult, otherResult), "Factory results must be independently distinguishable.");
        int computations = 0;
        Func<IReadOnlyList<PptxRenderer.ChartSeriesNameRecord>> firstFactory = () => { computations++; return firstResult; };
        Func<IReadOnlyList<PptxRenderer.ChartSeriesNameRecord>> otherFactory = () => { computations++; return otherResult; };
        var first = workbook.GetOrAddSeriesNames(plot, firstFactory);
        var repeat = workbook.GetOrAddSeriesNames(plot, otherFactory);
        TestAssert.True(ReferenceEquals(first, repeat) && ReferenceEquals(first, firstResult), "The same plot shares its names.");
        TestAssert.Equal("First", repeat.Single().ActiveName);
        TestAssert.Equal(1, computations);
        TestAssert.True(ReferenceEquals(otherResult, workbook.GetOrAddSeriesNames(otherPlot, otherFactory)), "Another plot owns its names.");
        workbook.ClearRangeMemo();
        var next = workbook.GetOrAddSeriesNames(plot, otherFactory);
        TestAssert.True(ReferenceEquals(next, otherResult), "Clearing recomputes the original key.");
        TestAssert.Equal("Other", next.Single().ActiveName);
        TestAssert.Equal(3, computations);
    }

    public static void SeriesVectorsMemoSharesResultsWithinFrame()
    {
        var (workbook, plot, otherPlot) = CreateMemoContext();
        IReadOnlyList<PptxRenderer.ChartIndexedNumberVector> firstResult =
            [new([new(0, PptxRenderer.ChartPointIndexSource.OoxmlIndex, 5d, "5", true, default)], 1, null, "0", default, [], false)];
        IReadOnlyList<PptxRenderer.ChartIndexedNumberVector> otherResult =
            [new([new(0, PptxRenderer.ChartPointIndexSource.OoxmlIndex, 9d, "9", true, default)], 1, null, "0", default, [], false)];
        TestAssert.True(!ReferenceEquals(firstResult, otherResult), "Factory results must be independently distinguishable.");
        int computations = 0;
        Func<IReadOnlyList<PptxRenderer.ChartIndexedNumberVector>> firstFactory = () => { computations++; return firstResult; };
        Func<IReadOnlyList<PptxRenderer.ChartIndexedNumberVector>> otherFactory = () => { computations++; return otherResult; };
        var first = workbook.GetOrAddSeriesVectors(plot, false, firstFactory);
        var repeat = workbook.GetOrAddSeriesVectors(plot, false, otherFactory);
        TestAssert.True(ReferenceEquals(first, repeat) && ReferenceEquals(first, firstResult), "The same plot and visibility share one vector list.");
        TestAssert.Equal<double?>(5d, repeat.Single().DenseValues().Single());
        TestAssert.Equal(1, computations);
        TestAssert.True(ReferenceEquals(otherResult, workbook.GetOrAddSeriesVectors(plot, true, otherFactory)), "Visibility owns a distinct vector list.");
        TestAssert.True(ReferenceEquals(otherResult, workbook.GetOrAddSeriesVectors(otherPlot, false, otherFactory)), "Another plot owns a distinct vector list.");
        workbook.ClearRangeMemo();
        var next = workbook.GetOrAddSeriesVectors(plot, false, otherFactory);
        TestAssert.True(ReferenceEquals(next, otherResult), "Clearing recomputes the original key.");
        TestAssert.Equal<double?>(9d, next.Single().DenseValues().Single());
        TestAssert.Equal(4, computations);
    }

    public static void ValueExtentsMemoSharesResultsWithinFrame()
    {
        var (workbook, plot, otherPlot) = CreateMemoContext();
        PptxRenderer.ChartValueExtents firstResult = new(0d, 10d), otherResult = new(0d, 99d);
        TestAssert.Equal(firstResult, workbook.GetOrAddValueExtents(plot, PptxSceneChartGrouping.Clustered, false, () => firstResult));
        TestAssert.Equal(firstResult, workbook.GetOrAddValueExtents(plot, PptxSceneChartGrouping.Clustered, false, () => otherResult));
        TestAssert.Equal(otherResult, workbook.GetOrAddValueExtents(otherPlot, PptxSceneChartGrouping.Clustered, false, () => otherResult));
        TestAssert.Equal(otherResult, workbook.GetOrAddValueExtents(plot, PptxSceneChartGrouping.Stacked, false, () => otherResult));
        TestAssert.Equal(otherResult, workbook.GetOrAddValueExtents(plot, PptxSceneChartGrouping.Clustered, true, () => otherResult));
        workbook.ClearRangeMemo();
        TestAssert.Equal(otherResult, workbook.GetOrAddValueExtents(plot, PptxSceneChartGrouping.Clustered, false, () => otherResult));
    }

    public static void LineValueExtentsMemoSharesResultsWithinFrame()
    {
        var (workbook, plot, otherPlot) = CreateMemoContext();
        PptxRenderer.ChartValueExtents firstResult = new(0d, 10d), otherResult = new(0d, 99d), percentResult = new(0d, 1d);
        TestAssert.Equal(firstResult, workbook.GetOrAddLineValueExtents(plot, false, false, false, () => firstResult));
        TestAssert.Equal(firstResult, workbook.GetOrAddLineValueExtents(plot, false, false, false, () => otherResult));
        TestAssert.Equal(otherResult, workbook.GetOrAddLineValueExtents(otherPlot, false, false, false, () => otherResult));
        TestAssert.Equal(otherResult, workbook.GetOrAddLineValueExtents(plot, true, false, false, () => otherResult));
        TestAssert.Equal(percentResult, workbook.GetOrAddLineValueExtents(plot, true, true, false, () => percentResult));
        TestAssert.Equal(otherResult, workbook.GetOrAddLineValueExtents(plot, false, false, true, () => otherResult));
        workbook.ClearRangeMemo();
        TestAssert.Equal(otherResult, workbook.GetOrAddLineValueExtents(plot, false, false, false, () => otherResult));
    }

    private static (PptxRenderer.ChartWorkbookData Workbook, PptxSceneChartPlot Plot, PptxSceneChartPlot OtherPlot) CreateMemoContext()
    {
        // These contracts concern frame-local memo ownership. Workbook ZIP/range
        // parsing has separate tests; no workbook package is needed here.
        var workbook = new PptxRenderer.ChartWorkbookData(
            new Dictionary<string, PptxRenderer.ChartWorksheetData>(), PptxRenderer.ChartWorkbookStyles.Empty,
            new Dictionary<string, string>(), new Dictionary<string, PptxRenderer.ChartWorkbookTable>(), [], default);
        var first = ReadPlot();
        var other = ReadPlot();
        TestAssert.True(!ReferenceEquals(first.Source, other.Source), "Independently parsed plots must retain distinct source owners.");
        return (workbook, first, other);

        static PptxSceneChartPlot ReadPlot()
        {
            const string chart = "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart><c:plotArea><c:barChart/></c:plotArea></c:chart></c:chartSpace>";
            using var input = PptxTests.CreateSingleChartPackageStream(chart);
            using var package = OoxPackage.Open(input, default);
            var document = new PptxReader().Read(package, default);
            var scene = new PptxSceneBuilder().Build(document, package, default);
            return TestAssert.NotNull(scene.Slides[0].SlideNodes[0].Chart).Plots.Single();
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
        MethodInfo densify = typeof(PptxRenderer).GetMethod("DensifyChartValueSeries", BindingFlags.NonPublic | BindingFlags.Static)
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
            return ((System.Collections.IList)vectorType.GetMethod("DenseValues")!.Invoke(vector, null)!).Count;
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
    }

    public static void CategoryLabelLookupsScaleLinearly()
    {
        // RV14: the per-frame index enumerates points once; subsequent lookups
        // never touch the points again.
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
        MethodInfo build = typeof(PptxRenderer).GetMethod("BuildCategoryLabelIndex", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("Expected category label index builder.");
        Dictionary<int, string> index;
        try
        {
            index = (Dictionary<int, string>)build.Invoke(null, [vector])!;
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }

        TestAssert.Equal((long)count, counting.Visits);
        TestAssert.Equal(count, index.Count);
        for (int i = 0; i < count; i++)
        {
            TestAssert.Equal("label" + i, index[i]);
        }

        TestAssert.Equal((long)count, counting.Visits);
    }

    public static void CategoryLabelLookupKeepsFirstMatchPrecedence()
    {
        // RV14: duplicates keep the first text; sparse, blank, and out-of-range
        // indices yield empty; out-of-order points resolve correctly.
        Type vectorType = typeof(PptxRenderer).GetNestedType("ChartIndexedTextVector", BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("Expected text vector type.");
        object vector = BuildTextVector(vectorType, [(2, "b", true), (0, "a", true), (2, "B2", true), (1, "", true), (3, "c", false)], null);
        MethodInfo build = typeof(PptxRenderer).GetMethod("BuildCategoryLabelIndex", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("Expected category label index builder.");
        Dictionary<int, string> index;
        try
        {
            index = (Dictionary<int, string>)build.Invoke(null, [vector])!;
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }

        TestAssert.Equal("a", index[0]);
        TestAssert.Equal("", index[1]);
        TestAssert.Equal("b", index[2]);
        TestAssert.True(!index.ContainsKey(3), "Uncovered indices stay absent.");
        TestAssert.True(!index.ContainsKey(5), "Out-of-range indices stay absent.");
    }

    public static void ValueOnlyLabelsSkipCategoryLookup()
    {
        // RV14: value-only labels must not touch the category lookup at all.
        Type optionsType = typeof(PptxRenderer).GetNestedType("ChartDataLabelOptions", BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("Expected label options type.");
        object options = Activator.CreateInstance(optionsType)!;
        Type namesType = typeof(PptxRenderer).GetNestedType("ChartSeriesNameRecord", BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("Expected series name type.");
        Array names = Array.CreateInstance(namesType, 0);
        MethodInfo format = typeof(PptxRenderer).GetMethod("FormatCartesianDataLabel", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("Expected cartesian label formatter.");
        string label;
        try
        {
            label = (string)format.Invoke(null, [1.5, 0, 0, null, null, null, options, new ThrowingCategoryIndex(), names])!;
        }
        catch (TargetInvocationException ex)
        {
            throw new InvalidOperationException("Value-only labels must not consult the category lookup.", ex.InnerException ?? ex);
        }

        TestAssert.Equal("", label);
    }

    // Explicit sparse-slot expectations cover holes, null values, last-wins
    // duplicates, negative indexes and the declared slot count.
    public static void DenseNumberValueViewPreservesSparseSlots()
    {
        Type vectorType = NumberVectorType();
        object vector = BuildNumberVector(vectorType, NumberPointType(),
            [(0, 5d), (2, null), (5, 7d), (5, 9d), (-1, 1d)], 8);
        double?[] values = InvokeDenseValues(vectorType, vector);
        double?[] expected = [5d, null, null, null, null, 9d, null, null];
        TestAssert.Equal(expected.Length, values.Length);
        for (int i = 0; i < values.Length; i++)
        {
            TestAssert.True(expected[i] == values[i], $"Slot {i} must hold {expected[i]}.");
        }
    }

    // RV20: the compact view must stay near one nullable double per slot at
    // N/2N/4N declared slots, empty or sparse, instead of one fat provenance
    // struct per slot.
    public static void DenseNumberValueViewAllocationPerSlotIsCompact()
    {
        Type vectorType = NumberVectorType();
        Type pointType = NumberPointType();
        foreach (int count in new[] { 1000, 2000, 4000 })
        {
            object vector = BuildNumberVector(vectorType, pointType, [], count);
            long allocated = MeasureDenseValuesAllocation(vectorType, vector);
            TestAssert.True(allocated <= 32L * count, string.Format("Compact dense values must stay near 16 bytes per slot, allocated {0} bytes for {1} slots.", allocated, count));
        }

        object sparse = BuildNumberVector(vectorType, pointType, [(0, 1d), (3999, 2d)], 4000);
        long sparseAllocated = MeasureDenseValuesAllocation(vectorType, sparse);
        TestAssert.True(sparseAllocated <= 32L * 4000, string.Format("Sparse compact dense values must stay compact, allocated {0} bytes.", sparseAllocated));
    }

    public static void DenseNumberPointIndexPreservesSparseProvenance()
    {
        Type vectorType = NumberVectorType();
        object vector = BuildNumberVector(vectorType, NumberPointType(),
            [(0, 5d), (2, null), (5, 7d), (5, 9d), (-1, 1d)], 8);
        MethodInfo build = typeof(PptxRenderer).GetMethod("BuildDensePointIndex", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("Expected rendering provenance index.");
        var index = (System.Collections.IDictionary)build.Invoke(null, [vector])!;
        var source = ((System.Collections.IEnumerable)vectorType.GetProperty("Points")!.GetValue(vector)!).Cast<object>().ToArray();
        TestAssert.Equal(2, index.Count);
        TestAssert.True(source[0].Equals(index[0]), "First source point must retain complete provenance.");
        TestAssert.True(source[3].Equals(index[5]), "Last valid duplicate must retain complete provenance.");
        foreach (int absent in new[] { -1, 1, 2, 3, 4, 6, 7, 8 })
        {
            TestAssert.True(!index.Contains(absent), $"Slot {absent} must be absent.");
        }
    }

    private static double?[] InvokeDenseValues(Type vectorType, object vector)
    {
        try
        {
            MethodInfo dense = vectorType.GetMethod("DenseValues")
                ?? throw new InvalidOperationException("Expected dense value view.");
            return (double?[])dense.Invoke(vector, null)!;
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
    }

    private static long MeasureDenseValuesAllocation(Type vectorType, object vector)
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        long before = GC.GetAllocatedBytesForCurrentThread();
        double?[] values = InvokeDenseValues(vectorType, vector);
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        GC.KeepAlive(values);
        return allocated;
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

    private sealed class ThrowingCategoryIndex : IReadOnlyDictionary<int, string>
    {
        public string this[int key] => throw new InvalidOperationException("Category lookup must be lazy.");

        public IEnumerable<int> Keys => throw new InvalidOperationException("Category lookup must be lazy.");

        public IEnumerable<string> Values => throw new InvalidOperationException("Category lookup must be lazy.");

        public int Count => throw new InvalidOperationException("Category lookup must be lazy.");

        public bool ContainsKey(int key) => throw new InvalidOperationException("Category lookup must be lazy.");

        public bool TryGetValue(int key, out string value) => throw new InvalidOperationException("Category lookup must be lazy.");

        public IEnumerator<KeyValuePair<int, string>> GetEnumerator() => throw new InvalidOperationException("Category lookup must be lazy.");

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }

}
