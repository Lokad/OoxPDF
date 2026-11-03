using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    public static void PptxChartAxisExpansionAdmitsGeneratedStorageWithoutChangingPaint()
    {
        var small = ObserveAxisConstruction(10);
        var dense = ObserveAxisConstruction(99_998);
        TestAssert.True(dense.Request > small.Request && dense.Request > 8L * 100_000,
            "The real producer must fund the large generated tick store before painting.");
        TestAssert.Equal(99_998, CountAxisGridlines(dense.Content));
        var ungoverned = ObserveAxisConstruction(99_998, governed: false);
        TestAssert.Equal(ungoverned.Content, dense.Content);
    }

    public static void PptxChartAxisFramesReuseCapacityAndRefundOnPageExit()
    {
        foreach (bool early in new[] { false, true })
        {
            var observed = ObserveAxisConstruction(1_000, frames: 2, repeats: 2, early: early);
            TestAssert.Equal(2_000, CountAxisGridlines(observed.Content));
            TestAssert.True(observed.Growths == (early ? 1 : 2),
                "Repeated measurement/painting and sequential frames reuse the same page window.");
        }
    }

    public static void PptxHiddenAxisPaintAvoidsGeneratedStorageAdmission()
    {
        var small = ObserveAxisConstruction(10, visible: false);
        var dense = ObserveAxisConstruction(99_998, visible: false);
        TestAssert.Equal(small.Request, dense.Request);
        TestAssert.Equal(0, CountAxisGridlines(dense.Content));
    }

    public static void PptxChartAxisRefusalAndCancellationPreservePublication()
    {
        long request = ObserveAxisConstruction(10_000).Request;
        foreach (bool cancel in new[] { false, true })
        foreach (bool file in new[] { false, true })
        {
            string directory = Path.Combine(Path.GetTempPath(), "axis-admission-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                using var cancellation = new CancellationTokenSource();
                int selected = 0;
                var budget = new TestMemoryBudget { RefuseChange = (before, after) => {
                    if (before <= 0 || after != request) return false;
                    selected++;
                    if (cancel) { cancellation.Cancel(); return false; }
                    return true;
                }};
                using var source = AxisConstructionPackage(10_000);
                string input = Path.Combine(directory, "input.pptx"), target = Path.Combine(directory, "output.pdf");
                File.WriteAllBytes(input, source.ToArray()); File.WriteAllBytes(target, [7, 8, 9]);
                using var output = new MemoryStream();
                var options = new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx, FontResolver = new TestFaceFontResolver(), MemoryBudget = budget };
                void Convert() => (file ? OoxPdfConverter.ConvertAsync(input, target, options, cancellation.Token) :
                    OoxPdfConverter.ConvertAsync(source, output, options, cancellation.Token)).GetAwaiter().GetResult();
                if (cancel) TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(Convert).CancellationToken);
                else
                {
                    var failure = TestAssert.Throws<OoxPdfLimitExceededException>(Convert);
                    TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
                    TestAssert.Equal(OoxPdfLimitScope.Conversion, failure.Scope);
                }
                TestAssert.Equal(1, selected);
                TestAssert.Equal(0L, output.Length);
                TestAssert.True(File.ReadAllBytes(target).SequenceEqual(new byte[] { 7, 8, 9 }), "Preserve the destination.");
                TestAssert.True(source.CanRead && output.CanWrite, "Caller streams remain open.");
                TestAssert.Equal(0, Directory.GetFiles(directory, "*.tmp-*").Length);
                TestAssert.Equal(0L, budget.Reserved);
                TestAssert.Equal(budget.ScopesCreated, budget.ScopesDisposed);
                TestAssert.Equal(0, budget.CallerDisposals);
            }
            finally { Directory.Delete(directory, true); }
        }
    }

    private static (long Request, int Growths, string Content) ObserveAxisConstruction(int intervals,
        int frames = 1, int repeats = 1, bool early = false, bool governed = true, bool visible = true)
    {
        using var input = AxisConstructionPackage(intervals, frames, visible);
        using var package = OoxPackage.Open(input, default);
        var document = new PptxReader().Read(package, default);
        document = document with { Slides = Enumerable.Range(0, repeats).Select(i => new PptxSlide(document.Slides[0].PartName, i)).ToArray() };
        var budget = new PageMemoryBudget();
        var growths = new List<long>();
        budget.Pool.RefuseChange = (before, after) => { if (before > 0 && after > before) growths.Add(after); return false; };
        long request = 0; string content = "";
        using (var operation = OoxConversionBudget.BeginScope(null, governed ? budget : null))
        {
            var reader = new PptxRenderer(new TestFaceFontResolver()).RenderPagesAsync(document, package, null, default).GetAsyncEnumerator();
            IMemoryReservation? previous = null;
            try
            {
                for (int i = 0; i < repeats; i++)
                {
                    TestAssert.True(reader.MoveNextAsync().AsTask().GetAwaiter().GetResult(), "Render the generated axis.");
                    if (previous is not null) TestAssert.Equal(0L, previous.Bytes);
                    if (governed) { previous = budget.LargestLiveGrant(); request = Math.Max(request, previous.Bytes); }
                    content = reader.Current.Content;
                    TestAssert.True(content.Length > 100, "Exercise painted chart content.");
                    if (early) break;
                }
                if (!early) TestAssert.True(!reader.MoveNextAsync().AsTask().GetAwaiter().GetResult(), "Drain all pages.");
            }
            finally { reader.DisposeAsync().AsTask().GetAwaiter().GetResult(); }
            if (previous is not null) TestAssert.Equal(0L, previous.Bytes);
        }
        TestAssert.Equal(0L, budget.Pool.Reserved);
        TestAssert.Equal(budget.Pool.ScopesCreated, budget.Pool.ScopesDisposed);
        return (request, growths.Count(value => value == request), content);
    }

    private static int CountAxisGridlines(string content)
    {
        string[] lines = content.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n');
        int count = 0;
        for (int i = 0; i < lines.Length; i++)
        {
            if (lines[i] != "1 0 0 RG") continue;
            while (++i < lines.Length && lines[i] != "S")
                if (lines[i].EndsWith(" l", StringComparison.Ordinal)) count++;
        }
        return count;
    }

    private static MemoryStream AxisConstructionPackage(int intervals, int frames = 1, bool visible = true)
    {
        string chart = $$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea>
              <c:lineChart><c:grouping val="standard"/><c:ser><c:idx val="0"/><c:order val="0"/>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val>
              </c:ser><c:axId val="10"/><c:axId val="20"/></c:lineChart>
              <c:catAx><c:axId val="10"/><c:delete val="1"/><c:crossAx val="20"/></c:catAx>
              <c:valAx><c:axId val="20"/><c:scaling><c:min val="0"/><c:max val="{{intervals}}"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:majorUnit val="1"/>
                <c:majorGridlines><c:spPr><a:ln w="12700">{{(visible ? "<a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill>" : "<a:noFill/>")}}</a:ln></c:spPr></c:majorGridlines><c:crossAx val="10"/>
              </c:valAx></c:plotArea></c:chart></c:chartSpace>
            """;
        using var original = PptxTests.CreateSingleChartPackageStream(chart);
        if (frames == 1) return new MemoryStream(original.ToArray());
        var parts = ModelAdmissionParts(original.ToArray());
        var slide = XDocument.Parse(Encoding.UTF8.GetString(parts["ppt/slides/slide1.xml"]));
        XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
        var tree = slide.Descendants(p + "spTree").Single();
        var frame = tree.Element(p + "graphicFrame")!;
        for (int i = 1; i < frames; i++) tree.Add(new XElement(frame));
        parts["ppt/slides/slide1.xml"] = TestFixtures.Utf8(slide.ToString());
        return TestFixtures.CreateZipPackage(parts);
    }
}
