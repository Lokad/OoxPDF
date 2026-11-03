using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    public static void PptxWorkbookRangeConstructionFollowsResolvedSourceAndRefunds()
    {
        var compact = ObserveWorkbookRange(2);
        foreach (string form in new[] { "direct", "reversed", "union", "name" })
        {
            var large = ObserveWorkbookRange(10_000, form: form);
            TestAssert.True(large.Request > compact.Request + 8L * 10_000,
                "Significant formula expansion must be admitted even with only two cached points.");
            TestAssert.Equal(compact.Content, large.Content);
        }
    }

    public static void PptxWorkbookRangeFramesReusePageCapacity()
    {
        var one = ObserveWorkbookRange(10_000);
        var two = ObserveWorkbookRange(10_000, frames: 2);
        TestAssert.True(two.Request >= one.Request && two.Request < one.Request + one.Request / 4,
            "Sequential frames share the page's existing construction window.");
    }

    public static void PptxWorkbookRangeCapacityFollowsPageAdvanceAndEarlyDisposal()
    {
        foreach (bool early in new[] { false, true }) ObserveWorkbookRange(1_000, repeats: 2, early: early);
    }

    public static void PptxWorkbookRangeRefusalAndCancellationPreservePublication()
    {
        foreach (bool cancel in new[] { false, true })
        foreach (bool file in new[] { false, true })
        {
            string directory = Path.Combine(Path.GetTempPath(), "workbook-admission-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                using var cancellation = new CancellationTokenSource();
                int selected = 0;
                var budget = new TestMemoryBudget { RefuseChange = (before, after) => {
                    if (before <= 0 || after - before < 8L * 1024 * 1024) return false;
                    selected++;
                    if (cancel) { cancellation.Cancel(); return false; }
                    return true;
                }};
                using var source = WorkbookRangePackage(10_000);
                string input = Path.Combine(directory, "input.pptx"), target = Path.Combine(directory, "output.pdf");
                File.WriteAllBytes(input, source.ToArray()); File.WriteAllBytes(target, [7, 8, 9]);
                using var output = new MemoryStream();
                var options = new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx, FontResolver = new TestFaceFontResolver(), MemoryBudget = budget };
                void Convert() => (file ? OoxPdfConverter.ConvertAsync(input, target, options, cancellation.Token) :
                    OoxPdfConverter.ConvertAsync(source, output, options, cancellation.Token)).GetAwaiter().GetResult();
                if (cancel) TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(Convert).CancellationToken);
                else
                {
                    var refusal = TestAssert.Throws<OoxPdfLimitExceededException>(Convert);
                    TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, refusal.Kind);
                    TestAssert.Equal(OoxPdfLimitScope.Conversion, refusal.Scope);
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

    private static (long Request, string Content) ObserveWorkbookRange(int cells, int frames = 1, string form = "direct", int repeats = 1, bool early = false)
    {
        using var input = WorkbookRangePackage(cells, frames, form);
        using var package = OoxPackage.Open(input, default);
        var document = new PptxReader().Read(package, default);
        document = document with { Slides = Enumerable.Range(0, repeats).Select(i => new PptxSlide(document.Slides[0].PartName, i)).ToArray() };
        var budget = new PageMemoryBudget();
        long request = 0; string content = "";
        using (var operation = OoxConversionBudget.BeginScope(null, budget))
        {
            var reader = new PptxRenderer(new TestFaceFontResolver()).RenderPagesAsync(document, package, null, default).GetAsyncEnumerator();
            IMemoryReservation? previous = null;
            try
            {
                for (int i = 0; i < repeats; i++)
                {
                    TestAssert.True(reader.MoveNextAsync().AsTask().GetAwaiter().GetResult(), "Paint the workbook-backed chart.");
                    if (previous is not null) TestAssert.Equal(0L, previous.Bytes);
                    previous = budget.LargestLiveGrant();
                    request = Math.Max(request, previous.Bytes); content = reader.Current.Content;
                    TestAssert.True(content.Length > 100, "Do not qualify a blank/unrenderable control.");
                    if (early) break;
                }
                if (!early) TestAssert.True(!reader.MoveNextAsync().AsTask().GetAwaiter().GetResult(), "Drain all pages.");
            }
            finally { reader.DisposeAsync().AsTask().GetAwaiter().GetResult(); }
            TestAssert.Equal(0L, previous!.Bytes);
        }
        TestAssert.Equal(0L, budget.Pool.Reserved);
        TestAssert.Equal(budget.Pool.ScopesCreated, budget.Pool.ScopesDisposed);
        return (request, content);
    }

    private static MemoryStream WorkbookRangePackage(int cells, int frames = 1, string form = "direct")
    {
        var parts = ModelAdmissionParts(File.ReadAllBytes(Case("chart-workbook-range-source.pptx")));
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        var chart = XDocument.Parse(Encoding.UTF8.GetString(parts["ppt/charts/chart1.xml"]));
        foreach (var source in chart.Descendants(c + "strRef").Concat(chart.Descendants(c + "numRef")))
        {
            string column = source.Name.LocalName == "strRef" ? "A" : "B";
            source.Element(c + "f")!.Value = form switch {
                "reversed" => $"Sheet1!${column}${cells+1}:${column}$2",
                "union" => $"Sheet1!${column}$2:${column}${cells/2+1},Sheet1!${column}${cells/2+2}:${column}${cells+1}",
                "name" when column == "B" => "GrowingValues",
                _ => $"Sheet1!${column}$2:${column}${cells+1}"
            };
        }
        if (form == "name")
        {
            var workbook = ModelAdmissionParts(parts["ppt/embeddings/book.xlsx"]);
            XNamespace s = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var xml = XDocument.Parse(Encoding.UTF8.GetString(workbook["xl/workbook.xml"]));
            XElement names = xml.Root!.Element(s + "definedNames") ?? new XElement(s + "definedNames");
            if (names.Parent is null) xml.Root.Add(names);
            names.Add(new XElement(s + "definedName", new XAttribute("name", "GrowingValues"), $"Sheet1!$B$2:$B${cells+1}"));
            workbook["xl/workbook.xml"] = TestFixtures.Utf8(xml.ToString());
            using var book = TestFixtures.CreateZipPackage(workbook); parts["ppt/embeddings/book.xlsx"] = book.ToArray();
        }
        parts["ppt/charts/chart1.xml"] = TestFixtures.Utf8(chart.ToString());
        if (frames > 1)
        {
            XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
            var slide = XDocument.Parse(Encoding.UTF8.GetString(parts["ppt/slides/slide1.xml"]));
            var tree = slide.Descendants(p + "spTree").Single();
            var frame = tree.Element(p + "graphicFrame")!;
            for (int i = 1; i < frames; i++) tree.Add(new XElement(frame));
            parts["ppt/slides/slide1.xml"] = TestFixtures.Utf8(slide.ToString());
        }
        return TestFixtures.CreateZipPackage(parts);
    }
}
