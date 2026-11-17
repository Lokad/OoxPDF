using System.Reflection;
using System.Xml.Linq;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxWorkbookOwnershipTests
{
    public static void PptxWorkbookFreshRangesUseOwnedCancellationAndPreserveMemoHits()
    {
        using var cancellation = new CancellationTokenSource();
        var workbook = Read(Parts(PptxTests.EmbeddedChartWorkbook()), cancellation.Token);
        var cached = workbook.ReadRangeCells("Sheet1!A2");
        cancellation.Cancel();
        TestAssert.True(ReferenceEquals(cached, workbook.ReadRangeCells("Sheet1!A2")),
            "A completed memo hit still returns its existing immutable range without expansion.");
        var failure = TestAssert.Throws<OperationCanceledException>(() => workbook.ReadRangeCells("Sheet1!A1:A100000"));
        TestAssert.Equal(cancellation.Token, failure.CancellationToken);
        workbook.ClearRangeMemo();
        TestAssert.Throws<OperationCanceledException>(() => workbook.ReadRangeCells("Sheet1!A2"));
    }

    public static void PptxWorkbookLiveProjectionsObserveOwnedCancellation()
    {
        foreach (string consumer in new[] { "ReadWorkbookNumberPoints", "ReadWorkbookTextPoints" })
        {
            using var cancellation = new CancellationTokenSource();
            var workbook = Read(Parts(PptxTests.EmbeddedChartWorkbook()), cancellation.Token);
            var source = new PptxSceneChartDataSource("Sheet1!A1:A100000",
                PptxSceneChartDataSourceReferenceKind.NumberReference, "numRef",
                PptxSceneChartDataSourceCacheKind.NumberCache, "numCache", true);
            MethodInfo projection = typeof(PptxRenderer).GetMethod(consumer, BindingFlags.NonPublic | BindingFlags.Static)!;
            cancellation.Cancel();
            var failure = TestAssert.Throws<OperationCanceledException>(() => Invoke());
            TestAssert.Equal(cancellation.Token, failure.CancellationToken);

            void Invoke()
            {
                try { projection.Invoke(null, [workbook, source]); }
                catch (TargetInvocationException ex) { throw ex.InnerException ?? ex; }
            }
        }
    }

    public static void PptxWorkbookBaseTextExcludesPhoneticsAcrossStorageAndRuns()
    {
        XNamespace s = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        foreach (bool inline in new[] { false, true })
        foreach (bool rich in new[] { false, true })
        foreach (bool annotated in new[] { false, true })
        {
            var parts = Parts(PptxTests.EmbeddedChartWorkbook());
            var item = new XElement(s + (inline ? "is" : "si"));
            XElement Text(string value) => new(s + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), value);
            if (rich) item.Add(new XElement(s + "r", Text(" No")), new XElement(s + "r", Text("rth ")));
            else item.Add(Text(" North "));
            if (annotated) item.Add(Phonetic(s));
            if (inline)
                ChangeXml(parts, "xl/worksheets/sheet1.xml", xml =>
                {
                    var cell = xml.Descendants(s + "c").Single(c => (string?)c.Attribute("r") == "A2");
                    cell.SetAttributeValue("t", "inlineStr"); cell.ReplaceNodes(item);
                });
            else ChangeXml(parts, "xl/sharedStrings.xml", xml => xml.Root!.Elements(s + "si").First().ReplaceWith(item));
            var value = Read(parts).ReadRangeCells("Sheet1!A2").Single();
            TestAssert.Equal(" North ", value.Text);
            TestAssert.True(value.HasValue, "Both storage forms preserve the authored string value.");

        }
    }

    public static void PptxWorkbookPhoneticWhitespacePreservesBaseText()
    {
        var parts = Parts(PptxTests.EmbeddedChartWorkbook());
        ChangeXml(parts, "xl/sharedStrings.xml", xml =>
        {
            XNamespace s = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var phonetic = Phonetic(s);
            phonetic.Element(s + "t")!.SetAttributeValue(XNamespace.Xml + "space", "preserve");
            xml.Root!.Elements(s + "si").First().ReplaceNodes(new XElement(s + "t", "North"), phonetic);
        });
        var value = Read(parts).ReadRangeCells("Sheet1!A2").Single();
        TestAssert.Equal("North", value.Text);
        TestAssert.True(value.HasValue, "Phonetic whitespace must not alter the stored base value.");
    }

    public static void PptxWorkbookPhoneticsAndStorageKeepCachedChartPaint()
    {
        string source = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
            "pptx-ladder-11-chart-bar-clustered-port.pptx"));
        var original = Parts(File.ReadAllBytes(source));
        var baseline = Convert(original);
        foreach (bool inline in new[] { false, true })
        {
            var parts = new Dictionary<string, byte[]>(original);
            const string embedded = "ppt/embeddings/Microsoft_Excel_Sheet1.xlsx";
            var workbook = Parts(parts[embedded]);
            XNamespace s = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var strings = XDocument.Parse(System.Text.Encoding.UTF8.GetString(workbook["xl/sharedStrings.xml"]))
                .Root!.Elements(s + "si").Select(item => new XElement(item)).ToArray();
            foreach (var item in strings.Where(item => item.Value.Length != 0)) item.Add(Phonetic(s));
            if (inline)
            {
                foreach (string sheet in workbook.Keys.Where(key => key.StartsWith("xl/worksheets/", StringComparison.Ordinal) &&
                    key.EndsWith(".xml", StringComparison.Ordinal)).ToArray())
                    ChangeXml(workbook, sheet, xml =>
                    {
                        foreach (var cell in xml.Descendants(s + "c").Where(c => (string?)c.Attribute("t") == "s"))
                        {
                            int index = int.Parse(cell.Element(s + "v")!.Value, System.Globalization.CultureInfo.InvariantCulture);
                            cell.SetAttributeValue("t", "inlineStr");
                            cell.ReplaceNodes(new XElement(s + "is", strings[index].Nodes()));
                        }
                    });
            }
            else ChangeXml(workbook, "xl/sharedStrings.xml", xml => xml.Root!.ReplaceNodes(strings));
            using (var bytes = TestFixtures.CreateZipPackage(workbook)) parts[embedded] = bytes.ToArray();
            var current = Convert(parts);
            TestAssert.True(baseline.Bytes.AsSpan().SequenceEqual(current.Bytes),
                "Workbook string representation/provenance must not change active chart-cache paint.");
            TestAssert.True(baseline.Diagnostics.SequenceEqual(current.Diagnostics), "Cached-chart diagnostics remain exact.");
        }
    }

    private static XElement Phonetic(XNamespace s) => new(s + "rPh", new XAttribute("sb", "0"),
        new XAttribute("eb", "1"), new XElement(s + "t", "HINT"));
}
