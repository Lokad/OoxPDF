using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    public static void SvgSampledPageConstructionCoversReturnedPaintAndOpacityGroups()
    {
        foreach (var kind in new[] { OoxPdfInputKind.Docx, OoxPdfInputKind.Pptx })
        foreach (bool groups in new[] { false, true })
        {
            var actual = ObserveSampledSvg(kind, groups: groups);
            TestAssert.Equal(64 * 512, actual.Curves);
            TestAssert.True(actual.Request >= actual.PayloadBytes,
                "The existing page grant must cover significant generated content, including group snapshots.");
            TestAssert.Equal(groups ? 2 : 0, actual.Groups);
            var control = ObserveSampledSvg(kind, groups: groups, governed: false);
            TestAssert.Equal(control.Paint, actual.Paint);
        }
    }

    public static void SvgNativeShadingAvoidsSampledConstructionGrowth()
    {
        foreach (var kind in new[] { OoxPdfInputKind.Docx, OoxPdfInputKind.Pptx })
        {
            var sampled = ObserveSampledSvg(kind);
            var native = ObserveSampledSvg(kind, native: true);
            TestAssert.Equal(0, native.Curves);
            TestAssert.True(native.Shadings > 0, "Exercise native shading rather than an omitted picture.");
            TestAssert.True(native.Request < sampled.Request / 4,
                "Native shading must not reserve the sampled approximation's expansion.");
        }
    }

    public static void SvgSampledPicturePlacementsShareCapacityAndRefundPerPage()
    {
        foreach (var kind in new[] { OoxPdfInputKind.Docx, OoxPdfInputKind.Pptx })
        foreach (bool early in new[] { false, true })
        {
            var one = ObserveSampledSvg(kind, count: 16, groups: true);
            var repeated = ObserveSampledSvg(kind, count: 16, groups: true, copies: 2, pages: 2, early: early);
            TestAssert.Equal(16 * 512 * 2, repeated.Curves);
            TestAssert.True(repeated.Request >= repeated.PayloadBytes && repeated.Request < one.Request * 3,
                "Coexisting picture output is funded without accumulating an allowance for every paint attempt.");
        }
    }

    public static void SvgSampledConstructionRefusalAndCancellationPreservePublication()
    {
        foreach (var kind in new[] { OoxPdfInputKind.Docx, OoxPdfInputKind.Pptx })
        {
            long request = ObserveSampledSvg(kind, count: 16, groups: true).Request;
            foreach (bool cancel in new[] { false, true })
            foreach (bool file in new[] { false, true })
            {
                string directory = Path.Combine(Path.GetTempPath(), "svg-expansion-" + Guid.NewGuid().ToString("N"));
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
                    using var source = SampledSvgPackage(kind, 16, true, false, 1, 1);
                    string input = Path.Combine(directory, kind == OoxPdfInputKind.Docx ? "input.docx" : "input.pptx");
                    string target = Path.Combine(directory, "output.pdf");
                    File.WriteAllBytes(input, source.ToArray()); File.WriteAllBytes(target, [7, 8, 9]);
                    using var output = new MemoryStream();
                    var options = new OoxPdfOptions { InputKind = kind, FontResolver = new TestFaceFontResolver(), MemoryBudget = budget };
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
    }

    public static void SvgSampledConstructionUsesBodyTableAndStaticStoryPageOwners()
    {
        foreach (string template in new[] { "docx-svg-inline.docx", "docx-svg-table.docx", "docx-svg-header.docx", "docx-svg-footer.docx" })
        {
            var actual = ObserveSampledSvg(OoxPdfInputKind.Docx, count: 16, groups: true, template: template);
            TestAssert.Equal(16 * 512, actual.Curves);
            TestAssert.True(actual.Request >= actual.PayloadBytes, "Fund the emitted SVG independent of its placement container.");
        }
    }

    private static (long Request, long PayloadBytes, int Curves, int Groups, int Shadings, string Paint) ObserveSampledSvg(
        OoxPdfInputKind kind, int count = 64, bool groups = false, bool native = false, int copies = 1, int pages = 1,
        bool early = false, bool governed = true, string? template = null)
    {
        using var input = SampledSvgPackage(kind, count, groups, native, copies, pages, template);
        using var package = Lokad.OoxPdf.Ooxml.OoxPackage.Open(input, default);
        var reader = PageMemoryPages(package, kind, new TestFaceFontResolver(), pages, default).GetAsyncEnumerator();
        var budget = new PageMemoryBudget();
        long request = 0, payload = 0; int curves = 0, groupCount = 0, shadings = 0; string paint = "";
        using (var operation = OoxConversionBudget.BeginScope(null, governed ? budget : null))
        {
            IMemoryReservation? previous = null;
            try
            {
                for (int i = 0; i < pages; i++)
                {
                    TestAssert.True(reader.MoveNextAsync().AsTask().GetAwaiter().GetResult(), "Render the SVG page.");
                    if (previous is not null) TestAssert.Equal(0L, previous.Bytes);
                    if (governed) { previous = budget.LargestLiveGrant(); request = Math.Max(request, previous.Bytes); }
                    var page = reader.Current;
                    var contents = new List<string> { page.Content };
                    groupCount = 0; shadings = page.Shadings.Count;
                    ReadGroups(page.Groups);
                    payload = contents.Sum(value => 2L * value.Length);
                    paint = string.Join("\n", contents);
                    curves = paint.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n').Count(line => line.EndsWith(" c", StringComparison.Ordinal));
                    if (early) break;
                    void ReadGroups(IReadOnlyList<PdfTransparencyGroupResource> entries)
                    {
                        foreach (var entry in entries)
                        {
                            groupCount++; contents.Add(entry.Group.Content); shadings += entry.Group.Shadings.Count;
                            ReadGroups(entry.Group.Groups);
                        }
                    }
                }
                if (!early) TestAssert.True(!reader.MoveNextAsync().AsTask().GetAwaiter().GetResult(), "Drain the expected pages.");
            }
            finally { reader.DisposeAsync().AsTask().GetAwaiter().GetResult(); }
            if (previous is not null) TestAssert.Equal(0L, previous.Bytes);
        }
        TestAssert.Equal(0L, budget.Pool.Reserved);
        TestAssert.Equal(budget.Pool.ScopesCreated, budget.Pool.ScopesDisposed);
        return (request, payload, curves, groupCount, shadings, paint);
    }

    private static MemoryStream SampledSvgPackage(OoxPdfInputKind kind, int count, bool groups, bool native, int copies, int pages, string? template = null)
    {
        string name = template ?? (kind == OoxPdfInputKind.Docx ? "docx-svg-native-only.docx" : "pptx-svg-linear-gradient-stroke.pptx");
        var parts = ModelAdmissionParts(File.ReadAllBytes(Case(name)));
        string paths = string.Concat(Enumerable.Range(0, count).Select(i =>
            (groups && (i == 0 || i == count / 2) ? "<g opacity=\"0.5\">" : "") +
            $"<rect x=\"{i}\" y=\"0\" width=\"1024\" height=\"1024\" fill=\"url(#g)\"/>" +
            (groups && (i == count / 2 - 1 || i == count - 1) ? "</g>" : "")));
        string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"1024\" height=\"1024\" viewBox=\"0 0 1024 1024\"><defs><radialGradient id=\"g\" gradientUnits=\"userSpaceOnUse\" cx=\"512\" cy=\"512\" r=\"512\">" +
            (native ? "" : "<stop offset=\"0\" stop-color=\"#FF0000\"/>") +
            "<stop offset=\"0\" stop-color=\"#0000FF\"/><stop offset=\"1\" stop-color=\"#00FF00\"/></radialGradient></defs>" + paths + "</svg>";
        foreach (string image in parts.Keys.Where(key => key.EndsWith(".svg", StringComparison.OrdinalIgnoreCase)).ToArray())
            parts[image] = Encoding.UTF8.GetBytes(svg);
        if (kind == OoxPdfInputKind.Docx && pages > 1)
        {
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var xml = XDocument.Parse(Encoding.UTF8.GetString(parts["word/document.xml"]));
            var body = xml.Root!.Element(w + "body")!; var section = body.Element(w + "sectPr"); section?.Remove();
            var original = body.Elements().Select(e => new XElement(e)).ToArray();
            for (int i = 1; i < pages; i++)
            {
                body.Add(new XElement(w + "p", new XElement(w + "r", new XElement(w + "br", new XAttribute(w + "type", "page")))));
                foreach (var element in original) body.Add(new XElement(element));
            }
            if (section is not null) body.Add(section);
            parts["word/document.xml"] = TestFixtures.Utf8(xml.ToString());
        }
        if (kind == OoxPdfInputKind.Pptx && copies > 1)
        {
            XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
            var xml = XDocument.Parse(Encoding.UTF8.GetString(parts["ppt/slides/slide1.xml"]));
            var tree = xml.Descendants(p + "spTree").Single(); var picture = tree.Element(p + "pic")!;
            for (int i = 1; i < copies; i++) tree.Add(new XElement(picture));
            parts["ppt/slides/slide1.xml"] = TestFixtures.Utf8(xml.ToString());
        }
        if (kind == OoxPdfInputKind.Docx && copies > 1)
        {
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var xml = XDocument.Parse(Encoding.UTF8.GetString(parts["word/document.xml"]));
            foreach (var paragraph in xml.Descendants(w + "p").Where(e => e.Descendants(w + "drawing").Any()).ToArray())
                for (int i = 1; i < copies; i++) paragraph.AddAfterSelf(new XElement(paragraph));
            parts["word/document.xml"] = TestFixtures.Utf8(xml.ToString());
        }
        return TestFixtures.CreateZipPackage(parts);
    }
}
