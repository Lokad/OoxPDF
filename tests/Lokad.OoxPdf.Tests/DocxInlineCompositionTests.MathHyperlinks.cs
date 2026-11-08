using System.Xml.Linq;
using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxMathHyperlinksKeepRootAndArgumentSourceOwners()
    {
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        foreach (bool argument in new[] { false, true })
        foreach (bool metadata in new[] { false, true })
        {
            var document = ReadMutation("docx-omml-isolated-inline-body-fraction-left", mode, true, false,
                (xml, w, m) =>
                {
                    var math = xml.Descendants(m + "oMath").First();
                    var owner = argument ? math.Descendants(m + "num").First() : math;
                    var children = owner.Elements().Where(child => !child.Name.LocalName.EndsWith("Pr", StringComparison.Ordinal)).ToArray();
                    foreach (var child in children) child.Remove();
                    var link = new XElement(w + "hyperlink", new XAttribute(w + "anchor", "PublicCalculation"), children);
                    owner.Add(metadata ? new XElement(w + "customXml", link) : link);
                });
            var link = DocxBlockTraversal.EnumerateBodyParagraphs(document.BodyElements)
                .SelectMany(paragraph => paragraph.Hyperlinks).Single(item => item.IsMathContent);
            TestAssert.Equal(!argument, link.IsMathRootContent);
        }
    }

    public static void DocxMathHyperlinksMergeRootFragmentsOnTheirPlacedLine()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string container in new[] { "body", "cell", "header", "footer" })
        foreach (bool partial in new[] { false, true })
        {
            var renderer = MathContentControlRenderer(mode);
            var before = renderer.RenderPagesAsync(Read(false), null, default).ToBlockingEnumerable().ToArray();
            var after = renderer.RenderPagesAsync(Read(true), null, default).ToBlockingEnumerable().ToArray();
            AssertMathLinkPaint(before, after);
            var links = after.SelectMany(page => page.Annotations).ToList();
            foreach (var existing in before.SelectMany(page => page.Annotations))
                TestAssert.True(links.Remove(existing), "Adding a root link preserves existing navigation.");
            TestAssert.Equal(1, links.Count);
            TestAssert.True(links[0].Destination is not null, "One source target owns adjacent root fragments on the same line.");

            DocxDocument Read(bool linked) => ReadMutation("CompositionProbes/" + (container == "cell" ? "body" : container),
                mode, true, false, (xml, w, m) =>
                {
                    var math = xml.Descendants(m + "oMath").First();
                    math.RemoveNodes();
                    var run = new XElement(m + "r", new XElement(m + "t", "u+v=w"));
                    math.Add(linked ? new XElement(w + "hyperlink", new XAttribute(w + "anchor", "PublicCalculation"), run) : run);
                    if (partial) math.Add(new XElement(m + "r", new XElement(m + "t", "+z")));
                    if (container == "cell")
                    {
                        var paragraph = math.Parent!;
                        paragraph.ReplaceWith(new XElement(w + "tbl",
                            new XElement(w + "tblGrid", new XElement(w + "gridCol", new XAttribute(w + "w", "8640"))),
                            new XElement(w + "tr", new XElement(w + "tc", new XElement(paragraph)))));
                    }
                }, partOverride: container == "header" ? "word/header1.xml" : container == "footer" ? "word/footer1.xml" : null);
        }
    }

    public static void DocxMathHyperlinksMatchIndependentRootLineRectangles()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent root-line rectangles use installed Windows font metrics.");
        string directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var gold = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "math-link-line-office.json")));
        double tolerance = gold.RootElement.GetProperty("TolerancePoints").GetDouble();
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string input = Path.Combine(directory, row.GetProperty("Input").GetString()!);
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            var pages = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .RenderPagesAsync(DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final), null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(row.GetProperty("Pages").GetInt32(), pages.Length);
            string target = row.GetProperty("Target").GetString()!;
            var links = pages.SelectMany((page, index) => page.Annotations.Where(link => link.Uri == target)
                .Select(link => (Page: index + 1, Link: link))).OrderBy(item => item.Page).ThenBy(item => item.Link.X).ToArray();
            var expected = row.GetProperty("Links").EnumerateArray().ToArray();
            TestAssert.Equal(expected.Length, links.Length);
            for (int i = 0; i < links.Length; i++)
            {
                TestAssert.Equal(expected[i].GetProperty("Page").GetInt32(), links[i].Page);
                var rectangle = expected[i].GetProperty("Rect");
                var link = links[i].Link;
                double height = pages[links[i].Page - 1].Height;
                TestAssert.True(Math.Abs(link.X - rectangle[0].GetDouble()) < tolerance &&
                    Math.Abs(height - link.Y - link.Height - rectangle[1].GetDouble()) < tolerance &&
                    Math.Abs(link.X + link.Width - rectangle[2].GetDouble()) < tolerance &&
                    Math.Abs(height - link.Y - rectangle[3].GetDouble()) < tolerance,
                    row.GetProperty("Id").GetString() + ": all four edges retain the independent one-point Word gate.");
            }
        }
    }

    public static void DocxMathHyperlinksPreservePaintAndOwnedTargetsAcrossContainersAndViews()
    {
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        foreach (string fixture in new[] { "docx-omml-isolated-inline-body-fraction-left",
            "CompositionProbes/math-picture-first-table-single-cell",
            "CompositionProbes/math-composition-unnumbered-header",
            "CompositionProbes/math-composition-unnumbered-footer" })
        {
            var renderer = MathContentControlRenderer(mode);
            var before = renderer.RenderPagesAsync(Read(null), null, default).ToBlockingEnumerable().ToArray();
            var linked = new List<PdfLinkAnnotation>();
            foreach (string scope in new[] { "row", "numerator" })
            {
                var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
                var after = renderer.RenderPagesAsync(Read(scope), diagnostics.Add, default).ToBlockingEnumerable().ToArray();
                AssertMathLinkPaint(before, after);
                TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"),
                    "Adding a mathematical hyperlink cannot discard supported content.");
                var baselineAnnotations = before.SelectMany(page => page.Annotations).ToList();
                var annotations = after.SelectMany(page => page.Annotations).ToList();
                foreach (var existing in baselineAnnotations)
                    TestAssert.True(annotations.Remove(existing), "Existing navigation remains unchanged.");
                TestAssert.Equal(1, annotations.Count);
                TestAssert.True(annotations[0].Destination is not null && annotations[0].Width > 0d && annotations[0].Height > 0d,
                    "The actual linked mathematical paint publishes its resolved destination.");
                linked.Add(annotations[0]);
            }
            TestAssert.True(linked[1].Height < linked[0].Height,
                "A numerator link owns its glyphs, not the denominator or fraction bar.");

            DocxDocument Read(string? scope) => ReadMutation(fixture, mode, true, false, (xml, w, m) =>
            {
                if (scope is null) return;
                var equation = xml.Descendants(m + "oMath").First();
                var owner = scope == "row" ? equation : equation.Descendants(m + "num").First();
                var children = owner.Elements().Where(child => !child.Name.LocalName.EndsWith("Pr", StringComparison.Ordinal)).ToArray();
                foreach (var child in children) child.Remove();
                owner.Add(new XElement(w + "hyperlink", new XAttribute(w + "anchor", "PublicCalculation"), children));
            }, partOverride: fixture.EndsWith("-header", StringComparison.Ordinal) ? "word/header1.xml"
                : fixture.EndsWith("-footer", StringComparison.Ordinal) ? "word/footer1.xml" : null);
        }
    }

    public static void DocxMathHyperlinksKeepActualRevisionsAndExcludeDeletedNavigation()
    {
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        foreach (string kind in new[] { "ins", "del" })
        foreach (bool control in new[] { false, true })
        foreach (string placement in new[] { "inside", "outside" })
        {
            var renderer = MathContentControlRenderer(mode);
            var before = renderer.RenderPagesAsync(ReadMathRevisionBoundary(mode, kind, control, "bookmark"), null, default)
                .ToBlockingEnumerable().ToArray();
            var source = ReadMathRevisionBoundary(mode, kind, control, "bookmark", placement, "hyperlink");
            TestAssert.True(DocxInlineContent.HasSupportedParagraphRevisionOwners(source.Paragraphs.First(), default),
                "Navigation retains its actual source revision owner.");
            var after = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            AssertMathLinkPaint(before, after);
            bool included = kind == "ins" ? mode != OoxPdfDocxMarkupMode.Original
                : mode is OoxPdfDocxMarkupMode.Original or OoxPdfDocxMarkupMode.AllMarkup;
            TestAssert.Equal(before.Sum(page => page.Annotations.Count) + (included ? 1 : 0),
                after.Sum(page => page.Annotations.Count));
        }
    }

    public static void DocxMathHyperlinksFollowWrappedPaintAndKeepMissingTargetFallback()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = MathContentControlRenderer(mode);
            var before = renderer.RenderPagesAsync(Read(null), null, default).ToBlockingEnumerable().ToArray();
            var linked = renderer.RenderPagesAsync(Read("PublicCalculation"), null, default).ToBlockingEnumerable().ToArray();
            AssertMathLinkPaint(before, linked);
            TestAssert.True(linked.Sum(page => page.Annotations.Count) > before.Sum(page => page.Annotations.Count) + 1,
                "Each wrapped mathematical fragment retains its source target.");
            var missing = renderer.RenderPagesAsync(Read("MissingMathTarget"), null, default).ToBlockingEnumerable().ToArray();
            AssertMathContentControlPages(before, missing);

            DocxDocument Read(string? target) => ReadMutation("docx-omml-isolated-inline-body-fraction-left", mode, true, false,
                (xml, w, m) =>
                {
                    var equation = xml.Descendants(m + "oMath").First();
                    equation.AddBeforeSelf(new XElement(w + "r", new XElement(w + "t", "Prefix ")));
                    equation.RemoveNodes();
                    var run = new XElement(m + "r", new XElement(m + "t", string.Concat(Enumerable.Repeat("a+b=c+", 36)) + "d"));
                    equation.Add(target is null ? run : new XElement(w + "hyperlink", new XAttribute(w + "anchor", target), run));
                });
        }
    }

    public static void DocxMathHyperlinksIncludeStructuralPaintWithoutAbsorbingOtherRanges()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = MathContentControlRenderer(mode);
            var before = renderer.RenderPagesAsync(Read(null), null, default).ToBlockingEnumerable().Single();
            var control = renderer.RenderPagesAsync(Read("control"), null, default).ToBlockingEnumerable().Single();
            var argument = renderer.RenderPagesAsync(Read("argument"), null, default).ToBlockingEnumerable().Single();
            AssertMathLinkPaint([before], [control]); AssertMathLinkPaint([before], [argument]);
            var wholeLink = control.Annotations.Except(before.Annotations).Single();
            var argumentLink = argument.Annotations.Except(before.Annotations).Single();
            TestAssert.True(wholeLink.X < argumentLink.X && wholeLink.Width > argumentLink.Width &&
                wholeLink.Height >= argumentLink.Height,
                "The radical symbol and overbar belong to the control link;an argument link cannot absorb them.");

            DocxDocument Read(string? scope) => ReadMutation("docx-omml-isolated-inline-body-fraction-left", mode, true, false,
                (xml, w, m) =>
                {
                    var equation = xml.Descendants(m + "oMath").First(); equation.RemoveNodes();
                    var run = new XElement(m + "r", new XElement(m + "t", "a+b"));
                    var argument = new XElement(m + "e", scope == "argument" ? Link(run) : run);
                    var radical = new XElement(m + "rad", new XElement(m + "radPr",
                        new XElement(m + "degHide", new XAttribute(m + "val", "1"))), new XElement(m + "deg"), argument);
                    equation.Add(scope == "control" ? Link(radical) : radical);
                    XElement Link(XElement child) => new(w + "hyperlink", new XAttribute(w + "anchor", "PublicCalculation"), child);
                });
        }
    }

    private static void AssertMathLinkPaint(IReadOnlyList<PdfPage> before, IReadOnlyList<PdfPage> after)
    {
        TestAssert.Equal(before.Count, after.Count);
        for (int index = 0; index < before.Count; index++)
        {
            TestAssert.Equal(before[index].Content, after[index].Content);
            TestAssert.Equal(before[index].Fonts.Count, after[index].Fonts.Count);
            TestAssert.Equal(before[index].Images.Count, after[index].Images.Count);
        }
    }
}
