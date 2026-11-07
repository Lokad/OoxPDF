using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
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
