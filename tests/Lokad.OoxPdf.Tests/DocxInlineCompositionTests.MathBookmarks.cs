using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInnerMathBookmarksRetainActualSourceBoundaries()
    {
        foreach (string scope in new[] { "outer", "math-row", "numerator" })
        {
            var paragraph = ReadMutation("CompositionProbes/math-inner-bookmark-" + scope,
                OoxPdfDocxMarkupMode.Final, false, false).Paragraphs.First();
            var equation = paragraph.Equations.Single();
            var bookmark = paragraph.BookmarkAnchors.Single();
            TestAssert.Equal("MathSource", bookmark.Name);
            TestAssert.Equal(scope == "outer", bookmark.MathSource is null);
            if (scope != "outer")
            {
                var owner = scope == "math-row" ? equation.Content
                    : ((DocxMathFraction)equation.Content.Children.Single()).Numerator;
                TestAssert.True(ReferenceEquals(bookmark.MathSource, owner.Bookmarks.Single().Source),
                    "Navigation and mathematical measurement share the actual source boundary.");
                TestAssert.Equal(0, owner.Bookmarks.Single().ChildIndex);
            }
            TestAssert.True(DocxMathLayout.TryGetRuns(equation, default, out var runs) && runs.Count == 4,
                "Adding a bookmark cannot remove the supported numerator or denominator.");
        }
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            var document = ReadMutation("CompositionProbes/equation-revision-body-del", mode, false, false,
                (xml, w, m) => xml.Descendants(m + "num").First().AddFirst(new XElement(w + "bookmarkStart",
                    new XAttribute(w + "id", "18732"), new XAttribute(w + "name", "DeletedMathTarget"))));
            TestAssert.Equal(mode is OoxPdfDocxMarkupMode.Final or OoxPdfDocxMarkupMode.SimpleMarkup ? 0 : 1,
                document.Paragraphs.First().BookmarkAnchors.Count);
        }
    }

    public static void DocxInnerMathBookmarksKeepPortablePaintAndInternalLinks()
    {
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.AllMarkup ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                    : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var direct = renderer.RenderPagesAsync(ReadMutation("CompositionProbes/math-inner-bookmark-outer",
                mode, false, false), null, default).ToBlockingEnumerable().Single();
            foreach (string scope in new[] { "math-row", "numerator" })
            {
                var page = renderer.RenderPagesAsync(ReadMutation("CompositionProbes/math-inner-bookmark-" + scope,
                    mode, false, false), null, default).ToBlockingEnumerable().Single();
                TestAssert.Equal(direct.Content, page.Content);
                TestAssert.Equal(direct.Fonts.Count, page.Fonts.Count);
                var link = page.Annotations.Single();
                TestAssert.Equal(0, link.Destination!.Value.PageIndex);
                TestAssert.True(double.IsFinite(link.Destination.Value.Left!.Value) &&
                    double.IsFinite(link.Destination.Value.Top!.Value), "The target is a finite measured source point.");
            }
        }
    }

    public static void DocxInnerMathBookmarksMatchIndependentWordViewports()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent viewport geometry uses installed Windows math fonts.");
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var source = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "math-inner-bookmarks-office.json")));
        double tolerance = source.RootElement.GetProperty("ViewportTolerancePt").GetDouble();
        foreach (var row in source.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string path = Path.Combine(folder, row.GetProperty("Fixture").GetString()! + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            var page = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .RenderPagesAsync(DocxTests.ReadDocx(path, OoxPdfDocxMarkupMode.Final), null, default).ToBlockingEnumerable().Single();
            var target = page.Annotations.Single().Destination!.Value;
            TestAssert.Equal(row.GetProperty("PageIndex").GetInt32(), target.PageIndex);
            TestAssert.True(Math.Abs(target.Left!.Value - row.GetProperty("Left").GetDouble()) < tolerance &&
                Math.Abs(page.Height - target.Top!.Value - row.GetProperty("TopFromPage").GetDouble()) < tolerance,
                "The destination uses the independently measured row/argument viewport: " + row.GetProperty("Fixture").GetString() +
                " at " + target.Left + "/" + (page.Height - target.Top));
        }
    }

    public static void DocxInnerMathBookmarksSurviveCompositeBoxTranslations()
    {
        const string run = "<m:r><m:t>x</m:t></m:r>";
        foreach (var shape in new[]
        {
            (Xml: $"<m:f><m:num>{run}</m:num><m:den>{run}</m:den></m:f>", Role: "den"),
            (Xml: $"<m:f><m:fPr><m:type m:val='skw'/></m:fPr><m:num>{run}</m:num><m:den>{run}</m:den></m:f>", Role: "den"),
            (Xml: $"<m:sSup><m:e>{run}</m:e><m:sup>{run}</m:sup></m:sSup>", Role: "sup"),
            (Xml: $"<m:func><m:fName><m:r><m:rPr><m:nor/></m:rPr><m:t>sin</m:t></m:r></m:fName><m:e>{run}</m:e></m:func>", Role: "e"),
            (Xml: $"<m:rad><m:deg>{run}</m:deg><m:e>{run}</m:e></m:rad>", Role: "deg"),
            (Xml: $"<m:d><m:e>{run}</m:e></m:d>", Role: "e"),
            (Xml: $"<m:nary><m:sub>{run}</m:sub><m:sup>{run}</m:sup><m:e>{run}</m:e></m:nary>", Role: "sup"),
            (Xml: $"<m:m><m:mr><m:e>{run}</m:e><m:e>{run}</m:e></m:mr><m:mr><m:e>{run}</m:e><m:e>{run}</m:e></m:mr></m:m>", Role: "e")
        })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true),
                OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            string? content = null;
            foreach (bool inner in new[] { false, true })
            {
                var document = ReadMutation("CompositionProbes/math-inner-bookmark-outer", OoxPdfDocxMarkupMode.Final,
                    false, false, (xml, w, m) =>
                    {
                        var old = xml.Descendants(m + "oMath").Single();
                        var math = XElement.Parse($"<m:oMath xmlns:m='{m}'>{shape.Xml}</m:oMath>");
                        old.ReplaceWith(math);
                        if (!inner) return;
                        var start = xml.Descendants(w + "bookmarkStart").Single();
                        var end = xml.Descendants(w + "bookmarkEnd").Single();
                        start.Remove(); end.Remove();
                        var role = math.Descendants(m + shape.Role).Last();
                        role.AddFirst(start); role.Add(end);
                    });
                var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
                var page = renderer.RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().Single();
                TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"),
                    "The translation control must render its mathematical structure: " + shape.Xml);
                TestAssert.Equal(0, page.Annotations.Single().Destination!.Value.PageIndex);
                if (inner) TestAssert.Equal(content, page.Content);
                else content = page.Content;
            }
        }
    }
}
