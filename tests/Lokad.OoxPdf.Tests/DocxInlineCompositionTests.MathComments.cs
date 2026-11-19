using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxMathCommentsRetainTheContainingInlineSourceFrame()
    {
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        foreach (string scope in new[] { "outer", "math-row", "numerator", "cross-row" })
        {
            var document = ReadMathCommentRange(mode, scope);
            var paragraph = document.Paragraphs.First(item => item.CommentRanges.Count != 0);
            var range = paragraph.CommentRanges.Single();
            TestAssert.Equal("0", range.Id);
            TestAssert.Equal(scope == "numerator", range.InFractionNumerator);
            TestAssert.True(range.StartInlineItemIndex is int start && range.EndInlineItemIndex == start + 1 &&
                paragraph.InlineContent![start] is DocxInlineEquation,
                "A mathematical range retains its actual containing inline equation.");
            TestAssert.True(range.ReferenceSourceRunIndex is not null &&
                paragraph.InlineReferences.Count(item => item.Kind == DocxRelatedStoryKind.Comment && item.Id == "0") == 1,
                "The existing comment reference remains associated with that source range.");
            TestAssert.True(DocxMathLayout.TryGetRuns(paragraph.Equations.Single(), default, out var runs) &&
                runs.Any(item => item.Run.Text.Text == "x+2"),
                "Empty comment boundaries cannot turn a supported formula into an unsupported operand.");
        }
    }

    public static void DocxMathCommentsPreservePaintAcrossViewsAndNestedBoundaries()
    {
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.AllMarkup ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                    : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var direct = renderer.RenderPagesAsync(ReadMathCommentRange(mode, "outer"), null, default)
                .ToBlockingEnumerable().ToArray();
            foreach (string scope in new[] { "math-row", "numerator", "cross-row" })
            {
                var marked = renderer.RenderPagesAsync(ReadMathCommentRange(mode, scope), null, default)
                    .ToBlockingEnumerable().ToArray();
                TestAssert.Equal(direct.Length, marked.Length);
                for (int page = 0; page < direct.Length; page++)
                {
                    if (mode == OoxPdfDocxMarkupMode.AllMarkup && scope == "numerator" && page == 0)
                        AssertNumeratorCommentPoint(direct[page], marked[page]);
                    else
                        TestAssert.Equal(direct[page].Content, marked[page].Content);
                    TestAssert.Equal(direct[page].Fonts.Count, marked[page].Fonts.Count);
                    TestAssert.Equal(direct[page].Images.Count, marked[page].Images.Count);
                    TestAssert.Equal(direct[page].Annotations.Count, marked[page].Annotations.Count);
                }
            }
        }
    }

    private static void AssertNumeratorCommentPoint(PdfPage outer, PdfPage numerator)
    {
        const string fills = @"0\.973 0\.863 0\.867 rg\s+(?<x>[-\d.]+) [-\d.]+ [-\d.]+ [-\d.]+ re f";
        double leadingEdge = Regex.Matches(outer.Content, fills).Select(match => Coordinate(match, "x")).Single(x => x < 400d);
        TestAssert.True(Regex.Matches(numerator.Content, fills).All(match => Coordinate(match, "x") >= 400d),
            "A numerator-only comment prints brackets and a connector without a body selection fill.");

        var strokes = Regex.Matches(numerator.Content, @"(?m)^(?<x>[-\d.]+) (?<y>[-\d.]+) m (?<endX>[-\d.]+) (?<endY>[-\d.]+) l S\r?$")
            .Select(match => (X: Coordinate(match, "x"), Y: Coordinate(match, "y"),
                EndX: Coordinate(match, "endX"), EndY: Coordinate(match, "endY"))).ToArray();
        var brackets = strokes.Where(s => s.X == s.EndX && Math.Abs(s.X - leadingEdge) < .25d).ToArray();
        TestAssert.Equal(2, brackets.Length);
        TestAssert.True(brackets.Min(s => s.X) < leadingEdge && brackets.Max(s => s.X) > leadingEdge,
            "The two brackets enclose the measured leading edge of the fraction.");
        TestAssert.True(strokes.Any(s => Math.Abs(s.X - leadingEdge) < .002d && s.Y == s.EndY && s.EndX > 400d),
            "The connector starts at the same point as the brackets.");

        var outerInk = DocxCommentsTests.ReadEmbeddedGlyphTextShows(outer).Select(show => (show.Text, show.Size, show.X, show.Y)).ToArray();
        var pointInk = DocxCommentsTests.ReadEmbeddedGlyphTextShows(numerator).Select(show => (show.Text, show.Size, show.X, show.Y)).ToArray();
        TestAssert.True(outerInk.SequenceEqual(pointInk), "The formula, surrounding prose and comment text retain their paint and positions.");

        static double Coordinate(Match match, string name) => double.Parse(match.Groups[name].Value, CultureInfo.InvariantCulture);
    }

    public static void DocxMathCommentsRetainExistingCellStoryPlacement()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.AllMarkup ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                    : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var direct = renderer.RenderPagesAsync(Read(false), null, default).ToBlockingEnumerable().ToArray();
            var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
            var inside = renderer.RenderPagesAsync(Read(true), diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "A supported comment/formula combination survives its existing cell owner.");
            TestAssert.Equal(direct.Length, inside.Length);
            for (int page = 0; page < direct.Length; page++)
            {
                TestAssert.Equal(direct[page].Content, inside[page].Content);
                TestAssert.Equal(direct[page].Fonts.Count, inside[page].Fonts.Count);
                TestAssert.Equal(direct[page].Annotations.Count, inside[page].Annotations.Count);
            }

            DocxDocument Read(bool inner) => ReadMutation("CompositionProbes/math-comment-formula-cell", mode,
                false, false, (xml, w, m) =>
                {
                    if (!inner) return;
                    var start = xml.Descendants(w + "commentRangeStart").Single();
                    var end = xml.Descendants(w + "commentRangeEnd").Single();
                    start.Remove(); end.Remove();
                    var math = xml.Descendants(m + "oMath").First();
                    math.AddFirst(start); math.Add(end);
                });
        }
    }

    private static DocxDocument ReadMathCommentRange(OoxPdfDocxMarkupMode mode, string scope) =>
        ReadMutation("Composition/report", mode, false, false, (xml, w, m) =>
        {
            var start = xml.Descendants(w + "commentRangeStart").Single();
            var end = xml.Descendants(w + "commentRangeEnd").Single();
            var reference = xml.Descendants(w + "commentReference").Single().Parent!;
            start.Remove(); end.Remove(); reference.Remove();
            var math = xml.Descendants(m + "oMath").First();
            math.AddAfterSelf(reference);
            if (scope == "outer") { math.AddBeforeSelf(start); math.AddAfterSelf(end); }
            else
            {
                var numerator = math.Descendants(m + "num").Single();
                (scope == "numerator" ? numerator : math).AddFirst(start);
                (scope == "math-row" ? math : numerator).Add(end);
            }
        });
}
