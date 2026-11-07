using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

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
                    TestAssert.Equal(direct[page].Content, marked[page].Content);
                    TestAssert.Equal(direct[page].Fonts.Count, marked[page].Fonts.Count);
                    TestAssert.Equal(direct[page].Images.Count, marked[page].Images.Count);
                    TestAssert.Equal(direct[page].Annotations.Count, marked[page].Annotations.Count);
                }
            }
        }
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
