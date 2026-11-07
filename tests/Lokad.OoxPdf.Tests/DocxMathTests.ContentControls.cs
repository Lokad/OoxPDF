using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxMathTests
{
    public static void DocxMathContentControlsKeepStoredContentAnnotationsAndCancellation()
    {
        var xml = EditingMarkerFraction();
        var children = xml.Elements().ToArray(); xml.RemoveNodes();
        var content = new XElement(WordprocessingNamespace + "sdtContent", Boundary("commentRangeStart"),
            new XElement(WordprocessingNamespace + "bookmarkStart", new XAttribute(WordprocessingNamespace + "id", "7"),
                new XAttribute(WordprocessingNamespace + "name", "ControlledMath")), children,
            Boundary("commentRangeEnd"));
        var metadata = new XElement(WordprocessingNamespace + "sdt",
            new XElement(WordprocessingNamespace + "sdtPr", new XElement(MathNamespace + "r", new XElement(MathNamespace + "t", "Ghost"))), content);
        xml.Add(metadata);
        var comments = new List<string>(); var bookmarks = new List<DocxMathBookmark>();
        var parsed = DocxMathParser.Read(xml, (_, text) => PlainRun(text), default,
            readBookmark: bookmarks.Add, readCommentBoundary: item => comments.Add(item.Name.LocalName));
        TestAssert.Equal("commentRangeStart|commentRangeEnd", string.Join("|", comments));
        TestAssert.Equal("ControlledMath", bookmarks.Single().Name);
        TestAssert.True(DocxMathLayout.TryGetRuns(new(parsed, 0, 0, false, null, null, null), default, out var runs) &&
            string.Concat(runs.Select(item => item.Run.Text.Text)) == "a+bc",
            "Only sdtContent supplies stored math;control properties cannot inject glyphs.");
        using var cancellation = new CancellationTokenSource();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathParser.Read(xml, (_, text) =>
        { cancellation.Cancel(); return PlainRun(text); }, cancellation.Token));
    }

    public static void DocxMathContentControlsRetainOrphanUnknownAndNestedRevisionFallbacks()
    {
        foreach (var child in new[]
        {
            new XElement(MathNamespace + "num", new XElement(MathNamespace + "r", new XElement(MathNamespace + "t", "Orphan"))),
            new XElement(XName.Get("future", "urn:future"), new XElement(MathNamespace + "r", new XElement(MathNamespace + "t", "Future"))),
            new XElement(WordprocessingNamespace + "ins", new XElement(WordprocessingNamespace + "del",
                new XElement(MathNamespace + "r", new XElement(MathNamespace + "t", "Nested"))))
        })
        {
            var xml = new XElement(MathNamespace + "oMath", new XElement(WordprocessingNamespace + "sdt",
                new XElement(WordprocessingNamespace + "sdtContent", child)));
            var parsed = DocxMathParser.Read(xml, (_, text) => PlainRun(text), default, includeRevision: _ => true);
            TestAssert.True(!DocxMathLayout.TryGetRuns(new(parsed, 0, 0, false, null, null, null), default, out _),
                "Content controls preserve stored content without creating missing role/revision owners.");
        }
        var orphan = new XElement(MathNamespace + "oMath", new XElement(WordprocessingNamespace + "sdtContent",
            new XElement(MathNamespace + "r", new XElement(MathNamespace + "t", "Orphan"))));
        TestAssert.True(!DocxMathLayout.TryGetRuns(new(Parse(orphan), 0, 0, false, null, null, null), default, out _),
            "sdtContent needs its actual content-control parent.");
    }
}
