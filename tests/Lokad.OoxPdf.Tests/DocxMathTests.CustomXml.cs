using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxMathTests
{
    public static void DocxMathCustomXmlKeepsStoredContentAnnotationsFallbacksAndCancellation()
    {
        var xml = EditingMarkerFraction();
        var children = xml.Elements().ToArray(); xml.RemoveNodes();
        var metadata = new XElement(WordprocessingNamespace + "customXml",
            new XElement(WordprocessingNamespace + "customXmlPr",
                new XElement(MathNamespace + "r", new XElement(MathNamespace + "t", "Ghost"))),
            Boundary("commentRangeStart"),
            new XElement(WordprocessingNamespace + "bookmarkStart",
                new XAttribute(WordprocessingNamespace + "id", "7"),
                new XAttribute(WordprocessingNamespace + "name", "WrappedMath")),
            children, Boundary("commentRangeEnd"));
        xml.Add(metadata);
        var comments = new List<string>(); var bookmarks = new List<DocxMathBookmark>();
        var parsed = DocxMathParser.Read(xml, (_, text) => PlainRun(text), default,
            readBookmark: bookmarks.Add, readCommentBoundary: item => comments.Add(item.Name.LocalName));
        TestAssert.Equal("commentRangeStart|commentRangeEnd", string.Join("|", comments));
        TestAssert.Equal("WrappedMath", bookmarks.Single().Name);
        TestAssert.True(DocxMathLayout.TryGetRuns(new(parsed, 0, 0, false, null, null, null), default, out var runs) &&
            string.Concat(runs.Select(item => item.Run.Text.Text)) == "a+bc",
            "Custom-XML properties cannot supply glyphs;stored content keeps annotations and operands.");
        using var cancellation = new CancellationTokenSource();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathParser.Read(xml, (_, text) =>
        { cancellation.Cancel(); return PlainRun(text); }, cancellation.Token));

        foreach (var unsupported in new[]
        {
            new XElement(MathNamespace + "num", new XElement(MathNamespace + "r", new XElement(MathNamespace + "t", "Orphan"))),
            new XElement(XName.Get("future", "urn:future"), new XElement(MathNamespace + "r", new XElement(MathNamespace + "t", "Unknown"))),
            new XElement(WordprocessingNamespace + "ins", new XElement(WordprocessingNamespace + "del",
                new XElement(MathNamespace + "r", new XElement(MathNamespace + "t", "Nested"))))
        })
        {
            var unknown = new XElement(MathNamespace + "oMath", new XElement(WordprocessingNamespace + "customXml", unsupported));
            var result = DocxMathParser.Read(unknown, (_, text) => PlainRun(text), default, includeRevision: _ => true);
            TestAssert.True(!DocxMathLayout.TryGetRuns(new(result, 0, 0, false, null, null, null), default, out _),
                "Metadata cannot create missing argument/revision owners or erase unknown content.");
        }
    }
}
