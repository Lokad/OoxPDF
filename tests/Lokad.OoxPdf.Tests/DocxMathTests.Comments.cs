using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxMathTests
{
    public static void DocxMathCommentBoundariesFollowSourceOrderAndRevisionFiltering()
    {
        var xml = EditingMarkerFraction();
        xml.AddFirst(Boundary("commentRangeStart"));
        xml.Descendants(MathNamespace + "num").Single().Add(Boundary("commentRangeEnd"));
        var events = new List<string>();
        Read(xml);
        TestAssert.Equal("commentRangeStart|commentRangeEnd", string.Join("|", events));
        foreach (string kind in new[] { "ins", "del" })
        foreach (bool included in new[] { false, true })
        {
            events.Clear();
            var wrapper = new XElement(WordprocessingNamespace + kind, xml.Elements().Select(item => new XElement(item)));
            var revised = new XElement(MathNamespace + "oMath", wrapper);
            DocxMathParser.Read(revised, (_, text) => PlainRun(text), default,
                includeRevision: _ => included, preserveDeletedStructure: true,
                readCommentBoundary: item => events.Add(item.Name.LocalName));
            TestAssert.Equal(included ? 2 : 0, events.Count);
        }

        void Read(XElement source) => DocxMathParser.Read(source, (_, text) => PlainRun(text), default,
            readCommentBoundary: item => events.Add(item.Name.LocalName));
    }

    public static void DocxMathCommentBoundariesKeepUnknownContentAndCancellation()
    {
        foreach (var marker in new[] { new XElement(XName.Get("commentRangeStart", "urn:future")),
            new XElement(WordprocessingNamespace + "commentRangeStart", "Visible") })
        {
            var xml = EditingMarkerFraction(); xml.AddFirst(marker);
            TestAssert.True(Parse(xml).Children.First() is DocxMathUnsupported,
                "Foreign or content-bearing annotation elements keep the unsupported-content path.");
        }
        var source = EditingMarkerFraction();
        source.AddFirst(Boundary("commentRangeStart")); source.Add(Boundary("commentRangeEnd"));
        using var cancellation = new CancellationTokenSource();
        int events = 0;
        var failure = TestAssert.Throws<OperationCanceledException>(() => DocxMathParser.Read(source,
            (_, text) => PlainRun(text), cancellation.Token, readCommentBoundary: _ =>
            { events++; cancellation.Cancel(); }));
        TestAssert.Equal(1, events);
        TestAssert.Equal(cancellation.Token, failure.CancellationToken);
    }

    private static XElement Boundary(string name) => new(WordprocessingNamespace + name,
        new XAttribute(WordprocessingNamespace + "id", "0"));
}
