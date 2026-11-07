using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxMathTests
{
    public static void DocxMathHyperlinksKeepSourceOwnersFallbacksAndCancellation()
    {
        var equation = EditingMarkerFraction();
        var children = equation.Elements().ToArray(); equation.RemoveNodes();
        equation.Add(new XElement(WordprocessingNamespace + "hyperlink", children));
        var link = new DocxHyperlinkSpan("rMath", null, null, null, "https://example.org/math", "External", null,
            0, 0, 0, 0, 0) { IsMathContent = true, StartInlineItemIndex = 0, EndInlineItemIndex = 1 };
        int sources = 0;
        var parsed = DocxMathParser.Read(equation, (_, text) => PlainRun(text), default,
            readHyperlink: _ => { sources++; return link; });
        TestAssert.Equal(1, sources);
        TestAssert.True(DocxMathLayout.TryGetRuns(new(parsed, 0, 0, false, null, null, null), default, out var runs) &&
            runs.All(item => ReferenceEquals(item.Run.Hyperlink, link)) &&
            DocxMathLayout.GetControlSources(new(parsed, 0, 0, false, null, null, null), default)
                .All(item => ReferenceEquals(((DocxMathNode)item).Hyperlink, link)),
            "Literal and structural paint retain the actual resolved source hyperlink.");
        using var cancellation = new CancellationTokenSource();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathParser.Read(equation, (_, text) => PlainRun(text),
            cancellation.Token, readHyperlink: _ => { cancellation.Cancel(); return link; }));
        foreach (var unknown in new[]
        {
            new XElement(MathNamespace + "num", new XElement(MathNamespace + "r", new XElement(MathNamespace + "t", "Orphan"))),
            new XElement(XName.Get("future", "urn:future"), new XElement(MathNamespace + "r", new XElement(MathNamespace + "t", "Unknown"))),
            new XElement(WordprocessingNamespace + "ins", new XElement(WordprocessingNamespace + "del",
                new XElement(MathNamespace + "r", new XElement(MathNamespace + "t", "Nested"))))
        })
        {
            var xml = new XElement(MathNamespace + "oMath", new XElement(WordprocessingNamespace + "hyperlink", unknown));
            var result = DocxMathParser.Read(xml, (_, text) => PlainRun(text), default,
                includeRevision: _ => true, readHyperlink: _ => link);
            TestAssert.True(!DocxMathLayout.TryGetRuns(new(result, 0, 0, false, null, null, null), default, out _),
                "Navigation cannot create argument/revision owners or hide unknown mathematical content.");
        }
    }
}
