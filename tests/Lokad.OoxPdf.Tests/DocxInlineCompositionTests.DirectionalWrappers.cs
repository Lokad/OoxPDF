using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxDirectionalWrappersRetainPictureTextAndMath()
    {
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (var (place, mode, element, direction) in new[]
        {
            ("body", OoxPdfDocxMarkupMode.Final, "dir", "ltr"),
            ("cell", OoxPdfDocxMarkupMode.AllMarkup, "bdo", "rtl"),
            ("footer", OoxPdfDocxMarkupMode.AllMarkup, "dir", "rtl")
        })
        {
            string story = place == "body" ? "" : "-" + place;
            string fixture = "PortableFonts/math-composition-unnumbered" + story + "-normal";
            string part = place == "footer" ? "word/footer1.xml" : "word/document.xml";
            var control = new List<OoxPdfDiagnostic>();
            var candidate = new List<OoxPdfDiagnostic>();
            var direct = ReadMutation(fixture, mode, false, false, partOverride: part, diagnosticSink: control.Add);
            var wrapped = ReadMutation(fixture, mode, false, false, (xml, w, m) =>
            {
                XElement drawingRun = xml.Descendants(w + "drawing").Single().Parent!;
                XElement paragraph = drawingRun.Parent!;
                XElement formula = paragraph.Element(m + "oMath")!;
                XElement first = drawingRun.ElementsBeforeSelf().Last();
                XElement[] children = first.ElementsAfterSelf().Prepend(first)
                    .TakeWhile(child => child != formula).Append(formula).ToArray();
                var wrapper = new XElement(w + element, new XAttribute(w + "val", direction));
                first.AddBeforeSelf(wrapper);
                foreach (XElement child in children)
                {
                    child.Remove();
                    wrapper.Add(child);
                }
            }, partOverride: part, diagnosticSink: candidate.Add);
            var renderer = new DocxRenderer(resolver, mode, CachedMathGeometry(mode));
            var expectedPage = renderer.RenderPagesAsync(direct, control.Add, default).ToBlockingEnumerable().Single();
            var actualPage = renderer.RenderPagesAsync(wrapped, candidate.Add, default).ToBlockingEnumerable().Single();
            TestAssert.Contains("trend marker", DocxPageFootnoteNumberingTests.DecodeText(actualPage));
            TestAssert.Contains("a+b=c", ReadEquationText([actualPage]));
            TestAssert.Equal(1, actualPage.Images.Count);
            TestAssert.Equal(expectedPage.Content, actualPage.Content);
            var warning = candidate.Single(item => item.Id == "DOCX_INLINE_BIDI_APPROXIMATION");
            TestAssert.Equal(OoxPdfSeverity.Warning, warning.Severity);
            TestAssert.Equal("Source order", warning.Fallback);
            TestAssert.Equal("/" + part, warning.PartName);
            TestAssert.True(control.SequenceEqual(candidate.Where(item => item != warning)),
                "The directional fallback must preserve the other diagnostic outcomes.");
        }
    }

    public static void DocxDirectionalWrappersKeepHardBreaksAndExcludedRevisions()
    {
        const OoxPdfDocxMarkupMode mode = OoxPdfDocxMarkupMode.Final;
        const string fixture = "docx-ladder-01-plain-paragraph";
        var warnings = new List<OoxPdfDiagnostic>();
        var direct = ReadMutation(fixture, mode, false, false, (xml, w, _) => Author(xml, w, false));
        var wrapped = ReadMutation(fixture, mode, false, false, (xml, w, _) => Author(xml, w, true),
            diagnosticSink: warnings.Add);
        var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode, CachedMathGeometry(mode));
        var expected = renderer.RenderPagesAsync(direct, null, default).ToBlockingEnumerable().ToArray();
        var actual = renderer.RenderPagesAsync(wrapped, null, default).ToBlockingEnumerable().ToArray();
        TestAssert.Equal(2, actual.Length);
        TestAssert.True(expected.Select(page => page.Content).SequenceEqual(actual.Select(page => page.Content)),
            "The directional wrapper must keep both sides of the authored page break.");
        string before = DocxPageFootnoteNumberingTests.DecodeText(actual[0]);
        string after = DocxPageFootnoteNumberingTests.DecodeText(actual[1]);
        TestAssert.Contains("Before directional break", before);
        TestAssert.Contains("After directional break", after);
        TestAssert.True(!after.Contains("Excluded direction", StringComparison.Ordinal),
            "An excluded deletion stays invisible through its directional wrapper.");
        TestAssert.Equal(1, warnings.Count(item => item.Id == "DOCX_INLINE_BIDI_APPROXIMATION"));

        warnings.Clear();
        _ = ReadMutation(fixture, mode, false, false, (xml, w, _) => Author(xml, w, false),
            diagnosticSink: warnings.Add);
        TestAssert.True(warnings.All(item => item.Id != "DOCX_INLINE_BIDI_APPROXIMATION"),
            "An excluded deletion alone must not trigger a directional fallback warning.");

        static void Author(XDocument xml, XNamespace w, bool wrap)
        {
            var paragraph = xml.Descendants(w + "body").Single().Elements(w + "p").Single();
            var content = new[]
            {
                new XElement(w + "r", new XElement(w + "t", "Before directional break")),
                new XElement(w + "r", new XElement(w + "br", new XAttribute(w + "type", "page"))),
                new XElement(w + "r", new XElement(w + "t", "After directional break"))
            };
            paragraph.ReplaceNodes(wrap ? [new XElement(w + "dir", new XAttribute(w + "val", "ltr"), content)] : content);
            paragraph.Add(new XElement(w + "del", new XAttribute(w + "id", 816), new XAttribute(w + "author", "Author"),
                new XElement(w + "bdo", new XAttribute(w + "val", "rtl"),
                    new XElement(w + "r", new XElement(w + "delText", "Excluded direction")))));
        }
    }
}
