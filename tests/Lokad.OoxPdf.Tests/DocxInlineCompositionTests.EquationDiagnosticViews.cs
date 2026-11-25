using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxEquationWarningsFollowSelectedNoteRevisionSources()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (bool deleted in new[] { false, true })
        foreach (bool block in new[] { false, true })
        foreach (bool withReaderSink in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.Original, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string part = "word/" + kind + "s.xml";
            var readDiagnostics = new List<OoxPdfDiagnostic>();
            var source = ReadMutation("CompositionProbes/math-note-content-" + kind + "-inline", mode, false, false,
                (xml, w, m) =>
                {
                    int revisionId = 951;
                    var targets = block ? xml.Descendants(w + "p").Where(item => item.Descendants(m + "oMath").Any()).ToArray() :
                        xml.Descendants(m + "oMath").ToArray();
                    foreach (var equation in targets)
                    {
                        var wrapper = new XElement(w + (deleted ? "del" : "ins"), new XAttribute(w + "id", revisionId++),
                            new XAttribute(w + "author", "Public Reviewer"));
                        equation.AddBeforeSelf(wrapper); equation.Remove(); wrapper.Add(equation);
                    }
                }, part, withReaderSink ? readDiagnostics.Add : null);
            // Final retains supported structural paint for inline deleted math;
            // an excluded paragraph block omits the whole equation instead.
            bool hidden = deleted ? block && mode == OoxPdfDocxMarkupMode.Final : mode == OoxPdfDocxMarkupMode.Original;
            var owner = source.EquationParts.SingleOrDefault(item => item.PartName == "/" + part);
            TestAssert.Equal(hidden ? 0 : 2, owner?.SourceCount ?? 0);
            var diagnostics = RenderEquationViewDiagnostics(source, mode);
            var equations = diagnostics.Where(item => item.PartName == "/" + part &&
                item.Id is "DOCX_UNSUPPORTED_EQUATION" or "DOCX_APPROXIMATED_EQUATION").ToArray();
            TestAssert.True(hidden ? equations.Length == 0 : equations.Length != 0,
                $"Selected note diagnostics: {kind}/{mode}/{(deleted ? "del" : "ins")}/block={block}/readerSink={withReaderSink}, " +
                $"selected={owner?.SourceCount}, emitted={string.Join(',', diagnostics.Select(item => item.Id + ":" + item.PartName))}.");
        }
    }

    public static void DocxEquationWarningsRetainVisibleEmptyAndPartialPayloads()
    {
        foreach (bool empty in new[] { false, true })
        {
            const string part = "word/inline-box-header.xml";
            var source = ReadMutation("InlineBoxMath/inline-box-math-header", OoxPdfDocxMarkupMode.Final, false, false,
                (xml, _, m) =>
                {
                    var equation = xml.Descendants(m + "oMath").Single();
                    if (empty) equation.ReplaceWith(new XElement(m + "oMathPara", new XElement(m + "oMathParaPr")));
                    else equation.AddAfterSelf(new XElement(m + "oMath", new XElement(m + "future")));
                }, part);
            TestAssert.Equal(empty ? 1 : 2, source.EquationParts.Single(item => item.PartName == "/" + part).SourceCount);
            var diagnostics = RenderEquationViewDiagnostics(source, OoxPdfDocxMarkupMode.Final);
            TestAssert.True(diagnostics.Any(item => item.Id == "DOCX_UNSUPPORTED_EQUATION" && item.PartName == "/" + part),
                "Visible malformed or unparsed payloads retain their part warning.");
            if (!empty)
                TestAssert.True(diagnostics.Any(item => item.Id == "DOCX_APPROXIMATED_EQUATION" && item.PartName == "/" + part),
                    "Supported neighbors still paint and retain their approximation diagnostic.");
        }
    }

    public static void DocxEquationWarningsIgnoreRemovedLiteralsAndRetainStructure()
    {
        foreach (string shape in new[] { "literal", "display-run", "drawing", "display-picture", "fraction", "visible-unknown" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.Original, OoxPdfDocxMarkupMode.AllMarkup })
        {
            bool final = mode == OoxPdfDocxMarkupMode.Final;
            if (!final && shape is "fraction" or "visible-unknown") continue;
            var source = ReadMutation("CompositionProbes/math-display-image-body-between", mode, false, false, (xml, w, m) =>
            {
                var body = xml.Root!.Element(w + "body")!;
                var section = body.Element(w + "sectPr");
                var drawing = xml.Descendants(w + "drawing").Single();
                var fraction = new XElement(xml.Descendants(m + "f").First());
                var controlSize = new XElement(w + "sz", new XAttribute(w + "val", "22"));
                var control = new XElement(m + "ctrlPr", new XElement(w + "rPr", controlSize));
                fraction.AddFirst(new XElement(m + "fPr", control));
                var math = shape switch
                {
                    "drawing" or "fraction" => new XElement(m + "r", new XElement(m + "t", "314"),
                        new XElement(w + "delText", "QA"), drawing, new XElement(m + "t", "793")),
                    "display-picture" => new XElement(m + "r", new XElement(w + "delText", "QA"), drawing),
                    _ => new XElement(m + "r", new XElement(m + "t", "606"))
                };
                XElement payload = shape is "display-run" or "display-picture" ? new XElement(m + "oMathPara", math) :
                    new XElement(m + "oMath", shape == "fraction" ? fraction : null, math);
                var deletion = new XElement(w + "del", new XAttribute(w + "id", "536"), new XAttribute(w + "author", "Public Reviewer"), payload);
                var visibleUnknown = shape == "visible-unknown" ? new XElement(m + "oMath", new XElement(m + "future")) : null;
                body.ReplaceNodes(new XElement(w + "p", new XElement(w + "r", new XElement(w + "t", "Before")),
                    deletion, visibleUnknown, new XElement(w + "r", new XElement(w + "t", "After"))), section);
            });
            int expectedCount = final ? (shape is "fraction" or "visible-unknown" ? 1 : 0) : shape switch
            {
                "drawing" => 2,
                "display-picture" => 0,
                _ => 1
            };
            TestAssert.Equal(expectedCount, source.EquationParts.Single(item => item.PartName == source.MainPartName).SourceCount);
            if (shape == "fraction")
                TestAssert.True(source.Paragraphs.Single().Equations.Single().Content.Children.Single() is DocxMathFraction
                    { Numerator.Children.Count: 0, Denominator.Children.Count: 0 }, "Final retains only the fraction structure.");
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = PictureRangeRenderer(mode).RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().Single();
            PdfContentValidator.ValidatePage(page, 0, default);
            TestAssert.Equal(shape == "visible-unknown", diagnostics.Any(item => item.Id == "DOCX_UNSUPPORTED_EQUATION" && item.PartName == source.MainPartName));
            TestAssert.Equal(shape == "fraction" || !final && shape != "display-picture", diagnostics.Any(item => item.Id == "DOCX_APPROXIMATED_EQUATION"));
            TestAssert.Equal(!final && shape is "drawing" or "display-picture" ? 1 : 0, page.Images.Count);
            string text = DocxPageFootnoteNumberingTests.DecodeText(page);
            TestAssert.Contains("Before", text);
            TestAssert.Contains("After", text);
            if (final)
                foreach (string label in new[] { "606", "314", "793", "QA" }) TestAssert.DoesNotContain(label, text);
            else if (shape == "display-picture")
                TestAssert.Equal(2, PortableRenderingTests.ReadPaintedSourceLabel(page, "QA").Count);
            else
                TestAssert.Equal(3, PortableRenderingTests.ReadPaintedSourceLabel(page, shape == "drawing" ? "314" : "606", mathOnly: true).Count);
        }
    }

    public static void DocxEquationWarningsFollowSelectedTableRows()
    {
        foreach (bool deleted in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.Original, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadMutation("InlineBoxMath/inline-box-math-cell", mode, false, false, (xml, w, _) =>
            {
                var row = xml.Descendants(w + "tr").Single();
                var properties = row.Element(w + "trPr");
                if (properties is null) { properties = new XElement(w + "trPr"); row.AddFirst(properties); }
                properties.Add(new XElement(w + (deleted ? "del" : "ins"), new XAttribute(w + "id", "952"),
                    new XAttribute(w + "author", "Public Reviewer")));
            });
            bool hidden = deleted ? mode == OoxPdfDocxMarkupMode.Final : mode == OoxPdfDocxMarkupMode.Original;
            TestAssert.Equal(hidden ? 0 : 1, source.EquationParts.Single(item => item.PartName == source.MainPartName).SourceCount);
            var diagnostics = RenderEquationViewDiagnostics(source, mode);
            TestAssert.True(hidden ? diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION") :
                diagnostics.Any(item => item.Id is "DOCX_UNSUPPORTED_EQUATION" or "DOCX_APPROXIMATED_EQUATION"),
                "Source diagnostics use the same selected-row rule as table reading.");
        }
    }

    private static IReadOnlyList<OoxPdfDiagnostic> RenderEquationViewDiagnostics(DocxDocument source, OoxPdfDocxMarkupMode mode)
    {
        var diagnostics = new List<OoxPdfDiagnostic>();
        var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
            mode == OoxPdfDocxMarkupMode.AllMarkup ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup :
                OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var pages = renderer.RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
        TestAssert.True(pages.Length != 0, "The selected view still produces a valid document.");
        foreach (var page in pages) PdfContentValidator.ValidatePage(page, 0, default);
        return diagnostics;
    }
}
