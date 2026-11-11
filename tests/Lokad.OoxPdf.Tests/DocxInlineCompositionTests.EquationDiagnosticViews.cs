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
        foreach (bool withReaderSink in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.Original, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string part = "word/" + kind + "s.xml";
            var readDiagnostics = new List<OoxPdfDiagnostic>();
            var source = ReadMutation("CompositionProbes/math-note-content-" + kind + "-inline", mode, false, false,
                (xml, w, m) =>
                {
                    int revisionId = 951;
                    foreach (var equation in xml.Descendants(m + "oMath").ToArray())
                    {
                        var wrapper = new XElement(w + (deleted ? "del" : "ins"), new XAttribute(w + "id", revisionId++),
                            new XAttribute(w + "author", "Public Reviewer"));
                        equation.AddBeforeSelf(wrapper); equation.Remove(); wrapper.Add(equation);
                    }
                }, part, withReaderSink ? readDiagnostics.Add : null);
            bool hidden = deleted ? mode == OoxPdfDocxMarkupMode.Final : mode == OoxPdfDocxMarkupMode.Original;
            var owner = source.EquationParts.SingleOrDefault(item => item.PartName == "/" + part);
            TestAssert.Equal(hidden ? 0 : 2, owner?.SourceCount ?? 0);
            var diagnostics = RenderEquationViewDiagnostics(source, mode);
            var equations = diagnostics.Where(item => item.PartName == "/" + part &&
                item.Id is "DOCX_UNSUPPORTED_EQUATION" or "DOCX_APPROXIMATED_EQUATION").ToArray();
            TestAssert.True(hidden ? equations.Length == 0 : equations.Length != 0,
                "Selected note sources determine equation diagnostics with or without an initial reader sink.");
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
