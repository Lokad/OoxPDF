using System.IO.Compression;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxNoteTableInlineBoxesKeepTheirEquationAndProse()
    {
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        string input = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
            "InlineBoxMath", "inline-box-boxed-equation-valid.docx"));
        XElement boxRun;
        using (var zip = ZipFile.OpenRead(input))
        using (var stream = zip.GetEntry("word/document.xml")!.Open())
        {
            boxRun = XDocument.Load(stream).Descendants(w + "txbxContent").Single().Ancestors(w + "r").First();
        }

        var box = boxRun.Descendants(w + "txbxContent").Single();
        box.Descendants(w + "t").First().Value = "BoxInkStart ";
        box.Descendants(w + "t").Last().Value = " BoxInkFinish";
        var proseBoxRun = new XElement(boxRun);
        proseBoxRun.Descendants(m + "oMath").Single().Remove();

        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool withEquation in new[] { true, false })
        {
            string part = "word/" + kind + "s.xml";
            var source = ReadMutation("CompositionProbes/mixed-note-" + kind + "-table", mode, false, false,
                (xml, _, _) =>
                {
                    var note = xml.Root!.Elements(w + kind).Single(item => (string?)item.Attribute(w + "id") == "37");
                    var paragraph = note.Element(w + "tbl")!.Descendants(w + "p").Single();
                    paragraph.Elements().Where(e => e.Name != w + "pPr").Remove();
                    paragraph.Add(Run("SpanBefore"), new XElement(withEquation ? boxRun : proseBoxRun), Run("SpanAfter"));
                }, partOverride: part);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = PictureRangeRenderer(mode).RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable();
            PdfPage page = pages.Single(p => DocxCommentsTests.ReadEmbeddedGlyphTextShows(p, usePageCoordinates: true,
                includeTextTransforms: true).Any(show => show.Text.Contains("BoxInkStart", StringComparison.Ordinal)));
            AssertSourceLabelsPaint(page, mode, ["SpanBefore", "BoxInkStart", "BoxInkFinish", "SpanAfter"],
                kind + "/note-table-inline-box");
            if (withEquation)
                TestAssert.Equal(4, PortableRenderingTests.ReadPaintedSourceLabel(page, "x+2y", mathOnly: true).Count);
            TestAssert.Equal(withEquation ? 1 : 0,
                diagnostics.Count(d => d.PartName == "/" + part && d.Id == "DOCX_APPROXIMATED_EQUATION"));
            TestAssert.Equal(0, diagnostics.Count(d => d.PartName == "/" + part && d.Id == "DOCX_UNSUPPORTED_EQUATION"));
        }

        XElement Run(string text) => new(w + "r", new XElement(w + "t", text));
    }
}
