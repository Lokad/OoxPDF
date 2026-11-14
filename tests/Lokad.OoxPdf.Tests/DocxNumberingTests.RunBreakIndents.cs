using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxNumberingTests
{
    public static void DocxRunBreakFragmentsKeepContinuationInsets()
    {
        foreach (string owner in new[] { "styled", "level-only", "first-line", "hanging" })
        foreach (string kind in new[] { "page", "column" })
        foreach (OoxPdfDocxMarkupMode mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        foreach (OoxPdfDocxMarkupGeometryMode geometry in Enum.GetValues<OoxPdfDocxMarkupGeometryMode>())
        {
            double RelativeX(string breakKind)
            {
                var source = ReadIndentedRunBreak(owner, breakKind, mode);
                var layout = new DocxLayoutEngine(geometry).Create(source, new DocxTests.FamilyWidthTextMeasurer(), default);
                var placed = layout.Pages.SelectMany(page => page.Items.OfType<DocxTextLineLayout>()
                    .Where(line => line.Text == "BETA").Select(line => (Page: page, Line: line))).Single();
                var frame = placed.Page.ColumnFrames[breakKind == "column" ? 1 : 0];
                return placed.Line.X - frame.X;
            }
            double control = RelativeX("soft"), actual = RelativeX(kind);
            TestAssert.True(Math.Abs(control - actual) < .000001d,
                $"A source continuation keeps its inset across {kind}/{owner}/{mode}/{geometry}: expected {control}, got {actual}.");
            var fragments = ReadIndentedRunBreak(owner, kind, mode).Paragraphs;
            var continuation = fragments.Single(p => string.Concat(p.Runs.Select(run => run.Text)) == "BETA");
            TestAssert.True(continuation.ListLabel is null && continuation.Indent.FirstLinePoints is null &&
                continuation.Indent.HangingPoints is null, "Continuation geometry does not paint a repeated marker or repeat first-line indentation.");
        }
    }

    public static void DocxRunBreakContinuationPaintKeepsSourceInsetsAndWords()
    {
        foreach (string owner in new[] { "styled", "level-only", "first-line", "hanging" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            double PaintX(string kind)
            {
                var source = ReadIndentedRunBreak(owner, kind, mode);
                var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
                string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText));
                foreach (string word in new[] { "ALPHA", "BETA", "Expand visual cases", "Document diagnostics" }) TestAssert.Contains(word, text);
                if (owner is "styled" or "level-only")
                    foreach (string label in new[] { "1.", "2.", "3." }) TestAssert.Contains(label, text);
                var beta = pages.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page,
                    usePageCoordinates: true, includeTextTransforms: true)).Single(show => show.Text == "BETA");
                for (int index = 0; index < pages.Length; index++) PdfContentValidator.ValidatePage(pages[index], index, default);
                return beta.X - (kind == "column" ? 252d : 0d);
            }
            double control = PaintX("soft");
            foreach (string kind in new[] { "page", "column" })
                TestAssert.True(Math.Abs(control - PaintX(kind)) < .001d,
                    "Painted continuation words keep the same inset in equal-width frames: " + kind + "/" + owner + "/" + mode);
            using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(ReadIndentedRunBreak(owner, "page", mode),
                null, cancellation.Token).ToBlockingEnumerable().ToArray());
        }
    }

    private static DocxDocument ReadIndentedRunBreak(string owner, string kind, OoxPdfDocxMarkupMode mode) =>
        DocxInlineCompositionTests.ReadMutation("docx-numbering", mode, false, false, (xml, w, _) =>
        {
            var body = xml.Root!.Element(w + "body")!;
            var paragraph = body.Elements(w + "p").First(p => p.Element(w + "pPr")?.Element(w + "numPr") is not null);
            var properties = paragraph.Element(w + "pPr")!;
            if (owner == "level-only") properties.Elements(w + "pStyle").Remove();
            if (owner is "first-line" or "hanging")
            {
                properties.Elements(w + "numPr").Remove();
                properties.Add(new XElement(w + "ind", new XAttribute(w + "left", "720"),
                    new XAttribute(w + (owner == "hanging" ? "hanging" : "firstLine"), "720")));
            }
            paragraph.Elements().Where(child => child.Name != w + "pPr").Remove();
            paragraph.Add(new XElement(w + "r", new XElement(w + "t", "ALPHA"),
                new XElement(w + "br", kind == "soft" ? null : new XAttribute(w + "type", kind)), new XElement(w + "t", "BETA")));
            var section = body.Element(w + "sectPr")!;
            section.Elements(w + "cols").Remove();
            section.Add(new XElement(w + "cols", new XAttribute(w + "num", "2"), new XAttribute(w + "space", "720")));
        });
}
