using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxContainerPropertyOwnersShareThePrintedPageLane()
    {
        var context = DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup,
            OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        foreach (string container in new[] { "body", "header", "footer", "floating-table" })
        {
            var control = ReadContainerPropertySource(container, OoxPdfDocxMarkupMode.AllMarkup, history: false);
            var source = ReadContainerPropertySource(container, OoxPdfDocxMarkupMode.AllMarkup, history: true);
            TestAssert.Equal(1d, DocxRenderer.ResolveWordCompatiblePrintScale(control, context));
            TestAssert.True(DocxRenderer.ResolveWordCompatiblePrintScale(source, context) < 1d,
                container + ": a printed child formatting caption must reserve its page lane without a competing comment.");
            var empty = ReadContainerPropertySource(container, OoxPdfDocxMarkupMode.AllMarkup, history: true, empty: true);
            TestAssert.Equal(1d, DocxRenderer.ResolveWordCompatiblePrintScale(empty, context));
            TestAssert.Equal(1d, DocxRenderer.ResolveWordCompatiblePrintScale(source,
                DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.Final)));
            var wideMargin = source with { MarginRightPoints = 288d };
            TestAssert.Equal(1d, DocxRenderer.ResolveWordCompatiblePrintScale(wideMargin, context));
        }
    }

    public static void DocxContainerPropertyCaptionsKeepVisibleSourceContent()
    {
        foreach (string container in new[] { "body", "header", "footer", "floating-table" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode, CachedMathGeometry(mode));
            var control = renderer.RenderPagesAsync(ReadContainerPropertySource(container, mode, history: false),
                null, default).ToBlockingEnumerable().Single();
            var source = ReadContainerPropertySource(container, mode, history: true);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = renderer.RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().Single();
            string text = DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC);
            foreach (string label in container switch
            {
                "body" => new[] { "Inline box content." },
                "floating-table" => new[] { "BOX TABLE NORTH", "BOX TABLE SOUTH" },
                _ => new[] { "Before inline", "after equation." }
            }) TestAssert.Contains(label, text);
            string MathText(Lokad.OoxPdf.Pdf.PdfPage value) => string.Concat(
                DocxCommentsTests.ReadEmbeddedGlyphTextShows(value, true).Where(show => value.Fonts.Any(font =>
                    font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(font.Font, show.Font)))
                    .Select(show => show.Text));
            TestAssert.Equal(MathText(control), MathText(page));
            TestAssert.Equal(container == "floating-table", MathText(page).Length == 0);
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Supported formulas remain painted.");
            if (mode == OoxPdfDocxMarkupMode.Final) TestAssert.Equal(control.Content, page.Content);
            else
            {
                TestAssert.Equal(1, renderer.InspectMarkupBalloons(source).Count);
                TestAssert.Contains(container == "floating-table" ? "row height" : "Not Bold", text);
            }
            using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(source, null, cancellation.Token)
                .ToBlockingEnumerable().ToArray());
        }
    }

    public static void DocxNestedBoxPropertyHistoryDoesNotReserveAnUnpaintedLane()
    {
        var source = ReadMutation("InlineBoxMath/inline-box-body-equation-inline-box-valid",
            OoxPdfDocxMarkupMode.AllMarkup, false, false, (xml, w, _) =>
            {
                var outer = xml.Descendants(w + "txbxContent").Single();
                var drawing = new XElement(outer.Ancestors(w + "drawing").First());
                var run = new XElement(w + "r", new XElement(w + "rPr", ContainerPropertyChange(w, "rPr", empty: false)),
                    new XElement(w + "t", "Unpainted nested label."));
                drawing.Descendants(w + "txbxContent").Single().ReplaceNodes(new XElement(w + "p", run));
                outer.Elements(w + "p").First().Add(new XElement(w + "r", drawing));
            });
        var context = DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup,
            OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        TestAssert.Equal(1d, DocxRenderer.ResolveWordCompatiblePrintScale(source, context));
        var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true),
            OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        TestAssert.Equal(0, renderer.InspectMarkupBalloons(source).Count);
        var diagnostics = new List<OoxPdfDiagnostic>();
        var page = renderer.RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().Single();
        string text = DocxPageFootnoteNumberingTests.DecodeText(page);
        TestAssert.Contains("Inline box content.", text);
        TestAssert.True(!text.Contains("Unpainted nested label.", StringComparison.Ordinal), "The unsupported nested story stays ignored.");
        TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_NESTED_INLINE_TEXTBOX"), "The existing fallback remains explicit.");
    }

    private static DocxDocument ReadContainerPropertySource(string container, OoxPdfDocxMarkupMode mode,
        bool history, bool empty = false)
    {
        string fixture = container switch
        {
            "body" => "InlineBoxMath/inline-box-body-equation-inline-box-valid",
            "floating-table" => "TableGridOrigins/Containers/floating-table-ordered",
            _ => "InlineBoxMath/inline-box-math-" + container
        };
        string? part = container is "header" or "footer" ? "word/inline-box-" + container + ".xml" : null;
        return ReadMutation(fixture, mode, false, false, (xml, w, _) =>
        {
            // The floating fixture has a main-story comment. Remove its anchors
            // from both controls so it cannot mask a missing row-owner trigger.
            foreach (var anchor in xml.Descendants().Where(element => element.Name == w + "commentReference" ||
                element.Name == w + "commentRangeStart" || element.Name == w + "commentRangeEnd").ToArray()) anchor.Remove();
            if (!history) return;
            var box = xml.Descendants(w + "txbxContent").Single();
            bool table = container == "floating-table";
            var owner = table ? box.Descendants(w + "tr").First() : box.Descendants(w + "r").First(run => run.Element(w + "t") is not null);
            string name = table ? "trPr" : "rPr";
            var properties = owner.Element(w + name);
            if (properties is null) { properties = new XElement(w + name); owner.AddFirst(properties); }
            properties.Add(ContainerPropertyChange(w, name, empty));
        }, partOverride: part);
    }

    private static XElement ContainerPropertyChange(XNamespace w, string name, bool empty) =>
        new(w + (name + "Change"), new XAttribute(w + "id", "7701"), new XAttribute(w + "author", "Container Reviewer"),
            new XAttribute(w + "date", "2026-11-14T12:00:00Z"), new XElement(w + name, empty ? null :
                name == "rPr" ? new XElement(w + "b") : new XElement(w + "trHeight", new XAttribute(w + "val", "20"),
                    new XAttribute(w + "hRule", "atLeast"))));
}
