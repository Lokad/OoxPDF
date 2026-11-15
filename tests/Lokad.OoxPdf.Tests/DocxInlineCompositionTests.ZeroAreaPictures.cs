using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxZeroAreaPicturesRetainNeighboringMathAcrossContainers()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string container in new[] { "body", "cell", "header", "footer" })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode, CachedMathGeometry(mode));
            var normal = renderer.RenderPagesAsync(ReadZeroAreaPicture(container, mode, null), null, default).ToBlockingEnumerable().ToArray();
            string expected = ZeroAreaMathText(normal);
            TestAssert.True(expected.Length > 0 && normal.Sum(page => page.Images.Count) > 0,
                "The independent source control contains actual formulas and a painting picture.");
            foreach (string axes in new[] { "width", "height", "both" })
            {
                var diagnostics = new List<OoxPdfDiagnostic>();
                var source = ReadZeroAreaPicture(container, mode, axes);
                var pages = renderer.RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
                TestAssert.Equal(expected, ZeroAreaMathText(pages));
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                    "A non-painting picture frame cannot discard supported neighboring formulas: " + container + "/" + axes + "/" + mode);
                TestAssert.Equal(0, pages.Sum(page => page.Images.Count));
                using var canceled = new CancellationTokenSource(); canceled.Cancel();
                TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(source, null, canceled.Token).ToBlockingEnumerable().ToArray());
            }
        }
    }

    public static void DocxZeroAreaPicturesKeepAuthoredFragmentDimensionsAndOwners()
    {
        foreach (string axes in new[] { "width", "height", "both" })
        foreach (double scale in new[] { 1d, .75d })
        {
            var document = ReadZeroAreaPicture("body", OoxPdfDocxMarkupMode.Final, axes);
            var paragraph = document.Paragraphs.Single(p => p.Equations.Count == 2 && p.Images.Count == 1);
            var image = paragraph.Images.Single();
            var equations = paragraph.Equations.ToDictionary(equation => equation,
                equation => new DocxPreparedMathParagraph(equation, new(18d, 10d, 4d, []), 10d, 16d));
            TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineAtoms(paragraph, equations, new TrackingMeasurer(),
                12d, 9d, 16d, scale, default, out var measured), "A finite nonnegative frame remains measurable even without paint area.");
            var atoms = TestAssert.NotNull(measured);
            var picture = atoms.Fragments.Single(fragment => fragment.Source is DocxInlinePicture);
            TestAssert.True(ReferenceEquals(image, ((DocxInlinePicture)picture.Source).Image),
                "The measured fragment retains its actual source owner.");
            TestAssert.Equal(image.WidthPoints * scale, picture.Advance);
            TestAssert.Equal(image.HeightPoints * scale, picture.BaselineInset);
            TestAssert.True(DocxLayoutEngine.TryWrapOrderedInlineAtoms(atoms, 468d, 468d, default, out var wrapped),
                "A zero advance or paint height cannot prevent source cursor progress.");
            var lines = TestAssert.NotNull(wrapped);
            TestAssert.Equal(2, lines.SelectMany(line => line.Fragments).Count(fragment => fragment.Math is not null));
            TestAssert.Equal(1, lines.SelectMany(line => line.Fragments).Count(fragment => fragment.Source is DocxInlinePicture));
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxLayoutEngine.TryWrapOrderedInlineAtoms(
                atoms, 468d, 468d, canceled.Token, out _));
        }
    }

    private static string ZeroAreaMathText(PdfPage[] pages) => string.Concat(pages.SelectMany(page =>
        DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true, includeTextTransforms: true,
            honorActualText: false, fontResourceFilter: name => name.StartsWith("Eq", StringComparison.Ordinal)))
        .Select(show => show.Text));

    private static DocxDocument ReadZeroAreaPicture(string container, OoxPdfDocxMarkupMode mode, string? axes) =>
        ReadMutation("CompositionProbes/inline-math-navigation" + (container == "body" ? "" : "-" + container),
            mode, false, false, (xml, w, m) =>
            {
                if (axes is null) return;
                XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
                var paragraph = xml.Descendants(w + "p").Single(p =>
                    p.Descendants(m + "oMath").Count() == 2 && p.Descendants(wp + "inline").Count() == 1);
                var extent = paragraph.Descendants(wp + "inline").Single().Element(wp + "extent")!;
                if (axes is "width" or "both") extent.SetAttributeValue("cx", "0");
                if (axes is "height" or "both") extent.SetAttributeValue("cy", "0");
            }, partOverride: container is "header" or "footer" ? "word/" + container + "1.xml" : null);
}
