using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PortableRenderingTests
{
    public static void PortableRenderingRunProvenanceKeepsBodyCellAndHeaderInk()
    {
        var resolver = new PinnedResolver();
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string? control = null;
            foreach (string scope in new[] { "implicit", "explicit" })
            {
                var source = DocxTextTests.ReadRunStyleProvenanceSource(scope, null, mode);
                var renderer = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final
                    ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                var page = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().Single();
                foreach (string label in new[] { "BODY", "CELL", "HEADER" })
                    TestAssert.Equal(label.Length, ReadPaintedSourceLabel(page, label).Count);
                var runs = DocxBlockTraversal.EnumerateSourceParagraphs(source).SelectMany(paragraph => paragraph.Runs)
                    .Where(run => run.Text.Length != 0).ToArray();
                TestAssert.True(runs.All(run => run.StyleResolution.HasParagraphStyleRunProperties),
                    "All painted owners report their actual inherited run style.");
                control ??= page.Content;
                TestAssert.Equal(control, page.Content);
            }
        }
    }
}
