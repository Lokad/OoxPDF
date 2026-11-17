using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PortableRenderingTests
{
    public static void PortableRenderingCommentNumericIdentityKeepsVisibleBalloonBodies()
    {
        var resolver = new PinnedResolver();
        foreach (var (story, reference, start, end, cell) in new[]
        {
            ("42", "42", "42", "42", false),
            ("0042", "42", "42", "42", false),
            ("42", "+0042", "42", "42", false),
            ("42", "42", "0042", "42", false),
            ("42", "42", "42", "+42", false),
            ("0042", "+42", "42", "0042", true),
        })
        {
            var source = DocxCommentsTests.ReadCommentIdentitySource(story, reference, start, end, cell: cell);
            var page = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup,
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderPagesAsync(source, null, default).ToBlockingEnumerable().Single();
            foreach (string label in new[] { "ANCHOR", "REVIEWNOTE" })
                TestAssert.Equal(label.Length, ReadPaintedSourceLabel(page, label).Count);
        }
    }
}
