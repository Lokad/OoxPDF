using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineBoxOutlineEnvelopeAdmitsCompleteFrames()
    {
        // The retained body/header/footer Word controls paint a full outline
        // where its interior touches the cell top. The child clip stays exact.
        var content = new DocxTableCellClipSnapshot(false, 62d, 100d, 300d, 68d);
        foreach (double stroke in new[] { .377d, .5d, 2d })
        foreach (double inset in new[] { 0d, 4d })
        {
            var frame = new PdfRectangle(72d, 108d - inset, 216d, 60d);
            var outline = DocxTableCellClipGeometry.ResolveInlineOutline(content, frame, stroke, 100d);
            TestAssert.True(outline.Y <= frame.Y - stroke / 2d + 1e-9 &&
                outline.Y + outline.Height >= frame.Y + frame.Height + stroke / 2d - 1e-9,
                "A complete admitted frame retains both halves of its top and bottom stroke.");
            TestAssert.Equal(62d, content.X);
            TestAssert.Equal(100d, content.Y);
            TestAssert.Equal(300d, content.Width);
            TestAssert.Equal(68d, content.Height);
        }
    }

    public static void DocxInlineBoxOutlineEnvelopeKeepsPartialAndInvalidFallbacks()
    {
        var content = new DocxTableCellClipSnapshot(false, 62d, 100d, 300d, 68d);
        var frame = new PdfRectangle(72d, 108d, 216d, 60d);
        foreach (var candidate in new (PdfRectangle Frame, double Stroke, double? Bottom)[]
        {
            (Frame: frame, Stroke: .5d, Bottom: (double?)109d),
            (frame with { X = 61d }, .5d, null),
            (frame with { Y = 99d }, .5d, null),
            (frame with { Width = 301d }, .5d, null),
            (frame with { Height = 61d }, .5d, null),
            (frame with { X = double.NaN }, .5d, null),
            (frame with { Width = double.PositiveInfinity }, .5d, null),
            (frame, 0d, null),
            (frame, double.PositiveInfinity, null),
            (frame, .5d, (double?)double.NaN)
        })
            TestAssert.True(ReferenceEquals(content, DocxTableCellClipGeometry.ResolveInlineOutline(
                content, candidate.Frame, candidate.Stroke, candidate.Bottom)),
                "Partial, overflowing or invalid frames keep their original cell clip.");
    }
}
