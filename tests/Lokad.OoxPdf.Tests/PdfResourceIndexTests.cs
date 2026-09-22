using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PdfResourceIndexTests
{
    public static void AlphaRegistrationDedupesAndGrowsLinearly()
    {
        var graphics = new PdfGraphicsBuilder();
        graphics.SetAlpha(0.5d, 0.25d);
        TestAssert.True(graphics.ToString().Contains("/GS50000F25000S gs", StringComparison.Ordinal), "Alpha state names must stay deterministic.");

        const int first = 300;
        for (int i = 0; i < first; i++)
        {
            graphics.SetAlpha(i / (double)first, 1d);
        }

        TestAssert.Equal(first + 1, graphics.ExtGStates.Count);
        for (int i = 0; i < first; i++)
        {
            graphics.SetAlpha(i / (double)first, 1d);
        }

        TestAssert.Equal(first + 1, graphics.ExtGStates.Count);

        const int second = 600;
        for (int i = first; i < second; i++)
        {
            graphics.SetAlpha(i / (double)second, 0.5d);
        }

        TestAssert.Equal(first + 1 + (second - first), graphics.ExtGStates.Count);
        graphics.SetAlpha(0.5d, 0.25d);
        TestAssert.Equal(first + 1 + (second - first), graphics.ExtGStates.Count);
    }

    public static void SoftMaskEpsilonSemanticsPreserved()
    {
        var graphics = new PdfGraphicsBuilder();
        PdfImageXObject pixels = PdfImageXObject.Jpeg(2, 1, [1, 2, 3], 3, 8);
        var mask = new PdfLuminositySoftMask(pixels, 0d, 0d, 10d, 10d, 0d, 0d, 0d, 0d);
        graphics.SetLuminositySoftMask(mask, 0.5d, 1d);
        TestAssert.Equal(1, graphics.ExtGStates.Count);
        graphics.SetLuminositySoftMask(mask, 0.5d, 1d);
        TestAssert.Equal(1, graphics.ExtGStates.Count);
        graphics.SetLuminositySoftMask(mask, 0.5d + 0.0000005d, 1d);
        TestAssert.Equal(1, graphics.ExtGStates.Count);
        graphics.SetLuminositySoftMask(mask, 0.5d + 0.000002d, 1d);
        TestAssert.Equal(2, graphics.ExtGStates.Count);
    }

    public static void StateIndexesRollBackWithContent()
    {
        var graphics = new PdfGraphicsBuilder();
        graphics.SetAlpha(0.5d, 0.25d);
        PdfGraphicsBuilder.ContentMark mark = graphics.MarkContent();
        graphics.SetAlpha(0.1d, 0.2d);
        graphics.PaintAxialShading(0d, 0d, 10d, 10d, 255, 0, 0, 0, 0, 255);
        TestAssert.Equal(2, graphics.ExtGStates.Count);
        TestAssert.Equal(1, graphics.Shadings.Count);
        graphics.TruncateContent(mark);
        TestAssert.Equal(1, graphics.ExtGStates.Count);
        TestAssert.Equal(0, graphics.Shadings.Count);
        graphics.SetAlpha(0.1d, 0.2d);
        TestAssert.Equal(2, graphics.ExtGStates.Count);
        TestAssert.Equal("GS10000F20000S", graphics.ExtGStates[1].ResourceName);
        graphics.PaintAxialShading(0d, 0d, 10d, 10d, 255, 0, 0, 0, 0, 255);
        TestAssert.Equal(1, graphics.Shadings.Count);
        TestAssert.Equal("Sh1", graphics.Shadings[0].ResourceName);
        TestAssert.True(graphics.ToString().Contains("/GS10000F20000S gs", StringComparison.Ordinal), "Re-registered states must emit the same deterministic name.");
    }

    public static void PruneKeepsOnlyReferencedNamesWithPrefixGuard()
    {
        PdfImageXObject pixels = PdfImageXObject.Jpeg(2, 1, [1, 2, 3], 3, 8);
        var im1 = new PdfImageResource("Im1", pixels);
        var im12 = new PdfImageResource("Im12", pixels);
        var im2 = new PdfImageResource("Im2", pixels);
        List<PdfImageResource> kept = PptxRenderer.PruneUnreferencedImages("q /Im12 Do Q", [im1, im12, im2]);
        TestAssert.Equal(1, kept.Count);
        TestAssert.True(ReferenceEquals(im12, kept[0]), "Im12 must be kept without matching Im1 as a prefix.");
        List<PdfImageResource> conservative = PptxRenderer.PruneUnreferencedImages("BT (see /Im2 figure) Tj", [im2]);
        TestAssert.Equal(1, conservative.Count);
    }

    public static void PruneObservesCancellation()
    {
        PdfImageXObject pixels = PdfImageXObject.Jpeg(2, 1, [1, 2, 3], 3, 8);
        var images = new List<PdfImageResource> { new PdfImageResource("Im1", pixels) };
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => PptxRenderer.PruneUnreferencedImages(new string('x', 100000), images, cancelled.Token));
    }

    public static void PruneScalesAcrossManyResources()
    {
        string filler = new string('x', 20000);
        List<PdfImageResource> hundred = BuildImages(100);
        List<PdfImageResource> keptHundred = PptxRenderer.PruneUnreferencedImages(filler + " /Im100 Do", hundred);
        TestAssert.Equal(1, keptHundred.Count);
        TestAssert.True(ReferenceEquals(hundred[99], keptHundred[0]), "Only the referenced image must survive at N=100.");
        List<PdfImageResource> twoHundred = BuildImages(200);
        List<PdfImageResource> keptTwoHundred = PptxRenderer.PruneUnreferencedImages(filler + " /Im200 Do", twoHundred);
        TestAssert.Equal(1, keptTwoHundred.Count);
        TestAssert.True(ReferenceEquals(twoHundred[199], keptTwoHundred[0]), "Only the referenced image must survive at N=200.");
    }

    private static List<PdfImageResource> BuildImages(int count)
    {
        var images = new List<PdfImageResource>(count);
        for (int i = 1; i <= count; i++)
        {
            images.Add(new PdfImageResource("Im" + i.ToString(System.Globalization.CultureInfo.InvariantCulture), PdfImageXObject.Jpeg(2, 1, [1, 2, 3], 3, 8)));
        }

        return images;
    }
}
