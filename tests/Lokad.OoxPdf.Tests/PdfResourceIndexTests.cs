using System.Globalization;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class PdfResourceIndexTests
{
    public static void AlphaRegistrationUsesEmittedPairsAndRetainsRawValues()
    {
        CultureInfo previous = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
            var graphics = new PdfGraphicsBuilder();
            graphics.SetAlpha(.50001d, .25001d);
            graphics.SetAlpha(.50049d, .25049d);
            TestAssert.Equal(1, graphics.ExtGStates.Count);
            TestAssert.Equal("GS50000F25000S", graphics.ExtGStates[0].ResourceName);
            TestAssert.Equal(.50001d, graphics.ExtGStates[0].FillAlpha);
            TestAssert.Equal(.25001d, graphics.ExtGStates[0].StrokeAlpha);
            graphics.SetAlpha(.5005d, .25001d);
            graphics.SetAlpha(.50001d, .2505d);
            TestAssert.Equal(3, graphics.ExtGStates.Count);
            TestAssert.Equal("GS50100F25000S", graphics.ExtGStates[1].ResourceName);
            TestAssert.Equal("GS50000F25100S", graphics.ExtGStates[2].ResourceName);
            graphics.SetAlpha(.49949d, .25001d);
            TestAssert.Equal("GS49900F25000S", graphics.ExtGStates[3].ResourceName);
            graphics.SetAlpha(.49951d, .25001d);
            TestAssert.Equal(4, graphics.ExtGStates.Count);
            graphics.SetAlpha(-.5d, 2d);
            graphics.SetAlpha(.00049d, .99951d);
            TestAssert.Equal(5, graphics.ExtGStates.Count);
            TestAssert.Equal("GS0F100000S", graphics.ExtGStates[4].ResourceName);
            string before = graphics.ToString();
            TestAssert.Throws<ArgumentOutOfRangeException>(() => graphics.SetAlpha(double.NaN, 1d));
            TestAssert.Throws<ArgumentOutOfRangeException>(() => graphics.SetAlpha(.5d, double.NaN));
            TestAssert.Equal(before, graphics.ToString());
            TestAssert.Equal(5, graphics.ExtGStates.Count);
        }
        finally { CultureInfo.CurrentCulture = previous; }
    }

    public static void AlphaRegistrationKeepsMaskAndFormNamespacesDistinct()
    {
        var child = new PdfGraphicsBuilder();
        child.SetAlpha(.50049d, 1d);
        child.FillRectangle(0d, 0d, 10d, 10d);
        var graphics = new PdfGraphicsBuilder();
        graphics.SetAlpha(.50001d, 1d);
        var mark = graphics.MarkContent();
        graphics.SetVectorLuminositySoftMask(child, new(0d, 0d, 10d, 10d), .50002d, 1d);
        graphics.SetAlpha(.50049d, 1d);
        TestAssert.Equal(2, graphics.ExtGStates.Count);
        TestAssert.Equal(null, graphics.ExtGStates[0].LuminosityGroupName);
        TestAssert.Equal("Tr1", graphics.ExtGStates[1].LuminosityGroupName);
        TestAssert.Equal(.50002d, graphics.ExtGStates[1].FillAlpha);
        TestAssert.Equal(graphics.ExtGStates[0].ResourceName, child.ExtGStates.Single().ResourceName);
        graphics.SetVectorLuminositySoftMask(child, new(0d, 0d, 10d, 10d), .50002d, 1d);
        TestAssert.Equal(3, graphics.ExtGStates.Count);
        TestAssert.Equal("Tr2", graphics.ExtGStates[2].LuminosityGroupName);
        PdfContentValidator.ValidatePage(new(100d, 100d, graphics.ToString(), [], [], graphics.ExtGStates, [], [], [], groups: graphics.Groups), 0, default);
        graphics.TruncateContent(mark);
        TestAssert.Equal(1, graphics.ExtGStates.Count);
        TestAssert.Equal(0, graphics.Groups.Count);
        graphics.SetAlpha(.50049d, 1d);
        TestAssert.Equal(1, graphics.ExtGStates.Count);
        graphics.SetVectorLuminositySoftMask(child, new(0d, 0d, 10d, 10d), .50002d, 1d);
        TestAssert.Equal("GSV2", graphics.ExtGStates[1].ResourceName);
        TestAssert.Equal("Tr1", graphics.ExtGStates[1].LuminosityGroupName);
    }

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

    public static void StateIndexesRollBackWithContent()
    {
        var graphics = new PdfGraphicsBuilder();
        graphics.SetAlpha(0.5d, 0.25d);
        PdfGraphicsBuilder.ContentMark mark = graphics.MarkContent();
        graphics.SetAlpha(.50049d, .25049d);
        TestAssert.Equal(1, graphics.ExtGStates.Count);
        graphics.SetAlpha(0.1d, 0.2d);
        graphics.PaintAxialShading(0d, 0d, 10d, 10d, 255, 0, 0, 0, 0, 255);
        TestAssert.Equal(2, graphics.ExtGStates.Count);
        TestAssert.Equal(1, graphics.Shadings.Count);
        graphics.TruncateContent(mark);
        TestAssert.Equal(1, graphics.ExtGStates.Count);
        TestAssert.Equal(0, graphics.Shadings.Count);
        graphics.SetAlpha(.50001d, .25001d);
        TestAssert.Equal(1, graphics.ExtGStates.Count);
        graphics.SetAlpha(0.1d, 0.2d);
        TestAssert.Equal(2, graphics.ExtGStates.Count);
        TestAssert.Equal("GS10000F20000S", graphics.ExtGStates[1].ResourceName);
        graphics.PaintAxialShading(0d, 0d, 10d, 10d, 255, 0, 0, 0, 0, 255);
        TestAssert.Equal(1, graphics.Shadings.Count);
        TestAssert.Equal("Sh1", graphics.Shadings[0].ResourceName);
        TestAssert.True(graphics.ToString().Contains("/GS10000F20000S gs", StringComparison.Ordinal), "Re-registered states must emit the same deterministic name.");
    }
}
