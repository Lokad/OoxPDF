using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static IReadOnlyList<PdfFontResource> RenderChartCategoryLabels(PptxDocument document, PptxTheme theme, PdfGraphicsBuilder graphics, ChartPlotBox plotBox, XDocument chartXml, PptxSceneChart? sceneChart, PptxSceneChartAxis? sceneAxis, XElement? categoryAxis, ChartIndexedTextVector labelVector, bool horizontalBars, double? verticalAxisY, bool categoryLabelsOnTickMarks, bool categoryLabelsTopSide, PresentationFontResolver? fontResolver)
    {
        IReadOnlyList<ChartIndexedTextPoint?> labels = labelVector.DensePoints();
        if (labels.Count == 0)
        {
            return [];
        }

        ChartTextStyle style = ReadSceneOrXmlChartTextStyle(theme, sceneChart, sceneAxis, chartXml, categoryAxis, fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize, chartStyleRole: "categoryAxis");
        double fontSize = style.FontSize;
        RgbColor color = style.Color;
        double labelOffsetScale = ResolveSceneOrXmlCategoryAxisLabelOffsetScale(sceneAxis, categoryAxis);
        int tickLabelSkip = ResolveSceneOrXmlCategoryAxisTickLabelSkip(sceneAxis, categoryAxis);
        var runs = new List<TextRun>(labels.Count);
        for (int i = 0; i < labels.Count; i++)
        {
            if (i % tickLabelSkip != 0)
            {
                continue;
            }

            string? label = labels[i]?.Text;
            if (string.IsNullOrWhiteSpace(label))
            {
                continue;
            }

            double x;
            double y;
            double width;
            double height = fontSize * PptxChartMetricRules.AxisLabelHeightFactor;
            TextAlignment alignment;
            if (horizontalBars)
            {
                double slotHeight = plotBox.Height / labels.Count;
                x = Math.Max(0d, plotBox.X - plotBox.Width * PptxChartMetricRules.CategoryAxisHorizontalLeftOffsetRatio * labelOffsetScale);
                y = plotBox.Y + slotHeight * (i + 0.5d) - height * PptxChartMetricRules.CategoryAxisHorizontalBaselineRatio;
                width = plotBox.Width * PptxChartMetricRules.CategoryAxisHorizontalWidthRatio;
                alignment = TextAlignment.Right;
            }
            else
            {
                double slotWidth = plotBox.Width / labels.Count;
                width = slotWidth * PptxChartMetricRules.CategoryAxisVerticalWidthFactor;
                double labelCenterX = categoryLabelsOnTickMarks && labels.Count > 1
                    ? plotBox.X + plotBox.Width * i / (labels.Count - 1)
                    : plotBox.X + slotWidth * (i + 0.5d);
                x = labelCenterX - width / 2d;
                double axisY = verticalAxisY ?? plotBox.Y;
                y = categoryLabelsTopSide
                    ? axisY + height * PptxChartMetricRules.CategoryAxisVerticalTopSideOffsetFactor * labelOffsetScale
                    : axisY - height * PptxChartMetricRules.CategoryAxisVerticalTopOffsetFactor * labelOffsetScale;
                alignment = TextAlignment.Center;
            }

            double labelWidth = Math.Max(1d, width);
            runs.Add(new TextRun(
                label,
                x,
                y,
                labelWidth,
                height,
                x,
                y - height * PptxChartMetricRules.AxisLabelClipTopOffsetFactor,
                labelWidth,
                height * PptxChartMetricRules.AxisLabelClipHeightFactor,
                fontSize,
                style.CharacterSpacing,
                0d,
                color,
                1d,
                null,
                Bold: style.Bold,
                Italic: style.Italic,
                Underline: style.Underline,
                Strike: style.Strike,
                KerningEnabled: true,
                alignment,
                FontFamily: style.FontFamily,
                RotationDegrees: 0d,
                RotationCenterX: 0d,
                RotationCenterY: 0d,
                FlipHorizontal: false,
                FlipVertical: false, PreventCoalesce: false, Outline: null, StrictClip: false));
        }

        return RenderTextRuns(runs, graphics, "CCA", fontResolver);
    }

    private static IReadOnlyList<PdfFontResource> RenderChartValueAxisLabels(PptxDocument document, PptxTheme theme, PdfGraphicsBuilder graphics, ChartPlotBox plotBox, XDocument chartXml, PptxSceneChart? sceneChart, XElement? valueAxis, PptxSceneChartAxis? sceneAxis, ChartValueExtents extents, ChartAxisUnits axisUnits, bool valueAxisReversed, bool horizontalBars, bool rightSide, int axisSideSlot, bool manualPlotLayoutApplied, bool useTextSizedWidth, string? defaultNumberFormat, PresentationFontResolver? fontResolver)
    {
        double range = Math.Max(1d, extents.Max - extents.Min);
        ChartTextStyle style = ReadSceneOrXmlChartTextStyle(theme, sceneChart, sceneAxis, chartXml, valueAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
        double fontSize = style.FontSize;
        double height = fontSize * PptxChartMetricRules.AxisLabelHeightFactor;
        RgbColor color = style.Color;
        var textMeasurer = new ChartTextMeasurer(fontResolver);
        double autoTickTargetCount = GetValueAxisAutoTickTargetCount(horizontalBars, valueAxisLabelsVisible: true, manualPlotLayoutApplied);
        IReadOnlyList<double> tickValues = GetChartAxisTickValues(extents, axisUnits.MajorUnit, includeEndpoints: true, autoTickTargetCount);
        double maxLabelWidth = tickValues
            .Select(value => FormatSceneOrXmlChartAxisLabel(value, sceneAxis, valueAxis, defaultNumberFormat))
            .DefaultIfEmpty("0")
            .Max(label => textMeasurer.Measure(label, style));
        double valueAxisLabelWidth = Math.Max(
            fontSize * PptxChartMetricRules.ValueAxisMinimumLabelWidthFactor,
            maxLabelWidth + fontSize * PptxChartMetricRules.ValueAxisLabelPaddingFactor);
        var runs = new List<TextRun>(tickValues.Count);
        foreach (double value in tickValues)
        {
            string label = FormatSceneOrXmlChartAxisLabel(value, sceneAxis, valueAxis, defaultNumberFormat);
            double offset = GetChartValuePlotRatio(extents, value, valueAxisReversed);
            double x;
            double y;
            double width;
            TextAlignment alignment;
            if (horizontalBars)
            {
                width = plotBox.Width / PptxChartMetricRules.HorizontalValueAxisSlotCount;
                x = plotBox.X + plotBox.Width * offset - width / 2d;
                y = plotBox.Y - height * PptxChartMetricRules.HorizontalValueAxisTopOffsetFactor;
                alignment = TextAlignment.Center;
            }
            else
            {
                bool labelsRightSide = ResolveSceneOrXmlValueAxisLabelsRightSide(sceneAxis, valueAxis, rightSide);
                width = useTextSizedWidth ? valueAxisLabelWidth : plotBox.Width * PptxChartMetricRules.VerticalValueAxisWidthRatio;
                double sideGap = Math.Max(3d, fontSize * PptxChartMetricRules.ValueAxisLabelSideGapFactor);
                x = labelsRightSide
                    ? plotBox.X + plotBox.Width + sideGap + axisSideSlot * (width + sideGap)
                    : Math.Max(0d, plotBox.X - (axisSideSlot + 1) * (width + sideGap));
                y = plotBox.Y + plotBox.Height * offset - height * PptxChartMetricRules.VerticalValueAxisBaselineRatio;
                alignment = labelsRightSide ? TextAlignment.Left : TextAlignment.Right;
            }

            double labelWidth = Math.Max(1d, width);
            double clipY = horizontalBars ? y : plotBox.Y - height;
            double clipHeight = horizontalBars ? height : plotBox.Height + height * PptxChartMetricRules.ValueAxisVerticalClipExtraHeightFactor;
            runs.Add(new TextRun(
                label,
                x,
                y,
                labelWidth,
                height,
                x,
                clipY,
                labelWidth,
                clipHeight,
                fontSize,
                style.CharacterSpacing,
                0d,
                color,
                1d,
                null,
                Bold: style.Bold,
                Italic: style.Italic,
                Underline: style.Underline,
                Strike: style.Strike,
                KerningEnabled: true,
                alignment,
                FontFamily: style.FontFamily,
                RotationDegrees: 0d,
                RotationCenterX: 0d,
                RotationCenterY: 0d,
                FlipHorizontal: false,
                FlipVertical: false, PreventCoalesce: false, Outline: null, StrictClip: false));
        }

        return RenderTextRuns(runs, graphics, "CVA", fontResolver);
    }

    private static IReadOnlyList<PdfFontResource> RenderSecondaryChartValueAxisLabels(PptxDocument document, PptxTheme theme, PdfGraphicsBuilder graphics, ChartPlotBox plotBox, XDocument chartXml, PptxSceneChart? sceneChart, ChartValueExtents fallback, PresentationFontResolver fontResolver)
    {
        ChartAxisSource rightValueAxisSource = ReadSceneOrXmlSecondaryRightValueAxis(sceneChart, chartXml);
        XElement? rightValueAxis = rightValueAxisSource.XmlAxis;
        PptxSceneChartAxis? rightSceneAxis = rightValueAxisSource.SceneAxis;
        if (rightValueAxis is null && rightSceneAxis is null)
        {
            return Array.Empty<PdfFontResource>();
        }

        ChartValueExtents extents = ReadSceneOrXmlChartValueAxisExtents(rightSceneAxis, rightValueAxis, fallback, false, PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
        ChartValueAxisRenderOptions axisOptions = ReadSceneOrXmlChartValueAxisRenderOptions(rightSceneAxis, rightValueAxis, theme, extents, percentStacked: false);
        return RenderChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, rightValueAxis, rightSceneAxis, extents, axisOptions.Units, axisOptions.Reversed, horizontalBars: false, rightSide: true, axisSideSlot: 0, manualPlotLayoutApplied: false, useTextSizedWidth: false, defaultNumberFormat: null, fontResolver: fontResolver);
    }
}
