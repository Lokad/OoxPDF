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
    private static IReadOnlyList<XElement> ReadChartValueAxesForChart(XDocument chartXml, XElement chartElement)
    {
        string[] axisIds = chartElement
            .Elements(ChartNamespace + "axId")
            .Select(PptxSceneBuilder.ReadChartValueAttribute)
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .ToArray();
        if (axisIds.Length == 0)
        {
            return chartXml.Descendants(ChartNamespace + "valAx").ToArray();
        }

        var axes = new List<XElement>();
        foreach (string axisId in axisIds)
        {
            XElement? axis = chartXml
                .Descendants(ChartNamespace + "valAx")
                .FirstOrDefault(candidate => string.Equals(ReadChartAxisId(candidate), axisId, StringComparison.Ordinal));
            if (axis is not null)
            {
                axes.Add(axis);
            }
        }

        return axes;
    }

    private static IReadOnlyList<XElement> ReadChartCategoryAxesForChart(XDocument chartXml, XElement chartElement)
    {
        string[] axisIds = chartElement
            .Elements(ChartNamespace + "axId")
            .Select(PptxSceneBuilder.ReadChartValueAttribute)
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .ToArray();
        if (axisIds.Length == 0)
        {
            return ReadChartCategoryAxes(chartXml).ToArray();
        }

        var axes = new List<XElement>();
        foreach (string axisId in axisIds)
        {
            XElement? axis = ReadChartCategoryAxes(chartXml)
                .FirstOrDefault(candidate => string.Equals(ReadChartAxisId(candidate), axisId, StringComparison.Ordinal));
            if (axis is not null)
            {
                axes.Add(axis);
            }
        }

        return axes.Count == 0
            ? ReadChartCategoryAxes(chartXml).Take(1).ToArray()
            : axes;
    }

    private static IEnumerable<XElement> ReadChartCategoryAxes(XDocument chartXml)
    {
        return chartXml
            .Descendants(ChartNamespace + "plotArea")
            .FirstOrDefault()?
            .Elements()
            .Where(element => element.Name == ChartNamespace + "catAx" || element.Name == ChartNamespace + "dateAx") ??
            [];
    }

    private static ChartValueExtents ReadChartValueAxisExtents(XElement? valueAxis, ChartValueExtents fallback, bool useNearMaximumHeadroom, double nearMaximumHeadroomRatio)
    {
        return ReadChartValueAxisExtents(valueAxis, fallback, PptxChartMetricRules.AxisNiceTickTargetCount, useNearMaximumHeadroom, nearMaximumHeadroomRatio);
    }

    private static ChartValueExtents ReadBubbleChartValueAxisExtents(XElement? valueAxis, ChartValueExtents fallback)
    {
        return ReadChartValueAxisExtents(valueAxis, fallback, PptxChartMetricRules.BubbleAxisBoundsTickTargetCount, false, PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
    }

    private static ChartValueExtents ReadSceneOrXmlBubbleChartValueAxisExtents(PptxSceneChartAxis? axis, XElement? valueAxis, ChartValueExtents fallback)
    {
        if (axis is null)
        {
            return ReadBubbleChartValueAxisExtents(valueAxis, fallback);
        }

        if (!axis.HasScaling)
        {
            return fallback;
        }

        double min = axis.Minimum ?? GetNiceChartAxisMin(fallback.Min, fallback.Max);
        double max = axis.Maximum ?? GetNiceChartAxisMax(fallback.Max, min, PptxChartMetricRules.BubbleAxisBoundsTickTargetCount, false, PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
        return max > min
            ? new ChartValueExtents(min, max)
            : fallback;
    }

    private static ChartValueExtents ReadChartValueAxisExtents(XElement? valueAxis, ChartValueExtents fallback, double boundsTickTargetCount, bool useNearMaximumHeadroom, double nearMaximumHeadroomRatio)
    {
        XElement? scaling = valueAxis?.Element(ChartNamespace + "scaling");
        if (valueAxis is null || scaling is null)
        {
            return fallback;
        }

        double min = PptxSceneBuilder.ReadChartAxisScalingValueWithValue(valueAxis, "min").Value ?? GetNiceChartAxisMin(fallback.Min, fallback.Max);
        double max = PptxSceneBuilder.ReadChartAxisScalingValueWithValue(valueAxis, "max").Value ?? GetNiceChartAxisMax(fallback.Max, min, boundsTickTargetCount, useNearMaximumHeadroom, nearMaximumHeadroomRatio);
        return max > min
            ? new ChartValueExtents(min, max)
            : fallback;
    }

    private static ChartValueExtents ReadSceneOrXmlChartValueAxisExtents(PptxSceneChartAxis? axis, XElement? valueAxis, ChartValueExtents fallback, bool useNearMaximumHeadroom, double nearMaximumHeadroomRatio)
    {
        if (axis is null)
        {
            return ReadChartValueAxisExtents(valueAxis, fallback, useNearMaximumHeadroom, nearMaximumHeadroomRatio);
        }

        if (!axis.HasScaling)
        {
            return fallback;
        }

        double min = axis.Minimum ?? GetNiceChartAxisMin(fallback.Min, fallback.Max);
        double max = axis.Maximum ?? GetNiceChartAxisMax(fallback.Max, min, PptxChartMetricRules.AxisNiceTickTargetCount, useNearMaximumHeadroom, nearMaximumHeadroomRatio);
        return max > min
            ? new ChartValueExtents(min, max)
            : fallback;
    }

    private static ChartValueExtents ReadPercentStackedAwareValueAxisExtents(PptxSceneChartAxis? axis, XElement? valueAxis, ChartValueExtents fallback, bool percentStacked, bool useNearMaximumHeadroom, double nearMaximumHeadroomRatio)
    {
        ChartValueExtents extents = ReadSceneOrXmlChartValueAxisExtents(axis, valueAxis, fallback, useNearMaximumHeadroom, nearMaximumHeadroomRatio);
        if (!percentStacked)
        {
            return extents;
        }

        double min = HasSceneOrXmlAxisScalingValue(axis, valueAxis, ChartAxisScalingBound.Minimum) ? extents.Min : 0d;
        double max = HasSceneOrXmlAxisScalingValue(axis, valueAxis, ChartAxisScalingBound.Maximum) ? extents.Max : 1d;
        return max > min
            ? new ChartValueExtents(min, max)
            : extents;
    }

    private static bool HasSceneOrXmlAxisScalingValue(PptxSceneChartAxis? axis, XElement? valueAxis, ChartAxisScalingBound bound)
    {
        if (axis is not null)
        {
            return bound == ChartAxisScalingBound.Minimum
                ? axis.Minimum is not null
                : axis.Maximum is not null;
        }

        string elementName = bound == ChartAxisScalingBound.Minimum ? "min" : "max";
        return valueAxis?
            .Element(ChartNamespace + "scaling")
            ?.Element(ChartNamespace + elementName) is not null;
    }

    private static ChartAxisUnits ReadChartValueAxisUnits(XElement? valueAxis)
    {
        if (valueAxis is null)
        {
            return ChartAxisUnits.Empty;
        }

        return new ChartAxisUnits(
            PptxSceneBuilder.ReadChartAxisUnitValueWithValue(valueAxis, "majorUnit").Value,
            PptxSceneBuilder.ReadChartAxisUnitValueWithValue(valueAxis, "minorUnit").Value);
    }

    private static ChartAxisUnits ReadSceneOrXmlChartValueAxisUnits(PptxSceneChartAxis? axis, XElement? valueAxis)
    {
        return axis is null
            ? ReadChartValueAxisUnits(valueAxis)
            : new ChartAxisUnits(axis.MajorUnit, axis.MinorUnit);
    }

    private static ChartAxisUnits ResolvePercentStackedAxisUnits(ChartAxisUnits axisUnits, bool percentStacked)
    {
        return percentStacked && axisUnits.MajorUnit is null
            ? axisUnits with { MajorUnit = 0.1d }
            : axisUnits;
    }

    private static ChartAxisUnits ResolveBubbleAxisUnits(ChartAxisUnits axisUnits, ChartValueExtents extents)
    {
        return axisUnits.MajorUnit is null
            ? axisUnits with { MajorUnit = ChooseChartAxisMajorUnit(Math.Max(1d, extents.Max - extents.Min), PptxChartMetricRules.BubbleAxisNiceTickTargetCount) }
            : axisUnits;
    }

    private static void RenderChartShapeStyle(PdfGraphicsBuilder graphics, double x, double y, double width, double height, ChartShapeStyle style)
    {
        if (ToGlow(style.Glow) is { } glow)
        {
            DrawGlow(graphics, "rect", x, y, width, height, glow);
        }

        if (ToOuterShadow(style.OuterShadow) is { } outerShadow)
        {
            DrawOuterShadow(graphics, "rect", x, y, width, height, outerShadow);
        }

        if (style.GradientFill is { } gradientFill)
        {
            DrawLinearGradientFill(graphics, gradientFill, x, y, width, height);
        }
        else if (style.Fill is { } fill)
        {
            FillChartRectangle(graphics, x, y, width, height, fill);
        }

        if (style.Stroke is { } stroke)
        {
            if (stroke.Alpha < 1d)
            {
                graphics.SaveState();
                graphics.SetAlpha(1d, stroke.Alpha);
            }

            graphics.SetStrokeRgb(stroke.Color.Red, stroke.Color.Green, stroke.Color.Blue);
            graphics.SetLineWidth(stroke.Width);
            if (stroke.DashPattern is { Count: > 0 })
            {
                graphics.SetLineDash(stroke.DashPattern);
            }

            if (stroke.Cap is { } cap)
            {
                graphics.SetLineCap(cap);
            }

            if (stroke.Join is { } join)
            {
                graphics.SetLineJoin(join);
            }

            graphics.StrokeRectangle(x, y, width, height);
            if (stroke.DashPattern is { Count: > 0 })
            {
                graphics.ClearLineDash();
            }

            if (stroke.Cap is not null)
            {
                graphics.SetLineCap(0);
            }

            if (stroke.Join is not null)
            {
                graphics.SetLineJoin(0);
            }
            if (stroke.Alpha < 1d)
            {
                graphics.RestoreState();
            }
        }
    }

    private static void RenderInChartPlotAreaClip(PdfGraphicsBuilder graphics, ChartPlotBox plotBox, Action render)
    {
        graphics.SaveState();
        try
        {
            ClipChartPlotArea();
            render();
        }
        finally
        {
            graphics.RestoreState();
        }

        void ClipChartPlotArea()
        {
            graphics.ClipRectangleEvenOdd(plotBox.X, plotBox.Y, plotBox.Width, plotBox.Height);
        }
    }

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

    private static double GetValueAxisAutoTickTargetCount(bool horizontalBars, bool valueAxisLabelsVisible, bool manualPlotLayoutApplied)
    {
        if (!horizontalBars)
        {
            return PptxChartMetricRules.AxisNiceVerticalValueTickTargetCount;
        }

        return valueAxisLabelsVisible && manualPlotLayoutApplied
            ? PptxChartMetricRules.AxisNiceTickTargetCount
            : PptxChartMetricRules.AxisNiceHorizontalValueTickTargetCount;
    }

    private sealed class ChartTextMeasurer
    {
        private readonly TextAdvanceEstimator estimator;

        public ChartTextMeasurer(PresentationFontResolver? fontResolver)
        {
            estimator = new TextAdvanceEstimator(fontResolver, CancellationToken.None);
        }

        public double Measure(string text, ChartTextStyle style)
        {
            return Measure(text, style.FontSize, style.FontFamily, style.Bold, style.Italic, style.CharacterSpacing);
        }

        public double Measure(string text, double fontSize, string? fontFamily, bool bold, bool italic, double characterSpacing)
        {
            return estimator.Measure(
                text,
                fontSize,
                fontFamily,
                bold,
                italic,
                characterSpacing,
                kerningEnabled: true);
        }
    }

    private static int GetValueAxisSideSlot(PptxSceneChartAxis? primarySceneAxis, XElement? primaryAxis, PptxSceneChartAxis? secondarySceneAxis, XElement secondaryAxis, bool defaultPrimaryRightSide, bool defaultSecondaryRightSide)
    {
        bool primaryRight = ResolveSceneOrXmlValueAxisLabelsRightSide(primarySceneAxis, primaryAxis, defaultPrimaryRightSide);
        bool secondaryRight = ResolveSceneOrXmlValueAxisLabelsRightSide(secondarySceneAxis, secondaryAxis, defaultSecondaryRightSide);
        return primaryRight == secondaryRight ? 1 : 0;
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

    private static XElement? ReadSecondaryRightValueAxis(XDocument chartXml)
    {
        return chartXml
            .Descendants(ChartNamespace + "valAx")
            .Where(IsRightValueAxis)
            .FirstOrDefault();
    }

    private static XElement? ReadSecondaryValueAxisForChart(XDocument chartXml, XElement? primaryValueAxis)
    {
        string? primaryAxisId = ReadChartAxisId(primaryValueAxis);
        if (string.IsNullOrWhiteSpace(primaryAxisId))
        {
            return ReadSecondaryRightValueAxis(chartXml);
        }

        return chartXml
            .Descendants(ChartNamespace + "valAx")
            .FirstOrDefault(axis =>
                !string.Equals(ReadChartAxisId(axis), primaryAxisId, StringComparison.Ordinal) &&
                !IsChartAxisDeleted(axis));
    }

    private static PptxSceneChartAxis? ReadSceneSecondaryRightValueAxis(PptxSceneChart? sceneChart, XElement? rightValueAxis)
    {
        if (sceneChart is null)
        {
            return null;
        }

        string? axisId = ReadChartAxisId(rightValueAxis);
        if (!string.IsNullOrWhiteSpace(axisId))
        {
            PptxSceneChartAxis? matchingAxis = sceneChart.Axes.FirstOrDefault(axis =>
                string.Equals(axis.Id, axisId, StringComparison.Ordinal) &&
                axis.AxisKind == PptxSceneChartAxisKind.Value);
            if (matchingAxis is not null)
            {
                return matchingAxis;
            }
        }

        return sceneChart.Axes.FirstOrDefault(axis =>
            axis.AxisKind == PptxSceneChartAxisKind.Value &&
            axis.PositionKind == PptxSceneChartAxisPosition.Right);
    }

    private static bool IsSceneOrXmlVisibleValueAxis(PptxSceneChartAxis? sceneAxis, XElement? axis)
    {
        if (sceneAxis is not null)
        {
            return sceneAxis.AxisKind == PptxSceneChartAxisKind.Value &&
                sceneAxis.IsDeleted != true;
        }

        return axis is not null &&
            !PptxSceneBuilder.IsOoxmlBooleanElementEnabled(axis.Element(ChartNamespace + "delete"));
    }

    private static bool IsRightValueAxis(XElement? axis)
    {
        return axis is not null &&
            PptxSceneBuilder.ParseChartAxisPosition(PptxSceneBuilder.ReadChartElementValue(axis, "axPos")) == PptxSceneChartAxisPosition.Right &&
            !PptxSceneBuilder.IsOoxmlBooleanElementEnabled(axis.Element(ChartNamespace + "delete"));
    }

    private static string? ReadChartAxisId(XElement? axis)
    {
        return PptxSceneBuilder.ReadOptionalChartValueAttribute(axis?.Element(ChartNamespace + "axId"));
    }

    private static IReadOnlyList<double> GetChartAxisTickValues(ChartValueExtents extents, double? explicitUnit, bool includeEndpoints, double autoTickTargetCount)
    {
        double range = Math.Max(1d, extents.Max - extents.Min);
        if (explicitUnit is not { } unit || unit <= 0d)
        {
            unit = ChooseChartAxisMajorUnit(range, autoTickTargetCount);
        }

        var values = new List<double>();
        if (includeEndpoints)
        {
            values.Add(extents.Min);
        }

        double first = Math.Ceiling(extents.Min / unit) * unit;
        for (double value = first; value < extents.Max - PptxChartMetricRules.AxisValueEpsilon; value += unit)
        {
            if (value > extents.Min + PptxChartMetricRules.AxisValueEpsilon)
            {
                values.Add(value);
            }
        }

        if (includeEndpoints)
        {
            values.Add(extents.Max);
        }

        return values;
    }

    private static IReadOnlyList<double> GetChartGridlineValues(ChartValueExtents extents, double? explicitUnit, double? crossingValue, double autoTickTargetCount)
    {
        return GetChartAxisTickValues(extents, explicitUnit, includeEndpoints: true, autoTickTargetCount)
            .Where(value => crossingValue is not { } crossing || Math.Abs(value - crossing) > PptxChartMetricRules.AxisValueEpsilon)
            .ToArray();
    }

    private static double? ReadSceneOrXmlValueAxisCrossingValue(PptxSceneChartAxis? sceneAxis, XElement? valueAxis, ChartValueExtents extents)
    {
        double? crossesAt = sceneAxis is not null
            ? sceneAxis.CrossesAt
            : valueAxis is null
                ? null
                : PptxSceneBuilder.ReadChartElementDoubleWithValue(valueAxis, "crossesAt").Value;
        if (crossesAt is { } explicitCrossing)
        {
            return explicitCrossing;
        }

        PptxSceneChartAxisCrosses crosses = sceneAxis is not null
            ? ResolveChartAxisCrosses(sceneAxis.CrossesKind)
            : ResolveChartAxisCrosses(PptxSceneBuilder.ParseChartAxisCrosses(PptxSceneBuilder.ReadChartElementValue(valueAxis, "crosses")));
        if (crosses == PptxSceneChartAxisCrosses.Maximum)
        {
            return extents.Max;
        }

        if (crosses == PptxSceneChartAxisCrosses.Minimum)
        {
            return extents.Min;
        }

        if (extents.Min <= 0d && extents.Max >= 0d)
        {
            return 0d;
        }

        return extents.Min > 0d ? extents.Min : extents.Max;

        PptxSceneChartAxisCrosses ResolveChartAxisCrosses(PptxSceneChartAxisCrosses crosses)
        {
            return crosses == PptxSceneChartAxisCrosses.Unknown
                ? PptxSceneChartAxisCrosses.AutoZero
                : crosses;
        }
    }

    private static bool ReadSceneOrXmlValueAxisReversed(PptxSceneChartAxis? sceneAxis, XElement? valueAxis)
    {
        if (sceneAxis is not null)
        {
            return ResolveChartAxisReversed(sceneAxis.OrientationKind);
        }

        return ResolveChartAxisReversed(PptxSceneBuilder.ParseChartAxisOrientation(
            PptxSceneBuilder.ReadChartElementValue(valueAxis?.Element(ChartNamespace + "scaling"), "orientation")));

        bool ResolveChartAxisReversed(PptxSceneChartAxisOrientation orientation)
        {
            return orientation == PptxSceneChartAxisOrientation.MaximumMinimum;
        }
    }

    private static double GetChartValuePlotRatio(ChartValueExtents extents, double value, bool reversed)
    {
        double range = Math.Max(1d, extents.Max - extents.Min);
        double boundedValue = Math.Clamp(value, extents.Min, extents.Max);
        double ratio = (boundedValue - extents.Min) / range;
        return reversed ? 1d - ratio : ratio;
    }

    private static double ChartValueToPlotCoordinate(ChartValueExtents extents, double? value, double plotStart, double plotLength, bool reversed)
    {
        return plotStart + plotLength * GetChartValuePlotRatio(extents, value ?? 0d, reversed);
    }

    private static double ChooseChartAxisMajorUnit(double range, double tickTargetCount)
    {
        double target = Math.Max(range / Math.Max(1d, tickTargetCount), double.Epsilon);
        double magnitude = Math.Pow(PptxChartMetricRules.AxisNiceTickStepMaximum, Math.Floor(Math.Log10(target)));
        double normalized = target / magnitude;
        double nice = normalized <= PptxChartMetricRules.AxisNiceTickStepSmall
            ? PptxChartMetricRules.AxisNiceTickStepSmall
            : normalized <= PptxChartMetricRules.AxisNiceTickStepMedium
                ? PptxChartMetricRules.AxisNiceTickStepMedium
                : normalized <= PptxChartMetricRules.AxisNiceTickStepLarge
                    ? PptxChartMetricRules.AxisNiceTickStepLarge
                    : PptxChartMetricRules.AxisNiceTickStepMaximum;
        return nice * magnitude;
    }

    private static double GetNiceChartAxisMin(double dataMin, double dataMax)
    {
        if (dataMin >= 0d)
        {
            return dataMin;
        }

        double range = dataMax - dataMin;
        if (Math.Abs(range) < PptxChartMetricRules.AxisValueEpsilon)
        {
            return dataMin;
        }

        double unit = ChooseChartAxisMajorUnit(range, PptxChartMetricRules.AxisNiceTickTargetCount);
        double niceMin = Math.Floor(dataMin / unit) * unit;
        return niceMin < dataMax ? niceMin : dataMin;
    }

    private static double GetNiceChartAxisMax(double dataMax, double dataMin, double tickTargetCount, bool useNearMaximumHeadroom, double nearMaximumHeadroomRatio)
    {
        if (Math.Abs(dataMax) < PptxChartMetricRules.AxisValueEpsilon && Math.Abs(dataMin) < PptxChartMetricRules.AxisValueEpsilon)
        {
            return 1d;
        }

        double range = dataMax - Math.Min(0d, dataMin);
        if (Math.Abs(range) < PptxChartMetricRules.AxisValueEpsilon)
        {
            return dataMax > 0d ? dataMax * PptxChartMetricRules.AxisSingleValueHeadroomFactor : 1d;
        }

        double unit = ChooseChartAxisMajorUnit(range, tickTargetCount);
        double niceMax = Math.Ceiling(dataMax / unit) * unit;
        if (niceMax < dataMax + PptxChartMetricRules.AxisValueEpsilon)
        {
            return niceMax + unit;
        }

        if (useNearMaximumHeadroom &&
            niceMax > 0d &&
            dataMax > 0d &&
            dataMax / niceMax >= nearMaximumHeadroomRatio)
        {
            return niceMax + unit;
        }

        return niceMax;
    }

    private static IReadOnlyDictionary<int, double> ReadChartPointExplosions(XElement chartElement)
    {
        var explosions = new Dictionary<int, double>();
        XElement? series = chartElement.Element(ChartNamespace + "ser");
        if (series is null)
        {
            return explosions;
        }

        (double? seriesExplosion, _) = PptxSceneBuilder.ReadChartElementDoubleWithValue(series, "explosion");
        if (seriesExplosion is { } seriesExplosionValue)
        {
            double fraction = Math.Clamp(seriesExplosionValue / 100d, 0d, 1d);
            int pointCount = PptxSceneBuilder
                .ReadChartNumberPoints(
                    series
                        .Elements(ChartNamespace + "val")
                        .Descendants(ChartNamespace + "pt"),
                    requireNonNegativeIndex: true)
                .Select(point => point.HasParsedIndex ? point.Index : -1)
                .Where(index => index >= 0)
                .DefaultIfEmpty(-1)
                .Max() + 1;
            for (int index = 0; index < pointCount; index++)
            {
                explosions[index] = fraction;
            }
        }

        foreach (XElement point in series.Elements(ChartNamespace + "dPt"))
        {
            if (!PptxSceneBuilder.TryReadChartNonNegativeIndex(point, out int index, out _))
            {
                continue;
            }

            (double? explosion, _) = PptxSceneBuilder.ReadChartElementDoubleWithValue(point, "explosion");
            if (explosion is { } explosionValue)
            {
                explosions[index] = Math.Clamp(explosionValue / 100d, 0d, 1d);
            }
        }

        return explosions;
    }

    private static bool ReadSceneOrXmlMajorGridlines(PptxSceneChartAxis? sceneAxis, XElement? axis)
    {
        return sceneAxis?.HasMajorGridlines ?? HasMajorGridlines();

        bool HasMajorGridlines()
        {
            return PptxSceneBuilder.IsChartGridlineVisible(axis?.Element(ChartNamespace + "majorGridlines"));
        }
    }

    private static bool ReadSceneOrXmlMinorGridlines(PptxSceneChartAxis? sceneAxis, XElement? axis)
    {
        return sceneAxis?.HasMinorGridlines ?? HasMinorGridlines();

        bool HasMinorGridlines()
        {
            return PptxSceneBuilder.IsChartGridlineVisible(axis?.Element(ChartNamespace + "minorGridlines"));
        }
    }

    private static ChartGridlineStyle ReadSceneOrXmlChartGridlineStyle(PptxSceneChartAxis? sceneAxis, XElement? xmlAxis, PptxTheme theme)
    {
        return new ChartGridlineStyle(
            ReadSceneOrXmlChartGridlineStroke(sceneAxis, sceneAxis?.MajorGridlineLine ?? default, sceneAxis?.MajorGridlineStyleLine ?? default, xmlAxis?.Element(ChartNamespace + "majorGridlines"), theme),
            ReadSceneOrXmlChartGridlineStroke(sceneAxis, sceneAxis?.MinorGridlineLine ?? default, sceneAxis?.MinorGridlineStyleLine ?? default, xmlAxis?.Element(ChartNamespace + "minorGridlines"), theme));
    }

    private static ChartValueAxisRenderOptions ReadSceneOrXmlChartValueAxisRenderOptions(
        PptxSceneChartAxis? sceneAxis,
        XElement? xmlAxis,
        PptxTheme theme,
        ChartValueExtents extents,
        bool percentStacked)
    {
        return new ChartValueAxisRenderOptions(
            ResolvePercentStackedAxisUnits(ReadSceneOrXmlChartValueAxisUnits(sceneAxis, xmlAxis), percentStacked),
            ReadSceneOrXmlValueAxisReversed(sceneAxis, xmlAxis),
            ReadSceneOrXmlValueAxisCrossingValue(sceneAxis, xmlAxis, extents),
            ReadSceneOrXmlMajorGridlines(sceneAxis, xmlAxis),
            ReadSceneOrXmlMinorGridlines(sceneAxis, xmlAxis),
            ReadSceneOrXmlChartGridlineStyle(sceneAxis, xmlAxis, theme));
    }

    private static ChartBubbleValueAxisOptions ReadSceneOrXmlChartBubbleValueAxisOptions(
        PptxSceneChartAxis? sceneAxis,
        XElement? xmlAxis,
        PptxTheme theme,
        ChartValueExtents extents)
    {
        return new ChartBubbleValueAxisOptions(
            ResolveBubbleAxisUnits(ReadSceneOrXmlChartValueAxisUnits(sceneAxis, xmlAxis), extents),
            ReadSceneOrXmlChartGridlineStyle(sceneAxis, xmlAxis, theme));
    }

    private static ChartSeriesStroke? ReadSceneOrXmlChartGridlineStroke(PptxSceneChartAxis? sceneAxis, PptxSceneLineStyle sceneLine, PptxSceneLineStyle sceneStyleLine, XElement? gridlines, PptxTheme theme)
    {
        if (sceneAxis is not null)
        {
            return ToChartSeriesStroke(sceneLine, null) ?? ToChartSeriesStroke(sceneStyleLine, null);
        }

        return ReadChartGridlineStroke(gridlines, theme);
    }

    private static ChartSeriesStroke? ReadChartGridlineStroke(XElement? gridlines, PptxTheme theme)
    {
        return ToChartSeriesStroke(PptxSceneBuilder.ReadChartGridlineLine(gridlines, theme), null);
    }

    private static ChartAxesStyle ReadSceneOrXmlChartAxesStyle(PptxSceneChart? sceneChart, PptxSceneChartPlot? plot, XDocument chartXml, PptxTheme theme, XElement chartElement)
    {
        ChartAxisSource valueAxisSource = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, plot, chartXml, chartElement).FirstOrDefault();
        XElement? valueAxisElement = ResolveXmlValueAxisForSource(sceneChart, valueAxisSource, chartXml);
        ChartAxisSource categoryAxisSource = ReadSceneOrXmlChartCategoryAxisForPlot(sceneChart, plot, chartXml, chartElement);
        XElement? categoryAxisElement = categoryAxisSource.XmlAxis;
        ChartAxisSource secondaryValueAxisSource = ReadSceneOrXmlSecondaryValueAxisForChart(sceneChart, chartXml, valueAxisSource);
        XElement? secondaryValueAxisElement = secondaryValueAxisSource.XmlAxis;
        PptxSceneChartAxis? valueAxis = valueAxisSource.SceneAxis;
        PptxSceneChartAxis? categoryAxis = categoryAxisSource.SceneAxis;
        PptxSceneChartAxis? secondaryValueAxis = secondaryValueAxisSource.SceneAxis;
        return new ChartAxesStyle(
            ReadSceneOrXmlChartAxisStroke(valueAxis, valueAxisElement, theme),
            ReadSceneOrXmlChartAxisStroke(secondaryValueAxis, secondaryValueAxisElement, theme),
            ReadSceneOrXmlChartAxisStroke(categoryAxis, categoryAxisElement, theme),
            ResolveSceneOrXmlValueAxisRightSide(valueAxis, valueAxisElement, defaultRightSide: false),
            ResolveSceneOrXmlValueAxisRightSide(secondaryValueAxis, secondaryValueAxisElement, defaultRightSide: true),
            ResolveSceneOrXmlValueAxisBottomSide(valueAxis, valueAxisElement, defaultBottomSide: true),
            ResolveSceneOrXmlCategoryAxisRightSide(categoryAxis, categoryAxisElement, defaultRightSide: false),
            valueAxis is null ? !IsChartAxisDeleted(valueAxisElement) : valueAxis.IsDeleted != true,
            categoryAxis is null ? !IsChartAxisDeleted(categoryAxisElement) : categoryAxis.IsDeleted != true,
            ReadSceneOrXmlChartAxisMajorTickMark(categoryAxis, categoryAxisElement));
    }

    private static ChartSeriesStroke? ReadSceneOrXmlChartAxisStroke(PptxSceneChartAxis? sceneAxis, XElement? xmlAxis, PptxTheme theme)
    {
        return sceneAxis is not null
            ? ToChartSeriesStroke(sceneAxis.Line, null)
            : ReadChartAxisStroke(xmlAxis, theme);
    }

    private static PptxSceneChartAxisTickMark ReadSceneOrXmlChartAxisMajorTickMark(PptxSceneChartAxis? sceneAxis, XElement? xmlAxis)
    {
        if (sceneAxis is not null)
        {
            return ResolveChartAxisTickMark(sceneAxis.MajorTickMarkKind);
        }

        string majorTickMark = PptxSceneBuilder.ReadChartElementValue(xmlAxis, "majorTickMark");
        return ResolveChartAxisTickMark(PptxSceneBuilder.ParseChartAxisTickMark(
            string.IsNullOrEmpty(majorTickMark) ? "none" : majorTickMark));

        PptxSceneChartAxisTickMark ResolveChartAxisTickMark(PptxSceneChartAxisTickMark tickMark)
        {
            return tickMark == PptxSceneChartAxisTickMark.Unknown
                ? PptxSceneChartAxisTickMark.None
                : tickMark;
        }
    }

    private static PptxSceneChartTickLabelPosition ResolveChartTickLabelPosition(PptxSceneChartTickLabelPosition position)
    {
        return position == PptxSceneChartTickLabelPosition.Unknown
            ? PptxSceneChartTickLabelPosition.NextTo
            : position;
    }

    private static bool ResolveSceneOrXmlValueAxisRightSide(PptxSceneChartAxis? sceneAxis, XElement? axis, bool defaultRightSide)
    {
        if (sceneAxis is not null)
        {
            return sceneAxis.PositionKind switch
            {
                PptxSceneChartAxisPosition.Right => true,
                PptxSceneChartAxisPosition.Left => false,
                _ => defaultRightSide
            };
        }

        string position = PptxSceneBuilder.ReadChartElementValue(axis, "axPos");
        return PptxSceneBuilder.ParseChartAxisPosition(position) switch
        {
            PptxSceneChartAxisPosition.Right => true,
            PptxSceneChartAxisPosition.Left => false,
            _ => defaultRightSide
        };
    }

    private static bool ResolveSceneOrXmlValueAxisBottomSide(PptxSceneChartAxis? sceneAxis, XElement? axis, bool defaultBottomSide)
    {
        if (sceneAxis is not null)
        {
            return sceneAxis.PositionKind switch
            {
                PptxSceneChartAxisPosition.Bottom => true,
                PptxSceneChartAxisPosition.Top => false,
                _ => defaultBottomSide
            };
        }

        string position = PptxSceneBuilder.ReadChartElementValue(axis, "axPos");
        return PptxSceneBuilder.ParseChartAxisPosition(position) switch
        {
            PptxSceneChartAxisPosition.Bottom => true,
            PptxSceneChartAxisPosition.Top => false,
            _ => defaultBottomSide
        };
    }

    private static bool ResolveSceneOrXmlCategoryAxisRightSide(PptxSceneChartAxis? sceneAxis, XElement? axis, bool defaultRightSide)
    {
        if (sceneAxis is not null)
        {
            return sceneAxis.PositionKind switch
            {
                PptxSceneChartAxisPosition.Right => true,
                PptxSceneChartAxisPosition.Left => false,
                _ => defaultRightSide
            };
        }

        string position = PptxSceneBuilder.ReadChartElementValue(axis, "axPos");
        return PptxSceneBuilder.ParseChartAxisPosition(position) switch
        {
            PptxSceneChartAxisPosition.Right => true,
            PptxSceneChartAxisPosition.Left => false,
            _ => defaultRightSide
        };
    }

    private static bool ResolveSceneOrXmlCategoryAxisTopSide(PptxSceneChartAxis? sceneAxis, XElement? axis, bool defaultTopSide)
    {
        if (sceneAxis is not null)
        {
            return sceneAxis.PositionKind switch
            {
                PptxSceneChartAxisPosition.Top => true,
                PptxSceneChartAxisPosition.Bottom => false,
                _ => defaultTopSide
            };
        }

        string position = PptxSceneBuilder.ReadChartElementValue(axis, "axPos");
        return PptxSceneBuilder.ParseChartAxisPosition(position) switch
        {
            PptxSceneChartAxisPosition.Top => true,
            PptxSceneChartAxisPosition.Bottom => false,
            _ => defaultTopSide
        };
    }

    private static bool IsChartAxisDeleted(XElement? axis)
    {
        XElement? delete = axis?.Element(ChartNamespace + "delete");
        return PptxSceneBuilder.IsOoxmlBooleanElementEnabled(delete);
    }

    private static bool IsSceneOrXmlChartAxisLabelVisible(PptxSceneChartAxis? sceneAxis, XElement? axis)
    {
        if (sceneAxis is null)
        {
            return IsChartAxisLabelVisible();
        }

        return sceneAxis.IsDeleted != true &&
            ResolveChartTickLabelPosition(sceneAxis.TickLabelPositionKind) != PptxSceneChartTickLabelPosition.None;

        bool IsChartAxisLabelVisible()
        {
            if (IsChartAxisDeleted(axis))
            {
                return false;
            }

            string tickLabelPosition = PptxSceneBuilder.ReadChartElementValue(axis, "tickLblPos");
            return ResolveChartTickLabelPosition(PptxSceneBuilder.ParseChartTickLabelPosition(tickLabelPosition)) != PptxSceneChartTickLabelPosition.None;
        }
    }

    private static bool ResolveSceneOrXmlValueAxisLabelsRightSide(PptxSceneChartAxis? sceneAxis, XElement? axis, bool defaultRightSide)
    {
        if (sceneAxis is null)
        {
            return ResolveValueAxisLabelsRightSide(axis, defaultRightSide);
        }

        return ResolveChartTickLabelPosition(sceneAxis.TickLabelPositionKind) switch
        {
            PptxSceneChartTickLabelPosition.High => true,
            PptxSceneChartTickLabelPosition.Low => false,
            _ => defaultRightSide
        };

        bool ResolveValueAxisLabelsRightSide(XElement? axis, bool defaultRightSide)
        {
            string tickLabelPosition = PptxSceneBuilder.ReadChartElementValue(axis, "tickLblPos");
            return ResolveChartTickLabelPosition(PptxSceneBuilder.ParseChartTickLabelPosition(tickLabelPosition)) switch
            {
                PptxSceneChartTickLabelPosition.High => true,
                PptxSceneChartTickLabelPosition.Low => false,
                _ => defaultRightSide
            };
        }
    }

    private static double ResolveSceneOrXmlCategoryAxisLabelOffsetScale(PptxSceneChartAxis? sceneAxis, XElement? axis)
    {
        int offset = sceneAxis is not null
            ? sceneAxis.LabelOffset ?? PptxChartMetricRules.CategoryAxisDefaultLabelOffset
            : PptxSceneBuilder.ReadChartElementIntWithValue(axis, "lblOffset").Value ?? PptxChartMetricRules.CategoryAxisDefaultLabelOffset;

        offset = Math.Clamp(
            offset,
            PptxChartMetricRules.CategoryAxisMinimumLabelOffset,
            PptxChartMetricRules.CategoryAxisMaximumLabelOffset);

        return offset / (double)PptxChartMetricRules.CategoryAxisDefaultLabelOffset;
    }

    private static int ResolveSceneOrXmlCategoryAxisTickLabelSkip(PptxSceneChartAxis? sceneAxis, XElement? axis)
    {
        int skip = sceneAxis is not null
            ? sceneAxis.TickLabelSkip ?? PptxChartMetricRules.CategoryAxisDefaultTickLabelSkip
            : PptxSceneBuilder.ReadChartElementIntWithValue(axis, "tickLblSkip").Value ?? PptxChartMetricRules.CategoryAxisDefaultTickLabelSkip;

        return Math.Max(PptxChartMetricRules.CategoryAxisDefaultTickLabelSkip, skip);
    }

    private static bool ResolveSceneOrXmlCategoryAxisLabelsOnTickMarks(PptxSceneChartAxis? sceneAxis, XElement? axis)
    {
        if (sceneAxis is not null)
        {
            return ResolveChartAxisCrossBetween(sceneAxis.CrossBetweenKind) == PptxSceneChartAxisCrossBetween.MidpointCategory;
        }

        string crossBetween = PptxSceneBuilder.ReadChartElementValue(axis, "crossBetween");
        return ResolveChartAxisCrossBetween(PptxSceneBuilder.ParseChartAxisCrossBetween(crossBetween)) == PptxSceneChartAxisCrossBetween.MidpointCategory;

        PptxSceneChartAxisCrossBetween ResolveChartAxisCrossBetween(PptxSceneChartAxisCrossBetween crossBetween)
        {
            return crossBetween == PptxSceneChartAxisCrossBetween.Unknown
                ? PptxSceneChartAxisCrossBetween.Between
                : crossBetween;
        }
    }

    private static ChartSeriesStroke? ReadChartAxisStroke(XElement? axis, PptxTheme theme)
    {
        return axis is null
            ? null
            : ToChartSeriesStroke(PptxSceneBuilder.ReadChartAxisLine(axis, theme), null);
    }

    private static IReadOnlyList<ChartIndexedScatterSeries> ReadScatterSeriesVectors(XElement chartElement, bool readBubbleSize, ChartWorkbookData? workbook, bool plotVisibleOnly)
    {
        var series = new List<ChartIndexedScatterSeries>();
        foreach (XElement element in chartElement.Elements(ChartNamespace + "ser"))
        {
            ChartIndexedNumberVector xValues = ReadChartNumberVector(element.Element(ChartNamespace + "xVal"), workbook, plotVisibleOnly);
            ChartIndexedNumberVector yValues = ReadChartNumberVector(element.Element(ChartNamespace + "yVal"), workbook, plotVisibleOnly);
            ChartIndexedNumberVector bubbleSizes = readBubbleSize
                ? ReadChartNumberVector(element.Element(ChartNamespace + "bubbleSize"), workbook, plotVisibleOnly)
                : default;
            if (xValues.DensePoints().Count == 0 && yValues.DensePoints().Count == 0)
            {
                continue;
            }

            series.Add(new ChartIndexedScatterSeries(xValues, yValues, bubbleSizes, readBubbleSize));
        }

        return series;
    }
}
