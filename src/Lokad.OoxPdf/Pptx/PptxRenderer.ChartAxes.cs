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
        return ReadChartValueAxisExtents(valueAxis, fallback, PptxChartMetricRules.BubbleAxisBoundsTickTargetCount, true, PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio, preferUnitOneOverTwo: true);
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
        double max = axis.Maximum ?? GetNiceChartAxisMax(fallback.Max, min, PptxChartMetricRules.BubbleAxisBoundsTickTargetCount, true, PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio, preferUnitOneOverTwo: true);
        return max > min
            ? new ChartValueExtents(min, max)
            : fallback;
    }

    private static ChartValueExtents ReadChartValueAxisExtents(XElement? valueAxis, ChartValueExtents fallback, double boundsTickTargetCount, bool useNearMaximumHeadroom, double nearMaximumHeadroomRatio, bool preferUnitOneOverTwo = false)
    {
        XElement? scaling = valueAxis?.Element(ChartNamespace + "scaling");
        if (valueAxis is null || scaling is null)
        {
            return fallback;
        }

        double min = PptxSceneBuilder.ReadChartAxisScalingValueWithValue(valueAxis, "min").Value ?? GetNiceChartAxisMin(fallback.Min, fallback.Max);
        double max = PptxSceneBuilder.ReadChartAxisScalingValueWithValue(valueAxis, "max").Value ?? GetNiceChartAxisMax(fallback.Max, min, boundsTickTargetCount, useNearMaximumHeadroom, nearMaximumHeadroomRatio, preferUnitOneOverTwo);
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

}
