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
}
