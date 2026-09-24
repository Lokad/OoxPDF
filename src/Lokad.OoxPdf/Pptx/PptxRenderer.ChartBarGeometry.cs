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
    private static (double Min, double Max) GetClusteredPointValueExtents(IReadOnlyList<IReadOnlyList<double?>> series)
    {
        double maxValue = Math.Max(0d, series.SelectMany(values => values).OfType<double>().DefaultIfEmpty(0d).Max());
        double minValue = Math.Min(0d, series.SelectMany(values => values).OfType<double>().DefaultIfEmpty(0d).Min());
        return (minValue, maxValue);
    }

    private static (double Min, double Max) GetStackedPointValueExtents(IReadOnlyList<IReadOnlyList<double?>> series, int categoryCount, bool percentStacked)
    {
        if (percentStacked)
        {
            return (0d, 1d);
        }

        double minValue = 0d;
        double maxValue = 0d;
        for (int category = 0; category < categoryCount; category++)
        {
            double positive = 0d;
            double negative = 0d;
            foreach (IReadOnlyList<double?> values in series)
            {
                if (category >= values.Count || values[category] is not { } value)
                {
                    continue;
                }

                if (value >= 0d)
                {
                    positive += value;
                }
                else
                {
                    negative += value;
                }
            }

            maxValue = Math.Max(maxValue, positive);
            minValue = Math.Min(minValue, negative);
        }

        return (minValue, maxValue);
    }

    private static double GetStackedBarWidth(double categoryBand, double gapWidthPercent)
    {
        return Math.Max(0.5d, categoryBand * 100d / (100d + Math.Max(0d, gapWidthPercent)));
    }

    // Clustered width shares the gap divisor with the step overlap term: Office composite
    // vectors read barW 23.64 with step 30.0 on band 135.45, gap 219, overlap -27, i.e. a
    // 573 divisor (step honors overlap at 30/23.64 = 1.2691), so separation widens the
    // effective gap on both flanks. Zero overlap keeps the legacy divisor exactly.
    private static double GetClusteredBarWidth(double categoryBand, int seriesCount, double gapWidthPercent, double overlapPercent)
    {
        int count = Math.Max(1, seriesCount);
        return Math.Max(0.5d, categoryBand * 100d / (100d * count + Math.Max(0d, gapWidthPercent) - 2d * overlapPercent));
    }

    private static double GetClusteredBarStep(double barWidth, double overlapPercent)
    {
        return Math.Max(0d, barWidth * (1d - overlapPercent / 100d));
    }

    private static double GetCategoryPositiveTotal(IReadOnlyList<IReadOnlyList<double?>> series, int category, bool percentStacked)
    {
        if (!percentStacked)
        {
            return 1d;
        }

        double total = 0d;
        foreach (IReadOnlyList<double?> values in series)
        {
            if (category < values.Count && values[category] is { } value)
            {
                total += Math.Max(0d, value);
            }
        }

        return Math.Max(1d, total);
    }

    private static double[] GetCategoryPositiveTotals(IReadOnlyList<IReadOnlyList<double?>> series, int categoryCount, bool percentStacked)
    {
        var totals = new double[categoryCount];
        if (!percentStacked)
        {
            Array.Fill(totals, 1d);
            return totals;
        }

        for (int category = 0; category < categoryCount; category++)
        {
            totals[category] = GetCategoryPositiveTotal(series, category, percentStacked);
        }

        return totals;
    }

    private static double NormalizeStackedValue(double value, double positiveTotal, bool percentStacked)
    {
        return percentStacked && value > 0d ? value / positiveTotal : value;
    }
}
