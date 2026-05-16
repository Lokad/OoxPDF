using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static void RenderCharts(PptxRenderContext context, PdfGraphicsBuilder graphics)
    {
        foreach (XElement frame in context.SlideXml.Descendants(PresentationNamespace + "graphicFrame"))
        {
            XElement? graphicData = frame
                .Element(DrawingNamespace + "graphic")
                ?.Element(DrawingNamespace + "graphicData");
            if (graphicData?.Attribute("uri") is not { } uri ||
                !uri.Value.Contains("drawingml/2006/chart", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            ShapeBounds? bounds = ReadGraphicFrameBounds(frame);
            string? relationshipId = (string?)graphicData
                .Element(ChartNamespace + "chart")
                ?.Attribute(RelationshipsNamespace + "id");
            if (bounds is null || relationshipId is null || !context.SlideRelationships.TryGetValue(relationshipId, out OoxRelationship? relationship) || relationship.ResolvedTarget is null)
            {
                EmitChartDiagnostic(context.DiagnosticSink, "PPTX_UNSUPPORTED_CHART", OoxPdfSeverity.Warning, "Chart frame could not be resolved and was ignored.", context.Slide.PartName, context.SlideNumber, "Ignored");
                continue;
            }

            OoxPart? chartPart = context.Package.GetPart(relationship.ResolvedTarget);
            if (chartPart is null)
            {
                EmitChartDiagnostic(context.DiagnosticSink, "PPTX_UNSUPPORTED_CHART", OoxPdfSeverity.Warning, "Chart part was missing and was ignored.", relationship.ResolvedTarget, context.SlideNumber, "Ignored");
                continue;
            }

            using Stream chartStream = chartPart.OpenRead();
            XDocument chartXml = SafeXml.Load(chartStream);
            if (TryRenderChartFallback(graphics, context.Document, bounds.Value, chartXml))
            {
                EmitChartDiagnostic(context.DiagnosticSink, "PPTX_CHART_STATIC_FALLBACK", OoxPdfSeverity.Info, "PPTX chart was rendered with an approximate static chart fallback.", chartPart.Name, context.SlideNumber, "Static chart fallback");
                continue;
            }

            EmitChartDiagnostic(context.DiagnosticSink, "PPTX_UNSUPPORTED_CHART", OoxPdfSeverity.Warning, "Only bar, line, and pie chart cached numeric values have a static fallback.", chartPart.Name, context.SlideNumber, "Ignored");
        }
    }

    private static bool TryRenderChartFallback(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, XDocument chartXml)
    {
        XElement? barChart = chartXml.Descendants(ChartNamespace + "barChart").FirstOrDefault();
        if (barChart is not null)
        {
            IReadOnlyList<IReadOnlyList<double>> barSeries = ReadChartSeries(barChart);
            if (barSeries.Count != 0)
            {
                bool horizontalBars = string.Equals((string?)barChart.Element(ChartNamespace + "barDir")?.Attribute("val"), "bar", StringComparison.Ordinal);
                RenderBarChartFallback(graphics, document, bounds, barSeries, horizontalBars);
                return true;
            }
        }

        IReadOnlyList<IReadOnlyList<double>> lineSeries = ReadChartSeries(chartXml, "lineChart");
        if (lineSeries.Count != 0)
        {
            RenderLineChartFallback(graphics, document, bounds, lineSeries);
            return true;
        }

        IReadOnlyList<IReadOnlyList<double>> pieSeries = ReadChartSeries(chartXml, "pieChart");
        if (pieSeries.Count != 0)
        {
            RenderPieChartFallback(graphics, document, bounds, pieSeries[0]);
            return true;
        }

        return false;
    }

    private static IReadOnlyList<IReadOnlyList<double>> ReadChartSeries(XDocument chartXml, string chartElementName)
    {
        XElement? chartElement = chartXml.Descendants(ChartNamespace + chartElementName).FirstOrDefault();
        return chartElement is null ? [] : ReadChartSeries(chartElement);
    }

    private static IReadOnlyList<IReadOnlyList<double>> ReadChartSeries(XElement chartElement)
    {
        var series = new List<IReadOnlyList<double>>();
        foreach (XElement element in chartElement.Elements(ChartNamespace + "ser"))
        {
            double[] values = element
                .Elements(ChartNamespace + "val")
                .Descendants(ChartNamespace + "pt")
                .Select(point => (string?)point.Element(ChartNamespace + "v"))
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) ? parsed : double.NaN)
                .Where(value => !double.IsNaN(value))
                .ToArray();
            if (values.Length > 0)
            {
                series.Add(values);
            }
        }

        return series;
    }

    private static RgbColor ChartPalette(int index)
    {
        RgbColor[] palette =
        [
            new RgbColor(68, 114, 196),
            new RgbColor(237, 125, 49),
            new RgbColor(165, 165, 165),
            new RgbColor(255, 192, 0),
            new RgbColor(91, 155, 213),
            new RgbColor(112, 173, 71)
        ];
        return palette[index % palette.Length];
    }

    private static void RenderBarChartFallback(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, IReadOnlyList<IReadOnlyList<double>> series, bool horizontalBars)
    {
        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        double plotX = x + width * 0.1d;
        double plotY = y + height * 0.14d;
        double plotWidth = width * 0.82d;
        double plotHeight = height * 0.72d;
        int categoryCount = Math.Max(1, series.Max(values => values.Count));
        double maxValue = Math.Max(1d, series.SelectMany(values => values).DefaultIfEmpty(1d).Max(value => Math.Abs(value)));

        graphics.SetStrokeRgb(90, 90, 90);
        graphics.SetLineWidth(0.75d);
        graphics.StrokeLine(plotX, plotY, plotX + plotWidth, plotY);
        graphics.StrokeLine(plotX, plotY, plotX, plotY + plotHeight);

        if (horizontalBars)
        {
            RenderHorizontalBars(graphics, plotX, plotY, plotWidth, plotHeight, series, categoryCount, maxValue);
            return;
        }

        double categoryWidth = plotWidth / categoryCount;
        double barSlot = categoryWidth * 0.82d / Math.Max(1, series.Count);
        for (int category = 0; category < categoryCount; category++)
        {
            double categoryX = plotX + category * categoryWidth + categoryWidth * 0.09d;
            for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
            {
                IReadOnlyList<double> values = series[seriesIndex];
                if (category >= values.Count)
                {
                    continue;
                }

                double value = Math.Max(0d, values[category]);
                double barHeight = value / maxValue * plotHeight;
                RgbColor fill = ChartPalette(seriesIndex);
                graphics.SetFillRgb(fill.Red, fill.Green, fill.Blue);
                graphics.FillRectangle(categoryX + seriesIndex * barSlot, plotY, Math.Max(0.5d, barSlot * 0.86d), barHeight);
            }
        }
    }

    private static void RenderHorizontalBars(PdfGraphicsBuilder graphics, double plotX, double plotY, double plotWidth, double plotHeight, IReadOnlyList<IReadOnlyList<double>> series, int categoryCount, double maxValue)
    {
        double categoryHeight = plotHeight / categoryCount;
        double barSlot = categoryHeight * 0.82d / Math.Max(1, series.Count);
        for (int category = 0; category < categoryCount; category++)
        {
            double categoryY = plotY + plotHeight - (category + 1) * categoryHeight + categoryHeight * 0.09d;
            for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
            {
                IReadOnlyList<double> values = series[seriesIndex];
                if (category >= values.Count)
                {
                    continue;
                }

                double value = Math.Max(0d, values[category]);
                double barWidth = value / maxValue * plotWidth;
                RgbColor fill = ChartPalette(seriesIndex);
                graphics.SetFillRgb(fill.Red, fill.Green, fill.Blue);
                graphics.FillRectangle(plotX, categoryY + seriesIndex * barSlot, barWidth, Math.Max(0.5d, barSlot * 0.86d));
            }
        }
    }

    private static void RenderLineChartFallback(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, IReadOnlyList<IReadOnlyList<double>> series)
    {
        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        double plotX = x + width * 0.12d;
        double plotY = y + height * 0.16d;
        double plotWidth = width * 0.76d;
        double plotHeight = height * 0.68d;
        int pointCount = Math.Max(1, series.Max(values => values.Count));
        double maxValue = Math.Max(1d, series.SelectMany(values => values).DefaultIfEmpty(1d).Max());
        double minValue = Math.Min(0d, series.SelectMany(values => values).DefaultIfEmpty(0d).Min());
        double valueRange = Math.Max(1d, maxValue - minValue);

        graphics.SetStrokeRgb(90, 90, 90);
        graphics.SetLineWidth(0.75d);
        graphics.StrokeLine(plotX, plotY, plotX + plotWidth, plotY);
        graphics.StrokeLine(plotX, plotY, plotX, plotY + plotHeight);

        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
        {
            IReadOnlyList<double> values = series[seriesIndex];
            if (values.Count == 0)
            {
                continue;
            }

            RgbColor stroke = ChartPalette(seriesIndex);
            graphics.SetStrokeRgb(stroke.Red, stroke.Green, stroke.Blue);
            graphics.SetLineWidth(1.5d);
            double previousX = 0d;
            double previousY = 0d;
            for (int i = 0; i < values.Count; i++)
            {
                double pointX = plotX + (pointCount == 1 ? plotWidth / 2d : plotWidth * i / (pointCount - 1));
                double pointY = plotY + (values[i] - minValue) / valueRange * plotHeight;
                if (i > 0)
                {
                    graphics.StrokeLine(previousX, previousY, pointX, pointY);
                }

                graphics.SetFillRgb(stroke.Red, stroke.Green, stroke.Blue);
                graphics.FillEllipse(pointX - 2d, pointY - 2d, 4d, 4d);
                previousX = pointX;
                previousY = pointY;
            }
        }
    }

    private static void RenderPieChartFallback(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, IReadOnlyList<double> values)
    {
        double total = values.Where(value => value > 0d).Sum();
        if (total <= 0d)
        {
            return;
        }

        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        double radius = Math.Min(width, height) * 0.34d;
        double centerX = x + width * 0.46d;
        double centerY = y + height * 0.52d;
        double angle = -Math.PI / 2d;

        for (int i = 0; i < values.Count; i++)
        {
            double value = Math.Max(0d, values[i]);
            if (value <= 0d)
            {
                continue;
            }

            double sweep = value / total * Math.PI * 2d;
            int segments = Math.Max(2, (int)Math.Ceiling(Math.Abs(sweep) / (Math.PI / 18d)));
            var points = new (double X, double Y)[segments + 2];
            points[0] = (centerX, centerY);
            for (int segment = 0; segment <= segments; segment++)
            {
                double segmentAngle = angle + sweep * segment / segments;
                points[segment + 1] = (
                    centerX + Math.Cos(segmentAngle) * radius,
                    centerY + Math.Sin(segmentAngle) * radius);
            }

            RgbColor fill = ChartPalette(i);
            graphics.SetFillRgb(fill.Red, fill.Green, fill.Blue);
            graphics.FillPolygon(points);
            angle += sweep;
        }
    }

    private static void EmitChartDiagnostic(Action<OoxPdfDiagnostic>? diagnosticSink, string id, OoxPdfSeverity severity, string message, string? partName, int slideIndex, string fallback)
    {
        diagnosticSink?.Invoke(new OoxPdfDiagnostic(
            id,
            severity,
            message,
            partName,
            SlideIndex: slideIndex,
            Feature: "chart",
            Fallback: fallback));
    }
}
