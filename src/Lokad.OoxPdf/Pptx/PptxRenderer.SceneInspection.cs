namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    internal static PptxSceneNodeSnapshot InspectSceneNode(PptxSceneNode node)
    {
        int textParagraphCount = node.TextBody?.Paragraphs.Count ?? 0;
        int textRunCount = node.TextBody?.Paragraphs.Sum(paragraph => paragraph.Runs.Count) ?? 0;
        int tableRowCount = node.Table?.Rows.Count ?? 0;
        int tableCellCount = node.Table?.Rows.Sum(row => row.Cells.Count) ?? 0;
        PptxSceneBounds? bounds = node.Bounds;
        PptxSceneImageRecolor? recolor = node.Picture?.Recolor;
        IReadOnlyList<PptxSceneChartPlot> chartPlots = node.Chart?.Plots ?? [];
        IReadOnlyList<PptxSceneChartSeries> chartSeries = chartPlots
            .SelectMany(plot => plot.Series)
            .ToArray();
        IReadOnlyList<PptxSceneChartTextBodyProperties> chartTextBodyProperties = ReadChartTextBodyProperties(node.Chart);
        IReadOnlyList<PptxSceneChartColorDeclaration> chartColorDeclarations =
            node.Chart?.ColorStyle.Declarations ?? [];
        IReadOnlyList<string> chartColorDeclarationKinds = chartColorDeclarations
            .Select(declaration => declaration.Kind)
            .ToArray();
        IReadOnlyList<PptxSceneChartStyleEntry> chartStyleEntries = node.Chart?.StylePart.Entries ?? [];
        IReadOnlyList<string> chartStyleEntryRoles = chartStyleEntries
            .Select(entry => entry.Role)
            .ToArray();
        IReadOnlyList<int> chartStyleEntrySourceIndexes = chartStyleEntries
            .Select(entry => entry.SourceIndex)
            .ToArray();
        IReadOnlyList<string> chartStyleEntryNamespaceUris = chartStyleEntries
            .Select(entry => entry.NamespaceUri)
            .ToArray();
        return new PptxSceneNodeSnapshot(
            node.Kind.ToString(),
            node.IsPlaceholder,
            bounds is not null,
            bounds?.RotationDegrees ?? 0d,
            bounds?.FlipHorizontal ?? false,
            bounds?.FlipVertical ?? false,
            node.Shape?.Preset ?? string.Empty,
            node.Shape?.HasCustomGeometry ?? false,
            node.Shape?.CustomGeometry.HasUnsupportedGeometry ?? false,
            node.Shape?.GradientFill.HasGradientSource ?? false,
            node.Shape?.GradientFill.HasUnsupportedGradient ?? false,
            node.Shape?.PatternFill.HasPatternSource ?? false,
            node.Shape?.PatternFill.HasUnsupportedPattern ?? false,
            node.Shape?.NoFill ?? false,
            node.Shape?.LineNoFill ?? false,
            node.Shape?.Effects.HasEffectList ?? false,
            node.Shape?.Effects.HasEffectDag ?? false,
            node.Shape?.Effects.UnsupportedEffectNames.Count ?? 0,
            node.Shape?.Effects.UnsupportedEffectNames ?? [],
            node.TextBody is not null,
            textParagraphCount,
            textRunCount,
            node.TextBody?.HasUnsupportedTextOrientation ?? false,
            node.TextBody?.HasUnsupportedVerticalOverflow ?? false,
            node.Picture is not null,
            recolor?.Kind.ToString() ?? string.Empty,
            node.Table is not null,
            tableRowCount,
            tableCellCount,
            node.Chart is not null,
            node.Chart?.Plots.Count ?? 0,
            node.Chart?.Axes.Count ?? 0,
            chartSeries.Count,
            chartSeries.Count(series => series.Marker.IsDefined),
            chartSeries.Sum(series => series.PointStyles.Count),
            chartSeries.Sum(series => series.PointStyles.Count(point => point.Explosion is not null)) +
                chartSeries.Count(series => series.Explosion is not null),
            chartPlots.Count(plot => plot.DataLabels.IsDefined) +
                chartSeries.Count(series => series.DataLabels.IsDefined),
            chartPlots.Sum(plot => plot.DataLabels.Overrides.Count) +
                chartSeries.Sum(series => series.DataLabels.Overrides.Count),
            chartPlots.Count(plot => plot.DataLabels.Layout.HasLayout) +
                chartSeries.Count(series => series.DataLabels.Layout.HasLayout) +
                chartPlots.Sum(plot => plot.DataLabels.Overrides.Count(label => label.Layout.HasLayout)) +
                chartSeries.Sum(series => series.DataLabels.Overrides.Count(label => label.Layout.HasLayout)),
            chartTextBodyProperties.Count(properties => !string.IsNullOrWhiteSpace(properties.OrientationValue)),
            chartTextBodyProperties.Count(properties => !string.IsNullOrWhiteSpace(properties.VerticalOverflowValue)),
            node.Chart?.ColorStyle.IsDefined ?? false,
            node.Chart?.ColorStyle.Method ?? string.Empty,
            node.Chart?.ColorStyle.Colors.Count ?? 0,
            node.Chart?.ColorStyle.VariationCount ?? 0,
            chartColorDeclarations.Count,
            node.Chart?.ColorStyle.RootDeclarations.Count ?? 0,
            chartColorDeclarations.Count(declaration => declaration.IsResolved),
            chartColorDeclarationKinds,
            node.Chart?.StylePart.IsDefined ?? false,
            chartStyleEntries.Count,
            chartStyleEntryRoles,
            chartStyleEntrySourceIndexes,
            chartStyleEntryNamespaceUris,
            chartStyleEntries.Count(entry => HasChartShapeStyle(entry.ShapeStyle)),
            chartStyleEntries.Count(entry => entry.ShapeStyle?.Fill.HasFill == true ||
                entry.ShapeStyle?.PatternFill.HasPattern == true ||
                entry.ShapeStyle?.GradientFill?.HasGradient == true),
            chartStyleEntries.Count(entry => entry.FillReferenceIndex is not null),
            chartStyleEntries.Count(entry => entry.FillReferenceFill.HasFill),
            chartStyleEntries.Count(entry => entry.EffectReferenceIndex is not null),
            chartStyleEntries.Count(entry => entry.EffectReferenceEffects.HasEffectList ||
                entry.EffectReferenceEffects.HasEffectDag ||
                entry.EffectReferenceEffects.UnsupportedEffectNames?.Count > 0),
            chartStyleEntries.Count(entry => !string.IsNullOrWhiteSpace(entry.FontReferenceIndex)),
            node.Kind == PptxSceneNodeKind.Group);
    }

    private static IReadOnlyList<PptxSceneChartTextBodyProperties> ReadChartTextBodyProperties(PptxSceneChart? chart)
    {
        if (chart is null)
        {
            return [];
        }

        var properties = new List<PptxSceneChartTextBodyProperties>
        {
            chart.Title.TextBodyProperties,
            chart.Legend.TextBodyProperties
        };

        foreach (PptxSceneChartAxis axis in chart.Axes)
        {
            properties.Add(axis.TextBodyProperties);
            properties.Add(axis.Title.TextBodyProperties);
        }

        foreach (PptxSceneChartPlot plot in chart.Plots)
        {
            AddChartDataLabelTextBodyProperties(properties, plot.DataLabels);
            foreach (PptxSceneChartSeries series in plot.Series)
            {
                AddChartDataLabelTextBodyProperties(properties, series.DataLabels);
            }
        }

        return properties;
    }

    private static void AddChartDataLabelTextBodyProperties(
        List<PptxSceneChartTextBodyProperties> properties,
        PptxSceneChartDataLabels labels)
    {
        properties.Add(labels.TextBodyProperties);
        foreach (PptxSceneChartDataLabelOverride labelOverride in labels.Overrides)
        {
            properties.Add(labelOverride.TextBodyProperties);
        }
    }

    private static bool HasChartShapeStyle(PptxSceneChartShapeStyle? style)
    {
        return style is not null &&
            (style.NoFill ||
                style.Fill.HasFill ||
                style.GradientFill?.HasGradient == true ||
                style.PatternFill.HasPattern ||
                style.Line.HasLine ||
                style.Glow.HasGlow ||
                style.OuterShadow.HasShadow);
    }
}
