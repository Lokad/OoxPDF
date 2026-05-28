using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    internal static PptxSceneSnapshot InspectScene(PptxDocument document, OoxPackage package)
    {
        PptxScene scene = new PptxSceneBuilder().Build(document, package);
        return new PptxSceneSnapshot(scene.Slides.Select(ToSnapshot).ToArray());
    }

    private static PptxSceneSlideSnapshot ToSnapshot(PptxSceneSlide slide)
    {
        return new PptxSceneSlideSnapshot(
            slide.Index,
            slide.PartName,
            slide.MasterPartName,
            slide.LayoutPartName,
            slide.MasterXml is not null,
            slide.LayoutXml is not null,
            slide.SlideXml.Root is not null,
            slide.MasterRelationships.Count,
            slide.LayoutRelationships.Count,
            slide.SlideRelationships.Count,
            slide.MasterBackground.HasFill,
            slide.LayoutBackground.HasFill,
            slide.SlideBackground.HasFill,
            slide.HasTransition,
            slide.HasTiming,
            slide.HasOleObject,
            ToSnapshot(slide.MasterColorMap),
            ToSnapshot(slide.LayoutColorMap),
            ToSnapshot(slide.SlideColorMap),
            slide.MasterNodes.Select(ToSnapshot).ToArray(),
            slide.LayoutNodes.Select(ToSnapshot).ToArray(),
            slide.SlideNodes.Select(ToSnapshot).ToArray());
    }

    private static IReadOnlyDictionary<string, string> ToSnapshot(PptxColorMap colorMap)
    {
        var snapshot = new SortedDictionary<string, string>(StringComparer.Ordinal);
        foreach (KeyValuePair<string, string> mapping in colorMap.Mappings)
        {
            snapshot[mapping.Key] = mapping.Value;
        }

        return snapshot;
    }

    private static PptxSceneNodeSnapshot ToSnapshot(PptxSceneNode node)
    {
        int textParagraphCount = node.TextBody?.Paragraphs.Count ?? 0;
        int textRunCount = node.TextBody?.Paragraphs.Sum(paragraph => paragraph.Runs.Count) ?? 0;
        int tableRowCount = node.Table?.Rows.Count ?? 0;
        int tableCellCount = node.Table?.Rows.Sum(row => row.Cells.Count) ?? 0;
        IReadOnlyList<PptxSceneTableCell> tableCells = node.Table?.Rows
            .SelectMany(row => row.Cells)
            .ToArray() ?? [];
        PptxSceneTableStyle tableStyle = node.Table?.Style ?? default;
        PptxSceneBounds? bounds = node.Bounds;
        PptxSceneImageRecolor? recolor = node.Picture?.Recolor;
        PptxSceneShapePictureFill shapePictureFill = node.Shape?.PictureFill ?? default;
        PptxSceneChartLegend? legend = node.Chart?.Legend;
        PptxSceneChartManualLayout? legendLayout = legend?.Layout;
        PptxSceneChartManualLayout? plotAreaLayout = node.Chart?.PlotAreaLayout;
        IReadOnlyList<PptxSceneChartPlot> chartPlots = node.Chart?.Plots ?? [];
        IReadOnlyList<PptxSceneChartSeries> chartSeries = chartPlots
            .SelectMany(plot => plot.Series)
            .ToArray();
        IReadOnlyList<PptxSceneChartColorDeclaration> chartColorDeclarations =
            node.Chart?.ColorStyle.Declarations ?? [];
        IReadOnlyList<string> chartColorDeclarationKinds = chartColorDeclarations
            .Select(declaration => declaration.Kind)
            .ToArray();
        IReadOnlyList<string> chartStyleEntryRoles = node.Chart?.StylePart.Entries
            .Select(entry => entry.Role)
            .ToArray() ?? [];
        IReadOnlyList<PptxSceneChartStyleEntry> chartStyleEntries = node.Chart?.StylePart.Entries ?? [];
        IReadOnlyList<int> chartStyleEntrySourceIndexes = chartStyleEntries
            .Select(entry => entry.SourceIndex)
            .ToArray();
        IReadOnlyList<string> rejectedPointStyleIndexValues = ReadRejectedPointStyleIndexValues(node.Chart);
        IReadOnlyList<string> rejectedDataLabelOverrideIndexValues = ReadRejectedDataLabelOverrideIndexValues(node.Chart);
        return new PptxSceneNodeSnapshot(
            node.Kind.ToString(),
            node.IsPlaceholder,
            node.IsSmartArtGraphicFrame,
            node.Kind == PptxSceneNodeKind.UnknownGraphicFrame && !node.IsSmartArtGraphicFrame,
            node.HyperlinkClick.IsDefined,
            node.HyperlinkClick.RelationshipId,
            node.HyperlinkClick.Action,
            bounds is not null,
            bounds?.RotationDegrees ?? 0d,
            bounds?.FlipHorizontal ?? false,
            bounds?.FlipVertical ?? false,
            node.Shape is not null,
            node.Shape?.Preset ?? string.Empty,
            node.Shape?.HasCustomGeometry ?? false,
            node.Shape?.CustomGeometry.HasUnsupportedGeometry ?? false,
            node.Shape?.GradientFill.HasGradientSource ?? false,
            node.Shape?.GradientFill.HasUnsupportedGradient ?? false,
            node.Shape?.PatternFill.HasPatternSource ?? false,
            node.Shape?.PatternFill.HasUnsupportedPattern ?? false,
            node.Shape?.HasUnsupportedTransparency ?? false,
            node.Shape?.NoFill ?? false,
            node.Shape?.FillReference.Index ?? 0,
            node.Shape?.FillReference.Style is not null,
            node.Shape?.LineNoFill ?? false,
            node.Shape?.LineReference.Index ?? 0,
            node.Shape?.LineReference.Style is not null,
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
            node.Picture?.Resource is not null,
            node.Picture?.HasVideo ?? false,
            node.Picture?.HasAudio ?? false,
            node.Picture?.Resource?.ContentType ?? string.Empty,
            node.Picture?.Alpha ?? 1d,
            node.Picture?.AlphaValue ?? string.Empty,
            shapePictureFill.Resource is not null,
            shapePictureFill.Resource?.ContentType ?? string.Empty,
            shapePictureFill.HasPicture ? shapePictureFill.Alpha : 1d,
            shapePictureFill.AlphaValue ?? string.Empty,
            recolor?.Kind.ToString() ?? string.Empty,
            recolor?.KindValue ?? string.Empty,
            recolor?.Kind == PptxSceneImageRecolorKind.Luminance ? recolor?.Brightness : null,
            recolor?.Kind == PptxSceneImageRecolorKind.Luminance ? recolor?.Contrast : null,
            recolor?.Kind == PptxSceneImageRecolorKind.BiLevel ? recolor?.Threshold : null,
            recolor?.BrightnessValue ?? string.Empty,
            recolor?.ContrastValue ?? string.Empty,
            recolor?.ThresholdValue ?? string.Empty,
            node.Table is not null,
            tableRowCount,
            tableCellCount,
            tableStyle.StyleId ?? string.Empty,
            tableStyle.Name,
            tableStyle.Accent,
            tableStyle.IsSupported,
            tableStyle.FirstRow,
            tableStyle.FirstRowValue,
            tableStyle.LastRow,
            tableStyle.LastRowValue,
            tableStyle.FirstColumn,
            tableStyle.FirstColumnValue,
            tableStyle.LastColumn,
            tableStyle.LastColumnValue,
            tableStyle.BandRow,
            tableStyle.BandRowValue,
            tableStyle.BandColumn,
            tableStyle.BandColumnValue,
            tableCells.Count(cell => cell.StyleFill.HasFill),
            tableCells.Count(cell => cell.StyleText.Color is not null),
            tableCells.Count(cell => cell.StyleText.Bold),
            node.Chart is not null,
            node.Chart?.RelationshipId ?? string.Empty,
            node.Chart?.TargetPartName ?? string.Empty,
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
            node.Chart?.ColorStyle.IsDefined ?? false,
            node.Chart?.ColorStyle.PartName ?? string.Empty,
            node.Chart?.ColorStyle.Method ?? string.Empty,
            node.Chart?.ColorStyle.Id ?? string.Empty,
            node.Chart?.ColorStyle.Colors.Count ?? 0,
            node.Chart?.ColorStyle.VariationCount ?? 0,
            chartColorDeclarations.Count,
            node.Chart?.ColorStyle.RootDeclarations.Count ?? 0,
            chartColorDeclarations.Count(declaration => declaration.IsResolved),
            chartColorDeclarationKinds,
            node.Chart?.StylePart.IsDefined ?? false,
            node.Chart?.StylePart.PartName ?? string.Empty,
            node.Chart?.StylePart.Id ?? string.Empty,
            chartStyleEntryRoles.Count,
            chartStyleEntryRoles,
            chartStyleEntrySourceIndexes,
            chartStyleEntries.Count(entry => HasChartShapeStyle(entry.ShapeStyle)),
            chartStyleEntries.Count(entry => entry.ShapeStyle?.Fill.HasFill == true ||
                entry.ShapeStyle?.PatternFill.HasPattern == true ||
                entry.ShapeStyle?.GradientFill?.HasGradient == true),
            chartStyleEntries.Count(entry => entry.FillReferenceIndex is not null),
            chartStyleEntries.Count(entry => entry.EffectReferenceIndex is not null),
            chartStyleEntries.Count(entry => !string.IsNullOrWhiteSpace(entry.FontReferenceIndex)),
            rejectedPointStyleIndexValues.Count,
            rejectedPointStyleIndexValues,
            rejectedDataLabelOverrideIndexValues.Count,
            rejectedDataLabelOverrideIndexValues,
            node.Chart?.ExternalData.IsDefined ?? false,
            node.Chart?.ExternalData.RelationshipId ?? string.Empty,
            node.Chart?.ExternalData.TargetPartName ?? string.Empty,
            node.Chart?.ExternalData.AutoUpdate,
            node.Chart?.ExternalData.AutoUpdateValue ?? string.Empty,
            node.Chart?.ExternalData.Resource is not null,
            node.Chart?.ExternalData.Resource?.ContentType ?? string.Empty,
            node.Chart?.Options.Date1904,
            node.Chart?.Options.Date1904Value ?? string.Empty,
            node.Chart?.Options.RoundedCorners,
            node.Chart?.Options.RoundedCornersValue ?? string.Empty,
            node.Chart?.Options.PlotVisibleOnly,
            node.Chart?.Options.PlotVisibleOnlyValue ?? string.Empty,
            node.Chart?.Options.ShowDataLabelsOverMaximum,
            node.Chart?.Options.ShowDataLabelsOverMaximumValue ?? string.Empty,
            node.Chart?.Options.DisplayBlanksAs ?? string.Empty,
            plotAreaLayout?.HasLayout ?? false,
            plotAreaLayout?.X,
            plotAreaLayout?.Y,
            plotAreaLayout?.Width,
            plotAreaLayout?.Height,
            plotAreaLayout?.LayoutTarget ?? string.Empty,
            plotAreaLayout?.LayoutTargetKind.ToString() ?? string.Empty,
            plotAreaLayout?.XMode ?? string.Empty,
            plotAreaLayout?.XModeKind.ToString() ?? string.Empty,
            plotAreaLayout?.YMode ?? string.Empty,
            plotAreaLayout?.YModeKind.ToString() ?? string.Empty,
            plotAreaLayout?.WidthMode ?? string.Empty,
            plotAreaLayout?.WidthModeKind.ToString() ?? string.Empty,
            plotAreaLayout?.HeightMode ?? string.Empty,
            plotAreaLayout?.HeightModeKind.ToString() ?? string.Empty,
            legend?.IsDefined ?? false,
            legend?.Position ?? string.Empty,
            legend?.Overlay,
            legend?.OverlayValue ?? string.Empty,
            legend?.IsDeleted,
            legend?.IsDeletedValue ?? string.Empty,
            legendLayout?.HasLayout ?? false,
            legendLayout?.X,
            legendLayout?.Y,
            legendLayout?.Width,
            legendLayout?.Height,
            node.Kind == PptxSceneNodeKind.Group,
            node.Children.Select(ToSnapshot).ToArray());
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

    private static IReadOnlyList<string> ReadRejectedPointStyleIndexValues(PptxSceneChart? chart)
    {
        if (chart is null)
        {
            return [];
        }

        return chart.Plots
            .SelectMany(plot => plot.Series)
            .SelectMany(series => series.RejectedPointStyleIndexValues)
            .ToArray();
    }

    private static IReadOnlyList<string> ReadRejectedDataLabelOverrideIndexValues(PptxSceneChart? chart)
    {
        if (chart is null)
        {
            return [];
        }

        return chart.Plots
            .SelectMany(plot => plot.DataLabels.RejectedOverrideIndexValues)
            .ToArray();
    }
}
