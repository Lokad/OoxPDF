using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static void EmitUnsupportedFeatureDiagnostics(PptxSceneSlide sceneSlide, XDocument slideXml, string partName, int slideIndex, Action<OoxPdfDiagnostic>? diagnosticSink, HashSet<string> warnedParts,
        IReadOnlyDictionary<PptxTextSpanMemoKey, IReadOnlyList<PptxPositionedTextSpan>>? textSpanMemo)
    {
        if (diagnosticSink is null)
        {
            return;
        }

        var emitted = new HashSet<string>(StringComparer.Ordinal);
        void Emit(string id, string feature, string? sourcePartName = null)
        {
            if (!emitted.Add(id))
            {
                return;
            }

            diagnosticSink(new OoxPdfDiagnostic(
                id,
                OoxPdfSeverity.Warning,
                $"Unsupported PPTX feature '{feature}' was detected and ignored.",
                sourcePartName ?? partName,
                PageIndex: null,
                SlideIndex: slideIndex,
                Feature: feature,
                Fallback: "Ignored"));
        }

        void EmitTextOverflow(string? sourcePartName = null)
        {
            const string id = "PPTX_UNSUPPORTED_TEXT_OVERFLOW";
            if (!emitted.Add(id))
            {
                return;
            }

            diagnosticSink(new OoxPdfDiagnostic(
                id,
                OoxPdfSeverity.Warning,
                "PPTX text vertical overflow uses ellipsis; local clipping was applied, but the ellipsis marker is not rendered.",
                sourcePartName ?? partName,
                PageIndex: null,
                SlideIndex: slideIndex,
                Feature: "text vertical overflow ellipsis marker",
                Fallback: "Local clip"));
        }

        if (sceneSlide.HasTransition)
        {
            Emit("PPTX_UNSUPPORTED_TRANSITION", "transition");
        }

        OoxMarkupCompatibility.WarnMustUnderstandOnce(slideXml, partName, diagnosticSink, warnedParts, slideIndex);

        if (sceneSlide.MasterXml is not null)
        {
            OoxMarkupCompatibility.WarnMustUnderstandOnce(sceneSlide.MasterXml, sceneSlide.MasterPartName ?? partName, diagnosticSink, warnedParts, slideIndex);
        }

        if (sceneSlide.LayoutXml is not null)
        {
            OoxMarkupCompatibility.WarnMustUnderstandOnce(sceneSlide.LayoutXml, sceneSlide.LayoutPartName ?? partName, diagnosticSink, warnedParts, slideIndex);
        }

        if (sceneSlide.HasTiming)
        {
            Emit("PPTX_UNSUPPORTED_ANIMATION", "animation");
        }

        if (HasVideo(sceneSlide.SlideNodes) ||
            VisibleSlideElements(PresentationNamespace + "video").Any() ||
            VisibleSlideElements(DrawingNamespace + "videoFile").Any())
        {
            Emit("PPTX_UNSUPPORTED_VIDEO", "video");
        }

        if (HasAudio(sceneSlide.SlideNodes) ||
            VisibleSlideElements(PresentationNamespace + "audio").Any() ||
            VisibleSlideElements(DrawingNamespace + "audioFile").Any())
        {
            Emit("PPTX_UNSUPPORTED_AUDIO", "audio");
        }

        if (sceneSlide.HasOleObject)
        {
            Emit("PPTX_UNSUPPORTED_OLE_OBJECT", "OLE object");
        }

        if (HasSmartArtGraphicFrame(sceneSlide.SlideNodes))
        {
            Emit("PPTX_UNSUPPORTED_SMARTART", "SmartArt");
        }

        if (HasUnsupportedGraphicFrame(sceneSlide.SlideNodes) ||
            VisibleSlideElements(PresentationNamespace + "graphicFrame").Any(IsUnsupportedGraphicFrame))
        {
            Emit("PPTX_UNSUPPORTED_GRAPHIC_FRAME", "graphic frame");
        }

        if (sceneSlide.ShowMasterShapes)
        {
            EmitNodeDiagnostics(sceneSlide.MasterNodes, sceneSlide.MasterColorMap, false, sceneSlide.MasterPartName ?? partName);
        }
        EmitNodeDiagnostics(sceneSlide.LayoutNodes, sceneSlide.LayoutColorMap, false, sceneSlide.LayoutPartName ?? partName);
        EmitNodeDiagnostics(sceneSlide.SlideNodes, sceneSlide.SlideColorMap, true, partName);

        // Non-scene fallbacks retain background and orphan-content diagnostics,
        // while drawing-owned content shares scene visibility, including groups.
        // Part-level compatibility, timing and transition checks stay independent.
        if (sceneSlide.SlideBackground.HasUnsupportedGradient ||
            VisibleSlideElements(DrawingNamespace + "gradFill").Any(IsUnsupportedNonShapeGradientFill))
            Emit("PPTX_UNSUPPORTED_GRADIENT_FILL", "gradient fill");
        else if (!sceneSlide.SlideBackground.HasFill)
        {
            if (sceneSlide.LayoutBackground.HasUnsupportedGradient)
                Emit("PPTX_UNSUPPORTED_GRADIENT_FILL", "gradient fill", sceneSlide.LayoutPartName);
            else if (!sceneSlide.LayoutBackground.HasFill && sceneSlide.MasterBackground.HasUnsupportedGradient)
                Emit("PPTX_UNSUPPORTED_GRADIENT_FILL", "gradient fill", sceneSlide.MasterPartName);
        }
        if (VisibleSlideElements(DrawingNamespace + "pattFill").Any(IsUnsupportedNonShapePatternFill))
            Emit("PPTX_UNSUPPORTED_PATTERN_FILL", "pattern fill");
        if (VisibleSlideElements(DrawingNamespace + "bodyPr").Any(IsUnsupportedNonShapeTextOrientation))
            Emit("PPTX_UNSUPPORTED_TEXT_ORIENTATION", "vertical text");
        if (VisibleSlideElements(DrawingNamespace + "bodyPr").Any(IsUnsupportedNonShapeTextVerticalOverflow))
            EmitTextOverflow();
        if (PptxSceneBuilder.HasUnsupportedAlpha(slideXml, IsVisibleNonShapeColor))
            Emit("PPTX_UNSUPPORTED_TRANSPARENCY", "transparency");

        IEnumerable<XElement> VisibleSlideElements(XName name) =>
            slideXml.Descendants(name).Where(PptxSceneBuilder.IsVisibleDrawingContent);

        void EmitNodeDiagnostics(IReadOnlyList<PptxSceneNode> nodes, PptxColorMap colorMap, bool renderPlaceholders, string nodePartName)
        {
            bool gradient = false, pattern = false, orientation = false, overflow = false;
            bool pictureFill = false, tile = false, transparency = false, effect = false;
            bool geometry = false, callout = false;
            foreach (PptxSceneNode node in EnumerateRenderedSceneNodes(nodes, renderPlaceholders))
            {
                var shape = node.Shape;
                gradient = gradient || shape is { GradientFill.HasUnsupportedGradient: true } || HasUnsupportedChartGradientFill(node.Chart);
                pattern = pattern || shape is { PatternFill.HasUnsupportedPattern: true };
                orientation = orientation || HasUnsupportedSceneTextOrientation(node);
                overflow = overflow || HasUnsupportedSceneTextVerticalOverflow(node, colorMap, renderPlaceholders, textSpanMemo);
                pictureFill = pictureFill || HasUnsupportedPictureFill(node);
                tile = tile || node.Picture?.HasTile == true || shape?.PictureFill.HasTile == true;
                transparency = transparency || shape is { HasUnsupportedTransparency: true };
                effect = effect || HasUnsupportedEffect(node);
                geometry = geometry || shape?.CustomGeometry.HasUnsupportedGeometry == true;
                callout = callout || IsUnsupportedCalloutPreset(shape?.Preset);
            }

            // Emit in feature order within each part. The first visible part
            // that owns a feature supplies its diagnostic on this slide.
            if (gradient)
                Emit("PPTX_UNSUPPORTED_GRADIENT_FILL", "gradient fill", nodePartName);
            if (pattern)
                Emit("PPTX_UNSUPPORTED_PATTERN_FILL", "pattern fill", nodePartName);
            if (orientation)
                Emit("PPTX_UNSUPPORTED_TEXT_ORIENTATION", "vertical text", nodePartName);
            if (overflow)
                EmitTextOverflow(nodePartName);
            if (pictureFill)
                Emit("PPTX_UNSUPPORTED_PICTURE_FILL", "picture fill", nodePartName);
            if (tile)
                Emit("PPTX_UNSUPPORTED_IMAGE_TILE", "tiled image fill", nodePartName);
            if (transparency)
                Emit("PPTX_UNSUPPORTED_TRANSPARENCY", "transparency", nodePartName);
            if (effect)
                Emit("PPTX_UNSUPPORTED_EFFECT", "effect", nodePartName);
            if (geometry)
                Emit("PPTX_UNSUPPORTED_CUSTOM_GEOMETRY", "custom geometry", nodePartName);
            if (callout)
                Emit("PPTX_UNSUPPORTED_CALLOUT", "callout shape", nodePartName);
        }
    }

    private static bool HasUnsupportedPictureFill(PptxSceneNode node)
    {
        return (node.Shape is { PictureFill.HasPicture: true } shape &&
            (shape.CustomGeometry.HasGeometry
                ? !TryResolveCustomPictureFillPath(shape.CustomGeometry, out _)
                : !CanRenderPictureFillPreset(shape.Preset))) ||
            HasUnsupportedChartPictureFill(node.Chart);
    }

    private static bool HasUnsupportedChartPictureFill(PptxSceneChart? chart)
    {
        return chart is not null &&
            (HasUnsupportedChartPictureFill(chart.ChartAreaStyle) ||
                HasUnsupportedChartPictureFill(chart.PlotAreaStyle) ||
                HasUnsupportedChartPictureFill(chart.Title.ShapeStyle) ||
                HasUnsupportedChartPictureFill(chart.Legend.ShapeStyle) ||
                chart.StylePart.Entries.Any(entry => HasUnsupportedChartPictureFill(entry.ShapeStyle)) ||
                chart.Plots.Any(HasUnsupportedChartPictureFill));
    }

    private static bool HasUnsupportedChartPictureFill(PptxSceneChartPlot plot)
    {
        return HasUnsupportedChartPictureFill(plot.DataLabels) ||
            plot.Series.Any(HasUnsupportedChartPictureFill);
    }

    private static bool HasUnsupportedChartPictureFill(PptxSceneChartSeries series)
    {
        return HasUnsupportedChartPictureFill(series.DataLabels);
    }

    private static bool HasUnsupportedChartPictureFill(PptxSceneChartDataLabels labels)
    {
        return HasUnsupportedChartPictureFill(labels.ShapeStyle) ||
            labels.Overrides.Any(label => HasUnsupportedChartPictureFill(label.ShapeStyle));
    }

    private static bool HasUnsupportedChartPictureFill(PptxSceneChartShapeStyle style)
    {
        return style.PictureFill.HasPicture;
    }

    private static bool IsUnsupportedCalloutPreset(string? preset)
    {
        return preset?.Contains("Callout", StringComparison.OrdinalIgnoreCase) == true &&
            !string.Equals(preset, "wedgeRectCallout", StringComparison.Ordinal);
    }

    private static bool HasVideo(IReadOnlyList<PptxSceneNode> nodes)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if (node.Picture?.HasVideo == true ||
                HasVideo(node.Children))
            {
                return true;
            }
        }

        return false;
    }

    private static bool HasAudio(IReadOnlyList<PptxSceneNode> nodes)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if (node.Picture?.HasAudio == true ||
                HasAudio(node.Children))
            {
                return true;
            }
        }

        return false;
    }

    private static bool HasUnsupportedChartGradientFill(PptxSceneChart? chart)
    {
        return chart is not null &&
            (HasUnsupportedChartGradientFill(chart.ChartAreaStyle) ||
                HasUnsupportedChartGradientFill(chart.PlotAreaStyle) ||
                HasUnsupportedChartGradientFill(chart.Title.ShapeStyle) ||
                HasUnsupportedChartGradientFill(chart.Legend.ShapeStyle) ||
                chart.StylePart.Entries.Any(entry => HasUnsupportedChartGradientFill(entry.ShapeStyle)) ||
                chart.Plots.Any(HasUnsupportedChartGradientFill));
    }

    private static bool HasUnsupportedChartGradientFill(PptxSceneChartPlot plot)
    {
        return HasUnsupportedChartGradientFill(plot.DataLabels) ||
            plot.Series.Any(HasUnsupportedChartGradientFill);
    }

    private static bool HasUnsupportedChartGradientFill(PptxSceneChartSeries series)
    {
        return HasUnsupportedChartGradientFill(series.DataLabels);
    }

    private static bool HasUnsupportedChartGradientFill(PptxSceneChartDataLabels labels)
    {
        return HasUnsupportedChartGradientFill(labels.ShapeStyle) ||
            labels.Overrides.Any(label => HasUnsupportedChartGradientFill(label.ShapeStyle));
    }

    private static bool HasUnsupportedChartGradientFill(PptxSceneChartShapeStyle style)
    {
        return style.GradientFill?.HasUnsupportedGradient == true;
    }

    private static bool IsUnsupportedNonShapeGradientFill(XElement gradientFill)
    {
        return gradientFill.Parent?.Name != PresentationNamespace + "spPr" &&
            !PptxSceneBuilder.HasSupportedAuthoredGradientFill(gradientFill);
    }

    private static bool HasUnsupportedEffect(PptxSceneNode node)
    {
        return (node.Shape is { } shape && (shape.Effects.HasUnsupportedEffects ||
                (shape.Glow.HasGlow && !CanRenderGlowPreset(shape.Preset)))) ||
            (node.Chart is { } chart && HasUnsupportedEffect(chart));
    }

    private static bool HasUnsupportedEffect(PptxSceneChart chart)
    {
        return chart.ChartAreaStyle.Effects.HasUnsupportedEffects ||
            chart.PlotAreaStyle.Effects.HasUnsupportedEffects ||
            chart.Title.ShapeStyle.Effects.HasUnsupportedEffects ||
            chart.Legend.ShapeStyle.Effects.HasUnsupportedEffects ||
            chart.StylePart.Entries.Any(entry =>
                entry.ShapeStyle.Effects.HasUnsupportedEffects ||
                entry.EffectReferenceEffects.HasUnsupportedEffects) ||
            chart.Plots.Any(HasUnsupportedEffect);
    }

    private static bool HasUnsupportedEffect(PptxSceneChartPlot plot)
    {
        return HasUnsupportedEffect(plot.DataLabels) ||
            plot.Series.Any(HasUnsupportedEffect);
    }

    private static bool HasUnsupportedEffect(PptxSceneChartSeries series)
    {
        return series.Effects.HasUnsupportedEffects ||
            HasUnsupportedEffect(series.DataLabels) ||
            series.PointStyles.Any(point => point.Effects.HasUnsupportedEffects);
    }

    private static bool HasUnsupportedEffect(PptxSceneChartDataLabels labels)
    {
        return labels.ShapeStyle.Effects.HasUnsupportedEffects ||
            labels.Overrides.Any(label => label.ShapeStyle.Effects.HasUnsupportedEffects);
    }

    private static bool HasUnsupportedTextOrientation(XElement bodyProperties)
    {
        string? orientation = (string?)bodyProperties.Attribute("vert");
        return HasUnsupportedTextOrientation(orientation);
    }

    internal static bool HasUnsupportedTextOrientation(string? orientation) =>
        ParseTextOrientation(orientation) == PptxTextOrientation.Unknown;

    private static bool HasUnsupportedSceneTextOrientation(PptxSceneNode node)
    {
        return node.TextBody?.HasUnsupportedTextOrientation == true ||
            node.Table?.Rows.Any(row => row.Cells.Any(cell => cell.HasUnsupportedTextOrientation)) == true ||
            HasUnsupportedChartTextOrientation(node.Chart);

        static bool HasUnsupportedChartTextOrientation(PptxSceneChart? chart)
        {
            return chart is not null &&
                (HasUnsupportedTextOrientation(chart.Title.TextBodyProperties.OrientationValue) ||
                    HasUnsupportedTextOrientation(chart.Legend.TextBodyProperties.OrientationValue) ||
                    chart.Axes.Any(axis => HasUnsupportedTextOrientation(axis.Title.TextBodyProperties.OrientationValue)) ||
                    chart.Plots.Any(plot =>
                        HasUnsupportedTextOrientation(plot.DataLabels.TextBodyProperties.OrientationValue) ||
                        plot.DataLabels.Overrides.Any(label => HasUnsupportedTextOrientation(label.TextBodyProperties.OrientationValue)) ||
                        plot.Series.Any(series =>
                            HasUnsupportedTextOrientation(series.DataLabels.TextBodyProperties.OrientationValue) ||
                            series.DataLabels.Overrides.Any(label => HasUnsupportedTextOrientation(label.TextBodyProperties.OrientationValue)))));
        }
    }

    private static bool IsUnsupportedNonShapeTextOrientation(XElement bodyProperties)
    {
        return !IsSceneTextBodyProperties(bodyProperties) &&
            HasUnsupportedTextOrientation(bodyProperties);
    }

    private static bool HasUnsupportedTextVerticalOverflow(XElement bodyProperties)
    {
        string? overflow = (string?)bodyProperties.Attribute("vertOverflow");
        return HasUnsupportedTextVerticalOverflow(overflow);
    }

    private static bool HasUnsupportedTextVerticalOverflow(string? overflow)
    {
        return overflow?.Equals("ellipsis", StringComparison.OrdinalIgnoreCase) == true;
    }

    private static bool HasUnsupportedSceneTextVerticalOverflow(PptxSceneNode node, PptxColorMap colorMap, bool renderPlaceholders,
        IReadOnlyDictionary<PptxTextSpanMemoKey, IReadOnlyList<PptxPositionedTextSpan>>? textSpanMemo)
    {
        return HasUnmarkedShapeTextOverflow() || HasUnsupportedTableTextVerticalOverflow(node.Table) ||
            HasUnsupportedChartTextVerticalOverflow(node.Chart);

        bool HasUnmarkedShapeTextOverflow()
        {
            return node.TextBody is not null && textSpanMemo is not null &&
                textSpanMemo.TryGetValue(new PptxTextSpanMemoKey(node, colorMap, renderPlaceholders), out var spans) &&
                spans.Count > 0 && spans[0].FrameVerticalOverflowMode == PptxTextVerticalOverflow.Ellipsis &&
                spans.Any(UsesEllipsisClipFallback);
        }

        static bool HasUnsupportedChartTextVerticalOverflow(PptxSceneChart? chart)
        {
            return chart is not null &&
                (HasUnsupportedTextVerticalOverflow(chart.Title.TextBodyProperties.VerticalOverflowValue) ||
                    chart.Axes.Any(axis => HasUnsupportedTextVerticalOverflow(axis.Title.TextBodyProperties.VerticalOverflowValue)));
        }

        static bool HasUnsupportedTableTextVerticalOverflow(PptxSceneTable? table)
        {
            return table?.Rows.Any(row => row.Cells.Any(cell =>
                cell.HasUnsupportedVerticalOverflow &&
                TextBodyOrientationNeedsOverflowDiagnostic(cell.TextBody?.Element(DrawingNamespace + "bodyPr")))) == true;
        }
    }

    private static bool UsesEllipsisClipFallback(PptxPositionedTextSpan span)
    {
        TextRun run = span.Run;
        if (!HasTextTransform(run) || string.IsNullOrWhiteSpace(run.Text) ||
            (run.Alpha <= 0d && run.Outline is not { Alpha: > 0d }))
        {
            return false;
        }

        // Transformed text bypasses the ordinary ellipsis baseline filter. Its
        // baseline and the retained frame clip still share writing coordinates.
        double baseline = run.Y + run.BaselineOffset;
        return baseline < span.FrameClipY - PptxTextMetricRules.TextStateTolerance ||
            baseline > span.FrameClipY + span.FrameClipHeight + PptxTextMetricRules.TextStateTolerance;
    }

    private static bool TextBodyOrientationNeedsOverflowDiagnostic(XElement? bodyProperties)
    {
        string? orientation = (string?)bodyProperties?.Attribute("vert");
        return !string.IsNullOrEmpty(orientation) &&
            !orientation.Equals("horz", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsUnsupportedNonShapeTextVerticalOverflow(XElement bodyProperties)
    {
        return !IsSceneTextBodyProperties(bodyProperties) &&
            HasUnsupportedTextVerticalOverflow(bodyProperties);
    }

    private static bool IsSceneTextBodyProperties(XElement bodyProperties)
    {
        XName? parentName = bodyProperties.Parent?.Name;
        return parentName == PresentationNamespace + "txBody" ||
            parentName == DrawingNamespace + "txBody";
    }

    private static bool IsUnsupportedNonShapePatternFill(XElement patternFill)
    {
        return patternFill.Parent?.Name != PresentationNamespace + "spPr" &&
            !PptxSceneBuilder.IsSupportedDiagonalPatternFill((string?)patternFill.Attribute("prst"));
    }

    private static bool IsUnsupportedGraphicFrame(XElement graphicFrame)
    {
        XElement? graphicData = graphicFrame
            .Descendants(DrawingNamespace + "graphicData")
            .FirstOrDefault();
        if (graphicData is null || PptxSceneBuilder.IsSmartArtGraphicFrame(graphicFrame))
        {
            return false;
        }

        string uri = (string?)graphicData.Attribute("uri") ?? string.Empty;
        return !uri.Contains("chart", StringComparison.OrdinalIgnoreCase) &&
            !graphicData.Descendants(DrawingNamespace + "tbl").Any();
    }

    private static bool HasSmartArtGraphicFrame(IReadOnlyList<PptxSceneNode> nodes)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if (node.IsSmartArtGraphicFrame || HasSmartArtGraphicFrame(node.Children))
            {
                return true;
            }
        }

        return false;
    }

    private static bool HasUnsupportedGraphicFrame(IReadOnlyList<PptxSceneNode> nodes)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if ((node.Kind == PptxSceneNodeKind.UnknownGraphicFrame && !node.IsSmartArtGraphicFrame) ||
                HasUnsupportedGraphicFrame(node.Children))
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsVisibleNonShapeColor(XElement color) =>
        PptxSceneBuilder.IsVisibleDrawingContent(color) &&
        !color.Ancestors().Any(ancestor =>
            ancestor.Name == PresentationNamespace + "sp" ||
            ancestor.Name == PresentationNamespace + "cxnSp");
}
