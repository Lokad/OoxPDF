using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static void EmitUnsupportedFeatureDiagnostics(PptxSceneSlide sceneSlide, XDocument slideXml, string partName, int slideIndex, Action<OoxPdfDiagnostic>? diagnosticSink)
    {
        if (diagnosticSink is null)
        {
            return;
        }

        var emitted = new HashSet<string>(StringComparer.Ordinal);
        void Emit(string id, string feature)
        {
            if (!emitted.Add(id))
            {
                return;
            }

            diagnosticSink(new OoxPdfDiagnostic(
                id,
                OoxPdfSeverity.Warning,
                $"Unsupported PPTX feature '{feature}' was detected and ignored.",
                partName,
                SlideIndex: slideIndex,
                Feature: feature,
                Fallback: "Ignored"));
        }

        if (sceneSlide.HasTransition ||
            slideXml.Descendants(PresentationNamespace + "transition").Any())
        {
            Emit("PPTX_UNSUPPORTED_TRANSITION", "transition");
        }

        if (sceneSlide.HasTiming ||
            slideXml.Descendants(PresentationNamespace + "timing").Any())
        {
            Emit("PPTX_UNSUPPORTED_ANIMATION", "animation");
        }

        if (HasVideo(sceneSlide.SlideNodes) ||
            slideXml.Descendants(PresentationNamespace + "video").Any() ||
            slideXml.Descendants(DrawingNamespace + "videoFile").Any())
        {
            Emit("PPTX_UNSUPPORTED_VIDEO", "video");
        }

        if (HasAudio(sceneSlide.SlideNodes) ||
            slideXml.Descendants(PresentationNamespace + "audio").Any() ||
            slideXml.Descendants(DrawingNamespace + "audioFile").Any())
        {
            Emit("PPTX_UNSUPPORTED_AUDIO", "audio");
        }

        if (sceneSlide.HasOleObject ||
            slideXml.Descendants(PresentationNamespace + "oleObj").Any())
        {
            Emit("PPTX_UNSUPPORTED_OLE_OBJECT", "OLE object");
        }

        if (HasSmartArtGraphicFrame(sceneSlide.SlideNodes) ||
            HasGraphicDataUri(slideXml, "drawingml/2006/diagram"))
        {
            Emit("PPTX_UNSUPPORTED_SMARTART", "SmartArt");
        }

        if (HasUnsupportedGraphicFrame(sceneSlide.SlideNodes) ||
            slideXml.Descendants(PresentationNamespace + "graphicFrame").Any(IsUnsupportedGraphicFrame))
        {
            Emit("PPTX_UNSUPPORTED_GRAPHIC_FRAME", "graphic frame");
        }

        if (HasUnsupportedGradientFill(sceneSlide, slideXml))
        {
            Emit("PPTX_UNSUPPORTED_GRADIENT_FILL", "gradient fill");
        }

        if (HasUnsupportedPatternFill(sceneSlide, slideXml))
        {
            Emit("PPTX_UNSUPPORTED_PATTERN_FILL", "pattern fill");
        }

        if (HasUnsupportedTextOrientation(sceneSlide, slideXml))
        {
            Emit("PPTX_UNSUPPORTED_TEXT_ORIENTATION", "vertical text");
        }

        if (HasUnsupportedTextVerticalOverflow(sceneSlide, slideXml))
        {
            Emit("PPTX_UNSUPPORTED_TEXT_OVERFLOW", "text vertical overflow");
        }

        if (HasUnsupportedPictureFill(sceneSlide))
        {
            Emit("PPTX_UNSUPPORTED_PICTURE_FILL", "picture fill");
        }

        if (HasTiledImageFill(sceneSlide))
        {
            Emit("PPTX_UNSUPPORTED_IMAGE_TILE", "tiled image fill");
        }

        if (HasUnsupportedTransparency(sceneSlide, slideXml))
        {
            Emit("PPTX_UNSUPPORTED_TRANSPARENCY", "transparency");
        }

        if (HasUnsupportedEffect(sceneSlide))
        {
            Emit("PPTX_UNSUPPORTED_EFFECT", "effect");
        }

        if (HasUnsupportedCustomGeometry(sceneSlide))
        {
            Emit("PPTX_UNSUPPORTED_CUSTOM_GEOMETRY", "custom geometry");
        }

        if (HasUnsupportedCallout(sceneSlide))
        {
            Emit("PPTX_UNSUPPORTED_CALLOUT", "callout shape");
        }
    }

    private static bool HasTiledImageFill(PptxSceneSlide sceneSlide)
    {
        return HasTiledImageFill(sceneSlide.SlideNodes) ||
            HasTiledImageFill(sceneSlide.LayoutNodes) ||
            HasTiledImageFill(sceneSlide.MasterNodes);
    }

    private static bool HasTiledImageFill(IReadOnlyList<PptxSceneNode> nodes)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if (node.Picture?.Tile.HasTile == true ||
                node.Shape?.PictureFill.Tile.HasTile == true ||
                HasTiledImageFill(node.Children))
            {
                return true;
            }
        }

        return false;
    }

    private static bool HasUnsupportedPictureFill(PptxSceneSlide sceneSlide)
    {
        return HasUnsupportedPictureFill(sceneSlide.SlideNodes) ||
            HasUnsupportedPictureFill(sceneSlide.LayoutNodes) ||
            HasUnsupportedPictureFill(sceneSlide.MasterNodes);
    }

    private static bool HasUnsupportedPictureFill(IReadOnlyList<PptxSceneNode> nodes)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if ((node.Shape is { PictureFill.HasPicture: true } shape &&
                    !CanRenderPictureFillPreset(shape.Preset)) ||
                HasUnsupportedPictureFill(node.Children))
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsUnsupportedCalloutPreset(string? preset)
    {
        return preset?.Contains("Callout", StringComparison.OrdinalIgnoreCase) == true &&
            !string.Equals(preset, "wedgeRectCallout", StringComparison.Ordinal);
    }

    private static bool HasUnsupportedCallout(PptxSceneSlide sceneSlide)
    {
        return HasUnsupportedCallout(sceneSlide.SlideNodes);
    }

    private static bool HasUnsupportedCallout(IReadOnlyList<PptxSceneNode> nodes)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if ((node.Shape is not null && IsUnsupportedCalloutPreset(node.Shape.Preset)) ||
                HasUnsupportedCallout(node.Children))
            {
                return true;
            }
        }

        return false;
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

    private static bool IsUnsupportedGradientFill(XElement gradientFill)
    {
        XName? ownerName = gradientFill.Parent?.Name;
        if (ownerName != PresentationNamespace + "spPr" && ownerName != PresentationNamespace + "bgPr")
        {
            return true;
        }

        if (gradientFill.Element(DrawingNamespace + "lin") is null)
        {
            return true;
        }

        XElement[] stops = gradientFill
            .Element(DrawingNamespace + "gsLst")
            ?.Elements(DrawingNamespace + "gs")
            .ToArray() ?? [];
        return stops.Length < 2 ||
            stops.Any(stop => stop.Elements().FirstOrDefault(IsGradientColorElement) is null) ||
            !HasSupportedGradientStopAlpha(stops);
    }

    private static bool HasUnsupportedGradientFill(PptxSceneSlide sceneSlide, XDocument slideXml)
    {
        return HasUnsupportedShapeGradientFill(sceneSlide.SlideNodes) ||
            slideXml.Descendants(DrawingNamespace + "gradFill").Any(IsUnsupportedNonShapeGradientFill);
    }

    private static bool HasUnsupportedShapeGradientFill(IReadOnlyList<PptxSceneNode> nodes)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if ((node.Shape is { GradientFill.HasUnsupportedGradient: true }) ||
                HasUnsupportedShapeGradientFill(node.Children))
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsUnsupportedNonShapeGradientFill(XElement gradientFill)
    {
        return gradientFill.Parent?.Name != PresentationNamespace + "spPr" &&
            IsUnsupportedGradientFill(gradientFill);
    }

    private static bool IsGradientColorElement(XElement color)
    {
        return color.Name.LocalName is "srgbClr" or "schemeClr" or "prstClr" or "sysClr" or "scrgbClr" or "hslClr";
    }

    private static bool HasSupportedGradientStopAlpha(IReadOnlyList<XElement> stops)
    {
        int first = ReadGradientStopAlpha(stops[0]);
        return stops.All(stop => Math.Abs(ReadGradientStopAlpha(stop) - first) <= 100);
    }

    private static int ReadGradientStopAlpha(XElement stop)
    {
        XElement? alpha = stop
            .Elements()
            .FirstOrDefault(IsGradientColorElement)
            ?.Element(DrawingNamespace + "alpha");
        return alpha?.Attribute("val") is { } value &&
            int.TryParse(value.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)
            ? Math.Clamp(parsed, 0, 100000)
            : 100000;
    }

    private static bool HasUnsupportedEffect(PptxSceneSlide sceneSlide)
    {
        return HasUnsupportedEffect(sceneSlide.SlideNodes);
    }

    private static bool HasUnsupportedEffect(IReadOnlyList<PptxSceneNode> nodes)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if ((node.Shape is not null && HasUnsupportedEffect(node.Shape.Effects)) ||
                (node.Chart is not null && HasUnsupportedEffect(node.Chart)) ||
                HasUnsupportedEffect(node.Children))
            {
                return true;
            }
        }

        return false;
    }

    private static bool HasUnsupportedEffect(PptxSceneShapeEffectFamily effects)
    {
        return effects.HasEffectDag ||
            effects.UnsupportedEffectNames?.Count > 0;
    }

    private static bool HasUnsupportedEffect(PptxSceneChart chart)
    {
        return HasUnsupportedEffect(chart.ChartAreaStyle) ||
            HasUnsupportedEffect(chart.PlotAreaStyle) ||
            HasUnsupportedEffect(chart.Title.ShapeStyle) ||
            HasUnsupportedEffect(chart.Legend.ShapeStyle) ||
            chart.StylePart.Entries.Any(entry => HasUnsupportedEffect(entry.ShapeStyle)) ||
            chart.Plots.Any(HasUnsupportedEffect);
    }

    private static bool HasUnsupportedEffect(PptxSceneChartPlot plot)
    {
        return HasUnsupportedEffect(plot.DataLabels) ||
            plot.Series.Any(HasUnsupportedEffect);
    }

    private static bool HasUnsupportedEffect(PptxSceneChartSeries series)
    {
        return HasUnsupportedEffect(series.Effects) ||
            HasUnsupportedEffect(series.DataLabels) ||
            series.PointStyles.Any(point => HasUnsupportedEffect(point.Effects));
    }

    private static bool HasUnsupportedEffect(PptxSceneChartDataLabels labels)
    {
        return HasUnsupportedEffect(labels.ShapeStyle) ||
            labels.Overrides.Any(label => HasUnsupportedEffect(label.ShapeStyle));
    }

    private static bool HasUnsupportedEffect(PptxSceneChartShapeStyle style)
    {
        return HasUnsupportedEffect(style.Effects);
    }

    private static bool HasUnsupportedEffect(PptxSceneChartEffectFamily effects)
    {
        return effects.HasEffectDag ||
            effects.UnsupportedEffectNames?.Count > 0;
    }

    private static bool HasUnsupportedCustomGeometry(PptxSceneSlide sceneSlide)
    {
        return HasUnsupportedCustomGeometry(sceneSlide.SlideNodes);
    }

    private static bool HasUnsupportedCustomGeometry(IReadOnlyList<PptxSceneNode> nodes)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if ((node.Shape is not null && node.Shape.CustomGeometry.HasUnsupportedGeometry) ||
                HasUnsupportedCustomGeometry(node.Children))
            {
                return true;
            }
        }

        return false;
    }

    private static bool HasUnsupportedTextOrientation(XElement bodyProperties)
    {
        string? orientation = (string?)bodyProperties.Attribute("vert");
        return !string.IsNullOrEmpty(orientation) &&
            !orientation.Equals("horz", StringComparison.OrdinalIgnoreCase) &&
            !orientation.Equals("vert", StringComparison.OrdinalIgnoreCase) &&
            !orientation.Equals("vert270", StringComparison.OrdinalIgnoreCase) &&
            !orientation.Equals("eaVert", StringComparison.OrdinalIgnoreCase) &&
            !orientation.Equals("mongolianVert", StringComparison.OrdinalIgnoreCase) &&
            !orientation.Equals("wordArtVert", StringComparison.OrdinalIgnoreCase) &&
            !orientation.Equals("wordArtVertRtl", StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasUnsupportedTextOrientation(PptxSceneSlide sceneSlide, XDocument slideXml)
    {
        return HasUnsupportedSceneTextOrientation(sceneSlide.SlideNodes) ||
            slideXml.Descendants(DrawingNamespace + "bodyPr").Any(IsUnsupportedNonShapeTextOrientation);
    }

    private static bool HasUnsupportedSceneTextOrientation(IReadOnlyList<PptxSceneNode> nodes)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if ((node.TextBody?.HasUnsupportedTextOrientation == true) ||
                HasUnsupportedTableTextOrientation(node.Table) ||
                HasUnsupportedSceneTextOrientation(node.Children))
            {
                return true;
            }
        }

        return false;
    }

    private static bool HasUnsupportedTableTextOrientation(PptxSceneTable? table)
    {
        return table?.Rows.Any(row => row.Cells.Any(cell => cell.HasUnsupportedTextOrientation)) == true;
    }

    private static bool IsUnsupportedNonShapeTextOrientation(XElement bodyProperties)
    {
        return !IsSceneTextBodyProperties(bodyProperties) &&
            HasUnsupportedTextOrientation(bodyProperties);
    }

    private static bool HasUnsupportedTextVerticalOverflow(XElement bodyProperties)
    {
        string? overflow = (string?)bodyProperties.Attribute("vertOverflow");
        return overflow?.Equals("ellipsis", StringComparison.OrdinalIgnoreCase) == true;
    }

    private static bool HasUnsupportedTextVerticalOverflow(PptxSceneSlide sceneSlide, XDocument slideXml)
    {
        return HasUnsupportedSceneTextVerticalOverflow(sceneSlide.SlideNodes) ||
            slideXml.Descendants(DrawingNamespace + "bodyPr").Any(IsUnsupportedNonShapeTextVerticalOverflow);
    }

    private static bool HasUnsupportedSceneTextVerticalOverflow(IReadOnlyList<PptxSceneNode> nodes)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if ((node.TextBody?.HasUnsupportedVerticalOverflow == true) ||
                HasUnsupportedTableTextVerticalOverflow(node.Table) ||
                HasUnsupportedSceneTextVerticalOverflow(node.Children))
            {
                return true;
            }
        }

        return false;
    }

    private static bool HasUnsupportedTableTextVerticalOverflow(PptxSceneTable? table)
    {
        return table?.Rows.Any(row => row.Cells.Any(cell => cell.HasUnsupportedVerticalOverflow)) == true;
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

    private static bool IsUnsupportedPatternFill(XElement patternFill)
    {
        return !IsSupportedDiagonalPatternFill((string?)patternFill.Attribute("prst"));
    }

    private static bool HasUnsupportedPatternFill(PptxSceneSlide sceneSlide, XDocument slideXml)
    {
        return HasUnsupportedShapePatternFill(sceneSlide.SlideNodes) ||
            slideXml.Descendants(DrawingNamespace + "pattFill").Any(IsUnsupportedNonShapePatternFill);
    }

    private static bool HasUnsupportedShapePatternFill(IReadOnlyList<PptxSceneNode> nodes)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if ((node.Shape is { PatternFill.HasUnsupportedPattern: true }) ||
                HasUnsupportedShapePatternFill(node.Children))
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsUnsupportedNonShapePatternFill(XElement patternFill)
    {
        return patternFill.Parent?.Name != PresentationNamespace + "spPr" &&
            IsUnsupportedPatternFill(patternFill);
    }

    private static bool HasGraphicDataUri(XDocument slideXml, string marker)
    {
        return slideXml
            .Descendants(DrawingNamespace + "graphicData")
            .Select(element => (string?)element.Attribute("uri"))
            .Any(uri => uri?.Contains(marker, StringComparison.OrdinalIgnoreCase) == true);
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

    private static bool HasUnsupportedTransparency(PptxSceneSlide sceneSlide, XDocument slideXml)
    {
        return HasUnsupportedShapeTransparency(sceneSlide.SlideNodes) ||
            slideXml.Descendants(DrawingNamespace + "alpha").Any(IsUnsupportedNonShapeAlpha);
    }

    private static bool HasUnsupportedShapeTransparency(IReadOnlyList<PptxSceneNode> nodes)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if ((node.Shape is { HasUnsupportedTransparency: true }) ||
                HasUnsupportedShapeTransparency(node.Children))
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsUnsupportedNonShapeAlpha(XElement alpha)
    {
        return !IsShapeAlpha(alpha) && PptxSceneBuilder.IsUnsupportedAlpha(alpha);
    }

    private static bool IsShapeAlpha(XElement alpha)
    {
        return alpha.Ancestors().Any(ancestor =>
            ancestor.Name == PresentationNamespace + "sp" ||
            ancestor.Name == PresentationNamespace + "cxnSp");
    }
}
