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

        if (slideXml.Descendants(PresentationNamespace + "transition").Any())
        {
            Emit("PPTX_UNSUPPORTED_TRANSITION", "transition");
        }

        if (slideXml.Descendants(PresentationNamespace + "timing").Any())
        {
            Emit("PPTX_UNSUPPORTED_ANIMATION", "animation");
        }

        if (slideXml.Descendants(PresentationNamespace + "video").Any() ||
            slideXml.Descendants(DrawingNamespace + "videoFile").Any())
        {
            Emit("PPTX_UNSUPPORTED_VIDEO", "video");
        }

        if (slideXml.Descendants(PresentationNamespace + "audio").Any() ||
            slideXml.Descendants(DrawingNamespace + "audioFile").Any())
        {
            Emit("PPTX_UNSUPPORTED_AUDIO", "audio");
        }

        if (slideXml.Descendants(PresentationNamespace + "oleObj").Any())
        {
            Emit("PPTX_UNSUPPORTED_OLE_OBJECT", "OLE object");
        }

        if (HasGraphicDataUri(slideXml, "drawingml/2006/diagram"))
        {
            Emit("PPTX_UNSUPPORTED_SMARTART", "SmartArt");
        }

        if (slideXml.Descendants(PresentationNamespace + "graphicFrame").Any(IsUnsupportedGraphicFrame))
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

        if (slideXml.Descendants(DrawingNamespace + "bodyPr").Any(HasUnsupportedTextOrientation))
        {
            Emit("PPTX_UNSUPPORTED_TEXT_ORIENTATION", "vertical text");
        }

        if (slideXml.Descendants(DrawingNamespace + "bodyPr").Any(HasUnsupportedTextVerticalOverflow))
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

        if (slideXml.Descendants(DrawingNamespace + "alpha").Any(IsUnsupportedAlpha))
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

    private static bool HasUnsupportedTextVerticalOverflow(XElement bodyProperties)
    {
        string? overflow = (string?)bodyProperties.Attribute("vertOverflow");
        return overflow?.Equals("ellipsis", StringComparison.OrdinalIgnoreCase) == true;
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

    private static bool IsUnsupportedAlpha(XElement alpha)
    {
        if (alpha.Attribute("val") is not { } value ||
            !int.TryParse(value.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) ||
            parsed >= 100000)
        {
            return false;
        }

        XElement? color = alpha.Parent;
        XElement? fill = color?.Parent;
        XElement? owner = fill?.Parent;
        XElement? lineOwner = owner?.Parent;
        XElement? gradientStopList = fill?.Parent;
        XElement? gradientFill = gradientStopList?.Parent;
        XElement? gradientOwner = gradientFill?.Parent;
        bool supportedUniformGradientFill = fill?.Name == DrawingNamespace + "gs" &&
            gradientStopList?.Name == DrawingNamespace + "gsLst" &&
            gradientFill?.Name == DrawingNamespace + "gradFill" &&
            gradientOwner?.Name is { } ownerName &&
            (ownerName == PresentationNamespace + "spPr" || ownerName == PresentationNamespace + "bgPr") &&
            !IsUnsupportedGradientFill(gradientFill);
        bool supportedShapeFill = fill?.Name == DrawingNamespace + "solidFill" &&
            owner?.Name == PresentationNamespace + "spPr";
        bool supportedBackgroundFill = fill?.Name == DrawingNamespace + "solidFill" &&
            owner?.Name == PresentationNamespace + "bgPr";
        bool supportedShapeLine = fill?.Name == DrawingNamespace + "solidFill" &&
            owner?.Name == DrawingNamespace + "ln" &&
            lineOwner?.Name == PresentationNamespace + "spPr";
        bool supportedTextFill = fill?.Name == DrawingNamespace + "solidFill" &&
            owner?.Name == DrawingNamespace + "rPr";
        bool supportedTableCellFill = fill?.Name == DrawingNamespace + "solidFill" &&
            owner?.Name == DrawingNamespace + "tcPr";
        bool supportedTableBorder = fill?.Name == DrawingNamespace + "solidFill" &&
            owner is not null &&
            owner.Name.Namespace == DrawingNamespace &&
            owner.Name.LocalName is "lnL" or "lnR" or "lnT" or "lnB" &&
            lineOwner?.Name == DrawingNamespace + "tcPr";
        bool supportedOuterShadow = fill?.Name == DrawingNamespace + "outerShdw" &&
            owner?.Name == DrawingNamespace + "effectLst";
        bool supportedGlow = fill?.Name == DrawingNamespace + "glow" &&
            owner?.Name == DrawingNamespace + "effectLst";
        return !supportedUniformGradientFill && !supportedShapeFill && !supportedBackgroundFill && !supportedShapeLine && !supportedTextFill && !supportedTableCellFill && !supportedTableBorder && !supportedOuterShadow && !supportedGlow;
    }
}
