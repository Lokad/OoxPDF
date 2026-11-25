using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed record PptxScene(IReadOnlyList<PptxSceneSlide> Slides);

internal sealed partial class PptxSceneBuilder
{
    private const double MinimumStrokeWidth = 0.1d;
    private const double SceneEffectTolerance = 0.001d;
    private const string HyperlinkRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    private const string SlideLayoutRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
    private const string SlideMasterRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
    private const string ThemeRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
    private const string ChartExternalDataPackageRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
    private const string ChartColorStyleRelationshipType = "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";
    private const string ChartStyleRelationshipType = "http://schemas.microsoft.com/office/2011/relationships/chartStyle";

    public PptxScene Build(PptxDocument document, OoxPackage package, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        // Themes belong to masters, not to the presentation as a whole. Share
        // parsed models only by their actual part identity within this scene.
        var themes = new Dictionary<string, PptxTheme>(StringComparer.OrdinalIgnoreCase);
        PptxTheme? fallbackTheme = null;
        var slides = new List<PptxSceneSlide>(document.Slides.Count);
        IMemoryReservation? modelStorage = OoxMemory.Reserve(OoxConversionBudget.Current?.Models, 0);
        PptxScene? result = null;
        try
        {
            foreach (PptxSlide slide in document.Slides)
            {
                cancellationToken.ThrowIfCancellationRequested();
                OoxPart? slidePart = package.GetPart(slide.PartName);
                if (slidePart is null)
                {
                    slides.Add(new PptxSceneSlide(
                        slide.Index,
                        slide.PartName,
                        GetFallbackTheme(),
                        null,
                        null,
                        null,
                        null,
                        new XDocument(),
                        new Dictionary<string, OoxRelationship>(),
                        new Dictionary<string, OoxRelationship>(),
                        new Dictionary<string, OoxRelationship>(),
                        PptxColorMap.Default,
                        PptxColorMap.Default,
                        PptxColorMap.Default,
                        default,
                        default,
                        default,
                        false,
                        false,
                        false,
                        true,
                        [],
                        [],
                        []));
                    continue;
                }

                XDocument slideXml = package.LoadXml(slidePart, cancellationToken);
                OoxPart? layoutPart = GetRelatedPart(package, slide.PartName, SlideLayoutRelationshipType, cancellationToken);
                OoxPart? masterPart = layoutPart is null ? null : GetRelatedPart(package, layoutPart.Name, SlideMasterRelationshipType, cancellationToken);
                if (modelStorage is not null)
                {
                    // Master/layout nodes are copied into each slide's scene. Admit
                    // that fan-out before constructing the next complete slide model.
                    long sourceBytes = slidePart.Bytes.LongLength + (layoutPart?.Bytes.LongLength ?? 0) + (masterPart?.Bytes.LongLength ?? 0);
                    OoxMemory.Resize(modelStorage, checked(modelStorage.Bytes + 2 * (64 * 1024L + 64 * sourceBytes)));
                    cancellationToken.ThrowIfCancellationRequested();
                }
                PptxTheme theme = ReadTheme(masterPart?.Name) ?? GetFallbackTheme();
                XDocument? masterXml = masterPart is null ? null : package.LoadXml(masterPart, cancellationToken);
                XDocument? layoutXml = layoutPart is null ? null : package.LoadXml(layoutPart, cancellationToken);
                IReadOnlyDictionary<string, OoxRelationship> masterRelationships = masterPart is null ? new Dictionary<string, OoxRelationship>() : ReadRelationships(package, masterPart.Name, cancellationToken);
                IReadOnlyDictionary<string, OoxRelationship> layoutRelationships = layoutPart is null ? new Dictionary<string, OoxRelationship>() : ReadRelationships(package, layoutPart.Name, cancellationToken);
                IReadOnlyDictionary<string, OoxRelationship> slideRelationships = ReadRelationships(package, slide.PartName, cancellationToken);
                IReadOnlyList<XDocument> layoutSources = masterXml is null ? [] : [masterXml];
                IReadOnlyList<XDocument> slideSources = masterXml is null
                    ? layoutXml is null ? [] : [layoutXml]
                    : layoutXml is null ? [masterXml] : [masterXml, layoutXml];
                PptxColorMap masterColorMap = ReadMasterColorMap(masterXml);
                PptxColorMap layoutColorMap = ReadColorMapOverride(layoutXml, masterColorMap);
                PptxColorMap slideColorMap = ReadColorMapOverride(slideXml, layoutColorMap);
                slides.Add(new PptxSceneSlide(
                    slide.Index,
                    slide.PartName,
                    theme,
                    masterPart?.Name,
                    layoutPart?.Name,
                    masterXml,
                    layoutXml,
                    slideXml,
                    masterRelationships,
                    layoutRelationships,
                    slideRelationships,
                    masterColorMap,
                    layoutColorMap,
                    slideColorMap,
                    ReadBackground(masterXml, theme, masterColorMap),
                    ReadBackground(layoutXml, theme, layoutColorMap),
                    ReadBackground(slideXml, theme, slideColorMap),
                    HasSlideTransition(slideXml),
                    HasSlideTiming(slideXml),
                    HasSlideOleObject(slideXml),
                    ReadShowMasterShapes(slideXml, layoutXml),
                    masterXml is null ? [] : ReadNodes(masterXml, [], theme, masterColorMap, package, masterRelationships, cancellationToken, document.DefaultTextStyle),
                    layoutXml is null ? [] : ReadNodes(layoutXml, layoutSources, theme, layoutColorMap, package, layoutRelationships, cancellationToken, document.DefaultTextStyle),
                    ReadNodes(slideXml, slideSources, theme, slideColorMap, package, slideRelationships, cancellationToken, document.DefaultTextStyle)));
            }

            return result = new PptxScene(slides);
        }
        finally
        {
            if (result is null) modelStorage?.Dispose();
            else if (modelStorage is not null) OoxMemory.Resize(modelStorage, modelStorage.Bytes / 2);
        }

        PptxTheme? ReadTheme(string? ownerPartName)
        {
            cancellationToken.ThrowIfCancellationRequested();
            OoxPart? part = ownerPartName is null ? null :
                GetRelatedPart(package, ownerPartName, ThemeRelationshipType, cancellationToken);
            if (part is null) return null;
            if (!themes.TryGetValue(part.Name, out PptxTheme? parsed))
            {
                XDocument themeXml = package.LoadXml(part, cancellationToken);
                OoxMemory.GrowModel(modelStorage, themeXml, cancellationToken);
                themes.Add(part.Name, parsed = PptxTheme.Parse(themeXml, cancellationToken));
            }
            return parsed;
        }

        PptxTheme GetFallbackTheme()
        {
            if (fallbackTheme is not null) return fallbackTheme;
            var relationships = package.GetRelationshipDictionary(document.PresentationPartName, cancellationToken).Values;
            string? owner = relationships.Any(r => !r.IsExternal && r.Type == ThemeRelationshipType && r.ResolvedTarget is not null)
                ? document.PresentationPartName
                : relationships.FirstOrDefault(r => !r.IsExternal && r.Type == SlideMasterRelationshipType && r.ResolvedTarget is not null)?.ResolvedTarget;
            return fallbackTheme = ReadTheme(owner) ?? PptxTheme.Empty;
        }

        bool HasSlideOleObject(XDocument slideXml)
        {
            return slideXml.Descendants(PresentationNamespace + "oleObj").Any(IsVisibleDrawingContent);
        }

        bool HasSlideTiming(XDocument slideXml) => slideXml.Root?.Element(PresentationNamespace + "timing") is not null;

        // Both owners must allow master artwork. PowerPoint keeps a layout's
        // suppression even when the slide explicitly allows it; absent means show.
        bool ReadShowMasterShapes(XDocument slideXml, XDocument? layoutXml)
        {
            return !OoxBoolean.IsOff((string?)slideXml.Root?.Attribute("showMasterSp")) &&
                !OoxBoolean.IsOff((string?)layoutXml?.Root?.Attribute("showMasterSp"));
        }

        bool HasSlideTransition(XDocument slideXml) => slideXml.Root?.Element(PresentationNamespace + "transition") is not null;
    }

    private static PptxColorMap ReadMasterColorMap(XDocument? xml)
    {
        return PptxColorMap.FromElement(xml?.Root?.Element(PresentationNamespace + "clrMap"));
    }

    private static PptxColorMap ReadColorMapOverride(XDocument? xml, PptxColorMap inheritedColorMap)
    {
        XElement? overrideColorMap = xml?.Root?
            .Element(PresentationNamespace + "clrMapOvr")?
            .Element(DrawingNamespace + "overrideClrMapping");
        return PptxColorMap.FromElement(overrideColorMap, inheritedColorMap);
    }

    private static OoxPart? GetRelatedPart(OoxPackage package, string sourcePartName, string relationshipType, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxRelationship? relationship = package.GetRelationshipDictionary(sourcePartName, cancellationToken).Values
            .FirstOrDefault(r => !r.IsExternal && r.Type == relationshipType && r.ResolvedTarget is not null);
        return relationship?.ResolvedTarget is null ? null : package.GetPart(relationship.ResolvedTarget);
    }

    private static IReadOnlyDictionary<string, OoxRelationship> ReadRelationships(OoxPackage package, string sourcePartName, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        // External hyperlink targets are retained so shape-level hlinkClick entries can resolve;
        // all other external relationships stay excluded from resource resolution.
        // The package dictionary is shared (parsed once); the filtered view is rebuilt per
        // caller cheaply without reparsing.
        return package.GetRelationshipDictionary(sourcePartName, cancellationToken)
            .Where(pair => (!pair.Value.IsExternal && pair.Value.ResolvedTarget is not null) || (pair.Value.IsExternal && pair.Value.Type.Equals(HyperlinkRelationshipType, StringComparison.Ordinal)))
            .ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.Ordinal);
    }

    private static IReadOnlyList<PptxSceneNode> ReadNodes(
        XDocument xml,
        IReadOnlyList<XDocument> placeholderSources,
        PptxTheme theme,
        PptxColorMap colorMap,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        CancellationToken cancellationToken,
        XElement? defaultTextStyle)
    {
        return GetShapeTree(xml) is { } shapeTree
            ? ReadChildNodes(shapeTree, placeholderSources, theme, colorMap, package, relationships, cancellationToken, defaultTextStyle)
            : [];
    }

    // Each slide, layout or master owns its scene through common slide data.
    internal static XElement? GetShapeTree(XDocument source) =>
        source.Root?.Element(PresentationNamespace + "cSld")?.Element(PresentationNamespace + "spTree");

    private static PptxSceneBackground ReadBackground(XDocument? xml, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? background = xml?.Root?
            .Element(PresentationNamespace + "cSld")?
            .Element(PresentationNamespace + "bg");
        XElement? properties = background?.Element(PresentationNamespace + "bgPr");
        XElement? reference = properties is null ? background?.Element(PresentationNamespace + "bgRef") : null;
        XElement? gradientSource = properties?.Element(DrawingNamespace + "gradFill");
        if (reference is not null)
        {
            int index = PptxFormatSchemeResolver.ReadIndex(reference, defaultValue: -1);
            if (index is 0 or 1000)
                return default;

            XElement? style = theme.GetBackgroundFillStyle(index);
            if (style?.Name == DrawingNamespace + "noFill")
                return default;

            properties = style?.Name == DrawingNamespace + "solidFill" ? style : null;
            gradientSource = style?.Name == DrawingNamespace + "gradFill" ? style : null;
        }

        if (PptxColorResolver.TryReadSolidColorWithAlpha(properties, theme, colorMap, reference, out RgbColor color, out double alpha))
        {
            return new PptxSceneBackground(true, color, alpha);
        }

        if (gradientSource is not null &&
            ReadLinearGradient(gradientSource, theme, colorMap, reference) is { HasGradient: true, HasUnsupportedGradient: false } gradient)
        {
            return new PptxSceneBackground(true, default, 1d, gradient);
        }

        return reference is not null ? new PptxSceneBackground { HasUnsupportedStyleReference = true }
            : new PptxSceneBackground { HasUnsupportedGradient = gradientSource is not null };
    }

    private static IReadOnlyList<PptxSceneNode> ReadChildNodes(
        XElement container,
        IReadOnlyList<XDocument> placeholderSources,
        PptxTheme theme,
        PptxColorMap colorMap,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        CancellationToken cancellationToken,
        XElement? defaultTextStyle)
    {
        var nodes = new List<PptxSceneNode>();
        foreach (XElement child in container.Elements())
        {
            cancellationToken.ThrowIfCancellationRequested();
            PptxSceneNodeKind kind = ReadNodeKind(child);
            if (kind == PptxSceneNodeKind.Unknown)
            {
                continue;
            }

            XElement? nonVisualProperties = ReadNonVisualProperties(child);
            // Exclude hidden objects before their text, images, charts or group
            // children enter resource preflight and painting.
            if (IsHiddenNonVisualProperties(nonVisualProperties))
            {
                continue;
            }

            nodes.Add(PptxSceneNode.Create(
                kind,
                ReadNonVisualId(nonVisualProperties),
                ReadNonVisualName(nonVisualProperties),
                IsPlaceholder(child),
                kind == PptxSceneNodeKind.UnknownGraphicFrame && IsSmartArtGraphicData(GetGraphicData(child)),
                ReadHyperlinkClick(nonVisualProperties),
                ReadBounds(child),
                kind is PptxSceneNodeKind.Shape or PptxSceneNodeKind.Connector ? ReadShape(child, theme, colorMap, package, relationships) : null,
                ReadTextBody(child, placeholderSources, defaultTextStyle),
                kind == PptxSceneNodeKind.Picture ? ReadPicture(child, theme, colorMap, package, relationships, cancellationToken) : null,
                kind == PptxSceneNodeKind.Table ? ReadTable(child, theme, colorMap) : null,
                kind == PptxSceneNodeKind.Chart ? ReadChart(child, package, theme, colorMap, relationships, cancellationToken) : null,
                kind == PptxSceneNodeKind.Group ? ReadGroupTransform(child) : PptxSceneGroupTransform.Identity,
                kind == PptxSceneNodeKind.Group ? ReadChildNodes(child, placeholderSources, theme, colorMap, package, relationships, cancellationToken, defaultTextStyle) : [],
                child));
        }

        return nodes;
    }

    private static PptxSceneNodeKind ReadNodeKind(XElement element)
    {
        if (element.Name == PresentationNamespace + "sp")
        {
            return PptxSceneNodeKind.Shape;
        }

        if (element.Name == PresentationNamespace + "cxnSp")
        {
            return PptxSceneNodeKind.Connector;
        }

        if (element.Name == PresentationNamespace + "pic")
        {
            return PptxSceneNodeKind.Picture;
        }

        if (element.Name == PresentationNamespace + "grpSp")
        {
            return PptxSceneNodeKind.Group;
        }

        if (element.Name != PresentationNamespace + "graphicFrame")
        {
            return PptxSceneNodeKind.Unknown;
        }

        return ReadGraphicFrameKind(GetGraphicData(element));
    }

    // Nonvisual extensions may contain previews of other graphics. Only the
    // frame's direct graphic owns the payload used for drawing and diagnostics.
    internal static XElement? GetGraphicData(XElement frame) =>
        frame.Element(DrawingNamespace + "graphic")?.Element(DrawingNamespace + "graphicData");

    internal static PptxSceneNodeKind ReadGraphicFrameKind(XElement? graphicData)
    {
        string uri = (string?)graphicData?.Attribute("uri") ?? string.Empty;
        if (graphicData?.Element(DrawingNamespace + "tbl") is not null)
        {
            return PptxSceneNodeKind.Table;
        }

        return uri.Contains("chart", StringComparison.OrdinalIgnoreCase)
            ? PptxSceneNodeKind.Chart
            : PptxSceneNodeKind.UnknownGraphicFrame;
    }

    internal static bool IsSmartArtGraphicData(XElement? graphicData) =>
        ((string?)graphicData?.Attribute("uri"))?.Contains("drawingml/2006/diagram", StringComparison.OrdinalIgnoreCase) == true;

    // Raw diagnostic scans can encounter drawing markup copied into extensions.
    internal static bool IsVisibleDrawingContent(XElement element) =>
        !element.AncestorsAndSelf().Any(owner =>
            owner.Name == PresentationNamespace + "extLst" || owner.Name == DrawingNamespace + "extLst" ||
            IsHiddenNonVisualProperties(ReadNonVisualProperties(owner)));

    private static bool IsHiddenNonVisualProperties(XElement? properties) =>
        properties is not null && OoxXml.ReadBool(properties, "hidden");

    private static XElement? ReadNonVisualProperties(XElement element)
    {
        return element
            .Elements()
            .FirstOrDefault(e => e.Name.LocalName is "nvSpPr" or "nvPicPr" or "nvGrpSpPr" or "nvGraphicFramePr" or "nvCxnSpPr")
            ?.Elements()
            .FirstOrDefault(e => e.Name.LocalName == "cNvPr");
    }

    private static string ReadNonVisualId(XElement? nonVisualProperties)
    {
        return (string?)nonVisualProperties?.Attribute("id") ?? string.Empty;
    }

    private static string ReadNonVisualName(XElement? nonVisualProperties)
    {
        return (string?)nonVisualProperties?.Attribute("name") ?? string.Empty;
    }

    internal static PptxSceneHyperlinkClick ReadHyperlinkClick(XElement? properties)
    {
        XElement? hyperlink = properties?.Element(DrawingNamespace + "hlinkClick");
        if (hyperlink is null)
        {
            return default;
        }

        return new PptxSceneHyperlinkClick(
            true,
            (string?)hyperlink.Attribute(RelationshipsNamespace + "id"),
            (string?)hyperlink.Attribute("action"));
    }

    private static bool IsPlaceholder(XElement element)
    {
        return element
            .Elements()
            .FirstOrDefault(e => e.Name.LocalName is "nvSpPr" or "nvPicPr" or "nvGrpSpPr" or "nvGraphicFramePr" or "nvCxnSpPr")
            ?.Elements()
            .FirstOrDefault(e => e.Name.LocalName == "nvPr")
            ?.Element(PresentationNamespace + "ph") is not null;
    }

    private static PptxSceneBounds? ReadBounds(XElement element)
    {
        XElement? transform = element
            .Element(PresentationNamespace + "spPr")?
            .Element(DrawingNamespace + "xfrm") ??
            element.Element(PresentationNamespace + "grpSpPr")?
                .Element(DrawingNamespace + "xfrm") ??
            element.Element(PresentationNamespace + "xfrm");
        if (transform is null)
        {
            return null;
        }

        XElement? offset = transform.Element(DrawingNamespace + "off");
        XElement? extents = transform.Element(DrawingNamespace + "ext");
        if (offset is null || extents is null)
        {
            return null;
        }

        return new PptxSceneBounds(
            OoxXml.ParseOptionalLong(offset, "x", 0L),
            OoxXml.ParseOptionalLong(offset, "y", 0L),
            OoxXml.ParseOptionalLong(extents, "cx", 0L),
            OoxXml.ParseOptionalLong(extents, "cy", 0L),
            transform.Attribute("rot") is { } rotation ? long.Parse(rotation.Value, CultureInfo.InvariantCulture) / 60000d : 0d,
            OoxXml.ReadBool(transform, "flipH"),
            OoxXml.ReadBool(transform, "flipV"));
    }

    private static PptxScenePicture ReadPicture(
        XElement picture,
        PptxTheme theme,
        PptxColorMap colorMap,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        CancellationToken cancellationToken)
    {
        XElement? presentationFill = picture.Element(PresentationNamespace + "blipFill");
        XElement? blip = (presentationFill ?? picture.Element(DrawingNamespace + "blipFill"))?.Element(DrawingNamespace + "blip");
        string? svgRelationshipId = OoxXml.ReadDrawingSvgRelationshipId(blip, cancellationToken);
        string? relationshipId = (presentationFill is null ? null : (string?)blip?.Attribute(RelationshipsNamespace + "embed"))
            ?? svgRelationshipId;
        string? targetPartName = ResolveRelationshipTarget(relationshipId, relationships);
        XElement? shapeProperties = picture.Element(PresentationNamespace + "spPr");
        PptxSceneLineStyle line = ReadShapeLine(shapeProperties, theme, colorMap, default, widthSpecifiedByDefault: false);
        return new PptxScenePicture(
            relationshipId,
            targetPartName,
            ReadImageResource(package, targetPartName),
            ReadPictureCrop(picture),
            ReadPictureFill(picture),
            ReadPictureAlpha(picture),
            ReadPictureAlphaValue(picture),
            ReadImageRecolor(picture, theme, colorMap),
            HasPictureVideo(),
            HasPictureAudio(),
            HasPictureTile(picture),
            line,
            TryReadOuterShadow(shapeProperties, theme, colorMap, out PptxSceneOuterShadow outerShadow) ? outerShadow : default,
            ReadPresetGeometry(shapeProperties),
            ReadPresetAdjustments(shapeProperties),
            shapeProperties?.Element(DrawingNamespace + "custGeom") is not null)
        {
            SvgFallbackRelationshipId = svgRelationshipId,
            SvgFallback = svgRelationshipId != relationshipId
                ? ReadImageResource(package, ResolveRelationshipTarget(svgRelationshipId, relationships)) : null
        };

        bool HasPictureAudio()
        {
            return picture.Descendants(PresentationNamespace + "audio").Any(IsVisibleDrawingContent) ||
                picture.Descendants(DrawingNamespace + "audioFile").Any(IsVisibleDrawingContent);
        }

        bool HasPictureVideo()
        {
            return picture.Descendants(PresentationNamespace + "video").Any(IsVisibleDrawingContent) ||
                picture.Descendants(DrawingNamespace + "videoFile").Any(IsVisibleDrawingContent);
        }
    }
}
