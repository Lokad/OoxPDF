using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Pptx;

internal sealed record PptxScene(PptxDocument Document, PptxTheme Theme, IReadOnlyList<PptxSceneSlide> Slides);

internal sealed record PptxSceneSlide(
    int Index,
    string PartName,
    XDocument SlideXml,
    IReadOnlyList<PptxSceneNode> MasterNodes,
    IReadOnlyList<PptxSceneNode> LayoutNodes,
    IReadOnlyList<PptxSceneNode> SlideNodes);

internal sealed record PptxSceneNode(
    PptxSceneNodeKind Kind,
    string Id,
    string Name,
    bool IsPlaceholder,
    PptxSceneBounds? Bounds,
    XElement Source);

internal sealed record PptxSceneBounds(
    double X,
    double Y,
    double Width,
    double Height,
    double RotationDegrees,
    bool FlipHorizontal,
    bool FlipVertical);

internal enum PptxSceneNodeKind
{
    Shape,
    Picture,
    Table,
    Chart,
    Group,
    Connector,
    UnknownGraphicFrame,
    Unknown
}

internal sealed class PptxSceneBuilder
{
    private static readonly XNamespace PresentationNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string SlideLayoutRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
    private const string SlideMasterRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";

    public PptxScene Build(PptxDocument document, OoxPackage package)
    {
        PptxTheme theme = PptxTheme.Load(package, document.PresentationPartName);
        var slides = new List<PptxSceneSlide>(document.Slides.Count);
        foreach (PptxSlide slide in document.Slides)
        {
            OoxPart? slidePart = package.GetPart(slide.PartName);
            if (slidePart is null)
            {
                slides.Add(new PptxSceneSlide(slide.Index, slide.PartName, new XDocument(), [], [], []));
                continue;
            }

            XDocument slideXml = LoadXml(slidePart);
            OoxPart? layoutPart = GetRelatedPart(package, slide.PartName, SlideLayoutRelationshipType);
            OoxPart? masterPart = layoutPart is null ? null : GetRelatedPart(package, layoutPart.Name, SlideMasterRelationshipType);
            slides.Add(new PptxSceneSlide(
                slide.Index,
                slide.PartName,
                slideXml,
                masterPart is null ? [] : ReadNodes(LoadXml(masterPart)),
                layoutPart is null ? [] : ReadNodes(LoadXml(layoutPart)),
                ReadNodes(slideXml)));
        }

        return new PptxScene(document, theme, slides);
    }

    private static XDocument LoadXml(OoxPart part)
    {
        using Stream stream = part.OpenRead();
        return SafeXml.Load(stream);
    }

    private static OoxPart? GetRelatedPart(OoxPackage package, string sourcePartName, string relationshipType)
    {
        OoxRelationship? relationship = package.GetRelationships(sourcePartName)
            .FirstOrDefault(r => !r.IsExternal && r.Type == relationshipType && r.ResolvedTarget is not null);
        return relationship?.ResolvedTarget is null ? null : package.GetPart(relationship.ResolvedTarget);
    }

    private static IReadOnlyList<PptxSceneNode> ReadNodes(XDocument xml)
    {
        var nodes = new List<PptxSceneNode>();
        foreach (XElement shapeTree in xml.Descendants(PresentationNamespace + "spTree"))
        {
            foreach (XElement child in shapeTree.Elements())
            {
                PptxSceneNodeKind kind = ReadKind(child);
                if (kind == PptxSceneNodeKind.Unknown)
                {
                    continue;
                }

                (string id, string name) = ReadNonVisualProperties(child);
                nodes.Add(new PptxSceneNode(kind, id, name, IsPlaceholder(child), ReadBounds(child), child));
            }
        }

        return nodes;
    }

    private static PptxSceneNodeKind ReadKind(XElement element)
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

        XElement? graphicData = element
            .Descendants(DrawingNamespace + "graphicData")
            .FirstOrDefault();
        string uri = (string?)graphicData?.Attribute("uri") ?? string.Empty;
        if (graphicData?.Descendants(DrawingNamespace + "tbl").Any() == true)
        {
            return PptxSceneNodeKind.Table;
        }

        return uri.Contains("chart", StringComparison.OrdinalIgnoreCase)
            ? PptxSceneNodeKind.Chart
            : PptxSceneNodeKind.UnknownGraphicFrame;
    }

    private static (string Id, string Name) ReadNonVisualProperties(XElement element)
    {
        XElement? nonVisual = element
            .Elements()
            .FirstOrDefault(e => e.Name.LocalName is "nvSpPr" or "nvPicPr" or "nvGrpSpPr" or "nvGraphicFramePr" or "nvCxnSpPr")
            ?.Elements()
            .FirstOrDefault(e => e.Name.LocalName == "cNvPr");
        return ((string?)nonVisual?.Attribute("id") ?? string.Empty, (string?)nonVisual?.Attribute("name") ?? string.Empty);
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
            OoxUnits.EmuToPoints(ReadLong(offset, "x")),
            OoxUnits.EmuToPoints(ReadLong(offset, "y")),
            OoxUnits.EmuToPoints(ReadLong(extents, "cx")),
            OoxUnits.EmuToPoints(ReadLong(extents, "cy")),
            transform.Attribute("rot") is { } rotation ? long.Parse(rotation.Value, CultureInfo.InvariantCulture) / 60000d : 0d,
            ReadBool(transform, "flipH"),
            ReadBool(transform, "flipV"));
    }

    private static long ReadLong(XElement element, string name)
    {
        return element.Attribute(name) is { } attribute
            ? long.Parse(attribute.Value, CultureInfo.InvariantCulture)
            : 0L;
    }

    private static bool ReadBool(XElement element, string name)
    {
        string? value = (string?)element.Attribute(name);
        return value is "1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
    }
}
