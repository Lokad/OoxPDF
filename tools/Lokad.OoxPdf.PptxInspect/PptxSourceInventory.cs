using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

internal static class PptxSourceInventory
{
    private const string LayoutRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
    private const string MasterRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
    private const string DiagramNamespace = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
    private static readonly string[] RecolorKinds = ["grayscl", "biLevel", "lum", "duotone"];

    public static object Read(OoxPackage package)
    {
        OoxPart presentation = PptxReader.ReadPresentationPart(package, default);
        XDocument document = package.LoadXml(presentation, default);
        var parts = new Dictionary<string, Dictionary<string, object>>(StringComparer.OrdinalIgnoreCase);
        IReadOnlyDictionary<string, OoxRelationship> empty = new Dictionary<string, OoxRelationship>();

        // Report authored slides, including hidden ones. The renderer filters
        // visibility after it resolves these same slide owners.
        var slides = PptxReader.ReadSlideSources(package, presentation, document, default)
            .Select((source, index) =>
            {
                string slide = source.PartName;
                var slideRelationships = package.GetRelationshipDictionary(slide, default);
                string? layout = RelatedTarget(slideRelationships, LayoutRelationship);
                var layoutRelationships = layout is null ? empty : package.GetRelationshipDictionary(layout, default);
                string? master = RelatedTarget(layoutRelationships, MasterRelationship);
                return new
                {
                    SlideNumber = index + 1,
                    SlidePart = slide,
                    LayoutPart = layout,
                    MasterPart = master,
                    SlideContent = ReadPart(slide, slideRelationships),
                    Layout = ReadPart(layout, layoutRelationships),
                    Master = ReadPart(master, master is null ? empty : package.GetRelationshipDictionary(master, default)),
                    Relationships = SummarizeRelationships(slideRelationships),
                    LayoutRelationships = SummarizeRelationships(layoutRelationships)
                };
            }).ToArray();

        return new
        {
            Summary = new
            {
                SlideCount = slides.Length,
                SlidesWithCharts = slides.Count(s => Count(s.SlideContent, "Charts") > 0 || s.Relationships["Charts"] > 0),
                SlidesWithTables = slides.Count(s => Count(s.SlideContent, "Tables") > 0),
                SlidesWithPictures = slides.Count(s => Count(s.SlideContent, "Pictures") > 0 || s.Relationships["Images"] > 0),
                SlidesWithGroups = slides.Count(s => Count(s.SlideContent, "GroupShapes") > 0),
                SlidesWithEffects = slides.Count(s => Count(s.SlideContent, "Effects") + Count(s.Layout, "Effects") + Count(s.Master, "Effects") > 0),
                SlidesWithTransparency = slides.Count(s => Count(s.SlideContent, "Transparency") + Count(s.Layout, "Transparency") + Count(s.Master, "Transparency") > 0),
                SlidesWithPictureRecolor = slides.Count(s => Count(s.SlideContent, "PictureRecolors") + Count(s.Layout, "PictureRecolors") + Count(s.Master, "PictureRecolors") > 0),
                SlidesWithInheritedContent = slides.Count(s => Count(s.Layout, "Shapes") + Count(s.Layout, "Pictures") + Count(s.Master, "Shapes") + Count(s.Master, "Pictures") > 0),
                TotalSlideShapes = slides.Sum(s => (long)Count(s.SlideContent, "Shapes")),
                TotalSlidePictures = slides.Sum(s => (long)Count(s.SlideContent, "Pictures")),
                TotalSlidePictureRecolors = slides.Sum(s => (long)Count(s.SlideContent, "PictureRecolors")),
                TotalSlideTables = slides.Sum(s => (long)Count(s.SlideContent, "Tables")),
                TotalSlideCharts = slides.Sum(s => (long)Count(s.SlideContent, "Charts"))
            },
            Slides = slides
        };

        Dictionary<string, object>? ReadPart(string? name, IReadOnlyDictionary<string, OoxRelationship> relationships)
        {
            if (name is null || package.GetPart(name) is not { } part) return null;
            if (parts.TryGetValue(part.Name, out var inventory)) return inventory;
            inventory = CountPart(part.Name, package.LoadXml(part, default), package, relationships);
            parts.Add(part.Name, inventory);
            return inventory;
        }
    }

    private static int Count(Dictionary<string, object>? part, string name) => part is null ? 0 : (int)part[name];

    private static string? RelatedTarget(IReadOnlyDictionary<string, OoxRelationship> relationships, string type) =>
        relationships.Values.FirstOrDefault(r => !r.IsExternal && r.Type == type && r.ResolvedTarget is not null)?.ResolvedTarget;

    private static Dictionary<string, object> CountPart(string name, XDocument document, OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        // Count source XML, including stored properties and unsupported constructs.
        var elements = document.Descendants().ToLookup(e => e.Name.LocalName, StringComparer.Ordinal);
        int Count(string localName) => elements[localName].Count();
        int effects = Count("effectLst") + Count("effectDag");
        int transparency = Count("alpha");
        var pictureRecolorKinds = new Dictionary<string, int>(StringComparer.Ordinal);
        var pictureRecolorProfiles = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (XElement blip in elements["blip"])
        {
            foreach (XElement child in blip.Elements().Where(e => RecolorKinds.Contains(e.Name.LocalName)))
            {
                string kind = child.Name.LocalName;
                pictureRecolorKinds[kind] = pictureRecolorKinds.GetValueOrDefault(kind) + 1;
            }

            string? preferred = RecolorKinds.FirstOrDefault(kind => blip.Elements().Any(e => e.Name.LocalName == kind));
            if (preferred is null) continue;

            string? id = (string?)blip.Attribute(RelationshipsNamespace + "embed");
            string contentType = id is not null && relationships.TryGetValue(id, out var relationship) &&
                relationship.Type.EndsWith("/image", StringComparison.OrdinalIgnoreCase) && relationship.ResolvedTarget is { } target
                ? package.GetPart(target)?.ContentType ?? "(unknown)"
                : "(unresolved)";
            string profile = $"{preferred}|{contentType}";
            pictureRecolorProfiles[profile] = pictureRecolorProfiles.GetValueOrDefault(profile) + 1;
        }

        return new Dictionary<string, object>
        {
            ["PartName"] = name,
            ["Shapes"] = Count("sp"),
            ["ConnectorShapes"] = Count("cxnSp"),
            ["GroupShapes"] = Count("grpSp"),
            ["Pictures"] = Count("pic"),
            ["GraphicFrames"] = Count("graphicFrame"),
            ["Tables"] = Count("tbl"),
            ["Charts"] = Count("chart"),
            ["TextBodies"] = Count("txBody"),
            ["Placeholders"] = Count("ph"),
            ["Transforms"] = Count("xfrm"),
            ["RotatedTransforms"] = elements["xfrm"].Count(e => e.Attribute("rot") is not null),
            ["FlippedTransforms"] = elements["xfrm"].Count(e => e.Attribute("flipH") is not null || e.Attribute("flipV") is not null),
            ["SolidFills"] = Count("solidFill"),
            ["GradientFills"] = Count("gradFill"),
            ["PatternFills"] = Count("pattFill"),
            ["PictureFills"] = Count("blipFill"),
            ["PictureRecolors"] = pictureRecolorKinds.Values.Sum(),
            ["PictureRecolorKinds"] = pictureRecolorKinds,
            ["PictureRecolorProfiles"] = pictureRecolorProfiles,
            ["Transparency"] = transparency,
            ["Effects"] = effects,
            ["Clips"] = Count("srcRect"),
            ["SmartArtSignals"] = elements.SelectMany(group => group).Count(e => e.Name.NamespaceName == DiagramNamespace),
            ["MediaSignals"] = Count("video") + Count("audio") + Count("videoFile") + Count("audioFile"),
            ["OleSignals"] = Count("oleObj"),
            ["PresetGeometries"] = Count("prstGeom"),
            ["CustomGeometries"] = Count("custGeom"),
            ["PresetGeometryKinds"] = Histogram(elements["prstGeom"].Select(e => AttributeOrMissing(e, "prst"))),
            ["PresetShapeProfiles"] = Histogram(elements["spPr"].Select(ShapeProfile).OfType<string>()),
            ["LineHeadEndKinds"] = Histogram(elements["headEnd"].Select(e => AttributeOrMissing(e, "type"))),
            ["LineTailEndKinds"] = Histogram(elements["tailEnd"].Select(e => AttributeOrMissing(e, "type"))),
            ["TextBodyVerticalModes"] = Histogram(elements["bodyPr"].Where(e => e.Attribute("vert") is not null)
                .Select(e => AttributeOrMissing(e, "vert"))),
            ["Notes"] = new
            {
                HasBackground = Count("bg") != 0,
                HasHiddenPlaceholder = elements["ph"].Any(e => (string?)e.Attribute("hidden") is "1" or "true"),
                HasGroupedContent = Count("grpSp") != 0,
                HasEffects = effects != 0,
                HasTransparency = transparency != 0
            }
        };
    }

    private static string? ShapeProfile(XElement properties)
    {
        XElement? Child(XElement? parent, string name) => parent?.Elements().FirstOrDefault(e => e.Name.LocalName == name);
        if (Child(properties, "prstGeom") is not { } geometry) return null;

        XElement? transform = Child(properties, "xfrm");
        XElement? line = Child(properties, "ln");
        string owner = properties.Parent?.Name.LocalName ?? "(missing)";
        string preset = AttributeOrMissing(geometry, "prst");
        string rotation = transform?.Attribute("rot") is null ? "noRot" : "rot";
        string flipH = transform?.Attribute("flipH") is null ? "noFlipH" : "flipH";
        string flipV = transform?.Attribute("flipV") is null ? "noFlipV" : "flipV";
        string head = AttributeOrMissing(Child(line, "headEnd"), "type");
        string tail = AttributeOrMissing(Child(line, "tailEnd"), "type");
        return $"{owner}|{preset}|{rotation}|{flipH}|{flipV}|head={head}|tail={tail}";
    }

    private static string AttributeOrMissing(XElement? element, string name)
    {
        string? value = (string?)element?.Attribute(name);
        return string.IsNullOrWhiteSpace(value) ? "(missing)" : value;
    }

    private static Dictionary<string, int> Histogram(IEnumerable<string> values) =>
        values.GroupBy(value => value, StringComparer.Ordinal).ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);

    private static Dictionary<string, int> SummarizeRelationships(IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        var counts = new Dictionary<string, int>
        {
            ["Images"] = 0, ["Charts"] = 0, ["SlideLayouts"] = 0, ["SlideMasters"] = 0, ["Themes"] = 0,
            ["Hyperlinks"] = 0, ["OleObjects"] = 0, ["Media"] = 0, ["Other"] = 0, ["External"] = 0
        };
        foreach (var relationship in relationships.Values)
        {
            if (relationship.IsExternal) counts["External"]++;
            int separator = relationship.Type.LastIndexOf('/');
            string category = (separator < 0 ? "" : relationship.Type[(separator + 1)..].ToLowerInvariant()) switch
            {
                "image" => "Images",
                "chart" => "Charts",
                "slidelayout" => "SlideLayouts",
                "slidemaster" => "SlideMasters",
                "theme" => "Themes",
                "hyperlink" => "Hyperlinks",
                "oleobject" => "OleObjects",
                "audio" or "video" or "media" => "Media",
                _ => "Other"
            };
            counts[category]++;
        }
        return counts;
    }
}
