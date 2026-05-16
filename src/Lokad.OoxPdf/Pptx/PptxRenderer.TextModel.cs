using System.Globalization;
using System.Xml.Linq;

using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    internal static IReadOnlyList<PptxTextFrameModelSnapshot> InspectTextFrameModels(PptxDocument document, OoxPackage package, int slideIndex)
    {
        if (slideIndex < 0 || slideIndex >= document.Slides.Count)
        {
            return [];
        }

        PptxTheme theme = PptxTheme.Load(package, document.PresentationPartName);
        PptxSlide slide = document.Slides[slideIndex];
        OoxPart? slidePart = package.GetPart(slide.PartName);
        if (slidePart is null)
        {
            return [];
        }

        using Stream stream = slidePart.OpenRead();
        XDocument slideXml = SafeXml.Load(stream);
        IReadOnlyList<XDocument> inheritedXml = LoadInheritedSlideXml(package, slide.PartName);
        return inheritedXml
            .SelectMany(xml => BuildTextFrameModels(xml, document, theme, slideIndex + 1, includePlaceholders: false, placeholderSources: []))
            .Concat(BuildTextFrameModels(slideXml, document, theme, slideIndex + 1, includePlaceholders: true, inheritedXml))
            .Select(ToSnapshot)
            .ToArray();
    }

    private static PptxTextFrameModelSnapshot ToSnapshot(PptxTextFrameModel frame)
    {
        return new PptxTextFrameModelSnapshot(
            frame.TextX,
            frame.TextWidth,
            frame.FontScale,
            frame.Paragraphs.Select(ToSnapshot).ToArray());
    }

    private static PptxTextParagraphModelSnapshot ToSnapshot(PptxTextParagraphModel paragraph)
    {
        return new PptxTextParagraphModelSnapshot(
            paragraph.Level,
            paragraph.Style.Alignment.ToString(),
            paragraph.Style.FontSize,
            paragraph.Runs.Select(ToSnapshot).ToArray());
    }

    private static PptxTextRunModelSnapshot ToSnapshot(PptxTextRunModel run)
    {
        return new PptxTextRunModelSnapshot(
            run.Kind.ToString(),
            run.Text,
            run.Style.FontSize,
            run.Style.CharacterSpacing,
            run.Style.Typeface,
            run.Style.Underline,
            run.Style.Highlight);
    }

    private static IReadOnlyList<PptxTextFrameModel> BuildTextFrameModels(
        XDocument slideXml,
        PptxDocument document,
        PptxTheme theme,
        int slideNumber,
        bool includePlaceholders,
        IReadOnlyList<XDocument> placeholderSources)
    {
        var frames = new List<PptxTextFrameModel>();
        foreach (XElement shape in slideXml.Descendants(PresentationNamespace + "sp"))
        {
            PptxTextFrameModel? frame = BuildTextFrameModel(shape, document, theme, slideNumber, includePlaceholders, placeholderSources);
            if (frame is not null)
            {
                frames.Add(frame);
            }
        }

        return frames;
    }

    private static PptxTextFrameModel? BuildTextFrameModel(
        XElement shape,
        PptxDocument document,
        PptxTheme theme,
        int slideNumber,
        bool includePlaceholders,
        IReadOnlyList<XDocument> placeholderSources)
    {
        if (!includePlaceholders && IsPlaceholder(shape))
        {
            return null;
        }

        XElement? shapeProperties = shape.Element(PresentationNamespace + "spPr");
        XElement? textBody = shape.Element(PresentationNamespace + "txBody");
        IReadOnlyList<XElement> inheritedPlaceholders = FindInheritedPlaceholderShapes(shape, placeholderSources);
        XElement? inheritedPlaceholder = inheritedPlaceholders.LastOrDefault();
        XElement? inheritedTextBody = inheritedPlaceholder?.Element(PresentationNamespace + "txBody");
        ShapeBounds? bounds = shapeProperties is null ? null : ReadBounds(shapeProperties);
        bounds ??= inheritedPlaceholder?.Element(PresentationNamespace + "spPr") is { } inheritedProperties
            ? ReadBounds(inheritedProperties)
            : null;
        if (bounds is null || textBody is null)
        {
            return null;
        }

        bounds = ReadAncestorGroupTransform(shape).Apply(bounds.Value);
        double x = OoxUnits.EmuToPoints(bounds.Value.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Value.Y);
        double width = OoxUnits.EmuToPoints(bounds.Value.Width);
        double height = OoxUnits.EmuToPoints(bounds.Value.Height);
        TextInsets insets = ReadTextInsets(textBody);
        double fontScale = ReadNormAutofitFontScale(textBody);
        double textX = x + insets.Left;
        double textWidth = Math.Max(1d, width - insets.Left - insets.Right);
        double textHeight = Math.Max(1d, height - insets.Top - insets.Bottom);
        double rotationCenterX = x + width / 2d;
        double rotationCenterY = document.SlideHeightPoints - yTop - height / 2d;
        bool clipsVerticalOverflow = ClipsVerticalOverflow(textBody);
        double textClipY = clipsVerticalOverflow
            ? document.SlideHeightPoints - yTop - insets.Top - textHeight
            : 0d;
        double textClipHeight = clipsVerticalOverflow
            ? textHeight
            : document.SlideHeightPoints;
        RgbColor? shapeFontColor = TryReadShapeFontColor(shape, theme, out RgbColor fontColor)
            ? fontColor
            : null;
        XElement? defaultParagraphProperties = MergeParagraphProperties(
            textBody.Element(DrawingNamespace + "lstStyle")?.Element(DrawingNamespace + "lvl1pPr"),
            inheritedTextBody?.Element(DrawingNamespace + "lstStyle")?.Element(DrawingNamespace + "lvl1pPr"),
            FindInheritedTextStyle(shape, placeholderSources, "lvl1pPr"),
            FindDefaultTextStyle(placeholderSources, "lvl1pPr"));
        double verticalOffset = ReadVerticalAnchor(textBody) switch
        {
            TextVerticalAnchor.Top when inheritedTextBody is not null => ReadVerticalAnchor(inheritedTextBody) switch
            {
                TextVerticalAnchor.Middle => Math.Max(0d, (textHeight - EstimateTextHeight(textBody, defaultParagraphProperties)) / 2d),
                TextVerticalAnchor.Bottom => Math.Max(0d, textHeight - EstimateTextHeight(textBody, defaultParagraphProperties)),
                _ => 0d
            },
            TextVerticalAnchor.Middle => Math.Max(0d, (textHeight - EstimateTextHeight(textBody, defaultParagraphProperties)) / 2d),
            TextVerticalAnchor.Bottom => Math.Max(0d, textHeight - EstimateTextHeight(textBody, defaultParagraphProperties)),
            _ => 0d
        };
        IReadOnlyList<PptxTextParagraphModel> paragraphs = BuildParagraphModels(
            shape,
            textBody,
            inheritedPlaceholders,
            placeholderSources,
            theme,
            slideNumber,
            fontScale,
            shapeFontColor);

        return new PptxTextFrameModel(
            shape,
            textBody,
            inheritedTextBody,
            theme,
            bounds.Value,
            insets,
            fontScale,
            textX,
            textWidth,
            textHeight,
            textClipY,
            textClipHeight,
            rotationCenterX,
            rotationCenterY,
            verticalOffset,
            shapeFontColor,
            paragraphs);
    }

    private static IReadOnlyList<PptxTextParagraphModel> BuildParagraphModels(
        XElement shape,
        XElement textBody,
        IReadOnlyList<XElement> inheritedPlaceholders,
        IReadOnlyList<XDocument> placeholderSources,
        PptxTheme theme,
        int slideNumber,
        double fontScale,
        RgbColor? shapeFontColor)
    {
        var paragraphs = new List<PptxTextParagraphModel>();
        foreach (XElement paragraph in textBody.Elements(DrawingNamespace + "p"))
        {
            XElement? paragraphProperties = paragraph.Element(DrawingNamespace + "pPr");
            int paragraphLevel = paragraphProperties?.Attribute("lvl") is { } levelAttribute
                ? int.Parse(levelAttribute.Value, CultureInfo.InvariantCulture)
                : 0;
            string levelName = $"lvl{Math.Clamp(paragraphLevel + 1, 1, 9).ToString(CultureInfo.InvariantCulture)}pPr";
            var paragraphPropertySources = new List<XElement?>
            {
                textBody.Element(DrawingNamespace + "lstStyle")?.Element(DrawingNamespace + levelName)
            };
            paragraphPropertySources.AddRange(inheritedPlaceholders
                .Select(placeholder => placeholder.Element(PresentationNamespace + "txBody")?.Element(DrawingNamespace + "lstStyle")?.Element(DrawingNamespace + levelName))
                .Reverse());
            paragraphPropertySources.Add(FindInheritedTextStyle(shape, placeholderSources, levelName));
            paragraphPropertySources.Add(FindDefaultTextStyle(placeholderSources, levelName));
            XElement? defaultParagraphProperties = MergeParagraphProperties(paragraphPropertySources.ToArray());
            ResolvedParagraphTextStyle paragraphStyle = ResolveParagraphTextStyle(paragraph, paragraphProperties, defaultParagraphProperties, fontScale);
            paragraphs.Add(new PptxTextParagraphModel(
                paragraph,
                paragraphProperties,
                defaultParagraphProperties,
                paragraphLevel,
                paragraphStyle,
                BuildRunModels(paragraph, paragraphStyle, shapeFontColor, theme, slideNumber, fontScale)));
        }

        return paragraphs;
    }

    private static IReadOnlyList<PptxTextRunModel> BuildRunModels(
        XElement paragraph,
        ResolvedParagraphTextStyle paragraphStyle,
        RgbColor? shapeFontColor,
        PptxTheme theme,
        int slideNumber,
        double fontScale)
    {
        var runs = new List<PptxTextRunModel>();
        foreach (XElement child in paragraph.Elements())
        {
            if (child.Name == DrawingNamespace + "br")
            {
                XElement? breakProperties = child.Element(DrawingNamespace + "rPr");
                runs.Add(new PptxTextRunModel(
                    PptxTextRunKind.Break,
                    child,
                    breakProperties,
                    "\n",
                    ResolveRunTextStyle(breakProperties, paragraphStyle.DefaultRunProperties, shapeFontColor, theme, fontScale)));
                continue;
            }

            if (!IsTextRunElement(child))
            {
                continue;
            }

            XElement? runProperties = child.Element(DrawingNamespace + "rPr");
            runs.Add(new PptxTextRunModel(
                child.Name == DrawingNamespace + "fld" ? PptxTextRunKind.Field : PptxTextRunKind.Text,
                child,
                runProperties,
                ReadTextElementText(child, slideNumber),
                ResolveRunTextStyle(runProperties, paragraphStyle.DefaultRunProperties, shapeFontColor, theme, fontScale)));
        }

        return runs;
    }
}
