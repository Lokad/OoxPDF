using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Diagnostics;
using System.Globalization;
using System.Text;
using System.Xml.Linq;

namespace Lokad.OoxPdf.Pptx;

internal sealed class PptxRenderer
{
    private static readonly XNamespace PresentationNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace ChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private const string SlideLayoutRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
    private const string SlideMasterRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
    private static readonly XNamespace RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public IReadOnlyList<PdfPage> RenderBlankPages(PptxDocument document)
    {
        return document.Slides
            .Select(_ => new PdfPage(document.SlideWidthPoints, document.SlideHeightPoints))
            .ToArray();
    }

    public IReadOnlyList<PdfPage> RenderPages(PptxDocument document, OoxPackage package, Action<OoxPdfDiagnostic>? diagnosticSink = null)
    {
        var pages = new List<PdfPage>(document.Slides.Count);
        for (int slideIndex = 0; slideIndex < document.Slides.Count; slideIndex++)
        {
            PptxSlide slide = document.Slides[slideIndex];
            OoxPart? slidePart = package.GetPart(slide.PartName);
            if (slidePart is null)
            {
                pages.Add(new PdfPage(document.SlideWidthPoints, document.SlideHeightPoints));
                continue;
            }

            using Stream stream = slidePart.OpenRead();
            XDocument slideXml = SafeXml.Load(stream);
            EmitUnsupportedFeatureDiagnostics(slideXml, slide.PartName, slideIndex + 1, diagnosticSink);
            IReadOnlyList<XDocument> inheritedXml = LoadInheritedSlideXml(package, slide.PartName);
            var graphics = new PdfGraphicsBuilder();
            PptxTheme theme = PptxTheme.Load(package, document.PresentationPartName);

            foreach (XDocument inherited in inheritedXml)
            {
                RenderBackground(inherited, document, graphics, theme);
                RenderShapes(inherited, document, graphics, theme, renderPlaceholders: false);
            }

            RenderBackground(slideXml, document, graphics, theme);
            RenderShapes(slideXml, document, graphics, theme, renderPlaceholders: true);
            IReadOnlyList<TextRun> tableTextRuns = inheritedXml
                .Append(slideXml)
                .SelectMany(xml => RenderTables(xml, document, graphics, theme))
                .ToArray();
            IReadOnlyList<PdfImageResource> images = RenderPictures(package, slide.PartName, slideXml, document, graphics);
            IReadOnlyList<TextRun> textRuns = inheritedXml
                .SelectMany(xml => ReadTextRuns(xml, document, theme, includePlaceholders: false))
                .Concat(ReadTextRuns(slideXml, document, theme, includePlaceholders: true))
                .Concat(tableTextRuns)
                .ToArray();
            IReadOnlyList<PdfFontResource> fonts = RenderTextRuns(textRuns, graphics);
            pages.Add(new PdfPage(document.SlideWidthPoints, document.SlideHeightPoints, graphics.ToString(), fonts, images));
        }

        return pages;
    }

    private static void EmitUnsupportedFeatureDiagnostics(XDocument slideXml, string partName, int slideIndex, Action<OoxPdfDiagnostic>? diagnosticSink)
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

        if (slideXml.Descendants(ChartNamespace + "chart").Any() ||
            HasGraphicDataUri(slideXml, "drawingml/2006/chart"))
        {
            Emit("PPTX_UNSUPPORTED_CHART", "chart");
        }

        if (HasGraphicDataUri(slideXml, "drawingml/2006/diagram"))
        {
            Emit("PPTX_UNSUPPORTED_SMARTART", "SmartArt");
        }
    }

    private static bool HasGraphicDataUri(XDocument slideXml, string marker)
    {
        return slideXml
            .Descendants(DrawingNamespace + "graphicData")
            .Select(element => (string?)element.Attribute("uri"))
            .Any(uri => uri?.Contains(marker, StringComparison.OrdinalIgnoreCase) == true);
    }

    private static IReadOnlyList<XDocument> LoadInheritedSlideXml(OoxPackage package, string slidePartName)
    {
        var documents = new List<XDocument>();
        OoxPart? layoutPart = GetRelatedPart(package, slidePartName, SlideLayoutRelationshipType);
        OoxPart? masterPart = layoutPart is null ? null : GetRelatedPart(package, layoutPart.Name, SlideMasterRelationshipType);

        if (masterPart is not null)
        {
            using Stream masterStream = masterPart.OpenRead();
            documents.Add(SafeXml.Load(masterStream));
        }

        if (layoutPart is not null)
        {
            using Stream layoutStream = layoutPart.OpenRead();
            documents.Add(SafeXml.Load(layoutStream));
        }

        return documents;
    }

    private static OoxPart? GetRelatedPart(OoxPackage package, string sourcePartName, string relationshipType)
    {
        OoxRelationship? relationship = package.GetRelationships(sourcePartName)
            .FirstOrDefault(r => !r.IsExternal && r.Type == relationshipType && r.ResolvedTarget is not null);
        return relationship?.ResolvedTarget is null ? null : package.GetPart(relationship.ResolvedTarget);
    }

    private static void RenderBackground(XDocument slideXml, PptxDocument document, PdfGraphicsBuilder graphics, PptxTheme theme)
    {
        XElement? background = slideXml.Root?
            .Element(PresentationNamespace + "cSld")?
            .Element(PresentationNamespace + "bg")?
            .Element(PresentationNamespace + "bgPr");
        if (TryReadSolidColor(background, theme, out RgbColor color))
        {
            graphics.SetFillRgb(color.Red, color.Green, color.Blue);
            graphics.FillRectangle(0, 0, document.SlideWidthPoints, document.SlideHeightPoints);
        }
    }

    private static void RenderShapes(XDocument slideXml, PptxDocument document, PdfGraphicsBuilder graphics, PptxTheme theme, bool renderPlaceholders)
    {
        foreach (XElement shapeTree in slideXml.Descendants(PresentationNamespace + "spTree"))
        {
            RenderShapeContainer(shapeTree, document, graphics, theme, GroupTransform.Identity, renderPlaceholders);
        }
    }

    private static void RenderShapeContainer(XElement container, PptxDocument document, PdfGraphicsBuilder graphics, PptxTheme theme, GroupTransform transform, bool renderPlaceholders)
    {
        foreach (XElement shape in container.Elements(PresentationNamespace + "sp"))
        {
            if (renderPlaceholders || !IsPlaceholder(shape))
            {
                RenderShape(shape, document, graphics, theme, transform);
            }
        }

        foreach (XElement group in container.Elements(PresentationNamespace + "grpSp"))
        {
            GroupTransform childTransform = transform.Combine(ReadGroupTransform(group));
            RenderShapeContainer(group, document, graphics, theme, childTransform, renderPlaceholders);
        }
    }

    private static void RenderShape(XElement shape, PptxDocument document, PdfGraphicsBuilder graphics, PptxTheme theme, GroupTransform groupTransform)
    {
        XElement? shapeProperties = shape.Element(PresentationNamespace + "spPr");
        if (shapeProperties is null)
        {
            return;
        }

        ShapeBounds? rawBounds = ReadBounds(shapeProperties);
        if (rawBounds is null)
        {
            return;
        }

        ShapeBounds bounds = groupTransform.Apply(rawBounds.Value);
        string preset = (string?)shapeProperties
            .Element(DrawingNamespace + "prstGeom")
            ?.Attribute("prst") ?? "rect";

        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        bool transformed = bounds.RotationDegrees != 0d || bounds.FlipHorizontal || bounds.FlipVertical;

        bool hasFill = TryReadSolidColor(shapeProperties, theme, out RgbColor fill);
        bool hasStroke = TryReadLine(shapeProperties, theme, out RgbColor stroke, out double lineWidth);

        if (transformed)
        {
            graphics.SaveState();
            ApplyShapeTransform(graphics, x, y, width, height, bounds);
        }

        if (preset == "line")
        {
            if (hasStroke)
            {
                graphics.SetStrokeRgb(stroke.Red, stroke.Green, stroke.Blue);
                graphics.SetLineWidth(lineWidth);
                graphics.StrokeLine(x, document.SlideHeightPoints - yTop, x + width, document.SlideHeightPoints - yTop - height);
            }

            if (transformed)
            {
                graphics.RestoreState();
            }

            return;
        }

        if (hasFill)
        {
            graphics.SetFillRgb(fill.Red, fill.Green, fill.Blue);
            if (preset == "ellipse")
            {
                graphics.FillEllipse(x, y, width, height);
            }
            else
            {
                graphics.FillRectangle(x, y, width, height);
            }
        }

        if (hasStroke)
        {
            graphics.SetStrokeRgb(stroke.Red, stroke.Green, stroke.Blue);
            graphics.SetLineWidth(lineWidth);
            if (preset == "ellipse")
            {
                graphics.StrokeEllipse(x, y, width, height);
            }
            else
            {
                graphics.StrokeRectangle(x, y, width, height);
            }
        }

        if (transformed)
        {
            graphics.RestoreState();
        }
    }

    private static IReadOnlyList<TextRun> RenderTables(XDocument slideXml, PptxDocument document, PdfGraphicsBuilder graphics, PptxTheme theme)
    {
        var textRuns = new List<TextRun>();
        foreach (XElement frame in slideXml.Descendants(PresentationNamespace + "graphicFrame"))
        {
            ShapeBounds? bounds = ReadGraphicFrameBounds(frame);
            XElement? table = frame
                .Element(DrawingNamespace + "graphic")
                ?.Element(DrawingNamespace + "graphicData")
                ?.Element(DrawingNamespace + "tbl");
            if (bounds is null || table is null)
            {
                continue;
            }

            IReadOnlyList<double> rawColumnWidths = table
                .Element(DrawingNamespace + "tblGrid")
                ?.Elements(DrawingNamespace + "gridCol")
                .Select(column => Math.Max(1d, ParseOptionalLongAttribute(column, "w", 1)))
                .ToArray() ?? [];
            IReadOnlyList<XElement> rows = table.Elements(DrawingNamespace + "tr").ToArray();
            if (rawColumnWidths.Count == 0 || rows.Count == 0)
            {
                continue;
            }

            double frameX = OoxUnits.EmuToPoints(bounds.Value.X);
            double frameYTop = OoxUnits.EmuToPoints(bounds.Value.Y);
            double frameWidth = OoxUnits.EmuToPoints(bounds.Value.Width);
            double frameHeight = OoxUnits.EmuToPoints(bounds.Value.Height);
            double frameTop = document.SlideHeightPoints - frameYTop;
            double columnScale = frameWidth / rawColumnWidths.Sum();

            IReadOnlyList<double> rawRowHeights = rows
                .Select(row => Math.Max(1d, ParseOptionalLongAttribute(row, "h", 1)))
                .ToArray();
            double rowScale = frameHeight / rawRowHeights.Sum();

            double yTop = frameTop;
            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                double rowHeight = rawRowHeights[rowIndex] * rowScale;
                double cellX = frameX;
                double cellY = yTop - rowHeight;
                IReadOnlyList<XElement> cells = rows[rowIndex].Elements(DrawingNamespace + "tc").ToArray();

                for (int columnIndex = 0; columnIndex < Math.Min(cells.Count, rawColumnWidths.Count); columnIndex++)
                {
                    double columnWidth = rawColumnWidths[columnIndex] * columnScale;
                    XElement cell = cells[columnIndex];
                    XElement? cellProperties = cell.Element(DrawingNamespace + "tcPr");

                    if (TryReadSolidColor(cellProperties, theme, out RgbColor fill))
                    {
                        graphics.SetFillRgb(fill.Red, fill.Green, fill.Blue);
                        graphics.FillRectangle(cellX, cellY, columnWidth, rowHeight);
                    }

                    graphics.SetStrokeRgb(0, 0, 0);
                    graphics.SetLineWidth(0.75d);
                    graphics.StrokeRectangle(cellX, cellY, columnWidth, rowHeight);
                    AddTableCellTextRuns(cell, cellX, cellY, columnWidth, rowHeight, theme, textRuns);
                    cellX += columnWidth;
                }

                yTop -= rowHeight;
            }
        }

        return textRuns;
    }

    private static ShapeBounds? ReadGraphicFrameBounds(XElement frame)
    {
        XElement? transform = frame.Element(PresentationNamespace + "xfrm");
        return transform is null ? null : ReadBoundsFromTransform(transform);
    }

    private static void AddTableCellTextRuns(XElement cell, double x, double y, double width, double height, PptxTheme theme, List<TextRun> runs)
    {
        XElement? textBody = cell.Element(DrawingNamespace + "txBody");
        if (textBody is null)
        {
            return;
        }

        double cursorY = y + height - 14d;
        foreach (XElement paragraph in textBody.Elements(DrawingNamespace + "p"))
        {
            TextAlignment alignment = ReadAlignment(paragraph);
            double cursorX = x + 4d;
            double maxFontSize = 12d;
            foreach (XElement run in paragraph.Elements(DrawingNamespace + "r"))
            {
                string text = (string?)run.Element(DrawingNamespace + "t") ?? string.Empty;
                if (text.Length == 0)
                {
                    continue;
                }

                XElement? runProperties = run.Element(DrawingNamespace + "rPr");
                double fontSize = runProperties?.Attribute("sz") is { } size
                    ? int.Parse(size.Value, CultureInfo.InvariantCulture) / 100d
                    : 12d;
                maxFontSize = Math.Max(maxFontSize, fontSize);
                RgbColor color = TryReadSolidColor(runProperties, theme, out RgbColor runColor)
                    ? runColor
                    : new RgbColor(0, 0, 0);
                string? typeface = theme.ResolveTypeface((string?)runProperties?
                    .Element(DrawingNamespace + "latin")
                    ?.Attribute("typeface"));
                bool bold = ParseOptionalBoolAttribute(runProperties, "b");
                bool italic = ParseOptionalBoolAttribute(runProperties, "i");
                bool underline = ((string?)runProperties?.Attribute("u")) is { } underlineValue
                    && !underlineValue.Equals("none", StringComparison.OrdinalIgnoreCase);
                runs.Add(new TextRun(text, cursorX, cursorY, Math.Max(1d, width - 8d), Math.Max(1d, height - 8d), fontSize, color, bold, italic, underline, alignment, typeface));
                cursorX += text.Length * fontSize * 0.55d;
            }

            cursorY -= maxFontSize * 1.2d;
        }
    }

    private static ShapeBounds? ReadBounds(XElement shapeProperties)
    {
        XElement? transform = shapeProperties.Element(DrawingNamespace + "xfrm");
        return transform is null ? null : ReadBoundsFromTransform(transform);
    }

    private static ShapeBounds? ReadBoundsFromTransform(XElement transform)
    {
        XElement? offset = transform.Element(DrawingNamespace + "off");
        XElement? extents = transform.Element(DrawingNamespace + "ext");
        if (offset is null || extents is null)
        {
            return null;
        }

        double rotationDegrees = transform.Attribute("rot") is { } rotationAttribute
            ? long.Parse(rotationAttribute.Value, CultureInfo.InvariantCulture) / 60000d
            : 0d;
        bool flipHorizontal = ParseBoolAttribute(transform, "flipH");
        bool flipVertical = ParseBoolAttribute(transform, "flipV");

        return new ShapeBounds(
            ParseLongAttribute(offset, "x"),
            ParseLongAttribute(offset, "y"),
            ParseLongAttribute(extents, "cx"),
            ParseLongAttribute(extents, "cy"),
            rotationDegrees,
            flipHorizontal,
            flipVertical);
    }

    private static GroupTransform ReadGroupTransform(XElement group)
    {
        XElement? transform = group
            .Element(PresentationNamespace + "grpSpPr")
            ?.Element(DrawingNamespace + "xfrm");
        XElement? offset = transform?.Element(DrawingNamespace + "off");
        XElement? extents = transform?.Element(DrawingNamespace + "ext");
        XElement? childOffset = transform?.Element(DrawingNamespace + "chOff");
        XElement? childExtents = transform?.Element(DrawingNamespace + "chExt");
        if (offset is null || extents is null || childOffset is null || childExtents is null)
        {
            return GroupTransform.Identity;
        }

        long chWidth = Math.Max(1, ParseLongAttribute(childExtents, "cx"));
        long chHeight = Math.Max(1, ParseLongAttribute(childExtents, "cy"));
        return new GroupTransform(
            ParseLongAttribute(offset, "x"),
            ParseLongAttribute(offset, "y"),
            ParseLongAttribute(childOffset, "x"),
            ParseLongAttribute(childOffset, "y"),
            ParseLongAttribute(extents, "cx") / (double)chWidth,
            ParseLongAttribute(extents, "cy") / (double)chHeight);
    }

    private static void ApplyShapeTransform(PdfGraphicsBuilder graphics, double x, double y, double width, double height, ShapeBounds bounds)
    {
        double radians = -bounds.RotationDegrees * Math.PI / 180d;
        double sx = bounds.FlipHorizontal ? -1d : 1d;
        double sy = bounds.FlipVertical ? -1d : 1d;
        double cos = Math.Cos(radians);
        double sin = Math.Sin(radians);
        double centerX = x + width / 2d;
        double centerY = y + height / 2d;

        double a = cos * sx;
        double b = sin * sx;
        double c = -sin * sy;
        double d = cos * sy;
        double e = centerX - a * centerX - c * centerY;
        double f = centerY - b * centerX - d * centerY;
        graphics.Transform(a, b, c, d, e, f);
    }

    private static bool TryReadSolidColor(XElement? element, PptxTheme theme, out RgbColor color)
    {
        XElement? solidFill = element?.Element(DrawingNamespace + "solidFill");
        string? hex = (string?)solidFill?.Element(DrawingNamespace + "srgbClr")?.Attribute("val");
        if (RgbColor.TryParse(hex, out color))
        {
            return true;
        }

        string? schemeColor = (string?)solidFill?.Element(DrawingNamespace + "schemeClr")?.Attribute("val");
        return schemeColor is not null && theme.TryResolveColor(schemeColor, out color);
    }

    private static bool TryReadLine(XElement shapeProperties, PptxTheme theme, out RgbColor color, out double lineWidth)
    {
        XElement? line = shapeProperties.Element(DrawingNamespace + "ln");
        lineWidth = line?.Attribute("w") is { } widthAttribute
            ? OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture))
            : 1d;
        return TryReadSolidColor(line, theme, out color);
    }

    private static IReadOnlyList<TextRun> ReadTextRuns(XDocument slideXml, PptxDocument document, PptxTheme theme, bool includePlaceholders)
    {
        var runs = new List<TextRun>();
        foreach (XElement shape in slideXml.Descendants(PresentationNamespace + "sp"))
        {
            if (!includePlaceholders && IsPlaceholder(shape))
            {
                continue;
            }

            XElement? shapeProperties = shape.Element(PresentationNamespace + "spPr");
            XElement? textBody = shape.Element(PresentationNamespace + "txBody");
            ShapeBounds? bounds = shapeProperties is null ? null : ReadBounds(shapeProperties);
            if (bounds is null || textBody is null)
            {
                continue;
            }

            double x = OoxUnits.EmuToPoints(bounds.Value.X);
            double yTop = OoxUnits.EmuToPoints(bounds.Value.Y);
            double width = OoxUnits.EmuToPoints(bounds.Value.Width);
            double height = OoxUnits.EmuToPoints(bounds.Value.Height);
            double cursorY = document.SlideHeightPoints - yTop - 18d;

            foreach (XElement paragraph in textBody.Elements(DrawingNamespace + "p"))
            {
                TextAlignment alignment = ReadAlignment(paragraph);
                double cursorX = x + 4d;
                double maxFontSize = 18d;
                foreach (XElement run in paragraph.Elements(DrawingNamespace + "r"))
                {
                    string text = (string?)run.Element(DrawingNamespace + "t") ?? string.Empty;
                    if (text.Length == 0)
                    {
                        continue;
                    }

                    XElement? runProperties = run.Element(DrawingNamespace + "rPr");
                    double fontSize = runProperties?.Attribute("sz") is { } size
                        ? int.Parse(size.Value, CultureInfo.InvariantCulture) / 100d
                        : 18d;
                    maxFontSize = Math.Max(maxFontSize, fontSize);
                    RgbColor color = TryReadSolidColor(runProperties, theme, out RgbColor runColor)
                        ? runColor
                        : new RgbColor(0, 0, 0);
                    string? typeface = theme.ResolveTypeface((string?)runProperties?
                        .Element(DrawingNamespace + "latin")
                        ?.Attribute("typeface"));
                    bool bold = ParseOptionalBoolAttribute(runProperties, "b");
                    bool italic = ParseOptionalBoolAttribute(runProperties, "i");
                    bool underline = ((string?)runProperties?.Attribute("u")) is { } underlineValue
                        && !underlineValue.Equals("none", StringComparison.OrdinalIgnoreCase);
                    runs.Add(new TextRun(text, cursorX, cursorY, width, height, fontSize, color, bold, italic, underline, alignment, typeface));
                    cursorX += text.Length * fontSize * 0.55d;
                }

                cursorY -= maxFontSize * 1.2d;
            }
        }

        return runs;
    }

    private static bool IsPlaceholder(XElement shape)
    {
        return shape
            .Element(PresentationNamespace + "nvSpPr")
            ?.Element(PresentationNamespace + "nvPr")
            ?.Element(PresentationNamespace + "ph") is not null;
    }

    private static IReadOnlyList<PdfImageResource> RenderPictures(OoxPackage package, string slidePartName, XDocument slideXml, PptxDocument document, PdfGraphicsBuilder graphics)
    {
        var relationships = package.GetRelationships(slidePartName)
            .Where(r => !r.IsExternal && r.ResolvedTarget is not null)
            .ToDictionary(r => r.Id, StringComparer.Ordinal);
        var images = new List<PdfImageResource>();
        int index = 1;
        foreach (XElement picture in slideXml.Descendants(PresentationNamespace + "pic"))
        {
            string? relationshipId = (string?)picture
                .Element(PresentationNamespace + "blipFill")
                ?.Element(DrawingNamespace + "blip")
                ?.Attribute(RelationshipsNamespace + "embed");
            XElement? shapeProperties = picture.Element(PresentationNamespace + "spPr");
            ShapeBounds? bounds = shapeProperties is null ? null : ReadBounds(shapeProperties);
            if (relationshipId is null || bounds is null || !relationships.TryGetValue(relationshipId, out OoxRelationship? relationship) || relationship.ResolvedTarget is null)
            {
                continue;
            }

            OoxPart? imagePart = package.GetPart(relationship.ResolvedTarget);
            if (imagePart is null)
            {
                continue;
            }

            PdfImageXObject? image = CreateImage(imagePart);
            if (image is null)
            {
                continue;
            }

            string name = "Im" + index++;
            double x = OoxUnits.EmuToPoints(bounds.Value.X);
            double yTop = OoxUnits.EmuToPoints(bounds.Value.Y);
            double width = OoxUnits.EmuToPoints(bounds.Value.Width);
            double height = OoxUnits.EmuToPoints(bounds.Value.Height);
            double y = document.SlideHeightPoints - yTop - height;
            CropRect crop = ReadCrop(picture);
            if (crop.IsEmpty)
            {
                graphics.DrawImage(name, x, y, width, height);
            }
            else
            {
                graphics.DrawImageCropped(name, x, y, width, height, crop.Left, crop.Top, crop.Right, crop.Bottom);
            }

            images.Add(new PdfImageResource(name, image));
        }

        return images;
    }

    private static PdfImageXObject? CreateImage(OoxPart imagePart)
    {
        byte[] bytes = imagePart.Bytes;
        if (imagePart.ContentType.Equals("image/jpeg", StringComparison.OrdinalIgnoreCase) ||
            imagePart.ContentType.Equals("image/jpg", StringComparison.OrdinalIgnoreCase))
        {
            JpegInfo info = JpegInfo.Read(bytes);
            return PdfImageXObject.Jpeg(info.Width, info.Height, bytes);
        }

        if (imagePart.ContentType.Equals("image/png", StringComparison.OrdinalIgnoreCase))
        {
            PngImage png = PngImage.Read(bytes);
            return PdfImageXObject.RgbPng(png.Width, png.Height, png.Rgb, png.Alpha);
        }

        return null;
    }

    private static CropRect ReadCrop(XElement picture)
    {
        XElement? sourceRectangle = picture
            .Element(PresentationNamespace + "blipFill")
            ?.Element(DrawingNamespace + "srcRect");
        if (sourceRectangle is null)
        {
            return default;
        }

        return new CropRect(
            ParsePercentage(sourceRectangle, "l"),
            ParsePercentage(sourceRectangle, "t"),
            ParsePercentage(sourceRectangle, "r"),
            ParsePercentage(sourceRectangle, "b"));
    }

    private static double ParsePercentage(XElement element, string attribute)
    {
        return element.Attribute(attribute) is { } value
            ? Math.Clamp(int.Parse(value.Value, CultureInfo.InvariantCulture) / 100000d, 0d, 0.999d)
            : 0d;
    }

    private static IReadOnlyList<PdfFontResource> RenderTextRuns(IReadOnlyList<TextRun> textRuns, PdfGraphicsBuilder graphics)
    {
        if (textRuns.Count == 0)
        {
            return [];
        }

        string familyName = textRuns.Select(r => r.FontFamily).FirstOrDefault(f => !string.IsNullOrWhiteSpace(f)) ?? "Arial";
        FontResolution resolution = new WindowsFontResolver().Resolve(new FontRequest(familyName));
        if (resolution.FontFilePath is null || !File.Exists(resolution.FontFilePath))
        {
            return [];
        }

        OpenTypeFont font = OpenTypeFont.Load(resolution.FontFilePath);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, textRuns.SelectMany(r => r.Text.EnumerateRunes().Select(rune => rune.Value)));
        var resource = new PdfFontResource("F1", embedded);

        foreach (TextRun run in textRuns)
        {
            DrawWrappedRun(graphics, embedded, run);
        }

        return [resource];
    }

    private static void DrawWrappedRun(PdfGraphicsBuilder graphics, PdfEmbeddedFont embedded, TextRun run)
    {
        double cursorY = run.Y;
        double lineHeight = run.FontSize * 1.2d;
        foreach (string line in WrapWords(run.Text, run.Width, run.FontSize, embedded))
        {
            if (cursorY < run.Y - run.Height)
            {
                break;
            }

            string glyphHex = embedded.EncodeGlyphHex(line);
            if (glyphHex.Length != 0)
            {
                double lineWidth = embedded.MeasureTextPoints(line, run.FontSize);
                double x = run.Alignment switch
                {
                    TextAlignment.Center => run.X + Math.Max(0, run.Width - lineWidth) / 2d,
                    TextAlignment.Right => run.X + Math.Max(0, run.Width - lineWidth),
                    _ => run.X
                };

                graphics.DrawGlyphText("F1", run.FontSize, x, cursorY, run.Color.Red, run.Color.Green, run.Color.Blue, glyphHex, run.Italic);
                if (run.Bold)
                {
                    graphics.DrawGlyphText("F1", run.FontSize, x + 0.35d, cursorY, run.Color.Red, run.Color.Green, run.Color.Blue, glyphHex, run.Italic);
                }

                if (run.Underline)
                {
                    graphics.SetStrokeRgb(run.Color.Red, run.Color.Green, run.Color.Blue);
                    graphics.SetLineWidth(Math.Max(0.5d, run.FontSize / 18d));
                    graphics.StrokeLine(x, cursorY - run.FontSize * 0.12d, x + lineWidth, cursorY - run.FontSize * 0.12d);
                }
            }

            cursorY -= lineHeight;
        }
    }

    private static IEnumerable<string> WrapWords(string text, double maxWidth, double fontSize, PdfEmbeddedFont embedded)
    {
        string[] words = text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (words.Length == 0)
        {
            yield break;
        }

        var line = new StringBuilder();
        foreach (string word in words)
        {
            string candidate = line.Length == 0 ? word : line + " " + word;
            if (line.Length > 0 && embedded.MeasureTextPoints(candidate, fontSize) > maxWidth)
            {
                yield return line.ToString();
                line.Clear();
                line.Append(word);
            }
            else
            {
                line.Clear();
                line.Append(candidate);
            }
        }

        if (line.Length > 0)
        {
            yield return line.ToString();
        }
    }

    private static long ParseLongAttribute(XElement element, string name)
    {
        string value = (string?)element.Attribute(name)
            ?? throw new InvalidDataException($"Missing required PPTX shape attribute '{name}'.");
        return long.Parse(value, CultureInfo.InvariantCulture);
    }

    private static long ParseOptionalLongAttribute(XElement element, string name, long defaultValue)
    {
        return element.Attribute(name) is { } value
            ? long.Parse(value.Value, CultureInfo.InvariantCulture)
            : defaultValue;
    }

    private static bool ParseBoolAttribute(XElement element, string name)
    {
        string? value = (string?)element.Attribute(name);
        return value is not null && (value == "1" || value.Equals("true", StringComparison.OrdinalIgnoreCase));
    }

    private static bool ParseOptionalBoolAttribute(XElement? element, string name)
    {
        return element is not null && ParseBoolAttribute(element, name);
    }

    private static TextAlignment ReadAlignment(XElement paragraph)
    {
        string? value = (string?)paragraph.Element(DrawingNamespace + "pPr")?.Attribute("algn");
        return value switch
        {
            "ctr" => TextAlignment.Center,
            "r" => TextAlignment.Right,
            _ => TextAlignment.Left
        };
    }

    private readonly record struct ShapeBounds(
        long X,
        long Y,
        long Width,
        long Height,
        double RotationDegrees,
        bool FlipHorizontal,
        bool FlipVertical);

    private readonly record struct GroupTransform(long OffsetX, long OffsetY, long ChildOffsetX, long ChildOffsetY, double ScaleX, double ScaleY)
    {
        public static GroupTransform Identity { get; } = new(0, 0, 0, 0, 1d, 1d);

        public ShapeBounds Apply(ShapeBounds bounds)
        {
            return new ShapeBounds(
                OffsetX + (long)Math.Round((bounds.X - ChildOffsetX) * ScaleX),
                OffsetY + (long)Math.Round((bounds.Y - ChildOffsetY) * ScaleY),
                (long)Math.Round(bounds.Width * ScaleX),
                (long)Math.Round(bounds.Height * ScaleY),
                bounds.RotationDegrees,
                bounds.FlipHorizontal,
                bounds.FlipVertical);
        }

        public GroupTransform Combine(GroupTransform child)
        {
            return new GroupTransform(
                OffsetX + (long)Math.Round((child.OffsetX - ChildOffsetX) * ScaleX),
                OffsetY + (long)Math.Round((child.OffsetY - ChildOffsetY) * ScaleY),
                child.ChildOffsetX,
                child.ChildOffsetY,
                ScaleX * child.ScaleX,
                ScaleY * child.ScaleY);
        }
    }

    private readonly record struct TextRun(
        string Text,
        double X,
        double Y,
        double Width,
        double Height,
        double FontSize,
        RgbColor Color,
        bool Bold,
        bool Italic,
        bool Underline,
        TextAlignment Alignment,
        string? FontFamily);

    private enum TextAlignment
    {
        Left,
        Center,
        Right
    }

    private readonly record struct CropRect(double Left, double Top, double Right, double Bottom)
    {
        public bool IsEmpty => Left == 0d && Top == 0d && Right == 0d && Bottom == 0d;
    }
}
