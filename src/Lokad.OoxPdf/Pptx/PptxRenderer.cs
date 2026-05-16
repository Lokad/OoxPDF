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
            if (CanRenderSlideInOrder(slideXml))
            {
                var relationships = package.GetRelationships(slide.PartName)
                    .Where(r => !r.IsExternal && r.ResolvedTarget is not null)
                    .ToDictionary(r => r.Id, StringComparer.Ordinal);
                var orderedImages = new List<PdfImageResource>();
                int imageIndex = 1;
                IReadOnlyList<TextRun> inheritedTextRuns = inheritedXml
                    .SelectMany(xml => ReadTextRuns(xml, document, theme, includePlaceholders: false, placeholderSources: []))
                    .ToArray();
                IReadOnlyList<TextRun> slideTextRuns = ReadTextRuns(slideXml, document, theme, includePlaceholders: true, inheritedXml);
                IReadOnlyList<TextRun> slideTableTextRuns = RenderTables(slideXml, document, new PdfGraphicsBuilder(), theme);
                RenderedFonts renderedFonts = CreateRenderedFonts(inheritedTextRuns.Concat(slideTextRuns).Concat(slideTableTextRuns).ToArray());
                DrawTextRunsWithFonts(inheritedTextRuns, graphics, renderedFonts.Fonts);
                foreach (XElement shapeTree in slideXml.Descendants(PresentationNamespace + "spTree"))
                {
                    RenderOrderedShapeTextContainer(shapeTree, relationships, package, document, graphics, diagnosticSink, slideIndex + 1, theme, renderedFonts.Fonts, orderedImages, ref imageIndex, GroupTransform.Identity, renderPlaceholders: true, inheritedXml);
                }

                pages.Add(new PdfPage(document.SlideWidthPoints, document.SlideHeightPoints, graphics.ToString(), renderedFonts.Resources, orderedImages, graphics.ExtGStates.ToArray()));
                continue;
            }

            IReadOnlyList<PdfImageResource> images = RenderPictures(package, slide.PartName, slideXml, document, graphics, diagnosticSink, slideIndex + 1);
            RenderShapes(slideXml, document, graphics, theme, renderPlaceholders: true);
            IReadOnlyList<TextRun> tableTextRuns = inheritedXml
                .Append(slideXml)
                .SelectMany(xml => RenderTables(xml, document, graphics, theme))
                .ToArray();
            RenderCharts(package, slide.PartName, slideXml, document, graphics, diagnosticSink, slideIndex + 1);
            IReadOnlyList<TextRun> textRuns = inheritedXml
                .SelectMany(xml => ReadTextRuns(xml, document, theme, includePlaceholders: false, placeholderSources: []))
                .Concat(ReadTextRuns(slideXml, document, theme, includePlaceholders: true, inheritedXml))
                .Concat(tableTextRuns)
                .ToArray();
            IReadOnlyList<PdfFontResource> fonts = RenderTextRuns(textRuns, graphics);
            pages.Add(new PdfPage(document.SlideWidthPoints, document.SlideHeightPoints, graphics.ToString(), fonts, images, graphics.ExtGStates.ToArray()));
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

        if (HasGraphicDataUri(slideXml, "drawingml/2006/diagram"))
        {
            Emit("PPTX_UNSUPPORTED_SMARTART", "SmartArt");
        }

        if (slideXml.Descendants(DrawingNamespace + "gradFill").Any())
        {
            Emit("PPTX_UNSUPPORTED_GRADIENT_FILL", "gradient fill");
        }

        if (slideXml.Descendants(DrawingNamespace + "pattFill").Any())
        {
            Emit("PPTX_UNSUPPORTED_PATTERN_FILL", "pattern fill");
        }

        if (slideXml.Descendants(DrawingNamespace + "bodyPr").Any(bodyProperties =>
                bodyProperties.Element(DrawingNamespace + "normAutofit") is not null ||
                bodyProperties.Element(DrawingNamespace + "spAutoFit") is not null))
        {
            Emit("PPTX_UNSUPPORTED_TEXT_AUTOFIT", "text autofit");
        }

        if (slideXml.Descendants(DrawingNamespace + "bodyPr").Any(HasUnsupportedTextColumns))
        {
            Emit("PPTX_UNSUPPORTED_TEXT_COLUMNS", "multi-column text");
        }

        if (slideXml.Descendants(DrawingNamespace + "bodyPr").Any(HasUnsupportedTextOrientation))
        {
            Emit("PPTX_UNSUPPORTED_TEXT_ORIENTATION", "vertical text");
        }

        if (slideXml.Descendants(PresentationNamespace + "spPr").Any(shapeProperties =>
                shapeProperties.Element(DrawingNamespace + "blipFill") is not null))
        {
            Emit("PPTX_UNSUPPORTED_PICTURE_FILL", "picture fill");
        }

        if (slideXml.Descendants(DrawingNamespace + "alpha").Any(IsUnsupportedAlpha))
        {
            Emit("PPTX_UNSUPPORTED_TRANSPARENCY", "transparency");
        }

        if (slideXml.Descendants(DrawingNamespace + "effectLst").Any(effectList => effectList.Elements().Any()) ||
            slideXml.Descendants(DrawingNamespace + "effectDag").Any())
        {
            Emit("PPTX_UNSUPPORTED_EFFECT", "effect");
        }

        if (slideXml.Descendants(DrawingNamespace + "custGeom").Any())
        {
            Emit("PPTX_UNSUPPORTED_CUSTOM_GEOMETRY", "custom geometry");
        }

        if (slideXml.Descendants(DrawingNamespace + "prstGeom").Any(geometry =>
                IsUnsupportedCalloutPreset((string?)geometry.Attribute("prst"))))
        {
            Emit("PPTX_UNSUPPORTED_CALLOUT", "callout shape");
        }
    }

    private static bool IsUnsupportedCalloutPreset(string? preset)
    {
        return preset?.Contains("Callout", StringComparison.OrdinalIgnoreCase) == true &&
            !string.Equals(preset, "wedgeRectCallout", StringComparison.Ordinal);
    }

    private static bool HasUnsupportedTextColumns(XElement bodyProperties)
    {
        return bodyProperties.Attribute("numCol") is { } columns &&
            int.TryParse(columns.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int count) &&
            count > 1;
    }

    private static bool HasUnsupportedTextOrientation(XElement bodyProperties)
    {
        string? orientation = (string?)bodyProperties.Attribute("vert");
        return !string.IsNullOrEmpty(orientation) &&
            !orientation.Equals("horz", StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasGraphicDataUri(XDocument slideXml, string marker)
    {
        return slideXml
            .Descendants(DrawingNamespace + "graphicData")
            .Select(element => (string?)element.Attribute("uri"))
            .Any(uri => uri?.Contains(marker, StringComparison.OrdinalIgnoreCase) == true);
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
        bool supportedShapeFill = fill?.Name == DrawingNamespace + "solidFill" &&
            owner?.Name == PresentationNamespace + "spPr";
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
        return !supportedShapeFill && !supportedShapeLine && !supportedTextFill && !supportedTableCellFill && !supportedTableBorder;
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

    private static bool CanRenderSlideInOrder(XDocument slideXml)
    {
        return !slideXml
            .Descendants(PresentationNamespace + "graphicFrame")
            .Any(frame => !IsTableGraphicFrame(frame));
    }

    private static void RenderOrderedShapeTextContainer(
        XElement container,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        OoxPackage package,
        PptxDocument document,
        PdfGraphicsBuilder graphics,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex,
        PptxTheme theme,
        IReadOnlyDictionary<string, RenderedFont> fonts,
        List<PdfImageResource> images,
        ref int imageIndex,
        GroupTransform transform,
        bool renderPlaceholders,
        IReadOnlyList<XDocument> placeholderSources)
    {
        foreach (XElement child in container.Elements())
        {
            if (child.Name == PresentationNamespace + "sp")
            {
                if (renderPlaceholders || !IsPlaceholder(child))
                {
                    RenderShape(child, document, graphics, theme, transform);
                    DrawTextRunsWithFonts(ReadTextRunsForShape(child, document, theme, renderPlaceholders, placeholderSources), graphics, fonts);
                }

                continue;
            }

            if (child.Name == PresentationNamespace + "cxnSp")
            {
                RenderShape(child, document, graphics, theme, transform);
                continue;
            }

            if (child.Name == PresentationNamespace + "pic")
            {
                RenderPicture(child, relationships, package, document, graphics, diagnosticSink, slideIndex, transform, images, ref imageIndex);
                continue;
            }

            if (child.Name == PresentationNamespace + "graphicFrame")
            {
                IReadOnlyList<TextRun> tableTextRuns = RenderTableFrame(child, document, graphics, theme);
                DrawTextRunsWithFonts(tableTextRuns, graphics, fonts);
                continue;
            }

            if (child.Name == PresentationNamespace + "grpSp")
            {
                RenderOrderedShapeTextContainer(child, relationships, package, document, graphics, diagnosticSink, slideIndex, theme, fonts, images, ref imageIndex, transform.Combine(ReadGroupTransform(child)), renderPlaceholders, placeholderSources);
            }
        }
    }

    private static void RenderShapeContainer(XElement container, PptxDocument document, PdfGraphicsBuilder graphics, PptxTheme theme, GroupTransform transform, bool renderPlaceholders)
    {
        foreach (XElement child in container.Elements())
        {
            if (child.Name == PresentationNamespace + "sp")
            {
                if (renderPlaceholders || !IsPlaceholder(child))
                {
                    RenderShape(child, document, graphics, theme, transform);
                }

                continue;
            }

            if (child.Name == PresentationNamespace + "cxnSp")
            {
                RenderShape(child, document, graphics, theme, transform);
                continue;
            }

            if (child.Name == PresentationNamespace + "grpSp")
            {
                GroupTransform childTransform = transform.Combine(ReadGroupTransform(child));
                RenderShapeContainer(child, document, graphics, theme, childTransform, renderPlaceholders);
            }
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

        bool hasFill = TryReadSolidColorWithAlpha(shapeProperties, theme, out RgbColor fill, out double fillAlpha);
        bool hasStroke = TryReadLineWithAlpha(shapeProperties, theme, out RgbColor stroke, out double lineWidth, out double strokeAlpha);
        bool hasDash = TryReadPresetDash(shapeProperties, lineWidth, out IReadOnlyList<double> dashPattern);
        int? lineCap = ReadLineCap(shapeProperties) switch
        {
            "rnd" => 1,
            "sq" => 2,
            _ => null
        };
        int? lineJoin = ReadLineJoin(shapeProperties);

        if (transformed)
        {
            graphics.SaveState();
            ApplyShapeTransform(graphics, x, y, width, height, bounds);
        }

        if (preset is "line" or "straightConnector1")
        {
            if (hasStroke)
            {
                bool transparentStroke = strokeAlpha < 0.999d;
                if (transparentStroke)
                {
                    graphics.SaveState();
                    graphics.SetAlpha(strokeAlpha, strokeAlpha);
                }

                double x1 = x;
                double y1 = document.SlideHeightPoints - yTop;
                double x2 = x + width;
                double y2 = document.SlideHeightPoints - yTop - height;
                string? headArrowType = ReadLineEndType(shapeProperties, "headEnd");
                string? tailArrowType = ReadLineEndType(shapeProperties, "tailEnd");
                bool hasHeadArrow = IsFilledTriangleArrow(headArrowType);
                bool hasTailArrow = IsFilledTriangleArrow(tailArrowType);
                if ((hasHeadArrow || hasTailArrow) && !hasDash && lineCap is null)
                {
                    graphics.SetFillRgb(stroke.Red, stroke.Green, stroke.Blue);
                    bool usesOfficeArrowType = headArrowType == "arrow" || tailArrowType == "arrow";
                    if (usesOfficeArrowType)
                    {
                        FillOfficeArrowedLine(graphics, x1, y1, x2, y2, lineWidth, hasHeadArrow, hasTailArrow);
                    }
                    else
                    {
                        FillArrowedLine(graphics, x1, y1, x2, y2, lineWidth, hasHeadArrow, hasTailArrow);
                    }
                }
                else
                {
                    graphics.SetStrokeRgb(stroke.Red, stroke.Green, stroke.Blue);
                    graphics.SetLineWidth(lineWidth);
                    if (hasDash)
                    {
                        graphics.SetLineDash(dashPattern);
                    }

                    if (lineCap is { } cap)
                    {
                        graphics.SetLineCap(cap);
                        graphics.SetLineJoin(1);
                    }

                    graphics.StrokeLine(x1, y1, x2, y2);
                    if (hasDash)
                    {
                        graphics.ClearLineDash();
                    }

                    if (lineCap is not null)
                    {
                        graphics.SetLineCap(0);
                        graphics.SetLineJoin(0);
                    }
                }

                if (transparentStroke)
                {
                    graphics.RestoreState();
                }
            }

            if (transformed)
            {
                graphics.RestoreState();
            }

            return;
        }

        if (hasFill)
        {
            bool transparentFill = fillAlpha < 0.999d;
            if (transparentFill)
            {
                graphics.SaveState();
                graphics.SetAlpha(fillAlpha, 1d);
            }

            graphics.SetFillRgb(fill.Red, fill.Green, fill.Blue);
            if (preset == "ellipse")
            {
                graphics.FillEllipse(x, y, width, height);
            }
            else if (preset == "roundRect")
            {
                graphics.FillRoundedRectangle(x, y, width, height, Math.Min(width, height) * 0.16d);
            }
            else if (TryCreatePresetPolygonPoints(preset, x, y, width, height, out (double X, double Y)[] polygonPoints))
            {
                graphics.FillPolygon(polygonPoints);
            }
            else
            {
                graphics.FillRectangle(x, y, width, height);
            }

            if (transparentFill)
            {
                graphics.RestoreState();
            }
        }

        if (hasStroke)
        {
            bool transparentStroke = strokeAlpha < 0.999d;
            if (transparentStroke)
            {
                graphics.SaveState();
                graphics.SetAlpha(1d, strokeAlpha);
            }

            graphics.SetStrokeRgb(stroke.Red, stroke.Green, stroke.Blue);
            graphics.SetLineWidth(lineWidth);
            if (hasDash)
            {
                graphics.SetLineDash(dashPattern);
            }

            int? appliedLineJoin = lineJoin ?? (lineCap is null ? null : 1);
            if (lineCap is { } cap)
            {
                graphics.SetLineCap(cap);
            }

            if (appliedLineJoin is { } join)
            {
                graphics.SetLineJoin(join);
            }

            if (preset == "ellipse")
            {
                graphics.StrokeEllipse(x, y, width, height);
            }
            else if (preset == "roundRect")
            {
                graphics.StrokeRoundedRectangle(x, y, width, height, Math.Min(width, height) * 0.16d);
            }
            else if (TryCreatePresetPolygonPoints(preset, x, y, width, height, out (double X, double Y)[] polygonPoints))
            {
                graphics.StrokePolygon(polygonPoints);
            }
            else
            {
                graphics.StrokeRectangle(x, y, width, height);
            }

            if (hasDash)
            {
                graphics.ClearLineDash();
            }

            if (lineCap is not null)
            {
                graphics.SetLineCap(0);
            }

            if (appliedLineJoin is not null)
            {
                graphics.SetLineJoin(0);
            }

            if (transparentStroke)
            {
                graphics.RestoreState();
            }
        }

        if (transformed)
        {
            graphics.RestoreState();
        }
    }

    private static void FillLineArrowhead(PdfGraphicsBuilder graphics, double tipX, double tipY, double directionX, double directionY, double lineWidth)
    {
        double length = Math.Sqrt(directionX * directionX + directionY * directionY);
        if (length <= 0.001d)
        {
            return;
        }

        double ux = directionX / length;
        double uy = directionY / length;
        double nx = -uy;
        double ny = ux;
        double size = Math.Max(5d, lineWidth * 3.5d);
        double baseX = tipX - ux * size;
        double baseY = tipY - uy * size;
        double halfWidth = size * 0.45d;
        graphics.FillPolygon(
        [
            (tipX, tipY),
            (baseX + nx * halfWidth, baseY + ny * halfWidth),
            (baseX - nx * halfWidth, baseY - ny * halfWidth)
        ]);
    }

    private static void FillArrowedLine(PdfGraphicsBuilder graphics, double x1, double y1, double x2, double y2, double lineWidth, bool headArrow, bool tailArrow)
    {
        double dx = x2 - x1;
        double dy = y2 - y1;
        double length = Math.Sqrt(dx * dx + dy * dy);
        if (length <= 0.001d)
        {
            return;
        }

        double ux = dx / length;
        double uy = dy / length;
        double nx = -uy;
        double ny = ux;
        double half = lineWidth / 2d;
        double arrowLength = lineWidth * 3d;
        double arrowHalfWidth = lineWidth * 1.5d;
        double startX = headArrow ? x1 + ux * (arrowLength - half) : x1;
        double startY = headArrow ? y1 + uy * (arrowLength - half) : y1;
        double endX = tailArrow ? x2 - ux * (arrowLength - half) : x2;
        double endY = tailArrow ? y2 - uy * (arrowLength - half) : y2;

        graphics.FillPolygon(
        [
            (startX + nx * half, startY + ny * half),
            (endX + nx * half, endY + ny * half),
            (endX - nx * half, endY - ny * half),
            (startX - nx * half, startY - ny * half)
        ]);

        if (headArrow)
        {
            FillTriangle(graphics, x1, y1, x1 + ux * arrowLength, y1 + uy * arrowLength, nx, ny, arrowHalfWidth);
        }

        if (tailArrow)
        {
            FillTriangle(graphics, x2, y2, x2 - ux * arrowLength, y2 - uy * arrowLength, nx, ny, arrowHalfWidth);
        }
    }

    private static void FillTriangle(PdfGraphicsBuilder graphics, double tipX, double tipY, double baseX, double baseY, double nx, double ny, double halfWidth)
    {
        graphics.FillPolygon(
        [
            (tipX, tipY),
            (baseX + nx * halfWidth, baseY + ny * halfWidth),
            (baseX - nx * halfWidth, baseY - ny * halfWidth)
        ]);
    }

    private static void FillOfficeArrowedLine(PdfGraphicsBuilder graphics, double x1, double y1, double x2, double y2, double lineWidth, bool headArrow, bool tailArrow)
    {
        double dx = x2 - x1;
        double dy = y2 - y1;
        double length = Math.Sqrt(dx * dx + dy * dy);
        if (length <= 0.001d)
        {
            return;
        }

        double ux = dx / length;
        double uy = dy / length;
        double nx = -uy;
        double ny = ux;
        double half = lineWidth / 2d;
        double shaftInset = lineWidth * 0.99d;
        double startX = headArrow ? x1 + ux * shaftInset : x1;
        double startY = headArrow ? y1 + uy * shaftInset : y1;
        double endX = tailArrow ? x2 - ux * shaftInset : x2;
        double endY = tailArrow ? y2 - uy * shaftInset : y2;

        graphics.MoveTo(startX + nx * half, startY + ny * half);
        graphics.LineTo(endX + nx * half, endY + ny * half);
        graphics.LineTo(endX - nx * half, endY - ny * half);
        graphics.LineTo(startX - nx * half, startY - ny * half);
        graphics.ClosePath();

        if (headArrow)
        {
            AppendOfficeArrowHeadPath(graphics, x1, y1, ux, uy, nx, ny, lineWidth);
        }

        if (tailArrow)
        {
            AppendOfficeArrowHeadPath(graphics, x2, y2, -ux, -uy, -nx, -ny, lineWidth);
        }

        graphics.FillCurrentPath();
    }

    private static void AppendOfficeArrowHeadPath(PdfGraphicsBuilder graphics, double tipX, double tipY, double ux, double uy, double nx, double ny, double lineWidth)
    {
        tipX -= ux * 0.01d;
        tipY -= uy * 0.01d;

        (double X, double Y) Point(double along, double normal)
        {
            return (tipX + ux * along + nx * normal, tipY + uy * along + ny * normal);
        }

        (double X, double Y) p = Point(lineWidth * 3.74d, -lineWidth * 2.181667d);
        graphics.MoveTo(p.X, p.Y);
        p = Point(0d, 0d);
        graphics.LineTo(p.X, p.Y);
        p = Point(lineWidth * 3.74d, lineWidth * 2.181667d);
        graphics.LineTo(p.X, p.Y);
        (double X, double Y) c1 = Point(lineWidth * 3.978333d, lineWidth * 2.321667d);
        (double X, double Y) c2 = Point(lineWidth * 4.285d, lineWidth * 2.24d);
        (double X, double Y) c3 = Point(lineWidth * 4.423333d, lineWidth * 2.001667d);
        graphics.CurveTo(c1.X, c1.Y, c2.X, c2.Y, c3.X, c3.Y);
        p = Point(lineWidth * 4.243333d, lineWidth * 1.318333d);
        graphics.LineTo(p.X, p.Y);
        p = Point(lineWidth * 1.243333d, -lineWidth * 0.431667d);
        graphics.LineTo(p.X, p.Y);
        p = Point(lineWidth * 1.243333d, lineWidth * 0.431667d);
        graphics.LineTo(p.X, p.Y);
        p = Point(lineWidth * 4.243333d, -lineWidth * 1.318333d);
        graphics.LineTo(p.X, p.Y);
        c1 = Point(lineWidth * 4.481667d, -lineWidth * 1.456667d);
        c2 = Point(lineWidth * 4.563333d, -lineWidth * 1.763333d);
        c3 = Point(lineWidth * 4.423333d, -lineWidth * 2.001667d);
        graphics.CurveTo(c1.X, c1.Y, c2.X, c2.Y, c3.X, c3.Y);
        c1 = Point(lineWidth * 4.285d, -lineWidth * 2.24d);
        c2 = Point(lineWidth * 3.978333d, -lineWidth * 2.321667d);
        c3 = Point(lineWidth * 3.74d, -lineWidth * 2.181667d);
        graphics.CurveTo(c1.X, c1.Y, c2.X, c2.Y, c3.X, c3.Y);
        graphics.ClosePath();
    }

    private static string? ReadLineEndType(XElement shapeProperties, string elementName)
    {
        return (string?)shapeProperties
            .Element(DrawingNamespace + "ln")
            ?.Element(DrawingNamespace + elementName)
            ?.Attribute("type");
    }

    private static bool IsFilledTriangleArrow(string? type)
    {
        return type is "arrow" or "triangle";
    }

    private static bool TryReadPresetDash(XElement shapeProperties, double lineWidth, out IReadOnlyList<double> dashPattern)
    {
        string? presetDash = (string?)shapeProperties
            .Element(DrawingNamespace + "ln")
            ?.Element(DrawingNamespace + "prstDash")
            ?.Attribute("val");
        if (presetDash == "dash")
        {
            dashPattern = [lineWidth * 4d, lineWidth * 3d];
            return true;
        }

        if (presetDash == "dashDot")
        {
            dashPattern = [lineWidth * 4d, lineWidth * 3d, lineWidth, lineWidth * 3d];
            return true;
        }

        dashPattern = [];
        return false;
    }

    private static string? ReadLineCap(XElement shapeProperties)
    {
        return (string?)shapeProperties
            .Element(DrawingNamespace + "ln")
            ?.Attribute("cap");
    }

    private static int? ReadLineJoin(XElement shapeProperties)
    {
        XElement? line = shapeProperties.Element(DrawingNamespace + "ln");
        if (line?.Element(DrawingNamespace + "round") is not null)
        {
            return 1;
        }

        if (line?.Element(DrawingNamespace + "bevel") is not null)
        {
            return 2;
        }

        if (line?.Element(DrawingNamespace + "miter") is not null)
        {
            return 0;
        }

        return null;
    }

    private static bool TryCreatePresetPolygonPoints(string preset, double x, double y, double width, double height, out (double X, double Y)[] points)
    {
        points = preset switch
        {
            "triangle" => CreateTrianglePoints(x, y, width, height),
            "rtTriangle" => CreateRightTrianglePoints(x, y, width, height),
            "diamond" => CreateDiamondPoints(x, y, width, height),
            "pentagon" => CreatePentagonPoints(x, y, width, height),
            "hexagon" => CreateHexagonPoints(x, y, width, height),
            "octagon" => CreateOctagonPoints(x, y, width, height),
            "star4" => CreateStar4Points(x, y, width, height),
            "star5" => CreateStar5Points(x, y, width, height),
            "star6" => CreateStar6Points(x, y, width, height),
            "star8" => CreateStar8Points(x, y, width, height),
            "plus" => CreatePlusPoints(x, y, width, height),
            "chevron" => CreateChevronPoints(x, y, width, height),
            "homePlate" => CreateHomePlatePoints(x, y, width, height),
            "parallelogram" => CreateParallelogramPoints(x, y, width, height),
            "trapezoid" => CreateTrapezoidPoints(x, y, width, height),
            "downArrow" => CreateDownArrowPoints(x, y, width, height),
            "upArrow" => CreateUpArrowPoints(x, y, width, height),
            "leftArrow" => CreateLeftArrowPoints(x, y, width, height),
            "rightArrow" => CreateRightArrowPoints(x, y, width, height),
            "leftRightArrow" => CreateLeftRightArrowPoints(x, y, width, height),
            "upDownArrow" => CreateUpDownArrowPoints(x, y, width, height),
            "wedgeRectCallout" => CreateWedgeRectCalloutPoints(x, y, width, height),
            _ => []
        };
        return points.Length != 0;
    }

    private static (double X, double Y)[] CreateTrianglePoints(double x, double y, double width, double height)
    {
        return
        [
            (x + width / 2d, y + height),
            (x + width, y),
            (x, y)
        ];
    }

    private static (double X, double Y)[] CreateRightTrianglePoints(double x, double y, double width, double height)
    {
        return
        [
            (x, y),
            (x, y + height),
            (x + width, y)
        ];
    }

    private static (double X, double Y)[] CreateDiamondPoints(double x, double y, double width, double height)
    {
        return
        [
            (x + width / 2d, y + height),
            (x + width, y + height / 2d),
            (x + width / 2d, y),
            (x, y + height / 2d)
        ];
    }

    private static (double X, double Y)[] CreateStar4Points(double x, double y, double width, double height)
    {
        double innerLeftX = RoundOfficeShapeCoordinate(width * 0.41161d);
        double innerRightX = RoundOfficeShapeCoordinate(width * 0.58839d);
        double innerBottomY = RoundOfficeShapeCoordinate(height * 0.41159d);
        double innerTopY = RoundOfficeShapeCoordinate(height * 0.58841d);
        return
        [
            (x, y + height / 2d),
            (x + innerLeftX, y + innerTopY),
            (x + width / 2d, y + height),
            (x + innerRightX, y + innerTopY),
            (x + width, y + height / 2d),
            (x + innerRightX, y + innerBottomY),
            (x + width / 2d, y),
            (x + innerLeftX, y + innerBottomY)
        ];
    }

    private static (double X, double Y)[] CreateStar5Points(double x, double y, double width, double height)
    {
        double topShoulderY = RoundOfficeShapeCoordinate(height * 0.61803d);
        double lowerShoulderY = RoundOfficeShapeCoordinate(height * 0.38197d);
        double innerBottomY = RoundOfficeShapeCoordinate(height * 0.23607d);
        double innerTopLeftX = RoundOfficeShapeCoordinate(width * 0.38194d);
        double innerTopRightX = RoundOfficeShapeCoordinate(width * 0.61806d);
        double innerLeftX = RoundOfficeShapeCoordinate(width * 0.30902d);
        double innerRightX = RoundOfficeShapeCoordinate(width * 0.69098d);
        double lowerLeftX = RoundOfficeShapeCoordinate(width * 0.19098d);
        double lowerRightX = RoundOfficeShapeCoordinate(width * 0.80902d);
        return
        [
            (x, y + topShoulderY),
            (x + innerTopLeftX, y + topShoulderY),
            (x + width / 2d, y + height),
            (x + innerTopRightX, y + topShoulderY),
            (x + width, y + topShoulderY),
            (x + innerRightX, y + lowerShoulderY),
            (x + lowerRightX, y),
            (x + width / 2d, y + innerBottomY),
            (x + lowerLeftX, y),
            (x + innerLeftX, y + lowerShoulderY)
        ];
    }

    private static (double X, double Y)[] CreateStar6Points(double x, double y, double width, double height)
    {
        double quarterWidth = width / 6d;
        double quarterHeight = height / 4d;
        return
        [
            (x, y + height * 0.75d),
            (x + width / 3d, y + height * 0.75d),
            (x + width / 2d, y + height),
            (x + width * 2d / 3d, y + height * 0.75d),
            (x + width, y + height * 0.75d),
            (x + width - quarterWidth, y + height / 2d),
            (x + width, y + height * 0.25d),
            (x + width * 2d / 3d, y + height * 0.25d),
            (x + width / 2d, y),
            (x + width / 3d, y + height * 0.25d),
            (x, y + height * 0.25d),
            (x + quarterWidth, y + height / 2d)
        ];
    }

    private static (double X, double Y)[] CreateStar8Points(double x, double y, double width, double height)
    {
        return
        [
            (x, y + height / 2d),
            (x + RoundOfficeShapeCoordinate(width * 0.15356d), y + RoundOfficeShapeCoordinate(height * 0.64348d)),
            (x + RoundOfficeShapeCoordinate(width * 0.14644d), y + RoundOfficeShapeCoordinate(height * 0.85356d)),
            (x + RoundOfficeShapeCoordinate(width * 0.3565d), y + RoundOfficeShapeCoordinate(height * 0.84644d)),
            (x + width / 2d, y + height),
            (x + RoundOfficeShapeCoordinate(width * 0.6435d), y + RoundOfficeShapeCoordinate(height * 0.84644d)),
            (x + RoundOfficeShapeCoordinate(width * 0.85356d), y + RoundOfficeShapeCoordinate(height * 0.85356d)),
            (x + RoundOfficeShapeCoordinate(width * 0.84644d), y + RoundOfficeShapeCoordinate(height * 0.64348d)),
            (x + width, y + height / 2d),
            (x + RoundOfficeShapeCoordinate(width * 0.84644d), y + RoundOfficeShapeCoordinate(height * 0.35652d)),
            (x + RoundOfficeShapeCoordinate(width * 0.85356d), y + RoundOfficeShapeCoordinate(height * 0.14644d)),
            (x + RoundOfficeShapeCoordinate(width * 0.6435d), y + RoundOfficeShapeCoordinate(height * 0.15356d)),
            (x + width / 2d, y),
            (x + RoundOfficeShapeCoordinate(width * 0.3565d), y + RoundOfficeShapeCoordinate(height * 0.15356d)),
            (x + RoundOfficeShapeCoordinate(width * 0.14644d), y + RoundOfficeShapeCoordinate(height * 0.14644d)),
            (x + RoundOfficeShapeCoordinate(width * 0.15356d), y + RoundOfficeShapeCoordinate(height * 0.35652d))
        ];
    }

    private static (double X, double Y)[] CreatePentagonPoints(double x, double y, double width, double height)
    {
        double golden = (Math.Sqrt(5d) - 1d) / 2d;
        double bottomInset = RoundOfficeShapeCoordinate(width * (1d - golden) / 2d);
        double shoulderY = RoundOfficeShapeCoordinate(y + height * golden);
        return
        [
            (x, shoulderY),
            (x + width / 2d, y + height),
            (x + width, shoulderY),
            (x + width - bottomInset, y),
            (x + bottomInset, y)
        ];
    }

    private static (double X, double Y)[] CreateHexagonPoints(double x, double y, double width, double height)
    {
        double inset = Math.Min(width / 2d, height / 4d);
        return
        [
            (x, y + height / 2d),
            (x + inset, y + height),
            (x + width - inset, y + height),
            (x + width, y + height / 2d),
            (x + width - inset, y),
            (x + inset, y)
        ];
    }

    private static (double X, double Y)[] CreateOctagonPoints(double x, double y, double width, double height)
    {
        double inset = RoundOfficeShapeCoordinate(Math.Min(width, height) * (1d - Math.Sqrt(0.5d)));
        return
        [
            (x, y + height - inset),
            (x + inset, y + height),
            (x + width - inset, y + height),
            (x + width, y + height - inset),
            (x + width, y + inset),
            (x + width - inset, y),
            (x + inset, y),
            (x, y + inset)
        ];
    }

    private static double RoundOfficeShapeCoordinate(double value)
    {
        return Math.Round(value, 2, MidpointRounding.AwayFromZero);
    }

    private static (double X, double Y)[] CreatePlusPoints(double x, double y, double width, double height)
    {
        double armWidth = Math.Min(width, height) / 4d;
        double rightArmX = x + width - armWidth;
        double topArmY = y + height - armWidth;
        return
        [
            (x, topArmY),
            (x + armWidth, topArmY),
            (x + armWidth, y + height),
            (rightArmX, y + height),
            (rightArmX, topArmY),
            (x + width, topArmY),
            (x + width, y + armWidth),
            (rightArmX, y + armWidth),
            (rightArmX, y),
            (x + armWidth, y),
            (x + armWidth, y + armWidth),
            (x, y + armWidth)
        ];
    }

    private static (double X, double Y)[] CreateChevronPoints(double x, double y, double width, double height)
    {
        double inset = Math.Min(width / 2d, height / 2d);
        return
        [
            (x, y + height),
            (x + width - inset, y + height),
            (x + width, y + height / 2d),
            (x + width - inset, y),
            (x, y),
            (x + inset, y + height / 2d)
        ];
    }

    private static (double X, double Y)[] CreateHomePlatePoints(double x, double y, double width, double height)
    {
        double inset = Math.Min(width, height / 2d);
        return
        [
            (x, y + height),
            (x + width - inset, y + height),
            (x + width, y + height / 2d),
            (x + width - inset, y),
            (x, y)
        ];
    }

    private static (double X, double Y)[] CreateParallelogramPoints(double x, double y, double width, double height)
    {
        double inset = Math.Min(width, height / 4d);
        return
        [
            (x + inset, y + height),
            (x + width, y + height),
            (x + width - inset, y),
            (x, y)
        ];
    }

    private static (double X, double Y)[] CreateTrapezoidPoints(double x, double y, double width, double height)
    {
        double inset = Math.Min(width / 2d, height / 4d);
        return
        [
            (x + inset, y + height),
            (x + width - inset, y + height),
            (x + width, y),
            (x, y)
        ];
    }

    private static (double X, double Y)[] CreateDownArrowPoints(double x, double y, double width, double height)
    {
        double arrowShoulderY = y + Math.Min(height, width / 2d);
        return
        [
            (x + width * 0.25d, y + height),
            (x + width * 0.75d, y + height),
            (x + width * 0.75d, arrowShoulderY),
            (x + width, arrowShoulderY),
            (x + width * 0.5d, y),
            (x, arrowShoulderY),
            (x + width * 0.25d, arrowShoulderY)
        ];
    }

    private static (double X, double Y)[] CreateUpArrowPoints(double x, double y, double width, double height)
    {
        double arrowShoulderY = y + Math.Max(0d, height - width / 2d);
        return
        [
            (x + width * 0.25d, y),
            (x + width * 0.75d, y),
            (x + width * 0.75d, arrowShoulderY),
            (x + width, arrowShoulderY),
            (x + width * 0.5d, y + height),
            (x, arrowShoulderY),
            (x + width * 0.25d, arrowShoulderY)
        ];
    }

    private static (double X, double Y)[] CreateRightArrowPoints(double x, double y, double width, double height)
    {
        double arrowShoulderX = x + Math.Max(0d, width - height / 2d);
        return
        [
            (x, y + height * 0.25d),
            (arrowShoulderX, y + height * 0.25d),
            (arrowShoulderX, y),
            (x + width, y + height * 0.5d),
            (arrowShoulderX, y + height),
            (arrowShoulderX, y + height * 0.75d),
            (x, y + height * 0.75d)
        ];
    }

    private static (double X, double Y)[] CreateLeftArrowPoints(double x, double y, double width, double height)
    {
        double arrowShoulderX = x + Math.Min(width, height / 2d);
        return
        [
            (x + width, y + height * 0.25d),
            (arrowShoulderX, y + height * 0.25d),
            (arrowShoulderX, y),
            (x, y + height * 0.5d),
            (arrowShoulderX, y + height),
            (arrowShoulderX, y + height * 0.75d),
            (x + width, y + height * 0.75d)
        ];
    }

    private static (double X, double Y)[] CreateLeftRightArrowPoints(double x, double y, double width, double height)
    {
        double headDepth = Math.Min(width / 2d, height / 2d);
        double shaftBottom = y + height * 0.25d;
        double shaftTop = y + height * 0.75d;
        return
        [
            (x, y + height / 2d),
            (x + headDepth, y + height),
            (x + headDepth, shaftTop),
            (x + width - headDepth, shaftTop),
            (x + width - headDepth, y + height),
            (x + width, y + height / 2d),
            (x + width - headDepth, y),
            (x + width - headDepth, shaftBottom),
            (x + headDepth, shaftBottom),
            (x + headDepth, y)
        ];
    }

    private static (double X, double Y)[] CreateUpDownArrowPoints(double x, double y, double width, double height)
    {
        double headDepth = Math.Min(height / 2d, width / 2d);
        double shaftLeft = x + width * 0.25d;
        double shaftRight = x + width * 0.75d;
        return
        [
            (x, y + height - headDepth),
            (x + width / 2d, y + height),
            (x + width, y + height - headDepth),
            (shaftRight, y + height - headDepth),
            (shaftRight, y + headDepth),
            (x + width, y + headDepth),
            (x + width / 2d, y),
            (x, y + headDepth),
            (shaftLeft, y + headDepth),
            (shaftLeft, y + height - headDepth)
        ];
    }

    private static (double X, double Y)[] CreateWedgeRectCalloutPoints(double x, double y, double width, double height)
    {
        double tailLeftX = x + width / 6d;
        double tailTipX = x + width * 7d / 24d;
        double tailRightX = x + width * 5d / 12d;
        double tailTipY = y - height / 8d;
        return
        [
            (x, y + height),
            (x + width, y + height),
            (x + width, y),
            (tailRightX, y),
            (tailTipX, tailTipY),
            (tailLeftX, y),
            (x, y)
        ];
    }

    private static IReadOnlyList<TextRun> RenderTables(XDocument slideXml, PptxDocument document, PdfGraphicsBuilder graphics, PptxTheme theme)
    {
        var textRuns = new List<TextRun>();
        foreach (XElement frame in slideXml.Descendants(PresentationNamespace + "graphicFrame"))
        {
            textRuns.AddRange(RenderTableFrame(frame, document, graphics, theme));
        }

        return textRuns;
    }

    private static bool IsTableGraphicFrame(XElement frame)
    {
        return frame
            .Element(DrawingNamespace + "graphic")
            ?.Element(DrawingNamespace + "graphicData")
            ?.Element(DrawingNamespace + "tbl") is not null;
    }

    private static IReadOnlyList<TextRun> RenderTableFrame(XElement frame, PptxDocument document, PdfGraphicsBuilder graphics, PptxTheme theme)
    {
        var textRuns = new List<TextRun>();
        ShapeBounds? bounds = ReadGraphicFrameBounds(frame);
        XElement? table = frame
            .Element(DrawingNamespace + "graphic")
            ?.Element(DrawingNamespace + "graphicData")
            ?.Element(DrawingNamespace + "tbl");
        if (bounds is null || table is null)
        {
            return textRuns;
        }

        IReadOnlyList<double> rawColumnWidths = table
                .Element(DrawingNamespace + "tblGrid")
                ?.Elements(DrawingNamespace + "gridCol")
                .Select(column => Math.Max(1d, ParseOptionalLongAttribute(column, "w", 1)))
                .ToArray() ?? [];
        IReadOnlyList<XElement> rows = table.Elements(DrawingNamespace + "tr").ToArray();
        if (rawColumnWidths.Count == 0 || rows.Count == 0)
        {
            return textRuns;
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
        var rowTops = new double[rows.Count + 1];
        var skippedVerticalGridSegments = new bool[rawColumnWidths.Count + 1, rows.Count];
        var skippedHorizontalGridSegments = new bool[rows.Count + 1, rawColumnWidths.Count];
        var explicitBorders = new List<TableBorderLine>();
        rowTops[0] = yTop;
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            double rowHeight = rawRowHeights[rowIndex] * rowScale;
            double cellY = yTop - rowHeight;
            IReadOnlyList<XElement> cells = rows[rowIndex].Elements(DrawingNamespace + "tc").ToArray();

            double cellX = frameX;
            int columnIndex = 0;
            foreach (XElement cell in cells)
            {
                if (columnIndex >= rawColumnWidths.Count)
                {
                    break;
                }

                if (IsMergedTableCellContinuation(cell))
                {
                    cellX += rawColumnWidths[columnIndex] * columnScale;
                    columnIndex++;
                    continue;
                }

                int columnSpan = Math.Min(ReadTableCellColumnSpan(cell), rawColumnWidths.Count - columnIndex);
                int rowSpan = Math.Min(ReadTableCellRowSpan(cell), rows.Count - rowIndex);
                for (int boundary = columnIndex + 1; boundary < columnIndex + columnSpan; boundary++)
                {
                    skippedVerticalGridSegments[boundary, rowIndex] = true;
                }

                for (int boundary = rowIndex + 1; boundary < rowIndex + rowSpan; boundary++)
                {
                    for (int skippedColumn = columnIndex; skippedColumn < columnIndex + columnSpan; skippedColumn++)
                    {
                        skippedHorizontalGridSegments[boundary, skippedColumn] = true;
                    }
                }

                double columnWidth = rawColumnWidths
                        .Skip(columnIndex)
                        .Take(columnSpan)
                        .Sum() * columnScale;
                double cellHeight = rawRowHeights
                        .Skip(rowIndex)
                        .Take(rowSpan)
                        .Sum() * rowScale;
                double cellTop = yTop;
                double cellBottom = cellTop - cellHeight;
                XElement? cellProperties = cell.Element(DrawingNamespace + "tcPr");

                if (TryReadSolidColorWithAlpha(cellProperties, theme, out RgbColor fill, out double fillAlpha))
                {
                    bool transparentFill = fillAlpha < 0.999d;
                    if (transparentFill)
                    {
                        graphics.SaveState();
                        graphics.SetAlpha(fillAlpha, 1d);
                    }

                    graphics.SetFillRgb(fill.Red, fill.Green, fill.Blue);
                    graphics.FillRectangle(cellX, cellBottom, columnWidth, cellHeight);
                    if (transparentFill)
                    {
                        graphics.RestoreState();
                    }
                }

                AddTableCellBorders(explicitBorders, cellProperties, theme, cellX, cellBottom, columnWidth, cellHeight);
                AddTableCellTextRuns(cell, cellX, cellBottom, columnWidth, cellHeight, theme, textRuns);
                cellX += columnWidth;
                columnIndex += columnSpan;
            }

            yTop -= rowHeight;
            rowTops[rowIndex + 1] = yTop;
        }

        if (!TableHasExplicitBorders(table))
        {
            StrokeDefaultTableGrid(graphics, frameX, frameTop, frameWidth, frameHeight, rawColumnWidths.Select(width => width * columnScale).ToArray(), rowTops, table, skippedVerticalGridSegments, skippedHorizontalGridSegments);
        }
        else
        {
            StrokeTableBorders(graphics, explicitBorders);
        }

        return textRuns;
    }

    private static bool IsMergedTableCellContinuation(XElement cell)
    {
        return ParseOptionalBoolAttribute(cell, "hMerge") ||
            ParseOptionalBoolAttribute(cell, "vMerge");
    }

    private static int ReadTableCellColumnSpan(XElement cell)
    {
        return cell.Attribute("gridSpan") is { } spanAttribute &&
            int.TryParse(spanAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int span)
            ? Math.Max(1, span)
            : 1;
    }

    private static int ReadTableCellRowSpan(XElement cell)
    {
        return cell.Attribute("rowSpan") is { } spanAttribute &&
            int.TryParse(spanAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int span)
            ? Math.Max(1, span)
            : 1;
    }

    private static bool TableHasExplicitBorders(XElement table)
    {
        return table
            .Descendants(DrawingNamespace + "tcPr")
            .Any(cellProperties =>
                cellProperties.Element(DrawingNamespace + "lnL") is not null ||
                cellProperties.Element(DrawingNamespace + "lnR") is not null ||
                cellProperties.Element(DrawingNamespace + "lnT") is not null ||
                cellProperties.Element(DrawingNamespace + "lnB") is not null);
    }

    private static void StrokeDefaultTableGrid(PdfGraphicsBuilder graphics, double x, double yTop, double width, double height, IReadOnlyList<double> columnWidths, IReadOnlyList<double> rowTops, XElement table, bool[,] skippedVerticalGridSegments, bool[,] skippedHorizontalGridSegments)
    {
        bool hasTableStyle = table
            .Element(DrawingNamespace + "tblPr")
            ?.Element(DrawingNamespace + "tableStyleId") is not null;
        if (hasTableStyle)
        {
            graphics.SetStrokeRgb(255, 255, 255);
        }
        else
        {
            graphics.SetStrokeRgb(0, 0, 0);
        }

        double cursorX = x;
        graphics.SetLineWidth(1d);
        StrokeDefaultVerticalGridLine(graphics, cursorX, yTop, height, rowTops, skippedVerticalGridSegments, 0);
        for (int columnIndex = 0; columnIndex < columnWidths.Count; columnIndex++)
        {
            cursorX += columnWidths[columnIndex];
            StrokeDefaultVerticalGridLine(graphics, cursorX, yTop, height, rowTops, skippedVerticalGridSegments, columnIndex + 1);
        }

        IReadOnlyList<XElement> rows = table.Elements(DrawingNamespace + "tr").ToArray();
        for (int i = 0; i < rowTops.Count; i++)
        {
            bool firstRowBoundary = i == 1 &&
                table.Element(DrawingNamespace + "tblPr")?.Attribute("firstRow")?.Value == "1" &&
                rows.Count > 1;
            graphics.SetLineWidth(firstRowBoundary ? 3d : 1d);
            StrokeDefaultHorizontalGridLine(graphics, x, width, rowTops[i], columnWidths, skippedHorizontalGridSegments, i);
        }
    }

    private static void StrokeDefaultVerticalGridLine(PdfGraphicsBuilder graphics, double x, double yTop, double height, IReadOnlyList<double> rowTops, bool[,] skippedSegments, int boundaryIndex)
    {
        if (boundaryIndex == 0 || boundaryIndex == skippedSegments.GetLength(0) - 1)
        {
            graphics.StrokeLine(x, yTop + 0.5d, x, yTop - height - 0.5d);
            return;
        }

        int rowCount = skippedSegments.GetLength(1);
        int rowIndex = 0;
        while (rowIndex < rowCount)
        {
            while (rowIndex < rowCount && skippedSegments[boundaryIndex, rowIndex])
            {
                rowIndex++;
            }

            if (rowIndex >= rowCount)
            {
                break;
            }

            int startRow = rowIndex;
            while (rowIndex < rowCount && !skippedSegments[boundaryIndex, rowIndex])
            {
                rowIndex++;
            }

            graphics.StrokeLine(x, rowTops[startRow] + 0.5d, x, rowTops[rowIndex] - 0.5d);
        }
    }

    private static void StrokeDefaultHorizontalGridLine(PdfGraphicsBuilder graphics, double x, double width, double y, IReadOnlyList<double> columnWidths, bool[,] skippedSegments, int boundaryIndex)
    {
        if (boundaryIndex == 0 || boundaryIndex == skippedSegments.GetLength(0) - 1)
        {
            graphics.StrokeLine(x - 0.5d, y, x + width + 0.5d, y);
            return;
        }

        int columnCount = skippedSegments.GetLength(1);
        var columnLefts = new double[columnCount + 1];
        columnLefts[0] = x;
        for (int i = 0; i < columnCount; i++)
        {
            columnLefts[i + 1] = columnLefts[i] + columnWidths[i];
        }

        int columnIndex = 0;
        while (columnIndex < columnCount)
        {
            while (columnIndex < columnCount && skippedSegments[boundaryIndex, columnIndex])
            {
                columnIndex++;
            }

            if (columnIndex >= columnCount)
            {
                break;
            }

            int startColumn = columnIndex;
            while (columnIndex < columnCount && !skippedSegments[boundaryIndex, columnIndex])
            {
                columnIndex++;
            }

            graphics.StrokeLine(columnLefts[startColumn] - 0.5d, y, columnLefts[columnIndex] + 0.5d, y);
        }
    }

    private static void AddTableCellBorders(List<TableBorderLine> borders, XElement? cellProperties, PptxTheme theme, double x, double y, double width, double height)
    {
        AddTableBorder(borders, cellProperties?.Element(DrawingNamespace + "lnL"), theme, x, y, x, y + height);
        AddTableBorder(borders, cellProperties?.Element(DrawingNamespace + "lnR"), theme, x + width, y, x + width, y + height);
        AddTableBorder(borders, cellProperties?.Element(DrawingNamespace + "lnT"), theme, x, y + height, x + width, y + height);
        AddTableBorder(borders, cellProperties?.Element(DrawingNamespace + "lnB"), theme, x, y, x + width, y);
    }

    private static void AddTableBorder(List<TableBorderLine> borders, XElement? line, PptxTheme theme, double x1, double y1, double x2, double y2)
    {
        if (line is null || line.Element(DrawingNamespace + "noFill") is not null || !TryReadSolidColorWithAlpha(line, theme, out RgbColor color, out double alpha))
        {
            return;
        }

        double lineWidth = line.Attribute("w") is { } widthAttribute
            ? Math.Max(1d, OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture)) / 2d)
            : 0.75d;
        borders.Add(new TableBorderLine(x1, y1, x2, y2, lineWidth, color, alpha));
    }

    private static void StrokeTableBorders(PdfGraphicsBuilder graphics, List<TableBorderLine> borders)
    {
        foreach (IGrouping<TableBorderKey, TableBorderLine> group in borders.GroupBy(TableBorderKey.From))
        {
            IReadOnlyList<TableBorderLine> ordered = group
                .OrderBy(border => group.Key.Vertical ? Math.Min(border.Y1, border.Y2) : Math.Min(border.X1, border.X2))
                .ToArray();
            double start = group.Key.Vertical
                ? ordered.Min(border => Math.Min(border.Y1, border.Y2))
                : ordered.Min(border => Math.Min(border.X1, border.X2));
            double end = group.Key.Vertical
                ? ordered.Max(border => Math.Max(border.Y1, border.Y2))
                : ordered.Max(border => Math.Max(border.X1, border.X2));
            double halfWidth = group.Key.LineWidth / 2d;
            double x1 = group.Key.Vertical ? group.Key.FixedCoordinate : start - halfWidth;
            double y1 = group.Key.Vertical ? start - halfWidth : group.Key.FixedCoordinate;
            double x2 = group.Key.Vertical ? group.Key.FixedCoordinate : end + halfWidth;
            double y2 = group.Key.Vertical ? end + halfWidth : group.Key.FixedCoordinate;

            bool transparentStroke = group.Key.Alpha < 0.999d;
            if (transparentStroke)
            {
                graphics.SaveState();
                graphics.SetAlpha(1d, group.Key.Alpha);
            }

            graphics.SetStrokeRgb(group.Key.Color.Red, group.Key.Color.Green, group.Key.Color.Blue);
            graphics.SetLineWidth(group.Key.LineWidth);
            graphics.StrokeLine(x1, y1, x2, y2);
            if (transparentStroke)
            {
                graphics.RestoreState();
            }
        }
    }

    private static ShapeBounds? ReadGraphicFrameBounds(XElement frame)
    {
        XElement? transform = frame.Element(PresentationNamespace + "xfrm");
        return transform is null ? null : ReadBoundsFromTransform(transform);
    }

    private static void RenderCharts(OoxPackage package, string slidePartName, XDocument slideXml, PptxDocument document, PdfGraphicsBuilder graphics, Action<OoxPdfDiagnostic>? diagnosticSink, int slideIndex)
    {
        IReadOnlyDictionary<string, OoxRelationship> relationships = package.GetRelationships(slidePartName)
            .Where(r => !r.IsExternal && r.ResolvedTarget is not null)
            .ToDictionary(r => r.Id, StringComparer.Ordinal);

        foreach (XElement frame in slideXml.Descendants(PresentationNamespace + "graphicFrame"))
        {
            XElement? graphicData = frame
                .Element(DrawingNamespace + "graphic")
                ?.Element(DrawingNamespace + "graphicData");
            if (graphicData?.Attribute("uri") is not { } uri ||
                !uri.Value.Contains("drawingml/2006/chart", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            ShapeBounds? bounds = ReadGraphicFrameBounds(frame);
            string? relationshipId = (string?)graphicData
                .Element(ChartNamespace + "chart")
                ?.Attribute(RelationshipsNamespace + "id");
            if (bounds is null || relationshipId is null || !relationships.TryGetValue(relationshipId, out OoxRelationship? relationship) || relationship.ResolvedTarget is null)
            {
                EmitChartDiagnostic(diagnosticSink, "PPTX_UNSUPPORTED_CHART", OoxPdfSeverity.Warning, "Chart frame could not be resolved and was ignored.", slidePartName, slideIndex, "Ignored");
                continue;
            }

            OoxPart? chartPart = package.GetPart(relationship.ResolvedTarget);
            if (chartPart is null)
            {
                EmitChartDiagnostic(diagnosticSink, "PPTX_UNSUPPORTED_CHART", OoxPdfSeverity.Warning, "Chart part was missing and was ignored.", relationship.ResolvedTarget, slideIndex, "Ignored");
                continue;
            }

            using Stream chartStream = chartPart.OpenRead();
            XDocument chartXml = SafeXml.Load(chartStream);
            IReadOnlyList<IReadOnlyList<double>> series = ReadBarChartSeries(chartXml);
            if (series.Count == 0)
            {
                EmitChartDiagnostic(diagnosticSink, "PPTX_UNSUPPORTED_CHART", OoxPdfSeverity.Warning, "Only bar chart cached numeric values have a static fallback.", chartPart.Name, slideIndex, "Ignored");
                continue;
            }

            RenderBarChartFallback(graphics, document, bounds.Value, series);
            EmitChartDiagnostic(diagnosticSink, "PPTX_CHART_STATIC_FALLBACK", OoxPdfSeverity.Info, "PPTX chart was rendered with an approximate static bar-chart fallback.", chartPart.Name, slideIndex, "Static bar-chart fallback");
        }
    }

    private static IReadOnlyList<IReadOnlyList<double>> ReadBarChartSeries(XDocument chartXml)
    {
        var series = new List<IReadOnlyList<double>>();
        foreach (XElement element in chartXml.Descendants(ChartNamespace + "barChart").Elements(ChartNamespace + "ser"))
        {
            double[] values = element
                .Elements(ChartNamespace + "val")
                .Descendants(ChartNamespace + "pt")
                .Select(point => (string?)point.Element(ChartNamespace + "v"))
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) ? parsed : double.NaN)
                .Where(value => !double.IsNaN(value))
                .ToArray();
            if (values.Length > 0)
            {
                series.Add(values);
            }
        }

        return series;
    }

    private static void RenderBarChartFallback(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, IReadOnlyList<IReadOnlyList<double>> series)
    {
        RgbColor[] palette =
        [
            new RgbColor(68, 114, 196),
            new RgbColor(237, 125, 49),
            new RgbColor(165, 165, 165),
            new RgbColor(255, 192, 0),
            new RgbColor(91, 155, 213),
            new RgbColor(112, 173, 71)
        ];
        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        double plotX = x + width * 0.1d;
        double plotY = y + height * 0.14d;
        double plotWidth = width * 0.82d;
        double plotHeight = height * 0.72d;
        int categoryCount = Math.Max(1, series.Max(values => values.Count));
        double maxValue = Math.Max(1d, series.SelectMany(values => values).DefaultIfEmpty(1d).Max(value => Math.Abs(value)));

        graphics.SetStrokeRgb(90, 90, 90);
        graphics.SetLineWidth(0.75d);
        graphics.StrokeLine(plotX, plotY, plotX + plotWidth, plotY);
        graphics.StrokeLine(plotX, plotY, plotX, plotY + plotHeight);

        double categoryWidth = plotWidth / categoryCount;
        double barSlot = categoryWidth * 0.82d / Math.Max(1, series.Count);
        for (int category = 0; category < categoryCount; category++)
        {
            double categoryX = plotX + category * categoryWidth + categoryWidth * 0.09d;
            for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
            {
                IReadOnlyList<double> values = series[seriesIndex];
                if (category >= values.Count)
                {
                    continue;
                }

                double value = Math.Max(0d, values[category]);
                double barHeight = value / maxValue * plotHeight;
                RgbColor fill = palette[seriesIndex % palette.Length];
                graphics.SetFillRgb(fill.Red, fill.Green, fill.Blue);
                graphics.FillRectangle(categoryX + seriesIndex * barSlot, plotY, Math.Max(0.5d, barSlot * 0.86d), barHeight);
            }
        }
    }

    private static void EmitChartDiagnostic(Action<OoxPdfDiagnostic>? diagnosticSink, string id, OoxPdfSeverity severity, string message, string? partName, int slideIndex, string fallback)
    {
        diagnosticSink?.Invoke(new OoxPdfDiagnostic(
            id,
            severity,
            message,
            partName,
            SlideIndex: slideIndex,
            Feature: "chart",
            Fallback: fallback));
    }

    private static void AddTableCellTextRuns(XElement cell, double x, double y, double width, double height, PptxTheme theme, List<TextRun> runs)
    {
        XElement? textBody = cell.Element(DrawingNamespace + "txBody");
        if (textBody is null)
        {
            return;
        }

        TextInsets insets = ReadTextInsets(textBody);
        double textAreaHeight = Math.Max(0d, height - insets.Top - insets.Bottom);
        double verticalOffset = ReadTableCellVerticalAnchor(cell) switch
        {
            TextVerticalAnchor.Middle => textAreaHeight / 2d,
            TextVerticalAnchor.Bottom => textAreaHeight,
            _ => 0d
        };
        double firstFontSize = ReadFirstTableCellFontSize(textBody);
        double cursorY = y + height - insets.Top - firstFontSize + 0.54d - verticalOffset;
        var advanceEstimator = new TextAdvanceEstimator();
        foreach (XElement paragraph in textBody.Elements(DrawingNamespace + "p"))
        {
            TextAlignment alignment = ReadAlignment(paragraph, null);
            double cursorX = x + insets.Left;
            double maxFontSize = 12d;
            foreach (XElement run in paragraph.Elements(DrawingNamespace + "r"))
            {
                XElement? runProperties = run.Element(DrawingNamespace + "rPr");
                double fontSize = runProperties?.Attribute("sz") is { } size
                    ? int.Parse(size.Value, CultureInfo.InvariantCulture) / 100d
                    : 12d;
                double alpha = 1d;
                RgbColor color;
                if (TryReadSolidColorWithAlpha(runProperties, theme, out RgbColor runColor, out double runAlpha))
                {
                    color = runColor;
                    alpha = runAlpha;
                }
                else
                {
                    color = new RgbColor(0, 0, 0);
                }
                string? typeface = theme.ResolveTypeface((string?)runProperties?
                    .Element(DrawingNamespace + "latin")
                    ?.Attribute("typeface"));
                bool bold = ParseOptionalBoolAttribute(runProperties, "b");
                bool italic = ParseOptionalBoolAttribute(runProperties, "i");
                bool underline = ((string?)runProperties?.Attribute("u")) is { } underlineValue
                    && !underlineValue.Equals("none", StringComparison.OrdinalIgnoreCase);
                bool strike = IsStrikeEnabled(runProperties, null);
                foreach (TextCapsFragment fragment in ApplyTextCaps((string?)run.Element(DrawingNamespace + "t") ?? string.Empty, runProperties, null))
                {
                    if (fragment.Text.Length == 0)
                    {
                        continue;
                    }

                    double fragmentFontSize = fontSize * fragment.FontScale;
                    maxFontSize = Math.Max(maxFontSize, fragmentFontSize);
                    double advance = advanceEstimator.Measure(fragment.Text, fragmentFontSize, typeface, bold, italic, characterSpacing: 0d);
                    runs.Add(new TextRun(fragment.Text, cursorX, cursorY, Math.Max(1d, advance), textAreaHeight, x, y - height * 0.75d, Math.Max(1d, width), Math.Max(1d, height * 2.1d), fragmentFontSize, 0d, 0d, color, alpha, null, bold, italic, underline, strike, alignment, typeface, 0d, 0d, 0d));
                    cursorX += advance;
                }
            }

            cursorY -= maxFontSize * 1.2d;
        }
    }

    private static double ReadFirstTableCellFontSize(XElement textBody)
    {
        foreach (XElement runProperties in textBody
            .Elements(DrawingNamespace + "p")
            .Elements(DrawingNamespace + "r")
            .Elements(DrawingNamespace + "rPr"))
        {
            if (runProperties.Attribute("sz") is { } size &&
                int.TryParse(size.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int centipoints))
            {
                return Math.Max(1d, centipoints / 100d);
            }
        }

        return 12d;
    }

    private static TextVerticalAnchor ReadTableCellVerticalAnchor(XElement cell)
    {
        string? anchor = (string?)cell
            .Element(DrawingNamespace + "tcPr")
            ?.Attribute("anchor");
        return anchor switch
        {
            "ctr" => TextVerticalAnchor.Middle,
            "b" => TextVerticalAnchor.Bottom,
            _ => TextVerticalAnchor.Top
        };
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
        return TryReadSolidColorWithAlpha(element, theme, out color, out _);
    }

    private static bool TryReadSolidColorWithAlpha(XElement? element, PptxTheme theme, out RgbColor color, out double alpha)
    {
        XElement? solidFill = element?.Element(DrawingNamespace + "solidFill");
        XElement? colorContainer = solidFill ?? element;
        alpha = ReadAlpha(colorContainer);
        XElement? srgbColor = colorContainer?.Element(DrawingNamespace + "srgbClr");
        string? hex = (string?)srgbColor?.Attribute("val");
        if (RgbColor.TryParse(hex, out color))
        {
            color = ApplyColorTransforms(srgbColor, color);
            return true;
        }

        XElement? schemeColorElement = colorContainer?.Element(DrawingNamespace + "schemeClr");
        string? schemeColor = (string?)schemeColorElement?.Attribute("val");
        if (schemeColor is not null && theme.TryResolveColor(schemeColor, out color))
        {
            color = ApplyColorTransforms(schemeColorElement, color);
            return true;
        }

        return false;
    }

    private static double ReadAlpha(XElement? colorContainer)
    {
        XElement? alpha = colorContainer?
            .Elements()
            .FirstOrDefault(element => element.Name.LocalName is "srgbClr" or "schemeClr")
            ?.Element(DrawingNamespace + "alpha");
        if (alpha?.Attribute("val") is { } value &&
            int.TryParse(value.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed))
        {
            return Math.Clamp(parsed / 100000d, 0d, 1d);
        }

        return 1d;
    }

    private static RgbColor ApplyColorTransforms(XElement? colorElement, RgbColor color)
    {
        if (colorElement is null)
        {
            return color;
        }

        double red = color.Red;
        double green = color.Green;
        double blue = color.Blue;
        foreach (XElement transform in colorElement.Elements())
        {
            double value = ParseOptionalLongAttribute(transform, "val", 100000) / 100000d;
            switch (transform.Name.LocalName)
            {
                case "lumMod":
                case "shade":
                    red *= value;
                    green *= value;
                    blue *= value;
                    break;
                case "lumOff":
                    red += 255d * value;
                    green += 255d * value;
                    blue += 255d * value;
                    break;
                case "tint":
                    red += (255d - red) * value;
                    green += (255d - green) * value;
                    blue += (255d - blue) * value;
                    break;
            }
        }

        return new RgbColor(ToByte(red), ToByte(green), ToByte(blue));
    }

    private static byte ToByte(double value)
    {
        return (byte)Math.Clamp((int)Math.Round(value, MidpointRounding.AwayFromZero), 0, 255);
    }

    private static bool TryReadLine(XElement shapeProperties, PptxTheme theme, out RgbColor color, out double lineWidth)
    {
        return TryReadLineWithAlpha(shapeProperties, theme, out color, out lineWidth, out _);
    }

    private static bool TryReadLineWithAlpha(XElement shapeProperties, PptxTheme theme, out RgbColor color, out double lineWidth, out double alpha)
    {
        XElement? line = shapeProperties.Element(DrawingNamespace + "ln");
        lineWidth = line?.Attribute("w") is { } widthAttribute
            ? OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture))
            : 1d;
        return TryReadSolidColorWithAlpha(line, theme, out color, out alpha);
    }

    private static bool TryReadShapeFontColor(XElement shape, PptxTheme theme, out RgbColor color)
    {
        XElement? fontRef = shape
            .Element(PresentationNamespace + "style")
            ?.Element(DrawingNamespace + "fontRef");
        return TryReadSolidColor(fontRef, theme, out color);
    }

    private static IReadOnlyList<TextRun> ReadTextRuns(XDocument slideXml, PptxDocument document, PptxTheme theme, bool includePlaceholders, IReadOnlyList<XDocument> placeholderSources)
    {
        var runs = new List<TextRun>();
        var advanceEstimator = new TextAdvanceEstimator();
        foreach (XElement shape in slideXml.Descendants(PresentationNamespace + "sp"))
        {
            if (!includePlaceholders && IsPlaceholder(shape))
            {
                continue;
            }

            XElement? shapeProperties = shape.Element(PresentationNamespace + "spPr");
            XElement? textBody = shape.Element(PresentationNamespace + "txBody");
            XElement? inheritedPlaceholder = FindInheritedPlaceholderShape(shape, placeholderSources);
            XElement? inheritedTextBody = inheritedPlaceholder?.Element(PresentationNamespace + "txBody");
            ShapeBounds? bounds = shapeProperties is null ? null : ReadBounds(shapeProperties);
            bounds ??= inheritedPlaceholder?.Element(PresentationNamespace + "spPr") is { } inheritedProperties
                ? ReadBounds(inheritedProperties)
                : null;
            if (bounds is null || textBody is null)
            {
                continue;
            }

            bounds = ReadAncestorGroupTransform(shape).Apply(bounds.Value);
            double x = OoxUnits.EmuToPoints(bounds.Value.X);
            double yTop = OoxUnits.EmuToPoints(bounds.Value.Y);
            double width = OoxUnits.EmuToPoints(bounds.Value.Width);
            double height = OoxUnits.EmuToPoints(bounds.Value.Height);
            TextInsets insets = ReadTextInsets(textBody);
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
                FindInheritedTextStyle(shape, placeholderSources));
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
            double cursorLineTop = document.SlideHeightPoints - yTop - insets.Top - verticalOffset;
            int autoNumberValue = 1;

            foreach (XElement paragraph in textBody.Elements(DrawingNamespace + "p"))
            {
                XElement? paragraphProperties = paragraph.Element(DrawingNamespace + "pPr");
                XElement? defaultRunProperties = paragraphProperties?.Element(DrawingNamespace + "defRPr") ??
                    defaultParagraphProperties?.Element(DrawingNamespace + "defRPr");
                double paragraphFontSize = ReadFirstParagraphFontSize(paragraph, defaultRunProperties);
                double spacingBefore = ReadParagraphSpacing(paragraphProperties, defaultParagraphProperties, "spcBef", paragraphFontSize);
                double spacingAfter = ReadParagraphSpacing(paragraphProperties, defaultParagraphProperties, "spcAft", paragraphFontSize);
                LineSpacing lineSpacing = ReadLineSpacing(paragraphProperties, defaultParagraphProperties);
                if (!ParagraphHasVisibleContent(paragraph))
                {
                    if (ParagraphHasLayoutContent(paragraph))
                    {
                        XElement? endRunProperties = paragraph.Element(DrawingNamespace + "endParaRPr");
                        double emptyFontSize = ReadFontSize(endRunProperties, defaultRunProperties);
                        double emptySpacingBefore = ReadParagraphSpacing(paragraphProperties, defaultParagraphProperties, "spcBef", emptyFontSize);
                        double emptySpacingAfter = ReadParagraphSpacing(paragraphProperties, defaultParagraphProperties, "spcAft", emptyFontSize);
                        cursorLineTop -= emptySpacingBefore + ReadParagraphAdvance(lineSpacing, emptyFontSize) + emptySpacingAfter;
                    }

                    continue;
                }

                TextAlignment alignment = ReadAlignment(paragraph, defaultParagraphProperties);
                string? bulletText = ReadBulletText(paragraphProperties, ref autoNumberValue);
                bool bulletPending = bulletText is not null;
                ParagraphIndent indent = ReadParagraphIndent(paragraphProperties);
                IReadOnlyList<double> tabStops = ReadTabStops(paragraphProperties);
                double bulletX = textX + Math.Max(0d, indent.MarginLeft + indent.Hanging);
                double paragraphTextX = bulletText is null
                    ? textX + Math.Max(0d, indent.MarginLeft + indent.Hanging)
                    : textX + Math.Max(0d, indent.MarginLeft);
                cursorLineTop -= spacingBefore;
                double cursorY = cursorLineTop - ReadFirstLineBaselineOffset(paragraph, defaultRunProperties, lineSpacing);
                double cursorX = paragraphTextX;
                double maxFontSize = 18d;
                var paragraphRuns = new List<TextRun>();
                double paragraphEndX = paragraphTextX;
                foreach (XElement child in paragraph.Elements())
                {
                    if (child.Name == DrawingNamespace + "br")
                    {
                        AddAlignedParagraphRuns(runs, paragraphRuns, alignment, textX, textWidth, paragraphEndX);
                        paragraphRuns.Clear();
                        cursorLineTop -= ReadLineAdvance(lineSpacing, maxFontSize);
                        cursorY = double.NaN;
                        cursorX = paragraphTextX;
                        paragraphEndX = paragraphTextX;
                        maxFontSize = 18d;
                        continue;
                    }

                    if (child.Name != DrawingNamespace + "r")
                    {
                        continue;
                    }

                    XElement run = child;
                    XElement? runProperties = run.Element(DrawingNamespace + "rPr");
                    double nominalFontSize = ReadFontSize(runProperties, defaultRunProperties);
                    maxFontSize = Math.Max(maxFontSize, nominalFontSize);
                    double baselineOffset = ReadBaselineOffset(runProperties, defaultRunProperties, nominalFontSize);
                    if (double.IsNaN(cursorY))
                    {
                        cursorY = cursorLineTop - LineBaselineOffset(nominalFontSize, lineSpacing);
                    }

                    double fontSize = Math.Abs(baselineOffset) > 0.001d
                        ? nominalFontSize * 2d / 3d
                        : nominalFontSize;
                    double alpha = 1d;
                    RgbColor color;
                    if (TryReadSolidColorWithAlpha(runProperties, theme, out RgbColor runColor, out double runAlpha))
                    {
                        color = runColor;
                        alpha = runAlpha;
                    }
                    else if (TryReadSolidColorWithAlpha(defaultRunProperties, theme, out RgbColor defaultColor, out double defaultAlpha))
                    {
                        color = defaultColor;
                        alpha = defaultAlpha;
                    }
                    else
                    {
                        color = shapeFontColor ?? new RgbColor(0, 0, 0);
                    }
                    string? typeface = theme.ResolveTypeface((string?)(runProperties?
                        .Element(DrawingNamespace + "latin") ??
                        defaultRunProperties?.Element(DrawingNamespace + "latin"))
                        ?.Attribute("typeface"));
                    bool bold = ParseOptionalBoolAttribute(runProperties, "b") ||
                        (runProperties?.Attribute("b") is null && ParseOptionalBoolAttribute(defaultRunProperties, "b"));
                    bool italic = ParseOptionalBoolAttribute(runProperties, "i") ||
                        (runProperties?.Attribute("i") is null && ParseOptionalBoolAttribute(defaultRunProperties, "i"));
                    bool underline = ((string?)(runProperties?.Attribute("u") ?? defaultRunProperties?.Attribute("u"))) is { } underlineValue
                        && !underlineValue.Equals("none", StringComparison.OrdinalIgnoreCase);
                    bool strike = IsStrikeEnabled(runProperties, defaultRunProperties);
                    double characterSpacing = ReadCharacterSpacing(runProperties, defaultRunProperties);
                    RgbColor? highlight = TryReadHighlightColor(runProperties, out RgbColor highlightColor)
                        ? highlightColor
                        : null;
                    if (bulletPending)
                    {
                        BulletStyle bulletStyle = ReadBulletStyle(paragraphProperties, theme, fontSize, color, typeface);
                        double bulletWidth = Math.Max(1d, textWidth - (bulletX - textX));
                        paragraphRuns.Add(new TextRun(bulletText!, bulletX, cursorY, bulletWidth, textHeight, textX, textClipY, textWidth, textClipHeight, bulletStyle.FontSize, characterSpacing, 0d, bulletStyle.Color, 1d, null, bold, italic, underline, strike, alignment, bulletStyle.Typeface, bounds.Value.RotationDegrees, rotationCenterX, rotationCenterY));
                        paragraphEndX = Math.Max(paragraphEndX, bulletX + advanceEstimator.Measure(bulletText!, bulletStyle.FontSize, bulletStyle.Typeface, bold, italic, characterSpacing));
                        bulletPending = false;
                    }

                    string rawText = (string?)run.Element(DrawingNamespace + "t") ?? string.Empty;
                    string[] tabParts = rawText.Split('\t');
                    for (int tabPartIndex = 0; tabPartIndex < tabParts.Length; tabPartIndex++)
                    {
                        if (tabPartIndex > 0)
                        {
                            cursorX = ResolveNextTabX(cursorX, paragraphTextX, tabStops, fontSize);
                            paragraphEndX = Math.Max(paragraphEndX, cursorX);
                        }

                        foreach (TextCapsFragment fragment in ApplyTextCaps(tabParts[tabPartIndex], runProperties, defaultRunProperties))
                        {
                            if (fragment.Text.Length == 0)
                            {
                                continue;
                            }

                            double fragmentFontSize = fontSize * fragment.FontScale;
                            foreach (string segment in SplitFlowSegments(fragment.Text))
                            {
                                string currentSegment = segment;
                                double segmentWidth = advanceEstimator.Measure(currentSegment, fragmentFontSize, typeface, bold, italic, characterSpacing);
                                bool overflowsLine = cursorX > paragraphTextX &&
                                    (cursorX + segmentWidth > textX + textWidth ||
                                        (characterSpacing > 0d && cursorX + segmentWidth > textX + textWidth - fragmentFontSize));
                                if (overflowsLine)
                                {
                                    AddAlignedParagraphRuns(runs, paragraphRuns, alignment, textX, textWidth, paragraphEndX);
                                    paragraphRuns.Clear();
                                    cursorLineTop -= ReadLineAdvance(lineSpacing, maxFontSize);
                                    cursorY = cursorLineTop - LineBaselineOffset(fragmentFontSize, lineSpacing);
                                    cursorX = paragraphTextX;
                                    paragraphEndX = paragraphTextX;
                                    maxFontSize = Math.Max(nominalFontSize, fragmentFontSize);
                                    currentSegment = currentSegment.TrimStart();
                                    segmentWidth = advanceEstimator.Measure(currentSegment, fragmentFontSize, typeface, bold, italic, characterSpacing);
                                }

                                if (currentSegment.Length == 0)
                                {
                                    continue;
                                }

                                paragraphRuns.Add(new TextRun(currentSegment, cursorX, cursorY, Math.Max(1d, segmentWidth), textHeight, textX, textClipY, textWidth, textClipHeight, fragmentFontSize, characterSpacing, baselineOffset, color, alpha, highlight, bold, italic, underline, strike, alignment, typeface, bounds.Value.RotationDegrees, rotationCenterX, rotationCenterY));
                                cursorX += segmentWidth;
                                paragraphEndX = Math.Max(paragraphEndX, cursorX);
                            }
                        }
                    }
                }

                AddAlignedParagraphRuns(runs, paragraphRuns, alignment, textX, textWidth, paragraphEndX);
                cursorLineTop -= ReadParagraphAdvance(lineSpacing, maxFontSize) + spacingAfter;
            }
        }

        return runs;
    }

    private static GroupTransform ReadAncestorGroupTransform(XElement shape)
    {
        GroupTransform transform = GroupTransform.Identity;
        foreach (XElement group in shape.Ancestors(PresentationNamespace + "grpSp").Reverse())
        {
            transform = transform.Combine(ReadGroupTransform(group));
        }

        return transform;
    }

    private static IReadOnlyList<TextRun> ReadTextRunsForShape(
        XElement shape,
        PptxDocument document,
        PptxTheme theme,
        bool includePlaceholders,
        IReadOnlyList<XDocument> placeholderSources)
    {
        XElement current = new(shape);
        foreach (XElement group in shape.Ancestors(PresentationNamespace + "grpSp"))
        {
            var groupCopy = new XElement(PresentationNamespace + "grpSp");
            if (group.Element(PresentationNamespace + "grpSpPr") is { } properties)
            {
                groupCopy.Add(new XElement(properties));
            }

            groupCopy.Add(current);
            current = groupCopy;
        }

        var slide = new XDocument(
            new XElement(PresentationNamespace + "sld",
                new XElement(PresentationNamespace + "cSld",
                    new XElement(PresentationNamespace + "spTree", current))));
        return ReadTextRuns(slide, document, theme, includePlaceholders, placeholderSources);
    }

    private static XElement? FindInheritedPlaceholderShape(XElement shape, IReadOnlyList<XDocument> placeholderSources)
    {
        XElement? placeholder = shape
            .Element(PresentationNamespace + "nvSpPr")
            ?.Element(PresentationNamespace + "nvPr")
            ?.Element(PresentationNamespace + "ph");
        if (placeholder is null)
        {
            return null;
        }

        string? type = (string?)placeholder.Attribute("type");
        string? index = (string?)placeholder.Attribute("idx");
        foreach (XDocument source in placeholderSources.Reverse())
        {
            foreach (XElement candidate in source.Descendants(PresentationNamespace + "sp"))
            {
                XElement? candidatePlaceholder = candidate
                    .Element(PresentationNamespace + "nvSpPr")
                    ?.Element(PresentationNamespace + "nvPr")
                    ?.Element(PresentationNamespace + "ph");
                if (candidatePlaceholder is null)
                {
                    continue;
                }

                string? candidateType = (string?)candidatePlaceholder.Attribute("type");
                string? candidateIndex = (string?)candidatePlaceholder.Attribute("idx");
                bool indexMatches = index is not null && candidateIndex == index;
                bool typeMatches = index is null && type is not null && candidateType == type;
                if (!indexMatches && !typeMatches)
                {
                    continue;
                }

                return candidate;
            }
        }

        return null;
    }

    private static XElement? FindInheritedTextStyle(XElement shape, IReadOnlyList<XDocument> placeholderSources)
    {
        string? placeholderType = (string?)shape
            .Element(PresentationNamespace + "nvSpPr")
            ?.Element(PresentationNamespace + "nvPr")
            ?.Element(PresentationNamespace + "ph")
            ?.Attribute("type");
        string styleName = placeholderType switch
        {
            "title" or "ctrTitle" => "titleStyle",
            "body" or "subTitle" => "bodyStyle",
            _ => "otherStyle"
        };

        foreach (XDocument source in placeholderSources)
        {
            XElement? style = source.Root?
                .Element(PresentationNamespace + "txStyles")
                ?.Element(PresentationNamespace + styleName);
            XElement? level = style?.Element(DrawingNamespace + "lvl1pPr") ??
                style?.Element(DrawingNamespace + "defPPr");
            if (level is not null)
            {
                return level;
            }
        }

        return null;
    }

    private static IEnumerable<string> SplitFlowSegments(string text)
    {
        int index = 0;
        while (index < text.Length)
        {
            int start = index;
            while (index < text.Length && text[index] == ' ')
            {
                index++;
            }

            while (index < text.Length && text[index] != ' ')
            {
                index++;
            }

            while (index < text.Length && text[index] == ' ')
            {
                index++;
            }

            if (index > start)
            {
                yield return text[start..index];
            }
        }
    }

    private static void AddAlignedParagraphRuns(List<TextRun> runs, List<TextRun> paragraphRuns, TextAlignment alignment, double textX, double textWidth, double paragraphEndX)
    {
        if (paragraphRuns.Count == 0)
        {
            return;
        }

        double paragraphWidth = Math.Max(0d, paragraphEndX - textX);
        double offset = alignment switch
        {
            TextAlignment.Center => Math.Max(0d, textWidth - paragraphWidth) / 2d,
            TextAlignment.Right => Math.Max(0d, textWidth - paragraphWidth),
            _ => 0d
        };

        foreach (TextRun run in paragraphRuns)
        {
            runs.Add(run with
            {
                X = run.X + offset,
                Width = run.Width,
                Alignment = TextAlignment.Left
            });
        }
    }

    private static XElement? MergeParagraphProperties(params XElement?[] sources)
    {
        XElement? merged = null;
        foreach (XElement? source in sources.Reverse())
        {
            if (source is null)
            {
                continue;
            }

            merged = merged is null
                ? new XElement(source)
                : OverlayParagraphProperties(source, merged);
        }

        return merged;
    }

    private static XElement OverlayParagraphProperties(XElement primary, XElement fallback)
    {
        XElement merged = new(primary);
        foreach (XAttribute attribute in fallback.Attributes())
        {
            if (merged.Attribute(attribute.Name) is null)
            {
                merged.Add(new XAttribute(attribute));
            }
        }

        MergeChildElement(merged, fallback, DrawingNamespace + "defRPr");
        return merged;
    }

    private static void MergeChildElement(XElement primaryParent, XElement fallbackParent, XName childName)
    {
        XElement? primary = primaryParent.Element(childName);
        XElement? fallback = fallbackParent.Element(childName);
        if (fallback is null)
        {
            return;
        }

        if (primary is null)
        {
            primaryParent.Add(new XElement(fallback));
            return;
        }

        foreach (XAttribute attribute in fallback.Attributes())
        {
            if (primary.Attribute(attribute.Name) is null)
            {
                primary.Add(new XAttribute(attribute));
            }
        }

        foreach (XElement child in fallback.Elements())
        {
            if (primary.Element(child.Name) is null)
            {
                primary.Add(new XElement(child));
            }
        }
    }

    private static bool ParagraphHasVisibleContent(XElement paragraph)
    {
        return paragraph.Elements().Any(child =>
            child.Name == DrawingNamespace + "r" ||
            child.Name == DrawingNamespace + "br");
    }

    private static bool ParagraphHasLayoutContent(XElement paragraph)
    {
        return paragraph.Element(DrawingNamespace + "pPr") is not null ||
            paragraph.Element(DrawingNamespace + "endParaRPr") is not null;
    }

    private static double ReadFontSize(XElement? runProperties, XElement? defaultRunProperties)
    {
        return (runProperties?.Attribute("sz") ?? defaultRunProperties?.Attribute("sz")) is { } size
            ? int.Parse(size.Value, CultureInfo.InvariantCulture) / 100d
            : 18d;
    }

    private static double ReadCharacterSpacing(XElement? runProperties, XElement? defaultRunProperties)
    {
        return (runProperties?.Attribute("spc") ?? defaultRunProperties?.Attribute("spc")) is { } spacing
            ? int.Parse(spacing.Value, CultureInfo.InvariantCulture) / 100d
            : 0d;
    }

    private static double ReadBaselineOffset(XElement? runProperties, XElement? defaultRunProperties, double fontSize)
    {
        return (runProperties?.Attribute("baseline") ?? defaultRunProperties?.Attribute("baseline")) is { } baseline
            ? fontSize * int.Parse(baseline.Value, CultureInfo.InvariantCulture) / 100000d
            : 0d;
    }

    private static bool IsStrikeEnabled(XElement? runProperties, XElement? defaultRunProperties)
    {
        string? value = (string?)(runProperties?.Attribute("strike") ?? defaultRunProperties?.Attribute("strike"));
        return value is not null && !value.Equals("noStrike", StringComparison.OrdinalIgnoreCase);
    }

    private static IReadOnlyList<TextCapsFragment> ApplyTextCaps(string text, XElement? runProperties, XElement? defaultRunProperties)
    {
        string? value = (string?)(runProperties?.Attribute("cap") ?? defaultRunProperties?.Attribute("cap"));
        if (text.Length == 0)
        {
            return [];
        }

        if (value is "all")
        {
            return [new TextCapsFragment(text.ToUpperInvariant(), 1d)];
        }

        if (value is not "small")
        {
            return [new TextCapsFragment(text, 1d)];
        }

        var fragments = new List<TextCapsFragment>();
        var builder = new StringBuilder();
        bool? currentSmall = null;
        foreach (char character in text)
        {
            bool isSmall = char.IsLetter(character) && char.IsLower(character);
            if (currentSmall is not null && currentSmall != isSmall)
            {
                fragments.Add(new TextCapsFragment(builder.ToString(), currentSmall.Value ? 0.8d : 1d));
                builder.Clear();
            }

            currentSmall = isSmall;
            builder.Append(char.ToUpperInvariant(character));
        }

        if (builder.Length > 0 && currentSmall is not null)
        {
            fragments.Add(new TextCapsFragment(builder.ToString(), currentSmall.Value ? 0.8d : 1d));
        }

        return fragments;
    }

    private static TextInsets ReadTextInsets(XElement textBody)
    {
        XElement? bodyProperties = textBody.Element(DrawingNamespace + "bodyPr");
        return new TextInsets(
            ReadInset(bodyProperties, "lIns", 91440),
            ReadInset(bodyProperties, "rIns", 91440),
            ReadInset(bodyProperties, "tIns", 45720),
            ReadInset(bodyProperties, "bIns", 45720));
    }

    private static bool ClipsVerticalOverflow(XElement textBody)
    {
        string? overflow = (string?)textBody
            .Element(DrawingNamespace + "bodyPr")
            ?.Attribute("vertOverflow");
        return overflow?.Equals("clip", StringComparison.OrdinalIgnoreCase) == true;
    }

    private static double ReadInset(XElement? element, string attributeName, long defaultEmu)
    {
        long emu = element?.Attribute(attributeName) is { } attribute
            ? long.Parse(attribute.Value, CultureInfo.InvariantCulture)
            : defaultEmu;
        return OoxUnits.EmuToPoints(emu);
    }

    private static double ReadParagraphSpacing(
        XElement? paragraphProperties,
        XElement? defaultParagraphProperties,
        string elementName,
        double referenceFontSize)
    {
        XElement? spacing = paragraphProperties?.Element(DrawingNamespace + elementName) ??
            defaultParagraphProperties?.Element(DrawingNamespace + elementName);
        if (spacing?.Element(DrawingNamespace + "spcPts")?.Attribute("val") is { } points)
        {
            return int.Parse(points.Value, CultureInfo.InvariantCulture) / 100d;
        }

        if (spacing?.Element(DrawingNamespace + "spcPct")?.Attribute("val") is { } percent)
        {
            return referenceFontSize * Math.Max(0d, int.Parse(percent.Value, CultureInfo.InvariantCulture) / 100000d);
        }

        return 0d;
    }

    private static ParagraphIndent ReadParagraphIndent(XElement? paragraphProperties)
    {
        return new ParagraphIndent(
            ReadParagraphEmuAttribute(paragraphProperties, "marL"),
            ReadParagraphEmuAttribute(paragraphProperties, "indent"));
    }

    private static double ReadParagraphEmuAttribute(XElement? paragraphProperties, string attributeName)
    {
        return paragraphProperties?.Attribute(attributeName) is { } attribute
            ? OoxUnits.EmuToPoints(long.Parse(attribute.Value, CultureInfo.InvariantCulture))
            : 0d;
    }

    private static IReadOnlyList<double> ReadTabStops(XElement? paragraphProperties)
    {
        if (paragraphProperties?.Element(DrawingNamespace + "tabLst") is not { } tabList)
        {
            return Array.Empty<double>();
        }

        return tabList
            .Elements(DrawingNamespace + "tab")
            .Select(tab => tab.Attribute("pos") is { } position
                ? OoxUnits.EmuToPoints(long.Parse(position.Value, CultureInfo.InvariantCulture))
                : double.NaN)
            .Where(position => !double.IsNaN(position))
            .Order()
            .ToArray();
    }

    private static double ResolveNextTabX(double cursorX, double paragraphTextX, IReadOnlyList<double> tabStops, double fontSize)
    {
        double current = cursorX - paragraphTextX;
        foreach (double tabStop in tabStops)
        {
            if (tabStop > current + 0.01d)
            {
                return paragraphTextX + tabStop;
            }
        }

        return cursorX + fontSize * 2.2d;
    }

    private static LineSpacing ReadLineSpacing(XElement? paragraphProperties, XElement? defaultParagraphProperties)
    {
        XElement? spacing = paragraphProperties?.Element(DrawingNamespace + "lnSpc") ??
            defaultParagraphProperties?.Element(DrawingNamespace + "lnSpc");
        if (spacing?.Element(DrawingNamespace + "spcPts")?.Attribute("val") is { } points)
        {
            return LineSpacing.Absolute(Math.Max(0.1d, int.Parse(points.Value, CultureInfo.InvariantCulture) / 100d));
        }

        if (spacing?.Element(DrawingNamespace + "spcPct")?.Attribute("val") is { } percent)
        {
            return LineSpacing.Multiple(Math.Max(0.1d, int.Parse(percent.Value, CultureInfo.InvariantCulture) / 100000d), true);
        }

        return LineSpacing.Multiple(1d, false);
    }

    private static double ReadParagraphAdvance(LineSpacing lineSpacing, double fontSize)
    {
        return lineSpacing.IsExplicit
            ? lineSpacing.Resolve(fontSize)
            : fontSize * 1.2d;
    }

    private static double ReadLineAdvance(LineSpacing lineSpacing, double fontSize)
    {
        return lineSpacing.IsExplicit
            ? lineSpacing.Resolve(fontSize)
            : fontSize * 1.2d;
    }

    private static double ReadFirstLineBaselineOffset(XElement paragraph, XElement? defaultRunProperties, LineSpacing lineSpacing)
    {
        return LineBaselineOffset(ReadFirstParagraphFontSize(paragraph, defaultRunProperties), lineSpacing);
    }

    private static double ReadFirstParagraphFontSize(XElement paragraph, XElement? defaultRunProperties)
    {
        const double defaultFontSize = 18d;
        foreach (XElement child in paragraph.Elements())
        {
            if (child.Name == DrawingNamespace + "br")
            {
                return defaultFontSize;
            }

            if (child.Name != DrawingNamespace + "r")
            {
                continue;
            }

            XElement? runProperties = child.Element(DrawingNamespace + "rPr");
            if (runProperties?.Attribute("sz") is { } size)
            {
                return int.Parse(size.Value, CultureInfo.InvariantCulture) / 100d;
            }

            if (defaultRunProperties?.Attribute("sz") is { } defaultSize)
            {
                return int.Parse(defaultSize.Value, CultureInfo.InvariantCulture) / 100d;
            }

            return defaultFontSize;
        }

        return defaultFontSize;
    }

    private static double LineBaselineOffset(double fontSize, LineSpacing lineSpacing)
    {
        if (lineSpacing.IsAbsolute)
        {
            return Math.Max(BaselineOffset(fontSize), lineSpacing.Value - fontSize * 0.374d);
        }

        return lineSpacing.IsExplicit
            ? lineSpacing.Resolve(fontSize) - fontSize * 0.234d
            : BaselineOffset(fontSize);
    }

    private static double BaselineOffset(double fontSize)
    {
        const double baselineOffsetFactor = 0.974d;
        return fontSize * baselineOffsetFactor;
    }

    private static string? ReadBulletText(XElement? paragraphProperties, ref int autoNumberValue)
    {
        if (paragraphProperties is null || paragraphProperties.Element(DrawingNamespace + "buNone") is not null)
        {
            return null;
        }

        if ((string?)paragraphProperties.Element(DrawingNamespace + "buChar")?.Attribute("char") is { } bullet)
        {
            return bullet;
        }

        XElement? autoNumber = paragraphProperties.Element(DrawingNamespace + "buAutoNum");
        if (autoNumber is null)
        {
            return null;
        }

        if (autoNumber.Attribute("startAt") is { } startAt &&
            int.TryParse(startAt.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int start) &&
            start > 0)
        {
            autoNumberValue = start;
        }

        string result = FormatAutoNumber(autoNumberValue, (string?)autoNumber.Attribute("type"));
        autoNumberValue++;
        return result;
    }

    private static string FormatAutoNumber(int value, string? type)
    {
        return type switch
        {
            "arabicParenBoth" => $"({value})",
            "arabicParenR" => $"{value})",
            "alphaLcPeriod" => $"{FormatAlphaNumber(value, upper: false)}.",
            "alphaUcPeriod" => $"{FormatAlphaNumber(value, upper: true)}.",
            "alphaLcParenR" => $"{FormatAlphaNumber(value, upper: false)})",
            "alphaUcParenR" => $"{FormatAlphaNumber(value, upper: true)})",
            "romanLcPeriod" => $"{FormatRomanNumber(value, upper: false)}.",
            "romanUcPeriod" => $"{FormatRomanNumber(value, upper: true)}.",
            "romanLcParenR" => $"{FormatRomanNumber(value, upper: false)})",
            "romanUcParenR" => $"{FormatRomanNumber(value, upper: true)})",
            _ => $"{value}."
        };
    }

    private static string FormatAlphaNumber(int value, bool upper)
    {
        if (value <= 0)
        {
            return string.Empty;
        }

        var builder = new StringBuilder();
        int current = value;
        while (current > 0)
        {
            current--;
            char letter = (char)((upper ? 'A' : 'a') + current % 26);
            builder.Insert(0, letter);
            current /= 26;
        }

        return builder.ToString();
    }

    private static string FormatRomanNumber(int value, bool upper)
    {
        if (value <= 0)
        {
            return string.Empty;
        }

        (int Value, string Numeral)[] numerals =
        [
            (1000, "M"),
            (900, "CM"),
            (500, "D"),
            (400, "CD"),
            (100, "C"),
            (90, "XC"),
            (50, "L"),
            (40, "XL"),
            (10, "X"),
            (9, "IX"),
            (5, "V"),
            (4, "IV"),
            (1, "I")
        ];

        var builder = new StringBuilder();
        int current = value;
        foreach ((int numeralValue, string numeral) in numerals)
        {
            while (current >= numeralValue)
            {
                builder.Append(numeral);
                current -= numeralValue;
            }
        }

        string result = builder.ToString();
        return upper ? result : result.ToLowerInvariant();
    }

    private static BulletStyle ReadBulletStyle(XElement? paragraphProperties, PptxTheme theme, double textFontSize, RgbColor textColor, string? textTypeface)
    {
        XElement? bulletFont = FindBulletProperty(paragraphProperties, "buFont");
        XElement? bulletColor = FindBulletProperty(paragraphProperties, "buClr");
        XElement? bulletSizePercent = FindBulletProperty(paragraphProperties, "buSzPct");
        XElement? bulletSizePoints = FindBulletProperty(paragraphProperties, "buSzPts");

        string? typeface = theme.ResolveTypeface((string?)bulletFont?.Attribute("typeface"));
        RgbColor color = bulletColor is not null &&
            TryReadSolidColor(bulletColor, theme, out RgbColor explicitColor)
                ? explicitColor
                : textColor;
        double fontSize = textFontSize;
        if (bulletSizePercent?.Attribute("val") is { } sizePercent)
        {
            fontSize = textFontSize * Math.Max(0.1d, int.Parse(sizePercent.Value, CultureInfo.InvariantCulture) / 100000d);
        }
        else if (bulletSizePoints?.Attribute("val") is { } sizePoints)
        {
            fontSize = Math.Max(0.1d, int.Parse(sizePoints.Value, CultureInfo.InvariantCulture) / 100d);
        }

        return new BulletStyle(fontSize, color, typeface ?? textTypeface);
    }

    private static XElement? FindBulletProperty(XElement? paragraphProperties, string localName)
    {
        if (paragraphProperties is null)
        {
            return null;
        }

        XName propertyName = DrawingNamespace + localName;
        XElement? marker = paragraphProperties
            .Elements()
            .FirstOrDefault(element => element.Name == DrawingNamespace + "buChar" ||
                element.Name == DrawingNamespace + "buAutoNum" ||
                element.Name == DrawingNamespace + "buBlip");
        IEnumerable<XElement> candidates = marker is null
            ? paragraphProperties.Elements()
            : paragraphProperties.Elements().TakeWhile(element => element != marker);
        return candidates.FirstOrDefault(element => element.Name == propertyName);
    }

    private static bool TryReadHighlightColor(XElement? runProperties, out RgbColor color)
    {
        XElement? highlight = runProperties?.Element(DrawingNamespace + "highlight");
        string? hex = (string?)highlight?.Element(DrawingNamespace + "srgbClr")?.Attribute("val");
        return RgbColor.TryParse(hex, out color);
    }

    private static TextVerticalAnchor ReadVerticalAnchor(XElement textBody)
    {
        string? anchor = (string?)textBody
            .Element(DrawingNamespace + "bodyPr")
            ?.Attribute("anchor");
        return anchor switch
        {
            "ctr" => TextVerticalAnchor.Middle,
            "b" => TextVerticalAnchor.Bottom,
            _ => TextVerticalAnchor.Top
        };
    }

    private static double EstimateTextHeight(XElement textBody, XElement? defaultParagraphProperties)
    {
        double height = 0d;
        foreach (XElement paragraph in textBody.Elements(DrawingNamespace + "p"))
        {
            XElement? paragraphProperties = paragraph.Element(DrawingNamespace + "pPr");
            XElement? defaultRunProperties = paragraphProperties?.Element(DrawingNamespace + "defRPr") ??
                defaultParagraphProperties?.Element(DrawingNamespace + "defRPr");
            LineSpacing lineSpacing = ReadLineSpacing(paragraphProperties, defaultParagraphProperties);
            double paragraphFontSize = ReadFirstParagraphFontSize(paragraph, defaultRunProperties);
            height += ReadParagraphSpacing(paragraphProperties, defaultParagraphProperties, "spcBef", paragraphFontSize);
            if (!ParagraphHasVisibleContent(paragraph))
            {
                if (ParagraphHasLayoutContent(paragraph))
                {
                    XElement? endRunProperties = paragraph.Element(DrawingNamespace + "endParaRPr");
                    height += ReadParagraphAdvance(lineSpacing, ReadFontSize(endRunProperties, defaultRunProperties));
                    double emptyFontSize = ReadFontSize(endRunProperties, defaultRunProperties);
                    height += ReadParagraphSpacing(paragraphProperties, defaultParagraphProperties, "spcAft", emptyFontSize);
                }

                continue;
            }

            double maxFontSize = 18d;
            bool hasLineContent = false;
            foreach (XElement child in paragraph.Elements())
            {
                if (child.Name == DrawingNamespace + "br")
                {
                    height += ReadLineAdvance(lineSpacing, maxFontSize);
                    maxFontSize = 18d;
                    hasLineContent = false;
                    continue;
                }

                if (child.Name != DrawingNamespace + "r")
                {
                    continue;
                }

                XElement? runProperties = child.Element(DrawingNamespace + "rPr");
                double fontSize = ReadFontSize(runProperties, defaultRunProperties);
                maxFontSize = Math.Max(maxFontSize, fontSize);
                hasLineContent = true;
            }

            if (hasLineContent || maxFontSize > 0d)
            {
                height += ReadLineAdvance(lineSpacing, maxFontSize);
            }

            height += ReadParagraphSpacing(paragraphProperties, defaultParagraphProperties, "spcAft", paragraphFontSize);
        }

        return height;
    }

    private static bool IsPlaceholder(XElement shape)
    {
        return shape
            .Element(PresentationNamespace + "nvSpPr")
            ?.Element(PresentationNamespace + "nvPr")
            ?.Element(PresentationNamespace + "ph") is not null;
    }

    private static IReadOnlyList<PdfImageResource> RenderPictures(OoxPackage package, string slidePartName, XDocument slideXml, PptxDocument document, PdfGraphicsBuilder graphics, Action<OoxPdfDiagnostic>? diagnosticSink, int slideIndex)
    {
        var relationships = package.GetRelationships(slidePartName)
            .Where(r => !r.IsExternal && r.ResolvedTarget is not null)
            .ToDictionary(r => r.Id, StringComparer.Ordinal);
        var images = new List<PdfImageResource>();
        int index = 1;
        foreach (XElement shapeTree in slideXml.Descendants(PresentationNamespace + "spTree"))
        {
            RenderPictureContainer(shapeTree, relationships, package, document, graphics, diagnosticSink, slideIndex, GroupTransform.Identity, images, ref index);
        }

        return images;
    }

    private static void RenderPictureContainer(
        XElement container,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        OoxPackage package,
        PptxDocument document,
        PdfGraphicsBuilder graphics,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex,
        GroupTransform transform,
        List<PdfImageResource> images,
        ref int index)
    {
        foreach (XElement child in container.Elements())
        {
            if (child.Name == PresentationNamespace + "pic")
            {
                RenderPicture(child, relationships, package, document, graphics, diagnosticSink, slideIndex, transform, images, ref index);
                continue;
            }

            if (child.Name == PresentationNamespace + "grpSp")
            {
                GroupTransform childTransform = transform.Combine(ReadGroupTransform(child));
                RenderPictureContainer(child, relationships, package, document, graphics, diagnosticSink, slideIndex, childTransform, images, ref index);
                continue;
            }
        }
    }

    private static void RenderPicture(
        XElement picture,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        OoxPackage package,
        PptxDocument document,
        PdfGraphicsBuilder graphics,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex,
        GroupTransform transform,
        List<PdfImageResource> images,
        ref int index)
    {
        string? relationshipId = (string?)picture
            .Element(PresentationNamespace + "blipFill")
            ?.Element(DrawingNamespace + "blip")
            ?.Attribute(RelationshipsNamespace + "embed");
        XElement? shapeProperties = picture.Element(PresentationNamespace + "spPr");
        ShapeBounds? bounds = shapeProperties is null ? null : ReadBounds(shapeProperties);
        if (relationshipId is null || bounds is null || !relationships.TryGetValue(relationshipId, out OoxRelationship? relationship) || relationship.ResolvedTarget is null)
        {
            return;
        }

        OoxPart? imagePart = package.GetPart(relationship.ResolvedTarget);
        if (imagePart is null)
        {
            diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                "IMAGE_MISSING_PART",
                OoxPdfSeverity.Error,
                "Referenced image part was missing and the image was ignored.",
                relationship.ResolvedTarget,
                SlideIndex: slideIndex,
                Feature: "image",
                Fallback: "Ignored"));
            return;
        }

        PdfImageXObject? image = CreateImage(imagePart, diagnosticSink, slideIndex);
        if (image is null)
        {
            return;
        }

        ShapeBounds transformedBounds = transform.Apply(bounds.Value);
        string name = "Im" + index++;
        double x = OoxUnits.EmuToPoints(transformedBounds.X);
        double yTop = OoxUnits.EmuToPoints(transformedBounds.Y);
        double width = OoxUnits.EmuToPoints(transformedBounds.Width);
        double height = OoxUnits.EmuToPoints(transformedBounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        CropRect crop = ReadCrop(picture);
        FillRect fillRect = ReadFillRect(picture);
        double imageX = x + fillRect.Left * width;
        double imageY = y + fillRect.Bottom * height;
        double imageWidth = Math.Max(0.001d, width * (1d - fillRect.Left - fillRect.Right));
        double imageHeight = Math.Max(0.001d, height * (1d - fillRect.Top - fillRect.Bottom));
        double alpha = ReadPictureAlpha(picture);
        bool transparent = alpha < 0.999d;
        bool hasTransform = Math.Abs(transformedBounds.RotationDegrees) > 0.001d || transformedBounds.FlipHorizontal || transformedBounds.FlipVertical;
        if (hasTransform)
        {
            graphics.SaveState();
            ApplyShapeTransform(graphics, x, y, width, height, transformedBounds);
        }

        if (transparent)
        {
            graphics.SaveState();
            graphics.SetAlpha(alpha, 1d);
        }

        if (crop.IsEmpty)
        {
            graphics.DrawImage(name, imageX, imageY, imageWidth, imageHeight);
        }
        else
        {
            graphics.DrawImageCropped(name, imageX, imageY, imageWidth, imageHeight, crop.Left, crop.Top, crop.Right, crop.Bottom);
        }

        if (transparent)
        {
            graphics.RestoreState();
        }

        if (hasTransform)
        {
            graphics.RestoreState();
        }

        images.Add(new PdfImageResource(name, image));
    }

    private static PdfImageXObject? CreateImage(OoxPart imagePart, Action<OoxPdfDiagnostic>? diagnosticSink, int slideIndex)
    {
        byte[] bytes = imagePart.Bytes;
        try
        {
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

            diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                "IMAGE_UNSUPPORTED_FORMAT",
                OoxPdfSeverity.Error,
                $"Image '{imagePart.ContentType}' could not be rendered and was ignored: Unsupported image content type.",
                imagePart.Name,
                SlideIndex: slideIndex,
                Feature: imagePart.ContentType,
                Fallback: "Ignored"));
        }
        catch (Exception ex) when (ex is InvalidDataException or NotSupportedException)
        {
            diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                "IMAGE_UNSUPPORTED_FORMAT",
                OoxPdfSeverity.Error,
                $"Image '{imagePart.ContentType}' could not be rendered and was ignored: {ex.Message}",
                imagePart.Name,
                SlideIndex: slideIndex,
                Feature: imagePart.ContentType,
                Fallback: "Ignored"));
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

    private static FillRect ReadFillRect(XElement picture)
    {
        XElement? fillRectangle = picture
            .Element(PresentationNamespace + "blipFill")
            ?.Element(DrawingNamespace + "stretch")
            ?.Element(DrawingNamespace + "fillRect");
        if (fillRectangle is null)
        {
            return default;
        }

        return new FillRect(
            ParsePercentage(fillRectangle, "l"),
            ParsePercentage(fillRectangle, "t"),
            ParsePercentage(fillRectangle, "r"),
            ParsePercentage(fillRectangle, "b"));
    }

    private static double ReadPictureAlpha(XElement picture)
    {
        XElement? blip = picture
            .Element(PresentationNamespace + "blipFill")
            ?.Element(DrawingNamespace + "blip");
        XElement? alphaModFix = blip?.Element(DrawingNamespace + "alphaModFix");
        if (alphaModFix?.Attribute("amt") is { } amount &&
            int.TryParse(amount.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedAmount))
        {
            return Math.Clamp(parsedAmount / 100000d, 0d, 1d);
        }

        return 1d;
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

        textRuns = CoalesceAdjacentTextRuns(textRuns);
        textRuns = CoalesceUnderlineRuns(textRuns);
        RenderedFonts renderedFonts = CreateRenderedFonts(textRuns);
        DrawTextRunsWithFonts(textRuns, graphics, renderedFonts.Fonts);
        return renderedFonts.Resources;
    }

    private static RenderedFonts CreateRenderedFonts(IReadOnlyList<TextRun> textRuns)
    {
        if (textRuns.Count == 0)
        {
            return new RenderedFonts(new Dictionary<string, RenderedFont>(StringComparer.OrdinalIgnoreCase), []);
        }

        textRuns = CoalesceAdjacentTextRuns(textRuns);
        textRuns = CoalesceUnderlineRuns(textRuns);
        var resolver = new WindowsFontResolver();
        var fonts = new Dictionary<string, RenderedFont>(StringComparer.OrdinalIgnoreCase);
        var resources = new List<PdfFontResource>();
        foreach (IGrouping<string, TextRun> group in textRuns.GroupBy(r => FontKey(r), StringComparer.OrdinalIgnoreCase))
        {
            TextRun first = group.First();
            string familyName = string.IsNullOrWhiteSpace(first.FontFamily) ? "Arial" : first.FontFamily!;
            FontResolution resolution = resolver.Resolve(new FontRequest(familyName, first.Bold, first.Italic));
            if (resolution.FontFilePath is null || !File.Exists(resolution.FontFilePath))
            {
                continue;
            }

            OpenTypeFont font = OpenTypeFont.Load(resolution.FontFilePath);
            PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, group.SelectMany(r => r.Text.EnumerateRunes().Select(rune => rune.Value)));
            string resourceName = "F" + (resources.Count + 1).ToString(CultureInfo.InvariantCulture);
            fonts[group.Key] = new RenderedFont(resourceName, embedded, first.Bold && !resolution.Bold, first.Italic && !resolution.Italic);
            resources.Add(new PdfFontResource(resourceName, embedded));
        }

        return new RenderedFonts(fonts, resources);
    }

    private static void DrawTextRunsWithFonts(IReadOnlyList<TextRun> textRuns, PdfGraphicsBuilder graphics, IReadOnlyDictionary<string, RenderedFont> fonts)
    {
        textRuns = CoalesceAdjacentTextRuns(textRuns);
        textRuns = CoalesceUnderlineRuns(textRuns);
        foreach (TextRun run in textRuns)
        {
            if (fonts.TryGetValue(FontKey(run), out RenderedFont rendered))
            {
                DrawWrappedRun(graphics, rendered.ResourceName, rendered.Font, run, rendered.SyntheticBold, rendered.SyntheticItalic);
            }
        }
    }

    private static IReadOnlyList<TextRun> CoalesceAdjacentTextRuns(IReadOnlyList<TextRun> textRuns)
    {
        var coalesced = new List<TextRun>(textRuns.Count);
        foreach (TextRun run in textRuns)
        {
            if (run.Text.Length == 0)
            {
                continue;
            }

            if (coalesced.Count != 0 && CanCoalesceTextRun(coalesced[^1], run))
            {
                TextRun previous = coalesced[^1];
                coalesced[^1] = previous with
                {
                    Text = previous.Text + run.Text,
                    Width = run.X + run.Width - previous.X
                };
            }
            else
            {
                coalesced.Add(run);
            }
        }

        return coalesced;
    }

    private static bool CanCoalesceTextRun(TextRun left, TextRun right)
    {
        return Math.Abs(left.Y - right.Y) < 0.01d &&
            Math.Abs(left.FontSize - right.FontSize) < 0.01d &&
            Math.Abs(left.CharacterSpacing - right.CharacterSpacing) < 0.01d &&
            Math.Abs(left.BaselineOffset - right.BaselineOffset) < 0.01d &&
            Math.Abs(left.RotationDegrees - right.RotationDegrees) < 0.01d &&
            Math.Abs(left.RotationCenterX - right.RotationCenterX) < 0.01d &&
            Math.Abs(left.RotationCenterY - right.RotationCenterY) < 0.01d &&
            Math.Abs(left.ClipX - right.ClipX) < 0.01d &&
            Math.Abs(left.ClipY - right.ClipY) < 0.01d &&
            Math.Abs(left.ClipWidth - right.ClipWidth) < 0.01d &&
            Math.Abs(left.ClipHeight - right.ClipHeight) < 0.01d &&
            left.Color.Equals(right.Color) &&
            Math.Abs(left.Alpha - right.Alpha) < 0.001d &&
            left.HighlightColor.Equals(right.HighlightColor) &&
            left.Bold == right.Bold &&
            left.Italic == right.Italic &&
            left.Underline == right.Underline &&
            left.Strike == right.Strike &&
            left.Alignment == right.Alignment &&
            string.Equals(left.FontFamily, right.FontFamily, StringComparison.OrdinalIgnoreCase) &&
            right.X >= left.X &&
            Math.Abs(right.X - (left.X + left.Width)) < Math.Max(1d, left.FontSize * 0.2d);
    }

    private static IReadOnlyList<TextRun> CoalesceUnderlineRuns(IReadOnlyList<TextRun> textRuns)
    {
        var coalesced = new List<TextRun>(textRuns.Count);
        foreach (TextRun run in textRuns)
        {
            if (coalesced.Count > 0 && CanCoalesceUnderlineRun(coalesced[^1], run))
            {
                TextRun previous = coalesced[^1];
                coalesced[^1] = previous with
                {
                    Text = previous.Text + run.Text,
                    Width = Math.Max(previous.Width, run.X + run.Width - previous.X)
                };
                continue;
            }

            coalesced.Add(run);
        }

        return coalesced;
    }

    private static bool CanCoalesceUnderlineRun(TextRun left, TextRun right)
    {
        return left.Underline &&
            right.Underline &&
            !left.Strike &&
            !right.Strike &&
            left.Bold == right.Bold &&
            left.Italic == right.Italic &&
            left.Alignment == right.Alignment &&
            string.Equals(left.FontFamily, right.FontFamily, StringComparison.OrdinalIgnoreCase) &&
            left.Color.Equals(right.Color) &&
            NearlyEqual(left.Alpha, right.Alpha) &&
            left.HighlightColor.Equals(right.HighlightColor) &&
            NearlyEqual(left.Y, right.Y) &&
            NearlyEqual(left.Height, right.Height) &&
            NearlyEqual(left.ClipX, right.ClipX) &&
            NearlyEqual(left.ClipY, right.ClipY) &&
            NearlyEqual(left.ClipWidth, right.ClipWidth) &&
            NearlyEqual(left.ClipHeight, right.ClipHeight) &&
            NearlyEqual(left.FontSize, right.FontSize) &&
            NearlyEqual(left.CharacterSpacing, right.CharacterSpacing) &&
            NearlyEqual(left.BaselineOffset, right.BaselineOffset) &&
            NearlyEqual(left.RotationDegrees, right.RotationDegrees) &&
            NearlyEqual(left.RotationCenterX, right.RotationCenterX) &&
            NearlyEqual(left.RotationCenterY, right.RotationCenterY) &&
            Math.Abs((left.X + left.Width) - right.X) <= Math.Max(1d, left.FontSize * 0.08d);
    }

    private static bool NearlyEqual(double left, double right)
    {
        return Math.Abs(left - right) <= 0.001d;
    }

    private static string FontKey(TextRun run)
    {
        string familyName = string.IsNullOrWhiteSpace(run.FontFamily) ? "Arial" : run.FontFamily!;
        return familyName + "\u001f" + run.Bold.ToString(CultureInfo.InvariantCulture) + "\u001f" + run.Italic.ToString(CultureInfo.InvariantCulture);
    }

    private static void DrawWrappedRun(PdfGraphicsBuilder graphics, string resourceName, PdfEmbeddedFont embedded, TextRun run, bool syntheticBold, bool syntheticItalic)
    {
        graphics.SaveState();
        if (Math.Abs(run.RotationDegrees) > 0.001d)
        {
            ApplyTextRotation(graphics, run.RotationDegrees, run.RotationCenterX, run.RotationCenterY);
        }

        graphics.ClipRectangle(run.ClipX, run.ClipY, run.ClipWidth, run.ClipHeight);
        double cursorY = run.Y;
        double lineHeight = run.FontSize * 1.2d;
        foreach (string line in WrapWords(run.Text, run.Width, run.FontSize, run.CharacterSpacing, embedded))
        {
            if (cursorY < run.Y - run.Height ||
                cursorY - lineHeight < run.ClipY)
            {
                break;
            }

            string glyphHex = embedded.EncodeGlyphHex(line);
            if (glyphHex.Length != 0)
            {
                double lineWidth = MeasureRenderedText(embedded, line, run.FontSize, run.CharacterSpacing);
                double x = run.Alignment switch
                {
                    TextAlignment.Center => run.X + Math.Max(0, run.Width - lineWidth) / 2d,
                    TextAlignment.Right => run.X + Math.Max(0, run.Width - lineWidth),
                    _ => run.X
                };

                double baselineY = cursorY + run.BaselineOffset;
                if (run.HighlightColor is { } highlight)
                {
                    graphics.SetFillRgb(highlight.Red, highlight.Green, highlight.Blue);
                    graphics.FillRectangle(x, baselineY - run.FontSize * 0.2275d, lineWidth, run.FontSize * 1.119d);
                }

                bool transparentText = run.Alpha < 0.999d;
                if (transparentText)
                {
                    graphics.SaveState();
                    graphics.SetAlpha(run.Alpha, 1d);
                }

                DrawGlyphText(graphics, embedded, resourceName, run.FontSize, x, baselineY, run.Color, line, glyphHex, syntheticItalic, run.CharacterSpacing);
                if (syntheticBold)
                {
                    DrawGlyphText(graphics, embedded, resourceName, run.FontSize, x + 0.35d, baselineY, run.Color, line, glyphHex, syntheticItalic, run.CharacterSpacing);
                }

                if (run.Underline)
                {
                    graphics.SetFillRgb(run.Color.Red, run.Color.Green, run.Color.Blue);
                    double underlineScale = run.FontSize / embedded.Font.UnitsPerEm;
                    double underlineThickness = Math.Max(0.5d, Math.Abs(embedded.Font.Post.UnderlineThickness) * underlineScale);
                    double underlineY = baselineY + (embedded.Font.Post.UnderlinePosition - Math.Abs(embedded.Font.Post.UnderlineThickness)) * underlineScale;
                    graphics.FillRectangle(x, underlineY, lineWidth, underlineThickness);
                }

                if (run.Strike)
                {
                    graphics.SetFillRgb(run.Color.Red, run.Color.Green, run.Color.Blue);
                    graphics.FillRectangle(x, baselineY + run.FontSize * 0.211d, lineWidth, Math.Max(0.5d, run.FontSize * 0.05d));
                }

                if (transparentText)
                {
                    graphics.RestoreState();
                }
            }

            cursorY -= lineHeight;
        }

        graphics.RestoreState();
    }

    private static void ApplyTextRotation(PdfGraphicsBuilder graphics, double rotationDegrees, double centerX, double centerY)
    {
        double radians = -rotationDegrees * Math.PI / 180d;
        double cos = Math.Cos(radians);
        double sin = Math.Sin(radians);
        double e = centerX - cos * centerX + sin * centerY;
        double f = centerY - sin * centerX - cos * centerY;
        graphics.Transform(cos, sin, -sin, cos, e, f);
    }

    private static void DrawGlyphText(PdfGraphicsBuilder graphics, PdfEmbeddedFont embedded, string resourceName, double fontSize, double x, double y, RgbColor color, string text, string glyphHex, bool syntheticItalic, double characterSpacing)
    {
        string? positioningArray = embedded.EncodeGlyphPositioningArray(text, characterSpacing, fontSize, forcePositioningArray: true);
        if (positioningArray is null)
        {
            graphics.DrawGlyphText(resourceName, fontSize, x, y, color.Red, color.Green, color.Blue, glyphHex, syntheticItalic);
        }
        else
        {
            graphics.DrawGlyphPositionedText(resourceName, fontSize, x, y, color.Red, color.Green, color.Blue, positioningArray, syntheticItalic);
        }
    }

    private static IEnumerable<string> WrapWords(string text, double maxWidth, double fontSize, double characterSpacing, PdfEmbeddedFont embedded)
    {
        if (MeasureRenderedText(embedded, text, fontSize, characterSpacing) <= maxWidth)
        {
            yield return text;
            yield break;
        }

        string[] words = text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (words.Length == 0)
        {
            yield break;
        }

        var line = new StringBuilder();
        foreach (string word in words)
        {
            string candidate = line.Length == 0 ? word : line + " " + word;
            if (line.Length > 0 && MeasureRenderedText(embedded, candidate, fontSize, characterSpacing) > maxWidth)
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

    private static double MeasureRenderedText(PdfEmbeddedFont embedded, string text, double fontSize, double characterSpacing)
    {
        double width = embedded.MeasureTextPoints(text, fontSize);
        int runeCount = text.EnumerateRunes().Count();
        return Math.Max(0d, width + Math.Max(0, runeCount - 1) * characterSpacing);
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

    private static TextAlignment ReadAlignment(XElement paragraph, XElement? defaultParagraphProperties)
    {
        string? value = (string?)(paragraph.Element(DrawingNamespace + "pPr")?.Attribute("algn") ??
            defaultParagraphProperties?.Attribute("algn"));
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
        double ClipX,
        double ClipY,
        double ClipWidth,
        double ClipHeight,
        double FontSize,
        double CharacterSpacing,
        double BaselineOffset,
        RgbColor Color,
        double Alpha,
        RgbColor? HighlightColor,
        bool Bold,
        bool Italic,
        bool Underline,
        bool Strike,
        TextAlignment Alignment,
        string? FontFamily,
        double RotationDegrees,
        double RotationCenterX,
        double RotationCenterY);

    private readonly record struct TextCapsFragment(string Text, double FontScale);

    private readonly record struct TextInsets(double Left, double Right, double Top, double Bottom);

    private readonly record struct ParagraphIndent(double MarginLeft, double Hanging);

    private readonly record struct RenderedFont(string ResourceName, PdfEmbeddedFont Font, bool SyntheticBold, bool SyntheticItalic);

    private readonly record struct RenderedFonts(IReadOnlyDictionary<string, RenderedFont> Fonts, IReadOnlyList<PdfFontResource> Resources);

    private readonly record struct BulletStyle(double FontSize, RgbColor Color, string? Typeface);

    private readonly record struct TableBorderLine(double X1, double Y1, double X2, double Y2, double LineWidth, RgbColor Color, double Alpha);

    private readonly record struct TableBorderKey(bool Vertical, double FixedCoordinate, double LineWidth, RgbColor Color, double Alpha)
    {
        public static TableBorderKey From(TableBorderLine border)
        {
            bool vertical = Math.Abs(border.X1 - border.X2) < 0.001d;
            double fixedCoordinate = vertical ? border.X1 : border.Y1;
            return new TableBorderKey(vertical, Math.Round(fixedCoordinate, 3), Math.Round(border.LineWidth, 3), border.Color, Math.Round(border.Alpha, 5));
        }
    }

    private readonly record struct LineSpacing(double Value, bool IsAbsolute, bool IsExplicit)
    {
        public static LineSpacing Absolute(double points) => new(points, true, true);

        public static LineSpacing Multiple(double factor, bool isExplicit) => new(factor, false, isExplicit);

        public double Resolve(double fontSize)
        {
            return IsAbsolute ? Value : fontSize * Value * 1.2d;
        }
    }

    private sealed class TextAdvanceEstimator
    {
        private readonly WindowsFontResolver resolver = new();
        private readonly Dictionary<string, OpenTypeFont?> fonts = new(StringComparer.OrdinalIgnoreCase);

        public double Measure(string text, double fontSize, string? familyName, bool bold = false, bool italic = false, double characterSpacing = 0d)
        {
            OpenTypeFont? font = ResolveFont(string.IsNullOrWhiteSpace(familyName) ? "Arial" : familyName, bold, italic);
            if (font is null)
            {
                int fallbackRuneCount = text.EnumerateRunes().Count();
                return Math.Max(0d, text.Length * fontSize * 0.42d + Math.Max(0, fallbackRuneCount - 1) * characterSpacing);
            }

            double units = 0d;
            ushort previousGlyph = 0;
            foreach (Rune rune in text.EnumerateRunes())
            {
                ushort glyph = font.MapCodePoint(rune.Value);
                if (previousGlyph != 0 && glyph != 0)
                {
                    units += font.GetKerning(previousGlyph, glyph);
                }

                units += font.GetAdvanceWidth(glyph);
                previousGlyph = glyph;
            }

            int runeCount = text.EnumerateRunes().Count();
            return Math.Max(0d, units * fontSize / font.UnitsPerEm + Math.Max(0, runeCount - 1) * characterSpacing);
        }

        private OpenTypeFont? ResolveFont(string familyName, bool bold, bool italic)
        {
            string key = familyName + "\u001f" + bold.ToString(CultureInfo.InvariantCulture) + "\u001f" + italic.ToString(CultureInfo.InvariantCulture);
            if (fonts.TryGetValue(key, out OpenTypeFont? cached))
            {
                return cached;
            }

            try
            {
                FontResolution resolution = resolver.Resolve(new FontRequest(familyName, bold, italic));
                cached = resolution.FontFilePath is null || !File.Exists(resolution.FontFilePath)
                    ? null
                    : OpenTypeFont.Load(resolution.FontFilePath);
            }
            catch (Exception ex) when (ex is IOException or InvalidDataException or NotSupportedException or ArgumentOutOfRangeException)
            {
                cached = null;
            }

            fonts[key] = cached;
            return cached;
        }
    }

    private enum TextAlignment
    {
        Left,
        Center,
        Right
    }

    private enum TextVerticalAnchor
    {
        Top,
        Middle,
        Bottom
    }

    private readonly record struct CropRect(double Left, double Top, double Right, double Bottom)
    {
        public bool IsEmpty => Left == 0d && Top == 0d && Right == 0d && Bottom == 0d;
    }

    private readonly record struct FillRect(double Left, double Top, double Right, double Bottom);
}
