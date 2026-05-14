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
            IReadOnlyList<PdfImageResource> images = RenderPictures(package, slide.PartName, slideXml, document, graphics, diagnosticSink, slideIndex + 1);
            RenderShapes(slideXml, document, graphics, theme, renderPlaceholders: true);
            IReadOnlyList<TextRun> tableTextRuns = inheritedXml
                .Append(slideXml)
                .SelectMany(xml => RenderTables(xml, document, graphics, theme))
                .ToArray();
            RenderCharts(package, slide.PartName, slideXml, document, graphics, diagnosticSink, slideIndex + 1);
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
            else if (preset == "downArrow")
            {
                graphics.FillPolygon(CreateDownArrowPoints(x, y, width, height));
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
            else if (preset == "downArrow")
            {
                graphics.StrokePolygon(CreateDownArrowPoints(x, y, width, height));
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

    private static (double X, double Y)[] CreateDownArrowPoints(double x, double y, double width, double height)
    {
        return
        [
            (x + width * 0.25d, y + height),
            (x + width * 0.75d, y + height),
            (x + width * 0.75d, y + height * 0.45d),
            (x + width, y + height * 0.45d),
            (x + width * 0.5d, y),
            (x, y + height * 0.45d),
            (x + width * 0.25d, y + height * 0.45d)
        ];
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
                runs.Add(new TextRun(text, cursorX, cursorY, Math.Max(1d, width - 8d), Math.Max(1d, height - 8d), x + 4d, y + 4d, Math.Max(1d, width - 8d), Math.Max(1d, height - 8d), fontSize, color, null, bold, italic, underline, alignment, typeface));
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
        XElement? colorContainer = solidFill ?? element;
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
        XElement? line = shapeProperties.Element(DrawingNamespace + "ln");
        lineWidth = line?.Attribute("w") is { } widthAttribute
            ? OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture))
            : 1d;
        return TryReadSolidColor(line, theme, out color);
    }

    private static bool TryReadShapeFontColor(XElement shape, PptxTheme theme, out RgbColor color)
    {
        XElement? fontRef = shape
            .Element(PresentationNamespace + "style")
            ?.Element(DrawingNamespace + "fontRef");
        return TryReadSolidColor(fontRef, theme, out color);
    }

    private static IReadOnlyList<TextRun> ReadTextRuns(XDocument slideXml, PptxDocument document, PptxTheme theme, bool includePlaceholders)
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
            ShapeBounds? bounds = shapeProperties is null ? null : ReadBounds(shapeProperties);
            if (bounds is null || textBody is null)
            {
                continue;
            }

            double x = OoxUnits.EmuToPoints(bounds.Value.X);
            double yTop = OoxUnits.EmuToPoints(bounds.Value.Y);
            double width = OoxUnits.EmuToPoints(bounds.Value.Width);
            double height = OoxUnits.EmuToPoints(bounds.Value.Height);
            TextInsets insets = ReadTextInsets(textBody);
            double textX = x + insets.Left;
            double textWidth = Math.Max(1d, width - insets.Left - insets.Right);
            double textHeight = Math.Max(1d, height - insets.Top - insets.Bottom);
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
            XElement? defaultParagraphProperties = textBody
                .Element(DrawingNamespace + "lstStyle")
                ?.Element(DrawingNamespace + "lvl1pPr");
            double verticalOffset = ReadVerticalAnchor(textBody) switch
            {
                TextVerticalAnchor.Middle => Math.Max(0d, (textHeight - EstimateTextHeight(textBody, defaultParagraphProperties)) / 2d),
                TextVerticalAnchor.Bottom => Math.Max(0d, textHeight - EstimateTextHeight(textBody, defaultParagraphProperties)),
                _ => 0d
            };
            double cursorY = document.SlideHeightPoints - yTop - insets.Top - ReadFirstLineBaselineOffset(textBody) - verticalOffset;

            foreach (XElement paragraph in textBody.Elements(DrawingNamespace + "p"))
            {
                if (!ParagraphHasVisibleContent(paragraph))
                {
                    continue;
                }

                TextAlignment alignment = ReadAlignment(paragraph);
                XElement? paragraphProperties = paragraph.Element(DrawingNamespace + "pPr");
                XElement? defaultRunProperties = paragraphProperties?.Element(DrawingNamespace + "defRPr") ??
                    defaultParagraphProperties?.Element(DrawingNamespace + "defRPr");
                double spacingBefore = ReadParagraphSpacing(paragraphProperties, defaultParagraphProperties, "spcBef");
                double spacingAfter = ReadParagraphSpacing(paragraphProperties, defaultParagraphProperties, "spcAft");
                double lineSpacingFactor = ReadLineSpacingFactor(paragraphProperties, defaultParagraphProperties);
                string? bulletText = ReadBulletText(paragraphProperties);
                string? bulletTypeface = ReadBulletTypeface(paragraphProperties, theme);
                bool bulletPending = bulletText is not null;
                ParagraphIndent indent = ReadParagraphIndent(paragraphProperties);
                double bulletX = textX + Math.Max(0d, indent.MarginLeft + indent.Hanging);
                double paragraphTextX = bulletText is null
                    ? textX + Math.Max(0d, indent.MarginLeft + indent.Hanging)
                    : textX + Math.Max(0d, indent.MarginLeft);
                cursorY -= spacingBefore;
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
                        cursorY -= maxFontSize * lineSpacingFactor;
                        cursorX = paragraphTextX;
                        paragraphEndX = paragraphTextX;
                        maxFontSize = 18d;
                        continue;
                    }

                    if (child.Name == DrawingNamespace + "tab")
                    {
                        cursorX += maxFontSize * 2.2d;
                        continue;
                    }

                    if (child.Name != DrawingNamespace + "r")
                    {
                        continue;
                    }

                    XElement run = child;
                    string text = (string?)run.Element(DrawingNamespace + "t") ?? string.Empty;
                    if (text.Length == 0)
                    {
                        continue;
                    }

                    XElement? runProperties = run.Element(DrawingNamespace + "rPr");
                    double fontSize = (runProperties?.Attribute("sz") ?? defaultRunProperties?.Attribute("sz")) is { } size
                        ? int.Parse(size.Value, CultureInfo.InvariantCulture) / 100d
                        : 18d;
                    maxFontSize = Math.Max(maxFontSize, fontSize);
                    RgbColor color = TryReadSolidColor(runProperties, theme, out RgbColor runColor)
                        ? runColor
                        : TryReadSolidColor(defaultRunProperties, theme, out RgbColor defaultColor)
                            ? defaultColor
                            : shapeFontColor ?? new RgbColor(0, 0, 0);
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
                    RgbColor? highlight = TryReadHighlightColor(runProperties, out RgbColor highlightColor)
                        ? highlightColor
                        : null;
                    if (bulletPending)
                    {
                        double bulletWidth = Math.Max(1d, textWidth - (bulletX - textX));
                        paragraphRuns.Add(new TextRun(bulletText!, bulletX, cursorY, bulletWidth, textHeight, textX, textClipY, textWidth, textClipHeight, fontSize, color, null, bold, italic, underline, alignment, bulletTypeface ?? typeface));
                        paragraphEndX = Math.Max(paragraphEndX, bulletX + advanceEstimator.Measure(bulletText!, fontSize, bulletTypeface ?? typeface));
                        bulletPending = false;
                    }

                    foreach (string segment in SplitFlowSegments(text))
                    {
                        string currentSegment = segment;
                        double segmentWidth = advanceEstimator.Measure(currentSegment, fontSize, typeface);
                        if (cursorX > paragraphTextX && cursorX + segmentWidth > textX + textWidth)
                        {
                            AddAlignedParagraphRuns(runs, paragraphRuns, alignment, textX, textWidth, paragraphEndX);
                            paragraphRuns.Clear();
                            cursorY -= maxFontSize * lineSpacingFactor;
                            cursorX = paragraphTextX;
                            paragraphEndX = paragraphTextX;
                            maxFontSize = fontSize;
                            currentSegment = currentSegment.TrimStart();
                            segmentWidth = advanceEstimator.Measure(currentSegment, fontSize, typeface);
                        }

                        if (currentSegment.Length == 0)
                        {
                            continue;
                        }

                        paragraphRuns.Add(new TextRun(currentSegment, cursorX, cursorY, Math.Max(1d, segmentWidth), textHeight, textX, textClipY, textWidth, textClipHeight, fontSize, color, highlight, bold, italic, underline, alignment, typeface));
                        cursorX += segmentWidth;
                        paragraphEndX = Math.Max(paragraphEndX, cursorX);
                    }
                }

                AddAlignedParagraphRuns(runs, paragraphRuns, alignment, textX, textWidth, paragraphEndX);
                cursorY -= maxFontSize * lineSpacingFactor + spacingAfter;
            }
        }

        return runs;
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

        if (paragraphRuns.Count == 1)
        {
            runs.AddRange(paragraphRuns);
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
                Width = Math.Max(1d, run.Width - offset),
                Alignment = TextAlignment.Left
            });
        }
    }

    private static bool ParagraphHasVisibleContent(XElement paragraph)
    {
        return paragraph.Elements().Any(child =>
            child.Name == DrawingNamespace + "r" ||
            child.Name == DrawingNamespace + "br" ||
            child.Name == DrawingNamespace + "tab");
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

    private static double ReadParagraphSpacing(XElement? paragraphProperties, XElement? defaultParagraphProperties, string elementName)
    {
        XElement? spacing = paragraphProperties?.Element(DrawingNamespace + elementName) ??
            defaultParagraphProperties?.Element(DrawingNamespace + elementName);
        if (spacing?.Element(DrawingNamespace + "spcPts")?.Attribute("val") is { } points)
        {
            return int.Parse(points.Value, CultureInfo.InvariantCulture) / 100d;
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

    private static double ReadLineSpacingFactor(XElement? paragraphProperties, XElement? defaultParagraphProperties)
    {
        XElement? spacing = paragraphProperties?.Element(DrawingNamespace + "lnSpc") ??
            defaultParagraphProperties?.Element(DrawingNamespace + "lnSpc");
        if (spacing?.Element(DrawingNamespace + "spcPct")?.Attribute("val") is { } percent)
        {
            return Math.Max(0.1d, int.Parse(percent.Value, CultureInfo.InvariantCulture) / 100000d);
        }

        return 1d;
    }

    private static double ReadFirstLineBaselineOffset(XElement textBody)
    {
        const double defaultFontSize = 18d;
        foreach (XElement paragraph in textBody.Elements(DrawingNamespace + "p"))
        {
            foreach (XElement child in paragraph.Elements())
            {
                if (child.Name == DrawingNamespace + "br")
                {
                    return defaultFontSize * 0.75d;
                }

                if (child.Name != DrawingNamespace + "r")
                {
                    continue;
                }

                XElement? runProperties = child.Element(DrawingNamespace + "rPr");
                if (runProperties?.Attribute("sz") is { } size)
                {
                    double fontSize = int.Parse(size.Value, CultureInfo.InvariantCulture) / 100d;
                    return fontSize * 0.75d;
                }

                return defaultFontSize * 0.75d;
            }
        }

        return defaultFontSize * 0.75d;
    }

    private static string? ReadBulletText(XElement? paragraphProperties)
    {
        if (paragraphProperties is null || paragraphProperties.Element(DrawingNamespace + "buNone") is not null)
        {
            return null;
        }

        return (string?)paragraphProperties.Element(DrawingNamespace + "buChar")?.Attribute("char");
    }

    private static string? ReadBulletTypeface(XElement? paragraphProperties, PptxTheme theme)
    {
        return theme.ResolveTypeface((string?)paragraphProperties?
            .Element(DrawingNamespace + "buFont")
            ?.Attribute("typeface"));
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
            if (!ParagraphHasVisibleContent(paragraph))
            {
                continue;
            }

            XElement? paragraphProperties = paragraph.Element(DrawingNamespace + "pPr");
            XElement? defaultRunProperties = paragraphProperties?.Element(DrawingNamespace + "defRPr") ??
                defaultParagraphProperties?.Element(DrawingNamespace + "defRPr");
            double lineSpacingFactor = ReadLineSpacingFactor(paragraphProperties, defaultParagraphProperties);
            height += ReadParagraphSpacing(paragraphProperties, defaultParagraphProperties, "spcBef");
            double maxFontSize = 18d;
            bool hasLineContent = false;
            foreach (XElement child in paragraph.Elements())
            {
                if (child.Name == DrawingNamespace + "br")
                {
                    height += maxFontSize * lineSpacingFactor;
                    maxFontSize = 18d;
                    hasLineContent = false;
                    continue;
                }

                if (child.Name != DrawingNamespace + "r")
                {
                    continue;
                }

                XElement? runProperties = child.Element(DrawingNamespace + "rPr");
                double fontSize = (runProperties?.Attribute("sz") ?? defaultRunProperties?.Attribute("sz")) is { } size
                    ? int.Parse(size.Value, CultureInfo.InvariantCulture) / 100d
                    : 18d;
                maxFontSize = Math.Max(maxFontSize, fontSize);
                hasLineContent = true;
            }

            if (hasLineContent || maxFontSize > 0d)
            {
                height += maxFontSize * lineSpacingFactor;
            }

            height += ReadParagraphSpacing(paragraphProperties, defaultParagraphProperties, "spcAft");
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
        foreach (XElement picture in container.Elements(PresentationNamespace + "pic"))
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
                diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                    "IMAGE_MISSING_PART",
                    OoxPdfSeverity.Error,
                    "Referenced image part was missing and the image was ignored.",
                    relationship.ResolvedTarget,
                    SlideIndex: slideIndex,
                    Feature: "image",
                    Fallback: "Ignored"));
                continue;
            }

            PdfImageXObject? image = CreateImage(imagePart, diagnosticSink, slideIndex);
            if (image is null)
            {
                continue;
            }

            ShapeBounds transformedBounds = transform.Apply(bounds.Value);
            string name = "Im" + index++;
            double x = OoxUnits.EmuToPoints(transformedBounds.X);
            double yTop = OoxUnits.EmuToPoints(transformedBounds.Y);
            double width = OoxUnits.EmuToPoints(transformedBounds.Width);
            double height = OoxUnits.EmuToPoints(transformedBounds.Height);
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

        foreach (XElement group in container.Elements(PresentationNamespace + "grpSp"))
        {
            GroupTransform childTransform = transform.Combine(ReadGroupTransform(group));
            RenderPictureContainer(group, relationships, package, document, graphics, diagnosticSink, slideIndex, childTransform, images, ref index);
        }
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

        var resolver = new WindowsFontResolver();
        var fonts = new Dictionary<string, RenderedFont>(StringComparer.OrdinalIgnoreCase);
        var resources = new List<PdfFontResource>();
        foreach (IGrouping<string, TextRun> group in textRuns.GroupBy(r => string.IsNullOrWhiteSpace(r.FontFamily) ? "Arial" : r.FontFamily!, StringComparer.OrdinalIgnoreCase))
        {
            FontResolution resolution = resolver.Resolve(new FontRequest(group.Key));
            if (resolution.FontFilePath is null || !File.Exists(resolution.FontFilePath))
            {
                continue;
            }

            OpenTypeFont font = OpenTypeFont.Load(resolution.FontFilePath);
            PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, group.SelectMany(r => r.Text.EnumerateRunes().Select(rune => rune.Value)));
            string resourceName = "F" + (resources.Count + 1).ToString(CultureInfo.InvariantCulture);
            fonts[group.Key] = new RenderedFont(resourceName, embedded);
            resources.Add(new PdfFontResource(resourceName, embedded));
        }

        foreach (TextRun run in textRuns)
        {
            string familyName = string.IsNullOrWhiteSpace(run.FontFamily) ? "Arial" : run.FontFamily!;
            if (fonts.TryGetValue(familyName, out RenderedFont rendered))
            {
                DrawWrappedRun(graphics, rendered.ResourceName, rendered.Font, run);
            }
        }

        return resources;
    }

    private static void DrawWrappedRun(PdfGraphicsBuilder graphics, string resourceName, PdfEmbeddedFont embedded, TextRun run)
    {
        graphics.SaveState();
        graphics.ClipRectangle(run.ClipX, run.ClipY, run.ClipWidth, run.ClipHeight);
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

                if (run.HighlightColor is { } highlight)
                {
                    graphics.SetFillRgb(highlight.Red, highlight.Green, highlight.Blue);
                    graphics.FillRectangle(x, cursorY - run.FontSize * 0.22d, lineWidth, run.FontSize * 1.05d);
                }

                graphics.DrawGlyphText(resourceName, run.FontSize, x, cursorY, run.Color.Red, run.Color.Green, run.Color.Blue, glyphHex, run.Italic);
                if (run.Bold)
                {
                    graphics.DrawGlyphText(resourceName, run.FontSize, x + 0.35d, cursorY, run.Color.Red, run.Color.Green, run.Color.Blue, glyphHex, run.Italic);
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

        graphics.RestoreState();
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
        double ClipX,
        double ClipY,
        double ClipWidth,
        double ClipHeight,
        double FontSize,
        RgbColor Color,
        RgbColor? HighlightColor,
        bool Bold,
        bool Italic,
        bool Underline,
        TextAlignment Alignment,
        string? FontFamily);

    private readonly record struct TextInsets(double Left, double Right, double Top, double Bottom);

    private readonly record struct ParagraphIndent(double MarginLeft, double Hanging);

    private readonly record struct RenderedFont(string ResourceName, PdfEmbeddedFont Font);

    private sealed class TextAdvanceEstimator
    {
        private readonly WindowsFontResolver resolver = new();
        private readonly Dictionary<string, OpenTypeFont?> fonts = new(StringComparer.OrdinalIgnoreCase);

        public double Measure(string text, double fontSize, string? familyName)
        {
            OpenTypeFont? font = ResolveFont(string.IsNullOrWhiteSpace(familyName) ? "Arial" : familyName);
            if (font is null)
            {
                return text.Length * fontSize * 0.42d;
            }

            double units = 0d;
            foreach (Rune rune in text.EnumerateRunes())
            {
                ushort glyph = font.MapCodePoint(rune.Value);
                units += font.GetAdvanceWidth(glyph);
            }

            return units * fontSize / font.UnitsPerEm;
        }

        private OpenTypeFont? ResolveFont(string familyName)
        {
            if (fonts.TryGetValue(familyName, out OpenTypeFont? cached))
            {
                return cached;
            }

            try
            {
                FontResolution resolution = resolver.Resolve(new FontRequest(familyName));
                cached = resolution.FontFilePath is null || !File.Exists(resolution.FontFilePath)
                    ? null
                    : OpenTypeFont.Load(resolution.FontFilePath);
            }
            catch (Exception ex) when (ex is IOException or InvalidDataException or NotSupportedException or ArgumentOutOfRangeException)
            {
                cached = null;
            }

            fonts[familyName] = cached;
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
}
