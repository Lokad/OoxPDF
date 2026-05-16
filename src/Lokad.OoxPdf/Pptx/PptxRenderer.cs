using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Diagnostics;
using System.Globalization;
using System.Text;
using System.Xml.Linq;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
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
        PptxTheme theme = PptxTheme.Load(package, document.PresentationPartName);
        var imageCache = new Dictionary<string, PdfImageXObject?>(StringComparer.OrdinalIgnoreCase);
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
            IReadOnlyDictionary<string, OoxRelationship> slideRelationships = package.GetRelationships(slide.PartName)
                .Where(r => !r.IsExternal && r.ResolvedTarget is not null)
                .ToDictionary(r => r.Id, StringComparer.Ordinal);
            var context = new PptxRenderContext(package, document, theme, slide, slideXml, inheritedXml, slideRelationships, imageCache, diagnosticSink);

            foreach (XDocument inherited in context.InheritedXml)
            {
                RenderBackground(context, inherited, graphics);
                RenderShapes(context, inherited, graphics, renderPlaceholders: false);
            }

            RenderBackground(context, context.SlideXml, graphics);
            if (CanRenderSlideInOrder(context.SlideXml))
            {
                var orderedImages = new List<PdfImageResource>();
                int imageIndex = 1;
                IReadOnlyList<TextRun> inheritedTextRuns = ReadInheritedTextRuns(context);
                IReadOnlyList<TextRun> slideTextRuns = ReadSlideTextRuns(context);
                IReadOnlyList<TextRun> slideTableTextRuns = RenderTables(context, context.SlideXml, new PdfGraphicsBuilder());
                RenderedFonts renderedFonts = CreateRenderedFonts(inheritedTextRuns.Concat(slideTextRuns).Concat(slideTableTextRuns).ToArray());
                DrawTextRunsWithFonts(inheritedTextRuns, graphics, renderedFonts.Fonts);
                foreach (XElement shapeTree in context.SlideXml.Descendants(PresentationNamespace + "spTree"))
                {
                    RenderOrderedShapeTextContainer(shapeTree, context, graphics, renderedFonts.Fonts, orderedImages, ref imageIndex, GroupTransform.Identity, renderPlaceholders: true);
                }

                pages.Add(new PdfPage(context.Document.SlideWidthPoints, context.Document.SlideHeightPoints, graphics.ToString(), renderedFonts.Resources, orderedImages, graphics.ExtGStates.ToArray()));
                continue;
            }

            IReadOnlyList<PdfImageResource> images = RenderPictures(context, graphics);
            RenderShapes(context, context.SlideXml, graphics, renderPlaceholders: true);
            IReadOnlyList<TextRun> tableTextRuns = context.InheritedXml
                .Append(context.SlideXml)
                .SelectMany(xml => RenderTables(context, xml, graphics))
                .ToArray();
            RenderCharts(context, graphics);
            IReadOnlyList<TextRun> textRuns = ReadInheritedTextRuns(context)
                .Concat(ReadSlideTextRuns(context))
                .Concat(tableTextRuns)
                .ToArray();
            IReadOnlyList<PdfFontResource> fonts = RenderTextRuns(textRuns, graphics);
            pages.Add(new PdfPage(context.Document.SlideWidthPoints, context.Document.SlideHeightPoints, graphics.ToString(), fonts, images, graphics.ExtGStates.ToArray()));
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

        if (slideXml.Descendants(DrawingNamespace + "bodyPr").Any(HasUnsupportedTextColumns))
        {
            Emit("PPTX_UNSUPPORTED_TEXT_COLUMNS", "multi-column text");
        }

        if (slideXml.Descendants(DrawingNamespace + "bodyPr").Any(HasUnsupportedTextOrientation))
        {
            Emit("PPTX_UNSUPPORTED_TEXT_ORIENTATION", "vertical text");
        }

        if (slideXml.Descendants(PresentationNamespace + "spPr").Any(HasUnsupportedPictureFill))
        {
            Emit("PPTX_UNSUPPORTED_PICTURE_FILL", "picture fill");
        }

        if (slideXml.Descendants().Any(fill =>
                fill.Name.LocalName == "blipFill" &&
                fill.Element(DrawingNamespace + "tile") is not null))
        {
            Emit("PPTX_UNSUPPORTED_IMAGE_TILE", "tiled image fill");
        }

        if (slideXml.Descendants(DrawingNamespace + "blip").Any(blip =>
                blip.Element(DrawingNamespace + "grayscl") is not null ||
                blip.Element(DrawingNamespace + "duotone") is not null ||
                blip.Element(DrawingNamespace + "biLevel") is not null ||
                blip.Element(DrawingNamespace + "lum") is not null))
        {
            Emit("PPTX_UNSUPPORTED_IMAGE_RECOLOR", "image recolor");
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

    private static bool IsTableGraphicFrame(XElement frame)
    {
        return frame
            .Element(DrawingNamespace + "graphic")
            ?.Element(DrawingNamespace + "graphicData")
            ?.Element(DrawingNamespace + "tbl") is not null;
    }

    private static IReadOnlyList<TextRun> RenderTableFrame(PptxRenderContext context, XElement frame, PdfGraphicsBuilder graphics)
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
        double frameTop = context.Document.SlideHeightPoints - frameYTop;
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

                bool hasCellFill = TryReadSolidColorWithAlpha(cellProperties, context.Theme, out RgbColor fill, out double fillAlpha) ||
                    TryReadBuiltInTableStyleCellFill(table, rowIndex, context.Theme, out fill, out fillAlpha);
                if (hasCellFill)
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

                AddTableCellBorders(explicitBorders, cellProperties, context.Theme, cellX, cellBottom, columnWidth, cellHeight);
                RgbColor? tableStyleTextColor = TryReadBuiltInTableStyleTextColor(table, rowIndex, context.Theme, out RgbColor textColor)
                    ? textColor
                    : null;
                AddTableCellTextRuns(cell, cellX, cellBottom, columnWidth, cellHeight, context.Theme, textRuns, tableStyleTextColor);
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

    private static void AddTableCellTextRuns(XElement cell, double x, double y, double width, double height, PptxTheme theme, List<TextRun> runs, RgbColor? tableStyleTextColor = null)
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
            foreach (XElement run in paragraph.Elements().Where(IsTextRunElement))
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
                    color = tableStyleTextColor ?? new RgbColor(0, 0, 0);
                }
                string? typeface = theme.ResolveTypeface((string?)runProperties?
                    .Element(DrawingNamespace + "latin")
                    ?.Attribute("typeface"));
                bool bold = ParseOptionalBoolAttribute(runProperties, "b");
                bool italic = ParseOptionalBoolAttribute(runProperties, "i");
                bool underline = ((string?)runProperties?.Attribute("u")) is { } underlineValue
                    && !underlineValue.Equals("none", StringComparison.OrdinalIgnoreCase);
                bool strike = IsStrikeEnabled(runProperties, null);
                foreach (TextCapsFragment fragment in ApplyTextCaps(ReadTextElementText(run, slideNumber: 0), runProperties, null))
                {
                    if (fragment.Text.Length == 0)
                    {
                        continue;
                    }

                    double fragmentFontSize = fontSize * fragment.FontScale;
                    maxFontSize = Math.Max(maxFontSize, fragmentFontSize);
                    double advance = advanceEstimator.Measure(fragment.Text, fragmentFontSize, typeface, bold, italic, characterSpacing: 0d);
                    runs.Add(new TextRun(fragment.Text, cursorX, cursorY, Math.Max(1d, advance), textAreaHeight, x, y - height * 0.75d, Math.Max(1d, width), Math.Max(1d, height * 2.1d), fragmentFontSize, 0d, 0d, color, alpha, null, bold, italic, underline, strike, true, alignment, typeface, 0d, 0d, 0d));
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
            .Elements()
            .Where(IsTextRunElement)
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

    private static bool TryReadSolidColor(XElement? element, PptxTheme theme, out RgbColor color)
    {
        return TryReadSolidColorWithAlpha(element, theme, out color, out _);
    }

    private static bool TryReadSolidColorWithAlpha(XElement? element, PptxTheme theme, out RgbColor color, out double alpha)
    {
        return TryReadSolidColorWithAlpha(element, theme, placeholderColorContainer: null, out color, out alpha);
    }

    private static bool TryReadSolidColorWithAlpha(XElement? element, PptxTheme theme, XElement? placeholderColorContainer, out RgbColor color, out double alpha)
    {
        XElement? solidFill = element?.Element(DrawingNamespace + "solidFill");
        XElement? colorContainer = element?.Name == DrawingNamespace + "solidFill"
            ? element
            : solidFill ?? element;
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
        if (schemeColor == "phClr" &&
            placeholderColorContainer is not null &&
            TryReadSolidColorWithAlpha(placeholderColorContainer, theme, placeholderColorContainer: null, out color, out double placeholderAlpha))
        {
            color = ApplyColorTransforms(schemeColorElement, color);
            alpha *= placeholderAlpha;
            return true;
        }

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

    private static bool IsPlaceholder(XElement shape)
    {
        return shape
            .Element(PresentationNamespace + "nvSpPr")
            ?.Element(PresentationNamespace + "nvPr")
            ?.Element(PresentationNamespace + "ph") is not null;
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

    private static int ParseOptionalIntAttribute(XElement? element, string name, int defaultValue)
    {
        return element?.Attribute(name) is { } value
            ? int.Parse(value.Value, CultureInfo.InvariantCulture)
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
        bool KerningEnabled,
        TextAlignment Alignment,
        string? FontFamily,
        double RotationDegrees,
        double RotationCenterX,
        double RotationCenterY,
        bool PreventCoalesce = false);

    private readonly record struct TextCapsFragment(string Text, double FontScale);

    private readonly record struct TextFlowSegment(string Text, string AdvanceText, bool Draw, bool PreventCoalesce, double? AdvanceFontSizeFactor = null);

    private readonly record struct ResolvedParagraphTextStyle(
        TextAlignment Alignment,
        XElement? Properties,
        XElement? DefaultRunProperties,
        double FontSize,
        double SpacingBefore,
        double SpacingAfter,
        LineSpacing LineSpacing,
        ParagraphIndent Indent,
        IReadOnlyList<double> TabStops);

    private readonly record struct ResolvedRunTextStyle(
        double NominalFontSize,
        double FontSize,
        double CharacterSpacing,
        double BaselineOffset,
        RgbColor Color,
        double Alpha,
        RgbColor? Highlight,
        bool Bold,
        bool Italic,
        bool Underline,
        bool Strike,
        bool KerningEnabled,
        string? Typeface);

    private sealed class TextLayoutLine(double startX)
    {
        public List<TextRun> Runs { get; } = [];

        public double EndX { get; private set; } = startX;

        public void Add(TextRun run, double endX)
        {
            Runs.Add(run);
            AdvanceTo(endX);
        }

        public void AdvanceTo(double x)
        {
            EndX = Math.Max(EndX, x);
        }

        public void Reset(double startX)
        {
            Runs.Clear();
            EndX = startX;
        }
    }

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

        public double Measure(string text, double fontSize, string? familyName, bool bold = false, bool italic = false, double characterSpacing = 0d, bool kerningEnabled = true)
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
                if (kerningEnabled && previousGlyph != 0 && glyph != 0)
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
                    : OpenTypeFont.Load(resolution.FontFilePath, resolution.FontFaceIndex);
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
