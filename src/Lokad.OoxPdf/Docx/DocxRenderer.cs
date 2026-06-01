using System.Globalization;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed class DocxRenderer
{
    internal const string DefaultDocumentTypefaceRequest = "OOXPDF_DOCUMENT_DEFAULT";
    private readonly IFontResolver fontResolver;

    public DocxRenderer(IFontResolver? fontResolver = null)
    {
        this.fontResolver = fontResolver ?? new WindowsFontResolver();
    }

    public IReadOnlyList<PdfPage> RenderBlankPages(DocxDocument document, Action<OoxPdfDiagnostic>? diagnosticSink = null)
    {
        if (document.BodyElements.Count == 0 && document.HeaderParagraphs.Count == 0 && document.FooterParagraphs.Count == 0)
        {
            return [new PdfPage(document.PageWidthPoints, document.PageHeightPoints)];
        }

        return RenderParagraphs(document, fontResolver, diagnosticSink);
    }

    internal DocxLayoutSnapshot InspectLayout(DocxDocument document)
    {
        DocxFontResources fontResources = PrepareFontResources(document, fontResolver);
        DocxLayout layout = new DocxLayoutEngine().Create(document, fontResources.TextMeasurer);
        return DocxLayoutSnapshot.FromLayout(layout);
    }

    internal DocxFontPlanSnapshot InspectFontPlan(DocxDocument document)
    {
        return DocxFontPlanSnapshot.FromPlan(DocxFontPlan.Create(document, fontResolver));
    }

    private static IReadOnlyList<PdfPage> RenderParagraphs(DocxDocument document, IFontResolver fontResolver, Action<OoxPdfDiagnostic>? diagnosticSink)
    {
        DocxFontResources fontResources = PrepareFontResources(document, fontResolver);

        DocxLayout layout = new DocxLayoutEngine().Create(document, fontResources.TextMeasurer);
        var pages = new List<PdfPage>(layout.Pages.Count);
        int imageIndex = 1;

        double width = Math.Max(1d, document.PageWidthPoints - document.MarginLeftPoints - document.MarginRightPoints);
        for (int pageIndex = 0; pageIndex < layout.Pages.Count; pageIndex++)
        {
            DocxLayoutPage layoutPage = layout.Pages[pageIndex];
            var graphics = new PdfGraphicsBuilder();
            var pageImages = new List<PdfImageResource>();
            for (int itemIndex = 0; itemIndex < layoutPage.Items.Count; itemIndex++)
            {
                DocxLayoutItem item = layoutPage.Items[itemIndex];
                DocxTableRowLayout? previousRow = itemIndex > 0 ? layoutPage.Items[itemIndex - 1] as DocxTableRowLayout : null;
                DocxTableRowLayout? nextRow = itemIndex + 1 < layoutPage.Items.Count ? layoutPage.Items[itemIndex + 1] as DocxTableRowLayout : null;
                RenderLayoutItem(item, previousRow, nextRow, graphics, pageImages, fontResources, diagnosticSink, ref imageIndex);
            }

            int pageNumber = pageIndex + 1;
            RenderStaticParagraphs(
                SelectStaticHeaderFooter(document.HeaderParagraphsByType, document.HeaderParagraphs, document.PageSettings, pageNumber),
                graphics,
                fontResources,
                document.MarginLeftPoints,
                width,
                document.PageHeightPoints - ResolveHeaderDistance(document),
                pageNumber);
            RenderStaticParagraphs(
                SelectStaticHeaderFooter(document.FooterParagraphsByType, document.FooterParagraphs, document.PageSettings, pageNumber),
                graphics,
                fontResources,
                document.MarginLeftPoints,
                width,
                ResolveFooterDistance(document),
                pageNumber);

            pages.Add(new PdfPage(layoutPage.Width, layoutPage.Height, graphics.ToString(), fontResources.Resources, pageImages.ToArray()));
        }

        return pages;
    }

    private static DocxFontResources PrepareFontResources(DocxDocument document, IFontResolver fontResolver)
    {
        DocxFontPlan plan = DocxFontPlan.Create(document, fontResolver);
        var resources = new List<PdfFontResource>();
        var runResources = new Dictionary<DocxTextRun, DocxRunFontResource>();
        PrepareResolvedRunFontResources(plan, resources, runResources);
        DocxRunFontResource? fallback = PrepareFallbackFontResource(plan, fontResolver, resources, runResources);
        IDocxTextMeasurer? measurer = plan.Runs.Any(run => run.Resolution?.FontFilePath is not null && File.Exists(run.Resolution.FontFilePath)) || fallback is not null
            ? new DocxFontPlanTextMeasurer(plan, fallback?.Resolution)
            : null;
        return new DocxFontResources(plan, measurer, resources, runResources, fallback);
    }

    private static void PrepareResolvedRunFontResources(
        DocxFontPlan plan,
        List<PdfFontResource> resources,
        Dictionary<DocxTextRun, DocxRunFontResource> runResources)
    {
        foreach (IGrouping<(string Path, int FaceIndex), DocxResolvedRunTypeface> group in plan.Runs
            .Where(run => run.Resolution?.FontFilePath is not null && File.Exists(run.Resolution.FontFilePath))
            .GroupBy(run => (run.Resolution!.FontFilePath!, run.Resolution.FontFaceIndex)))
        {
            FontResolution resolution = group.First().Resolution!;
            IReadOnlyList<int> glyphs = CollectRunGlyphs(group);
            if (glyphs.Count == 0)
            {
                continue;
            }

            OpenTypeFont font = OpenTypeFont.Load(resolution.FontFilePath!, resolution.FontFaceIndex);
            PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, glyphs);
            string name = "F" + (resources.Count + 1).ToString(CultureInfo.InvariantCulture);
            var runResource = new DocxRunFontResource(name, embedded, resolution);
            resources.Add(new PdfFontResource(name, embedded));
            foreach (DocxResolvedRunTypeface run in group)
            {
                runResources[run.Run] = runResource;
            }
        }
    }

    private static DocxRunFontResource? PrepareFallbackFontResource(
        DocxFontPlan plan,
        IFontResolver fontResolver,
        List<PdfFontResource> resources,
        Dictionary<DocxTextRun, DocxRunFontResource> runResources)
    {
        FontResolution resolution = ResolveDocumentBaseFont(plan, fontResolver);
        if (resolution.FontFilePath is null || !File.Exists(resolution.FontFilePath))
        {
            return null;
        }

        DocxResolvedRunTypeface[] fallbackRuns = plan.Runs
            .Where(run => !runResources.ContainsKey(run.Run))
            .ToArray();
        if (fallbackRuns.Length == 0)
        {
            return null;
        }

        IReadOnlyList<int> glyphs = CollectRunGlyphs(fallbackRuns);
        if (glyphs.Count == 0)
        {
            return null;
        }

        OpenTypeFont font = OpenTypeFont.Load(resolution.FontFilePath, resolution.FontFaceIndex);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, glyphs);
        string name = "F" + (resources.Count + 1).ToString(CultureInfo.InvariantCulture);
        var runResource = new DocxRunFontResource(name, embedded, resolution);
        resources.Add(new PdfFontResource(name, embedded));
        foreach (DocxResolvedRunTypeface run in fallbackRuns)
        {
            runResources[run.Run] = runResource;
        }

        return runResource;
    }

    private static IReadOnlyList<int> CollectRunGlyphs(IEnumerable<DocxResolvedRunTypeface> runs)
    {
        return runs
            .SelectMany(run => run.Run.Text.EnumerateRunes().Select(rune => rune.Value))
            .Concat("0123456789".EnumerateRunes().Select(rune => rune.Value))
            .Distinct()
            .ToArray();
    }

    private static FontResolution ResolveDocumentBaseFont(DocxFontPlan plan, IFontResolver fontResolver)
    {
        foreach (DocxResolvedRunTypeface run in plan.Runs)
        {
            if (run.Resolution is { FontFilePath: not null } resolution)
            {
                return resolution;
            }
        }

        return fontResolver.Resolve(new FontRequest(DefaultDocumentTypefaceRequest));
    }

    private static void RenderLayoutItem(
        DocxLayoutItem item,
        DocxTableRowLayout? previousRow,
        DocxTableRowLayout? nextRow,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        DocxFontResources fontResources,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        ref int imageIndex)
    {
        switch (item)
        {
            case DocxTextLineLayout textLine:
                RenderTextLine(textLine, graphics, fontResources);
                break;
            case DocxInlineImageLayout image:
                RenderInlineImage(image, graphics, pageImages, diagnosticSink, ref imageIndex);
                break;
            case DocxTableRowLayout row:
                RenderTableRow(row, IsAdjacentTableRow(previousRow, row) ? previousRow : null, IsAdjacentTableRow(row, nextRow) ? nextRow : null, graphics, pageImages, fontResources, diagnosticSink, ref imageIndex);
                break;
        }
    }

    private static bool IsAdjacentTableRow(DocxTableRowLayout? first, DocxTableRowLayout? second)
    {
        return first is not null &&
            second is not null &&
            first.Table.TableIndex == second.Table.TableIndex &&
            first.RowIndex + 1 == second.RowIndex;
    }

    private static void RenderTextLine(DocxTextLineLayout line, PdfGraphicsBuilder graphics, DocxFontResources fontResources)
    {
        IReadOnlyList<DocxTextSegmentLayout> segments = line.Segments.Count == 0
            ? [new DocxTextSegmentLayout(line.Text, line.StyleRun, line.X, line.Width)]
            : line.Segments;
        foreach (DocxTextSegmentLayout segment in segments)
        {
            RenderTextSegment(segment, line.FontSize, line.BaselineY, graphics, fontResources);
        }
    }

    private static void RenderTextSegment(
        DocxTextSegmentLayout segment,
        double fontSize,
        double baselineY,
        PdfGraphicsBuilder graphics,
        DocxFontResources fontResources)
    {
        DocxRunFontResource? resource = ResolveFontResource(segment.StyleRun, fontResources);
        if (resource is null)
        {
            return;
        }

        DocxTextRun style = segment.StyleRun;
        RgbColor color = ReadColor(style.ColorHex);
        string glyphHex = resource.Embedded.EncodeGlyphHex(segment.Text);
        graphics.DrawGlyphText(resource.Name, fontSize, segment.X, baselineY, color.Red, color.Green, color.Blue, glyphHex, style.Italic);
        if (style.Bold)
        {
            graphics.DrawGlyphText(resource.Name, fontSize, segment.X + 0.35d, baselineY, color.Red, color.Green, color.Blue, glyphHex, style.Italic);
        }

        if (style.Underline)
        {
            graphics.SetStrokeRgb(color.Red, color.Green, color.Blue);
            graphics.SetLineWidth(Math.Max(0.5d, fontSize / 18d));
            graphics.StrokeLine(segment.X, baselineY - fontSize * 0.12d, segment.X + segment.Width, baselineY - fontSize * 0.12d);
        }
    }

    private static DocxRunFontResource? ResolveFontResource(DocxTextRun run, DocxFontResources fontResources)
    {
        return fontResources.RunResources.TryGetValue(run, out DocxRunFontResource? resource)
            ? resource
            : fontResources.Fallback;
    }

    private static void RenderInlineImage(
        DocxInlineImageLayout image,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        ref int imageIndex)
    {
        PdfImageXObject? xObject = CreateImage(image.Image, diagnosticSink, image.PageIndex);
        if (xObject is null)
        {
            return;
        }

        string imageName = "Im" + imageIndex++;
        graphics.DrawImage(imageName, image.X, image.Y, image.Width, image.Height);
        pageImages.Add(new PdfImageResource(imageName, xObject));
    }

    private static void RenderTableRow(
        DocxTableRowLayout row,
        DocxTableRowLayout? previousRow,
        DocxTableRowLayout? nextRow,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        DocxFontResources fontResources,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        ref int imageIndex)
    {
        foreach (DocxTableCellLayout cellLayout in row.Cells)
        {
            if (cellLayout.IsVerticalMergeContinuation)
            {
                continue;
            }

            DocxTableCell cell = cellLayout.Cell;
            if (RgbColor.TryParse(cell.FillHex, out RgbColor fill))
            {
                graphics.SetFillRgb(fill.Red, fill.Green, fill.Blue);
                graphics.FillRectangle(cellLayout.X, cellLayout.Y, cellLayout.Width, cellLayout.Height);
            }
        }

        RenderTableRowBorders(row, previousRow, nextRow, graphics);

        foreach (DocxTableCellLayout cellLayout in row.Cells)
        {
            if (cellLayout.IsVerticalMergeContinuation)
            {
                continue;
            }

            if (cellLayout.TextLines.Count != 0 || cellLayout.InlineImages.Count != 0)
            {
                graphics.SaveState();
                graphics.ClipRectangle(cellLayout.X, cellLayout.Y, cellLayout.Width, cellLayout.Height);
                foreach (DocxTextLineLayout line in cellLayout.TextLines)
                {
                    RenderTextLine(line, graphics, fontResources);
                }

                foreach (DocxInlineImageLayout image in cellLayout.InlineImages)
                {
                    RenderInlineImage(image, graphics, pageImages, diagnosticSink, ref imageIndex);
                }

                graphics.RestoreState();
            }
        }
    }

    private static void RenderTableRowBorders(
        DocxTableRowLayout row,
        DocxTableRowLayout? previousRow,
        DocxTableRowLayout? nextRow,
        PdfGraphicsBuilder graphics)
    {
        for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
        {
            DocxTableCellLayout cellLayout = row.Cells[cellIndex];
            if (cellLayout.IsVerticalMergeContinuation)
            {
                continue;
            }

            if (previousRow is null)
            {
                RenderHorizontalTableCellBorder(cellLayout, "top", graphics);
            }

            if (nextRow is null)
            {
                RenderHorizontalTableCellBorder(cellLayout, "bottom", graphics);
            }

            DocxTableCellBorder? left = FindBorder(cellLayout.Cell.Borders, "left") ?? FindBorder(cellLayout.Cell.Borders, "start");
            if (cellIndex == 0)
            {
                RenderVerticalTableCellBorder(cellLayout.X, cellLayout.Y, cellLayout.Height, left, graphics);
            }

            DocxTableCellBorder? right = FindBorder(cellLayout.Cell.Borders, "right") ?? FindBorder(cellLayout.Cell.Borders, "end");
            if (cellIndex == row.Cells.Count - 1)
            {
                RenderVerticalTableCellBorder(cellLayout.X + cellLayout.Width, cellLayout.Y, cellLayout.Height, right, graphics, alignInsideLeft: true);
                continue;
            }

            DocxTableCellLayout nextCell = row.Cells[cellIndex + 1];
            DocxTableCellBorder? nextLeft = FindBorder(nextCell.Cell.Borders, "left") ?? FindBorder(nextCell.Cell.Borders, "start");
            RenderSharedVerticalTableBorder(cellLayout.X + cellLayout.Width, cellLayout.Y, cellLayout.Height, right, nextLeft, graphics);
        }

        if (nextRow is not null)
        {
            RenderSharedHorizontalTableBorders(row, nextRow, graphics);
        }
    }

    private static void RenderSharedHorizontalTableBorders(
        DocxTableRowLayout row,
        DocxTableRowLayout nextRow,
        PdfGraphicsBuilder graphics)
    {
        foreach (DocxTableCellLayout cellLayout in row.Cells)
        {
            if (cellLayout.IsVerticalMergeContinuation)
            {
                continue;
            }

            DocxTableCellLayout[] overlappingNextCells = nextRow.Cells
                .Where(nextCell => !nextCell.IsVerticalMergeContinuation && HorizontalOverlap(cellLayout, nextCell) > 0d)
                .ToArray();
            if (overlappingNextCells.Length == 0)
            {
                RenderHorizontalTableCellBorder(cellLayout, "bottom", graphics);
                continue;
            }

            foreach (DocxTableCellLayout nextRowCell in overlappingNextCells)
            {
                RenderSharedHorizontalTableBorderSegment(cellLayout, nextRowCell, graphics);
            }
        }
    }

    private static void RenderSharedHorizontalTableBorderSegment(
        DocxTableCellLayout cellLayout,
        DocxTableCellLayout nextRowCell,
        PdfGraphicsBuilder graphics)
    {
        DocxTableCellBorder? bottom = FindBorder(cellLayout.Cell.Borders, "bottom");
        DocxTableCellBorder? nextTop = FindBorder(nextRowCell.Cell.Borders, "top");
        if (IsSuppressedBorder(bottom) || IsSuppressedBorder(nextTop))
        {
            return;
        }

        DocxTableCellBorder? border = SelectStrongerBorder(bottom, nextTop);
        if (border is null)
        {
            return;
        }

        RgbColor color = ReadColor(border.Color);
        graphics.SetFillRgb(color.Red, color.Green, color.Blue);
        double width = ReadBorderWidth(border.SizeValue);
        double x = Math.Max(cellLayout.X, nextRowCell.X);
        double right = Math.Min(cellLayout.X + cellLayout.Width, nextRowCell.X + nextRowCell.Width);
        if (right <= x)
        {
            return;
        }

        graphics.FillRectangle(x, cellLayout.Y - width / 2d, right - x, width);
    }

    private static double HorizontalOverlap(DocxTableCellLayout first, DocxTableCellLayout second)
    {
        return Math.Min(first.X + first.Width, second.X + second.Width) - Math.Max(first.X, second.X);
    }

    private static void RenderHorizontalTableCellBorder(DocxTableCellLayout cellLayout, string edge, PdfGraphicsBuilder graphics)
    {
        DocxTableCellBorder? border = FindBorder(cellLayout.Cell.Borders, edge);
        if (border is null || IsSuppressedBorder(border))
        {
            return;
        }

        RgbColor color = ReadColor(border.Color);
        graphics.SetFillRgb(color.Red, color.Green, color.Blue);
        double width = ReadBorderWidth(border.SizeValue);
        switch (edge)
        {
            case "top":
                graphics.FillRectangle(cellLayout.X, cellLayout.Y + cellLayout.Height - width, cellLayout.Width, width);
                break;
            case "bottom":
                graphics.FillRectangle(cellLayout.X, cellLayout.Y, cellLayout.Width, width);
                break;
        }
    }

    private static void RenderSharedVerticalTableBorder(
        double boundaryX,
        double y,
        double height,
        DocxTableCellBorder? leftCellRight,
        DocxTableCellBorder? rightCellLeft,
        PdfGraphicsBuilder graphics)
    {
        if (IsSuppressedBorder(leftCellRight) || IsSuppressedBorder(rightCellLeft))
        {
            return;
        }

        DocxTableCellBorder? border = SelectStrongerBorder(leftCellRight, rightCellLeft);
        if (border is null)
        {
            return;
        }

        RgbColor color = ReadColor(border.Color);
        graphics.SetFillRgb(color.Red, color.Green, color.Blue);
        double width = ReadBorderWidth(border.SizeValue);
        graphics.FillRectangle(boundaryX - width / 2d, y, width, height);
    }

    private static void RenderVerticalTableCellBorder(
        double boundaryX,
        double y,
        double height,
        DocxTableCellBorder? border,
        PdfGraphicsBuilder graphics,
        bool alignInsideLeft = false)
    {
        if (border is null || IsSuppressedBorder(border))
        {
            return;
        }

        RgbColor color = ReadColor(border.Color);
        graphics.SetFillRgb(color.Red, color.Green, color.Blue);
        double width = ReadBorderWidth(border.SizeValue);
        graphics.FillRectangle(alignInsideLeft ? boundaryX - width : boundaryX, y, width, height);
    }

    private static DocxTableCellBorder? SelectStrongerBorder(DocxTableCellBorder? first, DocxTableCellBorder? second)
    {
        if (first is null)
        {
            return second;
        }

        if (second is null)
        {
            return first;
        }

        return ReadBorderWidth(second.SizeValue) > ReadBorderWidth(first.SizeValue) ? second : first;
    }

    private static DocxTableCellBorder? FindBorder(IReadOnlyList<DocxTableCellBorder> borders, string edge)
    {
        return borders.FirstOrDefault(border => string.Equals(border.Edge, edge, StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsSuppressedBorder(DocxTableCellBorder? border)
    {
        return border is not null &&
            (string.Equals(border.Value, "nil", StringComparison.OrdinalIgnoreCase) ||
             string.Equals(border.Value, "none", StringComparison.OrdinalIgnoreCase));
    }

    private static double ReadBorderWidth(string? sizeValue)
    {
        return int.TryParse(sizeValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int eighths)
            ? Math.Max(0.25d, eighths / 8d)
            : 0.75d;
    }

    private static IReadOnlyList<DocxParagraph> SelectStaticHeaderFooter(
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> paragraphsByType,
        IReadOnlyList<DocxParagraph> fallbackParagraphs,
        DocxPageSettings settings,
        int pageNumber)
    {
        if (settings.TitlePage == true &&
            pageNumber == 1 &&
            paragraphsByType.TryGetValue("first", out IReadOnlyList<DocxParagraph>? first))
        {
            return first;
        }

        if (settings.EvenAndOddHeaders == true &&
            pageNumber % 2 == 0 &&
            paragraphsByType.TryGetValue("even", out IReadOnlyList<DocxParagraph>? even))
        {
            return even;
        }

        return paragraphsByType.TryGetValue("default", out IReadOnlyList<DocxParagraph>? defaultParagraphs)
            ? defaultParagraphs
            : fallbackParagraphs;
    }

    private static double ResolveHeaderDistance(DocxDocument document)
    {
        return document.PageSettings.HeaderDistancePoints ?? Math.Max(18d, document.MarginTopPoints / 2d);
    }

    private static double ResolveFooterDistance(DocxDocument document)
    {
        return document.PageSettings.FooterDistancePoints ?? Math.Max(18d, document.MarginBottomPoints / 2d);
    }

    private static void RenderStaticParagraphs(
        IReadOnlyList<DocxParagraph> paragraphs,
        PdfGraphicsBuilder graphics,
        DocxFontResources fontResources,
        double x,
        double width,
        double startY,
        int pageNumber)
    {
        double cursorY = startY;
        foreach (DocxParagraph paragraph in paragraphs)
        {
            if (paragraph.Runs.Count == 0)
            {
                continue;
            }

            (DocxTextRun Run, string Text, double FontSize, DocxRunFontResource Resource)[] segments = paragraph.Runs
                .Select(run => (
                    Run: run,
                    Text: run.Text.Replace("{PAGE}", pageNumber.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal),
                    FontSize: Math.Min(12d, run.FontSize),
                    Resource: ResolveFontResource(run, fontResources)))
                .Where(segment => segment.Resource is not null && segment.Text.Length != 0)
                .Select(segment => (segment.Run, segment.Text, segment.FontSize, segment.Resource!))
                .ToArray();
            if (segments.Length == 0)
            {
                continue;
            }

            double lineWidth = segments.Sum(segment => segment.Resource.Embedded.MeasureTextPoints(segment.Text, segment.FontSize));
            double lineX = paragraph.Alignment switch
            {
                DocxTextAlignment.Center => x + Math.Max(0, width - lineWidth) / 2d,
                DocxTextAlignment.Right => x + Math.Max(0, width - lineWidth),
                _ => x
            };
            double segmentX = lineX;
            foreach ((DocxTextRun run, string text, double fontSize, DocxRunFontResource resource) in segments)
            {
                RgbColor color = ReadColor(run.ColorHex);
                graphics.DrawGlyphText(resource.Name, fontSize, segmentX, cursorY, color.Red, color.Green, color.Blue, resource.Embedded.EncodeGlyphHex(text), run.Italic);
                segmentX += resource.Embedded.MeasureTextPoints(text, fontSize);
            }

            cursorY -= segments.Max(segment => segment.FontSize) * 1.2d;
        }
    }

    private static RgbColor ReadColor(string? hex)
    {
        return RgbColor.TryParse(hex, out RgbColor color) ? color : new RgbColor(0, 0, 0);
    }

    private static PdfImageXObject? CreateImage(DocxInlineImage image, Action<OoxPdfDiagnostic>? diagnosticSink, int pageIndex)
    {
        try
        {
            if (image.ContentType.Equals("image/jpeg", StringComparison.OrdinalIgnoreCase) ||
                image.ContentType.Equals("image/jpg", StringComparison.OrdinalIgnoreCase))
            {
                JpegInfo info = JpegInfo.Read(image.Bytes);
                return PdfImageXObject.Jpeg(info.Width, info.Height, image.Bytes, info.ComponentCount, info.BitsPerComponent);
            }

            if (image.ContentType.Equals("image/png", StringComparison.OrdinalIgnoreCase))
            {
                PngImage png = PngImage.Read(image.Bytes);
                return PdfImageXObject.RgbPng(png.Width, png.Height, png.Rgb, png.Alpha);
            }

            if (image.ContentType.Equals("image/bmp", StringComparison.OrdinalIgnoreCase) ||
                image.ContentType.Equals("image/x-ms-bmp", StringComparison.OrdinalIgnoreCase))
            {
                BmpImage bmp = BmpImage.Read(image.Bytes);
                return PdfImageXObject.RgbPng(bmp.Width, bmp.Height, bmp.Rgb, bmp.Alpha);
            }
        }
        catch (Exception ex) when (ex is InvalidDataException or NotSupportedException)
        {
            EmitImageDiagnostic(diagnosticSink, image, pageIndex, ex.Message);
            return null;
        }

        EmitImageDiagnostic(diagnosticSink, image, pageIndex, "Unsupported image content type.");
        return null;
    }

    private static void EmitImageDiagnostic(Action<OoxPdfDiagnostic>? diagnosticSink, DocxInlineImage image, int pageIndex, string reason)
    {
        diagnosticSink?.Invoke(new OoxPdfDiagnostic(
            "IMAGE_UNSUPPORTED_FORMAT",
            OoxPdfSeverity.Error,
            $"Image '{image.ContentType}' could not be rendered and was ignored: {reason}",
            image.PartName,
            PageIndex: pageIndex,
            Feature: image.ContentType,
            Fallback: "Ignored"));
    }

}
