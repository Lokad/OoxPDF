using System.Globalization;
using System.Text;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed class DocxRenderer
{
    public IReadOnlyList<PdfPage> RenderBlankPages(DocxDocument document, Action<OoxPdfDiagnostic>? diagnosticSink = null)
    {
        if (document.BodyElements.Count == 0 && document.HeaderParagraphs.Count == 0 && document.FooterParagraphs.Count == 0)
        {
            return [new PdfPage(document.PageWidthPoints, document.PageHeightPoints)];
        }

        return RenderParagraphs(document, diagnosticSink);
    }

    private static IReadOnlyList<PdfPage> RenderParagraphs(DocxDocument document, Action<OoxPdfDiagnostic>? diagnosticSink)
    {
        string familyName = document.Paragraphs
            .Concat(document.HeaderParagraphs)
            .Concat(document.FooterParagraphs)
            .SelectMany(p => p.Runs)
            .Select(r => r.FontFamily)
            .FirstOrDefault(f => !string.IsNullOrWhiteSpace(f)) ?? "Arial";
        FontResolution resolution = new WindowsFontResolver().Resolve(new FontRequest(familyName));
        IReadOnlyList<DocxTextRun> allRuns = document.Paragraphs
            .Concat(document.HeaderParagraphs)
            .Concat(document.FooterParagraphs)
            .SelectMany(p => p.Runs)
            .ToArray();
        IEnumerable<int> tableRunes = document.Tables
            .SelectMany(t => t.Rows)
            .SelectMany(r => r.Cells)
            .SelectMany(c => c.Text.EnumerateRunes().Select(rune => rune.Value));
        PdfEmbeddedFont? embedded = null;
        PdfFontResource? resource = null;
        IReadOnlyList<int> glyphs = allRuns
            .SelectMany(r => r.Text.EnumerateRunes().Select(rune => rune.Value))
            .Concat(tableRunes)
            .Concat("0123456789".EnumerateRunes().Select(rune => rune.Value))
            .ToArray();
        if (glyphs.Count > 0 && resolution.FontFilePath is not null && File.Exists(resolution.FontFilePath))
        {
            OpenTypeFont font = OpenTypeFont.Load(resolution.FontFilePath);
            embedded = PdfEmbeddedFont.Create(font, glyphs);
            resource = new PdfFontResource("F1", embedded);
        }

        var pages = new List<PdfPage>();
        var graphics = new PdfGraphicsBuilder();
        var pageImages = new List<PdfImageResource>();
        int imageIndex = 1;

        double x = document.MarginLeftPoints;
        double width = Math.Max(1d, document.PageWidthPoints - document.MarginLeftPoints - document.MarginRightPoints);
        double cursorY = document.PageHeightPoints - document.MarginTopPoints;
        double pendingSpacingAfter = 0d;
        const double baselineOffsetFactor = 0.94d;
        void FinishPage()
        {
            int pageNumber = pages.Count + 1;
            if (embedded is not null)
            {
                RenderStaticParagraphs(document.HeaderParagraphs, graphics, embedded, x, width, document.PageHeightPoints - Math.Max(18d, document.MarginTopPoints / 2d), pageNumber);
                RenderStaticParagraphs(document.FooterParagraphs, graphics, embedded, x, width, Math.Max(18d, document.MarginBottomPoints / 2d), pageNumber);
            }

            IReadOnlyList<PdfFontResource> fonts = resource is null ? [] : [resource];
            pages.Add(new PdfPage(document.PageWidthPoints, document.PageHeightPoints, graphics.ToString(), fonts, pageImages.ToArray()));
            graphics = new PdfGraphicsBuilder();
            pageImages.Clear();
            cursorY = document.PageHeightPoints - document.MarginTopPoints;
            pendingSpacingAfter = 0d;
        }

        foreach (DocxBodyElement element in document.BodyElements)
        {
            if (element is DocxPageBreakElement)
            {
                if (HasPageContent(graphics, pageImages))
                {
                    FinishPage();
                }
                pendingSpacingAfter = 0d;

                continue;
            }

            if (element is DocxTableElement tableElement)
            {
                cursorY -= pendingSpacingAfter;
                pendingSpacingAfter = 0d;
                RenderTable(tableElement.Table, document, ref graphics, pageImages, resource, embedded, ref cursorY, x, width, FinishPage);
                continue;
            }

            if (element is not DocxParagraphElement paragraphElement)
            {
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            cursorY -= Math.Max(pendingSpacingAfter, paragraph.SpacingBeforePoints);
            pendingSpacingAfter = 0d;
            double paragraphFontSize = paragraph.Runs.Count == 0 ? 11d : paragraph.Runs.Max(r => r.FontSize);
            double lineHeight = paragraph.LineSpacingPoints ?? paragraphFontSize * paragraph.LineSpacingFactor;
            if (embedded is not null && paragraph.Runs.Count > 0)
            {
                string text = paragraph.ListLabel is null
                    ? string.Concat(paragraph.Runs.Select(r => r.Text))
                    : paragraph.ListLabel.Text + " " + string.Concat(paragraph.Runs.Select(r => r.Text));
                DocxTextRun firstRun = paragraph.Runs[0];
                RgbColor color = ReadColor(firstRun.ColorHex);
                foreach (string line in WrapWords(text, width, paragraphFontSize, embedded))
                {
                    if (cursorY - lineHeight < document.MarginBottomPoints && HasPageContent(graphics, pageImages))
                    {
                        FinishPage();
                    }

                    double lineWidth = embedded.MeasureTextPoints(line, paragraphFontSize);
                    double lineX = paragraph.Alignment switch
                    {
                        DocxTextAlignment.Center => x + Math.Max(0, width - lineWidth) / 2d,
                        DocxTextAlignment.Right => x + Math.Max(0, width - lineWidth),
                        _ => x
                    };
                    string glyphHex = embedded.EncodeGlyphHex(line);
                    double baselineOffset = paragraph.LineSpacingPoints is null
                        ? paragraphFontSize * baselineOffsetFactor
                        : Math.Max(0d, lineHeight - paragraphFontSize * 0.299d);
                    double baselineY = cursorY - baselineOffset;
                    graphics.DrawGlyphText("F1", paragraphFontSize, lineX, baselineY, color.Red, color.Green, color.Blue, glyphHex, firstRun.Italic);
                    if (firstRun.Bold)
                    {
                        graphics.DrawGlyphText("F1", paragraphFontSize, lineX + 0.35d, baselineY, color.Red, color.Green, color.Blue, glyphHex, firstRun.Italic);
                    }

                    if (firstRun.Underline)
                    {
                        graphics.SetStrokeRgb(color.Red, color.Green, color.Blue);
                        graphics.SetLineWidth(Math.Max(0.5d, paragraphFontSize / 18d));
                        graphics.StrokeLine(lineX, baselineY - paragraphFontSize * 0.12d, lineX + lineWidth, baselineY - paragraphFontSize * 0.12d);
                    }

                    cursorY -= lineHeight;
                }
            }

            foreach (DocxInlineImage image in paragraph.Images)
            {
                PdfImageXObject? xObject = CreateImage(image, diagnosticSink, pages.Count + 1);
                if (xObject is null)
                {
                    continue;
                }

                double imageWidth = Math.Min(width, image.WidthPoints);
                double imageHeight = image.HeightPoints * imageWidth / Math.Max(1d, image.WidthPoints);
                if (cursorY - imageHeight < document.MarginBottomPoints && HasPageContent(graphics, pageImages))
                {
                    FinishPage();
                }

                string imageName = "Im" + imageIndex++;
                double imageX = paragraph.Alignment switch
                {
                    DocxTextAlignment.Center => x + Math.Max(0, width - imageWidth) / 2d,
                    DocxTextAlignment.Right => x + Math.Max(0, width - imageWidth),
                    _ => x
                };
                graphics.DrawImage(imageName, imageX, cursorY - imageHeight, imageWidth, imageHeight);
                pageImages.Add(new PdfImageResource(imageName, xObject));
                cursorY -= imageHeight + 6d;
            }

            pendingSpacingAfter = paragraph.SpacingAfterPoints;
        }

        if (HasPageContent(graphics, pageImages) || pages.Count == 0)
        {
            FinishPage();
        }

        return pages;
    }

    private static void RenderStaticParagraphs(
        IReadOnlyList<DocxParagraph> paragraphs,
        PdfGraphicsBuilder graphics,
        PdfEmbeddedFont embedded,
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

            double fontSize = Math.Min(12d, paragraph.Runs.Max(r => r.FontSize));
            string text = string.Concat(paragraph.Runs.Select(r => r.Text)).Replace("{PAGE}", pageNumber.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal);
            DocxTextRun firstRun = paragraph.Runs[0];
            RgbColor color = ReadColor(firstRun.ColorHex);
            double lineWidth = embedded.MeasureTextPoints(text, fontSize);
            double lineX = paragraph.Alignment switch
            {
                DocxTextAlignment.Center => x + Math.Max(0, width - lineWidth) / 2d,
                DocxTextAlignment.Right => x + Math.Max(0, width - lineWidth),
                _ => x
            };
            graphics.DrawGlyphText("F1", fontSize, lineX, cursorY, color.Red, color.Green, color.Blue, embedded.EncodeGlyphHex(text), firstRun.Italic);
            cursorY -= fontSize * 1.2d;
        }
    }

    private static void RenderTable(
        DocxTable table,
        DocxDocument document,
        ref PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        PdfFontResource? fontResource,
        PdfEmbeddedFont? embedded,
        ref double cursorY,
        double x,
        double availableWidth,
        Action finishPage)
    {
        const double defaultRowHeight = 16d;
        double rawTableWidth = table.ColumnWidthsPoints.Sum();
        double scale = rawTableWidth <= 0d ? 1d : Math.Min(1d, availableWidth / rawTableWidth);
        double tableHeight = table.Rows.Sum(row => row.HeightPoints ?? defaultRowHeight);
        if (cursorY - tableHeight < document.MarginBottomPoints && HasPageContent(graphics, pageImages))
        {
            finishPage();
        }

        foreach (DocxTableRow row in table.Rows)
        {
            double rowHeight = row.HeightPoints ?? defaultRowHeight;
            if (cursorY - rowHeight < document.MarginBottomPoints && HasPageContent(graphics, pageImages))
            {
                finishPage();
            }

            double cellX = x;
            double cellY = cursorY - rowHeight;
            for (int columnIndex = 0; columnIndex < row.Cells.Count; columnIndex++)
            {
                double cellWidth = table.ColumnWidthsPoints[Math.Min(columnIndex, table.ColumnWidthsPoints.Count - 1)] * scale;
                DocxTableCell cell = row.Cells[columnIndex];
                if (RgbColor.TryParse(cell.FillHex, out RgbColor fill))
                {
                    graphics.SetFillRgb(fill.Red, fill.Green, fill.Blue);
                    graphics.FillRectangle(cellX, cellY, cellWidth, rowHeight);
                }

                graphics.SetStrokeRgb(0, 0, 0);
                graphics.SetLineWidth(0.75d);
                graphics.StrokeRectangle(cellX, cellY, cellWidth, rowHeight);
                if (embedded is not null && fontResource is not null && cell.Text.Length != 0)
                {
                    string glyphHex = embedded.EncodeGlyphHex(cell.Text);
                    graphics.DrawGlyphText(fontResource.ResourceName, 11d, cellX + 4d, cursorY - 17d, 0, 0, 0, glyphHex);
                }

                cellX += cellWidth;
            }

            cursorY -= rowHeight;
        }

        cursorY -= 6d;
    }

    private static RgbColor ReadColor(string? hex)
    {
        return RgbColor.TryParse(hex, out RgbColor color) ? color : new RgbColor(0, 0, 0);
    }

    private static bool HasPageContent(PdfGraphicsBuilder graphics, IReadOnlyCollection<PdfImageResource> images)
    {
        return graphics.ToString().Length > 0 || images.Count > 0;
    }

    private static PdfImageXObject? CreateImage(DocxInlineImage image, Action<OoxPdfDiagnostic>? diagnosticSink, int pageIndex)
    {
        try
        {
            if (image.ContentType.Equals("image/jpeg", StringComparison.OrdinalIgnoreCase) ||
                image.ContentType.Equals("image/jpg", StringComparison.OrdinalIgnoreCase))
            {
                JpegInfo info = JpegInfo.Read(image.Bytes);
                return PdfImageXObject.Jpeg(info.Width, info.Height, image.Bytes);
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
}
