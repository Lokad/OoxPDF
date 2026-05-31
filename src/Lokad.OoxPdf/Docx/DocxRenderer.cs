using System.Globalization;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed class DocxRenderer
{
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
        DocxLayout layout = new DocxLayoutEngine().Create(document, fontResources.Embedded);
        return DocxLayoutSnapshot.FromLayout(layout);
    }

    private static IReadOnlyList<PdfPage> RenderParagraphs(DocxDocument document, IFontResolver fontResolver, Action<OoxPdfDiagnostic>? diagnosticSink)
    {
        DocxFontResources fontResources = PrepareFontResources(document, fontResolver);
        PdfEmbeddedFont? embedded = fontResources.Embedded;
        PdfFontResource? resource = fontResources.Resource;

        DocxLayout layout = new DocxLayoutEngine().Create(document, embedded);
        var pages = new List<PdfPage>(layout.Pages.Count);
        int imageIndex = 1;

        double width = Math.Max(1d, document.PageWidthPoints - document.MarginLeftPoints - document.MarginRightPoints);
        for (int pageIndex = 0; pageIndex < layout.Pages.Count; pageIndex++)
        {
            DocxLayoutPage layoutPage = layout.Pages[pageIndex];
            var graphics = new PdfGraphicsBuilder();
            var pageImages = new List<PdfImageResource>();
            foreach (DocxLayoutItem item in layoutPage.Items)
            {
                RenderLayoutItem(item, graphics, pageImages, resource, embedded, diagnosticSink, ref imageIndex);
            }

            if (embedded is not null)
            {
                int pageNumber = pageIndex + 1;
                RenderStaticParagraphs(document.HeaderParagraphs, graphics, embedded, document.MarginLeftPoints, width, document.PageHeightPoints - Math.Max(18d, document.MarginTopPoints / 2d), pageNumber);
                RenderStaticParagraphs(document.FooterParagraphs, graphics, embedded, document.MarginLeftPoints, width, Math.Max(18d, document.MarginBottomPoints / 2d), pageNumber);
            }

            IReadOnlyList<PdfFontResource> fonts = resource is null ? [] : [resource];
            pages.Add(new PdfPage(layoutPage.Width, layoutPage.Height, graphics.ToString(), fonts, pageImages.ToArray()));
        }

        return pages;
    }

    private static DocxFontResources PrepareFontResources(DocxDocument document, IFontResolver fontResolver)
    {
        string familyName = document.Paragraphs
            .Concat(document.HeaderParagraphs)
            .Concat(document.FooterParagraphs)
            .SelectMany(p => p.Runs)
            .Select(r => r.FontFamily)
            .FirstOrDefault(f => !string.IsNullOrWhiteSpace(f)) ?? "Arial";
        FontResolution resolution = fontResolver.Resolve(new FontRequest(familyName));
        IReadOnlyList<DocxTextRun> allRuns = document.Paragraphs
            .Concat(document.HeaderParagraphs)
            .Concat(document.FooterParagraphs)
            .SelectMany(p => p.Runs)
            .ToArray();
        IEnumerable<int> tableRunes = document.Tables
            .SelectMany(t => t.Rows)
            .SelectMany(r => r.Cells)
            .SelectMany(c => c.Paragraphs.Count == 0
                ? c.Text.EnumerateRunes().Select(rune => rune.Value)
                : c.Paragraphs.SelectMany(p => p.Runs).SelectMany(run => run.Text.EnumerateRunes().Select(rune => rune.Value)));
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

        return new DocxFontResources(embedded, resource);
    }

    private static void RenderLayoutItem(
        DocxLayoutItem item,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        PdfFontResource? fontResource,
        PdfEmbeddedFont? embedded,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        ref int imageIndex)
    {
        switch (item)
        {
            case DocxTextLineLayout textLine when embedded is not null:
                RenderTextLine(textLine, graphics, embedded);
                break;
            case DocxInlineImageLayout image:
                RenderInlineImage(image, graphics, pageImages, diagnosticSink, ref imageIndex);
                break;
            case DocxTableRowLayout row when embedded is not null && fontResource is not null:
                RenderTableRow(row, graphics, fontResource, embedded);
                break;
            case DocxTableRowLayout row:
                RenderTableRow(row, graphics, fontResource, embedded);
                break;
        }
    }

    private static void RenderTextLine(DocxTextLineLayout line, PdfGraphicsBuilder graphics, PdfEmbeddedFont embedded)
    {
        DocxTextRun style = line.StyleRun;
        RgbColor color = ReadColor(style.ColorHex);
        string glyphHex = embedded.EncodeGlyphHex(line.Text);
        graphics.DrawGlyphText("F1", line.FontSize, line.X, line.BaselineY, color.Red, color.Green, color.Blue, glyphHex, style.Italic);
        if (style.Bold)
        {
            graphics.DrawGlyphText("F1", line.FontSize, line.X + 0.35d, line.BaselineY, color.Red, color.Green, color.Blue, glyphHex, style.Italic);
        }

        if (style.Underline)
        {
            graphics.SetStrokeRgb(color.Red, color.Green, color.Blue);
            graphics.SetLineWidth(Math.Max(0.5d, line.FontSize / 18d));
            graphics.StrokeLine(line.X, line.BaselineY - line.FontSize * 0.12d, line.X + line.Width, line.BaselineY - line.FontSize * 0.12d);
        }
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

    private static void RenderTableRow(DocxTableRowLayout row, PdfGraphicsBuilder graphics, PdfFontResource? fontResource, PdfEmbeddedFont? embedded)
    {
        foreach (DocxTableCellLayout cellLayout in row.Cells)
        {
            DocxTableCell cell = cellLayout.Cell;
            if (RgbColor.TryParse(cell.FillHex, out RgbColor fill))
            {
                graphics.SetFillRgb(fill.Red, fill.Green, fill.Blue);
                graphics.FillRectangle(cellLayout.X, cellLayout.Y, cellLayout.Width, cellLayout.Height);
            }

            graphics.SetStrokeRgb(0, 0, 0);
            graphics.SetLineWidth(0.75d);
            graphics.StrokeRectangle(cellLayout.X, cellLayout.Y, cellLayout.Width, cellLayout.Height);
            if (embedded is not null && fontResource is not null)
            {
                foreach (DocxTextLineLayout line in cellLayout.TextLines)
                {
                    RenderTextLine(line, graphics, embedded);
                }
            }
        }
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
