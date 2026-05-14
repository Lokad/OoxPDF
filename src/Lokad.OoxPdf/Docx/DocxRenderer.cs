using System.Globalization;
using System.Text;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed class DocxRenderer
{
    public IReadOnlyList<PdfPage> RenderBlankPages(DocxDocument document)
    {
        if (document.Paragraphs.Count == 0)
        {
            return [new PdfPage(document.PageWidthPoints, document.PageHeightPoints)];
        }

        return RenderParagraphs(document);
    }

    private static IReadOnlyList<PdfPage> RenderParagraphs(DocxDocument document)
    {
        string familyName = document.Paragraphs
            .SelectMany(p => p.Runs)
            .Select(r => r.FontFamily)
            .FirstOrDefault(f => !string.IsNullOrWhiteSpace(f)) ?? "Arial";
        FontResolution resolution = new WindowsFontResolver().Resolve(new FontRequest(familyName));
        IReadOnlyList<DocxTextRun> allRuns = document.Paragraphs.SelectMany(p => p.Runs).ToArray();
        PdfEmbeddedFont? embedded = null;
        PdfFontResource? resource = null;
        if (allRuns.Count > 0 && resolution.FontFilePath is not null && File.Exists(resolution.FontFilePath))
        {
            OpenTypeFont font = OpenTypeFont.Load(resolution.FontFilePath);
            embedded = PdfEmbeddedFont.Create(font, allRuns.SelectMany(r => r.Text.EnumerateRunes().Select(rune => rune.Value)));
            resource = new PdfFontResource("F1", embedded);
        }

        var pages = new List<PdfPage>();
        var graphics = new PdfGraphicsBuilder();
        var pageImages = new List<PdfImageResource>();
        int imageIndex = 1;

        double x = document.MarginLeftPoints;
        double width = Math.Max(1d, document.PageWidthPoints - document.MarginLeftPoints - document.MarginRightPoints);
        double cursorY = document.PageHeightPoints - document.MarginTopPoints;
        void FinishPage()
        {
            IReadOnlyList<PdfFontResource> fonts = resource is null ? [] : [resource];
            pages.Add(new PdfPage(document.PageWidthPoints, document.PageHeightPoints, graphics.ToString(), fonts, pageImages.ToArray()));
            graphics = new PdfGraphicsBuilder();
            pageImages.Clear();
            cursorY = document.PageHeightPoints - document.MarginTopPoints;
        }

        foreach (DocxParagraph paragraph in document.Paragraphs)
        {
            cursorY -= paragraph.SpacingBeforePoints;
            double paragraphFontSize = paragraph.Runs.Count == 0 ? 11d : paragraph.Runs.Max(r => r.FontSize);
            double lineHeight = paragraphFontSize * paragraph.LineSpacingFactor;
            if (embedded is not null && paragraph.Runs.Count > 0)
            {
                string text = paragraph.ListLabel is null
                    ? string.Concat(paragraph.Runs.Select(r => r.Text))
                    : paragraph.ListLabel + " " + string.Concat(paragraph.Runs.Select(r => r.Text));
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
                    graphics.DrawGlyphText("F1", paragraphFontSize, lineX, cursorY, color.Red, color.Green, color.Blue, glyphHex, firstRun.Italic);
                    if (firstRun.Bold)
                    {
                        graphics.DrawGlyphText("F1", paragraphFontSize, lineX + 0.35d, cursorY, color.Red, color.Green, color.Blue, glyphHex, firstRun.Italic);
                    }

                    if (firstRun.Underline)
                    {
                        graphics.SetStrokeRgb(color.Red, color.Green, color.Blue);
                        graphics.SetLineWidth(Math.Max(0.5d, paragraphFontSize / 18d));
                        graphics.StrokeLine(lineX, cursorY - paragraphFontSize * 0.12d, lineX + lineWidth, cursorY - paragraphFontSize * 0.12d);
                    }

                    cursorY -= lineHeight;
                }
            }

            foreach (DocxInlineImage image in paragraph.Images)
            {
                PdfImageXObject? xObject = CreateImage(image);
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

            cursorY -= paragraph.SpacingAfterPoints;
        }

        if (HasPageContent(graphics, pageImages) || pages.Count == 0)
        {
            FinishPage();
        }

        return pages;
    }

    private static RgbColor ReadColor(string? hex)
    {
        return RgbColor.TryParse(hex, out RgbColor color) ? color : new RgbColor(0, 0, 0);
    }

    private static bool HasPageContent(PdfGraphicsBuilder graphics, IReadOnlyCollection<PdfImageResource> images)
    {
        return graphics.ToString().Length > 0 || images.Count > 0;
    }

    private static PdfImageXObject? CreateImage(DocxInlineImage image)
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

        return null;
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
