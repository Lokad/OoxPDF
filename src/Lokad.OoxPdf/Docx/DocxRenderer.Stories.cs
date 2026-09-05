using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private static void RenderPlacedRelatedStory(
        DocxPlacedRelatedStoryLayout story,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int pageNumber,
        int pageCount,
        ref int imageIndex)
    {
        if (story.SeparatorY is { } separatorY)
        {
            graphics.SetFillRgb(0, 0, 0);
            graphics.FillRectangle(story.X, separatorY, Math.Min(story.SeparatorWidth, story.Width), story.SeparatorThickness);
        }

        graphics.SaveState();
        graphics.ClipRectangle(story.X, story.TopY - story.Height, story.Width, story.Height);
        IReadOnlyList<DocxLayoutItem> items = story.TextLines
            .Cast<DocxLayoutItem>()
            .Concat(story.InlineImages)
            .Concat(story.TableRows)
            .OrderByDescending(ResolveLayoutItemTop)
            .ToArray();
        for (int itemIndex = 0; itemIndex < items.Count; itemIndex++)
        {
            DocxLayoutItem item = items[itemIndex];
            DocxTableRowLayout? previousRow = itemIndex > 0 ? items[itemIndex - 1] as DocxTableRowLayout : null;
            DocxTableRowLayout? nextRow = itemIndex + 1 < items.Count ? items[itemIndex + 1] as DocxTableRowLayout : null;
            RenderLayoutItem(item, previousRow, nextRow, graphics, pageImages, fontResources, markupContext, diagnosticSink, pageNumber, pageCount, ref imageIndex);
        }

        graphics.RestoreState();
    }

    private static void RenderFloatingDrawings(
        IReadOnlyList<DocxFloatingDrawingLayout> floatingDrawings,
        int pageIndex,
        bool behindDocument,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        int pageNumber,
        int pageCount,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        ref int imageIndex)
    {
        foreach (DocxFloatingDrawingLayout drawing in floatingDrawings
            .Where(drawing => drawing.AnchorPageIndex == pageIndex && IsBehindDocument(drawing.Drawing) == behindDocument)
            .OrderBy(drawing => ReadZOrder(drawing.Drawing.RelativeHeightValue)))
        {
            RenderFloatingDrawing(drawing, graphics, pageImages, fontResources, markupContext, pageNumber, pageCount, diagnosticSink, ref imageIndex);
        }
    }

    private static void RenderPlacedRelatedStoryDrawings(
        DocxPlacedRelatedStoryLayout story,
        bool behindDocument,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        int pageNumber,
        int pageCount,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        ref int imageIndex)
    {
        DocxFloatingDrawingLayout[] drawings = story.FloatingDrawings
            .Where(drawing => IsBehindDocument(drawing.Drawing) == behindDocument)
            .OrderBy(drawing => ReadZOrder(drawing.Drawing.RelativeHeightValue))
            .ToArray();
        if (drawings.Length == 0)
        {
            return;
        }

        graphics.SaveState();
        graphics.ClipRectangle(story.X, story.TopY - story.Height, story.Width, story.Height);
        foreach (DocxFloatingDrawingLayout drawing in drawings)
        {
            RenderFloatingDrawing(drawing, graphics, pageImages, fontResources, markupContext, pageNumber, pageCount, diagnosticSink, ref imageIndex);
        }

        graphics.RestoreState();
    }

    private static void RenderFloatingDrawing(
        DocxFloatingDrawingLayout drawing,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        int pageNumber,
        int pageCount,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        ref int imageIndex)
    {
        if (drawing.PlacedX is not { } placedX ||
            drawing.PlacedTop is not { } placedTop ||
            drawing.ExtentWidthPoints is not { } width ||
            drawing.ExtentHeightPoints is not { } height)
        {
            return;
        }

        if (drawing.Drawing.Image is { } image)
        {
            PdfImageXObject? xObject = CreateImage(image, diagnosticSink, drawing.AnchorPageIndex ?? 0);
            if (xObject is not null)
            {
                string imageName = "Im" + imageIndex++;
                graphics.DrawImage(imageName, placedX, placedTop - height, width, height);
                pageImages.Add(new PdfImageResource(imageName, xObject));
            }
        }

        if (drawing.TextBoxLayout is { } textBoxLayout)
        {
            RenderFloatingTextBox(drawing, textBoxLayout, placedX, placedTop, width, height, graphics, pageImages, fontResources, markupContext, pageNumber, pageCount, diagnosticSink, ref imageIndex);
        }
    }

    private static void RenderFloatingTextBox(
        DocxFloatingDrawingLayout drawing,
        DocxRelatedStoryLayout textBoxLayout,
        double placedX,
        double placedTop,
        double width,
        double height,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        int pageNumber,
        int pageCount,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        ref int imageIndex)
    {
        if (textBoxLayout.TextLines.Count == 0 &&
            textBoxLayout.InlineImages.Count == 0 &&
            textBoxLayout.TableRows.Count == 0)
        {
            return;
        }

        graphics.SaveState();
        graphics.ClipRectangle(placedX, placedTop - height, width, height);
        IReadOnlyList<DocxLayoutItem> items = textBoxLayout.TextLines
            .Select(line => TranslateTextLine(line, placedX, placedTop))
            .Cast<DocxLayoutItem>()
            .Concat(textBoxLayout.InlineImages.Select(image => image with
            {
                X = placedX + image.X,
                Y = placedTop + image.Y,
                PageIndex = drawing.AnchorPageIndex ?? image.PageIndex
            }))
            .Concat(textBoxLayout.TableRows.Select(row => TranslateTableRow(row, placedX, placedTop)))
            .OrderByDescending(ResolveLayoutItemTop)
            .ToArray();
        for (int itemIndex = 0; itemIndex < items.Count; itemIndex++)
        {
            DocxLayoutItem item = items[itemIndex];
            DocxTableRowLayout? previousRow = itemIndex > 0 ? items[itemIndex - 1] as DocxTableRowLayout : null;
            DocxTableRowLayout? nextRow = itemIndex + 1 < items.Count ? items[itemIndex + 1] as DocxTableRowLayout : null;
            RenderLayoutItem(item, previousRow, nextRow, graphics, pageImages, fontResources, markupContext, diagnosticSink, pageNumber, pageCount, ref imageIndex);
        }

        graphics.RestoreState();
    }

    private static double ResolveLayoutItemTop(DocxLayoutItem item)
    {
        return item switch
        {
            DocxTextLineLayout textLine => textLine.BaselineY,
            DocxInlineImageLayout image => image.Y + image.Height,
            DocxTableRowLayout row => row.Y + row.Height,
            _ => 0d
        };
    }

    private static DocxTextLineLayout TranslateTextLine(DocxTextLineLayout line, double deltaX, double deltaY)
    {
        return line with
        {
            X = line.X + deltaX,
            BaselineY = line.BaselineY + deltaY,
            Segments = line.Segments
                .Select(segment => segment with { X = segment.X + deltaX })
                .ToArray()
        };
    }

    private static bool IsBehindDocument(DocxFloatingDrawing drawing)
    {
        return IsOnOffTrue(drawing.BehindDocumentValue);
    }

    private static long ReadZOrder(string? value)
    {
        return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long zOrder)
            ? zOrder
            : 0L;
    }

    private static bool IsOnOffTrue(string? value)
    {
        return value is not null &&
            (value.Length == 0 ||
            value.Equals("1", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("true", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("on", StringComparison.OrdinalIgnoreCase));
    }
}
