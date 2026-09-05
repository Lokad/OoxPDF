using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    private static IReadOnlyList<DocxLayoutPage> ReindexPageOwnedLayouts(IReadOnlyList<DocxLayoutPage> pages)
    {
        return pages
            .Select((page, pageIndex) => page with
            {
                StaticInlineImages = ReindexInlineImages(page.StaticInlineImages, pageIndex),
                StaticTableRows = ReindexTableRows(page.StaticTableRows, pageIndex),
                PlacedRelatedStories = page.PlacedRelatedStories
                    .Select(story => ReindexPlacedRelatedStory(story, pageIndex))
                    .ToArray(),
                Items = page.Items
                    .Select(item => ReindexLayoutItem(item, pageIndex))
                    .ToArray()
            })
            .ToArray();
    }

    private static DocxLayoutItem ReindexLayoutItem(DocxLayoutItem item, int pageIndex)
    {
        return item switch
        {
            DocxInlineImageLayout image => image with { PageIndex = pageIndex },
            DocxTableRowLayout row => ReindexTableRow(row, pageIndex),
            _ => item
        };
    }

    private static DocxPlacedRelatedStoryLayout ReindexPlacedRelatedStory(DocxPlacedRelatedStoryLayout story, int pageIndex)
    {
        return story with
        {
            InlineImages = ReindexInlineImages(story.InlineImages, pageIndex),
            FloatingDrawings = story.FloatingDrawings
                .Select(drawing => drawing with
                {
                    PageStartIndex = pageIndex,
                    PageEndIndex = pageIndex,
                    AnchorPageIndex = pageIndex
                })
                .ToArray(),
            TableRows = ReindexTableRows(story.TableRows, pageIndex)
        };
    }

    private static IReadOnlyList<DocxTableRowLayout> ReindexTableRows(IReadOnlyList<DocxTableRowLayout> rows, int pageIndex)
    {
        return rows
            .Select(row => ReindexTableRow(row, pageIndex))
            .ToArray();
    }

    private static DocxTableRowLayout ReindexTableRow(DocxTableRowLayout row, int pageIndex)
    {
        return row with
        {
            Cells = row.Cells
                .Select(cell => ReindexTableCell(cell, pageIndex))
                .ToArray()
        };
    }

    private static DocxTableCellLayout ReindexTableCell(DocxTableCellLayout cell, int pageIndex)
    {
        return cell with
        {
            InlineImages = ReindexInlineImages(cell.InlineImages, pageIndex),
            NestedTableRows = ReindexTableRows(cell.NestedRows, pageIndex)
        };
    }

    private static IReadOnlyList<DocxInlineImageLayout> ReindexInlineImages(IReadOnlyList<DocxInlineImageLayout> images, int pageIndex)
    {
        return images
            .Select(image => image with { PageIndex = pageIndex })
            .ToArray();
    }
}
