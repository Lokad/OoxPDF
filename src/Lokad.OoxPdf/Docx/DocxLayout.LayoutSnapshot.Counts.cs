using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed partial record DocxLayoutSnapshot
{
    private static int CountBodyTextLength(IReadOnlyList<DocxBodyElement> elements)
    {
        int length = 0;
        foreach (DocxBodyElement element in elements)
        {
            if (element is DocxParagraphElement paragraphElement)
            {
                length += paragraphElement.Paragraph.Runs.Sum(run => run.Text.Length);
                continue;
            }

            if (element is DocxTableElement tableElement)
            {
                foreach (DocxTableRow row in tableElement.Table.Rows)
                {
                    foreach (DocxTableCell cell in row.Cells)
                    {
                        length += CountBodyTextLength(DocxTableCellContent.GetBodyElements(cell));
                    }
                }
            }
        }

        return length;
    }

    private static int CountTableCellTextLines(IReadOnlyList<DocxTableRowLayout> rows)
    {
        int count = 0;
        foreach (DocxTableRowLayout row in rows)
        {
            foreach (DocxTableCellLayout cell in row.Cells)
            {
                count += cell.TextLines.Count;
                count += CountTableCellTextLines(cell.NestedRows);
            }
        }

        return count;
    }

    private static int CountTableCellInlineImages(IReadOnlyList<DocxTableRowLayout> rows)
    {
        int count = 0;
        foreach (DocxTableRowLayout row in rows)
        {
            foreach (DocxTableCellLayout cell in row.Cells)
            {
                count += cell.InlineImages.Count;
                count += CountTableCellInlineImages(cell.NestedRows);
            }
        }

        return count;
    }

    private static int CountRevisions(DocxParagraph? paragraph)
    {
        return paragraph?.Revisions.Count ?? 0;
    }

    private static int CountRevisions(DocxParagraph? paragraph, DocxRevisionKind kind)
    {
        return paragraph?.Revisions.Count(revision => revision.Kind == kind) ?? 0;
    }

    private static int CountOtherRevisions(DocxParagraph? paragraph)
    {
        return paragraph?.Revisions.Count(revision =>
            revision.Kind is not (DocxRevisionKind.Insertion or DocxRevisionKind.Deletion or DocxRevisionKind.MoveFrom or DocxRevisionKind.MoveTo)) ?? 0;
    }

    private static int SumTableRowTextLineCount(DocxTableRowLayout row)
    {
        return row.Cells.Sum(SumTableCellTextLineCount);
    }

    private static int SumTableRowTextLength(DocxTableRowLayout row)
    {
        return row.Cells.Sum(SumTableCellTextLength);
    }

    private static int SumTableCellTextLineCount(DocxTableCellLayout cell)
    {
        return cell.TextLines.Count + cell.NestedRows.Sum(SumTableRowTextLineCount) + cell.InlineTextBoxes.Sum(SumInlineTextBoxTextLineCount);
    }

    private static int SumTableCellTextLength(DocxTableCellLayout cell)
    {
        return cell.TextLines.Sum(line => line.Text.Length) + cell.NestedRows.Sum(SumTableRowTextLength) + cell.InlineTextBoxes.Sum(SumInlineTextBoxTextLength);
    }

    private static int SumInlineTextBoxTextLineCount(DocxInlineTextBoxLayout box)
    {
        return box.TextLines.Count + box.TableRows.Sum(SumTableRowTextLineCount);
    }

    private static int SumInlineTextBoxTextLength(DocxInlineTextBoxLayout box)
    {
        return box.TextLines.Sum(line => line.Text.Length) + box.TableRows.Sum(SumTableRowTextLength);
    }
}
