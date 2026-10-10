namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private static void AddTableNoteBeforeSpacing(DocxDocument document, DocxLayout layout,
        Dictionary<DocxParagraph, NoteReferenceSpacing> candidates, CancellationToken cancellationToken)
    {
        for (int blockIndex = 0; blockIndex < document.BodyElements.Count; blockIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (document.BodyElements[blockIndex] is not DocxTableElement { Table: var table } ||
                table.Rows.Count != 1 || table.Rows[0].Cells.Count != 1 || table.Revisions.Count != 0 ||
                table.CellSpacingPoints is > 0d || table.UseLegacyTableGrid)
                continue;
            DocxTableRow sourceRow = table.Rows[0];
            DocxTableCell sourceCell = sourceRow.Cells[0];
            if (sourceRow.IsHeader || sourceRow.HeightPoints is not null || sourceRow.Revisions.Count != 0 ||
                sourceCell.Paragraphs.Count != 2 || sourceCell.BodyElements.Count != 2 ||
                sourceCell.BodyElements.Any(e => e is not DocxParagraphElement || e.Revisions.Count != 0) ||
                sourceCell.GridSpan != 1 || sourceCell.HasVerticalMerge || sourceCell.NoWrap || sourceCell.FitText ||
                sourceCell.VerticalAlignment != DocxTableCellVerticalAlignment.Top ||
                sourceCell.TextDirectionValue is not (null or "lrTb") || sourceCell.Revisions.Count != 0)
                continue;
            DocxParagraph previous = sourceCell.Paragraphs[0];
            DocxParagraph current = sourceCell.Paragraphs[1];
            if (!DocxLayoutEngine.IsPlainContextualSpacingParagraph(previous) ||
                !DocxLayoutEngine.IsPlainContextualSpacingParagraph(current) ||
                previous.Runs.All(r => string.IsNullOrWhiteSpace(r.Text)) || current.SpacingBeforePoints <= 0d ||
                current.InlineReferences.Count == 0)
                continue;
            DocxTableRowLayout[] rows = layout.Pages.SelectMany(p => p.Items.OfType<DocxTableRowLayout>())
                .Where(r => r.Table.SourceBlockIndex == blockIndex).ToArray();
            if (rows.Length != 1 || rows[0].FragmentCount != 1 || rows[0].Cells.Count != 1 ||
                rows[0].ReviewBorderPrintScale != 1d)
                continue;
            DocxTableCellLayout cell = rows[0].Cells[0];
            if (!ReferenceEquals(cell.Cell, sourceCell) || cell.NestedRows.Count != 0 ||
                cell.InlineImages.Count != 0 || cell.InlineTextBoxes.Count != 0)
                continue;
            DocxTextLineLayout[] previousLines = cell.TextLines.Where(l => ReferenceEquals(l.SourceParagraph, previous)).ToArray();
            DocxTextLineLayout[] currentLines = cell.TextLines.Where(l => ReferenceEquals(l.SourceParagraph, current)).ToArray();
            if (previousLines.Length != 1 || currentLines.Length <= 1) continue;
            DocxTextLineLayout first = currentLines[0];
            if (first.IsFirstParagraphLine != true || first.LineHeight is not > 0d ||
                first.LineHeightSource is not (DocxLineHeightSource.BodySingleLineAuto or DocxLineHeightSource.ExactLineSpacing) ||
                first.AppliedBeforeSpacing is not { } applied || first.PendingAfterSpacing is not { } pending)
                continue;
            bool suppressPreviousAfter = previous.Spacing.ContextualSpacing == true &&
                DocxLayoutEngine.HasSameContextualSpacingStyle(previous, current);
            double before = Math.Max(0d, applied - (suppressPreviousAfter ? 0d : pending));
            if (before > 0d && double.IsFinite(before))
                // Multiline marks own only the first-line before contribution.
                // Word retains final-cell after-spacing in the row, outside the link.
                candidates[current] = new NoteReferenceSpacing(before, 0d);
        }
    }
}
