using System.Globalization;

namespace Lokad.OoxPdf.Docx;

internal sealed partial record DocxStructureSnapshot
{
    // Single-caller pipeline stage; kept static.
    private static IEnumerable<DocxRevisionInfo> EnumerateDocumentRevisions(DocxDocument document)
    {
        foreach (DocxRevisionInfo revision in EnumerateBodyElementRevisions(document.BodyElements))
        {
            yield return revision;
        }

        foreach (DocxFloatingDrawing drawing in document.FloatingDrawings)
        {
            foreach (DocxRevisionInfo revision in EnumerateFloatingDrawingRevisions(drawing))
            {
                yield return revision;
            }
        }

        foreach (DocxRevisionInfo revision in EnumerateStaticStoryRevisions(
            document.HeaderBodyElementsByType,
            document.HeaderParagraphsByType,
            document.HeaderFloatingDrawingsByType))
        {
            yield return revision;
        }

        foreach (DocxRevisionInfo revision in EnumerateStaticStoryRevisions(
            document.FooterBodyElementsByType,
            document.FooterParagraphsByType,
            document.FooterFloatingDrawingsByType))
        {
            yield return revision;
        }

        foreach (DocxRevisionInfo revision in EnumeratePageSettingsRevisions(document.PageSettings))
        {
            yield return revision;
        }

        foreach (DocxSectionBreakElement sectionBreak in document.BodyElements.OfType<DocxSectionBreakElement>())
        {
            foreach (DocxRevisionInfo revision in EnumeratePageSettingsRevisions(sectionBreak.PageSettings))
            {
                yield return revision;
            }
        }

        foreach (DocxRelatedStory story in document.RelatedStories)
        {
            foreach (DocxRevisionInfo revision in EnumerateBodyElementRevisions(story.BodyElements))
            {
                yield return revision;
            }

            foreach (DocxFloatingDrawing drawing in story.FloatingDrawings)
            {
                foreach (DocxRevisionInfo revision in EnumerateFloatingDrawingRevisions(drawing))
                {
                    yield return revision;
                }
            }
        }

        if (document.FinalSectionBreak is { } finalSectionBreak)
        {
            foreach (DocxRevisionInfo revision in finalSectionBreak.Revisions)
            {
                yield return revision;
            }
        }
    }

    // Single-caller pipeline stage; kept static.
    private static IEnumerable<DocxRevisionInfo> EnumeratePageSettingsRevisions(DocxPageSettings settings)
    {
        foreach (DocxRevisionInfo revision in EnumerateStaticStoryRevisions(
            settings.HeaderBodyElementsByType,
            settings.HeaderParagraphsByType,
            settings.HeaderFloatingDrawingsByType))
        {
            yield return revision;
        }

        foreach (DocxRevisionInfo revision in EnumerateStaticStoryRevisions(
            settings.FooterBodyElementsByType,
            settings.FooterParagraphsByType,
            settings.FooterFloatingDrawingsByType))
        {
            yield return revision;
        }
    }

    private static IEnumerable<DocxRevisionInfo> EnumerateStaticStoryRevisions(
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> bodyElementsByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> fallbackParagraphsByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxFloatingDrawing>> floatingDrawingsByType)
    {
        if (bodyElementsByType.Count == 0)
        {
            foreach (DocxParagraph paragraph in fallbackParagraphsByType.Values.SelectMany(paragraphs => paragraphs))
            {
                foreach (DocxRevisionInfo revision in paragraph.Revisions)
                {
                    yield return revision;
                }
            }
        }
        else
        {
            foreach (IReadOnlyList<DocxBodyElement> bodyElements in bodyElementsByType.Values)
            {
                foreach (DocxRevisionInfo revision in EnumerateBodyElementRevisions(bodyElements))
                {
                    yield return revision;
                }
            }
        }

        foreach (DocxFloatingDrawing drawing in floatingDrawingsByType.Values.SelectMany(drawings => drawings))
        {
            foreach (DocxRevisionInfo revision in EnumerateFloatingDrawingRevisions(drawing))
            {
                yield return revision;
            }
        }
    }

    private static IEnumerable<DocxRevisionInfo> EnumerateFloatingDrawingRevisions(DocxFloatingDrawing drawing)
    {
        foreach (DocxRevisionInfo revision in drawing.Revisions)
        {
            yield return revision;
        }

        foreach (DocxRevisionInfo revision in EnumerateBodyElementRevisions(drawing.TextBoxBodyElements))
        {
            yield return revision;
        }
    }

    private static IEnumerable<DocxRevisionInfo> EnumerateBodyElementRevisions(IEnumerable<DocxBodyElement> bodyElements)
    {
        foreach (DocxBodyElement element in bodyElements)
        {
            switch (element)
            {
                case DocxParagraphElement paragraph:
                    foreach (DocxRevisionInfo revision in paragraph.Paragraph.Revisions)
                    {
                        yield return revision;
                    }

                    break;
                case DocxTableElement table:
                    foreach (DocxRevisionInfo revision in EnumerateTableRevisions(table.Table))
                    {
                        yield return revision;
                    }

                    break;
                case DocxSectionBreakElement sectionBreak:
                    foreach (DocxRevisionInfo revision in sectionBreak.Revisions)
                    {
                        yield return revision;
                    }

                    break;
            }
        }
    }

    private static IEnumerable<DocxRevisionInfo> EnumerateTableRevisions(DocxTable table)
    {
        foreach (DocxRevisionInfo revision in table.Revisions)
        {
            yield return revision;
        }

        foreach (DocxTableRow row in table.Rows)
        {
            foreach (DocxRevisionInfo revision in row.Revisions)
            {
                yield return revision;
            }

            foreach (DocxTableCell cell in row.Cells)
            {
                foreach (DocxRevisionInfo revision in EnumerateTableCellRevisions(cell))
                {
                    yield return revision;
                }
            }
        }

    }

    private static IEnumerable<DocxRevisionInfo> EnumerateTableCellRevisions(DocxTableCell cell)
    {
        foreach (DocxRevisionInfo revision in cell.Revisions)
        {
            yield return revision;
        }

        foreach (DocxParagraph paragraph in DocxBlockTraversal.EnumerateBodyParagraphs(DocxTableCellContent.GetBodyElements(cell)))
        {
            foreach (DocxRevisionInfo revision in paragraph.Revisions)
            {
                yield return revision;
            }
        }
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static DocxStructureTableAdjacencySnapshot ToTableAdjacencySnapshot(
        IReadOnlyList<DocxBodyElement> elements,
        DocxTable table,
        int tableIndex,
        int blockIndex,
        string? previousKind,
        string? nextKind)
    {
            DocxParagraph? TryGetAdjacentParagraph(IReadOnlyList<DocxBodyElement> elements, int index)
            {
                return index >= 0 &&
                    index < elements.Count &&
                    elements[index] is DocxParagraphElement paragraph
                        ? paragraph.Paragraph
                        : null;
            }

        DocxParagraph? previousParagraph = TryGetAdjacentParagraph(elements, blockIndex - 1);
        DocxParagraph? nextParagraph = TryGetAdjacentParagraph(elements, blockIndex + 1);
        return new DocxStructureTableAdjacencySnapshot(
            tableIndex,
            blockIndex,
            previousKind,
            nextKind,
            table.Rows.Count,
            MaxColumnCount(table),
            previousParagraph?.EffectiveProperties.StyleId,
            previousParagraph?.EffectiveProperties.SpacingAfterPoints,
            previousParagraph is null ? null : HasAfterSpacingToken(previousParagraph.EffectiveProperties.Spacing),
            nextParagraph?.EffectiveProperties.StyleId,
            nextParagraph?.EffectiveProperties.SpacingBeforePoints,
            nextParagraph is null ? null : HasBeforeSpacingToken(nextParagraph.EffectiveProperties.Spacing),
            nextParagraph is null ? null : TextLength(nextParagraph),
            nextParagraph?.ListLabel is not null,
            nextParagraph?.EffectiveProperties.KeepRules.KeepNext,
            nextParagraph?.EffectiveProperties.KeepRules.KeepLines);
    }
}
