namespace Lokad.OoxPdf.Docx;

internal static class DocxBlockTraversal
{
    public static IEnumerable<DocxParagraph> EnumerateBodyParagraphs(DocxDocument document)
    {
        return EnumerateBodyParagraphs(document.BodyElements);
    }

    public static IEnumerable<DocxParagraph> EnumerateBodyParagraphs(DocxRelatedStory story)
    {
        return EnumerateBodyParagraphs(story.BodyElements);
    }

    public static IEnumerable<DocxParagraph> EnumerateStaticStoryParagraphs(
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> bodyElementsByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> fallbackParagraphsByType)
    {
        return bodyElementsByType.Count == 0
            ? fallbackParagraphsByType.Values.SelectMany(paragraphs => paragraphs)
            : bodyElementsByType.Values.SelectMany(EnumerateBodyParagraphs);
    }

    public static IEnumerable<DocxParagraph> EnumerateStaticStoryParagraphs(DocxPageSettings settings)
    {
        return EnumerateStaticStoryParagraphs(settings.HeaderBodyElementsByType, settings.HeaderParagraphsByType)
            .Concat(EnumerateStaticStoryParagraphs(settings.FooterBodyElementsByType, settings.FooterParagraphsByType));
    }

    public static IEnumerable<DocxParagraph> EnumerateReferencedStaticStoryParagraphs(
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> bodyElementsByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> paragraphsByType,
        IReadOnlyList<DocxParagraph> fallbackParagraphs)
    {
        if (bodyElementsByType.Count != 0)
        {
            return bodyElementsByType.Values.SelectMany(EnumerateBodyParagraphs);
        }

        return paragraphsByType.Count == 0
            ? fallbackParagraphs
            : paragraphsByType.Values.SelectMany(paragraphs => paragraphs);
    }

    public static IReadOnlyList<DocxBodyElement> GetStaticStoryBodyElements(
        string variantType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> bodyElementsByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> fallbackParagraphsByType)
    {
        if (bodyElementsByType.TryGetValue(variantType, out IReadOnlyList<DocxBodyElement>? bodyElements))
        {
            return bodyElements;
        }

        return fallbackParagraphsByType.TryGetValue(variantType, out IReadOnlyList<DocxParagraph>? paragraphs)
            ? paragraphs.Select(DocxBodyElementFactory.CreateParagraph).Cast<DocxBodyElement>().ToArray()
            : [];
    }

    public static bool TryGetStaticStoryBodyElements(
        string variantType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> bodyElementsByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> fallbackParagraphsByType,
        out IReadOnlyList<DocxBodyElement> bodyElements)
    {
        if (bodyElementsByType.TryGetValue(variantType, out IReadOnlyList<DocxBodyElement>? foundElements))
        {
            bodyElements = foundElements;
            return true;
        }

        if (fallbackParagraphsByType.TryGetValue(variantType, out IReadOnlyList<DocxParagraph>? paragraphs))
        {
            bodyElements = paragraphs.Select(DocxBodyElementFactory.CreateParagraph).Cast<DocxBodyElement>().ToArray();
            return true;
        }

        bodyElements = [];
        return false;
    }

    public static IEnumerable<DocxParagraph> EnumerateBodyParagraphs(IEnumerable<DocxBodyElement> bodyElements)
    {
        foreach (DocxBodyElement element in bodyElements)
        {
            switch (element)
            {
                case DocxParagraphElement paragraph:
                    yield return paragraph.Paragraph;
                    break;
                case DocxTableElement table:
                    foreach (DocxParagraph cellParagraph in EnumerateTableParagraphs(table.Table))
                    {
                        yield return cellParagraph;
                    }

                    break;
            }
        }
    }

    public static IEnumerable<DocxParagraph> EnumerateDirectParagraphs(IEnumerable<DocxBodyElement> bodyElements)
    {
        return bodyElements
            .OfType<DocxParagraphElement>()
            .Select(element => element.Paragraph);
    }

    public static IEnumerable<DocxTable> EnumerateBodyTables(DocxDocument document)
    {
        return EnumerateBodyTables(document.BodyElements);
    }

    public static IEnumerable<DocxTable> EnumerateBodyTables(DocxRelatedStory story)
    {
        return EnumerateBodyTables(story.BodyElements);
    }

    public static IEnumerable<DocxTable> EnumerateBodyTables(IEnumerable<DocxBodyElement> bodyElements)
    {
        foreach (DocxTableElement table in bodyElements.OfType<DocxTableElement>())
        {
            yield return table.Table;
            foreach (DocxTable nestedTable in EnumerateTableTables(table.Table))
            {
                yield return nestedTable;
            }
        }

        IEnumerable<DocxTable> EnumerateTableTables(DocxTable table)
        {
            return table.Rows
                .SelectMany(row => row.Cells)
                .SelectMany(cell => EnumerateBodyTables(DocxTableCellContent.GetBodyElements(cell)));
        }
    }

    public static IEnumerable<DocxParagraph> EnumerateTableParagraphs(DocxTable table)
    {
        return table.Rows
            .SelectMany(row => row.Cells)
            .SelectMany(cell => EnumerateBodyParagraphs(DocxTableCellContent.GetBodyElements(cell)));
    }
}
