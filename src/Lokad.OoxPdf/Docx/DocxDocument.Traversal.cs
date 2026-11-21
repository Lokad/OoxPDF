namespace Lokad.OoxPdf.Docx;

internal static class DocxBlockTraversal
{
    // All-source inspection and review-author ordering include unprinted stories.
    public static IEnumerable<DocxParagraph> EnumerateSourceParagraphs(DocxDocument document)
        => EnumerateStoryParagraphs(document)
            .Concat(EnumerateStoryParagraphs(document).SelectMany(EnumerateInlineTextBoxParagraphs));

    // Section settings can own distinct static stories even when the document's
    // final header/footer has the same variant key. Child boxes follow the same
    // story owners as their enclosing paragraphs.
    private static IEnumerable<DocxParagraph> EnumerateStoryParagraphs(DocxDocument document)
    {
        return DocxBlockTraversal.EnumerateBodyParagraphs(document)
            .Concat(DocxBlockTraversal.EnumerateReferencedStaticStoryParagraphs(document.HeaderBodyElementsByType, document.HeaderParagraphsByType, document.HeaderParagraphs))
            .Concat(DocxBlockTraversal.EnumerateReferencedStaticStoryParagraphs(document.FooterBodyElementsByType, document.FooterParagraphsByType, document.FooterParagraphs))
            .Concat(DocxBlockTraversal.EnumerateStaticStoryParagraphs(document.PageSettings))
            .Concat(document.BodyElements
                .OfType<DocxSectionBreakElement>()
                .SelectMany(sectionBreak => DocxBlockTraversal.EnumerateStaticStoryParagraphs(sectionBreak.PageSettings)))
            .Concat(document.RelatedStories.SelectMany(DocxBlockTraversal.EnumerateBodyParagraphs))
            .Concat(EnumerateFloatingDrawingTextBoxParagraphs(document));
    }

    private static IEnumerable<DocxParagraph> EnumerateFloatingDrawingTextBoxParagraphs(DocxDocument document)
    {
        return EnumerateSourceFloatingDrawings(document).SelectMany(drawing => EnumerateBodyParagraphs(drawing.TextBoxBodyElements));
    }

    // Resource discovery uses the same effective inheritance and variant rules
    // as placement. Enabled variants remain possible before pagination.
    public static IEnumerable<(DocxPageSettings Settings, string Variant,
        IReadOnlyList<DocxBodyElement> BodyElements, IReadOnlyList<DocxFloatingDrawing> Drawings)>
        EnumerateStaticStorySources(DocxDocument document, CancellationToken token)
    {
        var sections = DocxLayoutEngine.BuildEffectiveSectionSettings(document, out var final, token);
        foreach (var settings in sections.Values.Select(section => section.PageSettings).Prepend(final.PageSettings))
        {
            token.ThrowIfCancellationRequested();
            for (int side = 0; side < 2; side++)
            {
                var bodies = side == 0 ? settings.HeaderBodyElementsByType : settings.FooterBodyElementsByType;
                var paragraphs = side == 0 ? settings.HeaderParagraphsByType : settings.FooterParagraphsByType;
                var drawings = side == 0 ? settings.HeaderFloatingDrawingsByType : settings.FooterFloatingDrawingsByType;
                foreach (string variant in bodies.Keys.Concat(paragraphs.Keys).Concat(drawings.Keys).Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    token.ThrowIfCancellationRequested();
                    yield return (settings, variant, FindStaticStoryBodyElements(variant, bodies, paragraphs) ?? [],
                        drawings.TryGetValue(variant, out var selected) ? selected : []);
                }
            }
        }
    }

    public static IEnumerable<DocxParagraph> EnumerateResourceParagraphs(DocxDocument document, CancellationToken token)
    {
        foreach (var paragraph in ResolveResourceContent(document, token).Paragraphs)
        {
            token.ThrowIfCancellationRequested();
            yield return paragraph;
        }
    }

    public static IEnumerable<DocxInlineImage> EnumerateResourceImages(DocxDocument document, CancellationToken token)
    {
        var (paragraphs, drawings) = ResolveResourceContent(document, token);
        foreach (var paragraph in paragraphs)
        {
            token.ThrowIfCancellationRequested();
            foreach (var image in paragraph.Images)
                yield return image;
        }
        foreach (var drawing in drawings)
        {
            token.ThrowIfCancellationRequested();
            if (drawing.Image is { } image)
                yield return image;
        }
    }

    private static (IEnumerable<DocxParagraph> Paragraphs, IEnumerable<DocxFloatingDrawing> Drawings)
        ResolveResourceContent(DocxDocument document, CancellationToken token)
    {
        var stories = EnumerateStaticStorySources(document, token)
            .Where(story => DocxLayoutEngine.IsStaticStoryVariantEnabled(story.Settings, story.Variant)).ToArray();
        var drawings = document.FloatingDrawings.Concat(stories.SelectMany(story => story.Drawings))
            .Concat(document.RelatedStories.SelectMany(story => story.FloatingDrawings));
        var parents = EnumerateBodyParagraphs(document)
            .Concat(stories.SelectMany(story => EnumerateBodyParagraphs(story.BodyElements)))
            .Concat(document.RelatedStories.SelectMany(EnumerateBodyParagraphs))
            .Concat(drawings.SelectMany(drawing => EnumerateBodyParagraphs(drawing.TextBoxBodyElements)));
        return (parents.Concat(parents.SelectMany(EnumerateInlineTextBoxParagraphs)), drawings);
    }

    private static IEnumerable<DocxFloatingDrawing> EnumerateSourceFloatingDrawings(DocxDocument document)
    {
        return document.FloatingDrawings
            .Concat(document.HeaderFloatingDrawingsByType.Values.SelectMany(drawings => drawings))
            .Concat(document.FooterFloatingDrawingsByType.Values.SelectMany(drawings => drawings))
            .Concat(document.PageSettings.HeaderFloatingDrawingsByType.Values.SelectMany(drawings => drawings))
            .Concat(document.PageSettings.FooterFloatingDrawingsByType.Values.SelectMany(drawings => drawings))
            .Concat(document.BodyElements
                .OfType<DocxSectionBreakElement>()
                .SelectMany(sectionBreak => sectionBreak.PageSettings.HeaderFloatingDrawingsByType.Values
                    .Concat(sectionBreak.PageSettings.FooterFloatingDrawingsByType.Values)
                    .SelectMany(drawings => drawings)))
            .Concat(document.RelatedStories.SelectMany(story => story.FloatingDrawings));
    }

    public static IEnumerable<DocxParagraph> EnumerateInlineTextBoxParagraphs(DocxParagraph paragraph)
    {
        foreach (DocxInlineTextBox textBox in paragraph.InlineTextBoxes)
        {
            foreach (DocxParagraph boxParagraph in DocxBlockTraversal.EnumerateBodyParagraphs(textBox.BodyElements))
            {
                yield return boxParagraph;
                foreach (DocxParagraph nested in EnumerateInlineTextBoxParagraphs(boxParagraph))
                {
                    yield return nested;
                }
            }
        }
    }




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

    // A defined empty story still owns its variant. Null means it is absent.
    public static IReadOnlyList<DocxBodyElement>? FindStaticStoryBodyElements(
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
            : null;
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
