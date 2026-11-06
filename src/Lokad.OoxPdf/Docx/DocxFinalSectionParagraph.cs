namespace Lokad.OoxPdf.Docx;

internal static class DocxFinalSectionParagraph
{
    internal static DocxDocument Materialize(DocxDocument document, CancellationToken token, double printScale = 1d)
    {
        if (printScale != 1d || !DocxPlainBodyFrame.HasKnownBodyFrame(document) ||
            !DocxPlainBodyFrame.HasNoStaticStories(document) || document.EmptyFinalSectionParagraph is not { } paragraph ||
            !DocxPlainBodyFrame.IsPlainParagraph(paragraph) || !OwnsDocumentEndnotes(document, token)) return document;
        return document with { BodyElements = [.. document.BodyElements, new DocxParagraphElement(paragraph)] };
    }

    // A closing sectPr owns the section before it. Scan backward so each
    // reference uses the note mode of the section closed by its next sectPr.
    // This rule also retains a pending frame when its default paragraph needs
    // an unsupported fallback; an unreferenced story cannot create a page.
    internal static bool OwnsDocumentEndnotes(DocxDocument document, CancellationToken token)
    {
        if (document.BodyElements.LastOrDefault() is not DocxSectionBreakElement ||
            document.FinalSectionBreak is not { TypeValue: null or DocxSectionBreakType.NextPage }) return false;
        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var story in document.RelatedStories)
        {
            token.ThrowIfCancellationRequested();
            if (story.Kind == DocxRelatedStoryKind.Endnote && story.IsNormalStoryType && story.Id is { } id) ids.Add(id);
        }
        if (ids.Count == 0) return false;
        var settings = document.FinalSectionBreak.PageSettings;
        for (int i = document.BodyElements.Count - 1; i >= 0; i--)
        {
            token.ThrowIfCancellationRequested();
            var element = document.BodyElements[i];
            if (element is DocxSectionBreakElement section) { settings = section.PageSettings; continue; }
            if (string.Equals(settings.EndnoteReferenceSettings.PositionValue, "sectEnd", StringComparison.OrdinalIgnoreCase)) continue;
            foreach (var paragraph in DocxBlockTraversal.EnumerateBodyParagraphs([element]))
            {
                token.ThrowIfCancellationRequested();
                foreach (var reference in paragraph.InlineReferences)
                {
                    token.ThrowIfCancellationRequested();
                    if (reference.Kind == DocxRelatedStoryKind.Endnote && reference.Id is { } id && ids.Contains(id)) return true;
                }
            }
        }
        return false;
    }
}
