namespace Lokad.OoxPdf.Docx;

internal sealed partial record DocxStructureSnapshot
{
    private static bool HasEffectiveKeepConstraint(DocxParagraph paragraph)
    {
        DocxParagraphKeepRules keepRules = paragraph.EffectiveProperties.KeepRules;
        return keepRules.KeepNext == true || keepRules.KeepLines == true;
    }

    private static int ParagraphInlineReferenceCount(DocxParagraph paragraph)
    {
        return paragraph.InlineReferences.Count;
    }

    private static int CountRevisions(DocxParagraph paragraph)
    {
        return paragraph.Revisions.Count;
    }

    private static int CountRevisions(IEnumerable<DocxRevisionInfo> revisions)
    {
        return revisions.Count();
    }

    private static int CountRevisions(DocxParagraph paragraph, DocxRevisionKind kind)
    {
        return paragraph.Revisions.Count(revision => revision.Kind == kind);
    }

    private static int CountRevisions(IEnumerable<DocxRevisionInfo> revisions, DocxRevisionKind kind)
    {
        return revisions.Count(revision => revision.Kind == kind);
    }

    private static int CountOtherRevisions(DocxParagraph paragraph)
    {
        return paragraph.Revisions.Count(revision =>
            revision.Kind is not (DocxRevisionKind.Insertion or DocxRevisionKind.Deletion or DocxRevisionKind.MoveFrom or DocxRevisionKind.MoveTo));
    }

    private static int CountOtherRevisions(IEnumerable<DocxRevisionInfo> revisions)
    {
        return revisions.Count(revision =>
            revision.Kind is not (DocxRevisionKind.Insertion or DocxRevisionKind.Deletion or DocxRevisionKind.MoveFrom or DocxRevisionKind.MoveTo));
    }

    // Overload pair; C# bans local-function overloads (CS0128); kept static.
    private static int CountFormattingRevisions(IEnumerable<DocxRevisionInfo> revisions)
    {
        return revisions.Count(revision => revision.PropertyChangeFamily is not null);
    }

    // Overload pair; C# bans local-function overloads (CS0128); kept static.
    private static int CountFormattingRevisions(IEnumerable<DocxRevisionInfo> revisions, DocxRevisionPropertyFamily family)
    {
        return revisions.Count(revision => revision.PropertyChangeFamily == family);
    }

    private static int ParagraphResolvedInlineReferenceCount(DocxParagraph paragraph, RelatedStoryIndex relatedStories)
    {
        return paragraph.InlineReferences.Count(reference => relatedStories.Resolve(reference.Kind, reference.Id).Index.HasValue);
    }

    private static int ParagraphFieldReferenceCount(DocxParagraph paragraph)
    {
        return paragraph.FieldReferences.Count;
    }

    private static int ParagraphPageFieldReferenceCount(DocxParagraph paragraph)
    {
        return paragraph.FieldReferences.Count(reference => reference.Kind == DocxFieldKind.Page);
    }

    private static int ParagraphNumPagesFieldReferenceCount(DocxParagraph paragraph)
    {
        return paragraph.FieldReferences.Count(reference => reference.Kind == DocxFieldKind.NumPages);
    }

    private static int ParagraphOtherFieldReferenceCount(DocxParagraph paragraph)
    {
        return paragraph.FieldReferences.Count(reference => reference.Kind == DocxFieldKind.Other);
    }

    // Single caller; kept static for symmetry with the multi-caller counter family.
    private static int ParagraphDynamicFieldReferenceCount(DocxParagraph paragraph)
    {
        return paragraph.FieldReferences.Count(IsDynamicFieldReference);
    }

    // Single caller; kept static for symmetry with the multi-caller counter family.
    private static int ParagraphDynamicPlaceholderFieldReferenceCount(DocxParagraph paragraph)
    {
        return paragraph.FieldReferences.Count(reference => IsDynamicFieldReference(reference) && reference.UsesPlaceholder);
    }

    // Single caller; kept static for symmetry with the multi-caller counter family.
    private static int ParagraphDynamicComplexWithoutCachedResultFieldReferenceCount(DocxParagraph paragraph)
    {
        return paragraph.FieldReferences.Count(reference => IsDynamicFieldReference(reference) && reference.SourceKind == DocxFieldSourceKind.ComplexInstruction && !reference.HasCachedResult);
    }

    // Single caller; kept static for symmetry with the multi-caller counter family.
    private static int ParagraphDynamicCachedResultNotRenderedFieldReferenceCount(DocxParagraph paragraph)
    {
        return paragraph.FieldReferences.Count(reference => IsDynamicFieldReference(reference) && reference.HasCachedResult && !reference.RendersCachedResult);
    }

    private static bool IsDynamicFieldReference(DocxFieldReference reference)
    {
        return reference.Kind.IsDynamic();
    }

    private static int ParagraphComplexFieldReferenceCount(DocxParagraph paragraph)
    {
        return paragraph.FieldReferences.Count(reference => reference.SourceKind == DocxFieldSourceKind.ComplexInstruction);
    }

    private static int ParagraphCachedResultFieldReferenceCount(DocxParagraph paragraph)
    {
        return paragraph.FieldReferences.Count(reference => reference.HasCachedResult);
    }

    private static int ParagraphRenderedCachedResultFieldReferenceCount(DocxParagraph paragraph)
    {
        return paragraph.FieldReferences.Count(reference => reference.RendersCachedResult);
    }

    private static int ParagraphPlaceholderFieldReferenceCount(DocxParagraph paragraph)
    {
        return paragraph.FieldReferences.Count(reference => reference.UsesPlaceholder);
    }

    private static int ParagraphNestedFieldReferenceCount(DocxParagraph paragraph)
    {
        return paragraph.FieldReferences.Count(reference => reference.NestingDepth > 0);
    }

    private static int ParagraphBookmarkAnchorCount(DocxParagraph paragraph)
    {
        return paragraph.BookmarkAnchors.Count;
    }

    private static int ParagraphHyperlinkCount(DocxParagraph paragraph)
    {
        return paragraph.Hyperlinks.Count;
    }

    private static int ParagraphExternalHyperlinkCount(DocxParagraph paragraph)
    {
        return paragraph.Hyperlinks.Count(link => string.Equals(link.TargetMode, "External", StringComparison.OrdinalIgnoreCase));
    }

    private static int ParagraphInternalHyperlinkCount(DocxParagraph paragraph)
    {
        return paragraph.Hyperlinks.Count(link => link.Anchor is not null || link.ResolvedTarget is not null);
    }

    private static bool HasInlineReferenceAnchor(DocxInlineReference reference)
    {
        return reference.SourceRunIndex >= 0 && reference.RunChildIndex >= 0;
    }

    private static int TextLength(DocxParagraph paragraph)
    {
        return paragraph.Runs.Sum(run => run.Text.Length);
    }

    private static int CountCharacters(DocxParagraph paragraph, Func<char, bool> predicate)
    {
        return paragraph.Runs.Sum(run => run.Text.Count(predicate));
    }

    private static int CountWhitespaceDelimitedTokens(DocxParagraph paragraph)
    {
        return EnumerateWhitespaceDelimitedTokenLengths(paragraph).Count();
    }

    private static int LongestWhitespaceDelimitedTokenLength(DocxParagraph paragraph)
    {
        return EnumerateWhitespaceDelimitedTokenLengths(paragraph).DefaultIfEmpty(0).Max();
    }

    private static IEnumerable<int> EnumerateWhitespaceDelimitedTokenLengths(DocxParagraph paragraph)
    {
        int currentLength = 0;
        foreach (DocxTextRun run in paragraph.Runs)
        {
            foreach (char c in run.Text)
            {
                if (char.IsWhiteSpace(c))
                {
                    if (currentLength > 0)
                    {
                        yield return currentLength;
                        currentLength = 0;
                    }

                    continue;
                }

                currentLength++;
            }
        }

        if (currentLength > 0)
        {
            yield return currentLength;
        }
    }

    private static int CountCharacters(IEnumerable<DocxParagraph> paragraphs, Func<char, bool> predicate)
    {
        return paragraphs.Sum(paragraph => CountCharacters(paragraph, predicate));
    }

    private static int MaxColumnCount(DocxTable table)
    {
        return table.Rows.Select(row => row.Cells.Sum(cell => Math.Max(1, cell.GridSpan))).DefaultIfEmpty(0).Max();
    }
}
