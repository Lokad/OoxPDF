namespace Lokad.OoxPdf.Docx;

// RV16: comment/reply story lookup built once per conversion from the laid-out
// related stories; per-page candidate collection reads it instead of rebuilding
// group dictionaries for every page.
internal sealed class MarkupCommentStoryIndex
{
    public IReadOnlyDictionary<string, DocxRelatedStoryLayout> CommentStories { get; }
    public IReadOnlyDictionary<string, DocxRelatedStoryLayout[]> CommentRepliesByParentId { get; }

    private MarkupCommentStoryIndex(
        IReadOnlyDictionary<string, DocxRelatedStoryLayout> commentStories,
        IReadOnlyDictionary<string, DocxRelatedStoryLayout[]> commentRepliesByParentId)
    {
        CommentStories = commentStories;
        CommentRepliesByParentId = commentRepliesByParentId;
    }

    public static MarkupCommentStoryIndex Build(IReadOnlyList<DocxRelatedStoryLayout> relatedStories)
    {
        Dictionary<string, DocxRelatedStoryLayout> commentStories = relatedStories
            .Where(story => story.Story.Kind == DocxRelatedStoryKind.Comment && story.Story.Id is not null)
            .GroupBy(story => story.Story.Id ?? string.Empty, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
        Dictionary<string, DocxRelatedStoryLayout[]> commentRepliesByParentId = relatedStories
            .Where(story => story.Story.Kind == DocxRelatedStoryKind.Comment && story.Story.CommentMetadata?.ParentCommentId is not null)
            .GroupBy(story => story.Story.CommentMetadata?.ParentCommentId ?? string.Empty, StringComparer.Ordinal)
            .ToDictionary(
                group => group.Key,
                group => group
                    .OrderBy(story => DocxRenderer.FormatCommentDate(story.Story.CommentMetadata?.Date), StringComparer.Ordinal)
                    .ThenBy(story => story.Story.Id, StringComparer.Ordinal)
                    .ToArray(),
                StringComparer.Ordinal);
        return new MarkupCommentStoryIndex(commentStories, commentRepliesByParentId);
    }
}

