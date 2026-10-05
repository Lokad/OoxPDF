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
    internal static string BuildCommentBalloonTitle(DocxRelatedStory? story, string? fallbackId)
    {
        DocxCommentMetadata? metadata = story?.CommentMetadata;
        string? author = FirstNonEmpty(metadata?.Author);
        string? initials = FirstNonEmpty(metadata?.Initials);
        string? owner = author is not null && initials is not null &&
            !string.Equals(author, initials, StringComparison.Ordinal)
            ? author + " (" + initials + ")"
            : author ?? initials;
        string label = owner is null
            ? "Comment " + (fallbackId ?? "?")
            : fallbackId is null ? owner : owner + " #" + fallbackId;
        string? date = FormatCommentDate(metadata?.Date);
        if (metadata?.ParentCommentId is not null)
        {
            label = "Reply " + label;
        }

        if (date is not null)
        {
            label += " " + date;
        }

        if (metadata?.IsResolved is not null)
        {
            label += metadata.IsResolved == true ? " resolved" : " open";
        }

        return label;
    }

    private static IReadOnlyDictionary<string, string>? BuildWordCompatibleCommentMarkerLabels(DocxDocument document, CancellationToken cancellationToken)
    {
        var stories = new Dictionary<string, DocxRelatedStory>(StringComparer.Ordinal);
        foreach (DocxRelatedStory story in document.RelatedStories)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (story.Kind == DocxRelatedStoryKind.Comment && story.Id is { } id)
            {
                stories.TryAdd(id, story);
            }
        }
        var labels = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (DocxParagraph paragraph in DocxBlockTraversal.EnumerateBodyParagraphs(document))
        {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (DocxInlineReference reference in paragraph.InlineReferences)
            {
                if (reference.Kind != DocxRelatedStoryKind.Comment || reference.Id is not { } id ||
                    labels.ContainsKey(id) || !stories.TryGetValue(id, out DocxRelatedStory? story) ||
                    story.CommentMetadata?.ParentCommentId is not null)
                {
                    continue;
                }
                string? initials = FirstNonEmpty(story.CommentMetadata?.Initials, story.CommentMetadata?.Author);
                string number = (labels.Count + 1).ToString(CultureInfo.InvariantCulture);
                labels.Add(id, "[" + (initials ?? "Comment ") + number + "]");
            }
        }
        return labels.Count == 0 ? null : labels;
    }

    private static string BuildWordCompatibleCommentBalloonTitle(DocxRelatedStory? story, string? fallbackId, IReadOnlyDictionary<string, string>? markerLabels = null)
    {
        if (fallbackId is not null && markerLabels?.TryGetValue(fallbackId, out string? label) == true)
        {
            return "Commented " + label + ": ";
        }
        DocxCommentMetadata? metadata = story?.CommentMetadata;
        string? initials = FirstNonEmpty(metadata?.Initials, metadata?.Author);
        string? id = FirstNonEmpty(fallbackId);
        if (initials is not null && id is not null)
        {
            return "Commented [" + initials + id + "]: ";
        }

        if (initials is not null)
        {
            return "Commented [" + initials + "]: ";
        }

        return id is null
            ? "Commented: "
            : "Commented [Comment " + id + "]: ";
    }

    internal static string BuildCommentBalloonPreview(DocxRelatedStoryLayout? storyLayout)
    {
        return BuildCommentBalloonPreview(storyLayout, []);
    }

    internal static string BuildCommentBalloonPreview(
        DocxRelatedStoryLayout? storyLayout,
        IReadOnlyList<DocxRelatedStoryLayout> replies)
    {
        if (storyLayout is null)
        {
            return string.Empty;
        }

        List<string> parts = BuildCommentStoryPreviewParts(storyLayout);
        if (replies.Count != 0)
        {
            parts.Add(replies.Count == 1 ? "1 reply" : replies.Count.ToString(CultureInfo.InvariantCulture) + " replies");
            foreach (string replyPreview in replies
                .Select(reply => string.Join(" ", BuildCommentStoryPreviewParts(reply)))
                .Where(preview => !string.IsNullOrWhiteSpace(preview)))
            {
                parts.Add("Reply: " + replyPreview);
            }
        }

        return string.Join(" ", parts);
    }

    internal static string BuildWordCompatibleCommentBalloonPreview(DocxRelatedStoryLayout? storyLayout)
    {
        if (storyLayout is null)
        {
            return string.Empty;
        }

        // Word 16 prints the parent only, for resolved/open threads and
        // multiple replies. Reply metadata stays in the model and inspection.
        // Office A/B (dense and threaded references): Word-compatible balloons show the
        // comment text with no date or resolved/open flag.
        return string.Join(" ", BuildCommentStoryPreviewParts(storyLayout));
    }

    private static List<string> BuildCommentStoryPreviewParts(DocxRelatedStoryLayout storyLayout)
    {
        var parts = DocxBlockTraversal
            .EnumerateBodyParagraphs(storyLayout.Story)
            .Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text)).Trim())
            .Where(text => text.Length != 0)
            .ToList();
        int tableCount = DocxBlockTraversal.EnumerateBodyTables(storyLayout.Story).Count();
        int inlineImageCount = storyLayout.InlineImages.Count + CountTableInlineImages(storyLayout.TableRows);
        int floatingDrawingCount = storyLayout.FloatingDrawings.Count(drawing => drawing.Drawing.Image is not null || !string.IsNullOrWhiteSpace(drawing.Drawing.ImageRelationshipId));

        if (tableCount != 0)
        {
            parts.Add(tableCount == 1 ? "[table]" : "[" + tableCount.ToString(CultureInfo.InvariantCulture) + " tables]");
        }

        int visualImageCount = inlineImageCount + floatingDrawingCount;
        if (visualImageCount != 0)
        {
            parts.Add(visualImageCount == 1 ? "[image]" : "[" + visualImageCount.ToString(CultureInfo.InvariantCulture) + " images]");
        }

        return parts;
    }

    private static int CountTableInlineImages(IReadOnlyList<DocxTableRowLayout> rows)
    {
        int count = 0;
        foreach (DocxTableRowLayout row in rows)
        {
            foreach (DocxTableCellLayout cell in row.Cells)
            {
                count += cell.InlineImages.Count;
                count += CountTableInlineImages(cell.NestedRows);
            }
        }

        return count;
    }
}
