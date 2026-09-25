using System.Globalization;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    // RV06: Word 16.0 paints comment balloons, body range washes and their
    // brackets per author. The slot is the first-seen order of distinct authors
    // in w:id order, and the 20-entry table below cycles (slot mod 20).
    // Measured from Word-COM balloon exports (ShowRevisions=True RevisionsView=0
    // MarkupMode=2 RevisionsMode=0 ShowComments=True; artifacts/rv06-colors/):
    // - 16-author probe (w:id order Zebra..Amber): entries 1..16 in order;
    // - swap probe (w:id 1=Apple anchored first): Apple wears entry 1, so the
    //   slot follows w:id order rather than the author name or anchor position;
    // - repeat probe (Zebra, Apple, Zebra, Mango, Coral anchored under w:ids
    //   5,1,2,3,4): balloons wear entries 4,1,2,1,3, ruling out anchor order
    //   and w:id-position rules and proving repeats reuse the first-seen entry;
    // - 32-author probe: entries 1..20, then 1..12 again, so the table cycles.
    // Entry 1 matches the long-standing review constants. Byte values are
    // round(OfficeComponent * 255).
    private static readonly DocxMarkupBalloonRgb[] CommentAuthorFillPalette =
    [
        new(248, 220, 221),
        new(213, 237, 255),
        new(234, 223, 244),
        new(236, 253, 215),
        new(247, 221, 237),
        new(222, 218, 250),
        new(214, 253, 254),
        new(255, 247, 213),
        new(252, 216, 219),
        new(217, 223, 251),
        new(242, 223, 244),
        new(232, 235, 236),
        new(217, 251, 217),
        new(252, 228, 216),
        new(252, 216, 221),
        new(235, 234, 233),
        new(245, 218, 250),
        new(232, 235, 236),
        new(213, 247, 255),
        new(244, 233, 223),
    ];

    private static readonly DocxMarkupBalloonRgb[] CommentAuthorStrokePalette =
    [
        new(209, 52, 56),
        new(0, 120, 212),
        new(92, 46, 145),
        new(73, 130, 5),
        new(204, 53, 149),
        new(113, 96, 232),
        new(3, 131, 135),
        new(109, 87, 0),
        new(207, 15, 31),
        new(78, 106, 237),
        new(177, 70, 194),
        new(57, 65, 70),
        new(11, 106, 11),
        new(202, 80, 16),
        new(117, 11, 28),
        new(93, 90, 88),
        new(136, 23, 152),
        new(105, 121, 126),
        new(0, 91, 112),
        new(142, 86, 46),
    ];

    internal static IReadOnlyDictionary<string, int> BuildCommentAuthorPaletteSlots(
        IEnumerable<(string? Id, string? Author)> commentsInFileOrder)
    {
        var slotsById = new Dictionary<string, int>(StringComparer.Ordinal);
        var slotByAuthor = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach ((string? id, string? author) in commentsInFileOrder
            .OrderBy(comment => TryParseCommentAuthorPaletteId(comment.Id, out int numericId) ? numericId : int.MaxValue)
            .ThenBy(comment => comment.Id, StringComparer.Ordinal))
        {
            string key = NormalizeRevisionAuthorBucketKey(author);
            if (!slotByAuthor.TryGetValue(key, out int slot))
            {
                slot = slotByAuthor.Count;
                slotByAuthor[key] = slot;
            }

            if (id is not null)
            {
                slotsById[id] = slot;
            }
        }

        return slotsById;
    }

    internal static IEnumerable<(string? Id, string? Author)> SelectCommentAuthorsForPalette(
        IEnumerable<DocxRelatedStoryLayout> relatedStories)
    {
        foreach (DocxRelatedStoryLayout storyLayout in relatedStories)
        {
            if (storyLayout.Story.Kind != DocxRelatedStoryKind.Comment || storyLayout.Story.Id is null)
            {
                continue;
            }

            yield return (storyLayout.Story.Id, storyLayout.Story.CommentMetadata?.Author);
        }
    }

    internal static int ResolveCommentAuthorSlot(IReadOnlyDictionary<string, int> slotsById, string? commentId)
    {
        return commentId is not null && slotsById.TryGetValue(commentId, out int slot) ? slot : 0;
    }

    private static DocxMarkupBalloonRgb CommentAuthorFillRgb(int slot)
    {
        return CommentAuthorFillPalette[PositivePaletteIndex(slot, CommentAuthorFillPalette.Length)];
    }

    private static DocxMarkupBalloonRgb CommentAuthorStrokeRgb(int slot)
    {
        return CommentAuthorStrokePalette[PositivePaletteIndex(slot, CommentAuthorStrokePalette.Length)];
    }

    internal static (byte Red, byte Green, byte Blue) ResolveCommentAuthorFillSnapshot(
        IReadOnlyDictionary<string, int> slotsById,
        string? commentId)
    {
        DocxMarkupBalloonRgb color = CommentAuthorFillRgb(ResolveCommentAuthorSlot(slotsById, commentId));
        return (color.Red, color.Green, color.Blue);
    }

    internal static (byte Red, byte Green, byte Blue) ResolveCommentAuthorStrokeSnapshot(
        IReadOnlyDictionary<string, int> slotsById,
        string? commentId)
    {
        DocxMarkupBalloonRgb color = CommentAuthorStrokeRgb(ResolveCommentAuthorSlot(slotsById, commentId));
        return (color.Red, color.Green, color.Blue);
    }

    private static bool TryParseCommentAuthorPaletteId(string? id, out int numericId)
    {
        numericId = 0;
        return id is not null && int.TryParse(id, NumberStyles.Integer, CultureInfo.InvariantCulture, out numericId);
    }

    private static int PositivePaletteIndex(int slot, int length)
    {
        return ((slot % length) + length) % length;
    }
}
