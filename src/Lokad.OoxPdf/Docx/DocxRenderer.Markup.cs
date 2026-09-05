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
    private static void RenderMarkupBalloons(
        DocxLayoutPage page,
        IReadOnlyList<DocxRelatedStoryLayout> relatedStories,
        IReadOnlyList<DocxFloatingDrawingLayout> floatingDrawings,
        PdfGraphicsBuilder graphics,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext)
    {
        DocxRunFontResource? labelResource = ResolveMarkupLabelFontResource(fontResources);
        if (labelResource is null)
        {
            return;
        }

        DocxRunFontResource bodyResource = ResolveMarkupBodyFontResource(fontResources) ?? labelResource;
        foreach (DocxMarkupBalloonPlacement placement in BuildMarkupBalloonPlacements(page, relatedStories, floatingDrawings, markupContext))
        {
            RenderMarkupBalloonPlacement(placement, graphics, labelResource, bodyResource, markupContext);
        }
    }

    private static DocxRunFontResource? ResolveMarkupLabelFontResource(DocxFontResources fontResources)
    {
        return fontResources.Fallback ?? fontResources.RunResources.Values.FirstOrDefault();
    }

    private static DocxRunFontResource? ResolveMarkupBodyFontResource(DocxFontResources fontResources)
    {
        return fontResources.RunResources.Values.FirstOrDefault(resource => !resource.Resolution.Bold) ??
            fontResources.Fallback;
    }

    private static IEnumerable<DocxFloatingDrawingLayout> EnumeratePageFloatingDrawings(DocxLayout layout, int pageIndex)
    {
        return layout.FloatingDrawings
            .Concat(layout.StaticFloatingDrawings)
            .Where(drawing => drawing.AnchorPageIndex == pageIndex);
    }

    private static IReadOnlyList<DocxMarkupBalloonPlacement> BuildMarkupBalloonPlacements(
        DocxLayoutPage page,
        IReadOnlyList<DocxRelatedStoryLayout> relatedStories,
        IReadOnlyList<DocxFloatingDrawingLayout> floatingDrawings,
        DocxMarkupContext markupContext)
    {
        if (!markupContext.RendersCommentBalloons && !markupContext.RendersRevisionBalloons)
        {
            return [];
        }

        DocxMarkupBalloonArea area = ResolveMarkupBalloonArea(page);
        DocxMarkupBalloonCandidate[] candidates = OrderMarkupBalloonCandidatesForPlacement(
                GroupNearbyMarkupBalloonCandidates(
                    OrderMarkupBalloonCandidatesForPlacement(CollectMarkupBalloonCandidates(page, relatedStories, floatingDrawings, markupContext, area.Width)),
                    area.Width))
            .ToArray();
        if (candidates.Length == 0)
        {
            return [];
        }

        IReadOnlyList<DocxMarkupBalloonLaneBand> laneBands = BuildMarkupBalloonLaneBands(candidates, page, markupContext);
        var placements = new List<DocxMarkupBalloonPlacement>();
        double nextTop = page.Height - page.MarginTop;
        int nextOverflowStartIndex = 1;
        foreach (DocxMarkupBalloonLaneBand laneBand in laneBands)
        {
            double bandTopLimit = Math.Max(laneBand.TopLimit, page.MarginBottom + laneBand.MaxBalloonHeight);
            nextTop = Math.Min(nextTop, bandTopLimit);
            var overflowCandidates = new List<DocxMarkupBalloonCandidate>();
            foreach (DocxMarkupBalloonCandidate candidate in laneBand.Candidates)
            {
                double anchorY = ResolveMarkupBalloonAnchorY(candidate.AnchorY, page, markupContext);
                double height = ResolveMarkupBalloonHeight(candidate, markupContext);
                double topInset = ResolveMarkupBalloonTopInset(markupContext);
                double desiredTop = Math.Min(nextTop, anchorY + topInset);
                double y = desiredTop - height;
                if (y < page.MarginBottom)
                {
                    if (nextTop - height < page.MarginBottom)
                    {
                        overflowCandidates.Add(candidate);
                        continue;
                    }

                    y = page.MarginBottom;
                }

                int nearbyAnchorConnectorCount = placements.Count(placement =>
                    !placement.IsOverflowSummary &&
                    placement.LaneBandIndex == laneBand.Index &&
                    Math.Abs(placement.AnchorY - anchorY) < MarkupBalloonConnectorCollisionAnchorYThresholdPoints);
                double connectorAnchorX = candidate.AnchorConnectorX;
                connectorAnchorX += area.Side == "Left"
                    ? -nearbyAnchorConnectorCount * 1.5d
                    : nearbyAnchorConnectorCount * 1.5d;
                double unclampedConnectorAnchorX = connectorAnchorX;
                connectorAnchorX = ClampMarkupBalloonConnectorAnchorX(connectorAnchorX, page);
                bool anchorConnectorClamped = Math.Abs(connectorAnchorX - unclampedConnectorAnchorX) > 0.001d;
                placements.Add(new DocxMarkupBalloonPlacement(
                    candidate.Kind,
                    area.Side,
                    candidate.Title,
                    candidate.Body,
                    candidate.WordCompatibleTitle,
                    candidate.WordCompatibleBody,
                    area.X,
                    y,
                    area.Width,
                    height,
                    anchorY,
                    connectorAnchorX,
                    area.ConnectorX,
                    anchorConnectorClamped,
                    candidate.FillRgb,
                    candidate.StrokeRgb,
                    candidate.TitleRgb,
                    candidate.BodyRgb,
                    IsOverflowSummary: false,
                    candidate.CandidateCount,
                    candidate.CommentCandidateCount,
                    candidate.RevisionCandidateCount,
                    candidate.CommentWithDateCount,
                    candidate.CommentResolvedCount,
                    candidate.CommentOpenCount,
                    candidate.CommentReplyCount,
                    candidate.BodySummaryPartCount,
                    candidate.WordCompatibleBodySummaryPartCount,
                    OverflowStartIndex: null, OverflowEndIndex: null, LaneBandIndex: laneBand.Index,
                    LaneBandCandidateCount: laneBand.CandidateCount));
                nextTop = y - MarkupBalloonMinimumSpacingPoints;
            }

            (nextTop, nextOverflowStartIndex) = AddOverflowContinuationPlacements(
                overflowCandidates,
                placements,
                area,
                page,
                nextTop,
                laneBand.Index,
                laneBand.CandidateCount,
                nextOverflowStartIndex);
        }

        return placements;
    }

    private static IReadOnlyList<DocxMarkupBalloonLaneBand> BuildMarkupBalloonLaneBands(
        IReadOnlyList<DocxMarkupBalloonCandidate> candidates,
        DocxLayoutPage page,
        DocxMarkupContext markupContext)
    {
        var bands = new List<DocxMarkupBalloonLaneBand>();
        var current = new List<DocxMarkupBalloonCandidate>();
        double currentTopLimit = 0d;
        double currentBottom = 0d;
        double currentMaxBalloonHeight = 0d;

        foreach (DocxMarkupBalloonCandidate candidate in candidates)
        {
            double anchorY = ResolveMarkupBalloonAnchorY(candidate.AnchorY, page, markupContext);
            double height = ResolveMarkupBalloonHeight(candidate, markupContext);
            double preferredTop = Math.Min(page.Height - page.MarginTop, anchorY + ResolveMarkupBalloonTopInset(markupContext));
            double preferredBottom = preferredTop - height;
            if (current.Count != 0 &&
                currentBottom - preferredTop > MarkupBalloonLaneBandSeparationPoints)
            {
                AddCurrentBand();
                current.Clear();
            }

            if (current.Count == 0)
            {
                currentTopLimit = preferredTop;
                currentBottom = preferredBottom;
                currentMaxBalloonHeight = height;
            }
            else
            {
                currentTopLimit = Math.Max(currentTopLimit, preferredTop);
                currentBottom = Math.Min(currentBottom, preferredBottom);
                currentMaxBalloonHeight = Math.Max(currentMaxBalloonHeight, height);
            }

            current.Add(candidate);
        }

        AddCurrentBand();
        return bands;

        void AddCurrentBand()
        {
            if (current.Count == 0)
            {
                return;
            }

            int index = bands.Count;
            bands.Add(new DocxMarkupBalloonLaneBand(
                index,
                current.ToArray(),
                currentTopLimit,
                currentMaxBalloonHeight,
                current.Sum(candidate => candidate.CandidateCount)));
        }
    }

    private static IEnumerable<DocxMarkupBalloonCandidate> OrderMarkupBalloonCandidatesForPlacement(
        IEnumerable<DocxMarkupBalloonCandidate> candidates)
    {
        DocxMarkupBalloonCandidate[] byAnchor = candidates
            .OrderByDescending(candidate => candidate.AnchorY)
            .ThenBy(candidate => candidate.Sequence)
            .ToArray();
        var band = new List<DocxMarkupBalloonCandidate>();
        double bandTopAnchorY = 0d;
        foreach (DocxMarkupBalloonCandidate candidate in byAnchor)
        {
            if (band.Count == 0)
            {
                bandTopAnchorY = candidate.AnchorY;
                band.Add(candidate);
                continue;
            }

            if (bandTopAnchorY - candidate.AnchorY <= MarkupBalloonConnectorCollisionAnchorYThresholdPoints)
            {
                band.Add(candidate);
                continue;
            }

            foreach (DocxMarkupBalloonCandidate orderedCandidate in OrderMarkupBalloonCandidateBand(band))
            {
                yield return orderedCandidate;
            }

            band.Clear();
            bandTopAnchorY = candidate.AnchorY;
            band.Add(candidate);
        }

        foreach (DocxMarkupBalloonCandidate orderedCandidate in OrderMarkupBalloonCandidateBand(band))
        {
            yield return orderedCandidate;
        }
    }

    private static IEnumerable<DocxMarkupBalloonCandidate> OrderMarkupBalloonCandidateBand(
        IReadOnlyList<DocxMarkupBalloonCandidate> band)
    {
        return band
            .OrderBy(candidate => candidate.Kind)
            .ThenByDescending(candidate => candidate.AnchorY)
            .ThenBy(candidate => candidate.Sequence);
    }


    private static double ClampMarkupBalloonConnectorAnchorX(double connectorAnchorX, DocxLayoutPage page)
    {
        return Math.Min(page.Width - 0.5d, Math.Max(0.5d, connectorAnchorX));
    }

    private static double ResolveMarkupBalloonAnchorY(
        double anchorY,
        DocxLayoutPage page,
        DocxMarkupContext markupContext)
    {
        if (markupContext.Mode != OoxPdfDocxMarkupMode.AllMarkup ||
            markupContext.GeometryMode != OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup ||
            !markupContext.ExpandsMarkupMargin)
        {
            return anchorY;
        }

        double resolved = anchorY - WordCompatibleAllMarkupBalloonAnchorYOffsetPoints;
        return Math.Min(page.Height - page.MarginTop, Math.Max(page.MarginBottom, resolved));
    }

    private static double ResolveMarkupBalloonHeight(
        DocxMarkupBalloonCandidate candidate,
        DocxMarkupContext markupContext)
    {
        if (markupContext.Mode == OoxPdfDocxMarkupMode.AllMarkup &&
            markupContext.GeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup &&
            markupContext.ExpandsMarkupMargin)
        {
            return WordCompatibleAllMarkupBalloonHeightPoints + ResolveWordCompatibleCommentThreadExtraHeight(candidate.CommentReplyCount);
        }

        return string.IsNullOrWhiteSpace(candidate.Body) ? 16d : 26d;
    }

    private static double ResolveWordCompatibleCommentThreadExtraHeight(int replyCount)
    {
        return ResolveCommentThreadSeparatorLineCount(replyCount) * WordCompatibleAllMarkupCommentThreadReplyHeightPoints;
    }

    private static int ResolveCommentThreadSeparatorLineCount(int replyCount)
    {
        return Math.Min(
            WordCompatibleAllMarkupCommentThreadMaxSeparatorLineCount,
            Math.Max(0, replyCount));
    }

    private static double ResolveMarkupBalloonTopInset(DocxMarkupContext markupContext)
    {
        return markupContext.Mode == OoxPdfDocxMarkupMode.AllMarkup &&
            markupContext.GeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup &&
            markupContext.ExpandsMarkupMargin
            ? WordCompatibleAllMarkupBalloonTopInsetPoints
            : 10d;
    }

    private static IReadOnlyList<DocxMarkupBalloonCandidate> GroupNearbyMarkupBalloonCandidates(
        IEnumerable<DocxMarkupBalloonCandidate> orderedCandidates,
        double textWidth)
    {
        var grouped = new List<DocxMarkupBalloonCandidate>();
        var group = new List<DocxMarkupBalloonCandidate>();
        foreach (DocxMarkupBalloonCandidate candidate in orderedCandidates)
        {
            if (group.Count == 0 ||
                CanGroupMarkupBalloonCandidates(group, candidate))
            {
                group.Add(candidate);
                continue;
            }

            grouped.Add(MergeMarkupBalloonGroup(group, textWidth));
            group.Clear();
            group.Add(candidate);
        }

        if (group.Count != 0)
        {
            grouped.Add(MergeMarkupBalloonGroup(group, textWidth));
        }

        return grouped;
    }

    private static bool CanGroupMarkupBalloonCandidates(IReadOnlyList<DocxMarkupBalloonCandidate> group, DocxMarkupBalloonCandidate candidate)
    {
        DocxMarkupBalloonCandidate previous = group[^1];
        if (Math.Abs(previous.AnchorY - candidate.AnchorY) <= 0.5d)
        {
            return true;
        }

        return group.Count < MarkupBalloonMaxNearbyRevisionGroupSize &&
            group.All(item => item.Kind == DocxMarkupBalloonKind.Revision) &&
            previous.Kind == DocxMarkupBalloonKind.Revision &&
            candidate.Kind == DocxMarkupBalloonKind.Revision &&
            Math.Abs(previous.AnchorY - candidate.AnchorY) <= 9d;
    }

    private static DocxMarkupBalloonCandidate MergeMarkupBalloonGroup(
        IReadOnlyList<DocxMarkupBalloonCandidate> group,
        double textWidth)
    {
        if (group.Count == 1)
        {
            return group[0];
        }

        bool allRevisions = group.All(candidate => candidate.Kind == DocxMarkupBalloonKind.Revision);
        string title = allRevisions
            ? group.Count.ToString(CultureInfo.InvariantCulture) + " tracked changes"
            : group.Count.ToString(CultureInfo.InvariantCulture) + " markup items";
        string[] bodyParts = BuildMarkupBalloonGroupBodyParts(group);
        string body = TrimBalloonText(string.Join("; ", bodyParts), textWidth);
        string[] wordCompatibleBodyParts = allRevisions
            ? bodyParts
            : BuildWordCompatibleMarkupBalloonGroupBodyParts(group);
        string? wordCompatibleTitle = allRevisions
            ? TrimBalloonText(title, textWidth)
            : group
                .Select(candidate => candidate.WordCompatibleTitle)
                .FirstOrDefault(title => !string.IsNullOrWhiteSpace(title));
        string? wordCompatibleBody = allRevisions
            ? body
            : wordCompatibleBodyParts.Length == 0 ? null : TrimBalloonText(string.Join("; ", wordCompatibleBodyParts), textWidth);
        return group[0] with
        {
            Kind = group.Select(candidate => candidate.Kind).Distinct().Count() == 1
                ? group[0].Kind
                : DocxMarkupBalloonKind.Markup,
            Title = TrimBalloonText(title, textWidth),
            Body = body,
            WordCompatibleTitle = wordCompatibleTitle,
            WordCompatibleBody = wordCompatibleBody,
            AnchorY = group.Max(candidate => candidate.AnchorY),
            AnchorConnectorX = allRevisions
                ? group.Select(candidate => candidate.AnchorConnectorX).DefaultIfEmpty(group[0].AnchorConnectorX).Average()
                : group[0].AnchorConnectorX,
            AnchorLeftX = group.Min(candidate => candidate.AnchorLeftX),
            AnchorRightX = group.Max(candidate => candidate.AnchorRightX),
            CandidateCount = group.Sum(candidate => candidate.CandidateCount),
            CommentCandidateCount = group.Sum(candidate => candidate.CommentCandidateCount),
            RevisionCandidateCount = group.Sum(candidate => candidate.RevisionCandidateCount),
            CommentWithDateCount = group.Sum(candidate => candidate.CommentWithDateCount),
            CommentResolvedCount = group.Sum(candidate => candidate.CommentResolvedCount),
            CommentOpenCount = group.Sum(candidate => candidate.CommentOpenCount),
            CommentReplyCount = group.Sum(candidate => candidate.CommentReplyCount),
            BodySummaryPartCount = bodyParts.Length,
            WordCompatibleBodySummaryPartCount = allRevisions ? bodyParts.Length : wordCompatibleBodyParts.Length
        };
    }

    private static string[] BuildMarkupBalloonGroupBodyParts(IReadOnlyList<DocxMarkupBalloonCandidate> group)
    {
        return group
            .Select(BuildMarkupBalloonGroupBodyPart)
            .Where(part => !string.IsNullOrWhiteSpace(part))
            .OfType<string>()
            .Distinct(StringComparer.Ordinal)
            .ToArray();
    }

    private static string? BuildMarkupBalloonGroupBodyPart(DocxMarkupBalloonCandidate candidate)
    {
        string? body = FirstNonEmpty(candidate.Body);
        string? title = FirstNonEmpty(candidate.Title);
        if (body is null)
        {
            return title;
        }

        return title is null ? body : title + ": " + body;
    }

    private static string[] BuildWordCompatibleMarkupBalloonGroupBodyParts(IReadOnlyList<DocxMarkupBalloonCandidate> group)
    {
        return group
            .Select(BuildWordCompatibleMarkupBalloonGroupBodyPart)
            .Where(part => !string.IsNullOrWhiteSpace(part))
            .OfType<string>()
            .Distinct(StringComparer.Ordinal)
            .ToArray();
    }

    private static string? BuildWordCompatibleMarkupBalloonGroupBodyPart(DocxMarkupBalloonCandidate candidate)
    {
        string? body = FirstNonEmpty(candidate.WordCompatibleBody, candidate.Body);
        if (body is not null)
        {
            return body;
        }

        return candidate.Kind == DocxMarkupBalloonKind.Comment
            ? null
            : FirstNonEmpty(candidate.WordCompatibleTitle, candidate.Title);
    }

    private static int CountBalloonSummaryPart(string? body)
    {
        return string.IsNullOrWhiteSpace(body) ? 0 : 1;
    }

    private static (double NextTop, int NextOverflowStartIndex) AddOverflowContinuationPlacements(
        IReadOnlyList<DocxMarkupBalloonCandidate> overflowCandidates,
        List<DocxMarkupBalloonPlacement> placements,
        DocxMarkupBalloonArea area,
        DocxLayoutPage page,
        double nextTop,
        int laneBandIndex,
        int laneBandCandidateCount,
        int nextOverflowStartIndex)
    {
        if (overflowCandidates.Count == 0)
        {
            return (nextTop, nextOverflowStartIndex);
        }

        const double height = 12d;
        const double gap = 2d;
        int slotCount = Math.Max(0, (int)Math.Floor((nextTop - page.MarginBottom + gap) / (height + gap)));
        if (slotCount == 0)
        {
            return (nextTop, nextOverflowStartIndex);
        }

        int continuationCount = Math.Min(slotCount, overflowCandidates.Count);
        int chunkSize = (int)Math.Ceiling(overflowCandidates.Count / (double)continuationCount);
        int overflowIndex = 0;
        for (int continuationIndex = 0; continuationIndex < continuationCount && overflowIndex < overflowCandidates.Count; continuationIndex++)
        {
            DocxMarkupBalloonCandidate[] chunk = overflowCandidates
                .Skip(overflowIndex)
                .Take(chunkSize)
                .ToArray();
            if (chunk.Length == 0)
            {
                break;
            }

            double y = nextTop - height;
            if (y < page.MarginBottom)
            {
                break;
            }

            string title = continuationCount == 1
                ? "More markup"
                : "More markup " + (continuationIndex + 1).ToString(CultureInfo.InvariantCulture) + "/" + continuationCount.ToString(CultureInfo.InvariantCulture);
            string body = TrimBalloonText(BuildOverflowContinuationBody(chunk), area.Width);
            placements.Add(new DocxMarkupBalloonPlacement(
                DocxMarkupBalloonKind.Overflow,
                area.Side,
                title,
                body,
                null,
                null,
                area.X,
                y,
                area.Width,
                height,
                chunk[0].AnchorY,
                area.ConnectorX,
                area.ConnectorX,
                false,
                new DocxMarkupBalloonRgb(245, 245, 245),
                new DocxMarkupBalloonRgb(120, 120, 120),
                new DocxMarkupBalloonRgb(70, 70, 70),
                new DocxMarkupBalloonRgb(0, 0, 0),
                IsOverflowSummary: true,
                CandidateCount: chunk.Sum(candidate => candidate.CandidateCount),
                CommentCandidateCount: chunk.Sum(candidate => candidate.CommentCandidateCount),
                RevisionCandidateCount: chunk.Sum(candidate => candidate.RevisionCandidateCount),
                CommentWithDateCount: chunk.Sum(candidate => candidate.CommentWithDateCount),
                CommentResolvedCount: chunk.Sum(candidate => candidate.CommentResolvedCount),
                CommentOpenCount: chunk.Sum(candidate => candidate.CommentOpenCount),
                CommentReplyCount: chunk.Sum(candidate => candidate.CommentReplyCount),
                BodySummaryPartCount: 0, WordCompatibleBodySummaryPartCount: 0, OverflowStartIndex: nextOverflowStartIndex + overflowIndex,
                OverflowEndIndex: nextOverflowStartIndex + overflowIndex + chunk.Length - 1,
                LaneBandIndex: laneBandIndex,
                LaneBandCandidateCount: laneBandCandidateCount));
            overflowIndex += chunk.Length;
            nextTop = y - gap;
        }

        return (nextTop, nextOverflowStartIndex + overflowIndex);
    }

    private static string BuildOverflowContinuationBody(IReadOnlyList<DocxMarkupBalloonCandidate> candidates)
    {
        int count = candidates.Sum(candidate => candidate.CandidateCount);
        string sample = candidates
            .Select(candidate => candidate.Body)
            .FirstOrDefault(body => !string.IsNullOrWhiteSpace(body)) ?? string.Empty;
        return string.IsNullOrWhiteSpace(sample)
            ? count.ToString(CultureInfo.InvariantCulture)
            : count.ToString(CultureInfo.InvariantCulture) + ": " + sample;
    }

    private static DocxCommentThreadBalloonMetrics CountCommentThreadBalloonMetrics(
        DocxRelatedStoryLayout? storyLayout,
        IReadOnlyList<DocxRelatedStoryLayout> replies)
    {
        int withDateCount = 0;
        int resolvedCount = 0;
        int openCount = 0;

        Count(storyLayout?.Story.CommentMetadata);
        foreach (DocxRelatedStoryLayout reply in replies)
        {
            Count(reply.Story.CommentMetadata);
        }

        return new DocxCommentThreadBalloonMetrics(withDateCount, resolvedCount, openCount, replies.Count);

        void Count(DocxCommentMetadata? metadata)
        {
            if (metadata is null)
            {
                return;
            }

            if (!string.IsNullOrWhiteSpace(metadata.Date))
            {
                withDateCount++;
            }

            if (metadata.IsResolved == true)
            {
                resolvedCount++;
            }
            else if (metadata.IsResolved == false)
            {
                openCount++;
            }
        }
    }

    private static IReadOnlyList<DocxMarkupBalloonCandidate> CollectMarkupBalloonCandidates(
        DocxLayoutPage page,
        IReadOnlyList<DocxRelatedStoryLayout> relatedStories,
        IReadOnlyList<DocxFloatingDrawingLayout> floatingDrawings,
        DocxMarkupContext markupContext,
        double textWidth)
    {
        var candidates = new List<DocxMarkupBalloonCandidate>();
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
                    .OrderBy(story => FormatCommentDate(story.Story.CommentMetadata?.Date), StringComparer.Ordinal)
                    .ThenBy(story => story.Story.Id, StringComparer.Ordinal)
                    .ToArray(),
                StringComparer.Ordinal);
        var renderedComments = new HashSet<string>(StringComparer.Ordinal);
        var renderedRevisions = new HashSet<int>();
        var renderedTableRevisions = new HashSet<string>(StringComparer.Ordinal);
        var renderedTableRowRevisions = new HashSet<string>(StringComparer.Ordinal);
        var renderedTableCellRevisions = new HashSet<string>(StringComparer.Ordinal);
        DocxTextLineLayout[] anchorTextLines = EnumerateMarkupBalloonAnchorTextLines(page, floatingDrawings).ToArray();
        int sequence = 0;
        foreach (DocxTextLineLayout line in anchorTextLines)
        {
            if (line.SourceParagraph is not { } paragraph)
            {
                continue;
            }

            if (markupContext.RendersCommentBalloons)
            {
                foreach (DocxInlineReference reference in paragraph.InlineReferences.Where(reference => reference.Kind == DocxRelatedStoryKind.Comment))
                {
                    string key = RuntimeHelpers.GetHashCode(paragraph).ToString(CultureInfo.InvariantCulture) + ":" + (reference.Id ?? string.Empty);
                    if (!renderedComments.Add(key))
                    {
                        continue;
                    }

                    commentStories.TryGetValue(reference.Id ?? string.Empty, out DocxRelatedStoryLayout? storyLayout);
                    commentRepliesByParentId.TryGetValue(reference.Id ?? string.Empty, out DocxRelatedStoryLayout[]? replies);
                    string commentBody = TrimBalloonText(BuildCommentBalloonPreview(storyLayout, replies ?? []), textWidth);
                    string wordCompatibleCommentBody = TrimBalloonText(BuildWordCompatibleCommentBalloonPreview(storyLayout, replies ?? []), textWidth);
                    DocxCommentThreadBalloonMetrics commentMetrics = CountCommentThreadBalloonMetrics(storyLayout, replies ?? []);
                    DocxTextLineLayout anchorLine = ResolveCommentAnchorLine(line, anchorTextLines, paragraph, reference);
                    candidates.Add(new DocxMarkupBalloonCandidate(
                        DocxMarkupBalloonKind.Comment,
                        TrimBalloonText(BuildCommentBalloonTitle(storyLayout?.Story, reference.Id), textWidth),
                        commentBody,
                        BuildWordCompatibleCommentBalloonTitle(storyLayout?.Story, reference.Id),
                        wordCompatibleCommentBody,
                        anchorLine.BaselineY,
                        ResolveCommentAnchorX(anchorLine, anchorTextLines, paragraph, reference, markupContext),
                        anchorLine.X - 2d,
                        anchorLine.X + Math.Max(0d, anchorLine.Width) + 2d,
                        sequence++,
                        new DocxMarkupBalloonRgb(255, 250, 220),
                        new DocxMarkupBalloonRgb(217, 151, 0),
                        new DocxMarkupBalloonRgb(70, 70, 70),
                        new DocxMarkupBalloonRgb(0, 0, 0),
                        CandidateCount: 1, RevisionCandidateCount: 0, CommentCandidateCount: 1,
                        CommentWithDateCount: commentMetrics.WithDateCount,
                        CommentResolvedCount: commentMetrics.ResolvedCount,
                        CommentOpenCount: commentMetrics.OpenCount,
                        CommentReplyCount: commentMetrics.ReplyCount,
                        BodySummaryPartCount: CountBalloonSummaryPart(commentBody),
                        WordCompatibleBodySummaryPartCount: CountBalloonSummaryPart(commentBody)));
                }
            }

            if (markupContext.RendersRevisionBalloons && paragraph.Revisions.Count != 0)
            {
                int key = RuntimeHelpers.GetHashCode(paragraph);
                if (!renderedRevisions.Add(key))
                {
                    continue;
                }

                DocxRevisionMarkupPalette paragraphRevisionPalette = ResolveRevisionMarkupPalette(paragraph.Revisions);
                string revisionTitle = TrimBalloonText(BuildRevisionBalloonTitle(paragraph.Revisions), textWidth);
                string revisionBody = TrimBalloonText(BuildRevisionBalloonPreview(paragraph), textWidth);
                candidates.Add(new DocxMarkupBalloonCandidate(
                    DocxMarkupBalloonKind.Revision,
                    revisionTitle,
                    revisionBody,
                    revisionTitle,
                    revisionBody,
                    line.BaselineY,
                    line.X + Math.Max(0d, line.Width) * 0.5d,
                    line.X - 2d,
                    line.X + Math.Max(0d, line.Width) + 2d,
                    sequence++,
                    paragraphRevisionPalette.FillRgb,
                    paragraphRevisionPalette.StrokeRgb,
                    paragraphRevisionPalette.TitleRgb,
                    new DocxMarkupBalloonRgb(0, 0, 0),
                    CandidateCount: 1, CommentCandidateCount: 0, CommentWithDateCount: 0, CommentResolvedCount: 0, CommentOpenCount: 0, CommentReplyCount: 0, RevisionCandidateCount: 1,
                    BodySummaryPartCount: CountBalloonSummaryPart(revisionBody),
                    WordCompatibleBodySummaryPartCount: CountBalloonSummaryPart(revisionBody)));
            }
        }

        if (markupContext.RendersRevisionBalloons)
        {
            foreach (DocxTableRowLayout row in EnumerateMarkupBalloonTableRows(page, floatingDrawings))
            {
                IReadOnlyList<DocxRevisionInfo> tableRevisions = row.Table.Revisions ?? [];
                if (tableRevisions.Count != 0)
                {
                    string tableKey = TableBalloonKey(row);
                    if (renderedTableRevisions.Add(tableKey))
                    {
                        DocxRevisionMarkupPalette tableRevisionPalette = ResolveRevisionMarkupPalette(tableRevisions);
                        string revisionTitle = TrimBalloonText(BuildRevisionBalloonTitle(tableRevisions), textWidth);
                        string revisionBody = TrimBalloonText(BuildRevisionBalloonPreview(tableRevisions), textWidth);
                        candidates.Add(new DocxMarkupBalloonCandidate(
                            DocxMarkupBalloonKind.Revision,
                            revisionTitle,
                            revisionBody,
                            revisionTitle,
                            revisionBody,
                            row.Y + Math.Max(0d, row.Height) * 0.5d,
                            row.Table.TableX + Math.Max(0d, row.Table.ResolvedTableWidth) * 0.5d,
                            row.Table.TableX - 2d,
                            row.Table.TableX + Math.Max(0d, row.Table.ResolvedTableWidth) + 2d,
                            sequence++,
                            tableRevisionPalette.FillRgb,
                            tableRevisionPalette.StrokeRgb,
                            tableRevisionPalette.TitleRgb,
                            new DocxMarkupBalloonRgb(0, 0, 0),
                            CandidateCount: 1, CommentCandidateCount: 0, CommentWithDateCount: 0, CommentResolvedCount: 0, CommentOpenCount: 0, CommentReplyCount: 0, RevisionCandidateCount: 1,
                            BodySummaryPartCount: CountBalloonSummaryPart(revisionBody),
                            WordCompatibleBodySummaryPartCount: CountBalloonSummaryPart(revisionBody)));
                    }
                }

                IReadOnlyList<DocxRevisionInfo> rowRevisions = row.Revisions ?? [];
                if (rowRevisions.Count != 0)
                {
                    string rowKey = TableRowBalloonKey(row);
                    if (renderedTableRowRevisions.Add(rowKey))
                    {
                        DocxRevisionMarkupPalette rowRevisionPalette = ResolveRevisionMarkupPalette(rowRevisions);
                        string revisionTitle = TrimBalloonText(BuildRevisionBalloonTitle(rowRevisions), textWidth);
                        string revisionBody = TrimBalloonText(BuildRevisionBalloonPreview(rowRevisions), textWidth);
                        candidates.Add(new DocxMarkupBalloonCandidate(
                            DocxMarkupBalloonKind.Revision,
                            revisionTitle,
                            revisionBody,
                            revisionTitle,
                            revisionBody,
                            row.Y + Math.Max(0d, row.Height) * 0.5d,
                            row.Table.TableX + Math.Max(0d, row.Table.ResolvedTableWidth) * 0.5d,
                            row.Table.TableX - 2d,
                            row.Table.TableX + Math.Max(0d, row.Table.ResolvedTableWidth) + 2d,
                            sequence++,
                            rowRevisionPalette.FillRgb,
                            rowRevisionPalette.StrokeRgb,
                            rowRevisionPalette.TitleRgb,
                            new DocxMarkupBalloonRgb(0, 0, 0),
                            CandidateCount: 1, CommentCandidateCount: 0, CommentWithDateCount: 0, CommentResolvedCount: 0, CommentOpenCount: 0, CommentReplyCount: 0, RevisionCandidateCount: 1,
                            BodySummaryPartCount: CountBalloonSummaryPart(revisionBody),
                            WordCompatibleBodySummaryPartCount: CountBalloonSummaryPart(revisionBody)));
                    }
                }

                for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
                {
                    DocxTableCellLayout cell = row.Cells[cellIndex];
                    IReadOnlyList<DocxRevisionInfo> cellRevisions = cell.Cell.Revisions;
                    if (cellRevisions.Count == 0)
                    {
                        continue;
                    }

                    string cellKey = TableCellBalloonKey(row, cellIndex);
                    if (!renderedTableCellRevisions.Add(cellKey))
                    {
                        continue;
                    }

                    DocxRevisionMarkupPalette cellRevisionPalette = ResolveRevisionMarkupPalette(cellRevisions);
                    string revisionTitle = TrimBalloonText(BuildRevisionBalloonTitle(cellRevisions), textWidth);
                    string revisionBody = TrimBalloonText(BuildRevisionBalloonPreview(cellRevisions), textWidth);
                    candidates.Add(new DocxMarkupBalloonCandidate(
                        DocxMarkupBalloonKind.Revision,
                        revisionTitle,
                        revisionBody,
                        revisionTitle,
                        revisionBody,
                        cell.Y + Math.Max(0d, cell.Height) * 0.5d,
                        cell.X + Math.Max(0d, cell.Width) * 0.5d,
                        cell.X - 2d,
                        cell.X + Math.Max(0d, cell.Width) + 2d,
                        sequence++,
                        cellRevisionPalette.FillRgb,
                        cellRevisionPalette.StrokeRgb,
                        cellRevisionPalette.TitleRgb,
                        new DocxMarkupBalloonRgb(0, 0, 0),
                        CandidateCount: 1, CommentCandidateCount: 0, CommentWithDateCount: 0, CommentResolvedCount: 0, CommentOpenCount: 0, CommentReplyCount: 0, RevisionCandidateCount: 1,
                        BodySummaryPartCount: CountBalloonSummaryPart(revisionBody),
                        WordCompatibleBodySummaryPartCount: CountBalloonSummaryPart(revisionBody)));
                }
            }
        }

        return candidates;
    }

    private static DocxTextLineLayout ResolveCommentAnchorLine(
        DocxTextLineLayout fallbackLine,
        IReadOnlyList<DocxTextLineLayout> anchorTextLines,
        DocxParagraph paragraph,
        DocxInlineReference reference)
    {
        DocxCommentRange? range = paragraph.CommentRanges.FirstOrDefault(range =>
            string.Equals(range.Id, reference.Id, StringComparison.Ordinal));
        if (range is not null &&
            TryResolveCommentRangeEndAnchor(fallbackLine, anchorTextLines, paragraph, range, out DocxTextLineLayout anchorLine, out _))
        {
            return anchorLine;
        }

        foreach (DocxTextLineLayout candidateLine in EnumerateCommentAnchorSearchLines(fallbackLine, anchorTextLines, paragraph))
        {
            if (TryResolveSourceOffsetAnchorX(candidateLine, reference.SourceRunIndex, reference.TextOffsetInRun, out _) ||
                TryResolvePreviousSourceRunEndAnchorX(candidateLine, reference.SourceRunIndex, minimumSourceRunIndex: null, out _))
            {
                return candidateLine;
            }
        }

        return fallbackLine;
    }

    private static bool TryResolveCommentRangeEndAnchor(
        DocxTextLineLayout line,
        IReadOnlyList<DocxTextLineLayout> anchorTextLines,
        DocxParagraph paragraph,
        DocxCommentRange range,
        out DocxTextLineLayout anchorLine,
        out double anchorX)
    {
        foreach (DocxTextLineLayout candidateLine in EnumerateCommentAnchorSearchLines(line, anchorTextLines, paragraph))
        {
            if (TryResolveSourceOffsetAnchorX(candidateLine, range.EndSourceRunIndex, range.EndTextOffset, out anchorX))
            {
                anchorLine = candidateLine;
                return true;
            }
        }

        foreach (DocxTextLineLayout candidateLine in EnumerateCommentAnchorSearchLines(line, anchorTextLines, paragraph))
        {
            if (TryResolvePreviousSourceRunEndAnchorX(candidateLine, range.EndSourceRunIndex, range.StartSourceRunIndex, out anchorX))
            {
                anchorLine = candidateLine;
                return true;
            }
        }

        foreach (DocxTextLineLayout candidateLine in EnumerateCommentAnchorSearchLines(line, anchorTextLines, paragraph))
        {
            if (TryResolveSourceRunEndAnchorX(candidateLine, range.StartSourceRunIndex, out anchorX))
            {
                anchorLine = candidateLine;
                return true;
            }
        }

        anchorLine = line;
        anchorX = 0d;
        return false;
    }

    private static double ResolveCommentAnchorX(
        DocxTextLineLayout line,
        IReadOnlyList<DocxTextLineLayout> anchorTextLines,
        DocxParagraph paragraph,
        DocxInlineReference reference,
        DocxMarkupContext markupContext)
    {
        DocxCommentRange? range = paragraph.CommentRanges.FirstOrDefault(range =>
            string.Equals(range.Id, reference.Id, StringComparison.Ordinal));
        if (range is not null)
        {
            if (UsesWordCompatibleAllMarkupTextProfile(markupContext) &&
                TryResolveCommentRangeEndAnchor(line, anchorTextLines, paragraph, range, out _, out double wordCompatibleRangeEndX))
            {
                return wordCompatibleRangeEndX +
                    WordCompatibleAllMarkupTextXOffsetPoints -
                    WordCompatibleAllMarkupConnectorBodyAnchorInsetPoints;
            }

            if (TryResolveSourceOffsetAnchorX(line, range.StartSourceRunIndex, range.StartTextOffset, out double startX))
            {
                return startX;
            }

            if (TryResolveSourceOffsetAnchorX(line, range.EndSourceRunIndex, range.EndTextOffset, out double endX))
            {
                return endX;
            }
        }

        return ResolveInlineReferenceAnchorX(line, reference);
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateCommentAnchorSearchLines(
        DocxTextLineLayout line,
        IReadOnlyList<DocxTextLineLayout> anchorTextLines,
        DocxParagraph paragraph)
    {
        yield return line;
        foreach (DocxTextLineLayout candidateLine in anchorTextLines)
        {
            if (!ReferenceEquals(candidateLine, line) &&
                ReferenceEquals(candidateLine.SourceParagraph, paragraph))
            {
                yield return candidateLine;
            }
        }
    }

    private static double ResolveInlineReferenceAnchorX(DocxTextLineLayout line, DocxInlineReference reference)
    {
        if (reference.SourceRunIndex < 0)
        {
            return line.X + Math.Max(0d, line.Width) * 0.5d;
        }

        if (TryResolveSourceOffsetAnchorX(line, reference.SourceRunIndex, reference.TextOffsetInRun, out double anchorX))
        {
            return anchorX;
        }

        return line.X + Math.Max(0d, line.Width) * 0.5d;
    }

    private static bool TryResolveSourceOffsetAnchorX(
        DocxTextLineLayout line,
        int? sourceRunIndex,
        int? textOffsetInRun,
        out double anchorX)
    {
        anchorX = 0d;
        if (sourceRunIndex is not { } runIndex || runIndex < 0)
        {
            return false;
        }

        int offset = Math.Max(0, textOffsetInRun ?? 0);
        foreach (DocxTextSegmentLayout segment in line.Segments)
        {
            if (segment.SourceTextRunIndex != runIndex)
            {
                continue;
            }

            int segmentStart = Math.Max(0, segment.SourceTextOffsetInRun);
            int segmentEnd = segmentStart + segment.Text.Length;
            if (offset < segmentStart || offset > segmentEnd)
            {
                continue;
            }

            if (segment.Text.Length == 0)
            {
                anchorX = segment.X;
                return true;
            }

            double ratio = Math.Clamp((offset - segmentStart) / (double)segment.Text.Length, 0d, 1d);
            anchorX = segment.X + Math.Max(0d, segment.Width) * ratio;
            return true;
        }

        return false;
    }

    private static bool TryResolveSourceRunStartAnchorX(
        DocxTextLineLayout line,
        int? sourceRunIndex,
        out double anchorX)
    {
        anchorX = 0d;
        if (sourceRunIndex is not { } runIndex)
        {
            return false;
        }

        DocxTextSegmentLayout[] segments = line.Segments
            .Where(segment => segment.SourceTextRunIndex == runIndex)
            .ToArray();
        if (segments.Length == 0)
        {
            return false;
        }

        anchorX = segments.Min(segment => segment.X);
        return true;
    }

    private static bool TryResolvePreviousSourceRunEndAnchorX(
        DocxTextLineLayout line,
        int? sourceRunIndex,
        int? minimumSourceRunIndex,
        out double anchorX)
    {
        anchorX = 0d;
        if (sourceRunIndex is not { } runIndex)
        {
            return false;
        }

        int previousRunIndex = line.Segments
            .Where(segment =>
                segment.SourceTextRunIndex < runIndex &&
                (minimumSourceRunIndex is null || segment.SourceTextRunIndex >= minimumSourceRunIndex.Value))
            .Select(segment => segment.SourceTextRunIndex)
            .DefaultIfEmpty(-1)
            .Max();
        return previousRunIndex >= 0 &&
            TryResolveSourceRunEndAnchorX(line, previousRunIndex, out anchorX);
    }

    private static bool TryResolveSourceRunEndAnchorX(
        DocxTextLineLayout line,
        int? sourceRunIndex,
        out double anchorX)
    {
        anchorX = 0d;
        if (sourceRunIndex is not { } runIndex)
        {
            return false;
        }

        DocxTextSegmentLayout[] segments = line.Segments
            .Where(segment => segment.SourceTextRunIndex == runIndex)
            .ToArray();
        if (segments.Length == 0)
        {
            return false;
        }

        anchorX = segments.Max(segment => segment.X + Math.Max(0d, segment.Width));
        return true;
    }

    private static IEnumerable<DocxTableRowLayout> EnumerateMarkupBalloonTableRows(
        DocxLayoutPage page,
        IReadOnlyList<DocxFloatingDrawingLayout> floatingDrawings)
    {
        foreach (DocxTableRowLayout row in EnumerateTableRows(page.StaticTableRows))
        {
            yield return row;
        }

        foreach (DocxTableRowLayout row in EnumerateTableRows(page.Items))
        {
            yield return row;
        }

        foreach (DocxPlacedRelatedStoryLayout story in page.PlacedRelatedStories)
        {
            foreach (DocxTableRowLayout row in EnumerateTableRows(story.TableRows))
            {
                yield return row;
            }
        }

        foreach (DocxTableRowLayout row in EnumerateFloatingDrawingTextBoxTableRows(floatingDrawings))
        {
            yield return row;
        }

        foreach (DocxPlacedRelatedStoryLayout story in page.PlacedRelatedStories)
        {
            foreach (DocxTableRowLayout row in EnumerateFloatingDrawingTextBoxTableRows(story.FloatingDrawings))
            {
                yield return row;
            }
        }
    }

    private static IEnumerable<DocxTableRowLayout> EnumerateTableRows(IEnumerable<DocxLayoutItem> items)
    {
        foreach (DocxLayoutItem item in items)
        {
            if (item is DocxTableRowLayout row)
            {
                foreach (DocxTableRowLayout nested in EnumerateTableRows(row))
                {
                    yield return nested;
                }
            }
        }
    }

    private static IEnumerable<DocxTableRowLayout> EnumerateTableRows(IEnumerable<DocxTableRowLayout> rows)
    {
        foreach (DocxTableRowLayout row in rows)
        {
            foreach (DocxTableRowLayout nested in EnumerateTableRows(row))
            {
                yield return nested;
            }
        }
    }

    private static IEnumerable<DocxTableRowLayout> EnumerateTableRows(DocxTableRowLayout row)
    {
        yield return row;
        foreach (DocxTableCellLayout cell in row.Cells)
        {
            foreach (DocxTableRowLayout nestedRow in cell.NestedRows)
            {
                foreach (DocxTableRowLayout nested in EnumerateTableRows(nestedRow))
                {
                    yield return nested;
                }
            }
        }
    }

    private static string TableBalloonKey(DocxTableRowLayout row)
    {
        return (row.StoryKind ?? string.Empty) +
            ":" + (row.StoryVariantType ?? string.Empty) +
            ":" + row.Table.SourceBlockIndex.ToString(CultureInfo.InvariantCulture) +
            ":" + row.Table.TableIndex.ToString(CultureInfo.InvariantCulture);
    }

    private static string TableRowBalloonKey(DocxTableRowLayout row)
    {
        return TableBalloonKey(row) +
            ":" + row.RowIndex.ToString(CultureInfo.InvariantCulture);
    }

    private static string TableCellBalloonKey(DocxTableRowLayout row, int cellIndex)
    {
        return TableRowBalloonKey(row) +
            ":" + cellIndex.ToString(CultureInfo.InvariantCulture);
    }

    private static DocxMarkupBalloonArea ResolveMarkupBalloonArea(DocxLayoutPage page)
    {
        const double mediaInset = 2d;
        double leftAvailable = Math.Max(0d, page.MarginLeft - 8d);
        double rightAvailable = Math.Max(0d, page.MarginRight - 8d);
        bool useLeft = leftAvailable > rightAvailable && leftAvailable >= 24d;
        double laneWidth = Math.Max(MinimumMarkupBalloonBodyWidthPoints, useLeft ? leftAvailable : rightAvailable);
        double stemWidth = Math.Min(WordMarkupBalloonStemWidthPoints, Math.Max(0d, laneWidth - MinimumMarkupBalloonBodyWidthPoints));
        double mediaBodyWidth = Math.Max(1d, page.Width - mediaInset * 2d);
        double bodyWidth = Math.Min(mediaBodyWidth, Math.Max(MinimumMarkupBalloonBodyWidthPoints, laneWidth - stemWidth));
        if (useLeft)
        {
            double leftLaneX = Math.Max(mediaInset, page.MarginLeft - laneWidth - 4d);
            double bodyX = ClampMarkupBalloonBodyX(leftLaneX, bodyWidth, page.Width, mediaInset);
            double connectorX = Math.Min(page.Width - mediaInset, bodyX + bodyWidth + stemWidth);
            return new DocxMarkupBalloonArea("Left", bodyX, bodyWidth, connectorX);
        }

        double rightLaneX = Math.Min(page.Width - laneWidth - 2d, page.Width - page.MarginRight + 4d);
        double rightBodyX = ClampMarkupBalloonBodyX(rightLaneX + stemWidth, bodyWidth, page.Width, mediaInset);
        double rightConnectorX = Math.Max(mediaInset, rightBodyX - stemWidth);
        return new DocxMarkupBalloonArea("Right", rightBodyX, bodyWidth, rightConnectorX);
    }

    private static double ClampMarkupBalloonBodyX(double x, double width, double pageWidth, double mediaInset)
    {
        double maxX = Math.Max(mediaInset, pageWidth - mediaInset - width);
        return Math.Min(maxX, Math.Max(mediaInset, x));
    }

    private static void RenderMarkupBalloonPlacement(
        DocxMarkupBalloonPlacement placement,
        PdfGraphicsBuilder graphics,
        DocxRunFontResource labelResource,
        DocxRunFontResource bodyResource,
        DocxMarkupContext markupContext)
    {
        DocxMarkupBalloonRgb fillRgb = ResolveMarkupBalloonBodyFillRgb(placement, markupContext);
        DocxMarkupBalloonRgb strokeRgb = ResolveMarkupBalloonBodyStrokeRgb(placement, markupContext);
        graphics.SetFillRgb(fillRgb.Red, fillRgb.Green, fillRgb.Blue);
        graphics.SetStrokeRgb(strokeRgb.Red, strokeRgb.Green, strokeRgb.Blue);
        graphics.SetLineWidth(ResolveMarkupBalloonBodyStrokeWidth(markupContext));
        graphics.FillStrokeRectangleEvenOdd(placement.X, placement.Y, placement.Width, placement.Height);
        if (!placement.IsOverflowSummary)
        {
            RenderMarkupBalloonConnector(placement, graphics, markupContext);
        }

        if (ShouldRenderWordCompatibleBalloonText(placement, markupContext))
        {
            RenderWordCompatibleCommentThreadSeparators(placement, graphics);
            RenderWordCompatibleBalloonText(placement, graphics, labelResource, bodyResource);
            return;
        }

        DrawBalloonText(graphics, labelResource, placement.Title, placement.X + 3d, placement.Y + placement.Height - 7d, 5.5d, placement.TitleRgb.Red, placement.TitleRgb.Green, placement.TitleRgb.Blue);
        if (!string.IsNullOrWhiteSpace(placement.Body))
        {
            DrawBalloonText(graphics, labelResource, placement.Body, placement.X + 3d, placement.Y + 4d, 5d, placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
        }
    }

    private static DocxMarkupBalloonRgb ResolveMarkupBalloonBodyFillRgb(
        DocxMarkupBalloonPlacement placement,
        DocxMarkupContext markupContext)
    {
        return !placement.IsOverflowSummary && UsesWordCompatibleAllMarkupTextProfile(markupContext)
            ? WordCompatibleAllMarkupReviewFillRgb
            : placement.FillRgb;
    }

    private static DocxMarkupBalloonRgb ResolveMarkupBalloonBodyStrokeRgb(
        DocxMarkupBalloonPlacement placement,
        DocxMarkupContext markupContext)
    {
        return !placement.IsOverflowSummary && UsesWordCompatibleAllMarkupTextProfile(markupContext)
            ? WordCompatibleAllMarkupReviewStrokeRgb
            : placement.StrokeRgb;
    }

    private static double ResolveMarkupBalloonBodyStrokeWidth(DocxMarkupContext markupContext)
    {
        return UsesWordCompatibleAllMarkupTextProfile(markupContext)
            ? WordCompatibleAllMarkupConnectorStrokeWidthPoints
            : 0.5d;
    }

    private static bool ShouldRenderWordCompatibleBalloonText(
        DocxMarkupBalloonPlacement placement,
        DocxMarkupContext markupContext)
    {
        return !placement.IsOverflowSummary &&
            markupContext.Mode == OoxPdfDocxMarkupMode.AllMarkup &&
            markupContext.GeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup &&
            markupContext.ExpandsMarkupMargin &&
            !string.IsNullOrWhiteSpace(placement.WordCompatibleTitle);
    }

    private static void RenderWordCompatibleBalloonText(
        DocxMarkupBalloonPlacement placement,
        PdfGraphicsBuilder graphics,
        DocxRunFontResource labelResource,
        DocxRunFontResource bodyResource)
    {
        string title = placement.WordCompatibleTitle ?? placement.Title;
        string body = placement.WordCompatibleBody ?? string.Empty;
        double fontSize = WordCompatibleAllMarkupBalloonTextFontSizePoints;
        double textX = placement.X + WordCompatibleAllMarkupBalloonTextInsetXPoints;
        double firstBaselineY = ResolveWordCompatibleBalloonFirstBaselineY(placement);
        double lineGap = fontSize * 1.2d;
        const byte titleRgb = 0;

        DrawBalloonText(
            graphics,
            labelResource,
            title,
            textX,
            firstBaselineY,
            fontSize,
            titleRgb,
            titleRgb,
            titleRgb,
            WordCompatibleAllMarkupBalloonTitlePositioningCharacterSpacingPoints);
        if (string.IsNullOrWhiteSpace(body))
        {
            return;
        }

        double titleWidth = labelResource.Embedded.MeasureTextPoints(title, fontSize);
        double bodyFirstLineX = textX + titleWidth + WordCompatibleAllMarkupBalloonBodyFirstLineXOffsetPoints;
        double rightEdge = placement.X + placement.Width - 0.5d;
        double firstLineWidth = Math.Max(0d, rightEdge - bodyFirstLineX);
        double continuationWidth = Math.Max(0d, rightEdge - textX);
        string[] lines = WrapWordCompatibleBalloonBody(body, bodyResource.Embedded, fontSize, firstLineWidth, continuationWidth);
        if (lines.Length == 0)
        {
            return;
        }

        DrawBalloonText(graphics, bodyResource, lines[0], bodyFirstLineX, firstBaselineY, fontSize, placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
        if (lines.Length == 1)
        {
            DrawBalloonText(graphics, bodyResource, " ", bodyFirstLineX + labelResource.Embedded.MeasureTextPoints(lines[0], fontSize), firstBaselineY, fontSize, placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
            return;
        }

        double secondBaselineY = firstBaselineY - lineGap;
        DrawBalloonText(
            graphics,
            bodyResource,
            lines[1],
            textX,
            secondBaselineY,
            fontSize,
            placement.BodyRgb.Red,
            placement.BodyRgb.Green,
            placement.BodyRgb.Blue,
            WordCompatibleAllMarkupBalloonContinuationPositioningCharacterSpacingPoints);
        DrawBalloonText(
            graphics,
            bodyResource,
            " ",
            textX + labelResource.Embedded.MeasureTextPoints(lines[1], fontSize) + WordCompatibleAllMarkupBalloonContinuationTerminalSpaceXOffsetPoints,
            secondBaselineY,
            fontSize,
            placement.BodyRgb.Red,
            placement.BodyRgb.Green,
            placement.BodyRgb.Blue);
    }

    private static double ResolveWordCompatibleBalloonFirstBaselineY(DocxMarkupBalloonPlacement placement)
    {
        return placement.Y + placement.Height - WordCompatibleAllMarkupBalloonFirstBaselineTopInsetPoints;
    }

    private static void RenderWordCompatibleCommentThreadSeparators(
        DocxMarkupBalloonPlacement placement,
        PdfGraphicsBuilder graphics)
    {
        int separatorCount = ResolveCommentThreadSeparatorLineCount(placement.CommentReplyCount);
        if (separatorCount == 0)
        {
            return;
        }

        double leftX = placement.X + WordCompatibleAllMarkupBalloonTextInsetXPoints;
        double rightX = placement.X + placement.Width - WordCompatibleAllMarkupBalloonTextInsetXPoints;
        double lineGap = WordCompatibleAllMarkupCommentThreadReplyHeightPoints;
        double firstSeparatorY = ResolveWordCompatibleBalloonFirstBaselineY(placement) -
            lineGap -
            WordCompatibleAllMarkupCommentThreadSeparatorYOffsetPoints;
        graphics.ClearLineDash();
        graphics.SetLineWidth(WordCompatibleAllMarkupConnectorStrokeWidthPoints);
        graphics.SetStrokeRgb(WordCompatibleAllMarkupReviewStrokeRgb.Red, WordCompatibleAllMarkupReviewStrokeRgb.Green, WordCompatibleAllMarkupReviewStrokeRgb.Blue);
        for (int i = 0; i < separatorCount; i++)
        {
            double separatorY = firstSeparatorY - i * lineGap;
            if (separatorY <= placement.Y + 1d)
            {
                break;
            }

            graphics.StrokeLine(leftX, separatorY, rightX, separatorY);
        }
    }

    private static void RenderMarkupBalloonConnector(
        DocxMarkupBalloonPlacement placement,
        PdfGraphicsBuilder graphics,
        DocxMarkupContext markupContext)
    {
        if (markupContext.Mode == OoxPdfDocxMarkupMode.AllMarkup &&
            markupContext.GeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup &&
            markupContext.ExpandsMarkupMargin)
        {
            double bodyConnectorX = placement.Side == "Left"
                ? placement.X + placement.Width
                : placement.X;
            double bodyConnectorY = placement.Y + placement.Height * 0.775d;
            graphics.SetLineWidth(WordCompatibleAllMarkupConnectorStrokeWidthPoints);
            graphics.SetLineDash(WordCompatibleAllMarkupConnectorStrokeWidthPoints, WordCompatibleAllMarkupConnectorStrokeWidthPoints);
            graphics.StrokeLine(placement.AnchorConnectorX, placement.AnchorY, placement.BalloonConnectorX, placement.AnchorY);
            graphics.StrokeLine(placement.BalloonConnectorX, placement.AnchorY, bodyConnectorX, bodyConnectorY);
            graphics.ClearLineDash();
            return;
        }

        graphics.StrokeLine(placement.AnchorConnectorX, placement.AnchorY, placement.BalloonConnectorX, placement.Y + placement.Height * 0.5d);
    }
}
