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
    private static bool ShouldRenderRevisionBalloon(DocxMarkupContext markupContext, IReadOnlyList<DocxRevisionInfo> revisions)
    {
        if (!markupContext.RendersRevisionBalloons || revisions.Count == 0)
        {
            return false;
        }

        if (markupContext.GeometryMode != OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
        {
            return true;
        }

        // Word-compatible print shows insertions, deletions, and moves inline and reserves
        // margin balloons for property (formatting) revisions and comments (Office A/B:
        // dense-revisions, balloon-lane-bands, and review references carry zero
        // insertion/deletion/move balloons alongside inline revision styling). A property
        // change with no property elements (for example an empty rPrChange) is void: Word
        // shows no balloon for it, so it cannot sustain one either.
        return revisions.Any(static revision => IsPropertyChangeRevision(revision.Kind) && revision.PropertyElementNames.Count != 0);
    }

    private static bool IsPropertyChangeRevision(DocxRevisionKind kind)
    {
        return kind is DocxRevisionKind.RunPropertiesChange
            or DocxRevisionKind.ParagraphPropertiesChange
            or DocxRevisionKind.TablePropertiesChange
            or DocxRevisionKind.TableRowPropertiesChange
            or DocxRevisionKind.TableCellPropertiesChange
            or DocxRevisionKind.SectionPropertiesChange;
    }

    private static DocxRunFontResource? EnsureMarkupBalloonTextResource(
        DocxLayout layout,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        CancellationToken cancellationToken)
    {
        // Balloon titles and bodies are synthetic strings (comment metadata, revision labels),
        // so run-collected font subsets can miss their glyphs and EncodeGlyphHex would silently
        // drop them. Build one post-layout subset of the label typeface covering every balloon
        // string when the label resource falls short. Returns null when no extra coverage is
        // needed, leaving existing subsets (and their snapshots) untouched.
        if (!markupContext.RendersCommentBalloons && !markupContext.RendersRevisionBalloons)
        {
            return null;
        }

        DocxRunFontResource? labelResource = ResolveMarkupLabelFontResource(fontResources);
        if (labelResource is null)
        {
            return null;
        }

        var codepoints = new HashSet<int>();
        for (int pageIndex = 0; pageIndex < layout.Pages.Count; pageIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxRunFontResource bodyCoverageResource = ResolveMarkupBodyFontResource(fontResources) ?? labelResource;
            foreach (DocxMarkupBalloonPlacement placement in BuildMarkupBalloonPlacements(
                layout.Pages[pageIndex],
                layout.RelatedStories,
                EnumeratePageFloatingDrawings(layout, pageIndex).ToArray(),
                markupContext,
                labelResource.Embedded,
                bodyCoverageResource.Embedded))
            {
                foreach (string? text in new string?[] { placement.Title, placement.Body, placement.WordCompatibleTitle, placement.WordCompatibleBody })
                {
                    if (string.IsNullOrEmpty(text))
                    {
                        continue;
                    }

                    foreach (Rune rune in text.EnumerateRunes())
                    {
                        codepoints.Add(rune.Value);
                    }
                }
            }
        }

        if (codepoints.Count == 0)
        {
            return null;
        }

        OpenTypeFont labelFont = labelResource.Embedded.Font;
        bool needsExtraCoverage = false;
        foreach (int codePoint in codepoints)
        {
            cancellationToken.ThrowIfCancellationRequested();
            ushort glyph = labelFont.MapCodePoint(codePoint);
            if (glyph != 0 && !labelResource.Embedded.TryGetEncodedCid(glyph, out _))
            {
                needsExtraCoverage = true;
                break;
            }
        }

        if (!needsExtraCoverage)
        {
            return null;
        }

        if (fontResources.Resources is not List<PdfFontResource> mutableResources)
        {
            return null;
        }

        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(labelFont, codepoints, cancellationToken);
        string name = "F" + (mutableResources.Count + 1).ToString(CultureInfo.InvariantCulture);
        mutableResources.Add(new PdfFontResource(name, embedded));
        return new DocxRunFontResource(name, embedded, labelResource.Resolution);
    }

    private static void RenderMarkupBalloons(
        DocxLayoutPage page,
        IReadOnlyList<DocxRelatedStoryLayout> relatedStories,
        IReadOnlyList<DocxFloatingDrawingLayout> floatingDrawings,
        PdfGraphicsBuilder graphics,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        DocxRunFontResource? balloonTextResource)
    {
        DocxRunFontResource? labelResource = balloonTextResource ?? ResolveMarkupLabelFontResource(fontResources);
        if (labelResource is null)
        {
            return;
        }

        DocxRunFontResource bodyResource = ResolveMarkupBodyFontResource(fontResources) ?? labelResource;
        foreach (DocxMarkupBalloonPlacement placement in BuildMarkupBalloonPlacements(page, relatedStories, floatingDrawings, markupContext, labelResource.Embedded, bodyResource.Embedded))
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
        DocxMarkupContext markupContext,
        PdfEmbeddedFont? labelEmbedded,
        PdfEmbeddedFont? bodyEmbedded)
    {
        if (!markupContext.RendersCommentBalloons && !markupContext.RendersRevisionBalloons)
        {
            return [];
        }

        DocxMarkupBalloonArea area = ResolveMarkupBalloonArea(page, markupContext);
        DocxMarkupBalloonCandidate[] candidates = OrderMarkupBalloonCandidatesForPlacement(
                GroupNearbyMarkupBalloonCandidates(
                    OrderMarkupBalloonCandidatesForPlacement(CollectMarkupBalloonCandidates(area.Width)),
                    area.Width))
            .ToArray();

        IReadOnlyList<DocxMarkupBalloonCandidate> GroupNearbyMarkupBalloonCandidates(
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

            bool CanGroupMarkupBalloonCandidates(IReadOnlyList<DocxMarkupBalloonCandidate> group, DocxMarkupBalloonCandidate candidate)
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
        }
        if (candidates.Length == 0)
        {
            return [];
        }

        IReadOnlyList<DocxMarkupBalloonLaneBand> laneBands = BuildMarkupBalloonLaneBands();
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
                double height = ResolveMarkupBalloonHeight(candidate, markupContext, labelEmbedded, bodyEmbedded, area.Width);
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
                connectorAnchorX = ClampMarkupBalloonConnectorAnchorX(connectorAnchorX);
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

        IReadOnlyList<DocxMarkupBalloonLaneBand> BuildMarkupBalloonLaneBands()
        {
            var bands = new List<DocxMarkupBalloonLaneBand>();
            var current = new List<DocxMarkupBalloonCandidate>();
            double currentTopLimit = 0d;
            double currentBottom = 0d;
            double currentMaxBalloonHeight = 0d;

            foreach (DocxMarkupBalloonCandidate candidate in candidates)
            {
                double anchorY = ResolveMarkupBalloonAnchorY(candidate.AnchorY, page, markupContext);
                double height = ResolveMarkupBalloonHeight(candidate, markupContext, labelEmbedded, bodyEmbedded, area.Width);
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

        double ClampMarkupBalloonConnectorAnchorX(double connectorAnchorX)
        {
            return Math.Min(page.Width - 0.5d, Math.Max(0.5d, connectorAnchorX));
        }

        IReadOnlyList<DocxMarkupBalloonCandidate> CollectMarkupBalloonCandidates(double textWidth)
        {
            var balloonCandidates = new List<DocxMarkupBalloonCandidate>();
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
            DocxTextLineLayout[] anchorTextLines = EnumerateMarkupBalloonAnchorTextLines(page, floatingDrawings, markupContext, page.Height).ToArray();
            // Office A/B (w6-tbxctl, w6-staticfloat, and w6-inline probes, Word-COM
            // rendered, plus uniformity for placed stories): Word balloons body-anchored
            // comments but never textbox ones (body-flow or static floatings, inline
            // boxes, or placed-story floatings - Word rejects anchor-in-footnote files so
            // the placed case extends the probed policy by uniformity), so
            // Word-compatible geometry suppresses comment candidates anchored there
            // (keyed by paragraph reference identity, the same idiom as the
            // rendered-comments key below). Revision candidates keep the legacy path
            // (unprobed).
            HashSet<int> textBoxParagraphs = UsesWordCompatibleAllMarkupTextProfile(markupContext)
                ? floatingDrawings
                    .SelectMany(EnumerateFloatingDrawingTextBoxTextLines)
                    .Select(line => line.SourceParagraph)
                    .Concat(EnumerateInlineTextBoxTextLines(page).Select(line => line.SourceParagraph))
                    .Concat(page.PlacedRelatedStories.SelectMany(story => story.FloatingDrawings.SelectMany(EnumerateFloatingDrawingTextBoxTextLines)).Select(line => line.SourceParagraph))
                    .Concat(EnumerateTableRowTextBoxLines(page).Select(line => line.SourceParagraph))
                    .Concat(EnumerateStaticTextBoxLines(page).Select(line => line.SourceParagraph))
                    .OfType<DocxParagraph>()
                    .Select(RuntimeHelpers.GetHashCode)
                    .ToHashSet()
                : [];
            int sequence = 0;
            foreach (DocxTextLineLayout line in anchorTextLines)
            {
                if (line.SourceParagraph is not { } paragraph)
                {
                    continue;
                }

                if (markupContext.RendersCommentBalloons &&
                    !textBoxParagraphs.Contains(RuntimeHelpers.GetHashCode(paragraph)))
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
                        balloonCandidates.Add(new DocxMarkupBalloonCandidate(
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

                if (ShouldRenderRevisionBalloon(markupContext, paragraph.Revisions))
                {
                    int key = RuntimeHelpers.GetHashCode(paragraph);
                    if (!renderedRevisions.Add(key))
                    {
                        continue;
                    }

                    DocxRevisionMarkupPalette paragraphRevisionPalette = ResolveRevisionMarkupPalette(paragraph.Revisions);
                    string revisionTitle = TrimBalloonText(BuildRevisionBalloonTitle(paragraph.Revisions), textWidth);
                    string revisionBody = TrimBalloonText(BuildRevisionBalloonPreview(paragraph), textWidth);
                    balloonCandidates.Add(new DocxMarkupBalloonCandidate(
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
                    if (ShouldRenderRevisionBalloon(markupContext, tableRevisions))
                    {
                        string tableKey = TableBalloonKey(row);
                        if (renderedTableRevisions.Add(tableKey))
                        {
                            DocxRevisionMarkupPalette tableRevisionPalette = ResolveRevisionMarkupPalette(tableRevisions);
                            string revisionTitle = TrimBalloonText(BuildRevisionBalloonTitle(tableRevisions), textWidth);
                            string revisionBody = TrimBalloonText(BuildRevisionBalloonPreview(tableRevisions), textWidth);
                            balloonCandidates.Add(new DocxMarkupBalloonCandidate(
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
                    if (ShouldRenderRevisionBalloon(markupContext, rowRevisions))
                    {
                        string rowKey = TableRowBalloonKey(row);
                        if (renderedTableRowRevisions.Add(rowKey))
                        {
                            DocxRevisionMarkupPalette rowRevisionPalette = ResolveRevisionMarkupPalette(rowRevisions);
                            string revisionTitle = TrimBalloonText(BuildRevisionBalloonTitle(rowRevisions), textWidth);
                            string revisionBody = TrimBalloonText(BuildRevisionBalloonPreview(rowRevisions), textWidth);
                            balloonCandidates.Add(new DocxMarkupBalloonCandidate(
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
                        if (!ShouldRenderRevisionBalloon(markupContext, cellRevisions))
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
                        balloonCandidates.Add(new DocxMarkupBalloonCandidate(
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

            return balloonCandidates;

            DocxCommentThreadBalloonMetrics CountCommentThreadBalloonMetrics(DocxRelatedStoryLayout? storyLayout, IReadOnlyList<DocxRelatedStoryLayout> replies)
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

        double resolved = anchorY - (markupContext.WordCompatibleTextYOffset + WordCompatibleAllMarkupBalloonAnchorRowInsetPoints);
        return Math.Min(page.Height - page.MarginTop, Math.Max(page.MarginBottom, resolved));
    }

    private static double ResolveMarkupBalloonHeight(
        DocxMarkupBalloonCandidate candidate,
        DocxMarkupContext markupContext,
        PdfEmbeddedFont? labelEmbedded,
        PdfEmbeddedFont? bodyEmbedded,
        double balloonWidth)
    {
        if (markupContext.Mode == OoxPdfDocxMarkupMode.AllMarkup &&
            markupContext.GeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup &&
            markupContext.ExpandsMarkupMargin)
        {
            // Office A/B (tbxrev one-line 12.3 plus dense two-line 21.4 balloon
            // rects, Word-COM rendered): reply-less balloon bodies fit the
            // rendered text rows while titles stay top-anchored, so the height
            // follows the wrapped body rows plus Office insets. Threaded
            // balloons keep the legacy flat height plus separator extra
            // (separator geometry is unprobed for row-derived sizing).
            if (candidate.CommentReplyCount == 0 &&
                !string.IsNullOrWhiteSpace(candidate.WordCompatibleTitle) &&
                labelEmbedded is not null &&
                bodyEmbedded is not null)
            {
                string title = candidate.WordCompatibleTitle ?? candidate.Title;
                string body = candidate.WordCompatibleBody ?? string.Empty;
                double fontSize = 9d * markupContext.WordCompatiblePrintScale;
                double titleWidth = labelEmbedded.MeasureTextPoints(title, fontSize);
                ComputeWordCompatibleBalloonWrapWidths(titleWidth, balloonWidth, out double firstLineWidth, out double continuationWidth);
                int rows = CountWordCompatibleBalloonTextRows(body, bodyEmbedded, fontSize, firstLineWidth, continuationWidth);
                return WordCompatibleAllMarkupBalloonFirstBaselineTopInsetPoints +
                    (rows - 1) * fontSize * 1.2d +
                    WordCompatibleAllMarkupBalloonBottomInsetPoints;
            }

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
        string[] bodyParts = BuildMarkupBalloonGroupBodyParts();
        string body = TrimBalloonText(string.Join("; ", bodyParts), textWidth);
        string[] wordCompatibleBodyParts = allRevisions
            ? bodyParts
            : BuildWordCompatibleMarkupBalloonGroupBodyParts();
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
                // Office: connectors attach at the range end, i.e. the last range line (lowest anchor).
                : (group.MinBy(candidate => candidate.AnchorY)?.AnchorConnectorX ?? group[0].AnchorConnectorX),
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

        string[] BuildMarkupBalloonGroupBodyParts()
        {
            return group
                .Select(BuildMarkupBalloonGroupBodyPart)
                .Where(part => !string.IsNullOrWhiteSpace(part))
                .OfType<string>()
                .Distinct(StringComparer.Ordinal)
                .ToArray();

            string? BuildMarkupBalloonGroupBodyPart(DocxMarkupBalloonCandidate candidate)
            {
                string? partBody = FirstNonEmpty(candidate.Body);
                string? partTitle = FirstNonEmpty(candidate.Title);
                if (partBody is null)
                {
                    return partTitle;
                }

                return partTitle is null ? partBody : partTitle + ": " + partBody;
            }
        }

        string[] BuildWordCompatibleMarkupBalloonGroupBodyParts()
        {
            return group
                .Select(BuildWordCompatibleMarkupBalloonGroupBodyPart)
                .Where(part => !string.IsNullOrWhiteSpace(part))
                .OfType<string>()
                .Distinct(StringComparer.Ordinal)
                .ToArray();

            string? BuildWordCompatibleMarkupBalloonGroupBodyPart(DocxMarkupBalloonCandidate candidate)
            {
                string? partBody = FirstNonEmpty(candidate.WordCompatibleBody, candidate.Body);
                if (partBody is not null)
                {
                    return partBody;
                }

                return candidate.Kind == DocxMarkupBalloonKind.Comment
                    ? null
                    : FirstNonEmpty(candidate.WordCompatibleTitle, candidate.Title);
            }
        }
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

        string BuildOverflowContinuationBody(IReadOnlyList<DocxMarkupBalloonCandidate> candidates)
        {
            int count = candidates.Sum(candidate => candidate.CandidateCount);
            string sample = candidates
                .Select(candidate => candidate.Body)
                .FirstOrDefault(text => !string.IsNullOrWhiteSpace(text)) ?? string.Empty;
            return string.IsNullOrWhiteSpace(sample)
                ? count.ToString(CultureInfo.InvariantCulture)
                : count.ToString(CultureInfo.InvariantCulture) + ": " + sample;
        }
    }

    private static DocxMarkupBalloonArea ResolveMarkupBalloonArea(DocxLayoutPage page, DocxMarkupContext markupContext)
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
        if (markupContext.Mode == OoxPdfDocxMarkupMode.AllMarkup &&
            markupContext.GeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup &&
            markupContext.ExpandsMarkupMargin &&
            Math.Abs(markupContext.WordCompatiblePrintScale - 1d) >= 0.000000001d &&
            !ShouldUseLeftMarkupLane(page))
        {
            // Office (W5-X1): lane geometry scales uniformly; the layout-derived lane ignores
            // the right margin (it saturates at the preferred reserve) and never scales.
            // Authored body end = layout right edge plus the reserved points.
            double designBodyEnd = page.Width - page.MarginRight + page.MarkupMarginReservePoints;
            double printScale = markupContext.WordCompatiblePrintScale;
            bodyWidth = Math.Max(MinimumMarkupBalloonBodyWidthPoints, WordCompatibleAllMarkupBalloonBodyWidthPoints * printScale);
            double scaledTextX = (designBodyEnd + WordCompatibleAllMarkupBalloonLaneGapPoints) * printScale;
            // Word prints balloon bodies to within 1.4pt of the media edge (R72/R207/dense rects),
            // so the scaled lane clamps at 1pt rather than the legacy 2pt guard.
            rightBodyX = ClampMarkupBalloonBodyX(scaledTextX - WordCompatibleAllMarkupBalloonTextInsetXPoints, bodyWidth, page.Width, 1d);
        }
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
            RenderWordCompatibleBalloonText(placement, graphics, labelResource, bodyResource, markupContext.WordCompatiblePrintScale);
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
        DocxRunFontResource bodyResource,
        double wordCompatiblePrintScale)
    {
        string title = placement.WordCompatibleTitle ?? placement.Title;
        string body = placement.WordCompatibleBody ?? string.Empty;
        double fontSize = 9d * wordCompatiblePrintScale;
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
        ComputeWordCompatibleBalloonWrapWidths(titleWidth, placement.Width, out double firstLineWidth, out double continuationWidth);
        double bodyFirstLineX = textX + titleWidth + WordCompatibleAllMarkupBalloonBodyFirstLineXOffsetPoints;
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
