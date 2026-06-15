using System.Globalization;
using System.Text.Encodings.Web;
using System.Text.Json;

using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf;

if (args.Length < 2)
{
    Console.Error.WriteLine("Usage: Lokad.OoxPdf.DocxInspect <input.docx> <output-directory> [--docx-markup final|original|simple|all] [--docx-markup-geometry preserve|reserve-margin|word-compatible]");
    Environment.Exit(2);
}

string inputPath = Path.GetFullPath(args[0]);
string outputDirectory = Path.GetFullPath(args[1]);
OoxPdfDocxMarkupMode markupMode = OoxPdfDocxMarkupMode.Final;
OoxPdfDocxMarkupGeometryMode markupGeometryMode = OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout;
for (int i = 2; i < args.Length; i++)
{
    if (args[i].Equals("--docx-markup", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
    {
        if (!TryParseDocxMarkupMode(args[++i], out markupMode))
        {
            Console.Error.WriteLine($"Invalid DOCX markup mode: {args[i]}");
            Environment.Exit(2);
        }

        continue;
    }

    if (args[i].Equals("--docx-markup-geometry", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
    {
        if (!TryParseDocxMarkupGeometryMode(args[++i], out markupGeometryMode))
        {
            Console.Error.WriteLine($"Invalid DOCX markup geometry mode: {args[i]}");
            Environment.Exit(2);
        }

        continue;
    }

    Console.Error.WriteLine($"Unknown or incomplete argument: {args[i]}");
    Environment.Exit(2);
}

Directory.CreateDirectory(outputDirectory);

using FileStream stream = File.OpenRead(inputPath);
OoxPackage package = OoxPackage.Open(stream);
DocxDocument document = new DocxReader().Read(package, markupMode: markupMode);
var renderer = new DocxRenderer(markupMode: markupMode, markupGeometryMode: markupGeometryMode);

var options = new JsonSerializerOptions
{
    WriteIndented = true,
    Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
};

DocxLayoutSnapshot layout = renderer.InspectLayout(document);
DocxFontPlanSnapshot fontPlan = renderer.InspectFontPlan(document);
DocxStructureSnapshot structure = renderer.InspectStructure(document);
DocxTextEmissionSnapshot textEmission = renderer.InspectTextEmission(document);
DocxMarkupBalloonPlacementSnapshot[] markupBalloons = renderer.InspectMarkupBalloons(document).ToArray();
File.WriteAllText(
    Path.Combine(outputDirectory, "layout-snapshot.json"),
    JsonSerializer.Serialize(layout, options));
File.WriteAllText(
    Path.Combine(outputDirectory, "font-plan-snapshot.json"),
    JsonSerializer.Serialize(fontPlan, options));
File.WriteAllText(
    Path.Combine(outputDirectory, "structure-snapshot.json"),
    JsonSerializer.Serialize(structure, options));
File.WriteAllText(
    Path.Combine(outputDirectory, "style-catalog.json"),
    JsonSerializer.Serialize(document.StyleCatalog, options));
File.WriteAllText(
    Path.Combine(outputDirectory, "style-catalog-summary.json"),
    JsonSerializer.Serialize(new
    {
        document.StyleCatalog.HasRunDefaults,
        document.StyleCatalog.HasParagraphDefaults,
        document.StyleCatalog.DefaultTableStyleId,
        ParagraphStyleCount = document.StyleCatalog.ParagraphStyles.Count,
        CharacterStyleCount = document.StyleCatalog.CharacterStyles.Count,
        TableStyleCount = document.StyleCatalog.TableStyles.Count,
        ParagraphBasedOnCount = document.StyleCatalog.ParagraphStyles.Count(style => style.BasedOnStyleId is not null),
        CharacterBasedOnCount = document.StyleCatalog.CharacterStyles.Count(style => style.BasedOnStyleId is not null),
        TableBasedOnCount = document.StyleCatalog.TableStyles.Count(style => style.BasedOnStyleId is not null),
        TableConditionalRegionCount = document.StyleCatalog.TableStyles.Sum(style => style.ConditionalRegionCount)
    }, options));
File.WriteAllText(
    Path.Combine(outputDirectory, "text-emission-snapshot.json"),
    JsonSerializer.Serialize(textEmission, options));
File.WriteAllText(
    Path.Combine(outputDirectory, "document-settings.json"),
    JsonSerializer.Serialize(document.Settings, options));
File.WriteAllText(
    Path.Combine(outputDirectory, "markup-mode.json"),
    JsonSerializer.Serialize(new { MarkupMode = markupMode.ToString(), MarkupGeometryMode = markupGeometryMode.ToString() }, options));
File.WriteAllText(
    Path.Combine(outputDirectory, "markup-summary.json"),
    JsonSerializer.Serialize(CreateMarkupSummary(markupMode, markupGeometryMode, document.Settings, document.PageSettings, structure, layout, textEmission, markupBalloons), options));
File.WriteAllText(
    Path.Combine(outputDirectory, "markup-balloons.json"),
    JsonSerializer.Serialize(markupBalloons, options));
File.WriteAllText(
    Path.Combine(outputDirectory, "page-summary.json"),
    JsonSerializer.Serialize(layout.Pages.Select((page, index) => new
    {
        Page = index + 1,
        page.Width,
        page.Height,
        page.ItemCount,
        page.TextLineCount,
        page.InlineImageCount,
        page.TableRowCount,
        BodyTableCellTextLineCount = CountChildTextLines(page.Items),
        page.StaticTextLineCount,
        page.StaticInlineImageCount,
        page.StaticTableRowCount,
        StaticTableCellTextLineCount = CountChildTextLines(page.StaticItems),
        StaticStoryCount = page.StaticStories.Count,
        page.SourceBlockCount,
        page.FirstSourceBlockIndex,
        page.LastSourceBlockIndex
    }), options));
File.WriteAllText(
    Path.Combine(outputDirectory, "source-block-summary.json"),
    JsonSerializer.Serialize(layout.SourceBlocks, options));
File.WriteAllText(
    Path.Combine(outputDirectory, "floating-drawing-summary.json"),
    JsonSerializer.Serialize(layout.FloatingDrawings
        .Select(drawing => ToFloatingDrawingSummary("Body", drawing))
        .Concat(layout.StaticFloatingDrawings.Select(drawing => ToFloatingDrawingSummary("Static", drawing))), options));
File.WriteAllText(
    Path.Combine(outputDirectory, "static-story-summary.json"),
    JsonSerializer.Serialize(layout.Pages.SelectMany((page, pageIndex) => page.StaticStories.Select(story => new
    {
        Page = pageIndex + 1,
        story.Kind,
        story.VariantType,
        story.TextLineCount,
        story.ParagraphCount,
        story.InlineImageCount,
        story.TableRowCount,
        story.SourceLineCount,
        story.TextLength,
        story.FirstParagraphLineCount,
        story.VerticalTop,
        story.VerticalBottom,
        story.FirstSourceParagraphIndex,
        story.LastSourceParagraphIndex,
        story.FirstSourceLineIndex,
        story.LastSourceLineIndex,
        ItemCount = story.Items.Count
    })), options));
File.WriteAllText(
    Path.Combine(outputDirectory, "static-story-item-summary.json"),
    JsonSerializer.Serialize(layout.Pages.SelectMany((page, pageIndex) => page.StaticStories.SelectMany(story => story.Items.Select(item => new
    {
        Page = pageIndex + 1,
        story.Kind,
        story.VariantType,
        ItemKind = item.Kind,
        item.SourceBlockIndex,
        item.SourceParagraphIndex,
        item.SourceLineIndex,
        item.X,
        item.Y,
        item.Width,
        item.Height,
        item.TextLength,
        ChildTextLineCount = item.TextLines?.Count ?? 0,
        item.LineHeightPoints,
        item.AppliedBeforeSpacingPoints,
        item.IsFirstParagraphLine,
        item.PendingAfterSpacingPoints,
        item.ParagraphBeforeSpacingPoints,
        item.ParagraphAfterSpacingPoints,
        item.ContextualSpacingSuppressed
    }))), options));
File.WriteAllText(
    Path.Combine(outputDirectory, "table-summary.json"),
    JsonSerializer.Serialize(layout.Tables.Select(table => new
    {
        table.TableIndex,
        table.SourceBlockIndex,
        table.StoryKind,
        table.StoryVariantType,
        PageStart = table.PageStartIndex + 1,
        PageEnd = table.PageEndIndex + 1,
        table.RowCount,
        table.LaidOutRowCount,
        table.HeaderRowLayoutCount,
        table.AuthoredHeaderRowCount,
        table.GridColumnCount,
        table.GridColumnsWidthSum,
        table.HasExplicitGrid,
        table.ResolvedColumnWidths,
        table.ResolvedTableWidth,
        table.X,
        table.PreferredWidthPoints,
        table.PreferredWidthValue,
        table.PreferredWidthType,
        table.IndentPoints,
        table.CellSpacingPoints,
        table.LayoutValue,
        table.DeclaredHeightRowCount,
        table.ExactHeightRowCount,
        table.AtLeastHeightRowCount,
        table.CantSplitRowCount,
        table.FragmentedRowCount,
        table.FragmentedRowLayoutCount,
        table.MaxRowFragmentCount,
        table.HasVerticalMerge,
        table.AuthoredVerticalMergeCellCount,
        table.AuthoredVerticalMergeRestartCellCount,
        table.AuthoredVerticalMergeContinuationCellCount,
        table.LaidOutVerticalMergeContinuationCellCount,
        table.MissingVerticalMergeOwnerCellCount
    }), options));
File.WriteAllText(
    Path.Combine(outputDirectory, "related-story-summary.json"),
    JsonSerializer.Serialize(layout.RelatedStories.Select(story => new
    {
        story.StoryIndex,
        story.Kind,
        story.PartName,
        story.Id,
        story.BlockCount,
        story.ParagraphCount,
        story.TableCount,
        story.TextLineCount,
        story.TableCellTextLineCount,
        story.TableRowCount,
        story.InlineImageCount,
        story.FloatingDrawingCount,
        story.TextLength,
        story.ContentHeight,
        ItemCount = story.Items.Count,
        SourceBlockCount = story.SourceBlocks.Count
    }), options));
File.WriteAllText(
    Path.Combine(outputDirectory, "related-story-source-block-summary.json"),
    JsonSerializer.Serialize(layout.RelatedStories.SelectMany(story => story.SourceBlocks.Select(block => new
    {
        story.StoryIndex,
        story.Kind,
        story.PartName,
        story.Id,
        block.SourceBlockIndex,
        SourceBlockKind = block.Kind,
        block.ItemCount,
        block.TextLineCount,
        block.InlineImageCount,
        block.TableRowCount,
        block.TextLength,
        block.VerticalTop,
        block.VerticalBottom,
        block.ConsumedHeight,
        block.AppliedBeforeSpacingSum
    })), options));
File.WriteAllText(
    Path.Combine(outputDirectory, "related-story-item-summary.json"),
    JsonSerializer.Serialize(layout.RelatedStories.SelectMany(story => story.Items.Select(item => new
    {
        story.StoryIndex,
        story.Kind,
        story.PartName,
        story.Id,
        ItemKind = item.Kind,
        item.ColumnIndex,
        item.SourceBlockIndex,
        item.SourceParagraphIndex,
        item.SourceLineIndex,
        item.X,
        item.Y,
        item.Width,
        item.Height,
        item.TextLength,
        item.CellCount,
        ChildTextLineCount = item.TextLines?.Count ?? 0,
        item.LineHeightPoints,
        item.AppliedBeforeSpacingPoints,
        item.SingleLineHeightPoints,
        item.ListLabelSingleLineHeightPoints,
        item.BodyWindowsLineHeightPoints,
        item.ListLabelWindowsLineHeightPoints,
        item.EffectiveLineSpacingFactor,
        item.LineSpacingFactorFloorApplied,
        item.IsFirstParagraphLine,
        item.PendingAfterSpacingPoints,
        item.ParagraphBeforeSpacingPoints,
        item.ParagraphAfterSpacingPoints,
        item.ContextualSpacingSuppressed
    })), options));
File.WriteAllText(
    Path.Combine(outputDirectory, "text-emission-summary.json"),
    JsonSerializer.Serialize(new
    {
        textEmission.LineCount,
        textEmission.SegmentCount,
        textEmission.TerminalSpaceSegmentCount,
        textEmission.NonzeroPdfCharacterSpacingSegmentCount,
        textEmission.CompensatedCharacterSpacingSegmentCount,
        CharacterProfile = SumCharacterProfiles(textEmission.Lines.SelectMany(line => line.Segments)),
        AdvanceProfile = SumAdvanceProfiles(textEmission.Lines.SelectMany(line => line.Segments)),
        LinesByStory = CreateTextEmissionStoryBuckets(textEmission.Lines),
        LinesByPage = textEmission.Lines
            .GroupBy(line => line.PageIndex)
            .OrderBy(group => group.Key)
            .Select(group => new
            {
                PageIndex = group.Key,
                LineCount = group.Count(),
                StaticLineCount = group.Count(line => line.IsStaticStory),
                BodyLineCount = group.Count(line => !line.IsStaticStory),
                LinesByStory = CreateTextEmissionStoryBuckets(group),
                SegmentCount = group.Sum(line => line.SegmentCount),
                TerminalSpaceSegmentCount = group.Sum(line => line.TerminalSpaceSegmentCount),
                NonzeroPdfCharacterSpacingSegmentCount = group.Sum(line => line.NonzeroPdfCharacterSpacingSegmentCount),
                CharacterProfile = SumCharacterProfiles(group.SelectMany(line => line.Segments)),
                AdvanceProfile = SumAdvanceProfiles(group.SelectMany(line => line.Segments)),
                SourceBlockCount = group
                    .Select(line => line.SourceBlockIndex)
                    .Where(index => index is not null)
                    .Distinct()
                    .Count(),
                FirstSourceBlockIndex = group
                    .Select(line => line.SourceBlockIndex)
                    .Where(index => index is not null)
                    .FirstOrDefault(),
                LastSourceBlockIndex = group
                    .Select(line => line.SourceBlockIndex)
                    .Where(index => index is not null)
                    .LastOrDefault()
            })
            .ToArray()
    }, options));
File.WriteAllText(
    Path.Combine(outputDirectory, "block-sequence.json"),
    JsonSerializer.Serialize(structure.Blocks, options));
File.WriteAllText(
    Path.Combine(outputDirectory, "table-adjacency-summary.json"),
    JsonSerializer.Serialize(structure.TableAdjacency, options));

static object ToFloatingDrawingSummary(string streamKind, DocxFloatingDrawingLayoutSnapshot drawing)
{
    return new
    {
        StreamKind = streamKind,
        Page = drawing.AnchorPageIndex is null ? (int?)null : drawing.AnchorPageIndex.Value + 1,
        drawing.SourceBlockIndex,
        drawing.SourceParagraphIndex,
        drawing.AnchorColumnIndex,
        drawing.StoryKind,
        drawing.StoryVariantType,
        drawing.WrapKind,
        drawing.WrapTextValue,
        drawing.HorizontalRelativeFromValue,
        drawing.HorizontalAlignValue,
        drawing.VerticalRelativeFromValue,
        drawing.VerticalAlignValue,
        drawing.ExtentWidthPoints,
        drawing.ExtentHeightPoints,
        drawing.DistanceTopPoints,
        drawing.DistanceBottomPoints,
        drawing.DistanceLeftPoints,
        drawing.DistanceRightPoints,
        drawing.HorizontalReferenceX,
        drawing.HorizontalReferenceWidth,
        drawing.VerticalReferenceTop,
        drawing.VerticalReferenceBottom,
        drawing.PlacedX,
        drawing.PlacedTop,
        drawing.WrapExclusionX,
        drawing.WrapExclusionTop,
        drawing.WrapExclusionWidth,
        drawing.WrapExclusionHeight,
        HasImage = drawing.ImageRelationshipId is not null,
        drawing.ImageContentType,
        drawing.ImageWidthPoints,
        drawing.ImageHeightPoints
    };
}

static int CountChildTextLines(IEnumerable<DocxLayoutItemSnapshot> items)
{
    return items.Sum(item => item.TextLines?.Count ?? 0);
}

static bool TryParseDocxMarkupMode(string value, out OoxPdfDocxMarkupMode mode)
{
    switch (value.Trim().ToLowerInvariant())
    {
        case "final":
            mode = OoxPdfDocxMarkupMode.Final;
            return true;
        case "original":
            mode = OoxPdfDocxMarkupMode.Original;
            return true;
        case "simple":
        case "simple-markup":
            mode = OoxPdfDocxMarkupMode.SimpleMarkup;
            return true;
        case "all":
        case "all-markup":
            mode = OoxPdfDocxMarkupMode.AllMarkup;
            return true;
        default:
            mode = OoxPdfDocxMarkupMode.Final;
            return false;
    }
}

static bool TryParseDocxMarkupGeometryMode(string value, out OoxPdfDocxMarkupGeometryMode mode)
{
    switch (value.Trim().ToLowerInvariant())
    {
        case "preserve":
        case "preserve-layout":
        case "preserve-document-layout":
            mode = OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout;
            return true;
        case "reserve":
        case "reserve-margin":
        case "markup-margin":
        case "reserve-markup-margin":
            mode = OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin;
            return true;
        case "word":
        case "word-compatible":
        case "word-compatible-all-markup":
        case "office":
        case "office-compatible":
        case "office-compatible-all-markup":
            mode = OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
            return true;
        default:
            mode = OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout;
            return false;
    }
}

static object CreateMarkupSummary(
    OoxPdfDocxMarkupMode markupMode,
    OoxPdfDocxMarkupGeometryMode markupGeometryMode,
    DocxDocumentSettings documentSettings,
    DocxPageSettings pageSettings,
    DocxStructureSnapshot structure,
    DocxLayoutSnapshot layout,
    DocxTextEmissionSnapshot textEmission,
    IReadOnlyList<DocxMarkupBalloonPlacementSnapshot> markupBalloons)
{
    DocxStructureInlineReferenceSnapshot[] commentReferences = structure.InlineReferences
        .Where(reference => reference.Kind == "Comment")
        .ToArray();
    DocxStructureCommentRangeSnapshot[] commentRanges = structure.CommentRanges?.ToArray() ?? [];
    DocxStructureRevisionRangeSnapshot[] revisionRanges = structure.RevisionRanges?.ToArray() ?? [];
    DocxStructureStorySnapshot[] commentStories = structure.Stories
        .Where(story => story.Kind == "Comment")
        .ToArray();
    DocxStructureBlockSnapshot[] revisionBlocks = structure.Blocks
        .Where(block => block.RevisionCount != 0)
        .ToArray();
    DocxStructureStyleUsageSnapshot[] paragraphStyleUsages = structure.StyleUsages
        .Where(usage => usage.Kind == "Paragraph")
        .ToArray();
    DocxLayoutItemSnapshot[] layoutItems = EnumerateLayoutItems(layout).ToArray();
    bool markupIndicatorsEnabled = markupMode is OoxPdfDocxMarkupMode.SimpleMarkup or OoxPdfDocxMarkupMode.AllMarkup;
    int rawCommentBalloonSignalCount = markupMode == OoxPdfDocxMarkupMode.AllMarkup ? commentReferences.Length : 0;
    int rawRevisionBalloonSignalCount = markupMode == OoxPdfDocxMarkupMode.AllMarkup ? layout.RevisionItemCount : 0;
    int renderedBalloonPlacementCount = markupBalloons.Count;
    int renderedCommentBalloonPlacementCount = markupBalloons.Count(placement => placement.CommentCandidateCount > 0);
    int renderedRevisionBalloonPlacementCount = markupBalloons.Count(placement => placement.RevisionCandidateCount > 0);
    int renderedBalloonCandidateCount = markupBalloons.Sum(placement => placement.CandidateCount);
    int renderedCommentBalloonCandidateCount = markupBalloons.Sum(placement => placement.CommentCandidateCount);
    int renderedRevisionBalloonCandidateCount = markupBalloons.Sum(placement => placement.RevisionCandidateCount);
    int expectedRenderedCommentStoryAnchorCount = markupMode == OoxPdfDocxMarkupMode.AllMarkup
        ? structure.ResolvedCommentStoryAnchorCount
        : 0;
    int visibleCommentStoryAnchorWithoutRenderedBalloonCount = Math.Max(
        0,
        expectedRenderedCommentStoryAnchorCount - renderedCommentBalloonCandidateCount);
    int complexFieldsWithoutCachedResult = Math.Max(
        0,
        structure.ComplexFieldReferenceCount -
        structure.CachedResultFieldReferenceCount -
        structure.DynamicComplexWithoutCachedResultFieldReferenceCount);
    int cachedResultFieldsNotRendered = Math.Max(
        0,
        structure.CachedResultFieldReferenceCount -
        structure.RenderedCachedResultFieldReferenceCount -
        structure.DynamicCachedResultNotRenderedFieldReferenceCount);
    return new
    {
        MarkupMode = markupMode.ToString(),
        MarkupGeometryMode = markupGeometryMode.ToString(),
        QualityCounters = new
        {
            Comments = new
            {
                UnresolvedInlineReferenceCount = commentReferences.Count(reference => reference.ResolvedStoryIndex is null),
                UnresolvedRangeCount = commentRanges.Count(range => range.ResolvedStoryIndex is null),
                VisibleStoryAnchorCount = structure.ResolvedCommentStoryAnchorCount,
                HiddenStoryAnchorCount = structure.HiddenCommentStoryAnchorCount,
                OrphanedStoryAnchorCount = structure.OrphanedCommentStoryAnchorCount,
                UnsupportedStoryAnchorCount = structure.UnsupportedCommentStoryAnchorCount,
                ExpectedRenderedStoryAnchorCount = expectedRenderedCommentStoryAnchorCount,
                RenderedStoryAnchorCandidateCount = renderedCommentBalloonCandidateCount,
                VisibleStoryAnchorWithoutRenderedBalloonCount = visibleCommentStoryAnchorWithoutRenderedBalloonCount,
                UnresolvedOrHiddenOrOrphanedStoryCount =
                    commentReferences.Count(reference => reference.ResolvedStoryIndex is null) +
                    structure.HiddenCommentStoryAnchorCount +
                    structure.OrphanedCommentStoryAnchorCount,
                UnresolvedHiddenOrphanedOrUnsupportedStoryCount =
                    commentReferences.Count(reference => reference.ResolvedStoryIndex is null) +
                    structure.HiddenCommentStoryAnchorCount +
                    structure.OrphanedCommentStoryAnchorCount +
                    structure.UnsupportedCommentStoryAnchorCount
            },
            UnsupportedFields = new
            {
                UnsupportedFieldClassCount =
                    complexFieldsWithoutCachedResult +
                    cachedResultFieldsNotRendered,
                ComplexWithoutCachedResultCount = complexFieldsWithoutCachedResult,
                CachedResultNotRenderedCount = cachedResultFieldsNotRendered,
                PlaceholderFieldReferenceCount = structure.PlaceholderFieldReferenceCount,
                DynamicFieldReferenceCount = structure.DynamicFieldReferenceCount,
                DynamicPlaceholderFieldReferenceCount = structure.DynamicPlaceholderFieldReferenceCount,
                DynamicComplexWithoutCachedResultFieldReferenceCount = structure.DynamicComplexWithoutCachedResultFieldReferenceCount,
                DynamicCachedResultNotRenderedFieldReferenceCount = structure.DynamicCachedResultNotRenderedFieldReferenceCount,
                OtherFieldReferenceCount = structure.OtherFieldReferenceCount,
                UnsupportedClassBuckets = NamedCounts(
                    ("complex-without-cached-result", complexFieldsWithoutCachedResult),
                    ("cached-result-not-rendered", cachedResultFieldsNotRendered)),
                FieldClassBuckets = NamedCounts(
                    ("page", structure.PageFieldReferenceCount),
                    ("num-pages", structure.NumPagesFieldReferenceCount),
                    ("other", structure.OtherFieldReferenceCount),
                    ("complex", structure.ComplexFieldReferenceCount),
                    ("cached-result", structure.CachedResultFieldReferenceCount),
                    ("rendered-cached-result", structure.RenderedCachedResultFieldReferenceCount),
                    ("placeholder", structure.PlaceholderFieldReferenceCount),
                    ("dynamic", structure.DynamicFieldReferenceCount),
                    ("dynamic-placeholder", structure.DynamicPlaceholderFieldReferenceCount),
                    ("nested", structure.NestedFieldReferenceCount))
            },
            FloatingDrawingAnchorClasses = new
            {
                BodyDrawingCount = structure.FloatingDrawingCount,
                LaidOutBodyDrawingCount = layout.FloatingDrawings.Count,
                BodyDrawingWithoutLayoutCount = Math.Max(0, structure.FloatingDrawingCount - layout.FloatingDrawings.Count),
                PayloadBuckets = CountBuckets(structure.FloatingDrawings, FloatingDrawingPayloadBucket),
                WrapBuckets = CountBuckets(structure.FloatingDrawings, drawing => drawing.WrapKind),
                HorizontalReferenceBuckets = CountBuckets(structure.FloatingDrawings, drawing => drawing.HorizontalRelativeFromValue),
                VerticalReferenceBuckets = CountBuckets(structure.FloatingDrawings, drawing => drawing.VerticalRelativeFromValue),
                HorizontalPlacementSourceBuckets = CountBuckets(layout.FloatingDrawings, drawing => drawing.HorizontalPlacementSource),
                VerticalPlacementSourceBuckets = CountBuckets(layout.FloatingDrawings, drawing => drawing.VerticalPlacementSource)
            },
            TextEmissionStoryClasses = new
            {
                Buckets = CreateTextEmissionStoryBuckets(textEmission.Lines),
                FloatingTextBoxLineCount = textEmission.Lines.Count(line => line.StoryKind == "TextBox"),
                RelatedStoryLineCount = textEmission.Lines.Count(line => line.ContainerStoryKind is "Footnote" or "Endnote" or "Comment"),
                StaticStoryLineCount = textEmission.Lines.Count(line => line.IsStaticStory)
            },
            NumberingIndentPatterns = new
            {
                PatternBuckets = CountBuckets(structure.ListUsages, NumberingIndentPatternBucket),
                ParagraphWithAnyIndentSignalCount = structure.ListUsages.Sum(usage =>
                    usage.LeftIndentParagraphCount +
                    usage.RightIndentParagraphCount +
                    usage.FirstLineIndentParagraphCount +
                    usage.HangingIndentParagraphCount +
                    usage.NumberingTabParagraphCount +
                    usage.ParagraphIndentOverrideCount +
                    usage.ParagraphNumberingTabStopCount),
                NumberingTabParagraphCount = structure.ListUsages.Sum(usage => usage.NumberingTabParagraphCount),
                ParagraphIndentOverrideCount = structure.ListUsages.Sum(usage => usage.ParagraphIndentOverrideCount)
            },
            StyleSpacingVariants = new
            {
                PatternBuckets = CountBuckets(paragraphStyleUsages, StyleSpacingPatternBucket),
                ParagraphWithAnySpacingSignalCount = paragraphStyleUsages.Sum(usage =>
                    usage.BeforeSpacingTokenParagraphCount +
                    usage.AfterSpacingTokenParagraphCount +
                    usage.BeforeAutoSpacingParagraphCount +
                    usage.AfterAutoSpacingParagraphCount +
                    usage.BeforeLinesSpacingParagraphCount +
                    usage.AfterLinesSpacingParagraphCount +
                    usage.ContextualSpacingParagraphCount +
                    usage.ExactLineSpacingParagraphCount +
                    usage.AtLeastLineSpacingParagraphCount +
                    usage.AutoLineSpacingParagraphCount),
                LayoutContextualSpacingSuppressedCount = layoutItems.Count(item => item.ContextualSpacingSuppressed == true) +
                    layout.RelatedStories.Sum(story => story.Items.Count(item => item.ContextualSpacingSuppressed == true))
            },
            TableBorderStyles = new
            {
                StyleBuckets = NamedCounts(
                    ("single", structure.Tables.Sum(table => table.SingleBorderCount)),
                    ("thick", structure.Tables.Sum(table => table.ThickBorderCount)),
                    ("double", structure.Tables.Sum(table => table.DoubleBorderCount)),
                    ("dotted", structure.Tables.Sum(table => table.DottedBorderCount)),
                    ("dashed", structure.Tables.Sum(table => table.DashedBorderCount)),
                    ("suppressed", structure.Tables.Sum(table => table.SuppressedBorderCount)),
                    ("other", structure.Tables.Sum(table => table.OtherBorderStyleCount))),
                TableWithMixedSupportedBorderStylesCount = structure.Tables.Count(table =>
                    (table.SingleBorderCount > 0 ? 1 : 0) +
                    (table.ThickBorderCount > 0 ? 1 : 0) +
                    (table.DoubleBorderCount > 0 ? 1 : 0) +
                    (table.DottedBorderCount > 0 ? 1 : 0) +
                    (table.DashedBorderCount > 0 ? 1 : 0) > 1),
                TableWithOtherBorderStyleCount = structure.Tables.Count(table => table.OtherBorderStyleCount > 0)
            },
            BalloonOverflow = new
            {
                OverflowPlacementCount = markupBalloons.Count(placement => placement.IsOverflowSummary),
                OverflowCandidateCount = markupBalloons.Where(placement => placement.IsOverflowSummary).Sum(placement => placement.CandidateCount),
                PageBuckets = CountBuckets(
                    markupBalloons.Where(placement => placement.IsOverflowSummary),
                    placement => (placement.PageIndex + 1).ToString(CultureInfo.InvariantCulture))
            },
            PageLocalMarkupDensity = layout.Pages
                .Select((page, index) => CreatePageMarkupDensity(page, index, markupBalloons))
                .ToArray()
        },
        Geometry = new
        {
            layout.MarkupGeometryMode,
            layout.MarkupMarginReservePoints,
            documentSettings.MirrorMargins,
            documentSettings.MirrorMarginsValue,
            pageSettings.GutterDistancePoints,
            pageSettings.GutterDistanceValue,
            ReserveMarginPageCount = layout.Pages.Count(page => page.MarkupMarginReservePoints > 0.001d),
            MaxRightMarginPoints = layout.Pages.Select(page => page.MarginRight).DefaultIfEmpty(0d).Max(),
            MinColumnFrameWidthPoints = layout.Pages
                .SelectMany(page => page.ColumnFrames)
                .Select(frame => frame.Width)
                .DefaultIfEmpty(0d)
                .Min(),
            PageFrames = layout.Pages
                .Select((page, index) => CreatePageGeometryFrame(page, index))
                .ToArray()
        },
        Revisions = new
        {
            structure.RevisionCount,
            structure.InsertionRevisionCount,
            structure.DeletionRevisionCount,
            structure.MoveFromRevisionCount,
            structure.MoveToRevisionCount,
            structure.OtherRevisionCount,
            RangeCount = revisionRanges.Length,
            ClosedRangeCount = revisionRanges.Count(range => range.IsClosed),
            LinkedCrossBlockRangeCount = revisionRanges.Count(range => range.IsLinkedAcrossBlocks),
            MoveFromRangeCount = revisionRanges.Count(range => range.Kind == "MoveFrom"),
            MoveToRangeCount = revisionRanges.Count(range => range.Kind == "MoveTo"),
            LayoutRevisionItemCount = layout.RevisionItemCount,
            LayoutRevisionCount = layout.RevisionCount,
            TextEmissionRevisionSegmentCount = textEmission.RevisionSegmentCount,
            structure.FormattingRevisionCount,
            structure.RunFormattingRevisionCount,
            structure.ParagraphFormattingRevisionCount,
            structure.TableFormattingRevisionCount,
            structure.RowFormattingRevisionCount,
            structure.CellFormattingRevisionCount,
            structure.SectionFormattingRevisionCount,
            FormattingRevisionPropertyCounts = (structure.FormattingRevisionProperties ?? [])
                .Select(property => new
                {
                    property.Family,
                    property.SourceElement,
                    property.PropertyElementName,
                    property.Count
                })
                .ToArray()
        },
        Comments = new
        {
            structure.CommentReferenceCount,
            InlineReferenceCount = commentReferences.Length,
            ResolvedInlineReferenceCount = commentReferences.Count(reference => reference.ResolvedStoryIndex is not null),
            InlineReferenceWithRevisionCount = commentReferences.Count(reference => reference.RevisionCount != 0),
            RangeCount = commentRanges.Length,
            ResolvedRangeCount = commentRanges.Count(range => range.ResolvedStoryIndex is not null),
            RangeWithStartCount = commentRanges.Count(range => range.StartSourceRunIndex is not null),
            RangeWithEndCount = commentRanges.Count(range => range.EndSourceRunIndex is not null),
            RangeWithReferenceCount = commentRanges.Count(range => range.ReferenceSourceRunIndex is not null),
            RelatedStoryCount = commentStories.Length,
            RelatedStoryWithAuthorCount = commentStories.Count(story => story.HasCommentAuthor),
            RelatedStoryWithInitialsCount = commentStories.Count(story => story.HasCommentInitials),
            RelatedStoryWithDateCount = commentStories.Count(story => story.HasCommentDate),
            ThreadedRelatedStoryCount = commentStories.Count(story => story.CommentParagraphId is not null),
            ReplyRelatedStoryCount = commentStories.Count(story => story.CommentParentId is not null || story.CommentParentParagraphId is not null),
            ResolvedRelatedStoryCount = commentStories.Count(story => story.CommentResolved == true),
            structure.PackageCommentAnchorIdCount,
            structure.HiddenCommentAnchorIdCount,
            structure.ResolvedCommentStoryAnchorCount,
            structure.HiddenCommentStoryAnchorCount,
            structure.OrphanedCommentStoryAnchorCount,
            structure.UnsupportedCommentStoryAnchorCount,
            ExpectedRenderedCommentStoryAnchorCount = expectedRenderedCommentStoryAnchorCount,
            RenderedCommentStoryAnchorCandidateCount = renderedCommentBalloonCandidateCount,
            VisibleCommentStoryAnchorWithoutRenderedBalloonCount = visibleCommentStoryAnchorWithoutRenderedBalloonCount,
            LayoutCommentReferenceItemCount = layout.CommentReferenceItemCount,
            LayoutCommentReferenceCount = layout.CommentReferenceCount,
            TextEmissionCommentReferenceLineCount = textEmission.CommentReferenceLineCount,
            TextEmissionCommentReferenceCount = textEmission.CommentReferenceCount
        },
        Fields = new
        {
            structure.FieldReferenceCount,
            structure.PageFieldReferenceCount,
            structure.NumPagesFieldReferenceCount,
            structure.OtherFieldReferenceCount,
            structure.ComplexFieldReferenceCount,
            structure.CachedResultFieldReferenceCount,
            structure.RenderedCachedResultFieldReferenceCount,
            structure.PlaceholderFieldReferenceCount,
            structure.DynamicFieldReferenceCount,
            structure.DynamicPlaceholderFieldReferenceCount,
            structure.DynamicComplexWithoutCachedResultFieldReferenceCount,
            structure.DynamicCachedResultNotRenderedFieldReferenceCount,
            structure.NestedFieldReferenceCount
        },
        FloatingDrawings = new
        {
            structure.FloatingDrawingCount,
            TextBoxDrawingCount = structure.FloatingDrawings.Count(drawing => drawing.TextBoxBlockCount != 0),
            TextBoxBlockCount = structure.FloatingDrawings.Sum(drawing => drawing.TextBoxBlockCount),
            TextBoxParagraphCount = structure.FloatingDrawings.Sum(drawing => drawing.TextBoxParagraphCount),
            TextBoxTextLength = structure.FloatingDrawings.Sum(drawing => drawing.TextBoxTextLength),
            LayoutFloatingDrawingCount = layout.FloatingDrawings.Count,
            LayoutTextBoxDrawingCount = layout.FloatingDrawings.Count(drawing => drawing.TextBoxTextLineCount != 0 || drawing.TextBoxInlineImageCount != 0 || drawing.TextBoxTableRowCount != 0),
            LayoutTextBoxTextLineCount = layout.FloatingDrawings.Sum(drawing => drawing.TextBoxTextLineCount),
            LayoutTextBoxInlineImageCount = layout.FloatingDrawings.Sum(drawing => drawing.TextBoxInlineImageCount),
            LayoutTextBoxTableRowCount = layout.FloatingDrawings.Sum(drawing => drawing.TextBoxTableRowCount)
        },
        Numbering = new
        {
            ListUsageCount = structure.ListUsages.Count,
            ListParagraphCount = structure.ListUsages.Sum(usage => usage.ParagraphCount),
            LeftIndentParagraphCount = structure.ListUsages.Sum(usage => usage.LeftIndentParagraphCount),
            RightIndentParagraphCount = structure.ListUsages.Sum(usage => usage.RightIndentParagraphCount),
            FirstLineIndentParagraphCount = structure.ListUsages.Sum(usage => usage.FirstLineIndentParagraphCount),
            HangingIndentParagraphCount = structure.ListUsages.Sum(usage => usage.HangingIndentParagraphCount),
            NumberingTabParagraphCount = structure.ListUsages.Sum(usage => usage.NumberingTabParagraphCount),
            ParagraphIndentOverrideCount = structure.ListUsages.Sum(usage => usage.ParagraphIndentOverrideCount),
            ParagraphNumberingTabStopCount = structure.ListUsages.Sum(usage => usage.ParagraphNumberingTabStopCount)
        },
        StyleSpacing = new
        {
            ParagraphStyleUsageCount = structure.StyleUsages.Count(usage => usage.Kind == "Paragraph"),
            StyledParagraphCount = structure.StyleUsages.Where(usage => usage.Kind == "Paragraph").Sum(usage => usage.ParagraphCount),
            BeforeSpacingTokenParagraphCount = structure.StyleUsages.Sum(usage => usage.BeforeSpacingTokenParagraphCount),
            AfterSpacingTokenParagraphCount = structure.StyleUsages.Sum(usage => usage.AfterSpacingTokenParagraphCount),
            BeforeAutoSpacingParagraphCount = structure.StyleUsages.Sum(usage => usage.BeforeAutoSpacingParagraphCount),
            AfterAutoSpacingParagraphCount = structure.StyleUsages.Sum(usage => usage.AfterAutoSpacingParagraphCount),
            BeforeLinesSpacingParagraphCount = structure.StyleUsages.Sum(usage => usage.BeforeLinesSpacingParagraphCount),
            AfterLinesSpacingParagraphCount = structure.StyleUsages.Sum(usage => usage.AfterLinesSpacingParagraphCount),
            ContextualSpacingParagraphCount = structure.StyleUsages.Sum(usage => usage.ContextualSpacingParagraphCount),
            ExactLineSpacingParagraphCount = structure.StyleUsages.Sum(usage => usage.ExactLineSpacingParagraphCount),
            AtLeastLineSpacingParagraphCount = structure.StyleUsages.Sum(usage => usage.AtLeastLineSpacingParagraphCount),
            AutoLineSpacingParagraphCount = structure.StyleUsages.Sum(usage => usage.AutoLineSpacingParagraphCount),
            TableStyleParagraphPropertiesCount = structure.StyleUsages.Sum(usage => usage.TableStyleParagraphPropertiesCount),
            LayoutContextualSpacingSuppressedCount = layoutItems.Count(item => item.ContextualSpacingSuppressed == true) +
                layout.RelatedStories.Sum(story => story.Items.Count(item => item.ContextualSpacingSuppressed == true))
        },
        TableBorders = new
        {
            TableCount = structure.Tables.Count,
            VisibleBorderCount = structure.Tables.Sum(table => table.VisibleBorderCount),
            SingleBorderCount = structure.Tables.Sum(table => table.SingleBorderCount),
            ThickBorderCount = structure.Tables.Sum(table => table.ThickBorderCount),
            DoubleBorderCount = structure.Tables.Sum(table => table.DoubleBorderCount),
            DottedBorderCount = structure.Tables.Sum(table => table.DottedBorderCount),
            DashedBorderCount = structure.Tables.Sum(table => table.DashedBorderCount),
            SuppressedBorderCount = structure.Tables.Sum(table => table.SuppressedBorderCount),
            OtherBorderStyleCount = structure.Tables.Sum(table => table.OtherBorderStyleCount)
        },
        ChangeBars = new
        {
            Enabled = markupIndicatorsEnabled,
            CandidateBlockCount = revisionBlocks.Length,
            CandidateParagraphBlockCount = revisionBlocks.Count(block => block.Kind == "Paragraph"),
            CandidateTableBlockCount = revisionBlocks.Count(block => block.Kind == "Table"),
            LayoutRevisionItemCount = layout.RevisionItemCount
        },
        CommentMarkers = new
        {
            Enabled = markupIndicatorsEnabled,
            CandidateInlineReferenceCount = commentReferences.Length,
            ResolvedCandidateInlineReferenceCount = commentReferences.Count(reference => reference.ResolvedStoryIndex is not null),
            LayoutCommentReferenceItemCount = layout.CommentReferenceItemCount
        },
        Balloons = new
        {
            Implemented = markupMode == OoxPdfDocxMarkupMode.AllMarkup,
            EstimatedPlacementCount = rawCommentBalloonSignalCount + rawRevisionBalloonSignalCount,
            EstimatedCommentPlacementCount = rawCommentBalloonSignalCount,
            EstimatedRevisionPlacementCount = rawRevisionBalloonSignalCount,
            RawMarkupSignalCount = rawCommentBalloonSignalCount + rawRevisionBalloonSignalCount,
            RawCommentReferenceSignalCount = rawCommentBalloonSignalCount,
            RawRevisionLayoutSignalCount = rawRevisionBalloonSignalCount,
            RenderedPlacementCount = renderedBalloonPlacementCount,
            RenderedCommentPlacementCount = renderedCommentBalloonPlacementCount,
            RenderedRevisionPlacementCount = renderedRevisionBalloonPlacementCount,
            PlacementCount = renderedBalloonPlacementCount,
            CommentPlacementCount = renderedCommentBalloonPlacementCount,
            RevisionPlacementCount = renderedRevisionBalloonPlacementCount,
            RenderedCandidateCount = renderedBalloonCandidateCount,
            RenderedCommentCandidateCount = renderedCommentBalloonCandidateCount,
            RenderedRevisionCandidateCount = renderedRevisionBalloonCandidateCount,
            CandidateCount = renderedBalloonCandidateCount,
            CommentCandidateCount = renderedCommentBalloonCandidateCount,
            RevisionCandidateCount = renderedRevisionBalloonCandidateCount,
            MixedPlacementCount = markupBalloons.Count(placement => placement.CommentCandidateCount > 0 && placement.RevisionCandidateCount > 0),
            SideBuckets = CountBuckets(markupBalloons, placement => placement.Side),
            OverflowPlacementCount = markupBalloons.Count(placement => placement.IsOverflowSummary),
            OverflowCandidateCount = markupBalloons.Where(placement => placement.IsOverflowSummary).Sum(placement => placement.CandidateCount),
            GroupedRevisionPlacementCount = markupBalloons.Count(placement => placement.RevisionCandidateCount > 1),
            MaxGroupedCandidateCount = markupBalloons.Select(placement => placement.CandidateCount).DefaultIfEmpty(0).Max(),
            ConnectorPlacementCount = markupBalloons.Count(placement => !placement.IsOverflowSummary),
            ClampedConnectorPlacementCount = markupBalloons.Count(placement => !placement.IsOverflowSummary && placement.AnchorConnectorClamped),
            ConnectorOffsetPlacementCount = markupBalloons.Count(placement => !placement.IsOverflowSummary && Math.Abs(placement.AnchorConnectorX - placement.BalloonConnectorX) > 0.001d),
            SideConsistentConnectorPlacementCount = markupBalloons.Count(placement => !placement.IsOverflowSummary && IsConnectorSideConsistent(placement)),
            SideInconsistentConnectorPlacementCount = markupBalloons.Count(placement => !placement.IsOverflowSummary && !IsConnectorSideConsistent(placement))
        }
    };
}

static IEnumerable<DocxLayoutItemSnapshot> EnumerateLayoutItems(DocxLayoutSnapshot layout)
{
    foreach (DocxLayoutPageSnapshot page in layout.Pages)
    {
        foreach (DocxLayoutItemSnapshot item in page.StaticItems)
        {
            yield return item;
        }

        foreach (DocxLayoutItemSnapshot item in page.PlacedRelatedItems)
        {
            yield return item;
        }

        foreach (DocxLayoutItemSnapshot item in page.Items)
        {
            yield return item;
        }
    }

    foreach (DocxRelatedStoryLayoutSnapshot story in layout.RelatedStories)
    {
        foreach (DocxLayoutItemSnapshot item in story.Items)
        {
            yield return item;
        }
    }
}

static object[] CountBuckets<T>(IEnumerable<T> items, Func<T, string?> selector)
{
    return items
        .Select(item => NormalizeBucketValue(selector(item)))
        .GroupBy(value => value, StringComparer.Ordinal)
        .OrderByDescending(group => group.Count())
        .ThenBy(group => group.Key, StringComparer.Ordinal)
        .Select(group => new { Value = group.Key, Count = group.Count() })
        .Cast<object>()
        .ToArray();
}

static object[] CreateTextEmissionStoryBuckets(IEnumerable<DocxTextEmissionLineSnapshot> lines)
{
    return lines
        .GroupBy(line => new
        {
            line.StoryKind,
            line.StoryVariantType,
            line.ContainerStoryKind,
            line.ContainerStoryVariantType
        })
        .OrderByDescending(group => group.Count())
        .ThenBy(group => group.Key.ContainerStoryKind ?? string.Empty, StringComparer.Ordinal)
        .ThenBy(group => group.Key.StoryKind ?? string.Empty, StringComparer.Ordinal)
        .Select(group => new
        {
            group.Key.StoryKind,
            group.Key.StoryVariantType,
            group.Key.ContainerStoryKind,
            group.Key.ContainerStoryVariantType,
            LineCount = group.Count(),
            SegmentCount = group.Sum(line => line.SegmentCount),
            RevisionSegmentCount = group.Sum(line => line.RevisionSegmentCount),
            CommentReferenceCount = group.Sum(line => line.CommentReferenceCount),
            CharacterProfile = SumCharacterProfiles(group.SelectMany(line => line.Segments)),
            AdvanceProfile = SumAdvanceProfiles(group.SelectMany(line => line.Segments))
        })
        .Cast<object>()
        .ToArray();
}

static object[] NamedCounts(params (string Value, int Count)[] counts)
{
    return counts
        .Select(count => new { count.Value, count.Count })
        .Cast<object>()
        .ToArray();
}

static string FloatingDrawingPayloadBucket(DocxStructureFloatingDrawingSnapshot drawing)
{
    bool hasImage = !string.IsNullOrWhiteSpace(drawing.ImageRelationshipId) ||
        !string.IsNullOrWhiteSpace(drawing.ImagePartName) ||
        !string.IsNullOrWhiteSpace(drawing.ImageContentType);
    bool hasTextBox = drawing.TextBoxBlockCount != 0 ||
        drawing.TextBoxParagraphCount != 0 ||
        drawing.TextBoxTextLength != 0;
    if (hasImage && hasTextBox)
    {
        return "image-and-textbox";
    }

    if (hasImage)
    {
        return "image";
    }

    if (hasTextBox)
    {
        return "textbox";
    }

    return "empty-or-unsupported";
}

static string NumberingIndentPatternBucket(DocxStructureListUsageSnapshot usage)
{
    return string.Join(
        ";",
        $"format={NormalizeBucketValue(usage.FormatValue)}",
        $"suffix={NormalizeBucketValue(usage.SuffixValue)}",
        FlagBucket("left", usage.LeftIndentParagraphCount),
        FlagBucket("right", usage.RightIndentParagraphCount),
        FlagBucket("first", usage.FirstLineIndentParagraphCount),
        FlagBucket("hanging", usage.HangingIndentParagraphCount),
        FlagBucket("numbering-tab", usage.NumberingTabParagraphCount),
        FlagBucket("paragraph-indent", usage.ParagraphIndentOverrideCount),
        FlagBucket("paragraph-tab", usage.ParagraphNumberingTabStopCount));
}

static string StyleSpacingPatternBucket(DocxStructureStyleUsageSnapshot usage)
{
    return string.Join(
        ";",
        FlagBucket("before-token", usage.BeforeSpacingTokenParagraphCount),
        FlagBucket("after-token", usage.AfterSpacingTokenParagraphCount),
        FlagBucket("before-auto", usage.BeforeAutoSpacingParagraphCount),
        FlagBucket("after-auto", usage.AfterAutoSpacingParagraphCount),
        FlagBucket("before-lines", usage.BeforeLinesSpacingParagraphCount),
        FlagBucket("after-lines", usage.AfterLinesSpacingParagraphCount),
        FlagBucket("contextual", usage.ContextualSpacingParagraphCount),
        FlagBucket("exact-line", usage.ExactLineSpacingParagraphCount),
        FlagBucket("at-least-line", usage.AtLeastLineSpacingParagraphCount),
        FlagBucket("auto-line", usage.AutoLineSpacingParagraphCount),
        FlagBucket("table-style-ppr", usage.TableStyleParagraphPropertiesCount));
}

static object CreatePageGeometryFrame(DocxLayoutPageSnapshot page, int pageIndex)
{
    return new
    {
        Page = pageIndex + 1,
        Width = RoundPoint(page.Width),
        Height = RoundPoint(page.Height),
        MarginLeft = RoundPoint(page.MarginLeft),
        MarginRight = RoundPoint(page.MarginRight),
        MarginTop = RoundPoint(page.MarginTop),
        MarginBottom = RoundPoint(page.MarginBottom),
        MarkupMarginReservePoints = RoundPoint(page.MarkupMarginReservePoints),
        page.ColumnFrameCount,
        ColumnFrameWidthSum = RoundPoint(page.ColumnFrameWidthSum),
        ColumnGutterWidthSum = RoundPoint(page.ColumnGutterWidthSum),
        BodyFrame = CreatePageBodyFrame(page),
        MarkupLane = CreatePageMarkupLane(page)
    };
}

static object? CreatePageBodyFrame(DocxLayoutPageSnapshot page)
{
    double minX;
    double maxX;
    if (page.ColumnFrames.Count == 0)
    {
        minX = page.MarginLeft;
        maxX = page.Width - page.MarginRight;
    }
    else
    {
        minX = page.ColumnFrames.Min(frame => frame.X);
        maxX = page.ColumnFrames.Max(frame => frame.X + frame.Width);
    }

    double minY = page.MarginBottom;
    double maxY = page.Height - page.MarginTop;
    if (maxX <= minX || maxY <= minY)
    {
        return null;
    }

    return new
    {
        X = RoundPoint(minX),
        Y = RoundPoint(minY),
        Width = RoundPoint(maxX - minX),
        Height = RoundPoint(maxY - minY)
    };
}

static object? CreatePageMarkupLane(DocxLayoutPageSnapshot page)
{
    double minY = page.MarginBottom;
    double maxY = page.Height - page.MarginTop;
    if (maxY <= minY)
    {
        return null;
    }

    double leftAvailable = Math.Max(0d, page.MarginLeft - 8d);
    double rightAvailable = Math.Max(0d, page.MarginRight - 8d);
    bool useLeft = leftAvailable > rightAvailable && leftAvailable >= 24d;
    double laneWidth = Math.Max(24d, useLeft ? leftAvailable : rightAvailable);
    if (laneWidth <= 0d)
    {
        return null;
    }

    double laneX = useLeft
        ? Math.Max(2d, page.MarginLeft - laneWidth - 4d)
        : Math.Min(page.Width - laneWidth - 2d, page.Width - page.MarginRight + 4d);

    return new
    {
        Side = useLeft ? "Left" : "Right",
        X = RoundPoint(laneX),
        Y = RoundPoint(minY),
        Width = RoundPoint(laneWidth),
        Height = RoundPoint(maxY - minY),
        Reserved = page.MarkupMarginReservePoints > 0.001d
    };
}

static object CreatePageMarkupDensity(
    DocxLayoutPageSnapshot page,
    int pageIndex,
    IReadOnlyList<DocxMarkupBalloonPlacementSnapshot> markupBalloons)
{
    DocxMarkupBalloonPlacementSnapshot[] pageBalloons = markupBalloons
        .Where(placement => placement.PageIndex == pageIndex)
        .ToArray();
    int rawMarkupSignalCount = page.RevisionItemCount + page.CommentReferenceItemCount;
    int renderedBalloonCandidateCount = pageBalloons.Sum(placement => placement.CandidateCount);
    int renderedItemCount =
        page.ItemCount +
        page.StaticTextLineCount +
        page.StaticInlineImageCount +
        page.StaticTableRowCount +
        page.PlacedRelatedStoryTextLineCount +
        page.PlacedRelatedStoryInlineImageCount +
        page.PlacedRelatedStoryTableRowCount;
    int markupSignalCount = rawMarkupSignalCount + pageBalloons.Length;
    double usableHeightInches = Math.Max(0d, page.Height - page.MarginTop - page.MarginBottom) / 72d;
    return new
    {
        Page = pageIndex + 1,
        page.SourceBlockCount,
        RenderedItemCount = renderedItemCount,
        page.RevisionItemCount,
        page.CommentReferenceItemCount,
        RawMarkupSignalCount = rawMarkupSignalCount,
        BalloonPlacementCount = pageBalloons.Length,
        MarkupSide = ResolveMarkupSide(pageBalloons),
        ConnectorPlacementCount = pageBalloons.Count(placement => !placement.IsOverflowSummary),
        ClampedConnectorPlacementCount = pageBalloons.Count(placement => !placement.IsOverflowSummary && placement.AnchorConnectorClamped),
        SideInconsistentConnectorPlacementCount = pageBalloons.Count(placement => !placement.IsOverflowSummary && !IsConnectorSideConsistent(placement)),
        RenderedBalloonCandidateCount = renderedBalloonCandidateCount,
        RenderedCommentBalloonCandidateCount = pageBalloons.Sum(placement => placement.CommentCandidateCount),
        RenderedRevisionBalloonCandidateCount = pageBalloons.Sum(placement => placement.RevisionCandidateCount),
        OverflowBalloonPlacementCount = pageBalloons.Count(placement => placement.IsOverflowSummary),
        OverflowCandidateCount = pageBalloons.Where(placement => placement.IsOverflowSummary).Sum(placement => placement.CandidateCount),
        MarkupSignalCount = markupSignalCount,
        MarkupSignalsPerSourceBlock = Ratio(markupSignalCount, page.SourceBlockCount),
        MarkupSignalsPerRenderedItem = Ratio(markupSignalCount, renderedItemCount),
        MarkupSignalsPerUsableVerticalInch = Ratio(markupSignalCount, usableHeightInches)
    };
}

static bool IsConnectorSideConsistent(DocxMarkupBalloonPlacementSnapshot placement)
{
    if (placement.IsOverflowSummary)
    {
        return true;
    }

    const double tolerance = 0.001d;
    if (string.Equals(placement.Side, "Left", StringComparison.Ordinal))
    {
        return placement.BalloonConnectorX >= placement.X + placement.Width - tolerance;
    }

    if (string.Equals(placement.Side, "Right", StringComparison.Ordinal))
    {
        return placement.BalloonConnectorX <= placement.X + tolerance;
    }

    return true;
}

static string? ResolveMarkupSide(IReadOnlyList<DocxMarkupBalloonPlacementSnapshot> placements)
{
    string[] sides = placements
        .Select(placement => placement.Side)
        .Where(side => !string.IsNullOrWhiteSpace(side))
        .Distinct(StringComparer.Ordinal)
        .OrderBy(side => side, StringComparer.Ordinal)
        .ToArray();
    return sides.Length switch
    {
        0 => null,
        1 => sides[0],
        _ => "Mixed"
    };
}

static string FlagBucket(string name, int count)
{
    return count == 0 ? $"{name}=no" : $"{name}=yes";
}

static string NormalizeBucketValue(string? value)
{
    return string.IsNullOrWhiteSpace(value) ? "(none)" : value.Trim();
}

static double? Ratio(double numerator, double denominator)
{
    return denominator <= 0d ? null : Math.Round(numerator / denominator, 6);
}

static double RoundPoint(double value)
{
    return Math.Round(value, 6);
}

static DocxTextEmissionCharacterProfile SumCharacterProfiles(IEnumerable<DocxTextEmissionSegmentSnapshot> segments)
{
    int digitCount = 0;
    int letterCount = 0;
    int whitespaceCount = 0;
    int punctuationCount = 0;
    int symbolCount = 0;
    int otherCount = 0;
    foreach (DocxTextEmissionSegmentSnapshot segment in segments)
    {
        digitCount += segment.CharacterProfile.DigitCount;
        letterCount += segment.CharacterProfile.LetterCount;
        whitespaceCount += segment.CharacterProfile.WhitespaceCount;
        punctuationCount += segment.CharacterProfile.PunctuationCount;
        symbolCount += segment.CharacterProfile.SymbolCount;
        otherCount += segment.CharacterProfile.OtherCount;
    }

    return new(digitCount, letterCount, whitespaceCount, punctuationCount, symbolCount, otherCount);
}

static object SumAdvanceProfiles(IEnumerable<DocxTextEmissionSegmentSnapshot> segments)
{
    int glyphCount = 0;
    int glyphGapCount = 0;
    double naturalPdfWidth = 0d;
    double roundedPdfWidth = 0d;
    double layoutWidth = 0d;
    double naturalResidual = 0d;
    double roundedResidual = 0d;
    foreach (DocxTextEmissionSegmentSnapshot segment in segments)
    {
        glyphCount += segment.AdvanceProfile.GlyphCount;
        glyphGapCount += segment.AdvanceProfile.GlyphGapCount;
        naturalPdfWidth += segment.AdvanceProfile.NaturalPdfWidth;
        roundedPdfWidth += segment.AdvanceProfile.RoundedPdfWidth;
        layoutWidth += segment.AdvanceProfile.LayoutWidth;
        naturalResidual += segment.AdvanceProfile.LayoutToNaturalResidual;
        roundedResidual += segment.AdvanceProfile.LayoutToRoundedResidual;
    }

    return new
    {
        GlyphCount = glyphCount,
        GlyphGapCount = glyphGapCount,
        NaturalPdfWidth = Math.Round(naturalPdfWidth, 6),
        RoundedPdfWidth = Math.Round(roundedPdfWidth, 6),
        LayoutWidth = Math.Round(layoutWidth, 6),
        LayoutToNaturalResidual = Math.Round(naturalResidual, 6),
        LayoutToRoundedResidual = Math.Round(roundedResidual, 6),
        UniformResidualPerGap = glyphGapCount == 0 ? (double?)null : Math.Round(naturalResidual / glyphGapCount, 6),
        RoundedResidualPerGap = glyphGapCount == 0 ? (double?)null : Math.Round(roundedResidual / glyphGapCount, 6)
    };
}
