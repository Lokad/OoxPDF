using System.Text.Encodings.Web;
using System.Text.Json;

using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;

if (args.Length < 2)
{
    Console.Error.WriteLine("Usage: Lokad.OoxPdf.DocxInspect <input.docx> <output-directory>");
    Environment.Exit(2);
}

string inputPath = Path.GetFullPath(args[0]);
string outputDirectory = Path.GetFullPath(args[1]);
Directory.CreateDirectory(outputDirectory);

using FileStream stream = File.OpenRead(inputPath);
OoxPackage package = OoxPackage.Open(stream);
DocxDocument document = new DocxReader().Read(package);
var renderer = new DocxRenderer();

var options = new JsonSerializerOptions
{
    WriteIndented = true,
    Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
};

DocxLayoutSnapshot layout = renderer.InspectLayout(document);
DocxFontPlanSnapshot fontPlan = renderer.InspectFontPlan(document);
DocxStructureSnapshot structure = renderer.InspectStructure(document);
DocxTextEmissionSnapshot textEmission = renderer.InspectTextEmission(document);
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
        LinesByPage = textEmission.Lines
            .GroupBy(line => line.PageIndex)
            .OrderBy(group => group.Key)
            .Select(group => new
            {
                PageIndex = group.Key,
                LineCount = group.Count(),
                StaticLineCount = group.Count(line => line.IsStaticStory),
                BodyLineCount = group.Count(line => !line.IsStaticStory),
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
