using System.Globalization;
using System.Text;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed record DocxLayout(IReadOnlyList<DocxLayoutPage> Pages);

internal sealed record DocxLineHeightProfile(
    double LineHeight,
    double? SingleLineHeight,
    double? ListLabelSingleLineHeight,
    double? BodyWindowsLineHeight,
    double? ListLabelWindowsLineHeight,
    double? EffectiveLineSpacingFactor,
    bool LineSpacingFactorFloorApplied);

internal sealed record DocxLayoutSnapshot(
    IReadOnlyList<DocxLayoutPageSnapshot> Pages,
    IReadOnlyList<DocxTableSnapshot> Tables,
    IReadOnlyList<DocxLayoutSourceBlockSnapshot> SourceBlocks)
{
    public static DocxLayoutSnapshot FromLayout(DocxLayout layout)
    {
        DocxLayoutPageSnapshot[] pages = layout.Pages.Select(ToSnapshot).ToArray();
        return new DocxLayoutSnapshot(pages, ToTableSnapshots(pages), ToSourceBlockSnapshots(pages));
    }

    private static IReadOnlyList<DocxTableSnapshot> ToTableSnapshots(IReadOnlyList<DocxLayoutPageSnapshot> pages)
    {
        return pages
            .SelectMany((page, pageIndex) => page.TableRows.Select(row => (pageIndex, row)))
            .GroupBy(entry => entry.row.TableIndex)
            .OrderBy(group => group.Key)
            .Select(group =>
            {
                DocxTableRowSnapshot first = group.First().row;
                DocxTableRowSnapshot[] distinctRows = group
                    .Select(entry => entry.row)
                    .GroupBy(row => row.RowIndex)
                    .Select(rowGroup => rowGroup.First())
                    .ToArray();
                return new DocxTableSnapshot(
                    group.Key,
                    first.SourceBlockIndex,
                    group.Min(entry => entry.pageIndex),
                    group.Max(entry => entry.pageIndex),
                    first.TableRowCount,
                    group.Count(),
                    group.Count(entry => entry.row.IsHeader),
                    distinctRows.Count(row => row.IsHeader),
                    first.GridColumnCount,
                    first.GridColumnsWidthSum,
                    first.HasExplicitGrid,
                    first.ResolvedColumnWidths,
                    first.ResolvedTableWidth,
                    first.TableX,
                    first.PreferredTableWidthPoints,
                    first.PreferredTableWidthValue,
                    first.PreferredTableWidthType,
                    first.IndentPoints,
                    first.CellSpacingPoints,
                    first.LayoutValue,
                    distinctRows.Count(row => row.DeclaredHeightPoints is not null),
                    distinctRows.Count(row => string.Equals(row.HeightRuleValue, "exact", StringComparison.OrdinalIgnoreCase)),
                    distinctRows.Count(row => string.Equals(row.HeightRuleValue, "atLeast", StringComparison.OrdinalIgnoreCase)),
                    distinctRows.Count(row => row.CantSplit),
                    distinctRows.Count(row => row.FragmentCount > 1),
                    group.Count(entry => entry.row.FragmentCount > 1),
                    group.Max(entry => entry.row.FragmentCount),
                    group.Any(entry => entry.row.Cells.Any(cell => cell.HasVerticalMerge)));
            })
            .ToArray();
    }

    private static IReadOnlyList<DocxLayoutSourceBlockSnapshot> ToSourceBlockSnapshots(IReadOnlyList<DocxLayoutPageSnapshot> pages)
    {
        return pages
            .SelectMany((page, pageIndex) => page.Items
                .Where(item => item.SourceBlockIndex is not null)
                .Select(item => (pageIndex, item)))
            .GroupBy(entry => entry.item.SourceBlockIndex!.Value)
            .OrderBy(group => group.Key)
            .Select(group => new DocxLayoutSourceBlockSnapshot(
                group.Key,
                ResolveSourceBlockKind(group.Select(entry => entry.item)),
                group.Min(entry => entry.pageIndex),
                group.Max(entry => entry.pageIndex),
                group.Max(entry => entry.item.Y + entry.item.Height),
                group.Min(entry => entry.item.Y),
                group.Count(),
                group.Count(entry => entry.item.Kind == "TextLine"),
                group.Count(entry => entry.item.Kind == "TableRow"),
                group.Count(entry => entry.item.Kind == "InlineImage"),
                group.Sum(entry => entry.item.TextLength),
                group.Sum(entry => entry.item.LineHeightPoints ?? entry.item.Height),
                group.Sum(entry => entry.item.AppliedBeforeSpacingPoints ?? 0d)))
            .ToArray();
    }

    private static string ResolveSourceBlockKind(IEnumerable<DocxLayoutItemSnapshot> items)
    {
        bool hasTextLine = false;
        bool hasTableRow = false;
        bool hasInlineImage = false;
        foreach (DocxLayoutItemSnapshot item in items)
        {
            hasTextLine |= item.Kind == "TextLine";
            hasTableRow |= item.Kind == "TableRow";
            hasInlineImage |= item.Kind == "InlineImage";
        }

        int kindCount = (hasTextLine ? 1 : 0) + (hasTableRow ? 1 : 0) + (hasInlineImage ? 1 : 0);
        if (kindCount > 1)
        {
            return "Mixed";
        }

        if (hasTableRow)
        {
            return "Table";
        }

        if (hasTextLine)
        {
            return "Paragraph";
        }

        return hasInlineImage ? "InlineImage" : "Unknown";
    }

    private static DocxLayoutPageSnapshot ToSnapshot(DocxLayoutPage page)
    {
        IReadOnlyList<DocxLayoutItemSnapshot> items = page.Items.Select(ToSnapshot).ToArray();
        int?[] sourceBlockIndexes = items
            .Select(item => item.SourceBlockIndex)
            .Where(index => index is not null)
            .ToArray();
        IReadOnlyList<DocxTableRowSnapshot> tableRows = page.Items
            .OfType<DocxTableRowLayout>()
            .Select((row, rowIndex) => ToTableRowSnapshot(row, rowIndex, page.MarginBottom))
            .ToArray();
        double verticalTop = items.Count == 0 ? 0d : items.Max(item => item.Y + item.Height);
        double verticalBottom = items.Count == 0 ? 0d : items.Min(item => item.Y);
        return new DocxLayoutPageSnapshot(
            page.Width,
            page.Height,
            page.MarginLeft,
            page.MarginRight,
            page.MarginTop,
            page.MarginBottom,
            items.Count,
            page.StaticTextLines.Count,
            items.Count(item => item.Kind == "TextLine"),
            items.Count(item => item.Kind == "InlineImage"),
            items.Count(item => item.Kind == "TableRow"),
            sourceBlockIndexes.Distinct().Count(),
            sourceBlockIndexes.FirstOrDefault(),
            sourceBlockIndexes.LastOrDefault(),
            Math.Max(0d, verticalTop - verticalBottom),
            items.Where(item => item.Kind == "TextLine").Sum(item => item.Height),
            items.Where(item => item.Kind == "InlineImage").Sum(item => item.Height),
            items.Where(item => item.Kind == "TableRow").Sum(item => item.Height),
            items,
            tableRows);
    }

    private static DocxLayoutItemSnapshot ToSnapshot(DocxLayoutItem item)
    {
        return item switch
        {
            DocxTextLineLayout text => new DocxLayoutItemSnapshot(
                "TextLine",
                text.X,
                text.BaselineY,
                text.Width,
                text.FontSize,
                TextLength: text.Text.Length,
                CellCount: 0,
                text.SourceBlockIndex,
                text.SourceLineIndex,
                text.LineHeight,
                text.AppliedBeforeSpacing,
                text.SingleLineHeight,
                text.ListLabelSingleLineHeight,
                text.BodyWindowsLineHeight,
                text.ListLabelWindowsLineHeight,
                text.EffectiveLineSpacingFactor,
                text.LineSpacingFactorFloorApplied,
                text.IsFirstParagraphLine),
            DocxInlineImageLayout image => new DocxLayoutItemSnapshot(
                "InlineImage",
                image.X,
                image.Y,
                image.Width,
                image.Height,
                TextLength: 0,
                CellCount: 0,
                image.SourceBlockIndex,
                SourceLineIndex: null,
                LineHeightPoints: null,
                AppliedBeforeSpacingPoints: null,
                SingleLineHeightPoints: null,
                ListLabelSingleLineHeightPoints: null,
                BodyWindowsLineHeightPoints: null,
                ListLabelWindowsLineHeightPoints: null,
                EffectiveLineSpacingFactor: null,
                LineSpacingFactorFloorApplied: null,
                IsFirstParagraphLine: null),
            DocxTableRowLayout row => new DocxLayoutItemSnapshot(
                "TableRow",
                row.Cells.Count == 0 ? 0d : row.Cells.Min(cell => cell.X),
                row.Y,
                row.Cells.Sum(cell => cell.Width),
                row.Height,
                TextLength: row.Cells.Sum(cell => cell.TextLines.Sum(line => line.Text.Length)),
                CellCount: row.Cells.Count,
                SourceBlockIndex: row.Table.SourceBlockIndex,
                SourceLineIndex: null,
                LineHeightPoints: null,
                AppliedBeforeSpacingPoints: null,
                SingleLineHeightPoints: null,
                ListLabelSingleLineHeightPoints: null,
                BodyWindowsLineHeightPoints: null,
                ListLabelWindowsLineHeightPoints: null,
                EffectiveLineSpacingFactor: null,
                LineSpacingFactorFloorApplied: null,
                IsFirstParagraphLine: null),
            _ => new DocxLayoutItemSnapshot("Unknown", 0d, 0d, 0d, 0d, 0, 0, null, null, null, null, null, null, null, null, null, null, null)
        };
    }

    private static DocxTableRowSnapshot ToTableRowSnapshot(DocxTableRowLayout row, int rowIndex, double pageMarginBottom)
    {
        IReadOnlyList<DocxTableCellSnapshot> cells = row.Cells
            .Select(ToTableCellSnapshot)
            .ToArray();
        double? firstBaselineY = cells
            .Select(cell => cell.FirstBaselineY)
            .Where(baseline => baseline is not null)
            .DefaultIfEmpty(null)
            .Min();
        double? lastBaselineY = cells
            .Select(cell => cell.LastBaselineY)
            .Where(baseline => baseline is not null)
            .DefaultIfEmpty(null)
            .Min();
        return new DocxTableRowSnapshot(
            row.Table.TableIndex,
            row.Table.SourceBlockIndex,
            rowIndex,
            row.RowIndex,
            row.Table.RowCount,
            row.Table.GridColumnCount,
            row.Table.GridColumnsWidthSum,
            row.Table.HasExplicitGrid,
            row.Table.ResolvedColumnWidths,
            row.Table.ResolvedTableWidth,
            row.Table.TableX,
            row.Table.PreferredWidthPoints,
            row.Table.PreferredWidthValue,
            row.Table.PreferredWidthType,
            row.Table.IndentPoints,
            row.Table.CellSpacingPoints,
            row.Table.LayoutValue,
            row.HeightValue,
            row.HeightRuleValue,
            row.DeclaredHeightPoints,
            row.FragmentIndex,
            row.FragmentCount,
            row.Cells.Count == 0 ? 0d : row.Cells.Min(cell => cell.X),
            row.Y,
            row.Cells.Sum(cell => cell.Width),
            row.Height,
            Math.Max(0d, pageMarginBottom - row.Y),
            firstBaselineY,
            lastBaselineY,
            cells.Count,
            cells.Sum(cell => cell.TextLineCount),
            cells.Sum(cell => cell.TextLength),
            cells.Select(cell => cell.MaxFontSize).DefaultIfEmpty(0d).Max(),
            row.IsHeader,
            row.HeaderValue,
            row.HasTablePropertyExceptionCellMargins,
            row.CantSplit,
            row.CantSplitValue,
            cells);
    }

    private static DocxTableCellSnapshot ToTableCellSnapshot(DocxTableCellLayout cellLayout, int cellIndex)
    {
        DocxTableCell cell = cellLayout.Cell;
        DocxTextLineLayout? firstLine = cellLayout.TextLines.FirstOrDefault();
        DocxTextLineLayout? lastLine = cellLayout.TextLines.LastOrDefault();
        IReadOnlyList<DocxParagraph> paragraphs = DocxTableCellContent.GetParagraphs(cell);
        IReadOnlyList<double> spacingBeforePoints = paragraphs.Select(paragraph => paragraph.SpacingBeforePoints).ToArray();
        IReadOnlyList<double> spacingAfterPoints = paragraphs.Select(paragraph => paragraph.SpacingAfterPoints).ToArray();
        string cellText = string.Concat(paragraphs.SelectMany(paragraph => paragraph.Runs).Select(run => run.Text));
        TextProfile textProfile = BuildTextProfile(cellText);
        return new DocxTableCellSnapshot(
            cellIndex,
            cellLayout.X,
            cellLayout.Y,
            cellLayout.Width,
            cellLayout.Height,
            cellLayout.TextLines.Count,
            cellLayout.TextLines.Sum(line => line.Text.Length),
            cellLayout.TextLines.Count == 0 ? 0d : cellLayout.TextLines.Max(line => line.FontSize),
            firstLine?.X,
            firstLine?.BaselineY,
            lastLine?.BaselineY,
            cellLayout.InlineImages.Count,
            paragraphs.Count,
            paragraphs.Count(paragraph => HasBeforeSpacingToken(paragraph.Spacing)),
            paragraphs.Count(paragraph => HasAfterSpacingToken(paragraph.Spacing)),
            paragraphs.Count(paragraph => string.Equals(paragraph.Spacing.BeforeValue, "0", StringComparison.Ordinal)),
            paragraphs.Count(paragraph => string.Equals(paragraph.Spacing.AfterValue, "0", StringComparison.Ordinal)),
            spacingBeforePoints.Count == 0 ? null : spacingBeforePoints.Min(),
            spacingBeforePoints.Count == 0 ? null : spacingBeforePoints.Max(),
            spacingAfterPoints.Count == 0 ? null : spacingAfterPoints.Min(),
            spacingAfterPoints.Count == 0 ? null : spacingAfterPoints.Max(),
            cell.GridSpan,
            cell.GridSpanValue,
            cell.PreferredWidthPoints,
            cell.PreferredWidthValue,
            cell.PreferredWidthType,
            cell.VerticalAlignmentValue,
            cell.Margins.TopPoints,
            cell.Margins.RightPoints,
            cell.Margins.BottomPoints,
            cell.Margins.LeftPoints,
            cell.Borders.Count,
            cell.FillHex is not null,
            cell.ShadingValue is not null,
            cell.ConditionalFormat?.IsDefined == true,
            cell.HasVerticalMerge,
            cell.VerticalMergeValue,
            cellLayout.IsVerticalMergeContinuation,
            textProfile.SpaceCharacterCount,
            textProfile.NonAsciiCharacterCount,
            textProfile.PunctuationCharacterCount,
            textProfile.DigitCharacterCount,
            textProfile.UppercaseCharacterCount,
            textProfile.LowercaseCharacterCount,
            textProfile.LongestBreakableTokenLength);
    }

    private static TextProfile BuildTextProfile(string text)
    {
        int spaceCharacterCount = 0;
        int nonAsciiCharacterCount = 0;
        int punctuationCharacterCount = 0;
        int digitCharacterCount = 0;
        int uppercaseCharacterCount = 0;
        int lowercaseCharacterCount = 0;
        foreach (char value in text)
        {
            if (char.IsWhiteSpace(value))
            {
                spaceCharacterCount++;
            }

            if (value > 0x7f)
            {
                nonAsciiCharacterCount++;
            }

            if (char.IsPunctuation(value))
            {
                punctuationCharacterCount++;
            }

            if (char.IsDigit(value))
            {
                digitCharacterCount++;
            }

            if (char.IsUpper(value))
            {
                uppercaseCharacterCount++;
            }

            if (char.IsLower(value))
            {
                lowercaseCharacterCount++;
            }
        }

        int longestBreakableTokenLength = ResolveLongestBreakableTokenLength(text);
        return new TextProfile(
            spaceCharacterCount,
            nonAsciiCharacterCount,
            punctuationCharacterCount,
            digitCharacterCount,
            uppercaseCharacterCount,
            lowercaseCharacterCount,
            longestBreakableTokenLength);
    }

    private static int ResolveLongestBreakableTokenLength(string text)
    {
        int longest = 0;
        int current = 0;
        foreach (char value in text)
        {
            if (char.IsWhiteSpace(value) &&
                value != '\u00A0' &&
                value != '\u202F' &&
                value != '\u2007')
            {
                longest = Math.Max(longest, current);
                current = 0;
            }
            else
            {
                current++;
            }
        }

        return Math.Max(longest, current);
    }

    private static bool HasBeforeSpacingToken(DocxParagraphSpacing spacing)
    {
        return spacing.BeforeValue is not null ||
            spacing.BeforeLinesValue is not null ||
            spacing.BeforeAutoSpacingValue is not null;
    }

    private static bool HasAfterSpacingToken(DocxParagraphSpacing spacing)
    {
        return spacing.AfterValue is not null ||
            spacing.AfterLinesValue is not null ||
            spacing.AfterAutoSpacingValue is not null;
    }
}

internal sealed record DocxLayoutPageSnapshot(
    double Width,
    double Height,
    double MarginLeft,
    double MarginRight,
    double MarginTop,
    double MarginBottom,
    int ItemCount,
    int StaticTextLineCount,
    int TextLineCount,
    int InlineImageCount,
    int TableRowCount,
    int SourceBlockCount,
    int? FirstSourceBlockIndex,
    int? LastSourceBlockIndex,
    double VerticalUsed,
    double TextLineHeightSum,
    double InlineImageHeightSum,
    double TableRowHeightSum,
    IReadOnlyList<DocxLayoutItemSnapshot> Items,
    IReadOnlyList<DocxTableRowSnapshot> TableRows);

internal sealed record DocxLayoutSourceBlockSnapshot(
    int SourceBlockIndex,
    string Kind,
    int FirstPageIndex,
    int LastPageIndex,
    double VerticalTop,
    double VerticalBottom,
    int ItemCount,
    int TextLineCount,
    int TableRowCount,
    int InlineImageCount,
    int TextLength,
    double ConsumedHeightSum,
    double AppliedBeforeSpacingSum);

internal sealed record DocxLayoutItemSnapshot(
    string Kind,
    double X,
    double Y,
    double Width,
    double Height,
    int TextLength,
    int CellCount,
    int? SourceBlockIndex,
    int? SourceLineIndex,
    double? LineHeightPoints,
    double? AppliedBeforeSpacingPoints,
    double? SingleLineHeightPoints,
    double? ListLabelSingleLineHeightPoints,
    double? BodyWindowsLineHeightPoints,
    double? ListLabelWindowsLineHeightPoints,
    double? EffectiveLineSpacingFactor,
    bool? LineSpacingFactorFloorApplied,
    bool? IsFirstParagraphLine);

internal sealed record DocxTableRowSnapshot(
    int TableIndex,
    int SourceBlockIndex,
    int PageRowIndex,
    int RowIndex,
    int TableRowCount,
    int GridColumnCount,
    double GridColumnsWidthSum,
    bool HasExplicitGrid,
    IReadOnlyList<double> ResolvedColumnWidths,
    double ResolvedTableWidth,
    double TableX,
    double? PreferredTableWidthPoints,
    string? PreferredTableWidthValue,
    string? PreferredTableWidthType,
    double? IndentPoints,
    double? CellSpacingPoints,
    string? LayoutValue,
    string? HeightValue,
    string? HeightRuleValue,
    double? DeclaredHeightPoints,
    int FragmentIndex,
    int FragmentCount,
    double X,
    double Y,
    double Width,
    double Height,
    double BottomOverflowPoints,
    double? FirstBaselineY,
    double? LastBaselineY,
    int CellCount,
    int TextLineCount,
    int TextLength,
    double MaxFontSize,
    bool IsHeader,
    string? HeaderValue,
    bool HasTablePropertyExceptionCellMargins,
    bool CantSplit,
    string? CantSplitValue,
    IReadOnlyList<DocxTableCellSnapshot> Cells);

internal sealed record DocxTableSnapshot(
    int TableIndex,
    int SourceBlockIndex,
    int PageStartIndex,
    int PageEndIndex,
    int RowCount,
    int LaidOutRowCount,
    int HeaderRowLayoutCount,
    int AuthoredHeaderRowCount,
    int GridColumnCount,
    double GridColumnsWidthSum,
    bool HasExplicitGrid,
    IReadOnlyList<double> ResolvedColumnWidths,
    double ResolvedTableWidth,
    double X,
    double? PreferredWidthPoints,
    string? PreferredWidthValue,
    string? PreferredWidthType,
    double? IndentPoints,
    double? CellSpacingPoints,
    string? LayoutValue,
    int DeclaredHeightRowCount,
    int ExactHeightRowCount,
    int AtLeastHeightRowCount,
    int CantSplitRowCount,
    int FragmentedRowCount,
    int FragmentedRowLayoutCount,
    int MaxRowFragmentCount,
    bool HasVerticalMerge);

internal sealed record DocxTableCellSnapshot(
    int CellIndex,
    double X,
    double Y,
    double Width,
    double Height,
    int TextLineCount,
    int TextLength,
    double MaxFontSize,
    double? FirstTextLineX,
    double? FirstBaselineY,
    double? LastBaselineY,
    int InlineImageCount,
    int ParagraphCount,
    int ParagraphsWithBeforeSpacingToken,
    int ParagraphsWithAfterSpacingToken,
    int ParagraphsWithZeroBeforeSpacing,
    int ParagraphsWithZeroAfterSpacing,
    double? MinSpacingBeforePoints,
    double? MaxSpacingBeforePoints,
    double? MinSpacingAfterPoints,
    double? MaxSpacingAfterPoints,
    int GridSpan,
    string? GridSpanValue,
    double? PreferredWidthPoints,
    string? PreferredWidthValue,
    string? PreferredWidthType,
    string? VerticalAlignmentValue,
    double? MarginTopPoints,
    double? MarginRightPoints,
    double? MarginBottomPoints,
    double? MarginLeftPoints,
    int BorderCount,
    bool HasFill,
    bool HasShadingValue,
    bool HasConditionalFormat,
    bool HasVerticalMerge,
    string? VerticalMergeValue,
    bool IsVerticalMergeContinuation,
    int SpaceCharacterCount,
    int NonAsciiCharacterCount,
    int PunctuationCharacterCount,
    int DigitCharacterCount,
    int UppercaseCharacterCount,
    int LowercaseCharacterCount,
    int LongestBreakableTokenLength);

internal sealed record DocxTextEmissionSnapshot(
    int LineCount,
    int SegmentCount,
    int TerminalSpaceSegmentCount,
    int NonzeroPdfCharacterSpacingSegmentCount,
    int CompensatedCharacterSpacingSegmentCount,
    IReadOnlyList<DocxTextEmissionLineSnapshot> Lines);

internal sealed record DocxTextEmissionLineSnapshot(
    int PageIndex,
    bool IsStaticStory,
    int? SourceBlockIndex,
    int? SourceLineIndex,
    bool EndsWithIntraTokenBreak,
    int SegmentCount,
    int TextLength,
    int TerminalSpaceSegmentCount,
    int NonzeroPdfCharacterSpacingSegmentCount,
    IReadOnlyList<DocxTextEmissionSegmentSnapshot> Segments);

internal sealed record DocxTextEmissionSegmentSnapshot(
    int TextLength,
    double X,
    double BaselineY,
    double Width,
    double FontSize,
    double PdfFontSize,
    double LayoutCharacterSpacing,
    double PdfCharacterSpacing,
    double PositioningCharacterSpacing,
    bool CompensatePdfCharacterSpacing,
    bool IsTerminalLineSpace,
    string? FontResourceName,
    bool SyntheticBold,
    bool SyntheticItalic);

internal sealed record TextProfile(
    int SpaceCharacterCount,
    int NonAsciiCharacterCount,
    int PunctuationCharacterCount,
    int DigitCharacterCount,
    int UppercaseCharacterCount,
    int LowercaseCharacterCount,
    int LongestBreakableTokenLength);

internal sealed record DocxLayoutPage(
    double Width,
    double Height,
    double MarginLeft,
    double MarginRight,
    double MarginTop,
    double MarginBottom,
    DocxPageSettings PageSettings,
    IReadOnlyList<DocxTextLineLayout> StaticTextLines,
    IReadOnlyList<DocxLayoutItem> Items);

internal abstract record DocxLayoutItem;

internal sealed record DocxTextLineLayout(
    string Text,
    DocxTextRun StyleRun,
    double FontSize,
    double X,
    double BaselineY,
    double Width,
    IReadOnlyList<DocxTextSegmentLayout> Segments,
    int? SourceBlockIndex = null,
    int? SourceLineIndex = null,
    double? LineHeight = null,
    double? AppliedBeforeSpacing = null,
    bool? IsFirstParagraphLine = null,
    bool EndsWithIntraTokenBreak = false,
    double? SingleLineHeight = null,
    double? ListLabelSingleLineHeight = null,
    double? BodyWindowsLineHeight = null,
    double? ListLabelWindowsLineHeight = null,
    double? EffectiveLineSpacingFactor = null,
    bool? LineSpacingFactorFloorApplied = null) : DocxLayoutItem;

internal sealed record DocxTextSegmentLayout(
    string Text,
    DocxTextRun StyleRun,
    double X,
    double Width,
    double? FontSize = null,
    double BaselineOffsetY = 0d,
    double PdfCharacterSpacing = 0d,
    bool CompensatePdfCharacterSpacing = true);

internal sealed record DocxTextSpan(
    string Text,
    DocxTextRun StyleRun);

internal sealed record DocxWrappedTextLine(
    string Text,
    IReadOnlyList<DocxTextSpan> Spans,
    bool EndsWithIntraTokenBreak = false);

internal sealed record DocxInlineImageLayout(
    DocxInlineImage Image,
    double X,
    double Y,
    double Width,
    double Height,
    int PageIndex,
    int? SourceBlockIndex = null) : DocxLayoutItem;

internal sealed record DocxTableRowLayout(
    DocxTableLayoutContext Table,
    int RowIndex,
    int FragmentIndex,
    int FragmentCount,
    IReadOnlyList<DocxTableCellLayout> Cells,
    double Y,
    double Height,
    double? DeclaredHeightPoints,
    string? HeightValue,
    string? HeightRuleValue,
    bool IsHeader,
    string? HeaderValue,
    bool HasTablePropertyExceptionCellMargins,
    bool CantSplit,
    string? CantSplitValue) : DocxLayoutItem;

internal sealed record DocxTableLayoutContext(
    int TableIndex,
    int SourceBlockIndex,
    int RowCount,
    int GridColumnCount,
    double GridColumnsWidthSum,
    bool HasExplicitGrid,
    IReadOnlyList<double> ResolvedColumnWidths,
    double ResolvedTableWidth,
    double TableX,
    double? PreferredWidthPoints,
    string? PreferredWidthValue,
    string? PreferredWidthType,
    double? IndentPoints,
    double? CellSpacingPoints,
    string? LayoutValue);

internal sealed record DocxTableCellLayout(
    DocxTableCell Cell,
    double X,
    double Y,
    double Width,
    double Height,
    IReadOnlyList<DocxTextLineLayout> TextLines,
    IReadOnlyList<DocxInlineImageLayout> InlineImages,
    bool IsVerticalMergeContinuation = false,
    DocxTableCell? VerticalMergeOwnerCell = null);

internal sealed record DocxRunFontResource(string Name, PdfEmbeddedFont Embedded, FontResolution Resolution);

internal sealed record DocxFontResources(
    DocxFontPlan Plan,
    IDocxTextMeasurer? TextMeasurer,
    IReadOnlyList<PdfFontResource> Resources,
    IReadOnlyDictionary<DocxTextRun, DocxRunFontResource> RunResources,
    DocxRunFontResource? Fallback);

internal sealed record DocxTextEmissionSegment(
    string Text,
    DocxTextRun StyleRun,
    DocxRunFontResource Resource,
    RgbColor Color,
    double X,
    double BaselineY,
    double Width,
    double FontSize,
    double PdfCharacterSpacing,
    bool CompensatePdfCharacterSpacing,
    bool SyntheticBold,
    bool SyntheticItalic,
    bool IsTerminalLineSpace);

internal readonly record struct DocxKeepBlockEstimate(
    double Height,
    int ParagraphCount,
    int FirstTableRowCount);

internal interface IDocxTextMeasurer
{
    double MeasureText(DocxTextRun? run, string text, double fontSize);
}

internal interface IDocxLineMetricsProvider
{
    double MeasureSingleLineHeight(DocxTextRun? run, double fontSize);
}

internal interface IDocxStaticTextMetricsProvider
{
    double MeasureWindowsAscender(DocxTextRun? run, double fontSize);

    double MeasureWindowsDescender(DocxTextRun? run, double fontSize);
}

internal static class DocxTextSpacing
{
    public static double AddCharacterSpacing(double measuredWidth, DocxTextRun? run, string text)
    {
        return measuredWidth + CountCharacterSpacingGaps(text) * (run?.CharacterSpacingPoints ?? 0d);
    }

    public static double BoundarySpacing(DocxTextRun? left, string leftText, string rightText)
    {
        return leftText.Length == 0 ||
            rightText.Length == 0 ||
            leftText[^1] == '\t' ||
            rightText[0] == '\t'
            ? 0d
            : left?.CharacterSpacingPoints ?? 0d;
    }

    private static int CountCharacterSpacingGaps(string text)
    {
        int count = 0;
        foreach (Rune _ in text.EnumerateRunes())
        {
            count++;
        }

        return Math.Max(0, count - 1);
    }
}

internal static class DocxLineMetrics
{
    private const double WordSingleLineMinimumEm = 1.15d;
    private const double WordAutoLineBaselineOffsetEm = 0.94d;
    private const double WordExactLineTextBottomInsetEm = 0.299d;

    public static double MeasureOpenTypeSingleLineHeight(OpenTypeFont font, double fontSize)
    {
        if (font.UnitsPerEm == 0)
        {
            return fontSize * WordSingleLineMinimumEm;
        }

        double units = font.Os2.TypographicAscender - font.Os2.TypographicDescender + font.Os2.TypographicLineGap;

        return Math.Max(fontSize * WordSingleLineMinimumEm, units * fontSize / font.UnitsPerEm);
    }

    public static double MeasureWindowsAscender(OpenTypeFont font, double fontSize)
    {
        return font.UnitsPerEm == 0
            ? fontSize
            : font.Os2.WindowsAscender * fontSize / font.UnitsPerEm;
    }

    public static double MeasureWindowsDescender(OpenTypeFont font, double fontSize)
    {
        return font.UnitsPerEm == 0
            ? 0d
            : font.Os2.WindowsDescender * fontSize / font.UnitsPerEm;
    }

    public static double ResolveBodyBaselineOffset(double fontSize, double lineHeight, bool hasExplicitLineSpacing)
    {
        return hasExplicitLineSpacing
            ? Math.Max(0d, lineHeight - fontSize * WordExactLineTextBottomInsetEm)
            : fontSize * WordAutoLineBaselineOffsetEm;
    }

    public static double ResolveTableCellFirstBaselineInset(IReadOnlyList<DocxParagraph> paragraphs)
    {
        DocxParagraph? firstTextParagraph = paragraphs.FirstOrDefault(paragraph => paragraph.Runs.Count != 0);
        return firstTextParagraph is null ? 0d : firstTextParagraph.Runs.Max(run => run.FontSize);
    }
}

internal static class DocxVerticalAlignMetrics
{
    private const double SuperscriptSubscriptScale = 2d / 3d;
    private const double HalfPointGrid = 2d;
    private const double SubscriptBaselineShiftEm = -0.06d;

    public static double ResolveFontSize(double nominalFontSize, DocxTextRun run)
    {
        if (!IsSuperscript(run) && !IsSubscript(run))
        {
            return nominalFontSize;
        }

        return Math.Max(0.5d, Math.Floor(nominalFontSize * SuperscriptSubscriptScale * HalfPointGrid) / HalfPointGrid);
    }

    public static double ResolveBaselineOffset(double nominalFontSize, double layoutFontSize, DocxTextRun run)
    {
        if (IsSuperscript(run))
        {
            return Math.Max(0d, nominalFontSize - layoutFontSize);
        }

        return IsSubscript(run) ? nominalFontSize * SubscriptBaselineShiftEm : 0d;
    }

    private static bool IsSuperscript(DocxTextRun run)
    {
        return run.VerticalAlignmentValue?.Equals("superscript", StringComparison.OrdinalIgnoreCase) == true;
    }

    private static bool IsSubscript(DocxTextRun run)
    {
        return run.VerticalAlignmentValue?.Equals("subscript", StringComparison.OrdinalIgnoreCase) == true;
    }
}

internal sealed class DocxEmbeddedTextMeasurer(PdfEmbeddedFont embedded) : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider
{
    public double MeasureText(DocxTextRun? run, string text, double fontSize)
    {
        return DocxTextSpacing.AddCharacterSpacing(embedded.MeasureTextPoints(text, fontSize), run, text);
    }

    public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize)
    {
        return DocxLineMetrics.MeasureOpenTypeSingleLineHeight(embedded.Font, fontSize);
    }

    public double MeasureWindowsAscender(DocxTextRun? run, double fontSize)
    {
        return DocxLineMetrics.MeasureWindowsAscender(embedded.Font, fontSize);
    }

    public double MeasureWindowsDescender(DocxTextRun? run, double fontSize)
    {
        return DocxLineMetrics.MeasureWindowsDescender(embedded.Font, fontSize);
    }
}

internal sealed class DocxLayoutEngine
{
    private const double WordDefaultTabStopPoints = 36d;
    private const double WordListMinimumAutoLineSpacingFactor = 1.19d;

    private sealed record DocxPageGeometry(
        double Width,
        double Height,
        double MarginLeft,
        double MarginRight,
        double MarginTop,
        double MarginBottom,
        DocxPageSettings PageSettings)
    {
        public double BodyWidth => Math.Max(1d, Width - MarginLeft - MarginRight);
    }

    public DocxLayout Create(DocxDocument document, PdfEmbeddedFont? embedded)
    {
        IDocxTextMeasurer? textMeasurer = embedded is null ? null : new DocxEmbeddedTextMeasurer(embedded);
        return Create(document, textMeasurer);
    }

    internal DocxLayout Create(DocxDocument document, IDocxTextMeasurer? textMeasurer)
    {
        var pages = new List<DocxLayoutPage>();
        var currentItems = new List<DocxLayoutItem>();
        IReadOnlyDictionary<int, DocxPageSettings> sectionSettingsByElementIndex = BuildEffectiveSectionSettings(document, out DocxPageSettings finalSectionSettings);
        DocxPageGeometry page = ResolveSectionGeometry(document, FindSectionSettingsAtOrAfter(document.BodyElements, 0, sectionSettingsByElementIndex) ?? finalSectionSettings);
        double x = page.MarginLeft;
        double width = page.BodyWidth;
        double cursorY = page.Height - page.MarginTop;
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        int tableIndex = 0;
        double defaultTabStopPoints = document.Settings.DefaultTabStopPoints ?? WordDefaultTabStopPoints;

        void FinishPage()
        {
            pages.Add(new DocxLayoutPage(
                page.Width,
                page.Height,
                page.MarginLeft,
                page.MarginRight,
                page.MarginTop,
                page.MarginBottom,
                page.PageSettings,
                [],
                currentItems.ToArray()));
            currentItems = [];
            cursorY = page.Height - page.MarginTop;
            pendingSpacingAfter = 0d;
            previousParagraph = null;
        }

        void ApplySectionAfterBreak(int elementIndex)
        {
            page = ResolveSectionGeometry(document, FindSectionSettingsAtOrAfter(document.BodyElements, elementIndex + 1, sectionSettingsByElementIndex) ?? finalSectionSettings);
            x = page.MarginLeft;
            width = page.BodyWidth;
            if (!HasPageContent())
            {
                cursorY = page.Height - page.MarginTop;
            }
        }

        bool HasPageContent() => currentItems.Count > 0;

        for (int elementIndex = 0; elementIndex < document.BodyElements.Count; elementIndex++)
        {
            DocxBodyElement element = document.BodyElements[elementIndex];
            if (element is DocxPageBreakElement pageBreak)
            {
                if (pageBreak.BreakParagraph is { } breakParagraph)
                {
                    double breakBeforeSpacing = ShouldSuppressContextualSpacing(previousParagraph, breakParagraph)
                        ? 0d
                        : Math.Max(pendingSpacingAfter, breakParagraph.SpacingBeforePoints);
                    double breakFontSize = GetParagraphFontSize(breakParagraph);
                    double breakLineHeight = ResolveLineHeight(breakParagraph, breakFontSize, textMeasurer);
                    double paragraphAdvance = breakBeforeSpacing + breakLineHeight;
                    if (cursorY - paragraphAdvance < page.MarginBottom && HasPageContent())
                    {
                        FinishPage();
                    }

                    cursorY -= paragraphAdvance;
                }

                if (HasPageContent() || pageBreak.BreakParagraph is not null)
                {
                    FinishPage();
                }

                pendingSpacingAfter = 0d;
                previousParagraph = null;
                continue;
            }

            if (element is DocxManualBreakElement)
            {
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                continue;
            }

            if (element is DocxSectionBreakElement sectionBreak)
            {
                bool startsNewPage = ShouldStartNewPageForSectionBreak(sectionBreak);
                if (startsNewPage && HasPageContent())
                {
                    FinishPage();
                }

                if (startsNewPage)
                {
                    ApplySectionAfterBreak(elementIndex);
                }

                pendingSpacingAfter = 0d;
                previousParagraph = null;
                continue;
            }

            if (element is DocxTableElement tableElement)
            {
                cursorY -= pendingSpacingAfter;
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                LayoutTable(tableElement.Table, tableIndex++, elementIndex, page.MarginBottom, page.Height - page.MarginTop - page.MarginBottom, textMeasurer, defaultTabStopPoints, () => pages.Count + 1, ref currentItems, ref cursorY, x, width, FinishPage, HasPageContent);
                continue;
            }

            if (element is not DocxParagraphElement paragraphElement)
            {
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            double appliedBeforeSpacing = ShouldSuppressContextualSpacing(previousParagraph, paragraph)
                ? 0d
                : Math.Max(pendingSpacingAfter, paragraph.SpacingBeforePoints);
            cursorY -= appliedBeforeSpacing;
            pendingSpacingAfter = 0d;
            double paragraphFontSize = GetParagraphFontSize(paragraph);
            DocxLineHeightProfile lineHeightProfile = ResolveLineHeightProfile(paragraph, paragraphFontSize, textMeasurer);
            double lineHeight = lineHeightProfile.LineHeight;
            if (textMeasurer is not null &&
                HasPageContent() &&
                ShouldKeepParagraphBlockTogether(paragraph) &&
                cursorY - EstimateKeptParagraphBlock(document.BodyElements, elementIndex, width, textMeasurer, defaultTabStopPoints).Height <= page.MarginBottom)
            {
                FinishPage();
            }

            IReadOnlyList<DocxTextSpan> textSpans = textMeasurer is null ? [] : CreateTextSpans(paragraph.Runs);
            if (textMeasurer is not null && textSpans.Count > 0)
            {
                double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, paragraphFontSize, textMeasurer);
                double continuationTextStartOffset = GetParagraphTextStartOffset(paragraph);
                double labelStartOffset = GetParagraphLabelStartOffset(paragraph);
                double paragraphX = x + textStartOffset;
                double paragraphWidth = Math.Max(1d, width - textStartOffset - GetParagraphRightInset(paragraph));
                DocxTextRun firstRun = paragraph.Runs[0];
                bool firstLine = true;
                double continuationParagraphWidth = Math.Max(1d, width - continuationTextStartOffset - GetParagraphRightInset(paragraph));
                DocxWrappedTextLine[] lines = WrapTextLines(textSpans, paragraphWidth, continuationParagraphWidth, paragraphFontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints, allowOverwideTokenBreaks: false).ToArray();
                if (ShouldMoveParagraphForWidowControl(paragraph, lines.Length, cursorY, lineHeight, page.MarginBottom, HasPageContent()))
                {
                    FinishPage();
                }

                for (int lineIndex = 0; lineIndex < lines.Length; lineIndex++)
                {
                    DocxWrappedTextLine line = lines[lineIndex];
                    if (cursorY - lineHeight < page.MarginBottom && HasPageContent())
                    {
                        FinishPage();
                    }

                    double lineWidth = MeasureTextSpans(line.Spans, paragraphFontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints);
                    double drawableLineWidth = MeasureDrawableTextSpans(line.Spans, paragraphFontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints);
                    double lineX = paragraph.Alignment switch
                    {
                        DocxTextAlignment.Center => paragraphX + Math.Max(0, paragraphWidth - lineWidth) / 2d,
                        DocxTextAlignment.Right => paragraphX + Math.Max(0, paragraphWidth - lineWidth),
                        _ => paragraphX
                    };
                    double baselineOffset = DocxLineMetrics.ResolveBodyBaselineOffset(paragraphFontSize, lineHeight, paragraph.LineSpacingPoints is not null);
                    bool justifyLine = (paragraph.ListLabel is null || !firstLine) &&
                        ShouldJustifyTextLine(paragraph.Alignment, lineIndex == lines.Length - 1, drawableLineWidth, paragraphWidth, line.Spans);
                    IReadOnlyList<DocxTextSegmentLayout> segments = firstLine && paragraph.ListLabel is not null
                        ? CreateNumberedLineSegments(paragraph.ListLabel, line.Spans, firstRun, x + labelStartOffset, lineX, paragraphFontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints)
                        : justifyLine
                            ? CreateJustifiedTextSegments(line.Spans, lineX, drawableLineWidth, paragraphWidth, paragraphFontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints)
                            : CreateTextSegments(line.Spans, lineX, paragraphFontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints);
                    double effectiveX = firstLine && paragraph.ListLabel is not null ? x + labelStartOffset : lineX;
                    double effectiveWidth = firstLine && paragraph.ListLabel is not null
                        ? Math.Max(lineX + lineWidth, x + labelStartOffset + MeasureListLabel(paragraph.ListLabel, firstRun, paragraphFontSize, textMeasurer)) - (x + labelStartOffset)
                        : justifyLine
                            ? paragraphWidth
                            : lineWidth;
                    currentItems.Add(new DocxTextLineLayout(
                        firstLine && paragraph.ListLabel is not null ? paragraph.ListLabel.Text + GetListLabelTextSeparator(paragraph.ListLabel) + line.Text : line.Text,
                        firstRun,
                        paragraphFontSize,
                        effectiveX,
                        cursorY - baselineOffset,
                        effectiveWidth,
                        segments,
                        elementIndex,
                        lineIndex,
                        lineHeight,
                        firstLine ? appliedBeforeSpacing : 0d,
                        firstLine,
                        line.EndsWithIntraTokenBreak,
                        lineHeightProfile.SingleLineHeight,
                        lineHeightProfile.ListLabelSingleLineHeight,
                        lineHeightProfile.BodyWindowsLineHeight,
                        lineHeightProfile.ListLabelWindowsLineHeight,
                        lineHeightProfile.EffectiveLineSpacingFactor,
                        lineHeightProfile.LineSpacingFactorFloorApplied));
                    firstLine = false;
                    paragraphX = x + continuationTextStartOffset;
                    paragraphWidth = Math.Max(1d, width - continuationTextStartOffset - GetParagraphRightInset(paragraph));
                    cursorY -= lineHeight;
                }
            }
            else if (paragraph.Images.Count == 0)
            {
                if (cursorY - lineHeight < page.MarginBottom && HasPageContent())
                {
                    FinishPage();
                }

                cursorY -= lineHeight;
            }

            foreach (DocxInlineImage image in paragraph.Images)
            {
                double imageWidth = Math.Min(width, image.WidthPoints);
                double imageHeight = image.HeightPoints * imageWidth / Math.Max(1d, image.WidthPoints);
                if (cursorY - imageHeight < page.MarginBottom && HasPageContent())
                {
                    FinishPage();
                }

                double imageX = paragraph.Alignment switch
                {
                    DocxTextAlignment.Center => x + Math.Max(0, width - imageWidth) / 2d,
                    DocxTextAlignment.Right => x + Math.Max(0, width - imageWidth),
                    _ => x
                };
                currentItems.Add(new DocxInlineImageLayout(
                    image,
                    imageX,
                    cursorY - imageHeight,
                    imageWidth,
                    imageHeight,
                    pages.Count + 1,
                    SourceBlockIndex: elementIndex));
                cursorY -= imageHeight + 6d;
            }

            pendingSpacingAfter = paragraph.SpacingAfterPoints;
            previousParagraph = paragraph;
        }

        if (HasPageContent() || pages.Count == 0)
        {
            FinishPage();
        }

        return new DocxLayout(AddStaticTextLines(pages, textMeasurer).ToArray());
    }

    private static IReadOnlyList<DocxLayoutPage> AddStaticTextLines(IReadOnlyList<DocxLayoutPage> pages, IDocxTextMeasurer? textMeasurer)
    {
        if (textMeasurer is not IDocxStaticTextMetricsProvider staticMetrics)
        {
            return pages;
        }

        var pagesWithStaticText = new DocxLayoutPage[pages.Count];
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++)
        {
            DocxLayoutPage page = pages[pageIndex];
            int pageNumber = pageIndex + 1;
            double bodyWidth = Math.Max(1d, page.Width - page.MarginLeft - page.MarginRight);
            DocxTextLineLayout[] staticLines = CreateStaticTextLines(
                    SelectStaticHeaderFooter(page.PageSettings.HeaderParagraphsByType, page.PageSettings, pageNumber),
                    page.MarginLeft,
                    bodyWidth,
                    page.Height - ResolveHeaderDistance(page),
                    true,
                    pageNumber,
                    pages.Count,
                    textMeasurer,
                    staticMetrics)
                .Concat(CreateStaticTextLines(
                    SelectStaticHeaderFooter(page.PageSettings.FooterParagraphsByType, page.PageSettings, pageNumber),
                    page.MarginLeft,
                    bodyWidth,
                    ResolveFooterDistance(page),
                    false,
                    pageNumber,
                    pages.Count,
                    textMeasurer,
                    staticMetrics))
                .ToArray();
            pagesWithStaticText[pageIndex] = page with { StaticTextLines = staticLines };
        }

        return pagesWithStaticText;
    }

    private static IReadOnlyList<DocxTextLineLayout> CreateStaticTextLines(
        IReadOnlyList<DocxParagraph> paragraphs,
        double x,
        double width,
        double startY,
        bool isHeader,
        int pageNumber,
        int pageCount,
        IDocxTextMeasurer textMeasurer,
        IDocxStaticTextMetricsProvider staticMetrics)
    {
        var lines = new List<DocxTextLineLayout>();
        double cursorY = startY;
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        foreach (DocxParagraph paragraph in paragraphs)
        {
            double appliedBeforeSpacing = ShouldSuppressContextualSpacing(previousParagraph, paragraph)
                ? 0d
                : Math.Max(pendingSpacingAfter, paragraph.SpacingBeforePoints);
            cursorY -= appliedBeforeSpacing;
            pendingSpacingAfter = 0d;
            DocxTextSpan[] spans = CreateStaticTextSpans(paragraph.Runs, pageNumber, pageCount);
            if (spans.Length == 0)
            {
                previousParagraph = paragraph;
                continue;
            }

            foreach (DocxWrappedTextLine line in WrapStaticTextLines(spans, width, textMeasurer))
            {
                if (line.Spans.Count == 0)
                {
                    continue;
                }

                double lineWidth = MeasureStaticTextSpans(line.Spans, textMeasurer);
                double lineX = paragraph.Alignment switch
                {
                    DocxTextAlignment.Center => x + Math.Max(0d, width - lineWidth) / 2d,
                    DocxTextAlignment.Right => x + Math.Max(0d, width - lineWidth),
                    _ => x
                };
                double ascender = line.Spans.Max(span => staticMetrics.MeasureWindowsAscender(span.StyleRun, span.StyleRun.FontSize));
                double descender = line.Spans.Max(span => staticMetrics.MeasureWindowsDescender(span.StyleRun, span.StyleRun.FontSize));
                double baselineY = isHeader ? cursorY - ascender : cursorY + descender;
                IReadOnlyList<DocxTextSegmentLayout> segments = CreateStaticTextSegments(line.Spans, lineX, textMeasurer);
                lines.Add(new DocxTextLineLayout(
                    line.Text,
                    line.Spans[0].StyleRun,
                    line.Spans.Max(span => span.StyleRun.FontSize),
                    lineX,
                    baselineY,
                    lineWidth,
                    segments,
                    LineHeight: ascender + descender,
                    AppliedBeforeSpacing: appliedBeforeSpacing,
                    IsFirstParagraphLine: true));
                appliedBeforeSpacing = 0d;
                cursorY -= ascender + descender;
            }

            pendingSpacingAfter = paragraph.SpacingAfterPoints;
            previousParagraph = paragraph;
        }

        return lines;
    }

    private static DocxTextSpan[] CreateStaticTextSpans(IReadOnlyList<DocxTextRun> runs, int pageNumber, int pageCount)
    {
        if (runs.Count != 0 && runs.All(run => run.Text.Length == 0 || run.Hidden))
        {
            DocxTextRun? paragraphMark = runs.FirstOrDefault(run => !run.Hidden);
            return paragraphMark is null ? [] : [new DocxTextSpan(" ", paragraphMark)];
        }

        return runs
            .Where(run => !run.Hidden)
            .Select(run => new DocxTextSpan(ResolveStaticFieldPlaceholders(run.Text, pageNumber, pageCount), run))
            .Where(span => span.Text.Length != 0)
            .ToArray();
    }

    private static IEnumerable<DocxWrappedTextLine> WrapStaticTextLines(
        IReadOnlyList<DocxTextSpan> spans,
        double maxWidth,
        IDocxTextMeasurer textMeasurer)
    {
        string text = string.Concat(spans.Select(span => span.Text));
        int segmentStart = 0;
        while (segmentStart <= text.Length)
        {
            int breakIndex = text.IndexOf('\n', segmentStart);
            int segmentLength = breakIndex < 0 ? text.Length - segmentStart : breakIndex - segmentStart;
            bool yielded = false;
            foreach (DocxWrappedTextLine line in WrapStaticWords(text, spans, segmentStart, segmentLength, maxWidth, textMeasurer))
            {
                yielded = true;
                yield return line;
            }

            if (!yielded && segmentLength == 0)
            {
                yield return new DocxWrappedTextLine(string.Empty, []);
            }

            if (breakIndex < 0)
            {
                yield break;
            }

            segmentStart = breakIndex + 1;
        }
    }

    private static IEnumerable<DocxWrappedTextLine> WrapStaticWords(
        string text,
        IReadOnlyList<DocxTextSpan> spans,
        int segmentStart,
        int segmentLength,
        double maxWidth,
        IDocxTextMeasurer textMeasurer)
    {
        IReadOnlyList<TextToken> tokens = TokenizeSpaces(text, segmentStart, segmentLength);
        if (tokens.Count == 0)
        {
            yield break;
        }

        int lineStart = tokens[0].Start;
        int lineLength = 0;
        foreach (TextToken token in tokens)
        {
            int candidateLength = token.Start + token.Length - lineStart;
            bool lineHasNonWhitespace = HasNonWhitespace(text, lineStart, lineLength);
            if (lineLength > 0 &&
                lineHasNonWhitespace &&
                !token.IsBreakableWhitespace &&
                MeasureStaticTextSpans(SliceTextSpans(spans, lineStart, candidateLength), textMeasurer) > maxWidth)
            {
                yield return CreateWrappedTextLine(text, spans, lineStart, lineLength);
                lineStart = token.Start;
                lineLength = token.Length;
            }
            else
            {
                lineLength = candidateLength;
            }
        }

        if (lineLength > 0)
        {
            yield return CreateWrappedTextLine(text, spans, lineStart, lineLength);
        }
    }

    private static IReadOnlyList<DocxTextSegmentLayout> CreateStaticTextSegments(
        IReadOnlyList<DocxTextSpan> spans,
        double lineX,
        IDocxTextMeasurer textMeasurer)
    {
        var segments = new List<DocxTextSegmentLayout>(spans.Count);
        double segmentX = lineX;
        for (int i = 0; i < spans.Count; i++)
        {
            DocxTextSpan span = spans[i];
            double nominalFontSize = span.StyleRun.FontSize;
            double layoutFontSize = DocxVerticalAlignMetrics.ResolveFontSize(nominalFontSize, span.StyleRun);
            double baselineOffset = DocxVerticalAlignMetrics.ResolveBaselineOffset(nominalFontSize, layoutFontSize, span.StyleRun);
            double width = textMeasurer.MeasureText(span.StyleRun, span.Text, layoutFontSize);
            segments.Add(new DocxTextSegmentLayout(span.Text, span.StyleRun, segmentX, width, layoutFontSize, baselineOffset));
            segmentX += width;
            if (i + 1 < spans.Count)
            {
                segmentX += DocxTextSpacing.BoundarySpacing(span.StyleRun, span.Text, spans[i + 1].Text);
            }
        }

        return segments;
    }

    private static double MeasureStaticTextSpans(IReadOnlyList<DocxTextSpan> spans, IDocxTextMeasurer textMeasurer)
    {
        double width = 0d;
        for (int i = 0; i < spans.Count; i++)
        {
            DocxTextSpan span = spans[i];
            double layoutFontSize = DocxVerticalAlignMetrics.ResolveFontSize(span.StyleRun.FontSize, span.StyleRun);
            width += textMeasurer.MeasureText(span.StyleRun, span.Text, layoutFontSize);
            if (i + 1 < spans.Count)
            {
                width += DocxTextSpacing.BoundarySpacing(span.StyleRun, span.Text, spans[i + 1].Text);
            }
        }

        return width;
    }

    private static IReadOnlyList<DocxParagraph> SelectStaticHeaderFooter(
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> paragraphsByType,
        DocxPageSettings settings,
        int pageNumber)
    {
        if (settings.TitlePage == true &&
            pageNumber == 1 &&
            paragraphsByType.TryGetValue("first", out IReadOnlyList<DocxParagraph>? first))
        {
            return first;
        }

        if (settings.EvenAndOddHeaders == true &&
            pageNumber % 2 == 0 &&
            paragraphsByType.TryGetValue("even", out IReadOnlyList<DocxParagraph>? even))
        {
            return even;
        }

        return paragraphsByType.TryGetValue("default", out IReadOnlyList<DocxParagraph>? defaults)
            ? defaults
            : [];
    }

    private static double ResolveHeaderDistance(DocxLayoutPage page)
    {
        return page.PageSettings.HeaderDistancePoints ?? Math.Max(18d, page.MarginTop / 2d);
    }

    private static double ResolveFooterDistance(DocxLayoutPage page)
    {
        return page.PageSettings.FooterDistancePoints ?? Math.Max(18d, page.MarginBottom / 2d);
    }

    private static string ResolveStaticFieldPlaceholders(string text, int pageNumber, int pageCount)
    {
        return text
            .Replace("{NUMPAGES}", pageCount.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal)
            .Replace("{PAGE}", pageNumber.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal);
    }

    private static IReadOnlyDictionary<int, DocxPageSettings> BuildEffectiveSectionSettings(DocxDocument document, out DocxPageSettings finalSectionSettings)
    {
        var sectionSettingsByElementIndex = new Dictionary<int, DocxPageSettings>();
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> inheritedHeadersByType =
            new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase);
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> inheritedFootersByType =
            new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase);

        for (int elementIndex = 0; elementIndex < document.BodyElements.Count; elementIndex++)
        {
            if (document.BodyElements[elementIndex] is not DocxSectionBreakElement sectionBreak)
            {
                continue;
            }

            DocxPageSettings effectiveSettings = ResolveEffectiveSectionSettings(
                sectionBreak.PageSettings,
                inheritedHeadersByType,
                inheritedFootersByType);
            sectionSettingsByElementIndex[elementIndex] = effectiveSettings;
            inheritedHeadersByType = effectiveSettings.HeaderParagraphsByType;
            inheritedFootersByType = effectiveSettings.FooterParagraphsByType;
        }

        finalSectionSettings = ResolveEffectiveSectionSettings(
            BuildFinalSectionSettings(document),
            inheritedHeadersByType,
            inheritedFootersByType);
        return sectionSettingsByElementIndex;
    }

    private static DocxPageSettings BuildFinalSectionSettings(DocxDocument document)
    {
        DocxPageSettings settings = document.PageSettings;
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> headersByType = settings.HeaderParagraphsByType.Count == 0 && document.HeaderParagraphsByType.Count > 0
            ? document.HeaderParagraphsByType
            : settings.HeaderParagraphsByType;
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> footersByType = settings.FooterParagraphsByType.Count == 0 && document.FooterParagraphsByType.Count > 0
            ? document.FooterParagraphsByType
            : settings.FooterParagraphsByType;

        if (headersByType.Count == 0 && document.HeaderParagraphs.Count > 0)
        {
            headersByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = document.HeaderParagraphs
            };
        }

        if (footersByType.Count == 0 && document.FooterParagraphs.Count > 0)
        {
            footersByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = document.FooterParagraphs
            };
        }

        return settings with
        {
            HeaderParagraphsByType = headersByType,
            FooterParagraphsByType = footersByType
        };
    }

    private static DocxPageSettings ResolveEffectiveSectionSettings(
        DocxPageSettings settings,
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> inheritedHeadersByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> inheritedFootersByType)
    {
        return settings with
        {
            HeaderParagraphsByType = MergeInheritedStaticParagraphs(inheritedHeadersByType, settings.HeaderParagraphsByType),
            FooterParagraphsByType = MergeInheritedStaticParagraphs(inheritedFootersByType, settings.FooterParagraphsByType)
        };
    }

    private static IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> MergeInheritedStaticParagraphs(
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> inheritedParagraphsByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> localParagraphsByType)
    {
        if (inheritedParagraphsByType.Count == 0)
        {
            return localParagraphsByType;
        }

        if (localParagraphsByType.Count == 0)
        {
            return inheritedParagraphsByType;
        }

        var merged = new Dictionary<string, IReadOnlyList<DocxParagraph>>(inheritedParagraphsByType, StringComparer.OrdinalIgnoreCase);
        foreach ((string type, IReadOnlyList<DocxParagraph> paragraphs) in localParagraphsByType)
        {
            merged[type] = paragraphs;
        }

        return merged;
    }

    private static DocxPageSettings? FindSectionSettingsAtOrAfter(
        IReadOnlyList<DocxBodyElement> elements,
        int startIndex,
        IReadOnlyDictionary<int, DocxPageSettings> sectionSettingsByElementIndex)
    {
        for (int i = Math.Max(0, startIndex); i < elements.Count; i++)
        {
            if (elements[i] is DocxSectionBreakElement && sectionSettingsByElementIndex.TryGetValue(i, out DocxPageSettings? settings))
            {
                return settings;
            }
        }

        return null;
    }

    private static DocxPageGeometry ResolveSectionGeometry(DocxDocument document, DocxPageSettings? settings)
    {
        DocxPageSettings effectiveSettings = settings ?? document.PageSettings;
        double width = ReadTwipsValue(effectiveSettings.WidthValue, document.PageWidthPoints);
        double height = ReadTwipsValue(effectiveSettings.HeightValue, document.PageHeightPoints);
        (width, height) = NormalizePageSize(width, height);
        if (effectiveSettings.OrientationValue?.Equals("landscape", StringComparison.OrdinalIgnoreCase) == true && height > width)
        {
            (width, height) = (height, width);
        }

        return new DocxPageGeometry(
            width,
            height,
            ReadTwipsValue(effectiveSettings.MarginLeftValue, document.MarginLeftPoints),
            ReadTwipsValue(effectiveSettings.MarginRightValue, document.MarginRightPoints),
            ReadTwipsValue(effectiveSettings.MarginTopValue, document.MarginTopPoints),
            ReadTwipsValue(effectiveSettings.MarginBottomValue, document.MarginBottomPoints),
            effectiveSettings);
    }

    private static double ReadTwipsValue(string? value, double fallback)
    {
        return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long twips)
            ? OoxUnits.TwipsToPoints(twips)
            : fallback;
    }

    private static (double Width, double Height) NormalizePageSize(double width, double height)
    {
        if (Math.Abs(width - 595d) < 0.01d && Math.Abs(height - 842d) < 0.01d)
        {
            return (594.96d, 842.04d);
        }

        return (width, height);
    }

    private static bool ShouldStartNewPageForSectionBreak(DocxSectionBreakElement sectionBreak)
    {
        return sectionBreak.TypeValue is null ||
            sectionBreak.TypeValue.Equals("nextPage", StringComparison.OrdinalIgnoreCase) ||
            sectionBreak.TypeValue.Equals("oddPage", StringComparison.OrdinalIgnoreCase) ||
            sectionBreak.TypeValue.Equals("evenPage", StringComparison.OrdinalIgnoreCase);
    }

    private static bool ShouldKeepParagraphBlockTogether(DocxParagraph paragraph)
    {
        return paragraph.KeepRules.KeepLines == true || paragraph.KeepRules.KeepNext == true;
    }

    private static double ResolveLineHeight(DocxParagraph paragraph, double fontSize, IDocxTextMeasurer? textMeasurer)
    {
        return ResolveLineHeightProfile(paragraph, fontSize, textMeasurer).LineHeight;
    }

    private static DocxLineHeightProfile ResolveLineHeightProfile(DocxParagraph paragraph, double fontSize, IDocxTextMeasurer? textMeasurer)
    {
        if (paragraph.LineSpacingPoints is { } exactLineHeight)
        {
            return new DocxLineHeightProfile(
                exactLineHeight,
                SingleLineHeight: null,
                ListLabelSingleLineHeight: null,
                BodyWindowsLineHeight: null,
                ListLabelWindowsLineHeight: null,
                EffectiveLineSpacingFactor: null,
                LineSpacingFactorFloorApplied: false);
        }

        DocxTextRun? bodyRun = paragraph.Runs.FirstOrDefault();
        DocxTextRun? listLabelRun = paragraph.ListLabel is null
            ? null
            : CreateListLabelRun(paragraph.ListLabel, bodyRun, fontSize);
        IDocxLineMetricsProvider? metricsProvider = textMeasurer as IDocxLineMetricsProvider;
        IDocxStaticTextMetricsProvider? staticMetrics = textMeasurer as IDocxStaticTextMetricsProvider;
        double singleLineHeight = metricsProvider is not null
            ? metricsProvider.MeasureSingleLineHeight(bodyRun, fontSize)
            : fontSize;
        double? listLabelSingleLineHeight = metricsProvider is not null && listLabelRun is not null
            ? metricsProvider.MeasureSingleLineHeight(listLabelRun, listLabelRun.FontSize)
            : null;
        double? bodyWindowsLineHeight = staticMetrics is not null
            ? staticMetrics.MeasureWindowsAscender(bodyRun, fontSize) + staticMetrics.MeasureWindowsDescender(bodyRun, fontSize)
            : null;
        double? listLabelWindowsLineHeight = staticMetrics is not null && listLabelRun is not null
            ? staticMetrics.MeasureWindowsAscender(listLabelRun, listLabelRun.FontSize) + staticMetrics.MeasureWindowsDescender(listLabelRun, listLabelRun.FontSize)
            : null;
        double effectiveLineSpacingFactor = ResolveAutoLineSpacingFactor(paragraph, out bool floorApplied);
        return new DocxLineHeightProfile(
            singleLineHeight * effectiveLineSpacingFactor,
            singleLineHeight,
            listLabelSingleLineHeight,
            bodyWindowsLineHeight,
            listLabelWindowsLineHeight,
            effectiveLineSpacingFactor,
            floorApplied);
    }

    private static double ResolveAutoLineSpacingFactor(DocxParagraph paragraph, out bool floorApplied)
    {
        if (paragraph.ListLabel is not null &&
            paragraph.LineSpacingPoints is null &&
            paragraph.SpacingBeforePoints > 0d &&
            paragraph.Spacing.LineRuleValue?.Equals("auto", StringComparison.OrdinalIgnoreCase) == true)
        {
            double effectiveFactor = Math.Max(paragraph.LineSpacingFactor, WordListMinimumAutoLineSpacingFactor);
            floorApplied = effectiveFactor > paragraph.LineSpacingFactor;
            return effectiveFactor;
        }

        floorApplied = false;
        return paragraph.LineSpacingFactor;
    }

    private static bool ShouldMoveParagraphForWidowControl(
        DocxParagraph paragraph,
        int lineCount,
        double cursorY,
        double lineHeight,
        double marginBottom,
        bool hasPageContent)
    {
        if (paragraph.KeepRules.WidowControl == false ||
            lineCount <= 1 ||
            !hasPageContent)
        {
            return false;
        }

        int fittingLineCount = (int)Math.Floor(Math.Max(0d, cursorY - marginBottom) / lineHeight);
        return fittingLineCount > 0 &&
            fittingLineCount < lineCount &&
            (fittingLineCount == 1 || lineCount - fittingLineCount == 1);
    }

    private static bool ShouldSuppressContextualSpacing(DocxParagraph? previousParagraph, DocxParagraph paragraph)
    {
        return paragraph.Spacing.ContextualSpacing == true &&
            previousParagraph?.StyleId is not null &&
            paragraph.StyleId is not null &&
            string.Equals(previousParagraph.StyleId, paragraph.StyleId, StringComparison.Ordinal);
    }

    private static DocxKeepBlockEstimate EstimateKeptParagraphBlock(
        IReadOnlyList<DocxBodyElement> elements,
        int elementIndex,
        double availableWidth,
        IDocxTextMeasurer textMeasurer,
        double defaultTabStopPoints)
    {
        if (elements[elementIndex] is not DocxParagraphElement paragraphElement)
        {
            return new DocxKeepBlockEstimate(0d, 0, 0);
        }

        DocxParagraph paragraph = paragraphElement.Paragraph;
        double height = EstimateParagraphContentHeight(paragraph, availableWidth, textMeasurer, defaultTabStopPoints);
        int paragraphCount = 1;
        int firstTableRowCount = 0;
        int nextSearchIndex = elementIndex + 1;
        while (paragraph.KeepRules.KeepNext == true &&
            TryFindNextKeepTarget(elements, nextSearchIndex, out int nextIndex, out DocxBodyElement? next))
        {
            if (next is DocxParagraphElement nextParagraph)
            {
                height += Math.Max(paragraph.SpacingAfterPoints, nextParagraph.Paragraph.SpacingBeforePoints);
                height += EstimateParagraphContentHeight(nextParagraph.Paragraph, availableWidth, textMeasurer, defaultTabStopPoints);
                paragraphCount++;
                paragraph = nextParagraph.Paragraph;
                nextSearchIndex = nextIndex + 1;
                continue;
            }

            if (next is DocxTableElement nextTable)
            {
                height += paragraph.SpacingAfterPoints;
                height += EstimateFirstTableRowHeight(nextTable.Table, availableWidth, textMeasurer, defaultTabStopPoints);
                firstTableRowCount++;
            }

            break;
        }

        return new DocxKeepBlockEstimate(height, paragraphCount, firstTableRowCount);
    }

    private static bool TryFindNextKeepTarget(IReadOnlyList<DocxBodyElement> elements, int startIndex, out DocxBodyElement? target)
    {
        bool found = TryFindNextKeepTarget(elements, startIndex, out _, out DocxBodyElement? indexedTarget);
        target = indexedTarget;
        return found;
    }

    private static bool TryFindNextKeepTarget(IReadOnlyList<DocxBodyElement> elements, int startIndex, out int targetIndex, out DocxBodyElement? target)
    {
        for (int i = startIndex; i < elements.Count; i++)
        {
            if (elements[i] is DocxParagraphElement or DocxTableElement)
            {
                targetIndex = i;
                target = elements[i];
                return true;
            }

            if (elements[i] is DocxPageBreakElement or DocxManualBreakElement or DocxSectionBreakElement)
            {
                break;
            }
        }

        targetIndex = -1;
        target = null;
        return false;
    }

    private static double EstimateParagraphContentHeight(DocxParagraph paragraph, double availableWidth, IDocxTextMeasurer textMeasurer, double defaultTabStopPoints)
    {
        double height = 0d;
        double fontSize = GetParagraphFontSize(paragraph);
        double lineHeight = ResolveLineHeight(paragraph, fontSize, textMeasurer);
        IReadOnlyList<DocxTextSpan> textSpans = CreateTextSpans(paragraph.Runs);
        if (textSpans.Count != 0)
        {
            double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, textMeasurer);
            double firstParagraphWidth = Math.Max(1d, availableWidth - textStartOffset - GetParagraphRightInset(paragraph));
            double continuationParagraphWidth = Math.Max(1d, availableWidth - GetParagraphTextStartOffset(paragraph) - GetParagraphRightInset(paragraph));
            height += WrapTextLines(textSpans, firstParagraphWidth, continuationParagraphWidth, fontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints, allowOverwideTokenBreaks: false).Count() * lineHeight;
        }
        else if (paragraph.Images.Count == 0)
        {
            height += lineHeight;
        }

        foreach (DocxInlineImage image in paragraph.Images)
        {
            double imageWidth = Math.Min(availableWidth, image.WidthPoints);
            double imageHeight = image.HeightPoints * imageWidth / Math.Max(1d, image.WidthPoints);
            height += imageHeight + 6d;
        }

        return height;
    }

    private static double EstimateFirstTableRowHeight(DocxTable table, double availableWidth, IDocxTextMeasurer textMeasurer, double defaultTabStopPoints)
    {
        DocxTableRow? row = table.Rows.FirstOrDefault();
        if (row is null)
        {
            return 0d;
        }

        double tableAvailableWidth = Math.Max(1d, availableWidth - Math.Max(0d, table.IndentPoints ?? 0d));
        double gridTableWidth = table.ColumnWidthsPoints.Sum();
        double fallbackTableWidth = table.HasExplicitGrid && gridTableWidth > 0d ? gridTableWidth : tableAvailableWidth;
        double targetTableWidth = ResolveTargetTableWidth(table, tableAvailableWidth, fallbackTableWidth);
        IReadOnlyList<double> effectiveColumns = GetEffectiveTableColumnWidths(table, targetTableWidth);
        double rawTableWidth = effectiveColumns.Sum();
        double scale = rawTableWidth <= 0d ? 1d : targetTableWidth / rawTableWidth;
        double[] cellWidths = GetTableRowCellWidths(row, effectiveColumns, scale);
        double rowTopPadding = ResolveTableRowTopPadding(row);
        double contentHeight = row.Cells
            .Select((cell, columnIndex) => MeasureTableCellContentHeight(cell, cellWidths[columnIndex], textMeasurer, defaultTabStopPoints, rowTopPadding))
            .DefaultIfEmpty(0d)
            .Max();
        return ResolveTableRowHeight(row, contentHeight);
    }

    private static double GetParagraphTextStartOffset(DocxParagraph paragraph)
    {
        if (paragraph.ListLabel is null)
        {
            return Math.Max(0d, paragraph.Indent.LeftPoints ?? 0d);
        }

        DocxNumberingIndent indent = paragraph.ListLabel.Indent;
        double left = indent.LeftPoints ?? 0d;
        double firstLine = indent.FirstLinePoints ?? 0d;
        return Math.Max(0d, left + firstLine);
    }

    private static double GetParagraphFirstLineTextStartOffset(DocxParagraph paragraph, double fontSize, IDocxTextMeasurer textMeasurer)
    {
        if (paragraph.ListLabel is null)
        {
            return GetParagraphFirstLineIndentOffset(paragraph);
        }

        if (IsNumberingTabSuffix(paragraph.ListLabel))
        {
            return GetParagraphTextStartOffset(paragraph);
        }

        double gap = IsNumberingSpaceSuffix(paragraph.ListLabel)
            ? textMeasurer.MeasureText(paragraph.Runs.FirstOrDefault(), " ", fontSize)
            : 0d;
        DocxTextRun labelRun = CreateListLabelRun(paragraph.ListLabel, paragraph.Runs.FirstOrDefault(), fontSize);
        return Math.Max(
            0d,
            GetParagraphLabelStartOffset(paragraph) + textMeasurer.MeasureText(labelRun, paragraph.ListLabel.Text, labelRun.FontSize) + gap);
    }

    private static double GetParagraphLabelStartOffset(DocxParagraph paragraph)
    {
        if (paragraph.ListLabel is null)
        {
            return 0d;
        }

        DocxNumberingIndent indent = paragraph.ListLabel.Indent;
        double left = indent.LeftPoints ?? 0d;
        double hanging = indent.HangingPoints ?? 0d;
        double firstLine = indent.FirstLinePoints ?? 0d;
        return Math.Max(0d, left - hanging + firstLine);
    }

    private static double GetParagraphStartOffset(DocxParagraph paragraph)
    {
        return paragraph.ListLabel is null
            ? GetParagraphFirstLineIndentOffset(paragraph)
            : GetParagraphLabelStartOffset(paragraph);
    }

    private static double GetParagraphRightInset(DocxParagraph paragraph)
    {
        return paragraph.ListLabel?.Indent.RightPoints ?? paragraph.Indent.RightPoints ?? 0d;
    }

    private static double GetParagraphFirstLineIndentOffset(DocxParagraph paragraph)
    {
        double left = paragraph.Indent.LeftPoints ?? 0d;
        double firstLine = paragraph.Indent.FirstLinePoints ?? 0d;
        double hanging = paragraph.Indent.HangingPoints ?? 0d;
        return Math.Max(0d, left + firstLine - hanging);
    }

    private static IReadOnlyList<DocxTextSegmentLayout> CreateNumberedLineSegments(
        DocxListLabel label,
        IReadOnlyList<DocxTextSpan> lineSpans,
        DocxTextRun styleRun,
        double labelX,
        double lineX,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints)
    {
        DocxTextRun labelRun = CreateListLabelRun(label, styleRun, fontSize);
        double labelWidth = textMeasurer.MeasureText(labelRun, label.Text, labelRun.FontSize);
        double pdfCharacterSpacing = OfficeNumberedTextStateCharacterSpacing(fontSize);
        var segments = new List<DocxTextSegmentLayout>
        {
            new(
                label.Text,
                labelRun,
                labelX,
                labelWidth,
                labelRun.FontSize,
                PdfCharacterSpacing: pdfCharacterSpacing,
                CompensatePdfCharacterSpacing: false)
        };

        string separator = GetListLabelPdfSeparator(label);
        if (separator.Length != 0)
        {
            double separatorX = labelX + labelWidth;
            double separatorWidth = textMeasurer.MeasureText(labelRun, separator, labelRun.FontSize);
            segments.Add(new DocxTextSegmentLayout(separator, labelRun, separatorX, separatorWidth, labelRun.FontSize));
        }

        segments.AddRange(CreateTextSegments(lineSpans, lineX, fontSize, textMeasurer, tabStops, defaultTabStopPoints));
        return segments;
    }

    private static double OfficeNumberedTextStateCharacterSpacing(double fontSize)
    {
        const double docxNumberedTextStateSpacingEm = 0.004d;
        return fontSize * docxNumberedTextStateSpacingEm;
    }

    private static double MeasureListLabel(DocxListLabel label, DocxTextRun? baseRun, double fontSize, IDocxTextMeasurer textMeasurer)
    {
        DocxTextRun labelRun = CreateListLabelRun(label, baseRun, fontSize);
        return textMeasurer.MeasureText(labelRun, label.Text, labelRun.FontSize);
    }

    internal static DocxTextRun CreateListLabelRun(DocxListLabel label, DocxTextRun? baseRun, double fontSize)
    {
        return label.Style.ApplyTo(baseRun, label.Text, fontSize);
    }

    private static string GetListLabelTextSeparator(DocxListLabel label)
    {
        return label.SuffixValue switch
        {
            "nothing" => string.Empty,
            "space" => " ",
            _ => "\t"
        };
    }

    private static string GetListLabelPdfSeparator(DocxListLabel label)
    {
        return label.SuffixValue.Equals("nothing", StringComparison.OrdinalIgnoreCase) ? string.Empty : " ";
    }

    private static bool IsNumberingTabSuffix(DocxListLabel label)
    {
        return string.IsNullOrEmpty(label.SuffixValue) ||
            label.SuffixValue.Equals("tab", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsNumberingSpaceSuffix(DocxListLabel label)
    {
        return label.SuffixValue.Equals("space", StringComparison.OrdinalIgnoreCase);
    }

    private static void LayoutTable(
        DocxTable table,
        int tableIndex,
        int sourceBlockIndex,
        double marginBottom,
        double pageContentHeight,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        Func<int> getPageIndex,
        ref List<DocxLayoutItem> currentItems,
        ref double cursorY,
        double x,
        double availableWidth,
        Action finishPage,
        Func<bool> hasPageContent)
    {
        double tableX = x + Math.Max(0d, table.IndentPoints ?? 0d);
        double tableAvailableWidth = Math.Max(1d, availableWidth - Math.Max(0d, table.IndentPoints ?? 0d));
        double gridTableWidth = table.ColumnWidthsPoints.Sum();
        double fallbackTableWidth = table.HasExplicitGrid && gridTableWidth > 0d ? gridTableWidth : tableAvailableWidth;
        double targetTableWidth = ResolveTargetTableWidth(table, tableAvailableWidth, fallbackTableWidth);
        IReadOnlyList<double> effectiveColumns = GetEffectiveTableColumnWidths(table, targetTableWidth);
        double rawTableWidth = effectiveColumns.Sum();
        double scale = rawTableWidth <= 0d ? 1d : targetTableWidth / rawTableWidth;
        var tableContext = new DocxTableLayoutContext(
            tableIndex,
            sourceBlockIndex,
            table.Rows.Count,
            table.ColumnWidthsPoints.Count,
            table.ColumnWidthsPoints.Sum(),
            table.HasExplicitGrid,
            effectiveColumns.Select(width => width * scale).ToArray(),
            targetTableWidth,
            tableX,
            table.PreferredWidthPoints,
            table.PreferredWidthValue,
            table.PreferredWidthType,
            table.IndentPoints,
            table.CellSpacingPoints,
            table.LayoutValue);
        double[] rowHeights = table.Rows
            .Select(row => MeasureTableRowHeight(table, row, effectiveColumns, scale, textMeasurer, defaultTabStopPoints))
            .ToArray();

        IReadOnlyList<(DocxTableRow Row, int RowIndex)> headerRows = table.Rows
            .Select((row, rowIndex) => (row, rowIndex))
            .TakeWhile(entry => entry.row.IsHeader)
            .Select(entry => (entry.row, entry.rowIndex))
            .ToArray();
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            DocxTableRow row = table.Rows[rowIndex];
            double rowHeight = rowHeights[rowIndex];
            double remainingPageHeight = Math.Max(0d, cursorY - marginBottom);
            if (!row.CantSplit &&
                rowHeight > remainingPageHeight &&
                remainingPageHeight > 0.001d &&
                CanSplitTableRowAtPageBoundary(row, effectiveColumns, scale, rowHeight, remainingPageHeight, textMeasurer, defaultTabStopPoints))
            {
                AddSplitTableRowLayout(table, tableContext, row, rowIndex, rowHeights, headerRows, effectiveColumns, scale, textMeasurer, defaultTabStopPoints, getPageIndex, ref currentItems, ref cursorY, tableX, remainingPageHeight, pageContentHeight, finishPage);
                continue;
            }

            if (cursorY - rowHeight < marginBottom && hasPageContent())
            {
                finishPage();
                if (!row.IsHeader)
                {
                    AddRepeatedTableHeaderRows(table, tableContext, rowHeights, headerRows, effectiveColumns, scale, textMeasurer, defaultTabStopPoints, getPageIndex, ref currentItems, ref cursorY, tableX);
                }
            }

            AddTableRowLayout(table, tableContext, row, rowIndex, rowHeights, effectiveColumns, scale, textMeasurer, defaultTabStopPoints, getPageIndex, ref currentItems, ref cursorY, tableX);
        }
    }

    private static IReadOnlyList<double> GetEffectiveTableColumnWidths(DocxTable table, double preferredTableWidth)
    {
        int columnCount = table.ColumnWidthsPoints.Count;
        if (columnCount == 0)
        {
            int inferredColumnCount = GetMaxGridColumnCount(table);
            return inferredColumnCount > 0
                ? Enumerable.Repeat(preferredTableWidth / inferredColumnCount, inferredColumnCount).ToArray()
                : table.ColumnWidthsPoints;
        }

        double?[] preferredWidths = new double?[columnCount];
        foreach (DocxTableRow row in table.Rows)
        {
            int gridColumnIndex = 0;
            foreach (DocxTableCell cell in row.Cells)
            {
                int span = Math.Max(1, cell.GridSpan);
                double? preferredWidth = ResolvePreferredCellWidth(cell, preferredTableWidth);
                if (span == 1 &&
                    gridColumnIndex < columnCount &&
                    preferredWidth is > 0d)
                {
                    preferredWidths[gridColumnIndex] = preferredWidth.Value;
                }

                gridColumnIndex += span;
            }

            if (preferredWidths.All(width => width is > 0d))
            {
                return preferredWidths.Select(width => width!.Value).ToArray();
            }
        }

        if (!table.HasExplicitGrid)
        {
            int inferredColumnCount = columnCount == 0 ? GetMaxGridColumnCount(table) : columnCount;
            if (inferredColumnCount > 0)
            {
                return Enumerable.Repeat(preferredTableWidth / inferredColumnCount, inferredColumnCount).ToArray();
            }
        }

        return table.ColumnWidthsPoints;
    }

    private static int GetMaxGridColumnCount(DocxTable table)
    {
        return table.Rows
            .Select(row => row.Cells.Sum(cell => Math.Max(1, cell.GridSpan)))
            .DefaultIfEmpty(0)
            .Max();
    }

    private static double? ResolvePreferredCellWidth(DocxTableCell cell, double preferredTableWidth)
    {
        if (cell.PreferredWidthPoints is { } points)
        {
            return points;
        }

        if (cell.PreferredWidthType?.Equals("pct", StringComparison.OrdinalIgnoreCase) == true &&
            int.TryParse(cell.PreferredWidthValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int fiftiethsPercent))
        {
            return Math.Max(0d, preferredTableWidth * fiftiethsPercent / 5000d);
        }

        return null;
    }

    private static double? ResolvePreferredTableWidth(DocxTable table, double availableWidth)
    {
        if (table.PreferredWidthPoints is { } points)
        {
            return points;
        }

        if (table.PreferredWidthType?.Equals("pct", StringComparison.OrdinalIgnoreCase) == true &&
            int.TryParse(table.PreferredWidthValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int fiftiethsPercent))
        {
            double normalPercentageWidth = availableWidth * fiftiethsPercent / 5000d;
            double explicitGridWidth = table.HasExplicitGrid ? table.ColumnWidthsPoints.Sum() : 0d;
            double percentageBasis = explicitGridWidth > 0d && explicitGridWidth < normalPercentageWidth - 0.001d
                ? availableWidth + ResolveOuterTableCellContentInset(table)
                : availableWidth;
            return Math.Max(0d, percentageBasis * fiftiethsPercent / 5000d);
        }

        return null;
    }

    private static double ResolveOuterTableCellContentInset(DocxTable table)
    {
        DocxTableRow? firstRow = table.Rows.FirstOrDefault();
        if (firstRow is null || firstRow.Cells.Count == 0)
        {
            return 0d;
        }

        DocxTableCell firstCell = firstRow.Cells[0];
        DocxTableCell lastCell = firstRow.Cells[^1];
        return ResolveTableCellHorizontalPadding(firstCell.Margins.LeftPoints) +
            ResolveTableCellBorderContentInset(firstCell, "left") +
            ResolveTableCellHorizontalPadding(lastCell.Margins.RightPoints) +
            ResolveTableCellBorderContentInset(lastCell, "right");
    }

    private static double ResolveTargetTableWidth(DocxTable table, double availableWidth, double fallbackWidth)
    {
        double preferredWidth = ResolvePreferredTableWidth(table, availableWidth) ?? fallbackWidth;
        return table.PreferredWidthType?.Equals("dxa", StringComparison.OrdinalIgnoreCase) == true ||
            table.PreferredWidthType?.Equals("pct", StringComparison.OrdinalIgnoreCase) == true
            ? Math.Max(1d, preferredWidth)
            : Math.Min(availableWidth, preferredWidth);
    }

    private static double MeasureTableRowHeight(
        DocxTable table,
        DocxTableRow row,
        IReadOnlyList<double> effectiveColumns,
        double scale,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints)
    {
        double[] cellWidths = GetTableRowCellWidths(row, effectiveColumns, scale);
        double rowTopPadding = ResolveTableRowTopPadding(row);
        double contentHeight = textMeasurer is null
            ? 0d
            : row.Cells
                .Select((cell, columnIndex) => MeasureTableCellContentHeight(cell, cellWidths[columnIndex], textMeasurer, defaultTabStopPoints, rowTopPadding))
                .DefaultIfEmpty(0d)
                .Max();
        return ResolveTableRowHeight(row, contentHeight);
    }

    private static double ResolveTableRowHeight(DocxTableRow row, double contentHeight)
    {
        if (string.Equals(row.HeightRuleValue, "exact", StringComparison.OrdinalIgnoreCase) &&
            row.HeightPoints is { } exactHeight)
        {
            return Math.Max(1d, exactHeight);
        }

        double declaredHeight = string.Equals(row.HeightRuleValue, "auto", StringComparison.OrdinalIgnoreCase)
            ? 0d
            : row.HeightPoints ?? 0d;
        if (row.HeightPoints is not null &&
            !string.Equals(row.HeightRuleValue, "auto", StringComparison.OrdinalIgnoreCase))
        {
            declaredHeight += ResolveTableRowTopPadding(row);
        }

        double height = Math.Max(declaredHeight, contentHeight);
        height += ResolveTableRowCollapsedHorizontalBorderAdvance(row);
        return Math.Max(1d, height);
    }

    private static double ResolveTableRowCollapsedHorizontalBorderAdvance(DocxTableRow row)
    {
        double maxBottom = row.Cells
            .Select(cell => DocxTableBorderGeometry.ResolveVisibleWidth(DocxTableBorderGeometry.Find(cell.Borders, "bottom")))
            .DefaultIfEmpty(0d)
            .Max();
        if (maxBottom > 0d)
        {
            return maxBottom;
        }

        return row.Cells
            .Select(cell => DocxTableBorderGeometry.ResolveVisibleWidth(DocxTableBorderGeometry.Find(cell.Borders, "top")))
            .DefaultIfEmpty(0d)
            .Max();
    }

    private static void AddTableRowLayout(
        DocxTable table,
        DocxTableLayoutContext tableContext,
        DocxTableRow row,
        int rowIndex,
        IReadOnlyList<double> rowHeights,
        IReadOnlyList<double> effectiveColumns,
        double scale,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        Func<int> getPageIndex,
        ref List<DocxLayoutItem> currentItems,
        ref double cursorY,
        double x)
    {
        double rowHeight = rowHeights[rowIndex];
        currentItems.Add(CreateTableRowLayout(
            table,
            tableContext,
            row,
            rowIndex,
            rowHeights,
            effectiveColumns,
            scale,
            textMeasurer,
            defaultTabStopPoints,
            getPageIndex,
            cursorY,
            rowHeight,
            logicalRowTopY: cursorY,
            FragmentIndex: 0,
            FragmentCount: 1));
        cursorY -= rowHeight;
    }

    private static void AddSplitTableRowLayout(
        DocxTable table,
        DocxTableLayoutContext tableContext,
        DocxTableRow row,
        int rowIndex,
        IReadOnlyList<double> rowHeights,
        IReadOnlyList<(DocxTableRow Row, int RowIndex)> headerRows,
        IReadOnlyList<double> effectiveColumns,
        double scale,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        Func<int> getPageIndex,
        ref List<DocxLayoutItem> currentItems,
        ref double cursorY,
        double x,
        double firstFragmentHeight,
        double pageContentHeight,
        Action finishPage)
    {
        double rowHeight = rowHeights[rowIndex];
        double continuationContentHeight = row.IsHeader
            ? pageContentHeight
            : Math.Max(1d, pageContentHeight - SumRepeatedTableHeaderRowsHeight(rowHeights, headerRows));
        IReadOnlyList<double> fragmentHeights = ComputeTableRowFragmentHeights(rowHeight, firstFragmentHeight, continuationContentHeight);
        double consumedHeight = 0d;
        for (int fragmentIndex = 0; fragmentIndex < fragmentHeights.Count; fragmentIndex++)
        {
            double fragmentHeight = fragmentHeights[fragmentIndex];
            currentItems.Add(CreateTableRowLayout(
                table,
                tableContext,
                row,
                rowIndex,
                rowHeights,
                effectiveColumns,
                scale,
                textMeasurer,
                defaultTabStopPoints,
                getPageIndex,
                cursorY,
                fragmentHeight,
                logicalRowTopY: cursorY + consumedHeight,
                FragmentIndex: fragmentIndex,
                FragmentCount: fragmentHeights.Count));
            cursorY -= fragmentHeight;
            consumedHeight += fragmentHeight;

            if (fragmentIndex + 1 < fragmentHeights.Count)
            {
                finishPage();
                if (!row.IsHeader)
                {
                    AddRepeatedTableHeaderRows(table, tableContext, rowHeights, headerRows, effectiveColumns, scale, textMeasurer, defaultTabStopPoints, getPageIndex, ref currentItems, ref cursorY, x);
                }
            }
        }
    }

    private static void AddRepeatedTableHeaderRows(
        DocxTable table,
        DocxTableLayoutContext tableContext,
        IReadOnlyList<double> rowHeights,
        IReadOnlyList<(DocxTableRow Row, int RowIndex)> headerRows,
        IReadOnlyList<double> effectiveColumns,
        double scale,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        Func<int> getPageIndex,
        ref List<DocxLayoutItem> currentItems,
        ref double cursorY,
        double x)
    {
        foreach ((DocxTableRow headerRow, int headerRowIndex) in headerRows)
        {
            AddTableRowLayout(table, tableContext, headerRow, headerRowIndex, rowHeights, effectiveColumns, scale, textMeasurer, defaultTabStopPoints, getPageIndex, ref currentItems, ref cursorY, x);
        }
    }

    private static double SumRepeatedTableHeaderRowsHeight(
        IReadOnlyList<double> rowHeights,
        IReadOnlyList<(DocxTableRow Row, int RowIndex)> headerRows)
    {
        return headerRows.Sum(entry => entry.RowIndex >= 0 && entry.RowIndex < rowHeights.Count ? rowHeights[entry.RowIndex] : 0d);
    }

    private static bool CanSplitTableRowAtPageBoundary(
        DocxTableRow row,
        IReadOnlyList<double> effectiveColumns,
        double scale,
        double rowHeight,
        double firstFragmentHeight,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints)
    {
        if (textMeasurer is null)
        {
            return false;
        }

        double fragmentBottomY = rowHeight - firstFragmentHeight;
        double[] cellWidths = GetTableRowCellWidths(row, effectiveColumns, scale);
        double rowTopPadding = ResolveTableRowTopPadding(row);
        for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
        {
            DocxTableCell cell = row.Cells[cellIndex];
            if (IsVerticalMergeContinuation(cell))
            {
                continue;
            }

            IReadOnlyList<DocxTextLineLayout> textLines = LayoutTableCellTextLines(cell, 0d, 0d, cellWidths[cellIndex], rowHeight, rowTopPadding, textMeasurer, defaultTabStopPoints);
            bool hasLineInFirstFragment = textLines.Any(line => firstFragmentHeight >= line.LineHeight && line.BaselineY >= fragmentBottomY);
            bool hasLineInContinuation = textLines.Any(line => line.BaselineY < fragmentBottomY);
            if (hasLineInFirstFragment && hasLineInContinuation)
            {
                return true;
            }
        }

        return false;
    }

    private static IReadOnlyList<double> ComputeTableRowFragmentHeights(double rowHeight, double firstFragmentHeight, double pageContentHeight)
    {
        var fragments = new List<double>();
        double remainingHeight = rowHeight;
        double firstHeight = Math.Min(remainingHeight, Math.Max(0d, firstFragmentHeight));
        if (firstHeight > 0.001d)
        {
            fragments.Add(firstHeight);
            remainingHeight -= firstHeight;
        }

        double fullPageHeight = Math.Max(1d, pageContentHeight);
        while (remainingHeight > fullPageHeight + 0.001d)
        {
            fragments.Add(fullPageHeight);
            remainingHeight -= fullPageHeight;
        }

        if (remainingHeight > 0.001d)
        {
            fragments.Add(remainingHeight);
        }

        return fragments.Count == 0 ? [Math.Max(1d, rowHeight)] : fragments;
    }

    private static DocxTableRowLayout CreateTableRowLayout(
        DocxTable table,
        DocxTableLayoutContext tableContext,
        DocxTableRow row,
        int rowIndex,
        IReadOnlyList<double> rowHeights,
        IReadOnlyList<double> effectiveColumns,
        double scale,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        Func<int> getPageIndex,
        double cursorY,
        double rowHeight,
        double logicalRowTopY,
        int FragmentIndex,
        int FragmentCount)
    {
        double[] cellWidths = GetTableRowCellWidths(row, effectiveColumns, scale);
        double rowTopPadding = ResolveTableRowTopPadding(row);
        double fullRowHeight = rowHeights[rowIndex];
        double cellX = tableContext.TableX;
        double cellY = cursorY - rowHeight;
        double fullCellY = logicalRowTopY - fullRowHeight;
        var cells = new List<DocxTableCellLayout>(row.Cells.Count);
        int gridColumnIndex = 0;
        for (int columnIndex = 0; columnIndex < row.Cells.Count; columnIndex++)
        {
            double cellWidth = cellWidths[columnIndex];
            DocxTableCell cell = row.Cells[columnIndex];
            bool isVerticalMergeContinuation = IsVerticalMergeContinuation(cell);
            DocxTableCell? verticalMergeOwnerCell = isVerticalMergeContinuation
                ? FindVerticalMergeRestartCell(table, rowIndex, gridColumnIndex)
                : null;
            double visualHeight = rowHeight;
            double visualY = cellY;
            double fullVisualHeight = fullRowHeight;
            double fullVisualY = fullCellY;
            if (IsVerticalMergeRestart(cell))
            {
                fullVisualHeight = GetVerticalMergeSpanHeight(table, rowIndex, gridColumnIndex, rowHeights);
                fullVisualY = cursorY - fullVisualHeight;
                if (FragmentCount == 1)
                {
                    visualHeight = fullVisualHeight;
                    visualY = fullVisualY;
                }
            }

            IReadOnlyList<DocxTextLineLayout> textLines = isVerticalMergeContinuation
                ? []
                : LayoutTableCellTextLines(cell, cellX, fullVisualY, cellWidth, fullVisualHeight, rowTopPadding, textMeasurer, defaultTabStopPoints);
            IReadOnlyList<DocxInlineImageLayout> inlineImages = isVerticalMergeContinuation
                ? []
                : LayoutTableCellInlineImages(cell, cellX, fullVisualY, cellWidth, fullVisualHeight, rowTopPadding, textMeasurer, defaultTabStopPoints, getPageIndex())
                    .Where(image => VerticalOverlap(image.Y, image.Height, visualY, visualHeight) > 0.001d)
                    .ToArray();
            cells.Add(new DocxTableCellLayout(cell, cellX, visualY, cellWidth, visualHeight, textLines, inlineImages, isVerticalMergeContinuation, verticalMergeOwnerCell));
            cellX += cellWidth + (table.CellSpacingPoints ?? 0d);
            gridColumnIndex += Math.Max(1, cell.GridSpan);
        }

        return new DocxTableRowLayout(
            tableContext,
            rowIndex,
            FragmentIndex,
            FragmentCount,
            cells.ToArray(),
            cellY,
            rowHeight,
            row.HeightPoints,
            row.HeightValue,
            row.HeightRuleValue,
            row.IsHeader,
            row.HeaderValue,
            row.TablePropertyExceptionCellMargins is not null,
            row.CantSplit,
            row.CantSplitValue);
    }

    private static bool IsVerticalMergeRestart(DocxTableCell cell)
    {
        return cell.HasVerticalMerge &&
            string.Equals(cell.VerticalMergeValue, "restart", StringComparison.OrdinalIgnoreCase);
    }

    private static double VerticalOverlap(double firstY, double firstHeight, double secondY, double secondHeight)
    {
        return Math.Min(firstY + firstHeight, secondY + secondHeight) - Math.Max(firstY, secondY);
    }

    private static bool IsVerticalMergeContinuation(DocxTableCell cell)
    {
        return cell.HasVerticalMerge && !IsVerticalMergeRestart(cell);
    }

    private static DocxTableCell? FindVerticalMergeRestartCell(
        DocxTable table,
        int rowIndex,
        int gridColumnIndex)
    {
        for (int previousRowIndex = rowIndex - 1; previousRowIndex >= 0; previousRowIndex--)
        {
            if (!TryGetCellAtGridColumn(table.Rows[previousRowIndex], gridColumnIndex, out DocxTableCell? previousCell) ||
                previousCell is null)
            {
                return null;
            }

            if (IsVerticalMergeRestart(previousCell))
            {
                return previousCell;
            }

            if (!IsVerticalMergeContinuation(previousCell))
            {
                return null;
            }
        }

        return null;
    }

    private static double GetVerticalMergeSpanHeight(
        DocxTable table,
        int rowIndex,
        int gridColumnIndex,
        IReadOnlyList<double> rowHeights)
    {
        double height = rowHeights[rowIndex];
        for (int nextRowIndex = rowIndex + 1; nextRowIndex < table.Rows.Count; nextRowIndex++)
        {
            if (!TryGetCellAtGridColumn(table.Rows[nextRowIndex], gridColumnIndex, out DocxTableCell? nextCell) ||
                nextCell is null ||
                !IsVerticalMergeContinuation(nextCell))
            {
                break;
            }

            height += rowHeights[nextRowIndex];
        }

        return height;
    }

    private static bool TryGetCellAtGridColumn(DocxTableRow row, int gridColumnIndex, out DocxTableCell? cell)
    {
        int currentGridColumnIndex = 0;
        foreach (DocxTableCell candidate in row.Cells)
        {
            int span = Math.Max(1, candidate.GridSpan);
            if (gridColumnIndex >= currentGridColumnIndex && gridColumnIndex < currentGridColumnIndex + span)
            {
                cell = candidate;
                return true;
            }

            currentGridColumnIndex += span;
        }

        cell = null;
        return false;
    }

    private static double[] GetTableRowCellWidths(DocxTableRow row, IReadOnlyList<double> effectiveColumns, double scale)
    {
        var widths = new double[row.Cells.Count];
        int gridColumnIndex = 0;
        for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
        {
            DocxTableCell cell = row.Cells[cellIndex];
            int span = Math.Max(1, cell.GridSpan);
            double width = 0d;
            for (int spanIndex = 0; spanIndex < span; spanIndex++)
            {
                width += effectiveColumns[Math.Min(gridColumnIndex + spanIndex, effectiveColumns.Count - 1)] * scale;
            }

            widths[cellIndex] = width;
            gridColumnIndex += span;
        }

        return widths;
    }

    private static double MeasureTableCellContentHeight(
        DocxTableCell cell,
        double cellWidth,
        IDocxTextMeasurer textMeasurer,
        double defaultTabStopPoints,
        double? rowTopPadding = null)
    {
        IReadOnlyList<DocxParagraph> paragraphs = DocxTableCellContent.GetParagraphs(cell);
        if (paragraphs.Count == 0)
        {
            return 0d;
        }

        double paddingLeft = ResolveTableCellHorizontalPadding(cell.Margins.LeftPoints) + ResolveTableCellBorderContentInset(cell, "left");
        double paddingRight = ResolveTableCellHorizontalPadding(cell.Margins.RightPoints) + ResolveTableCellBorderContentInset(cell, "right");
        double paddingTop = rowTopPadding ?? ResolveTableCellVerticalPadding(cell.Margins.TopPoints);
        double paddingBottom = ResolveTableCellVerticalPadding(cell.Margins.BottomPoints);
        double textWidth = Math.Max(1d, cellWidth - paddingLeft - paddingRight);
        double contentHeight = paddingTop + paddingBottom;
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        foreach (DocxParagraph paragraph in paragraphs)
        {
            contentHeight += ShouldSuppressContextualSpacing(previousParagraph, paragraph)
                ? 0d
                : Math.Max(pendingSpacingAfter, paragraph.SpacingBeforePoints);
            pendingSpacingAfter = 0d;
            double fontSize = GetParagraphFontSize(paragraph);
            DocxLineHeightProfile lineHeightProfile = ResolveLineHeightProfile(paragraph, fontSize, textMeasurer);
            double lineHeight = lineHeightProfile.LineHeight;
            IReadOnlyList<DocxTextSpan> textSpans = CreateTextSpans(paragraph.Runs);
            if (textSpans.Count != 0)
            {
                double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, textMeasurer);
                double firstParagraphWidth = Math.Max(1d, textWidth - textStartOffset - GetParagraphRightInset(paragraph));
                double continuationParagraphWidth = Math.Max(1d, textWidth - GetParagraphTextStartOffset(paragraph) - GetParagraphRightInset(paragraph));
                int lineCount = WrapTextLines(textSpans, firstParagraphWidth, continuationParagraphWidth, fontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints, allowOverwideTokenBreaks: true).Count();
                contentHeight += lineCount * lineHeight;
            }
            else if (paragraph.Images.Count == 0)
            {
                contentHeight += lineHeight;
            }

            foreach (DocxInlineImage image in paragraph.Images)
            {
                double imageWidth = Math.Min(textWidth, image.WidthPoints);
                double imageHeight = image.HeightPoints * imageWidth / Math.Max(1d, image.WidthPoints);
                contentHeight += imageHeight + 6d;
            }

            pendingSpacingAfter = paragraph.SpacingAfterPoints;
            previousParagraph = paragraph;
        }

        contentHeight += pendingSpacingAfter;
        return contentHeight;
    }

    private static IReadOnlyList<DocxTextLineLayout> LayoutTableCellTextLines(
        DocxTableCell cell,
        double cellX,
        double cellY,
        double cellWidth,
        double cellHeight,
        double rowTopPadding,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints)
    {
        if (textMeasurer is null)
        {
            return [];
        }

        IReadOnlyList<DocxParagraph> paragraphs = DocxTableCellContent.GetParagraphs(cell);
        if (paragraphs.Count == 0)
        {
            return [];
        }

        double paddingLeft = ResolveTableCellHorizontalPadding(cell.Margins.LeftPoints) + ResolveTableCellBorderContentInset(cell, "left");
        double paddingRight = ResolveTableCellHorizontalPadding(cell.Margins.RightPoints) + ResolveTableCellBorderContentInset(cell, "right");
        double paddingTop = rowTopPadding;
        double paddingBottom = ResolveTableCellVerticalPadding(cell.Margins.BottomPoints);
        double baselineInset = ResolveTableCellFirstBaselineInset(paragraphs);
        double textWidth = Math.Max(1d, cellWidth - paddingLeft - paddingRight);
        double startBaselineY = cellY + cellHeight - baselineInset - paddingTop;
        double cursorY = startBaselineY;
        var lines = new List<DocxTextLineLayout>();
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        foreach (DocxParagraph paragraph in paragraphs)
        {
            double appliedBeforeSpacing = ShouldSuppressContextualSpacing(previousParagraph, paragraph)
                ? 0d
                : Math.Max(pendingSpacingAfter, paragraph.SpacingBeforePoints);
            cursorY -= appliedBeforeSpacing;
            pendingSpacingAfter = 0d;
            double fontSize = GetParagraphFontSize(paragraph);
            DocxLineHeightProfile lineHeightProfile = ResolveLineHeightProfile(paragraph, fontSize, textMeasurer);
            double lineHeight = lineHeightProfile.LineHeight;
            IReadOnlyList<DocxTextSpan> textSpans = CreateTextSpans(paragraph.Runs);
            if (textSpans.Count == 0)
            {
                if (paragraph.Images.Count == 0)
                {
                    cursorY -= lineHeight;
                }
            }
            else
            {
                DocxTextRun firstRun = paragraph.Runs[0];
                double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, textMeasurer);
                double continuationTextStartOffset = GetParagraphTextStartOffset(paragraph);
                double labelStartOffset = GetParagraphLabelStartOffset(paragraph);
                double paragraphX = cellX + paddingLeft + textStartOffset;
                double paragraphWidth = Math.Max(1d, textWidth - textStartOffset - GetParagraphRightInset(paragraph));
                double continuationParagraphWidth = Math.Max(1d, textWidth - continuationTextStartOffset - GetParagraphRightInset(paragraph));
                bool firstLine = true;
                DocxWrappedTextLine[] wrappedLines = WrapTextLines(textSpans, paragraphWidth, continuationParagraphWidth, fontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints, allowOverwideTokenBreaks: true).ToArray();
                for (int lineIndex = 0; lineIndex < wrappedLines.Length; lineIndex++)
                {
                    DocxWrappedTextLine line = wrappedLines[lineIndex];
                    double lineWidth = MeasureTextSpans(line.Spans, fontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints);
                    double drawableLineWidth = MeasureDrawableTextSpans(line.Spans, fontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints);
                    double lineX = paragraph.Alignment switch
                    {
                        DocxTextAlignment.Center => paragraphX + Math.Max(0, paragraphWidth - lineWidth) / 2d,
                        DocxTextAlignment.Right => paragraphX + Math.Max(0, paragraphWidth - lineWidth),
                        _ => paragraphX
                    };
                    bool justifyLine = (paragraph.ListLabel is null || !firstLine) &&
                        ShouldJustifyTextLine(paragraph.Alignment, lineIndex == wrappedLines.Length - 1, drawableLineWidth, paragraphWidth, line.Spans);
                    IReadOnlyList<DocxTextSegmentLayout> segments = firstLine && paragraph.ListLabel is not null
                        ? CreateNumberedLineSegments(paragraph.ListLabel, line.Spans, firstRun, cellX + paddingLeft + labelStartOffset, lineX, fontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints)
                        : justifyLine
                            ? CreateJustifiedTextSegments(line.Spans, lineX, drawableLineWidth, paragraphWidth, fontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints)
                            : CreateTextSegments(line.Spans, lineX, fontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints);
                    double effectiveX = firstLine && paragraph.ListLabel is not null ? cellX + paddingLeft + labelStartOffset : lineX;
                    double effectiveWidth = firstLine && paragraph.ListLabel is not null
                        ? Math.Max(lineX + lineWidth, cellX + paddingLeft + labelStartOffset + MeasureListLabel(paragraph.ListLabel, firstRun, fontSize, textMeasurer)) - (cellX + paddingLeft + labelStartOffset)
                        : justifyLine
                            ? paragraphWidth
                            : lineWidth;
                    lines.Add(new DocxTextLineLayout(
                        firstLine && paragraph.ListLabel is not null ? paragraph.ListLabel.Text + GetListLabelTextSeparator(paragraph.ListLabel) + line.Text : line.Text,
                        firstRun,
                        fontSize,
                        effectiveX,
                        cursorY,
                        effectiveWidth,
                        segments,
                        SourceBlockIndex: null,
                        SourceLineIndex: lineIndex,
                        LineHeight: lineHeight,
                        AppliedBeforeSpacing: firstLine ? appliedBeforeSpacing : 0d,
                        IsFirstParagraphLine: firstLine,
                        EndsWithIntraTokenBreak: line.EndsWithIntraTokenBreak,
                        SingleLineHeight: lineHeightProfile.SingleLineHeight,
                        ListLabelSingleLineHeight: lineHeightProfile.ListLabelSingleLineHeight,
                        BodyWindowsLineHeight: lineHeightProfile.BodyWindowsLineHeight,
                        ListLabelWindowsLineHeight: lineHeightProfile.ListLabelWindowsLineHeight,
                        EffectiveLineSpacingFactor: lineHeightProfile.EffectiveLineSpacingFactor,
                        LineSpacingFactorFloorApplied: lineHeightProfile.LineSpacingFactorFloorApplied));
                    firstLine = false;
                    paragraphX = cellX + paddingLeft + continuationTextStartOffset;
                    paragraphWidth = Math.Max(1d, textWidth - continuationTextStartOffset - GetParagraphRightInset(paragraph));
                    cursorY -= lineHeight;
                }
            }

            pendingSpacingAfter = paragraph.SpacingAfterPoints;
            previousParagraph = paragraph;
        }

        cursorY -= pendingSpacingAfter;
        if (lines.Count == 0)
        {
            return lines;
        }

        double usedHeight = Math.Max(0d, startBaselineY - cursorY);
        double availableHeight = Math.Max(0d, cellHeight - paddingTop - paddingBottom - baselineInset);
        double extra = Math.Max(0d, availableHeight - usedHeight);
        double verticalOffset = cell.VerticalAlignmentValue?.Equals("bottom", StringComparison.OrdinalIgnoreCase) == true
            ? extra
            : cell.VerticalAlignmentValue?.Equals("center", StringComparison.OrdinalIgnoreCase) == true
                ? extra / 2d
                : 0d;
        return verticalOffset == 0d ? lines : ShiftTextLines(lines, -verticalOffset);
    }

    private static IReadOnlyList<DocxInlineImageLayout> LayoutTableCellInlineImages(
        DocxTableCell cell,
        double cellX,
        double cellY,
        double cellWidth,
        double cellHeight,
        double rowTopPadding,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        int pageIndex)
    {
        IReadOnlyList<DocxParagraph> paragraphs = cell.Paragraphs;
        if (paragraphs.Count == 0 || !paragraphs.Any(paragraph => paragraph.Images.Count != 0))
        {
            return [];
        }

        double paddingLeft = ResolveTableCellHorizontalPadding(cell.Margins.LeftPoints) + ResolveTableCellBorderContentInset(cell, "left");
        double paddingRight = ResolveTableCellHorizontalPadding(cell.Margins.RightPoints) + ResolveTableCellBorderContentInset(cell, "right");
        double paddingTop = rowTopPadding;
        double paddingBottom = ResolveTableCellVerticalPadding(cell.Margins.BottomPoints);
        double baselineInset = ResolveTableCellFirstBaselineInset(paragraphs);
        double textWidth = Math.Max(1d, cellWidth - paddingLeft - paddingRight);
        double startBaselineY = cellY + cellHeight - baselineInset - paddingTop;
        double cursorY = startBaselineY;
        var images = new List<DocxInlineImageLayout>();
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        foreach (DocxParagraph paragraph in paragraphs)
        {
            cursorY -= ShouldSuppressContextualSpacing(previousParagraph, paragraph)
                ? 0d
                : Math.Max(pendingSpacingAfter, paragraph.SpacingBeforePoints);
            pendingSpacingAfter = 0d;
            if (textMeasurer is not null)
            {
                double fontSize = GetParagraphFontSize(paragraph);
                double lineHeight = ResolveLineHeight(paragraph, fontSize, textMeasurer);
                IReadOnlyList<DocxTextSpan> textSpans = CreateTextSpans(paragraph.Runs);
                if (textSpans.Count != 0)
                {
                    double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, textMeasurer);
                    double firstParagraphWidth = Math.Max(1d, textWidth - textStartOffset - GetParagraphRightInset(paragraph));
                    double continuationParagraphWidth = Math.Max(1d, textWidth - GetParagraphTextStartOffset(paragraph) - GetParagraphRightInset(paragraph));
                    cursorY -= WrapTextLines(textSpans, firstParagraphWidth, continuationParagraphWidth, fontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints, allowOverwideTokenBreaks: true).Count() * lineHeight;
                }
                else if (paragraph.Images.Count == 0)
                {
                    cursorY -= lineHeight;
                }
            }

            foreach (DocxInlineImage image in paragraph.Images)
            {
                double paragraphX = cellX + paddingLeft + GetParagraphStartOffset(paragraph);
                double paragraphWidth = Math.Max(1d, textWidth - GetParagraphStartOffset(paragraph) - GetParagraphRightInset(paragraph));
                double imageWidth = Math.Min(paragraphWidth, image.WidthPoints);
                double imageHeight = image.HeightPoints * imageWidth / Math.Max(1d, image.WidthPoints);
                double imageX = paragraph.Alignment switch
                {
                    DocxTextAlignment.Center => paragraphX + Math.Max(0, paragraphWidth - imageWidth) / 2d,
                    DocxTextAlignment.Right => paragraphX + Math.Max(0, paragraphWidth - imageWidth),
                    _ => paragraphX
                };
                images.Add(new DocxInlineImageLayout(image, imageX, cursorY - imageHeight, imageWidth, imageHeight, pageIndex));
                cursorY -= imageHeight + 6d;
            }

            pendingSpacingAfter = paragraph.SpacingAfterPoints;
            previousParagraph = paragraph;
        }

        cursorY -= pendingSpacingAfter;
        if (images.Count == 0)
        {
            return images;
        }

        double usedHeight = Math.Max(0d, startBaselineY - cursorY);
        double availableHeight = Math.Max(0d, cellHeight - paddingTop - paddingBottom - baselineInset);
        double extra = Math.Max(0d, availableHeight - usedHeight);
        double verticalOffset = cell.VerticalAlignmentValue?.Equals("bottom", StringComparison.OrdinalIgnoreCase) == true
            ? extra
            : cell.VerticalAlignmentValue?.Equals("center", StringComparison.OrdinalIgnoreCase) == true
                ? extra / 2d
                : 0d;
        return verticalOffset == 0d
            ? images
            : images.Select(image => image with { Y = image.Y - verticalOffset }).ToArray();
    }

    private static IReadOnlyList<DocxTextLineLayout> ShiftTextLines(IReadOnlyList<DocxTextLineLayout> lines, double deltaY)
    {
        return lines
            .Select(line => line with { BaselineY = line.BaselineY + deltaY })
            .ToArray();
    }

    private static double ResolveTableCellHorizontalPadding(double? points)
    {
        return Math.Max(0d, points ?? 0d);
    }

    private static double ResolveTableCellBorderContentInset(DocxTableCell cell, string edge)
    {
        return DocxTableBorderGeometry.ResolveVisibleWidth(DocxTableBorderGeometry.Find(cell.Borders, edge)) / 2d;
    }

    private static double ResolveTableCellVerticalPadding(double? points)
    {
        return Math.Max(0d, points ?? 0d);
    }

    private static double ResolveTableRowTopPadding(DocxTableRow row)
    {
        return row.Cells
            .Select(cell => ResolveTableCellVerticalPadding(cell.Margins.TopPoints))
            .DefaultIfEmpty(0d)
            .Max();
    }

    private static double ResolveTableCellFirstBaselineInset(IReadOnlyList<DocxParagraph> paragraphs)
    {
        return DocxLineMetrics.ResolveTableCellFirstBaselineInset(paragraphs);
    }

    private static IReadOnlyList<DocxTextSpan> CreateTextSpans(IReadOnlyList<DocxTextRun> runs)
    {
        if (runs.Count != 0 && runs.All(run => run.Text.Length == 0 || run.Hidden))
        {
            DocxTextRun? paragraphMark = runs.FirstOrDefault(run => !run.Hidden);
            return paragraphMark is null ? [] : [new DocxTextSpan(" ", paragraphMark)];
        }

        return runs
            .Where(run => run.Text.Length != 0 && !run.Hidden)
            .Select(run => new DocxTextSpan(run.Text, run))
            .ToArray();
    }

    private static double GetParagraphFontSize(DocxParagraph paragraph)
    {
        return paragraph.Runs.Count == 0 ? 11d : paragraph.Runs.Max(run => run.FontSize);
    }

    private static IReadOnlyList<DocxTextSegmentLayout> CreateTextSegments(
        IReadOnlyList<DocxTextSpan> spans,
        double lineX,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints)
    {
        var segments = new List<DocxTextSegmentLayout>(spans.Count);
        double segmentX = lineX;
        for (int i = 0; i < spans.Count; i++)
        {
            DocxTextSpan span = spans[i];
            segmentX = AddTextSegments(segments, span, segmentX, lineX, fontSize, textMeasurer, tabStops, defaultTabStopPoints);
            if (i + 1 < spans.Count)
            {
                segmentX += DocxTextSpacing.BoundarySpacing(span.StyleRun, span.Text, spans[i + 1].Text);
            }
        }

        return segments;
    }

    private static double MeasureTextSpans(
        IReadOnlyList<DocxTextSpan> spans,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints)
    {
        double width = 0d;
        for (int i = 0; i < spans.Count; i++)
        {
            DocxTextSpan span = spans[i];
            width = MeasureTextSpanAdvance(span, width, fontSize, textMeasurer, tabStops, defaultTabStopPoints);
            if (i + 1 < spans.Count)
            {
                width += DocxTextSpacing.BoundarySpacing(span.StyleRun, span.Text, spans[i + 1].Text);
            }
        }

        return width;
    }

    private static double MeasureDrawableTextSpans(
        IReadOnlyList<DocxTextSpan> spans,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints)
    {
        int length = spans.Sum(span => span.Text.Length);
        int drawableLength = FindDrawableTextLength(spans);
        return drawableLength == length
            ? MeasureTextSpans(spans, fontSize, textMeasurer, tabStops, defaultTabStopPoints)
            : MeasureTextSpans(SliceTextSpans(spans, 0, drawableLength), fontSize, textMeasurer, tabStops, defaultTabStopPoints);
    }

    private static bool ShouldJustifyTextLine(
        DocxTextAlignment alignment,
        bool isLastLine,
        double drawableLineWidth,
        double paragraphWidth,
        IReadOnlyList<DocxTextSpan> spans)
    {
        return alignment == DocxTextAlignment.Justified &&
            !isLastLine &&
            paragraphWidth - drawableLineWidth > 0.001d &&
            CountStretchableJustificationSpaces(spans) > 0 &&
            !spans.Any(span => span.Text.IndexOf('\t') >= 0);
    }

    private static IReadOnlyList<DocxTextSegmentLayout> CreateJustifiedTextSegments(
        IReadOnlyList<DocxTextSpan> spans,
        double lineX,
        double drawableLineWidth,
        double paragraphWidth,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints)
    {
        int stretchableSpaces = CountStretchableJustificationSpaces(spans);
        if (stretchableSpaces == 0 || spans.Any(span => span.Text.IndexOf('\t') >= 0))
        {
            return CreateTextSegments(spans, lineX, fontSize, textMeasurer, tabStops, defaultTabStopPoints);
        }

        double extraPerSpace = Math.Max(0d, paragraphWidth - drawableLineWidth) / stretchableSpaces;
        int drawableLength = FindDrawableTextLength(spans);
        IReadOnlyList<DocxTextSpan> drawableSpans = SliceTextSpans(spans, 0, drawableLength);
        var segments = new List<DocxTextSegmentLayout>(drawableSpans.Count);
        double segmentX = lineX;
        for (int i = 0; i < drawableSpans.Count; i++)
        {
            DocxTextSpan span = drawableSpans[i];
            segmentX = AddJustifiedSpanSegments(segments, span, segmentX, fontSize, textMeasurer, extraPerSpace);
            if (i + 1 < drawableSpans.Count)
            {
                segmentX += DocxTextSpacing.BoundarySpacing(span.StyleRun, span.Text, drawableSpans[i + 1].Text);
            }
        }

        return segments;
    }

    private static double AddJustifiedSpanSegments(
        List<DocxTextSegmentLayout> segments,
        DocxTextSpan span,
        double segmentX,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        double extraPerSpace)
    {
        double spanFontSize = GetTextSpanFontSize(span, fontSize);
        double baselineOffset = GetTextSpanBaselineOffset(span, fontSize);
        int start = 0;
        while (start < span.Text.Length)
        {
            bool isSpace = IsJustificationSpace(span.Text[start]);
            int end = start + 1;
            while (end < span.Text.Length && IsJustificationSpace(span.Text[end]) == isSpace)
            {
                end++;
            }

            string text = span.Text[start..end];
            double width = textMeasurer.MeasureText(span.StyleRun, text, spanFontSize);
            if (isSpace)
            {
                segmentX += width + text.Length * extraPerSpace;
            }
            else
            {
                segments.Add(new DocxTextSegmentLayout(text, span.StyleRun, segmentX, width, spanFontSize, baselineOffset));
                segmentX += width;
            }

            start = end;
        }

        return segmentX;
    }

    private static int CountStretchableJustificationSpaces(IReadOnlyList<DocxTextSpan> spans)
    {
        int drawableLength = FindDrawableTextLength(spans);
        int seen = 0;
        int count = 0;
        foreach (DocxTextSpan span in spans)
        {
            foreach (char c in span.Text)
            {
                if (seen++ >= drawableLength)
                {
                    return count;
                }

                if (IsJustificationSpace(c))
                {
                    count++;
                }
            }
        }

        return count;
    }

    private static int FindDrawableTextLength(IReadOnlyList<DocxTextSpan> spans)
    {
        int length = spans.Sum(span => span.Text.Length);
        int index = length;
        for (int spanIndex = spans.Count - 1; spanIndex >= 0; spanIndex--)
        {
            string text = spans[spanIndex].Text;
            for (int i = text.Length - 1; i >= 0; i--)
            {
                index--;
                if (!IsBreakableWhitespaceChar(text[i]))
                {
                    return index + 1;
                }
            }
        }

        return 0;
    }

    private static bool IsJustificationSpace(char c)
    {
        return c == ' ';
    }

    private static double AddTextSegments(
        List<DocxTextSegmentLayout> segments,
        DocxTextSpan span,
        double segmentX,
        double lineX,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints)
    {
        double spanFontSize = GetTextSpanFontSize(span, fontSize);
        double baselineOffset = GetTextSpanBaselineOffset(span, fontSize);
        int start = 0;
        for (int i = 0; i <= span.Text.Length; i++)
        {
            if (i < span.Text.Length && span.Text[i] != '\t')
            {
                continue;
            }

            if (i > start)
            {
                string text = span.Text[start..i];
                segmentX = AddTextSegment(segments, span.StyleRun, text, segmentX, spanFontSize, baselineOffset, textMeasurer);
            }

            if (i < span.Text.Length)
            {
                segmentX = lineX + AdvanceToNextTabStop(segmentX - lineX, tabStops, defaultTabStopPoints);
                start = i + 1;
            }
        }

        return segmentX;
    }

    private static double AddTextSegment(
        List<DocxTextSegmentLayout> segments,
        DocxTextRun styleRun,
        string text,
        double segmentX,
        double fontSize,
        double baselineOffset,
        IDocxTextMeasurer textMeasurer)
    {
        int leadingSpaces = CountLeadingOfficeSeparatedSpaces(text);
        if (leadingSpaces == 0 || leadingSpaces == text.Length)
        {
            double width = textMeasurer.MeasureText(styleRun, text, fontSize);
            segments.Add(new DocxTextSegmentLayout(text, styleRun, segmentX, width, fontSize, baselineOffset));
            return segmentX + width;
        }

        string spaceText = text[..leadingSpaces];
        double spaceWidth = textMeasurer.MeasureText(styleRun, spaceText, fontSize);
        segments.Add(new DocxTextSegmentLayout(spaceText, styleRun, segmentX, spaceWidth, fontSize, baselineOffset));
        segmentX += spaceWidth + DocxTextSpacing.BoundarySpacing(styleRun, spaceText, text[leadingSpaces..]);

        string bodyText = text[leadingSpaces..];
        double bodyWidth = textMeasurer.MeasureText(styleRun, bodyText, fontSize);
        segments.Add(new DocxTextSegmentLayout(bodyText, styleRun, segmentX, bodyWidth, fontSize, baselineOffset));
        return segmentX + bodyWidth;
    }

    private static int CountLeadingOfficeSeparatedSpaces(string text)
    {
        int count = 0;
        while (count < text.Length && text[count] == ' ')
        {
            count++;
        }

        return count;
    }

    private static double MeasureTextSpanAdvance(
        DocxTextSpan span,
        double currentWidth,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints)
    {
        double spanFontSize = GetTextSpanFontSize(span, fontSize);
        int start = 0;
        for (int i = 0; i <= span.Text.Length; i++)
        {
            if (i < span.Text.Length && span.Text[i] != '\t')
            {
                continue;
            }

            if (i > start)
            {
                currentWidth += textMeasurer.MeasureText(span.StyleRun, span.Text[start..i], spanFontSize);
            }

            if (i < span.Text.Length)
            {
                currentWidth = AdvanceToNextTabStop(currentWidth, tabStops, defaultTabStopPoints);
                start = i + 1;
            }
        }

        return currentWidth;
    }

    private static double GetTextSpanFontSize(DocxTextSpan span, double fallbackFontSize)
    {
        double nominalFontSize = span.StyleRun.FontSize > 0d ? span.StyleRun.FontSize : fallbackFontSize;
        return DocxVerticalAlignMetrics.ResolveFontSize(nominalFontSize, span.StyleRun);
    }

    private static double GetTextSpanBaselineOffset(DocxTextSpan span, double fallbackFontSize)
    {
        double nominalFontSize = span.StyleRun.FontSize > 0d ? span.StyleRun.FontSize : fallbackFontSize;
        double layoutFontSize = DocxVerticalAlignMetrics.ResolveFontSize(nominalFontSize, span.StyleRun);
        return DocxVerticalAlignMetrics.ResolveBaselineOffset(nominalFontSize, layoutFontSize, span.StyleRun);
    }

    private static double AdvanceToNextDefaultTabStop(double width, double defaultTabStopPoints)
    {
        double tabStop = defaultTabStopPoints > 0d ? defaultTabStopPoints : WordDefaultTabStopPoints;
        return (Math.Floor(width / tabStop) + 1d) * tabStop;
    }

    private static double AdvanceToNextTabStop(double width, IReadOnlyList<DocxTabStop> tabStops, double defaultTabStopPoints)
    {
        foreach (DocxTabStop tabStop in tabStops
            .Where(tabStop => tabStop.PositionPoints is not null && IsPositioningTabStop(tabStop))
            .OrderBy(tabStop => tabStop.PositionPoints!.Value))
        {
            if (tabStop.PositionPoints!.Value > width + 0.001d)
            {
                return tabStop.PositionPoints.Value;
            }
        }

        return AdvanceToNextDefaultTabStop(width, defaultTabStopPoints);
    }

    private static bool IsPositioningTabStop(DocxTabStop tabStop)
    {
        return !string.Equals(tabStop.Value, "bar", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(tabStop.Value, "clear", StringComparison.OrdinalIgnoreCase);
    }

    private static IEnumerable<DocxWrappedTextLine> WrapTextLines(
        IReadOnlyList<DocxTextSpan> spans,
        double firstLineMaxWidth,
        double continuationLineMaxWidth,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints,
        bool allowOverwideTokenBreaks)
    {
        string text = string.Concat(spans.Select(span => span.Text));
        int lineIndex = 0;
        int segmentStart = 0;
        while (segmentStart <= text.Length)
        {
            int breakIndex = text.IndexOf('\n', segmentStart);
            int segmentLength = breakIndex < 0 ? text.Length - segmentStart : breakIndex - segmentStart;
            bool yielded = false;
            foreach (DocxWrappedTextLine line in WrapWords(text, spans, segmentStart, segmentLength, index => index == 0 && lineIndex == 0 ? firstLineMaxWidth : continuationLineMaxWidth, fontSize, textMeasurer, tabStops, defaultTabStopPoints, allowOverwideTokenBreaks))
            {
                yielded = true;
                yield return line;
                lineIndex++;
            }

            if (!yielded && segmentLength == 0)
            {
                yield return new DocxWrappedTextLine(string.Empty, []);
                lineIndex++;
            }

            if (breakIndex < 0)
            {
                yield break;
            }

            segmentStart = breakIndex + 1;
        }
    }

    private static IEnumerable<DocxWrappedTextLine> WrapWords(
        string text,
        IReadOnlyList<DocxTextSpan> spans,
        int segmentStart,
        int segmentLength,
        Func<int, double> maxWidth,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints,
        bool allowOverwideTokenBreaks)
    {
        IReadOnlyList<TextToken> tokens = TokenizeSpaces(text, segmentStart, segmentLength);
        if (tokens.Count == 0)
        {
            yield break;
        }

        int lineStart = tokens[0].Start;
        int lineLength = 0;
        int lineIndex = 0;
        for (int tokenIndex = 0; tokenIndex < tokens.Count; tokenIndex++)
        {
            TextToken token = tokens[tokenIndex];
            int candidateLength = token.Start + token.Length - lineStart;
            bool lineHasNonWhitespace = HasNonWhitespace(text, lineStart, lineLength);
            if (lineLength > 0 &&
                lineHasNonWhitespace &&
                !token.IsBreakableWhitespace &&
                MeasureTextSpans(SliceTextSpans(spans, lineStart, candidateLength), fontSize, textMeasurer, tabStops, defaultTabStopPoints) > maxWidth(lineIndex))
            {
                yield return CreateWrappedTextLine(text, spans, lineStart, lineLength);
                lineIndex++;
                lineStart = token.Start;
                lineLength = 0;
                tokenIndex--;
                continue;
            }

            if (lineLength == 0 &&
                allowOverwideTokenBreaks &&
                !token.IsBreakableWhitespace &&
                TryFindOverwideTokenBreak(text, spans, token, maxWidth(lineIndex), fontSize, textMeasurer, tabStops, defaultTabStopPoints, out int breakLength))
            {
                yield return CreateWrappedTextLine(text, spans, token.Start, breakLength, endsWithIntraTokenBreak: true);
                lineIndex++;
                lineStart = token.Start + breakLength;
                lineLength = 0;
                tokens = ReplaceToken(tokens, tokenIndex, new TextToken(text.Substring(lineStart, token.Length - breakLength), lineStart, token.Length - breakLength));
                tokenIndex--;
            }
            else
            {
                lineLength = candidateLength;
            }
        }

        if (lineLength > 0)
        {
            yield return CreateWrappedTextLine(text, spans, lineStart, lineLength);
        }
    }

    private static bool TryFindOverwideTokenBreak(
        string text,
        IReadOnlyList<DocxTextSpan> spans,
        TextToken token,
        double maxWidth,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints,
        out int breakLength)
    {
        breakLength = 0;
        double tokenWidth = MeasureTextSpans(SliceTextSpans(spans, token.Start, token.Length), fontSize, textMeasurer, tabStops, defaultTabStopPoints);
        if (tokenWidth <= maxWidth)
        {
            return false;
        }

        for (int length = token.Length - 1; length > 0; length--)
        {
            int absoluteIndex = token.Start + length - 1;
            if (!DocxLineBreakOpportunities.IsOpportunityAfter(text[absoluteIndex]))
            {
                continue;
            }

            double prefixWidth = MeasureTextSpans(SliceTextSpans(spans, token.Start, length), fontSize, textMeasurer, tabStops, defaultTabStopPoints);
            if (prefixWidth <= maxWidth)
            {
                breakLength = length;
                return true;
            }
        }

        for (int length = token.Length - 1; length > 0; length--)
        {
            if (!IsSafeEmergencyTokenBreak(text, token, length))
            {
                continue;
            }

            double prefixWidth = MeasureTextSpans(SliceTextSpans(spans, token.Start, length), fontSize, textMeasurer, tabStops, defaultTabStopPoints);
            if (prefixWidth <= maxWidth)
            {
                breakLength = length;
                return true;
            }
        }

        return false;
    }

    private static bool IsSafeEmergencyTokenBreak(string text, TextToken token, int length)
    {
        if (length <= 0 || length >= token.Length)
        {
            return false;
        }

        int breakIndex = token.Start + length;
        char before = text[breakIndex - 1];
        char after = text[breakIndex];
        if (char.IsHighSurrogate(before) && char.IsLowSurrogate(after))
        {
            return false;
        }

        UnicodeCategory afterCategory = char.GetUnicodeCategory(after);
        return afterCategory is not UnicodeCategory.NonSpacingMark and
            not UnicodeCategory.SpacingCombiningMark and
            not UnicodeCategory.EnclosingMark;
    }

    private static IReadOnlyList<TextToken> ReplaceToken(IReadOnlyList<TextToken> tokens, int index, TextToken replacement)
    {
        var result = new List<TextToken>(tokens.Count);
        for (int i = 0; i < tokens.Count; i++)
        {
            result.Add(i == index ? replacement : tokens[i]);
        }

        return result;
    }

    private static DocxWrappedTextLine CreateWrappedTextLine(
        string text,
        IReadOnlyList<DocxTextSpan> spans,
        int start,
        int length,
        bool endsWithIntraTokenBreak = false)
    {
        return new DocxWrappedTextLine(text.Substring(start, length), SliceTextSpans(spans, start, length), endsWithIntraTokenBreak);
    }

    private static IReadOnlyList<DocxTextSpan> SliceTextSpans(
        IReadOnlyList<DocxTextSpan> spans,
        int start,
        int length)
    {
        if (length == 0)
        {
            return [];
        }

        var sliced = new List<DocxTextSpan>();
        int spanStart = 0;
        int end = start + length;
        foreach (DocxTextSpan span in spans)
        {
            int spanEnd = spanStart + span.Text.Length;
            int sliceStart = Math.Max(start, spanStart);
            int sliceEnd = Math.Min(end, spanEnd);
            if (sliceStart < sliceEnd)
            {
                sliced.Add(new DocxTextSpan(span.Text[(sliceStart - spanStart)..(sliceEnd - spanStart)], span.StyleRun));
            }

            if (spanEnd >= end)
            {
                break;
            }

            spanStart = spanEnd;
        }

        return sliced;
    }

    private static IReadOnlyList<TextToken> TokenizeSpaces(string text, int start, int length)
    {
        if (length == 0)
        {
            return [];
        }

        string segment = text.Substring(start, length);
        return TokenizeSpaces(segment)
            .Select(token => new TextToken(token.Text, start + token.Start, token.Length))
            .ToArray();
    }

    private static bool HasNonWhitespace(string text, int start, int length)
    {
        for (int i = start; i < start + length; i++)
        {
            if (!char.IsWhiteSpace(text[i]))
            {
                return true;
            }
        }

        return false;
    }

    private static IReadOnlyList<TextToken> TokenizeSpaces(string text)
    {
        if (text.Length == 0)
        {
            return [];
        }

        var tokens = new List<TextToken>();
        int start = 0;
        bool inBreakableWhitespace = IsBreakableWhitespaceChar(text[0]);
        for (int i = 1; i < text.Length; i++)
        {
            bool breakableWhitespace = IsBreakableWhitespaceChar(text[i]);
            if (breakableWhitespace == inBreakableWhitespace)
            {
                continue;
            }

            tokens.Add(new TextToken(text[start..i], start, i - start));
            start = i;
            inBreakableWhitespace = breakableWhitespace;
        }

        tokens.Add(new TextToken(text[start..], start, text.Length - start));
        return tokens;
    }

    private static bool IsBreakableWhitespaceChar(char value)
    {
        return char.IsWhiteSpace(value) &&
            value != '\u00A0' &&
            value != '\u202F' &&
            value != '\u2007';
    }

    private static class DocxLineBreakOpportunities
    {
        public static bool IsOpportunityAfter(char value)
        {
            return value is '-' or '/' or '\\' or '\u2010' or '\u2011' or '\u2012' or '\u2013' or '\u2014';
        }
    }

    private readonly record struct TextToken(string Text, int Start, int Length)
    {
        public bool IsBreakableWhitespace => Text.All(IsBreakableWhitespaceChar);
    }
}
