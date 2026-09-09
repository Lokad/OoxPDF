using System.Globalization;

namespace Lokad.OoxPdf.Docx;

internal sealed partial record DocxStructureSnapshot
{
    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static DocxStructureBlockSnapshot FromParagraph(
        int blockIndex,
        string? previousKind,
        string? nextKind,
        DocxParagraph paragraph,
        IReadOnlyList<DocxRelatedStory> relatedStories)
    {
            bool HasVisibleText(DocxParagraph paragraph)
            {
                return paragraph.Runs.Any(run => !string.IsNullOrWhiteSpace(run.Text));
            }

        DocxEffectiveParagraphProperties effective = paragraph.EffectiveProperties;
        return new DocxStructureBlockSnapshot(
            blockIndex,
            "Paragraph",
            previousKind,
            nextKind,
            ParagraphStyleId: effective.StyleId,
            ParagraphStyleFound: effective.StyleResolution.StyleFound,
            ParagraphStyleDepth: effective.StyleResolution.StyleDepth,
            HasDocumentDefaultParagraphProperties: effective.StyleResolution.HasDocumentDefaultParagraphProperties,
            HasDirectParagraphProperties: effective.StyleResolution.HasDirectParagraphProperties,
            HasTableStyleParagraphProperties: effective.StyleResolution.HasTableStyleParagraphProperties,
            RunCount: paragraph.Runs.Count,
            TextLength: TextLength(paragraph),
            HasVisibleText: HasVisibleText(paragraph),
            RevisionCount: CountRevisions(paragraph),
            InsertionRevisionCount: CountRevisions(paragraph, DocxRevisionKind.Insertion),
            DeletionRevisionCount: CountRevisions(paragraph, DocxRevisionKind.Deletion),
            MoveFromRevisionCount: CountRevisions(paragraph, DocxRevisionKind.MoveFrom),
            MoveToRevisionCount: CountRevisions(paragraph, DocxRevisionKind.MoveTo),
            OtherRevisionCount: CountOtherRevisions(paragraph),
            RevisionRangeCount: paragraph.RevisionRanges.Count,
            ListFormatValue: paragraph.ListLabel?.FormatValue,
            InlineImageCount: paragraph.Images.Count,
            InlineReferenceCount: paragraph.InlineReferences.Count,
            AnchoredInlineReferenceCount: paragraph.InlineReferences.Count(HasInlineReferenceAnchor),
            ResolvedInlineReferenceCount: ParagraphResolvedInlineReferenceCount(paragraph, relatedStories),
            MaxInlineReferenceTextOffsetInRun: paragraph.InlineReferences.Select(reference => reference.TextOffsetInRun).DefaultIfEmpty(0).Max(),
            FieldReferenceCount: paragraph.FieldReferences.Count,
            PageFieldReferenceCount: paragraph.FieldReferences.Count(reference => reference.Kind == DocxFieldKind.Page),
            NumPagesFieldReferenceCount: paragraph.FieldReferences.Count(reference => reference.Kind == DocxFieldKind.NumPages),
            OtherFieldReferenceCount: paragraph.FieldReferences.Count(reference => reference.Kind == DocxFieldKind.Other),
            ComplexFieldReferenceCount: ParagraphComplexFieldReferenceCount(paragraph),
            CachedResultFieldReferenceCount: ParagraphCachedResultFieldReferenceCount(paragraph),
            RenderedCachedResultFieldReferenceCount: ParagraphRenderedCachedResultFieldReferenceCount(paragraph),
            PlaceholderFieldReferenceCount: ParagraphPlaceholderFieldReferenceCount(paragraph),
            NestedFieldReferenceCount: ParagraphNestedFieldReferenceCount(paragraph),
            BookmarkAnchorCount: paragraph.BookmarkAnchors.Count,
            CommentReferenceCount: paragraph.InlineReferences.Count(reference => reference.Kind == DocxRelatedStoryKind.Comment),
            FootnoteReferenceCount: paragraph.InlineReferences.Count(reference => reference.Kind == DocxRelatedStoryKind.Footnote),
            EndnoteReferenceCount: paragraph.InlineReferences.Count(reference => reference.Kind == DocxRelatedStoryKind.Endnote),
            HyperlinkCount: paragraph.Hyperlinks.Count,
            ExternalHyperlinkCount: paragraph.Hyperlinks.Count(link => string.Equals(link.TargetMode, "External", StringComparison.OrdinalIgnoreCase)),
            InternalHyperlinkCount: paragraph.Hyperlinks.Count(link => link.Anchor is not null || link.ResolvedTarget is not null),
            SpacingBeforePoints: effective.SpacingBeforePoints,
            SpacingAfterPoints: effective.SpacingAfterPoints,
            LineSpacingPoints: effective.LineSpacingPoints,
            LineSpacingFactor: effective.LineSpacingFactor,
            HasBeforeSpacingToken: DocxParagraphSpacing.HasBeforeSpacingSide(effective.Spacing),
            HasAfterSpacingToken: DocxParagraphSpacing.HasAfterSpacingSide(effective.Spacing),
            BeforeAutoSpacingValue: effective.Spacing.BeforeAutoSpacingValue,
            AfterAutoSpacingValue: effective.Spacing.AfterAutoSpacingValue,
            ContextualSpacing: effective.Spacing.ContextualSpacing,
            KeepNext: effective.KeepRules.KeepNext,
            KeepLines: effective.KeepRules.KeepLines,
            WidowControl: effective.KeepRules.WidowControl,
            WordWrap: effective.WordWrap,
            WordWrapValue: effective.WordWrapValue,
            ParagraphIndentLeftPoints: effective.Indent.LeftPoints,
            ParagraphIndentRightPoints: effective.Indent.RightPoints,
            ParagraphIndentFirstLinePoints: effective.Indent.FirstLinePoints,
            ParagraphIndentHangingPoints: effective.Indent.HangingPoints,
            WhitespaceDelimitedTokenCount: CountWhitespaceDelimitedTokens(paragraph),
            LongestWhitespaceDelimitedTokenLength: LongestWhitespaceDelimitedTokenLength(paragraph),
            SpaceCharacterCount: CountCharacters(paragraph, static c => c == ' '),
            NonAsciiCharacterCount: CountCharacters(paragraph, static c => c > 127),
            PunctuationCharacterCount: CountCharacters(paragraph, char.IsPunctuation),
            DigitCharacterCount: CountCharacters(paragraph, char.IsDigit),
            UppercaseCharacterCount: CountCharacters(paragraph, char.IsUpper),
            LowercaseCharacterCount: CountCharacters(paragraph, char.IsLower),
            TabStopCount: effective.TabStops.Count,
            SnapToGrid: effective.SnapToGrid,
            SnapToGridValue: effective.SnapToGridValue,
            TableRowCount: 0,
            TableIndex: null,
            TableMaxColumnCount: null,
            TablePreferredWidthPoints: null,
            TablePreferredWidthType: null,
            TableIndentPoints: null,
            TableCellSpacingPoints: null,
            TableLayoutValue: null,
            PageBreakSourceKind: null,
            PageBreakValue: null,
            ManualBreakSourceKind: null,
            ManualBreakValue: null,
            ImplicitSourceKind: null,
            SectionBreakTypeValue: null,
            SectionColumnCountValue: null,
            SectionColumnEqualWidthValue: null,
            SectionColumnSpaceValue: null,
            SectionColumnDefinitionCount: null,
            SectionColumnDefinitionWidthTokenCount: null,
            SectionColumnDefinitionSpaceTokenCount: null,
            PageBreakConsumesParagraphLine: null,
            PageBreakLineSpacingPoints: null,
            PageBreakLineSpacingFactor: null,
            ManualBreakConsumesParagraphLine: null,
            ManualBreakLineSpacingPoints: null,
            ManualBreakLineSpacingFactor: null);
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static DocxStructureBlockSnapshot FromTable(
        int blockIndex,
        string? previousKind,
        string? nextKind,
        DocxTable table,
        int tableIndex,
        IReadOnlyList<DocxRelatedStory> relatedStories)
    {
        DocxParagraph[] paragraphs = DocxBlockTraversal.EnumerateTableParagraphs(table).ToArray();
        DocxRevisionInfo[] revisions = EnumerateTableRevisions(table).ToArray();
        return new DocxStructureBlockSnapshot(
            blockIndex,
            "Table",
            previousKind,
            nextKind,
            RevisionCount: CountRevisions(revisions),
            InsertionRevisionCount: CountRevisions(revisions, DocxRevisionKind.Insertion),
            DeletionRevisionCount: CountRevisions(revisions, DocxRevisionKind.Deletion),
            MoveFromRevisionCount: CountRevisions(revisions, DocxRevisionKind.MoveFrom),
            MoveToRevisionCount: CountRevisions(revisions, DocxRevisionKind.MoveTo),
            OtherRevisionCount: CountOtherRevisions(revisions),
            RevisionRangeCount: paragraphs.Sum(paragraph => paragraph.RevisionRanges.Count),
            InlineReferenceCount: paragraphs.Sum(ParagraphInlineReferenceCount),
            AnchoredInlineReferenceCount: paragraphs.Sum(paragraph => paragraph.InlineReferences.Count(HasInlineReferenceAnchor)),
            ResolvedInlineReferenceCount: paragraphs.Sum(paragraph => ParagraphResolvedInlineReferenceCount(paragraph, relatedStories)),
            MaxInlineReferenceTextOffsetInRun: paragraphs.SelectMany(paragraph => paragraph.InlineReferences).Select(reference => reference.TextOffsetInRun).DefaultIfEmpty(0).Max(),
            CommentReferenceCount: paragraphs.Sum(paragraph => paragraph.InlineReferences.Count(reference => reference.Kind == DocxRelatedStoryKind.Comment)),
            FootnoteReferenceCount: paragraphs.Sum(paragraph => paragraph.InlineReferences.Count(reference => reference.Kind == DocxRelatedStoryKind.Footnote)),
            EndnoteReferenceCount: paragraphs.Sum(paragraph => paragraph.InlineReferences.Count(reference => reference.Kind == DocxRelatedStoryKind.Endnote)),
            FieldReferenceCount: paragraphs.Sum(ParagraphFieldReferenceCount),
            PageFieldReferenceCount: paragraphs.Sum(ParagraphPageFieldReferenceCount),
            NumPagesFieldReferenceCount: paragraphs.Sum(ParagraphNumPagesFieldReferenceCount),
            OtherFieldReferenceCount: paragraphs.Sum(ParagraphOtherFieldReferenceCount),
            ComplexFieldReferenceCount: paragraphs.Sum(ParagraphComplexFieldReferenceCount),
            CachedResultFieldReferenceCount: paragraphs.Sum(ParagraphCachedResultFieldReferenceCount),
            RenderedCachedResultFieldReferenceCount: paragraphs.Sum(ParagraphRenderedCachedResultFieldReferenceCount),
            PlaceholderFieldReferenceCount: paragraphs.Sum(ParagraphPlaceholderFieldReferenceCount),
            NestedFieldReferenceCount: paragraphs.Sum(ParagraphNestedFieldReferenceCount),
            BookmarkAnchorCount: paragraphs.Sum(ParagraphBookmarkAnchorCount),
            HyperlinkCount: paragraphs.Sum(ParagraphHyperlinkCount),
            ExternalHyperlinkCount: paragraphs.Sum(ParagraphExternalHyperlinkCount),
            InternalHyperlinkCount: paragraphs.Sum(ParagraphInternalHyperlinkCount),
            TableRowCount: table.Rows.Count,
            TableIndex: tableIndex,
            TableMaxColumnCount: MaxColumnCount(table),
            TablePreferredWidthPoints: table.PreferredWidthPoints,
            TablePreferredWidthType: table.PreferredWidthType,
            TableIndentPoints: table.IndentPoints,
            TableCellSpacingPoints: table.CellSpacingPoints,
            TableLayoutValue: table.LayoutValue,
            ParagraphStyleId: null,
            ParagraphStyleFound: null,
            ParagraphStyleDepth: null,
            HasDocumentDefaultParagraphProperties: null,
            HasDirectParagraphProperties: null,
            HasTableStyleParagraphProperties: null,
            RunCount: 0,
            TextLength: 0,
            HasVisibleText: false,
            ListFormatValue: null,
            InlineImageCount: 0,
            SpacingBeforePoints: null,
            SpacingAfterPoints: null,
            LineSpacingPoints: null,
            LineSpacingFactor: null,
            HasBeforeSpacingToken: false,
            HasAfterSpacingToken: false,
            BeforeAutoSpacingValue: null,
            AfterAutoSpacingValue: null,
            ContextualSpacing: null,
            KeepNext: null,
            KeepLines: null,
            WidowControl: null,
            WordWrap: null,
            WordWrapValue: null,
            PageBreakSourceKind: null,
            PageBreakValue: null,
            ManualBreakSourceKind: null,
            ManualBreakValue: null,
            ImplicitSourceKind: null,
            SectionBreakTypeValue: null,
            SectionColumnCountValue: null,
            SectionColumnEqualWidthValue: null,
            SectionColumnSpaceValue: null,
            SectionColumnDefinitionCount: null,
            SectionColumnDefinitionWidthTokenCount: null,
            SectionColumnDefinitionSpaceTokenCount: null,
            ParagraphIndentLeftPoints: null,
            ParagraphIndentRightPoints: null,
            ParagraphIndentFirstLinePoints: null,
            ParagraphIndentHangingPoints: null,
            WhitespaceDelimitedTokenCount: null,
            LongestWhitespaceDelimitedTokenLength: null,
            SpaceCharacterCount: null,
            NonAsciiCharacterCount: null,
            PunctuationCharacterCount: null,
            DigitCharacterCount: null,
            UppercaseCharacterCount: null,
            LowercaseCharacterCount: null,
            TabStopCount: null,
            SnapToGrid: null,
            SnapToGridValue: null,
            PageBreakConsumesParagraphLine: null,
            PageBreakLineSpacingPoints: null,
            PageBreakLineSpacingFactor: null,
            ManualBreakConsumesParagraphLine: null,
            ManualBreakLineSpacingPoints: null,
            ManualBreakLineSpacingFactor: null);
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static DocxStructureBlockSnapshot FromPageBreak(int blockIndex, string? previousKind, string? nextKind, DocxPageBreakElement pageBreak)
    {
        return DocxStructureBlockSnapshot.ForPageBreak(blockIndex, previousKind, nextKind, pageBreak);
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static DocxStructureBlockSnapshot FromImplicitParagraph(int blockIndex, string? previousKind, string? nextKind, DocxImplicitParagraphElement implicitParagraph)
    {
        return DocxStructureBlockSnapshot.ForImplicitParagraph(blockIndex, previousKind, nextKind, implicitParagraph);
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static DocxStructureBlockSnapshot FromManualBreak(int blockIndex, string? previousKind, string? nextKind, DocxManualBreakElement manualBreak)
    {
        return DocxStructureBlockSnapshot.ForManualBreak(blockIndex, previousKind, nextKind, manualBreak);
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static DocxStructureBlockSnapshot FromSectionBreak(int blockIndex, string? previousKind, string? nextKind, DocxSectionBreakElement sectionBreak)
    {
        return DocxStructureBlockSnapshot.ForSectionBreak(blockIndex, previousKind, nextKind, CountRevisions(sectionBreak.Revisions), CountRevisions(sectionBreak.Revisions, DocxRevisionKind.Insertion), CountRevisions(sectionBreak.Revisions, DocxRevisionKind.Deletion), CountRevisions(sectionBreak.Revisions, DocxRevisionKind.MoveFrom), CountRevisions(sectionBreak.Revisions, DocxRevisionKind.MoveTo), CountOtherRevisions(sectionBreak.Revisions), sectionBreak.TypeValue, sectionBreak.ColumnCountValue, sectionBreak.ColumnEqualWidthValue, sectionBreak.ColumnSpaceValue, sectionBreak.ColumnDefinitions.Count, sectionBreak.ColumnDefinitions.Count(column => column.WidthValue is not null), sectionBreak.ColumnDefinitions.Count(column => column.SpaceValue is not null));
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static DocxStructureTableSnapshot ToTableSnapshot(DocxTable table, int tableIndex, int blockIndex)
    {
        int cellCount = table.Rows.Sum(row => row.Cells.Count);
        DocxParagraph[] tableParagraphs = DocxBlockTraversal.EnumerateTableParagraphs(table).ToArray();
        DocxRevisionInfo[] tableRevisions = EnumerateTableRevisions(table).ToArray();
        DocxTableCellBorder[] tableCellBorders = table.Rows.SelectMany(row => row.Cells.SelectMany(cell => cell.Borders)).ToArray();
        int paragraphCount = tableParagraphs.Length;
        return new DocxStructureTableSnapshot(
            tableIndex,
            blockIndex,
            table.StyleId,
            table.Rows.Count,
            MaxColumnCount(table),
            table.ColumnWidthsPoints.Count,
            table.ColumnWidthsPoints.Sum(),
            table.HasExplicitGrid,
            table.PreferredWidthPoints,
            table.PreferredWidthValue,
            table.PreferredWidthType,
            table.IndentPoints,
            table.IndentValue,
            table.IndentType,
            table.CellSpacingPoints,
            table.CellSpacingValue,
            table.CellSpacingType,
            table.LayoutValue,
            table.Rows.Count(row => row.IsHeader),
            table.Rows.Count(row => row.CantSplit),
            table.Rows.Count(row => row.HeightPoints is not null),
            table.Rows.Count(row => string.Equals(row.HeightRuleValue, "exact", StringComparison.OrdinalIgnoreCase)),
            table.Rows.Count(row => string.Equals(row.HeightRuleValue, "atLeast", StringComparison.OrdinalIgnoreCase)),
            table.Rows.Count(row => row.TablePropertyExceptionCellMargins is not null),
            cellCount,
            table.Rows.Sum(row => row.Cells.Count(cell => cell.GridSpan > 1)),
            table.Rows.Sum(row => row.Cells.Count(cell => cell.HasVerticalMerge)),
            table.Rows.Sum(row => row.Cells.Count(cell => string.Equals(cell.VerticalMergeValue, "restart", StringComparison.OrdinalIgnoreCase))),
            table.Rows.Sum(row => row.Cells.Count(cell => cell.FillHex is not null || cell.ShadingValue is not null)),
            table.Rows.Sum(row => row.Cells.Count(cell => cell.VerticalAlignmentValue is not null)),
            table.Rows.Sum(row => row.Cells.Count(cell => cell.PreferredWidthPoints is not null || cell.PreferredWidthValue is not null)),
            tableCellBorders.Count(IsVisibleBorder),
            tableCellBorders.Count(border => IsBorderStyle(border, "single")),
            tableCellBorders.Count(border => IsBorderStyle(border, "thick")),
            tableCellBorders.Count(border => IsBorderStyle(border, "double")),
            tableCellBorders.Count(border => IsBorderStyle(border, "dotted")),
            tableCellBorders.Count(IsDashedBorderStyle),
            tableCellBorders.Count(IsSuppressedBorder),
            tableCellBorders.Count(border => IsVisibleBorder(border) &&
                !IsBorderStyle(border, "single") &&
                !IsBorderStyle(border, "thick") &&
                !IsBorderStyle(border, "double") &&
                !IsBorderStyle(border, "triple") &&
                !IsBorderStyle(border, "dotted") &&
                !IsDashedBorderStyle(border) &&
                !IsCompoundBorderStyle(border) &&
                !IsThreeDBorderStyle(border) &&
                !IsWaveBorderStyle(border) &&
                !IsSupportedSolidBorderStyle(border)),
            paragraphCount,
            tableParagraphs.Sum(paragraph => paragraph.Runs.Count),
            tableParagraphs.Sum(paragraph => TextLength(paragraph)),
            CountRevisions(tableRevisions),
            CountRevisions(tableRevisions, DocxRevisionKind.Insertion),
            CountRevisions(tableRevisions, DocxRevisionKind.Deletion),
            CountRevisions(tableRevisions, DocxRevisionKind.MoveFrom),
            CountRevisions(tableRevisions, DocxRevisionKind.MoveTo),
            CountOtherRevisions(tableRevisions),
            tableParagraphs.Sum(paragraph => CountWhitespaceDelimitedTokens(paragraph)),
            tableParagraphs.Select(paragraph => LongestWhitespaceDelimitedTokenLength(paragraph)).DefaultIfEmpty(0).Max(),
            tableParagraphs.Sum(paragraph => paragraph.Images.Count),
            tableParagraphs.Sum(ParagraphInlineReferenceCount),
            tableParagraphs.Sum(ParagraphHyperlinkCount),
            tableParagraphs.Sum(ParagraphExternalHyperlinkCount),
            tableParagraphs.Sum(ParagraphInternalHyperlinkCount),
            tableParagraphs.Count(paragraph => paragraph.ListLabel is not null),
            tableParagraphs.Count(HasEffectiveKeepConstraint),
            table.Look?.FirstRow,
            table.Look?.FirstColumn,
            table.Look?.NoHorizontalBand,
            table.Look?.NoVerticalBand,
            table.Rows.Select((row, rowIndex) => ToTableRowSnapshot(row, rowIndex)).ToArray());
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static DocxStructureTableRowSnapshot ToTableRowSnapshot(DocxTableRow row, int rowIndex)
    {
        DocxStructureTableCellSnapshot[] cells = row.Cells.Select(ToTableCellSnapshot).ToArray();
        DocxRevisionInfo[] rowRevisions = row.Revisions
            .Concat(row.Cells.SelectMany(EnumerateTableCellRevisions))
            .ToArray();
        return new DocxStructureTableRowSnapshot(
            rowIndex,
            row.Cells.Count,
            row.Cells.Sum(cell => Math.Max(1, cell.GridSpan)),
            row.IsHeader,
            row.HeaderValue,
            row.CantSplit,
            row.CantSplitValue,
            row.HeightPoints,
            row.HeightValue,
            row.HeightRuleValue,
            row.TablePropertyExceptionCellMargins is not null,
            cells.Count(cell => cell.GridSpan > 1),
            cells.Count(cell => cell.HasVerticalMerge),
            cells.Count(cell => string.Equals(cell.VerticalMergeValue, "restart", StringComparison.OrdinalIgnoreCase)),
            cells.Count(cell => cell.HasShading),
            cells.Count(cell => cell.VerticalAlignmentValue is not null),
            cells.Count(cell => cell.HasPreferredWidth),
            cells.Sum(cell => cell.VisibleBorderCount),
            cells.Sum(cell => cell.ParagraphCount),
            cells.Sum(cell => cell.RunCount),
            cells.Sum(cell => cell.TextLength),
            CountRevisions(rowRevisions),
            CountRevisions(rowRevisions, DocxRevisionKind.Insertion),
            CountRevisions(rowRevisions, DocxRevisionKind.Deletion),
            CountRevisions(rowRevisions, DocxRevisionKind.MoveFrom),
            CountRevisions(rowRevisions, DocxRevisionKind.MoveTo),
            CountOtherRevisions(rowRevisions),
            cells.Sum(cell => cell.WhitespaceDelimitedTokenCount),
            cells.Select(cell => cell.LongestWhitespaceDelimitedTokenLength).DefaultIfEmpty(0).Max(),
            cells.Sum(cell => cell.InlineImageCount),
            cells.Sum(cell => cell.InlineReferenceCount),
            cells.Sum(cell => cell.HyperlinkCount),
            cells.Sum(cell => cell.ExternalHyperlinkCount),
            cells.Sum(cell => cell.InternalHyperlinkCount),
            cells.Sum(cell => cell.NumberedParagraphCount),
            cells.Sum(cell => cell.KeepRuleParagraphCount),
            cells.Sum(cell => cell.BeforeSpacingTokenParagraphCount),
            cells.Sum(cell => cell.AfterSpacingTokenParagraphCount),
            cells.Select(cell => cell.MaxFontSize).DefaultIfEmpty(0d).Max(),
            cells);
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static DocxStructureTableCellSnapshot ToTableCellSnapshot(DocxTableCell cell, int cellIndex)
    {
        DocxParagraph[] paragraphs = DocxBlockTraversal
            .EnumerateBodyParagraphs(DocxTableCellContent.GetBodyElements(cell))
            .ToArray();
        DocxRevisionInfo[] cellRevisions = cell.Revisions
            .Concat(paragraphs.SelectMany(paragraph => paragraph.Revisions))
            .ToArray();
        return new DocxStructureTableCellSnapshot(
            cellIndex,
            Math.Max(1, cell.GridSpan),
            cell.GridSpanValue,
            cell.HasVerticalMerge,
            cell.VerticalMergeValue,
            cell.FillHex is not null || cell.ShadingValue is not null,
            cell.ShadingValue,
            cell.VerticalAlignmentValue,
            cell.PreferredWidthPoints is not null || cell.PreferredWidthValue is not null,
            cell.PreferredWidthPoints,
            cell.PreferredWidthValue,
            cell.PreferredWidthType,
            cell.Borders.Count(border => !string.Equals(border.Value, "nil", StringComparison.OrdinalIgnoreCase) && !string.Equals(border.Value, "none", StringComparison.OrdinalIgnoreCase)),
            paragraphs.Length,
            paragraphs.Sum(paragraph => paragraph.Runs.Count),
            paragraphs.Sum(paragraph => TextLength(paragraph)),
            CountRevisions(cellRevisions),
            CountRevisions(cellRevisions, DocxRevisionKind.Insertion),
            CountRevisions(cellRevisions, DocxRevisionKind.Deletion),
            CountRevisions(cellRevisions, DocxRevisionKind.MoveFrom),
            CountRevisions(cellRevisions, DocxRevisionKind.MoveTo),
            CountOtherRevisions(cellRevisions),
            paragraphs.Sum(paragraph => CountWhitespaceDelimitedTokens(paragraph)),
            paragraphs.Select(paragraph => LongestWhitespaceDelimitedTokenLength(paragraph)).DefaultIfEmpty(0).Max(),
            paragraphs.Sum(paragraph => paragraph.Images.Count),
            paragraphs.Sum(ParagraphInlineReferenceCount),
            paragraphs.Sum(ParagraphHyperlinkCount),
            paragraphs.Sum(ParagraphExternalHyperlinkCount),
            paragraphs.Sum(ParagraphInternalHyperlinkCount),
            paragraphs.Count(paragraph => paragraph.ListLabel is not null),
            paragraphs.Count(HasEffectiveKeepConstraint),
            paragraphs.Count(paragraph => DocxParagraphSpacing.HasBeforeSpacingSide(paragraph.EffectiveProperties.Spacing)),
            paragraphs.Count(paragraph => DocxParagraphSpacing.HasAfterSpacingSide(paragraph.EffectiveProperties.Spacing)),
            paragraphs.SelectMany(paragraph => paragraph.Runs).Select(run => run.EffectiveProperties.FontSize).DefaultIfEmpty(0d).Max(),
            CountCharacters(paragraphs, static c => c == ' '),
            CountCharacters(paragraphs, static c => c > 127),
            CountCharacters(paragraphs, char.IsPunctuation),
            CountCharacters(paragraphs, char.IsDigit),
            CountCharacters(paragraphs, char.IsUpper),
            CountCharacters(paragraphs, char.IsLower),
            cell.BodyElements.Count,
            cell.BodyElements.OfType<DocxManualBreakElement>().Count(),
            cell.BodyElements.OfType<DocxPageBreakElement>().Count(),
            cell.BodyElements.OfType<DocxTableElement>().Count());
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static DocxStructureFloatingDrawingSnapshot ToFloatingDrawingSnapshot(DocxFloatingDrawing drawing, int index)
    {
        return new DocxStructureFloatingDrawingSnapshot(
            index,
            drawing.WrapKind?.ToValueString(),
            drawing.WrapTextValue,
            drawing.BehindDocumentValue,
            drawing.LayoutInCellValue,
            drawing.AllowOverlapValue,
            drawing.HorizontalRelativeFromValue,
            drawing.HorizontalAlignValue,
            drawing.HorizontalOffsetValue,
            drawing.VerticalRelativeFromValue,
            drawing.VerticalAlignValue,
            drawing.VerticalOffsetValue,
            drawing.ExtentCxValue,
            drawing.ExtentCyValue,
            drawing.DistanceTopValue,
            drawing.DistanceBottomValue,
            drawing.DistanceLeftValue,
            drawing.DistanceRightValue,
            drawing.SimplePositionValue,
            drawing.RelativeHeightValue,
            drawing.LockedValue,
            drawing.ImageRelationshipId,
            drawing.Image?.PartName,
            drawing.Image?.ContentType,
            drawing.Image?.WidthPoints,
            drawing.Image?.HeightPoints,
            drawing.SourceParagraphIndex,
            drawing.SourceBlockIndex,
            drawing.TextBoxBodyElements.Count,
            DocxBlockTraversal.EnumerateBodyParagraphs(drawing.TextBoxBodyElements).Count(),
            DocxBlockTraversal.EnumerateBodyParagraphs(drawing.TextBoxBodyElements).Sum(TextLength),
            CountRevisions(drawing.Revisions),
            CountRevisions(drawing.Revisions, DocxRevisionKind.Insertion),
            CountRevisions(drawing.Revisions, DocxRevisionKind.Deletion),
            CountRevisions(drawing.Revisions, DocxRevisionKind.MoveFrom),
            CountRevisions(drawing.Revisions, DocxRevisionKind.MoveTo),
            CountOtherRevisions(drawing.Revisions));
    }
}
