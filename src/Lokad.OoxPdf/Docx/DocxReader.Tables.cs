using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxReader
{
    private static DocxTable? ReadTable(
        DocxRevisionScopedElement source,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        DocxNumberingCounters numberingCounters,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxNoteReferenceCounters? inlineReferenceCounters,
        DocxDocumentSettings? documentSettings,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken,
        IReadOnlyList<DocxRevisionScopedElement>? fragments = null,
        bool projectCommentSelections = false)
    {
        cancellationToken.ThrowIfCancellationRequested();
        XElement table = source.Element;
        DocxRevisionInfo? inheritedRevision = fragments is null || fragments.All(fragment => fragment.Revision == source.Revision)
            ? source.Revision : null;
        XElement? tableProperties = table.Element(WordprocessingNamespace + "tblPr");
        string? layoutValue = (string?)tableProperties
            ?.Element(WordprocessingNamespace + "tblLayout")
            ?.Attribute(WordprocessingNamespace + "type");
        string? tableStyleId = (string?)tableProperties
            ?.Element(WordprocessingNamespace + "tblStyle")
            ?.Attribute(WordprocessingNamespace + "val");
        XElement? tableWidth = tableProperties
            ?.Element(WordprocessingNamespace + "tblW");
        XElement? tableIndent = tableProperties
            ?.Element(WordprocessingNamespace + "tblInd");
        XElement? tableCellSpacing = tableProperties
            ?.Element(WordprocessingNamespace + "tblCellSpacing");
        var tableRevisions = new List<DocxRevisionInfo>();
        AddRevision(tableRevisions, inheritedRevision);
        AddRevisions(tableRevisions, ReadPropertyChangeRevisions(tableProperties));
        DocxTableLook tableLook = ReadTableLook(tableProperties);
        DocxTableStyle? resolvedStyle = tableStyleId is not null && styles.TableStyles.TryGetValue(tableStyleId, out DocxTableStyle? parsedTableStyle)
            ? parsedTableStyle
            : styles.DefaultTableStyle;
        DocxTableStyle tableStyle = resolvedStyle ?? DocxTableStyle.Empty;
        bool isCustomStyle = resolvedStyle is { HasCompleteInheritance: true, UsesTableNormal: false };
        bool hasAuthoredIndent = tableIndent is not null || tableStyle.Table.IndentValue is not null;
        bool anchorsStyleGrid = isCustomStyle && !hasAuthoredIndent;
        // Office A/B (m13/ladder compat15 probes): a present compatibilityMode pins the
        // table grid at margin/indent plus the outer border half (legacy outer-edge
        // alignment); without it the grid sits at the origin (modern) or pins text via
        // style margins (w72/m12, see PinTextToMargin below). Any stated mode counts
        // (survey: only 15 occurs in the wild).
        bool useLegacyTableGrid = documentSettings?.HasCompatibilityMode == true;
        double? resolvedTableIndentPoints = tableIndent is not null ? ReadDxaWidth(tableIndent) : tableStyle.Table.IndentPoints;
        var parts = (fragments ?? [source]).Select(fragment => (
            Properties: fragment.Element.Element(WordprocessingNamespace + "tblPr"),
            Grid: ReadTableGrid(fragment.Element, cancellationToken),
            LeadingBookmarks: ReferenceEquals(fragment.Element, table) ? source.LeadingBookmarks : fragment.LeadingBookmarks,
            Rows: EnumerateBookmarkedBlocks(fragment.Element.Elements(), markupMode, cancellationToken,
                fragment.Revision == inheritedRevision ? null : fragment.Revision,
                tableChildName: WordprocessingNamespace + "tr", readTableColumnBookmarks: true)
                .Where(row => IsIncludedTableRow(row.Element, markupMode, cancellationToken)).ToArray())).ToArray();
        bool hasExplicitGrid = parts.Any(part => part.Grid.Length != 0);
        IReadOnlyList<double> columns = parts[0].Grid;
        int[][]? gridMappings = null;
        DocxTableSourceRow[][]? joinedRows = null;
        if (parts.Length > 1)
        {
            var grids = new double[parts.Length][];
            joinedRows = new DocxTableSourceRow[parts.Length][];
            for (int index = 0; index < parts.Length; index++)
            {
                (grids[index], joinedRows[index]) = CompleteJoinedTableGrid(parts[index].Grid, parts[index].Rows,
                    markupMode, cancellationToken);
            }
            var offsets = parts.Select(part => (ReadDxaWidth(part.Properties?.Element(WordprocessingNamespace + "tblInd"))
                ?? tableStyle.Table.IndentPoints ?? 0d) - (resolvedTableIndentPoints ?? 0d)).ToArray();
            var joined = JoinTableGrids(grids, offsets, cancellationToken);
            columns = joined.Columns;
            gridMappings = joined.Boundaries;
        }
        int rowCount = parts.Sum(part => part.Rows.Length);
        bool leadingRowEmphasis = true;
        foreach (var part in parts)
        {
            if (part.Rows.Length == 0) continue;
            leadingRowEmphasis = ReadRowTableLook(part.Rows[0].Element, part.Properties).FirstRow != false;
            break;
        }
        var rows = new List<DocxTableRow>();
        bool inHeaderGroup = true;
        int rowIndex = 0;
        for (int sourceIndex = 0; sourceIndex < parts.Length; sourceIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var part = parts[sourceIndex];
            var leadingBookmarks = part.LeadingBookmarks;
            DocxTableCellMargins partMargins = ReadTableStyleCellMargins(part.Properties);
            IReadOnlyList<DocxTableCellBorder> partBorders = ReadTableBorders(part.Properties);
            for (int partRowIndex = 0; partRowIndex < part.Rows.Length; partRowIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var rowElement = part.Rows[partRowIndex];
                XElement row = rowElement.Element;
                XElement? rowExceptions = row.Element(WordprocessingNamespace + "tblPrEx");
                DocxTableLook rowLook = ReadRowTableLook(row, part.Properties);
                IReadOnlyList<DocxTableCellBorder> rowTableBorders = ReadTableBorders(rowExceptions);
                XElement? rowProperties = row.Element(WordprocessingNamespace + "trPr");
                var rowRevisions = new List<DocxRevisionInfo>();
                AddRevision(rowRevisions, rowElement.Revision);
                AddRevisions(rowRevisions, ReadPropertyChangeRevisions(rowProperties));
                DocxTableCellMargins rowExceptionMargins = ReadTablePropertyExceptionCellMargins(row);
                DocxTableCellMargins rowInheritedMargins = rowExceptionMargins.Merge(partMargins.Merge(tableStyle.Cell.Margins));
                if (sourceIndex != 0) AddRevisions(rowRevisions, ReadPropertyChangeRevisions(part.Properties));
                var cells = new List<DocxTableCell>();
                DocxTableSourceRow sourceRow = joinedRows is null
                    ? ReadTableSourceRow(rowElement, rowProperties, markupMode, cancellationToken)
                    : joinedRows[sourceIndex][partRowIndex];
                int sourceColumn = sourceRow.FirstColumn;
                int gridBefore = gridMappings is null ? sourceColumn : gridMappings[sourceIndex][sourceColumn];
                var cellElements = sourceRow.Cells;
                ReadOnlyMemory<DocxTableColumnBookmark> columnBookmarks = default;
                var bookmarkEndBoundaries = sourceRow.BookmarkEndBoundaries;
                int followingCell = 0;
                // An end inside this cell can close an interval deferred to a
                // later one. A bookmark already placed in this cell keeps its target.
                Action<string> recordCellBookmarkEnd = id =>
                {
                    bookmarkEndBoundaries ??= new(StringComparer.Ordinal);
                    if (!bookmarkEndBoundaries.TryGetValue(id, out int boundary) || followingCell < boundary)
                        bookmarkEndBoundaries[id] = followingCell;
                };
                for (int cellIndex = 0; cellIndex < cellElements.Count; cellIndex++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var (cellSource, sourceSpan) = cellElements[cellIndex];
                    XElement cell = cellSource.Element;
                    DocxRevisionInfo? inheritedCellRevision = cellSource.Revision ?? rowElement.Revision;
                    XElement? cellProperties = cell.Element(WordprocessingNamespace + "tcPr");
                    int nextSourceColumn = sourceColumn + sourceSpan;
                    if (!columnBookmarks.IsEmpty)
                    {
                        var bookmarksForCell = BindColumnBookmarks(cellSource.LeadingBookmarks, ref columnBookmarks,
                            sourceColumn, nextSourceColumn, cancellationToken, bookmarkEndBoundaries, cellIndex);
                        if (!ReferenceEquals(bookmarksForCell, cellSource.LeadingBookmarks))
                            cellSource = cellSource with { LeadingBookmarks = bookmarksForCell };
                    }
                    var cellRevisions = new List<DocxRevisionInfo>();
                    AddRevision(cellRevisions, cellSource.Revision);
                    AddRevisions(cellRevisions, ReadPropertyChangeRevisions(cellProperties));
                    DocxTableCellConditionalFormat? conditionalFormat = ReadTableConditionalFormat(cellProperties);
                    DocxTableCellStyle conditionalStyle = ResolveTableCellStyle(tableStyle, rowLook, conditionalFormat, rowIndex, cellIndex, rowCount, cellElements.Count, leadingRowEmphasis);
                    followingCell = cellIndex + 1;
                    IReadOnlyList<DocxBodyElement> cellBodyElements = ReadTableCellBodyElements(
                        cellSource,
                        styles,
                        numbering,
                        numberingCounters,
                        package,
                        relationships,
                        conditionalStyle,
                        inlineReferenceCounters,
                        documentSettings,
                        markupMode,
                        cancellationToken,
                        inheritedCellRevision,
                        (sourceColumn, nextSourceColumn),
                        ref columnBookmarks,
                        recordCellBookmarkEnd,
                        projectCommentSelections);
                    IReadOnlyList<DocxParagraph> paragraphs = DocxBlockTraversal.EnumerateDirectParagraphs(cellBodyElements).ToArray();
                    XElement? shading = cellProperties?.Element(WordprocessingNamespace + "shd");
                    string? fill = (string?)shading?.Attribute(WordprocessingNamespace + "fill") ?? conditionalStyle.FillHex;
                    string? shadingValue = (string?)shading?.Attribute(WordprocessingNamespace + "val") ?? conditionalStyle.ShadingValue;
                    string? shadingColor = (string?)shading?.Attribute(WordprocessingNamespace + "color") ?? conditionalStyle.ShadingColor;
                    string? verticalAlignment = (string?)cellProperties
                        ?.Element(WordprocessingNamespace + "vAlign")
                        ?.Attribute(WordprocessingNamespace + "val")
                        ?? conditionalStyle.VerticalAlignmentValue;
                    XElement? cellWidth = cellProperties?.Element(WordprocessingNamespace + "tcW");
                    XElement? verticalMerge = cellProperties?.Element(WordprocessingNamespace + "vMerge");
                    XElement? noWrap = cellProperties?.Element(WordprocessingNamespace + "noWrap");
                    XElement? fitText = cellProperties?.Element(WordprocessingNamespace + "tcFitText");
                    string? textDirectionValue = cellProperties?.Element(WordprocessingNamespace + "textDirection") is { } textDirection
                        ? (string?)textDirection.Attribute(WordprocessingNamespace + "val")
                        : conditionalStyle.TextDirectionValue;
                    bool resolvedNoWrap = noWrap is not null
                        ? ReadOnOff(noWrap) == true
                        : conditionalStyle.NoWrap == true;
                    string? resolvedNoWrapValue = noWrap is not null
                        ? (string?)noWrap.Attribute(WordprocessingNamespace + "val")
                        : conditionalStyle.NoWrapValue;
                    bool resolvedFitText = fitText is not null
                        ? ReadOnOff(fitText) == true
                        : conditionalStyle.FitText == true;
                    string? resolvedFitTextValue = fitText is not null
                        ? (string?)fitText.Attribute(WordprocessingNamespace + "val")
                        : conditionalStyle.FitTextValue;
                    IReadOnlyList<DocxTableCellBorder> directBorders = ReadTableCellBorders(cellProperties);
                    IReadOnlyList<DocxTableCellBorder> borders = ResolveTableCellBorders(
                        directBorders,
                        conditionalStyle.Borders,
                        tableStyle.TableBorders,
                        partBorders,
                        rowTableBorders,
                        rowIndex,
                        cellIndex,
                        rowCount,
                        cellElements.Count);
                    DocxTableCellMargins inheritedMargins = rowInheritedMargins.Merge(conditionalStyle.Margins);
                    DocxTableCellMargins margins = ReadTableCellMargins(cellProperties).Merge(inheritedMargins);
                    DocxTableCellMargins styleMargins = tableStyle.Cell.Margins.Merge(conditionalStyle.Margins);
                    // Without compat, an intrinsic style pins text through the first
                    // cell's effective padding. Table, row and cell declarations replace
                    // style padding for both that grid anchor and the content offset.
                    bool pinTextToMargin = !isCustomStyle && !useLegacyTableGrid && (resolvedTableIndentPoints ?? 0d) == 0d && (styleMargins.LeftPoints ?? 0d) > 0d;
                    // Cell-local padding owns its content inset, not an inherited
                    // table text anchor. With neither inherited left margin nor
                    // authored indentation, retain the containing grid origin.
                    // Explicit zero indentation keeps its existing text anchor.
                    bool anchorsCellGrid = !hasAuthoredIndent &&
                        inheritedMargins.LeftPoints is null && inheritedMargins.LeftValue is null;
                    DocxTableCellMargins? gridAnchorMargins = anchorsStyleGrid || anchorsCellGrid
                        ? DocxTableCellMargins.Empty with { LeftPoints = 0d, RightPoints = 0d }
                        : pinTextToMargin ? margins : null;
                    int gridSpan = gridMappings is null ? sourceSpan : Math.Max(1,
                        gridMappings[sourceIndex][nextSourceColumn] - gridMappings[sourceIndex][sourceColumn]);
                    sourceColumn = nextSourceColumn;
                    cells.Add(new DocxTableCell(
                        paragraphs,
                        fill,
                        shadingValue,
                        shadingColor,
                        verticalAlignment,
                        borders,
                        margins,
                        ReadDxaWidth(cellWidth),
                        (string?)cellWidth?.Attribute(WordprocessingNamespace + "w"),
                        (string?)cellWidth?.Attribute(WordprocessingNamespace + "type"),
                        gridSpan,
                        (string?)cellProperties
                            ?.Element(WordprocessingNamespace + "gridSpan")
                            ?.Attribute(WordprocessingNamespace + "val"),
                        conditionalFormat,
                        verticalMerge is not null,
                        (string?)verticalMerge?.Attribute(WordprocessingNamespace + "val"),
                        resolvedNoWrap,
                        resolvedNoWrapValue,
                        resolvedFitText,
                        resolvedFitTextValue,
                        textDirectionValue,
                        gridAnchorMargins,
                        pinTextToMargin)
                    {
                        BodyElements = cellBodyElements,
                        Revisions = cellRevisions
                    });
                }

                if (cells.Count > 0)
                {
                    DocxTableRowStyle rowStyle = ResolveTableRowStyle(tableStyle, rowLook, rowProperties, rowIndex, rowCount, leadingRowEmphasis);
                    // Only the contiguous leading group participates in header layout.
                    inHeaderGroup &= rowStyle.RepeatHeader?.Enabled == true;
                    XElement? rowHeight = rowProperties?.Element(WordprocessingNamespace + "trHeight");
                    DocxTableCellMargins effectiveExceptions = sourceIndex == 0 ? rowExceptionMargins : rowExceptionMargins.Merge(partMargins);
                    var bookmarks = leadingBookmarks ?? rowElement.LeadingBookmarks;
                    if (leadingBookmarks is not null && rowElement.LeadingBookmarks is { } rowBookmarks)
                        bookmarks = [.. leadingBookmarks, .. rowBookmarks];
                    rows.Add(new DocxTableRow(
                        cells,
                        ReadTableRowHeight(rowHeight),
                        inHeaderGroup,
                        rowStyle.RepeatHeader?.Value,
                        (string?)rowHeight?.Attribute(WordprocessingNamespace + "val"),
                        (string?)rowHeight?.Attribute(WordprocessingNamespace + "hRule"),
                        HasAnyTableCellMargin(effectiveExceptions) ? effectiveExceptions : null,
                        rowStyle.CantSplit?.Enabled == true,
                        rowStyle.CantSplit?.Value)
                    {
                        Revisions = rowRevisions,
                        GridBefore = gridBefore,
                        BookmarkAnchors = bookmarks
                    });
                    leadingBookmarks = null;
                }
                rowIndex++;
            }
        }

        if (rows.Count == 0)
        {
            return null;
        }

        if (columns.Count == 0)
        {
            int inferredGridColumns = 0;
            foreach (DocxTableRow row in rows)
            {
                long rowTotal = row.GridBefore;
                foreach (DocxTableCell cell in row.Cells)
                {
                    rowTotal = checked(rowTotal + Math.Max(1, cell.GridSpan));
                    if (rowTotal > MaxInferredGridColumns)
                    {
                        throw new OoxPdfLimitExceededException(
                            "DOCX table inferred grid exceeds the maximum supported column count of " + MaxInferredGridColumns + ".");
                    }
                }

                inferredGridColumns = Math.Max(inferredGridColumns, (int)rowTotal);
            }

            if (inferredGridColumns > MaxInferredGridColumns)
            {
                throw new OoxPdfLimitExceededException(
                    "DOCX table inferred grid exceeds the maximum supported column count of " + MaxInferredGridColumns + ".");
            }

            columns = Enumerable.Repeat(72d, inferredGridColumns).ToArray();
        }

        return new DocxTable(
            layoutValue ?? tableStyle.Table.LayoutValue,
            columns,
            rows,
            tableStyleId,
            tableWidth is not null ? ReadDxaWidth(tableWidth) : tableStyle.Table.PreferredWidthPoints,
            tableWidth is not null ? (string?)tableWidth.Attribute(WordprocessingNamespace + "w") : tableStyle.Table.PreferredWidthValue,
            tableWidth is not null ? (string?)tableWidth.Attribute(WordprocessingNamespace + "type") : tableStyle.Table.PreferredWidthType,
            resolvedTableIndentPoints,
            tableIndent is not null ? (string?)tableIndent.Attribute(WordprocessingNamespace + "w") : tableStyle.Table.IndentValue,
            tableIndent is not null ? (string?)tableIndent.Attribute(WordprocessingNamespace + "type") : tableStyle.Table.IndentType,
            tableCellSpacing is not null ? ReadDxaWidth(tableCellSpacing) : tableStyle.Table.CellSpacingPoints,
            tableCellSpacing is not null ? (string?)tableCellSpacing.Attribute(WordprocessingNamespace + "w") : tableStyle.Table.CellSpacingValue,
            tableCellSpacing is not null ? (string?)tableCellSpacing.Attribute(WordprocessingNamespace + "type") : tableStyle.Table.CellSpacingType,
            tableLook,
            hasExplicitGrid,
            useLegacyTableGrid)
        {
            Revisions = tableRevisions
        };
    }

    private static bool IsIncludedTableRow(XElement row, OoxPdfDocxMarkupMode mode, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        if (row.Element(WordprocessingNamespace + "trPr") is not { } properties) return true;
        foreach (XElement revision in properties.Elements())
        {
            token.ThrowIfCancellationRequested();
            if ((revision.Name == WordprocessingNamespace + "ins" || revision.Name == WordprocessingNamespace + "del") &&
                !IsIncludedRevisionContainer(revision, mode)) return false;
        }
        return true;
    }

    private static double? ReadDxaWidth(XElement? width)
    {
        string? type = (string?)width?.Attribute(WordprocessingNamespace + "type");
        if (type is not null && !type.Equals("dxa", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        return ReadTwipsAttribute(width, WordprocessingNamespace + "w");
    }

    internal const int MaxTableGridSpan = 1024;

    internal const int MaxInferredGridColumns = 1024;

    private static int ReadGridSpan(XElement? cellProperties)
    {
        if (cellProperties
            ?.Element(WordprocessingNamespace + "gridSpan")
            ?.Attribute(WordprocessingNamespace + "val") is not { } span)
        {
            return 1;
        }

        // Malformed spans fall back to 1 to preserve layout for Office-tolerated
        // content; oversized spans are rejected before column allocation (M03).
        if (!int.TryParse(span.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed))
        {
            return 1;
        }

        int normalized = Math.Max(1, parsed);
        if (normalized > MaxTableGridSpan)
        {
            throw new OoxPdfLimitExceededException(
                "DOCX table gridSpan exceeds the maximum supported column span of " + MaxTableGridSpan + ".");
        }

        return normalized;
    }

    private static IReadOnlyList<DocxBookmarkAnchor>? BindColumnBookmarks(
        IReadOnlyList<DocxBookmarkAnchor>? leadingBookmarks, ref ReadOnlyMemory<DocxTableColumnBookmark> columnBookmarks,
        int startColumn, int endColumn, CancellationToken token,
        IReadOnlyDictionary<string, int>? endBoundaries = null, int cellIndex = 0)
    {
        List<DocxBookmarkAnchor>? result = null;
        int consumed = 0;
        foreach (var bookmark in columnBookmarks.Span)
        {
            token.ThrowIfCancellationRequested();
            if (bookmark.First >= endColumn) break;
            consumed++;
            // An omitted leading grid column has no content. Use the first
            // intersecting authored cell, including a cell that spans the start.
            if (bookmark.Last < startColumn) continue;
            if (endBoundaries is not null &&
                endBoundaries.TryGetValue(OoxXml.NormalizeDecimalId(bookmark.Anchor.Id)!, out int beforeCell) &&
                beforeCell <= cellIndex) continue;
            result ??= leadingBookmarks is null ? [] : [.. leadingBookmarks];
            result.Add(bookmark.Anchor);
        }
        columnBookmarks = columnBookmarks[consumed..];
        return result ?? leadingBookmarks;
    }

    private readonly record struct DocxCellBookmarkScope(int EndColumn, Action<DocxTableColumnBookmark> Defer,
        Action<string> RecordEnd);

    private static IReadOnlyList<DocxBodyElement> ReadTableCellBodyElements(
        DocxRevisionScopedElement cell,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        DocxNumberingCounters numberingCounters,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxTableCellStyle tableCellStyle,
        DocxNoteReferenceCounters? inlineReferenceCounters,
        DocxDocumentSettings? documentSettings,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken,
        DocxRevisionInfo? inheritedRevision,
        (int Start, int End) cellColumns,
        ref ReadOnlyMemory<DocxTableColumnBookmark> columnBookmarks,
        Action<string> recordBookmarkEnd,
        bool projectCommentSelections)
    {
        var elements = new List<DocxBodyElement>();
        var flowFields = new List<DocxComplexFieldState>();
        var comments = projectCommentSelections ? new DocxCommentProjection(elements) : null;
        bool mergeDeletedMarks = !DocxMarkupContext.FromMode(markupMode).IncludesDeletions;
        List<DocxTableColumnBookmark>? deferred = null;
        DocxTableColumnBookmark[]? trailing = null;
        void Defer(DocxTableColumnBookmark bookmark) => (deferred ??= []).Add(bookmark);
        var bookmarkScope = new DocxCellBookmarkScope(cellColumns.End, Defer, recordBookmarkEnd);
        foreach (var (block, tableFragments) in EnumerateLogicalTableBlocks(cell.Element.Elements(), styles, markupMode,
            cancellationToken, inheritedRevision, cell.LeadingBookmarks, readTableColumnBookmarks: true,
            leadingColumnBookmarks: cell.LeadingColumnBookmarks, recordBookmarkEnd: recordBookmarkEnd,
            recordTrailingColumnBookmarks: starts => trailing = starts, readCommentBoundaries: projectCommentSelections))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var scopedChild = block;
            // Resolve starts that cover this block and carry intervals to the right
            // forward within the source row. Their ends may occur inside this cell.
            if (block.LeadingColumnBookmarks is { } starts)
            {
                ReadOnlyMemory<DocxTableColumnBookmark> pending = starts;
                var cellBookmarks = BindColumnBookmarks(block.LeadingBookmarks, ref pending,
                    cellColumns.Start, cellColumns.End, cancellationToken);
                if (!ReferenceEquals(cellBookmarks, block.LeadingBookmarks))
                    scopedChild = block with { LeadingBookmarks = cellBookmarks };
                foreach (var bookmark in pending.Span)
                    Defer(bookmark);
            }
            XElement child = scopedChild.Element;
            DocxRevisionInfo? childRevision = scopedChild.Revision ?? inheritedRevision;
            if (child.Name == WordprocessingNamespace + "tcPr")
            {
                continue;
            }

            if (child.Name == WordprocessingNamespace + "p")
            {
                var leadingBookmarks = scopedChild.LeadingBookmarks;
                bool HiddenBreak(XElement source) => IsHiddenHardRunBreak(source, styles, tableCellStyle);
                bool VisibleBreak(XElement source) => IsPageOrColumnBreak(source) && !HiddenBreak(source);
                bool VisiblePageBreak(XElement source) => IsPageBreak(source) && VisibleBreak(source);
                // Inline page/column tokens are transparent inside a cell. Keep
                // the source paragraph intact before preparation. A standalone
                // page-break paragraph retains the existing spill/flow handling.
                bool keepsPageBreak = IsRunBreakOnlyParagraph(child, VisibleBreak, HiddenBreak, markupMode) &&
                    child.Elements().Any(element => HasVisibleRunBreak(element, VisiblePageBreak, markupMode));
                var inlineCommentBoundaries = comments?.BeginParagraph(scopedChild.LeadingCommentBoundaries,
                    FindMergePredecessor(elements, mergeDeletedMarks) is not null, cancellationToken);
                if (keepsPageBreak)
                {
                    XElement? paragraphProperties = child.Element(WordprocessingNamespace + "pPr");
                    string? paragraphStyleId = ReadEffectiveParagraphStyleId(paragraphProperties, styles);
                    var resolvedParagraph = ResolveParagraphProperties(ReadParagraphProperties(paragraphProperties), paragraphStyleId, styles, tableCellStyle.Paragraph);
                    var sourceNumbering = ResolveParagraphNumbering(paragraphProperties, resolvedParagraph, paragraphStyleId, styles, numbering, cancellationToken);
                    bool hasListMarker = sourceNumbering is { } source && !string.IsNullOrWhiteSpace(source.Definition.Text);
                    // Put a consumed boundary before the break, exactly as an
                    // authored inline marker in the fragment that merges.
                    foreach (ParagraphBreakPart part in SplitParagraphAtRunBreaks(child, markupMode, VisiblePageBreak,
                        HiddenBreak, cancellationToken, retainFirstLine: hasListMarker, leadingCommentBoundaries: inlineCommentBoundaries))
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (part.BreakValue is not null)
                        {
                            elements.Add(DocxBodyElementFactory.CreatePageBreak(DocxBreakSourceKind.RunBreak, part.BreakValue, null, null));
                            continue;
                        }

                        if (part.Paragraph is null)
                        {
                            continue;
                        }

                        DocxParagraph? splitParagraph = ReadParagraph(
                            part.Paragraph,
                            styles,
                            numbering,
                            numberingCounters,
                            package,
                            relationships,
                            tableCellStyle,
                            inlineReferenceCounters: inlineReferenceCounters,
                            documentSettings: documentSettings,
                            inheritedRevision: childRevision,
                            markupMode: markupMode,
                            cancellationToken: cancellationToken,
                            isContinuationFragment: !part.IsFirstFragment,
                            leadingBookmarks: leadingBookmarks,
                            cellBookmarks: bookmarkScope,
                            flowFields: flowFields,
                            retainUnmatchedCommentStarts: projectCommentSelections);
                        leadingBookmarks = null;
                        if (splitParagraph is not null)
                        {
                            AppendFlowParagraph(elements, AdjustBreakParagraphFragment(splitParagraph, part, sourceNumbering?.Definition.Indent),
                                mergeDeletedMarks, flowFields, cancellationToken, comments);
                        }
                    }

                    continue;
                }

                DocxParagraph? parsed = ReadParagraph(
                    child,
                    styles,
                    numbering,
                    numberingCounters,
                    package,
                    relationships,
                    tableCellStyle,
                    inlineReferenceCounters: inlineReferenceCounters,
                    documentSettings: documentSettings,
                    inheritedRevision: childRevision,
                    markupMode: markupMode,
                    cancellationToken: cancellationToken,
                    leadingBookmarks: leadingBookmarks,
                    cellBookmarks: bookmarkScope,
                    flowFields: flowFields,
                    retainUnmatchedCommentStarts: projectCommentSelections,
                    leadingInlineCommentBoundaries: inlineCommentBoundaries);
                if (parsed is not null)
                {
                    AppendFlowParagraph(elements, parsed, mergeDeletedMarks, flowFields, cancellationToken, comments);
                }
            }
            else
            {
                flowFields.Clear();
                if (child.Name == WordprocessingNamespace + "tbl")
                {
                    DocxTable? nestedTable = ReadTable(scopedChild, styles, numbering, numberingCounters, package, relationships,
                        inlineReferenceCounters, documentSettings, markupMode, cancellationToken, tableFragments, projectCommentSelections);
                    if (nestedTable is not null)
                    {
                        elements.Add(DocxBodyElementFactory.CreateTable(nestedTable));
                    }
                }
                comments?.FinishGroup(cancellationToken);
            }
        }
        comments?.FinishGroup(cancellationToken);

        // A trailing start has no following block in this cell. Later cells can
        // still intersect its interval, subject to the same end-marker checks.
        if (trailing is not null)
        {
            foreach (var bookmark in trailing)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (bookmark.First >= cellColumns.End)
                    Defer(bookmark);
            }
        }
        if (deferred is not null)
        {
            DocxTableColumnBookmark[] combined = [.. columnBookmarks.Span, .. deferred];
            columnBookmarks = combined.OrderBy(bookmark => bookmark.First).ToArray();
        }
        return elements;
    }


    private static DocxTableCellMargins ReadTableCellMargins(XElement? cellProperties)
    {
        XElement? margins = cellProperties?.Element(WordprocessingNamespace + "tcMar");
        return ReadCellMargins(margins);
    }

    private static DocxTableCellMargins ReadTableStyleCellMargins(XElement? tableProperties)
    {
        XElement? margins = tableProperties?.Element(WordprocessingNamespace + "tblCellMar");
        return ReadCellMargins(margins);
    }

    private static DocxTableCellMargins ReadTablePropertyExceptionCellMargins(XElement row)
    {
        XElement? margins = row
            .Element(WordprocessingNamespace + "tblPrEx")
            ?.Element(WordprocessingNamespace + "tblCellMar");
        return ReadCellMargins(margins);
    }

    private static DocxTableCellMargins ReadCellMargins(XElement? margins)
    {
        XElement? top = margins?.Element(WordprocessingNamespace + "top");
        XElement? bottom = margins?.Element(WordprocessingNamespace + "bottom");
        // Resolve logical edges once for the supported left-to-right table layout.
        XElement? left = margins?.Element(WordprocessingNamespace + "start")
            ?? margins?.Element(WordprocessingNamespace + "left");
        XElement? right = margins?.Element(WordprocessingNamespace + "end")
            ?? margins?.Element(WordprocessingNamespace + "right");
        return new DocxTableCellMargins(
            ReadCellMarginPoints(top),
            ReadCellMarginPoints(right),
            ReadCellMarginPoints(bottom),
            ReadCellMarginPoints(left),
            (string?)top?.Attribute(WordprocessingNamespace + "w"),
            (string?)right?.Attribute(WordprocessingNamespace + "w"),
            (string?)bottom?.Attribute(WordprocessingNamespace + "w"),
            (string?)left?.Attribute(WordprocessingNamespace + "w"));
    }

    private static bool HasAnyTableCellMargin(DocxTableCellMargins margins)
    {
        return margins.TopValue is not null ||
            margins.RightValue is not null ||
            margins.BottomValue is not null ||
            margins.LeftValue is not null;
    }

    private static double? ReadCellMarginPoints(XElement? margin)
    {
        string? type = (string?)margin?.Attribute(WordprocessingNamespace + "type");
        if (type is not null && !type.Equals("dxa", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        return ReadTwipsAttribute(margin, WordprocessingNamespace + "w");
    }


    private static IReadOnlyList<DocxTableCellBorder> ReadTableCellBorders(XElement? cellProperties)
    {
        return ReadBorderElements(cellProperties?.Element(WordprocessingNamespace + "tcBorders"));
    }

    private static IReadOnlyList<DocxTableCellBorder> ReadTableBorders(XElement? tableProperties)
    {
        return ReadBorderElements(tableProperties?.Element(WordprocessingNamespace + "tblBorders"));
    }

    private static DocxTableLook ReadTableLook(XElement? tableProperties)
    {
        XElement? look = tableProperties?.Element(WordprocessingNamespace + "tblLook");
        if (look is null)
        {
            return DocxTableLook.Empty;
        }

        var parsed = new DocxTableLook(
            (string?)look.Attribute(WordprocessingNamespace + "val"),
            OoxBoolean.ParseOptionalAttribute(look, WordprocessingNamespace + "firstRow"),
            (string?)look.Attribute(WordprocessingNamespace + "firstRow"),
            OoxBoolean.ParseOptionalAttribute(look, WordprocessingNamespace + "lastRow"),
            (string?)look.Attribute(WordprocessingNamespace + "lastRow"),
            OoxBoolean.ParseOptionalAttribute(look, WordprocessingNamespace + "firstColumn"),
            (string?)look.Attribute(WordprocessingNamespace + "firstColumn"),
            OoxBoolean.ParseOptionalAttribute(look, WordprocessingNamespace + "lastColumn"),
            (string?)look.Attribute(WordprocessingNamespace + "lastColumn"),
            OoxBoolean.ParseOptionalAttribute(look, WordprocessingNamespace + "noHBand"),
            (string?)look.Attribute(WordprocessingNamespace + "noHBand"),
            OoxBoolean.ParseOptionalAttribute(look, WordprocessingNamespace + "noVBand"),
            (string?)look.Attribute(WordprocessingNamespace + "noVBand"));

        // Word ignores the bitmask when any explicit flag is present. Missing
        // flags in that declaration are enabled, including band suppression.
        if (parsed.FirstRow is not null || parsed.LastRow is not null ||
            parsed.FirstColumn is not null || parsed.LastColumn is not null ||
            parsed.NoHorizontalBand is not null || parsed.NoVerticalBand is not null)
        {
            return parsed with
            {
                FirstRow = parsed.FirstRow ?? true,
                LastRow = parsed.LastRow ?? true,
                FirstColumn = parsed.FirstColumn ?? true,
                LastColumn = parsed.LastColumn ?? true,
                NoHorizontalBand = parsed.NoHorizontalBand ?? true,
                NoVerticalBand = parsed.NoVerticalBand ?? true
            };
        }
        if (ushort.TryParse(parsed.Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out ushort mask))
        {
            return parsed with
            {
                FirstRow = (mask & 0x0020) != 0,
                LastRow = (mask & 0x0040) != 0,
                FirstColumn = (mask & 0x0080) != 0,
                LastColumn = (mask & 0x0100) != 0,
                NoHorizontalBand = (mask & 0x0200) != 0,
                NoVerticalBand = (mask & 0x0400) != 0
            };
        }

        // Empty or malformed declarations retain the existing absent-look fallback.
        return parsed;
    }

    private static DocxTableCellConditionalFormat? ReadTableConditionalFormat(XElement? properties)
    {
        XElement? conditional = properties?.Element(WordprocessingNamespace + "cnfStyle");
        if (conditional is null)
        {
            return null;
        }

        return new DocxTableCellConditionalFormat(
            (string?)conditional.Attribute(WordprocessingNamespace + "val"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "firstRow"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "lastRow"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "firstColumn"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "lastColumn"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "oddHBand"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "evenHBand"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "oddVBand"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "evenVBand"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "firstRowFirstColumn"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "firstRowLastColumn"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "lastRowFirstColumn"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "lastRowLastColumn"));
    }

    private static IReadOnlyList<DocxTableCellBorder> ReadBorderElements(XElement? borders)
    {
        if (borders is null)
        {
            return [];
        }

        return borders
            .Elements()
            .Where(border => border.Name.Namespace == WordprocessingNamespace)
            .Select(border => new DocxTableCellBorder(
                border.Name.LocalName,
                (string?)border.Attribute(WordprocessingNamespace + "val"),
                (string?)border.Attribute(WordprocessingNamespace + "color"),
                (string?)border.Attribute(WordprocessingNamespace + "sz")))
            .ToArray();
    }

    private static IReadOnlyList<DocxTableCellBorder> ResolveTableCellBorders(
        IReadOnlyList<DocxTableCellBorder> directBorders,
        IReadOnlyList<DocxTableCellBorder> styleBorders,
        IReadOnlyList<DocxTableCellBorder> styleTableBorders,
        IReadOnlyList<DocxTableCellBorder> tableBorders,
        IReadOnlyList<DocxTableCellBorder> rowTableBorders,
        int rowIndex,
        int cellIndex,
        int rowCount,
        int cellCount)
    {
        var resolved = new Dictionary<string, DocxTableCellBorder>(StringComparer.OrdinalIgnoreCase);
        AddBorders(resolved, ResolveTableBordersForCell(styleTableBorders, rowIndex, cellIndex, rowCount, cellCount));
        AddBorders(resolved, styleBorders);
        AddBorders(resolved, ResolveTableBordersForCell(tableBorders, rowIndex, cellIndex, rowCount, cellCount));
        AddBorders(resolved, ResolveTableBordersForCell(rowTableBorders, rowIndex, cellIndex, rowCount, cellCount));
        AddBorders(resolved, directBorders);
        return new[] { "top", "bottom", "left", "right" }
            .Where(resolved.ContainsKey)
            .Select(edge => resolved[edge])
            .Concat(resolved.Values.Where(border => !IsCanonicalCellBorderEdge(border.Edge)))
            .ToArray();
    }

    private static IEnumerable<DocxTableCellBorder> ResolveTableBordersForCell(
        IReadOnlyList<DocxTableCellBorder> tableBorders,
        int rowIndex,
        int cellIndex,
        int rowCount,
        int cellCount)
    {
        DocxTableCellBorder? top = DocxTableBorderGeometry.Find(tableBorders, rowIndex == 0 ? "top" : "insideH");
        if (top is not null)
        {
            yield return top with { Edge = "top" };
        }

        DocxTableCellBorder? bottom = DocxTableBorderGeometry.Find(tableBorders, rowIndex == rowCount - 1 ? "bottom" : "insideH");
        if (bottom is not null)
        {
            yield return bottom with { Edge = "bottom" };
        }

        DocxTableCellBorder? left = cellIndex == 0
            ? DocxTableBorderGeometry.Find(tableBorders, "left") ?? DocxTableBorderGeometry.Find(tableBorders, "start")
            : DocxTableBorderGeometry.Find(tableBorders, "insideV");
        if (left is not null)
        {
            yield return left with { Edge = "left" };
        }

        DocxTableCellBorder? right = cellIndex == cellCount - 1
            ? DocxTableBorderGeometry.Find(tableBorders, "right") ?? DocxTableBorderGeometry.Find(tableBorders, "end")
            : DocxTableBorderGeometry.Find(tableBorders, "insideV");
        if (right is not null)
        {
            yield return right with { Edge = "right" };
        }
    }

    private static void AddBorders(Dictionary<string, DocxTableCellBorder> target, IEnumerable<DocxTableCellBorder> borders)
    {
        foreach (DocxTableCellBorder border in borders)
        {
            target[border.Edge] = border;
        }
    }

    private static bool IsCanonicalCellBorderEdge(string edge)
    {
        return edge.Equals("top", StringComparison.OrdinalIgnoreCase) ||
            edge.Equals("bottom", StringComparison.OrdinalIgnoreCase) ||
            edge.Equals("left", StringComparison.OrdinalIgnoreCase) ||
            edge.Equals("right", StringComparison.OrdinalIgnoreCase);
    }

    private static double? ReadTableRowHeight(XElement? height)
    {
        if (height?.Attribute(WordprocessingNamespace + "val") is not { } value)
        {
            return null;
        }

        try
        {
            return OoxUnits.TwipsToPoints(long.Parse(value.Value, CultureInfo.InvariantCulture));
        }
        catch (Exception ex) when (ex is FormatException or OverflowException)
        {
            throw new InvalidDataException("Malformed DOCX table row height.", ex);
        }
    }
}
