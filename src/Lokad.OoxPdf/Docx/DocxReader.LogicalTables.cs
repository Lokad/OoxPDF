using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxReader
{
    private static readonly XName[] ParagraphFlowBlocks =
        [WordprocessingNamespace + "p", WordprocessingNamespace + "tbl", WordprocessingNamespace + "altChunk"];

    // Keep original XML/source scopes: grouping must not move or clone paragraphs,
    // drawings, references or revision owners. Only joined tables need a list.
    private static IEnumerable<(DocxRevisionScopedElement Scope, IReadOnlyList<DocxRevisionScopedElement>? Tables)>
        EnumerateLogicalTableBlocks(IEnumerable<XElement> elements, DocxStyleSet styles,
            OoxPdfDocxMarkupMode mode, CancellationToken token, DocxRevisionInfo? inheritedRevision = null,
            IReadOnlyList<DocxBookmarkAnchor>? leadingBookmarks = null, bool readTableColumnBookmarks = false,
            DocxTableColumnBookmark[]? leadingColumnBookmarks = null, Action<string>? recordBookmarkEnd = null,
            Action<DocxTableColumnBookmark[]>? recordTrailingColumnBookmarks = null, bool readCommentBoundaries = false)
    {
        DocxRevisionScopedElement? pending = null;
        List<DocxRevisionScopedElement>? tables = null;
        foreach (var block in EnumerateBookmarkedBlocks(elements, mode, token, inheritedRevision,
            leadingBookmarks: leadingBookmarks, readTableColumnBookmarks: readTableColumnBookmarks,
            leadingColumnBookmarks: leadingColumnBookmarks, recordBookmarkEnd: recordBookmarkEnd,
            recordTrailingColumnBookmarks: recordTrailingColumnBookmarks, readCommentBoundaries: readCommentBoundaries))
        {
            if (pending is not null && CanJoinTables(pending.Element, block.Element, styles))
            {
                tables ??= [pending];
                tables.Add(block);
                continue;
            }
            if (pending is not null) yield return (pending, tables);
            pending = block;
            tables = null;
        }
        if (pending is not null) yield return (pending, tables);
    }

    private static IEnumerable<DocxRevisionScopedElement> EnumerateBookmarkedBlocks(IEnumerable<XElement> elements,
        OoxPdfDocxMarkupMode mode, CancellationToken token, DocxRevisionInfo? inheritedRevision,
        XName? tableChildName = null, IReadOnlyList<DocxBookmarkAnchor>? leadingBookmarks = null,
        bool readTableColumnBookmarks = false, DocxTableColumnBookmark[]? leadingColumnBookmarks = null,
        Action<string>? recordBookmarkEnd = null, Action<DocxTableColumnBookmark[]>? recordTrailingColumnBookmarks = null,
        bool readCommentBoundaries = false)
    {
        XName importedContentName = WordprocessingNamespace + "altChunk";
        XName bookmarkStartName = WordprocessingNamespace + "bookmarkStart";
        XName bookmarkEndName = WordprocessingNamespace + "bookmarkEnd";
        XName commentStartName = WordprocessingNamespace + "commentRangeStart";
        XName commentEndName = WordprocessingNamespace + "commentRangeEnd";
        XName[] includedNames = tableChildName is not null
            ? [tableChildName, bookmarkStartName, bookmarkEndName]
            : [.. ParagraphFlowBlocks, bookmarkStartName, bookmarkEndName];
        if (readCommentBoundaries)
            includedNames = [.. includedNames, commentStartName, commentEndName];
        List<XElement>? commentBoundaries = null;
        // A cell inherits starts at its row boundary before choosing its first block.
        Dictionary<string, DocxBookmarkAnchor>? bookmarkStarts = null;
        if (leadingBookmarks is { Count: > 0 })
        {
            bookmarkStarts = new(StringComparer.Ordinal);
            foreach (var bookmark in leadingBookmarks)
            {
                token.ThrowIfCancellationRequested();
                bookmarkStarts.TryAdd(OoxXml.NormalizeDecimalId(bookmark.Id)!, bookmark);
            }
        }
        Dictionary<string, DocxTableColumnBookmark>? columnStarts = null;
        if (leadingColumnBookmarks is { Length: > 0 })
        {
            columnStarts = new(StringComparer.Ordinal);
            foreach (var bookmark in leadingColumnBookmarks)
            {
                token.ThrowIfCancellationRequested();
                columnStarts.TryAdd(OoxXml.NormalizeDecimalId(bookmark.Anchor.Id)!, bookmark);
            }
        }
        foreach (var scope in EnumerateRevisionScopedChildren(elements, mode, token, includedNames, inheritedRevision))
        {
            token.ThrowIfCancellationRequested();
            if (scope.Element.Name == commentStartName || scope.Element.Name == commentEndName)
            {
                (commentBoundaries ??= []).Add(scope.Element);
                continue;
            }
            if (scope.Element.Name == bookmarkStartName)
            {
                var anchor = ReadBookmarkAnchor(scope.Element, 0, 0, 0, 0);
                if (OoxXml.NormalizeDecimalId(anchor.Id) is not { Length: > 0 } id || string.IsNullOrEmpty(anchor.Name))
                    continue;
                XName firstColumn = WordprocessingNamespace + "colFirst";
                XName lastColumn = WordprocessingNamespace + "colLast";
                if (readTableColumnBookmarks &&
                    (scope.Element.Attribute(firstColumn) is not null || scope.Element.Attribute(lastColumn) is not null))
                {
                    if (ReadTableColumnBookmark(scope.Element, anchor) is { } bookmark)
                        (columnStarts ??= new(StringComparer.Ordinal)).TryAdd(id, bookmark);
                    continue;
                }
                (bookmarkStarts ??= new(StringComparer.Ordinal)).TryAdd(id, anchor);
                continue;
            }
            if (scope.Element.Name == bookmarkEndName)
            {
                if ((bookmarkStarts is not null || columnStarts is not null || recordBookmarkEnd is not null) &&
                    OoxXml.NormalizeDecimalId((string?)scope.Element.Attribute(WordprocessingNamespace + "id")) is { } id)
                {
                    bookmarkStarts?.Remove(id);
                    columnStarts?.Remove(id);
                    recordBookmarkEnd?.Invoke(id);
                }
                continue;
            }

            // Bind the start before layout, including rows filtered by the selected view.
            var block = (bookmarkStarts is { Count: > 0 } || columnStarts is { Count: > 0 }) && scope.Element.Name != importedContentName
                ? scope with {
                    LeadingBookmarks = bookmarkStarts is { Count: > 0 } ? bookmarkStarts.Values.ToArray() : null,
                    LeadingColumnBookmarks = columnStarts is { Count: > 0 }
                        ? columnStarts.Values.OrderBy(bookmark => bookmark.First).ToArray() : null
                }
                : scope;
            bookmarkStarts = null;
            columnStarts = null;
            if (commentBoundaries is not null)
            {
                block = block with { LeadingCommentBoundaries = commentBoundaries };
                commentBoundaries = null;
            }
            yield return block;
        }

        if (columnStarts is { Count: > 0 })
            recordTrailingColumnBookmarks?.Invoke(columnStarts.Values.ToArray());
    }

    private static DocxTableColumnBookmark? ReadTableColumnBookmark(XElement source, DocxBookmarkAnchor anchor)
    {
        return ReadNonNegativeIntAttribute(source, WordprocessingNamespace + "colFirst") is int first &&
            ReadNonNegativeIntAttribute(source, WordprocessingNamespace + "colLast") is int last && first <= last
            ? new DocxTableColumnBookmark(anchor, first, last) : null;
    }

    private static bool CanJoinTables(XElement first, XElement second, DocxStyleSet styles)
    {
        if (first.Name != WordprocessingNamespace + "tbl" || second.Name != first.Name) return false;
        var firstProperties = first.Element(WordprocessingNamespace + "tblPr");
        var secondProperties = second.Element(WordprocessingNamespace + "tblPr");
        // Word keeps positioned tables distinct even without a paragraph boundary.
        if (firstProperties?.Element(WordprocessingNamespace + "tblpPr") is not null ||
            secondProperties?.Element(WordprocessingNamespace + "tblpPr") is not null) return false;
        string? styleId = StyleId(firstProperties);
        if (!string.Equals(styleId, StyleId(secondProperties), StringComparison.Ordinal)) return false;
        DocxTableStyle? style = styleId is not null && styles.TableStyles.TryGetValue(styleId, out var selected)
            ? selected : styles.DefaultTableStyle;
        // Word keeps fixed and autofit tables separate. Matching modes join,
        // including differing widths and nested content (independent controls).
        return IsFixed(firstProperties) == IsFixed(secondProperties);

        bool IsFixed(XElement? properties) => string.Equals(
            (string?)properties?.Element(WordprocessingNamespace + "tblLayout")?.Attribute(WordprocessingNamespace + "type")
                ?? style?.Table.LayoutValue, "fixed", StringComparison.OrdinalIgnoreCase);

        string? StyleId(XElement? properties)
        {
            string? id = (string?)properties?.Element(WordprocessingNamespace + "tblStyle")?
                .Attribute(WordprocessingNamespace + "val");
            return id is not null && styles.TableStyles.ContainsKey(id) ? id : styles.DefaultTableStyleId;
        }
    }

    private static double[] ReadTableGrid(XElement table, CancellationToken token)
    {
        var columns = new List<double>();
        foreach (var column in table.Element(WordprocessingNamespace + "tblGrid")?
            .Elements(WordprocessingNamespace + "gridCol") ?? [])
        {
            token.ThrowIfCancellationRequested();
            columns.Add(ReadTwipsAttribute(column, WordprocessingNamespace + "w") ?? 72d);
        }
        return columns.ToArray();
    }

    private readonly record struct DocxTableSourceRow(int FirstColumn, int EndColumn,
        IReadOnlyList<(DocxRevisionScopedElement Scope, int Span)> Cells,
        Dictionary<string, int>? BookmarkEndBoundaries);

    private static DocxTableSourceRow ReadTableSourceRow(DocxRevisionScopedElement row, XElement? properties,
        OoxPdfDocxMarkupMode mode, CancellationToken token, bool joinedGrid = false,
        Action<string>? recordContainingCellBookmarkEnd = null)
    {
        int firstColumn = ReadRowGridBefore(properties);
        int endColumn = firstColumn;
        var cells = new List<(DocxRevisionScopedElement Scope, int Span)>();
        Dictionary<string, int>? bookmarkEndBoundaries = null;
        // Retain row-scope ends for column intervals still pending after an earlier cell.
        foreach (var scope in EnumerateBookmarkedBlocks(row.Element.Elements(), mode, token,
            inheritedRevision: null, tableChildName: WordprocessingNamespace + "tc",
            readTableColumnBookmarks: true, leadingColumnBookmarks: row.LeadingColumnBookmarks,
            recordBookmarkEnd: id =>
            {
                recordContainingCellBookmarkEnd?.Invoke(id);
                (bookmarkEndBoundaries ??= new(StringComparer.Ordinal)).TryAdd(id, cells.Count);
            }))
        {
            int span = ReadGridSpan(scope.Element.Element(WordprocessingNamespace + "tcPr"));
            endColumn += span;
            if (joinedGrid && endColumn > MaxInferredGridColumns) throw JoinedGridLimit();
            cells.Add((scope, span));
        }
        return new(firstColumn, endColumn, cells, bookmarkEndBoundaries);
    }

    private static (double[] Grid, DocxTableSourceRow[] Rows) CompleteJoinedTableGrid(double[] grid,
        IReadOnlyList<DocxRevisionScopedElement> rows, OoxPdfDocxMarkupMode mode, CancellationToken token,
        Action<string>? recordContainingCellBookmarkEnd)
    {
        int required = grid.Length;
        var sourceRows = new DocxTableSourceRow[rows.Count];
        for (int index = 0; index < sourceRows.Length; index++)
        {
            var row = rows[index];
            DocxTableSourceRow source = ReadTableSourceRow(row, row.Element.Element(WordprocessingNamespace + "trPr"),
                mode, token, joinedGrid: true, recordContainingCellBookmarkEnd: recordContainingCellBookmarkEnd);
            sourceRows[index] = source;
            required = Math.Max(required, source.EndColumn);
        }
        if (required > MaxInferredGridColumns) throw JoinedGridLimit();
        double[] completed = required == grid.Length ? grid : grid.Concat(Enumerable.Repeat(72d, required - grid.Length)).ToArray();
        return (completed, sourceRows);
    }

    private static (double[] Columns, int[][] Boundaries) JoinTableGrids(
        double[][] grids, double[] offsets, CancellationToken token)
    {
        // Decimal accumulation keeps equivalent twip partitions equal (3 * .05
        // and .15), without fitting a geometric tolerance to rendered output.
        // Source coordinates are Int64 twips; the existing 1024-column bound
        // keeps their decimal accumulation in range.
        var boundaries = new SortedSet<decimal> { 0m };
        var sourceBoundaries = new decimal[grids.Length][];
        for (int index = 0; index < grids.Length; index++)
        {
            token.ThrowIfCancellationRequested();
            var positions = new decimal[grids[index].Length + 1];
            positions[0] = (decimal)offsets[index];
            boundaries.Add(positions[0]);
            if (boundaries.Count > MaxInferredGridColumns + 1) throw JoinedGridLimit();
            for (int column = 0; column < grids[index].Length; column++)
            {
                token.ThrowIfCancellationRequested();
                double width = grids[index][column];
                if (width < 0d) throw new InvalidDataException("DOCX joined table grid has a negative column width.");
                decimal step = (decimal)width;
                positions[column + 1] = positions[column] + step;
                boundaries.Add(positions[column + 1]);
                if (boundaries.Count > MaxInferredGridColumns + 1) throw JoinedGridLimit();
            }
            sourceBoundaries[index] = positions;
        }
        decimal[] ordered = boundaries.ToArray();
        var columns = new double[ordered.Length - 1];
        for (int index = 0; index < columns.Length; index++) columns[index] = (double)(ordered[index + 1] - ordered[index]);
        var mappings = new int[grids.Length][];
        for (int index = 0; index < grids.Length; index++)
        {
            token.ThrowIfCancellationRequested();
            mappings[index] = sourceBoundaries[index].Select(position => Array.BinarySearch(ordered, position)).ToArray();
        }
        return (columns, mappings);
    }

    private static int ReadRowGridBefore(XElement? properties)
    {
        int count = ReadNonNegativeIntAttribute(properties?.Element(WordprocessingNamespace + "gridBefore"),
            WordprocessingNamespace + "val") ?? 0;
        if (count > MaxInferredGridColumns) throw JoinedGridLimit();
        return count;
    }

    private static OoxPdfLimitExceededException JoinedGridLimit() =>
        new("DOCX joined table grid exceeds the maximum supported column count of " + MaxInferredGridColumns + ".");

    private static DocxTableLook ReadRowTableLook(XElement row, XElement? tableProperties)
    {
        var exceptions = row.Element(WordprocessingNamespace + "tblPrEx");
        return ReadTableLook(exceptions?.Element(WordprocessingNamespace + "tblLook") is not null
            ? exceptions : tableProperties);
    }
}
