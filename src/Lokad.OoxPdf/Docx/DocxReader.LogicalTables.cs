using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxReader
{
    // Keep original XML/source scopes: grouping must not move or clone paragraphs,
    // drawings, references or revision owners. Only joined tables need a list.
    private static IEnumerable<(DocxRevisionScopedElement Scope, IReadOnlyList<DocxRevisionScopedElement>? Tables)>
        EnumerateLogicalTableBlocks(IEnumerable<XElement> elements, DocxStyleSet styles,
            OoxPdfDocxMarkupMode mode, CancellationToken token, DocxRevisionInfo? inheritedRevision = null)
    {
        DocxRevisionScopedElement? pending = null;
        List<DocxRevisionScopedElement>? tables = null;
        foreach (var scope in EnumerateRevisionScopedChildren(elements, mode, token,
            [WordprocessingNamespace + "p", WordprocessingNamespace + "tbl", WordprocessingNamespace + "altChunk"], inheritedRevision))
        {
            token.ThrowIfCancellationRequested();
            if (pending is not null && CanJoinTables(pending.Element, scope.Element, styles))
            {
                tables ??= [pending];
                tables.Add(scope);
                continue;
            }
            if (pending is not null) yield return (pending, tables);
            pending = scope;
            tables = null;
        }
        if (pending is not null) yield return (pending, tables);
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

    private static double[] CompleteJoinedTableGrid(double[] grid, IReadOnlyList<DocxRevisionScopedElement> rows,
        OoxPdfDocxMarkupMode mode, CancellationToken token)
    {
        int required = grid.Length;
        foreach (var row in rows)
        {
            int count = ReadRowGridBefore(row.Element.Element(WordprocessingNamespace + "trPr"));
            foreach (var cell in EnumerateRevisionScopedChildren(row.Element.Elements(), mode, token, [WordprocessingNamespace + "tc"]))
            {
                count += ReadGridSpan(cell.Element.Element(WordprocessingNamespace + "tcPr"));
                if (count > MaxInferredGridColumns) throw JoinedGridLimit();
            }
            required = Math.Max(required, count);
        }
        if (required > MaxInferredGridColumns) throw JoinedGridLimit();
        return required == grid.Length ? grid : grid.Concat(Enumerable.Repeat(72d, required - grid.Length)).ToArray();
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
