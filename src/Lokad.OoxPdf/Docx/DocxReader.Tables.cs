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
        IReadOnlyList<DocxRevisionScopedElement>? fragments = null)
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
        bool anchorsStyleGrid = isCustomStyle && tableIndent is null && tableStyle.Table.IndentValue is null;
        // Office A/B (m13/ladder compat15 probes): a present compatibilityMode pins the
        // table grid at margin/indent plus the outer border half (legacy outer-edge
        // alignment); without it the grid sits at the origin (modern) or pins text via
        // style margins (w72/m12, see PinTextToMargin below). Any stated mode counts
        // (survey: only 15 occurs in the wild).
        bool useLegacyTableGrid = documentSettings is not null && documentSettings.CompatSettings.Any(setting => string.Equals(setting.Name, "compatibilityMode", StringComparison.Ordinal));
        double? resolvedTableIndentPoints = tableIndent is not null ? ReadDxaWidth(tableIndent) : tableStyle.Table.IndentPoints;
        var parts = (fragments ?? [source]).Select(fragment => (
            Properties: fragment.Element.Element(WordprocessingNamespace + "tblPr"),
            Grid: ReadTableGrid(fragment.Element, cancellationToken),
            Rows: EnumerateRevisionScopedChildren(fragment.Element.Elements(), markupMode, cancellationToken,
                [WordprocessingNamespace + "tr"], fragment.Revision == inheritedRevision ? null : fragment.Revision)
                .Where(row => IsIncludedTableRow(row.Element, markupMode, cancellationToken)).ToArray())).ToArray();
        bool hasExplicitGrid = parts.Any(part => part.Grid.Length != 0);
        IReadOnlyList<double> columns = parts[0].Grid;
        int[][]? gridMappings = null;
        if (parts.Length > 1)
        {
            var grids = parts.Select(part => CompleteJoinedTableGrid(part.Grid, part.Rows, markupMode, cancellationToken)).ToArray();
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
        int rowIndex = 0;
        for (int sourceIndex = 0; sourceIndex < parts.Length; sourceIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var part = parts[sourceIndex];
            DocxTableCellMargins partMargins = ReadTableStyleCellMargins(part.Properties);
            IReadOnlyList<DocxTableCellBorder> partBorders = ReadTableBorders(part.Properties);
            foreach (var rowElement in part.Rows)
            {
                cancellationToken.ThrowIfCancellationRequested();
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
                int sourceColumn = ReadRowGridBefore(rowProperties);
                int gridBefore = gridMappings is null ? sourceColumn : gridMappings[sourceIndex][sourceColumn];
                DocxRevisionScopedElement[] cellElements = EnumerateRevisionScopedChildren(row.Elements(), markupMode, cancellationToken, [WordprocessingNamespace + "tc"]).ToArray();
                for (int cellIndex = 0; cellIndex < cellElements.Length; cellIndex++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    XElement cell = cellElements[cellIndex].Element;
                    DocxRevisionInfo? inheritedCellRevision = cellElements[cellIndex].Revision ?? rowElement.Revision;
                    XElement? cellProperties = cell.Element(WordprocessingNamespace + "tcPr");
                    var cellRevisions = new List<DocxRevisionInfo>();
                    AddRevision(cellRevisions, cellElements[cellIndex].Revision);
                    AddRevisions(cellRevisions, ReadPropertyChangeRevisions(cellProperties));
                    DocxTableCellConditionalFormat? conditionalFormat = ReadTableCellConditionalFormat(cellProperties);
                    DocxTableCellStyle conditionalStyle = ResolveTableCellStyle(tableStyle, rowLook, conditionalFormat, rowIndex, cellIndex, rowCount, cellElements.Length, leadingRowEmphasis);
                    IReadOnlyList<DocxBodyElement> cellBodyElements = ReadTableCellBodyElements(
                        cell,
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
                        inheritedCellRevision);
                    IReadOnlyList<DocxParagraph> paragraphs = DocxBlockTraversal.EnumerateDirectParagraphs(cellBodyElements).ToArray();
                    string text = string.Join(" ", paragraphs
                        .Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text)))
                        .Where(t => t.Length != 0));
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
                        cellElements.Length);
                    DocxTableCellMargins inheritedMargins = rowInheritedMargins.Merge(conditionalStyle.Margins);
                    DocxTableCellMargins margins = ReadTableCellMargins(cellProperties).Merge(inheritedMargins);
                    DocxTableCellMargins styleMargins = tableStyle.Cell.Margins.Merge(conditionalStyle.Margins);
                    // Without compat, an intrinsic style pins text through the first
                    // cell's effective padding. Table, row and cell declarations replace
                    // style padding for both that grid anchor and the content offset.
                    bool pinTextToMargin = !isCustomStyle && !useLegacyTableGrid && (resolvedTableIndentPoints ?? 0d) == 0d && (styleMargins.LeftPoints ?? 0d) > 0d;
                    DocxTableCellMargins? gridAnchorMargins = anchorsStyleGrid
                        ? DocxTableCellMargins.Empty with { LeftPoints = 0d, RightPoints = 0d }
                        : pinTextToMargin ? margins : null;
                    int sourceSpan = ReadGridSpan(cellProperties);
                    int gridSpan = gridMappings is null ? sourceSpan : Math.Max(1,
                        gridMappings[sourceIndex][sourceColumn + sourceSpan] - gridMappings[sourceIndex][sourceColumn]);
                    sourceColumn += sourceSpan;
                    cells.Add(new DocxTableCell(
                        text,
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
                    XElement? header = rowProperties?.Element(WordprocessingNamespace + "tblHeader");
                    XElement? cantSplit = rowProperties?.Element(WordprocessingNamespace + "cantSplit");
                    XElement? rowHeight = rowProperties?.Element(WordprocessingNamespace + "trHeight");
                    DocxTableCellMargins effectiveExceptions = sourceIndex == 0 ? rowExceptionMargins : rowExceptionMargins.Merge(partMargins);
                    rows.Add(new DocxTableRow(
                        cells,
                        ReadTableRowHeight(rowHeight),
                        ReadOnOff(header) == true,
                        (string?)header?.Attribute(WordprocessingNamespace + "val"),
                        (string?)rowHeight?.Attribute(WordprocessingNamespace + "val"),
                        (string?)rowHeight?.Attribute(WordprocessingNamespace + "hRule"),
                        HasAnyTableCellMargin(effectiveExceptions) ? effectiveExceptions : null,
                        ReadOnOff(cantSplit) == true,
                        (string?)cantSplit?.Attribute(WordprocessingNamespace + "val"))
                    {
                        Revisions = rowRevisions,
                        GridBefore = gridBefore
                    });
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

    private static IReadOnlyList<DocxBodyElement> ReadTableCellBodyElements(
        XElement cell,
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
        DocxRevisionInfo? inheritedRevision)
    {
        var elements = new List<DocxBodyElement>();
        foreach (var (scopedChild, tableFragments) in EnumerateLogicalTableBlocks(cell.Elements(), styles, markupMode, cancellationToken, inheritedRevision))
        {
            cancellationToken.ThrowIfCancellationRequested();
            XElement child = scopedChild.Element;
            DocxRevisionInfo? childRevision = scopedChild.Revision ?? inheritedRevision;
            if (child.Name == WordprocessingNamespace + "tcPr")
            {
                continue;
            }

            if (child.Name == WordprocessingNamespace + "p")
            {
                if (IsRunColumnBreakOnlyParagraph(child, markupMode))
                {
                    DocxParagraph? breakParagraph = ReadParagraph(
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
                        cancellationToken: cancellationToken);
                    elements.Add(DocxBodyElementFactory.CreateManualBreak(DocxBreakSourceKind.RunBreak, "column", breakParagraph));
                    continue;
                }

                if (HasRunPageOrColumnBreak(child, markupMode))
                {
                    foreach (ParagraphBreakPart part in SplitParagraphAtRunBreaks(child, markupMode))
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (part.BreakValue is not null)
                        {
                            elements.Add(string.Equals(part.BreakValue, "column", StringComparison.OrdinalIgnoreCase)
                                ? DocxBodyElementFactory.CreateManualBreak(DocxBreakSourceKind.RunBreak, "column", null)
                                : DocxBodyElementFactory.CreatePageBreak(DocxBreakSourceKind.RunBreak, part.BreakValue, null, null));
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
                            cancellationToken: cancellationToken);
                        if (splitParagraph is not null)
                        {
                            elements.Add(DocxBodyElementFactory.CreateParagraph(AdjustBreakParagraphFragment(splitParagraph, part)));
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
                    cancellationToken: cancellationToken);
                if (parsed is not null)
                {
                    elements.Add(DocxBodyElementFactory.CreateParagraph(parsed));
                }
            }
            else if (child.Name == WordprocessingNamespace + "tbl")
            {
                DocxTable? nestedTable = ReadTable(scopedChild, styles, numbering, numberingCounters, package, relationships, inlineReferenceCounters, documentSettings, markupMode, cancellationToken, tableFragments);
                if (nestedTable is not null)
                {
                    elements.Add(DocxBodyElementFactory.CreateTable(nestedTable));
                }
            }
        }

        return NormalizeDeletedParagraphMarkElements(elements, markupMode);
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

    private static DocxTableCellConditionalFormat? ReadTableCellConditionalFormat(XElement? cellProperties)
    {
        XElement? conditional = cellProperties?.Element(WordprocessingNamespace + "cnfStyle");
        if (conditional is null)
        {
            return null;
        }

        return new DocxTableCellConditionalFormat(
            (string?)conditional.Attribute(WordprocessingNamespace + "val"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "firstRow"),
            (string?)conditional.Attribute(WordprocessingNamespace + "firstRow"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "lastRow"),
            (string?)conditional.Attribute(WordprocessingNamespace + "lastRow"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "firstColumn"),
            (string?)conditional.Attribute(WordprocessingNamespace + "firstColumn"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "lastColumn"),
            (string?)conditional.Attribute(WordprocessingNamespace + "lastColumn"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "oddHBand"),
            (string?)conditional.Attribute(WordprocessingNamespace + "oddHBand"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "evenHBand"),
            (string?)conditional.Attribute(WordprocessingNamespace + "evenHBand"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "oddVBand"),
            (string?)conditional.Attribute(WordprocessingNamespace + "oddVBand"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "evenVBand"),
            (string?)conditional.Attribute(WordprocessingNamespace + "evenVBand"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "firstRowFirstColumn"),
            (string?)conditional.Attribute(WordprocessingNamespace + "firstRowFirstColumn"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "firstRowLastColumn"),
            (string?)conditional.Attribute(WordprocessingNamespace + "firstRowLastColumn"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "lastRowFirstColumn"),
            (string?)conditional.Attribute(WordprocessingNamespace + "lastRowFirstColumn"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "lastRowLastColumn"),
            (string?)conditional.Attribute(WordprocessingNamespace + "lastRowLastColumn"));
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
        DocxTableCellBorder? top = FindBorder(tableBorders, rowIndex == 0 ? "top" : "insideH");
        if (top is not null)
        {
            yield return top with { Edge = "top" };
        }

        DocxTableCellBorder? bottom = FindBorder(tableBorders, rowIndex == rowCount - 1 ? "bottom" : "insideH");
        if (bottom is not null)
        {
            yield return bottom with { Edge = "bottom" };
        }

        DocxTableCellBorder? left = cellIndex == 0
            ? FindBorder(tableBorders, "left") ?? FindBorder(tableBorders, "start")
            : FindBorder(tableBorders, "insideV");
        if (left is not null)
        {
            yield return left with { Edge = "left" };
        }

        DocxTableCellBorder? right = cellIndex == cellCount - 1
            ? FindBorder(tableBorders, "right") ?? FindBorder(tableBorders, "end")
            : FindBorder(tableBorders, "insideV");
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

    private static DocxTableCellBorder? FindBorder(IReadOnlyList<DocxTableCellBorder> borders, string edge)
    {
        return borders.FirstOrDefault(border => string.Equals(border.Edge, edge, StringComparison.OrdinalIgnoreCase));
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
