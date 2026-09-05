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
    private sealed record DocxTableStyle(
        string? BasedOnStyleId,
        DocxTableStyleProperties Table,
        DocxTableCellStyle Cell,
        IReadOnlyList<DocxTableCellBorder> TableBorders,
        IReadOnlyDictionary<string, DocxTableCellStyle> ConditionalRegions)
    {
        public static DocxTableStyle Empty { get; } = new(null, DocxTableStyleProperties.Empty, DocxTableCellStyle.Empty, [], new Dictionary<string, DocxTableCellStyle>());

        public DocxTableStyle Merge(DocxTableStyle other)
        {
            var conditional = new Dictionary<string, DocxTableCellStyle>(ConditionalRegions, StringComparer.Ordinal);
            foreach ((string region, DocxTableCellStyle regionStyle) in other.ConditionalRegions)
            {
                conditional[region] = conditional.TryGetValue(region, out DocxTableCellStyle? inherited)
                    ? inherited.Merge(regionStyle)
                    : regionStyle;
            }

            return new DocxTableStyle(
                other.BasedOnStyleId ?? BasedOnStyleId,
                Table.Merge(other.Table),
                Cell.Merge(other.Cell),
                other.TableBorders.Count == 0 ? TableBorders : other.TableBorders,
                conditional);
        }
    }

    private sealed record DocxTableStyleProperties(
        string? LayoutValue,
        double? PreferredWidthPoints,
        string? PreferredWidthValue,
        string? PreferredWidthType,
        double? IndentPoints,
        string? IndentValue,
        string? IndentType,
        double? CellSpacingPoints,
        string? CellSpacingValue,
        string? CellSpacingType,
        int? RowBandSize,
        int? ColumnBandSize)
    {
        public static DocxTableStyleProperties Empty { get; } = new(null, null, null, null, null, null, null, null, null, null, null, null);

        public DocxTableStyleProperties Merge(DocxTableStyleProperties other)
        {
            return new DocxTableStyleProperties(
                other.LayoutValue ?? LayoutValue,
                other.PreferredWidthPoints ?? PreferredWidthPoints,
                other.PreferredWidthValue ?? PreferredWidthValue,
                other.PreferredWidthType ?? PreferredWidthType,
                other.IndentPoints ?? IndentPoints,
                other.IndentValue ?? IndentValue,
                other.IndentType ?? IndentType,
                other.CellSpacingPoints ?? CellSpacingPoints,
                other.CellSpacingValue ?? CellSpacingValue,
                other.CellSpacingType ?? CellSpacingType,
                other.RowBandSize ?? RowBandSize,
                other.ColumnBandSize ?? ColumnBandSize);
        }
    }

    private sealed record DocxTableCellStyle(
        DocxResolvedParagraphProperties Paragraph,
        DocxResolvedRunProperties Run,
        string? FillHex,
        string? ShadingValue,
        string? ShadingColor,
        string? VerticalAlignmentValue,
        IReadOnlyList<DocxTableCellBorder> Borders,
        DocxTableCellMargins Margins,
        bool? NoWrap,
        string? NoWrapValue,
        bool? FitText,
        string? FitTextValue,
        string? TextDirectionValue)
    {
        public static DocxTableCellStyle Empty { get; } = new(DocxResolvedParagraphProperties.Empty, DocxResolvedRunProperties.Empty, null, null, null, null, [], DocxTableCellMargins.Empty, null, null, null, null, null);

        public DocxTableCellStyle Merge(DocxTableCellStyle other)
        {
            return new DocxTableCellStyle(
                Paragraph.Merge(other.Paragraph),
                Run.Merge(other.Run),
                other.FillHex ?? FillHex,
                other.ShadingValue ?? ShadingValue,
                other.ShadingColor ?? ShadingColor,
                other.VerticalAlignmentValue ?? VerticalAlignmentValue,
                other.Borders.Count == 0 ? Borders : other.Borders,
                other.Margins.Merge(Margins),
                other.NoWrap ?? NoWrap,
                other.NoWrapValue ?? NoWrapValue,
                other.FitText ?? FitText,
                other.FitTextValue ?? FitTextValue,
                other.TextDirectionValue ?? TextDirectionValue);
        }
    }

    private static DocxTableStyle ReadTableStyle(XElement style)
    {
        var conditional = new Dictionary<string, DocxTableCellStyle>(StringComparer.Ordinal);
        foreach (XElement region in style.Elements(WordprocessingNamespace + "tblStylePr"))
        {
            string? type = (string?)region.Attribute(WordprocessingNamespace + "type");
            if (type is not null)
            {
                conditional[type] = ReadTableCellStyle(
                    region.Element(WordprocessingNamespace + "tcPr"),
                    region.Element(WordprocessingNamespace + "pPr"),
                    region.Element(WordprocessingNamespace + "rPr"));
            }
        }

        XElement? tableProperties = style.Element(WordprocessingNamespace + "tblPr");
        return new DocxTableStyle(
            (string?)style.Element(WordprocessingNamespace + "basedOn")?.Attribute(WordprocessingNamespace + "val"),
            ReadTableStyleProperties(tableProperties),
            ReadTableCellStyle(
                style.Element(WordprocessingNamespace + "tcPr"),
                style.Element(WordprocessingNamespace + "pPr"),
                style.Element(WordprocessingNamespace + "rPr")) with
            {
                Margins = ReadTableStyleCellMargins(tableProperties)
            },
            ReadTableBorders(tableProperties),
            conditional);
    }

    private static DocxTableStyleProperties ReadTableStyleProperties(XElement? tableProperties)
    {
        XElement? tableWidth = tableProperties?.Element(WordprocessingNamespace + "tblW");
        XElement? tableIndent = tableProperties?.Element(WordprocessingNamespace + "tblInd");
        XElement? tableCellSpacing = tableProperties?.Element(WordprocessingNamespace + "tblCellSpacing");
        return new DocxTableStyleProperties(
            (string?)tableProperties
                ?.Element(WordprocessingNamespace + "tblLayout")
                ?.Attribute(WordprocessingNamespace + "type"),
            ReadDxaWidth(tableWidth),
            (string?)tableWidth?.Attribute(WordprocessingNamespace + "w"),
            (string?)tableWidth?.Attribute(WordprocessingNamespace + "type"),
            ReadDxaWidth(tableIndent),
            (string?)tableIndent?.Attribute(WordprocessingNamespace + "w"),
            (string?)tableIndent?.Attribute(WordprocessingNamespace + "type"),
            ReadDxaWidth(tableCellSpacing),
            (string?)tableCellSpacing?.Attribute(WordprocessingNamespace + "w"),
            (string?)tableCellSpacing?.Attribute(WordprocessingNamespace + "type"),
            ReadPositiveIntAttribute(tableProperties?.Element(WordprocessingNamespace + "tblStyleRowBandSize"), WordprocessingNamespace + "val"),
            ReadPositiveIntAttribute(tableProperties?.Element(WordprocessingNamespace + "tblStyleColBandSize"), WordprocessingNamespace + "val"));
    }

    private static IReadOnlyDictionary<string, DocxTableStyle> ResolveTableStyles(IReadOnlyDictionary<string, DocxTableStyle> tableStyles)
    {
        var resolved = new Dictionary<string, DocxTableStyle>(StringComparer.Ordinal);
        foreach (string styleId in tableStyles.Keys)
        {
            resolved[styleId] = ResolveTableStyle(styleId, tableStyles);
        }

        return resolved;
    }

    private static DocxTableStyle ResolveTableStyle(string styleId, IReadOnlyDictionary<string, DocxTableStyle> tableStyles)
    {
        var chain = new Stack<DocxTableStyle>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        string? currentStyleId = styleId;
        while (currentStyleId is not null && seen.Add(currentStyleId) && tableStyles.TryGetValue(currentStyleId, out DocxTableStyle? style))
        {
            chain.Push(style);
            currentStyleId = style.BasedOnStyleId;
        }

        DocxTableStyle resolved = DocxTableStyle.Empty;
        while (chain.Count != 0)
        {
            resolved = resolved.Merge(chain.Pop());
        }

        return resolved with
        {
            BasedOnStyleId = tableStyles.TryGetValue(styleId, out DocxTableStyle? source)
                ? source.BasedOnStyleId
                : null
        };
    }

    private static DocxTableCellStyle ReadTableCellStyle(
        XElement? cellProperties,
        XElement? paragraphProperties,
        XElement? runProperties)
    {
        XElement? shading = cellProperties?.Element(WordprocessingNamespace + "shd");
        XElement? noWrap = cellProperties?.Element(WordprocessingNamespace + "noWrap");
        XElement? fitText = cellProperties?.Element(WordprocessingNamespace + "tcFitText");
        return new DocxTableCellStyle(
            ReadParagraphProperties(paragraphProperties),
            ReadRunProperties(runProperties),
            (string?)shading?.Attribute(WordprocessingNamespace + "fill"),
            (string?)shading?.Attribute(WordprocessingNamespace + "val"),
            (string?)shading?.Attribute(WordprocessingNamespace + "color"),
            (string?)cellProperties
                ?.Element(WordprocessingNamespace + "vAlign")
                ?.Attribute(WordprocessingNamespace + "val"),
            ReadTableCellBorders(cellProperties),
            DocxTableCellMargins.Empty,
            ReadOnOff(noWrap),
            (string?)noWrap?.Attribute(WordprocessingNamespace + "val"),
            ReadOnOff(fitText),
            (string?)fitText?.Attribute(WordprocessingNamespace + "val"),
            (string?)cellProperties
                ?.Element(WordprocessingNamespace + "textDirection")
                ?.Attribute(WordprocessingNamespace + "val"));
    }

    private static DocxTableCellStyle ResolveTableCellStyle(
        DocxTableStyle tableStyle,
        DocxTableLook tableLook,
        DocxTableCellConditionalFormat? conditionalFormat,
        int rowIndex,
        int cellIndex,
        int rowCount,
        int cellCount)
    {
        DocxTableCellStyle resolved = tableStyle.Cell;
        IEnumerable<string> regions = conditionalFormat?.IsDefined == true
            ? EnumerateTableStyleRegions(conditionalFormat)
            : EnumerateTableStyleRegions(tableLook, tableStyle.Table.RowBandSize, tableStyle.Table.ColumnBandSize, rowIndex, cellIndex, rowCount, cellCount);
        foreach (string region in regions)
        {
            if (tableStyle.ConditionalRegions.TryGetValue(region, out DocxTableCellStyle? style))
            {
                resolved = resolved.Merge(style);
            }
        }

        return resolved;
    }

    private static IEnumerable<string> EnumerateTableStyleRegions(DocxTableCellConditionalFormat conditionalFormat)
    {
        if (conditionalFormat.FirstRow == true)
        {
            yield return "firstRow";
        }

        if (conditionalFormat.LastRow == true)
        {
            yield return "lastRow";
        }

        if (conditionalFormat.FirstColumn == true)
        {
            yield return "firstCol";
        }

        if (conditionalFormat.LastColumn == true)
        {
            yield return "lastCol";
        }

        if (conditionalFormat.FirstRowFirstColumn == true)
        {
            yield return "nwCell";
        }

        if (conditionalFormat.FirstRowLastColumn == true)
        {
            yield return "neCell";
        }

        if (conditionalFormat.LastRowFirstColumn == true)
        {
            yield return "swCell";
        }

        if (conditionalFormat.LastRowLastColumn == true)
        {
            yield return "seCell";
        }

        if (conditionalFormat.OddHorizontalBand == true)
        {
            yield return "band1Horz";
        }

        if (conditionalFormat.EvenHorizontalBand == true)
        {
            yield return "band2Horz";
        }

        if (conditionalFormat.OddVerticalBand == true)
        {
            yield return "band1Vert";
        }

        if (conditionalFormat.EvenVerticalBand == true)
        {
            yield return "band2Vert";
        }
    }

    private static IEnumerable<string> EnumerateTableStyleRegions(
        DocxTableLook tableLook,
        int? rowBandSize,
        int? columnBandSize,
        int rowIndex,
        int cellIndex,
        int rowCount,
        int cellCount)
    {
        bool firstRow = tableLook.FirstRow != false;
        bool lastRow = tableLook.LastRow == true;
        bool firstColumn = tableLook.FirstColumn == true;
        bool lastColumn = tableLook.LastColumn == true;
        bool horizontalBand = tableLook.NoHorizontalBand != true;
        bool verticalBand = tableLook.NoVerticalBand != true;

        if (firstRow && rowIndex == 0)
        {
            yield return "firstRow";
        }

        if (lastRow && rowIndex == rowCount - 1)
        {
            yield return "lastRow";
        }

        if (firstColumn && cellIndex == 0)
        {
            yield return "firstCol";
        }

        if (lastColumn && cellIndex == cellCount - 1)
        {
            yield return "lastCol";
        }

        if (firstRow && firstColumn && rowIndex == 0 && cellIndex == 0)
        {
            yield return "nwCell";
        }

        if (firstRow && lastColumn && rowIndex == 0 && cellIndex == cellCount - 1)
        {
            yield return "neCell";
        }

        if (lastRow && firstColumn && rowIndex == rowCount - 1 && cellIndex == 0)
        {
            yield return "swCell";
        }

        if (lastRow && lastColumn && rowIndex == rowCount - 1 && cellIndex == cellCount - 1)
        {
            yield return "seCell";
        }

        string? horizontalBandRegion = ResolveBandRegion(rowIndex, rowBandSize ?? 1, "band1Horz", "band2Horz");
        if (horizontalBand && horizontalBandRegion is not null)
        {
            yield return horizontalBandRegion;
        }

        string? verticalBandRegion = ResolveBandRegion(cellIndex, columnBandSize ?? 1, "band1Vert", "band2Vert");
        if (verticalBand && verticalBandRegion is not null)
        {
            yield return verticalBandRegion;
        }
    }

    private static string? ResolveBandRegion(int index, int bandSize, string firstBand, string secondBand)
    {
        if (index == 0)
        {
            return null;
        }

        int effectiveBandSize = Math.Max(1, bandSize);
        int bandIndex = (index - 1) / effectiveBandSize;
        return bandIndex % 2 == 0 ? firstBand : secondBand;
    }
}
