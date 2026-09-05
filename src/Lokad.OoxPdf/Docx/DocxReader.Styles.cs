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
    private static DocxStyleSet LoadStyles(OoxPackage package, string documentPartName, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxRelationship? styleRelationship = package.GetRelationships(documentPartName, cancellationToken)
            .FirstOrDefault(r => !r.IsExternal && r.Type == StylesRelationshipType && r.ResolvedTarget is not null);
        OoxPart? stylesPart = styleRelationship?.ResolvedTarget is null
            ? package.Parts.FirstOrDefault(p => p.ContentType == StylesContentType)
            : package.GetPart(styleRelationship.ResolvedTarget);
        if (stylesPart is null)
        {
            return DocxStyleSet.Empty;
        }

        using Stream stream = stylesPart.OpenRead();
        XDocument stylesXml = SafeXml.Load(stream, cancellationToken);
        DocxResolvedRunProperties runDefaults = ReadRunProperties(stylesXml
            .Root?
            .Element(WordprocessingNamespace + "docDefaults")
            ?.Element(WordprocessingNamespace + "rPrDefault")
            ?.Element(WordprocessingNamespace + "rPr"));
        DocxResolvedParagraphProperties paragraphDefaults = ReadParagraphProperties(stylesXml
            .Root?
            .Element(WordprocessingNamespace + "docDefaults")
            ?.Element(WordprocessingNamespace + "pPrDefault")
            ?.Element(WordprocessingNamespace + "pPr"));

        var paragraphStyles = new Dictionary<string, DocxStyle>(StringComparer.Ordinal);
        var characterStyles = new Dictionary<string, DocxStyle>(StringComparer.Ordinal);
        var tableStyles = new Dictionary<string, DocxTableStyle>(StringComparer.Ordinal);
        string? defaultTableStyleId = null;
        foreach (XElement style in stylesXml.Root?.Elements(WordprocessingNamespace + "style") ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();
            string? styleId = (string?)style.Attribute(WordprocessingNamespace + "styleId");
            string? type = (string?)style.Attribute(WordprocessingNamespace + "type");
            if (string.IsNullOrWhiteSpace(styleId))
            {
                continue;
            }

            var parsed = new DocxStyle(
                (string?)style.Element(WordprocessingNamespace + "basedOn")?.Attribute(WordprocessingNamespace + "val"),
                ReadParagraphProperties(style.Element(WordprocessingNamespace + "pPr")),
                ReadRunProperties(style.Element(WordprocessingNamespace + "rPr")));
            if (type == "paragraph")
            {
                paragraphStyles[styleId] = parsed;
            }
            else if (type == "character")
            {
                characterStyles[styleId] = parsed;
            }
            else if (type == "table")
            {
                tableStyles[styleId] = ReadTableStyle(style);
                if (OoxBoolean.ParseAttribute(style, WordprocessingNamespace + "default", false))
                {
                    defaultTableStyleId = styleId;
                }
            }
        }

        IReadOnlyDictionary<string, DocxTableStyle> resolvedTableStyles = ResolveTableStyles(tableStyles);
        DocxTableStyle? defaultTableStyle = defaultTableStyleId is not null && resolvedTableStyles.TryGetValue(defaultTableStyleId, out DocxTableStyle? resolvedDefault)
            ? resolvedDefault
            : null;
        return new DocxStyleSet(runDefaults, paragraphDefaults, paragraphStyles, characterStyles, resolvedTableStyles, defaultTableStyleId, defaultTableStyle);
    }

    private static DocxStyleCatalog ToStyleCatalog(DocxStyleSet styles)
    {
        return new DocxStyleCatalog(
            styles.RunDefaults != DocxResolvedRunProperties.Empty,
            styles.ParagraphDefaults != DocxResolvedParagraphProperties.Empty,
            styles.DefaultTableStyleId,
            styles.ParagraphStyles
                .OrderBy(style => style.Key, StringComparer.Ordinal)
                .Select(style => ToStyleDefinitionSummary(style.Key, style.Value))
                .ToArray(),
            styles.CharacterStyles
                .OrderBy(style => style.Key, StringComparer.Ordinal)
                .Select(style => ToStyleDefinitionSummary(style.Key, style.Value))
                .ToArray(),
            styles.TableStyles
                .OrderBy(style => style.Key, StringComparer.Ordinal)
                .Select(style => ToTableStyleDefinitionSummary(style.Key, style.Value))
                .ToArray());
    }

    private static DocxStyleDefinitionSummary ToStyleDefinitionSummary(string styleId, DocxStyle style)
    {
        return new DocxStyleDefinitionSummary(
            styleId,
            style.BasedOnStyleId,
            style.Paragraph != DocxResolvedParagraphProperties.Empty,
            style.Run != DocxResolvedRunProperties.Empty);
    }

    private static DocxTableStyleDefinitionSummary ToTableStyleDefinitionSummary(string styleId, DocxTableStyle style)
    {
        return new DocxTableStyleDefinitionSummary(
            styleId,
            style.BasedOnStyleId,
            style.Table != DocxTableStyleProperties.Empty,
            style.Cell != DocxTableCellStyle.Empty,
            style.Cell.Paragraph != DocxResolvedParagraphProperties.Empty,
            style.Cell.Run != DocxResolvedRunProperties.Empty,
            style.TableBorders.Count,
            style.ConditionalRegions.Count);
    }

    private static DocxResolvedParagraphProperties ResolveParagraphProperties(
        XElement? directProperties,
        string? paragraphStyleId,
        DocxStyleSet styles,
        DocxResolvedParagraphProperties? tableStyleProperties)
    {
        DocxResolvedParagraphProperties result = styles.ParagraphDefaults;
        foreach (DocxStyle style in EnumerateStyleInheritance(paragraphStyleId, styles.ParagraphStyles))
        {
            result = result.Merge(style.Paragraph);
        }

        if (tableStyleProperties is { } tableProperties)
        {
            result = result.Merge(tableProperties);
        }

        return result.Merge(ReadParagraphProperties(directProperties));
    }

    private static DocxParagraphStyleResolution CreateParagraphStyleResolution(
        XElement? directProperties,
        string? paragraphStyleId,
        DocxStyleSet styles,
        DocxResolvedParagraphProperties? tableStyleProperties)
    {
        DocxStyle[] styleChain = EnumerateStyleInheritance(paragraphStyleId, styles.ParagraphStyles).ToArray();
        return new DocxParagraphStyleResolution(
            paragraphStyleId,
            paragraphStyleId is not null && styleChain.Length != 0,
            styleChain.Length,
            styles.ParagraphDefaults != DocxResolvedParagraphProperties.Empty,
            HasDirectParagraphProperties(directProperties),
            tableStyleProperties is not null && tableStyleProperties.Value != DocxResolvedParagraphProperties.Empty);
    }

    private static bool HasDirectParagraphProperties(XElement? properties)
    {
        return properties?.Elements().Any(element =>
            element.Name != WordprocessingNamespace + "pStyle" &&
            element.Name != WordprocessingNamespace + "rPr") == true;
    }

    private static DocxResolvedRunProperties ResolveRunProperties(
        XElement? directProperties,
        string? paragraphStyleId,
        string? characterStyleId,
        DocxStyleSet styles,
        DocxResolvedRunProperties? tableStyleProperties)
    {
        DocxResolvedRunProperties result = styles.RunDefaults;
        if (tableStyleProperties is { } tableProperties)
        {
            result = result.Merge(tableProperties);
        }

        foreach (DocxStyle paragraphStyle in EnumerateStyleInheritance(paragraphStyleId, styles.ParagraphStyles))
        {
            result = result.Merge(paragraphStyle.Run);
        }

        foreach (DocxStyle characterStyle in EnumerateStyleInheritance(characterStyleId, styles.CharacterStyles))
        {
            result = result.Merge(characterStyle.Run);
        }

        return result.Merge(ReadRunProperties(directProperties));
    }

    private static DocxRunStyleResolution CreateRunStyleResolution(
        XElement? directProperties,
        string? paragraphStyleId,
        string? characterStyleId,
        DocxStyleSet styles,
        DocxResolvedRunProperties? tableStyleProperties)
    {
        DocxStyle[] paragraphStyleChain = EnumerateStyleInheritance(paragraphStyleId, styles.ParagraphStyles).ToArray();
        DocxStyle[] characterStyleChain = EnumerateStyleInheritance(characterStyleId, styles.CharacterStyles).ToArray();
        return new DocxRunStyleResolution(
            characterStyleId,
            characterStyleId is not null && characterStyleChain.Length != 0,
            characterStyleChain.Length,
            styles.RunDefaults != DocxResolvedRunProperties.Empty,
            paragraphStyleChain.Any(style => style.Run != DocxResolvedRunProperties.Empty),
            characterStyleChain.Any(style => style.Run != DocxResolvedRunProperties.Empty),
            HasDirectRunProperties(directProperties),
            tableStyleProperties is not null && tableStyleProperties.Value != DocxResolvedRunProperties.Empty);
    }

    private static bool HasDirectRunProperties(XElement? properties)
    {
        return properties?.Elements().Any(element => element.Name != WordprocessingNamespace + "rStyle") == true;
    }

    private static IEnumerable<DocxStyle> EnumerateStyleInheritance(string? styleId, IReadOnlyDictionary<string, DocxStyle> styles)
    {
        if (styleId is null)
        {
            yield break;
        }

        var chain = new Stack<DocxStyle>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        string? currentStyleId = styleId;
        while (currentStyleId is not null && seen.Add(currentStyleId) && styles.TryGetValue(currentStyleId, out DocxStyle? style))
        {
            chain.Push(style);
            currentStyleId = style.BasedOnStyleId;
        }

        while (chain.Count != 0)
        {
            yield return chain.Pop();
        }
    }

    private static DocxResolvedParagraphProperties ReadParagraphProperties(XElement? properties)
    {
        string? alignmentValue = ReadAlignmentValue(properties);
        DocxTextAlignment? alignment = ReadAlignment(alignmentValue);
        XElement? spacing = properties?.Element(WordprocessingNamespace + "spacing");
        double? before = ReadTwipsAttribute(spacing, WordprocessingNamespace + "before");
        double? after = ReadTwipsAttribute(spacing, WordprocessingNamespace + "after");
        double? lineFactor = null;
        double? linePoints = null;
        if (spacing?.Attribute(WordprocessingNamespace + "line") is { } line)
        {
            string? lineRule = (string?)spacing.Attribute(WordprocessingNamespace + "lineRule");
            if (string.Equals(lineRule, "exact", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(lineRule, "atLeast", StringComparison.OrdinalIgnoreCase))
            {
                linePoints = OoxUnits.TwipsToPoints(long.Parse(line.Value, CultureInfo.InvariantCulture));
            }
            else
            {
                lineFactor = int.Parse(line.Value, CultureInfo.InvariantCulture) / 240d;
            }
        }

        DocxParagraphSpacing paragraphSpacing = new(
            (string?)spacing?.Attribute(WordprocessingNamespace + "before"),
            (string?)spacing?.Attribute(WordprocessingNamespace + "after"),
            (string?)spacing?.Attribute(WordprocessingNamespace + "beforeLines"),
            (string?)spacing?.Attribute(WordprocessingNamespace + "afterLines"),
            (string?)spacing?.Attribute(WordprocessingNamespace + "beforeAutospacing"),
            (string?)spacing?.Attribute(WordprocessingNamespace + "afterAutospacing"),
            (string?)spacing?.Attribute(WordprocessingNamespace + "line"),
            (string?)spacing?.Attribute(WordprocessingNamespace + "lineRule"),
            ReadOnOff(properties?.Element(WordprocessingNamespace + "contextualSpacing")));
        DocxParagraphKeepRules keepRules = new(
            ReadOnOff(properties?.Element(WordprocessingNamespace + "keepNext")),
            (string?)properties?.Element(WordprocessingNamespace + "keepNext")?.Attribute(WordprocessingNamespace + "val"),
            ReadOnOff(properties?.Element(WordprocessingNamespace + "keepLines")),
            (string?)properties?.Element(WordprocessingNamespace + "keepLines")?.Attribute(WordprocessingNamespace + "val"),
            ReadOnOff(properties?.Element(WordprocessingNamespace + "widowControl")),
            (string?)properties?.Element(WordprocessingNamespace + "widowControl")?.Attribute(WordprocessingNamespace + "val"));
        DocxParagraphIndent indent = ReadParagraphIndent(properties);
        IReadOnlyList<DocxTabStop> tabStops = ReadParagraphTabStops(properties);
        XElement? snapToGrid = properties?.Element(WordprocessingNamespace + "snapToGrid");
        XElement? pageBreakBefore = properties?.Element(WordprocessingNamespace + "pageBreakBefore");
        XElement? wordWrap = properties?.Element(WordprocessingNamespace + "wordWrap");

        return new DocxResolvedParagraphProperties(
            alignment,
            alignmentValue,
            before,
            after,
            lineFactor,
            linePoints,
            paragraphSpacing,
            keepRules,
            indent,
            tabStops,
            ReadOnOff(snapToGrid),
            (string?)snapToGrid?.Attribute(WordprocessingNamespace + "val"),
            ReadOnOff(pageBreakBefore),
            (string?)pageBreakBefore?.Attribute(WordprocessingNamespace + "val"),
            ReadOnOff(wordWrap),
            (string?)wordWrap?.Attribute(WordprocessingNamespace + "val"));
    }

    private static IReadOnlyList<DocxTabStop> ReadParagraphTabStops(XElement? properties)
    {
        return properties?
            .Element(WordprocessingNamespace + "tabs")
            ?.Elements(WordprocessingNamespace + "tab")
            .Select(tab => new DocxTabStop(
                ReadTwipsAttribute(tab, WordprocessingNamespace + "pos"),
                (string?)tab.Attribute(WordprocessingNamespace + "pos"),
                (string?)tab.Attribute(WordprocessingNamespace + "val"),
                (string?)tab.Attribute(WordprocessingNamespace + "leader")))
            .ToArray() ?? [];
    }

    private static DocxParagraphIndent ReadParagraphIndent(XElement? properties)
    {
        XElement? indent = properties?.Element(WordprocessingNamespace + "ind");
        return new DocxParagraphIndent(
            ReadLogicalStartTwips(indent),
            ReadLogicalEndTwips(indent),
            ReadTwipsAttribute(indent, WordprocessingNamespace + "firstLine"),
            ReadTwipsAttribute(indent, WordprocessingNamespace + "hanging"),
            ReadLogicalStartValue(indent),
            ReadLogicalEndValue(indent),
            (string?)indent?.Attribute(WordprocessingNamespace + "firstLine"),
            (string?)indent?.Attribute(WordprocessingNamespace + "hanging"));
    }

    private static string? ReadAlignmentValue(XElement? properties)
    {
        return (string?)properties
            ?.Element(WordprocessingNamespace + "jc")
            ?.Attribute(WordprocessingNamespace + "val");
    }

    private static DocxTextAlignment? ReadAlignment(string? value)
    {
        return value switch
        {
            "center" => DocxTextAlignment.Center,
            "right" => DocxTextAlignment.Right,
            "both" => DocxTextAlignment.Justified,
            null => null,
            _ => DocxTextAlignment.Left
        };
    }

    private static DocxResolvedRunProperties ReadRunProperties(XElement? properties)
    {
        double? fontSize = properties?
            .Element(WordprocessingNamespace + "sz")
            ?.Attribute(WordprocessingNamespace + "val") is { } size
            ? int.Parse(size.Value, CultureInfo.InvariantCulture) / 2d
            : null;
        string? color = (string?)properties?
            .Element(WordprocessingNamespace + "color")
            ?.Attribute(WordprocessingNamespace + "val");
        string? fontFamily = (string?)properties?
            .Element(WordprocessingNamespace + "rFonts")
            ?.Attribute(WordprocessingNamespace + "ascii");
        bool? bold = ReadOnOff(properties?.Element(WordprocessingNamespace + "b"));
        bool? italic = ReadOnOff(properties?.Element(WordprocessingNamespace + "i"));
        bool? complexScriptBold = ReadOnOff(properties?.Element(WordprocessingNamespace + "bCs"));
        bool? complexScriptItalic = ReadOnOff(properties?.Element(WordprocessingNamespace + "iCs"));
        bool? allCaps = ReadOnOff(properties?.Element(WordprocessingNamespace + "caps"));
        XElement? smallCapsElement = properties?.Element(WordprocessingNamespace + "smallCaps");
        bool? smallCaps = ReadOnOff(smallCapsElement);
        string? smallCapsValue = (string?)smallCapsElement?.Attribute(WordprocessingNamespace + "val");
        XElement? hiddenElement = properties?.Element(WordprocessingNamespace + "vanish");
        bool? hidden = ReadOnOff(hiddenElement);
        string? hiddenValue = (string?)hiddenElement?.Attribute(WordprocessingNamespace + "val");
        XElement? strikeElement = properties?.Element(WordprocessingNamespace + "strike");
        XElement? doubleStrikeElement = properties?.Element(WordprocessingNamespace + "dstrike");
        bool? strike = ReadOnOff(strikeElement);
        bool? doubleStrike = ReadOnOff(doubleStrikeElement);
        string? strikeValue = (string?)strikeElement?.Attribute(WordprocessingNamespace + "val");
        string? doubleStrikeValue = (string?)doubleStrikeElement?.Attribute(WordprocessingNamespace + "val");
        double? characterSpacingPoints = ReadSignedTwipsElement(properties?.Element(WordprocessingNamespace + "spacing"));
        string? verticalAlignmentValue = (string?)properties?
            .Element(WordprocessingNamespace + "vertAlign")
            ?.Attribute(WordprocessingNamespace + "val");
        string? highlightValue = (string?)properties?
            .Element(WordprocessingNamespace + "highlight")
            ?.Attribute(WordprocessingNamespace + "val");
        XElement? shading = properties?.Element(WordprocessingNamespace + "shd");
        string? shadingFill = (string?)shading?.Attribute(WordprocessingNamespace + "fill");
        string? shadingValue = (string?)shading?.Attribute(WordprocessingNamespace + "val");
        string? shadingColor = (string?)shading?.Attribute(WordprocessingNamespace + "color");
        XElement? underlineElement = properties?.Element(WordprocessingNamespace + "u");
        string? underlineValue = (string?)underlineElement?.Attribute(WordprocessingNamespace + "val");
        string? underlineColor = (string?)underlineElement?.Attribute(WordprocessingNamespace + "color");
        bool? underline = underlineElement is not null
            ? !string.Equals(underlineValue, "none", StringComparison.OrdinalIgnoreCase)
            : null;
        return new DocxResolvedRunProperties(fontSize, color, fontFamily, bold, italic, complexScriptBold, complexScriptItalic, underline, underlineValue, ReadRunFonts(properties), characterSpacingPoints, allCaps, verticalAlignmentValue, strike, strikeValue, doubleStrike, doubleStrikeValue, highlightValue, shadingFill, shadingValue, shadingColor, smallCaps, smallCapsValue, hidden, hiddenValue, underlineColor);
    }

    private static DocxRunFonts ReadRunFonts(XElement? properties)
    {
        XElement? fonts = properties?.Element(WordprocessingNamespace + "rFonts");
        return fonts is null
            ? DocxRunFonts.Empty
            : new DocxRunFonts(
                (string?)fonts.Attribute(WordprocessingNamespace + "ascii"),
                (string?)fonts.Attribute(WordprocessingNamespace + "hAnsi"),
                (string?)fonts.Attribute(WordprocessingNamespace + "eastAsia"),
                (string?)fonts.Attribute(WordprocessingNamespace + "cs"),
                (string?)fonts.Attribute(WordprocessingNamespace + "asciiTheme"),
                (string?)fonts.Attribute(WordprocessingNamespace + "hAnsiTheme"),
                (string?)fonts.Attribute(WordprocessingNamespace + "eastAsiaTheme"),
                (string?)fonts.Attribute(WordprocessingNamespace + "csTheme"));
    }

    private static bool? ReadOnOff(XElement? element)
    {
        if (element is null)
        {
            return null;
        }

        return OoxBoolean.ParseElement(element, false, valueAttributeName: WordprocessingNamespace + "val");
    }

    private static double? ReadTwipsAttribute(XElement? element, XName name)
    {
        return element?.Attribute(name) is { } value
            ? OoxUnits.TwipsToPoints(long.Parse(value.Value, CultureInfo.InvariantCulture))
            : null;
    }

    private static double? ReadSignedTwipsElement(XElement? element)
    {
        return int.TryParse((string?)element?.Attribute(WordprocessingNamespace + "val"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int twips)
            ? OoxUnits.TwipsToPoints(twips)
            : null;
    }

    private static double? ReadLogicalStartTwips(XElement? indent)
    {
        return ReadTwipsAttribute(indent, WordprocessingNamespace + "start") ??
            ReadTwipsAttribute(indent, WordprocessingNamespace + "left");
    }

    private static double? ReadLogicalEndTwips(XElement? indent)
    {
        return ReadTwipsAttribute(indent, WordprocessingNamespace + "end") ??
            ReadTwipsAttribute(indent, WordprocessingNamespace + "right");
    }

    private static string? ReadLogicalStartValue(XElement? indent)
    {
        return (string?)indent?.Attribute(WordprocessingNamespace + "start") ??
            (string?)indent?.Attribute(WordprocessingNamespace + "left");
    }

    private static string? ReadLogicalEndValue(XElement? indent)
    {
        return (string?)indent?.Attribute(WordprocessingNamespace + "end") ??
            (string?)indent?.Attribute(WordprocessingNamespace + "right");
    }

    private static bool HasCharacterUnitIndent(XElement indent)
    {
        return indent.Attribute(WordprocessingNamespace + "leftChars") is not null ||
            indent.Attribute(WordprocessingNamespace + "rightChars") is not null ||
            indent.Attribute(WordprocessingNamespace + "startChars") is not null ||
            indent.Attribute(WordprocessingNamespace + "endChars") is not null ||
            indent.Attribute(WordprocessingNamespace + "firstLineChars") is not null ||
            indent.Attribute(WordprocessingNamespace + "hangingChars") is not null;
    }

    private static bool HasUnsupportedNumberingIndent(XElement level)
    {
        XElement? indent = level
            .Element(WordprocessingNamespace + "pPr")
            ?.Element(WordprocessingNamespace + "ind");
        return indent is not null && HasCharacterUnitIndent(indent);
    }

    private static bool HasUnsupportedParagraphSpacingVariant(XElement spacing)
    {
        string? lineRule = (string?)spacing.Attribute(WordprocessingNamespace + "lineRule");
        return lineRule is not null &&
            !lineRule.Equals("auto", StringComparison.OrdinalIgnoreCase) &&
            !lineRule.Equals("exact", StringComparison.OrdinalIgnoreCase) &&
            !lineRule.Equals("atLeast", StringComparison.OrdinalIgnoreCase);
    }

    private static int? ReadPositiveIntAttribute(XElement? element, XName name)
    {
        if (element?.Attribute(name) is not { } value ||
            !int.TryParse(value.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) ||
            parsed <= 0)
        {
            return null;
        }

        return parsed;
    }

    private static double ReadMargin(XElement? margins, XName name, double defaultValue)
    {
        return margins?.Attribute(name) is { } margin
            ? OoxUnits.TwipsToPoints(long.Parse(margin.Value, CultureInfo.InvariantCulture))
            : defaultValue;
    }

    private sealed record DocxStyleSet(
        DocxResolvedRunProperties RunDefaults,
        DocxResolvedParagraphProperties ParagraphDefaults,
        IReadOnlyDictionary<string, DocxStyle> ParagraphStyles,
        IReadOnlyDictionary<string, DocxStyle> CharacterStyles,
        IReadOnlyDictionary<string, DocxTableStyle> TableStyles,
        string? DefaultTableStyleId,
        DocxTableStyle? DefaultTableStyle)
    {
        public static DocxStyleSet Empty { get; } = new(
            new DocxResolvedRunProperties(null, null, null, null, null, null, null, null, null, DocxRunFonts.Empty, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null),
            new DocxResolvedParagraphProperties(null, null, null, null, null, null, DocxParagraphSpacing.Empty, DocxParagraphKeepRules.Empty, DocxParagraphIndent.Empty, [], null, null, null, null, null, null),
            new Dictionary<string, DocxStyle>(),
            new Dictionary<string, DocxStyle>(),
            new Dictionary<string, DocxTableStyle>(),
            null,
            null);
    }

    private sealed record DocxStyle(string? BasedOnStyleId, DocxResolvedParagraphProperties Paragraph, DocxResolvedRunProperties Run);

    private sealed record ParagraphBreakPart(
        XElement? Paragraph,
        string? BreakValue,
        bool StartsAfterBreak,
        bool EndsBeforeBreak);

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
                MergeTableCellMargins(other.Margins, Margins),
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

    private static DocxTextRunStyle ReadTextRunStyle(XElement? properties)
    {
        DocxResolvedRunProperties run = ReadRunProperties(properties);
        return new DocxTextRunStyle(
            run.FontSize,
            run.ColorHex,
            run.Bold,
            run.Italic,
            run.Underline,
            run.UnderlineValue,
            run.FontFamily,
            run.Fonts,
            run.CharacterSpacingPoints,
            run.AllCaps,
            run.VerticalAlignmentValue,
            run.Strike,
            run.StrikeValue,
            run.DoubleStrike,
            run.DoubleStrikeValue,
            run.HighlightValue,
            run.ShadingFillHex,
            run.ShadingValue,
            run.ShadingColor,
            run.SmallCaps,
            run.SmallCapsValue,
            run.Hidden,
            run.HiddenValue,
            run.UnderlineColorHex);
    }

    private readonly record struct DocxResolvedParagraphProperties(
        DocxTextAlignment? Alignment,
        string? AlignmentValue,
        double? SpacingBeforePoints,
        double? SpacingAfterPoints,
        double? LineSpacingFactor,
        double? LineSpacingPoints,
        DocxParagraphSpacing Spacing,
        DocxParagraphKeepRules KeepRules,
        DocxParagraphIndent Indent,
        IReadOnlyList<DocxTabStop> TabStops,
        bool? SnapToGrid,
        string? SnapToGridValue,
        bool? PageBreakBefore,
        string? PageBreakBeforeValue,
        bool? WordWrap,
        string? WordWrapValue)
    {
        public static DocxResolvedParagraphProperties Empty { get; } = new(null, null, null, null, null, null, DocxParagraphSpacing.Empty, DocxParagraphKeepRules.Empty, DocxParagraphIndent.Empty, [], null, null, null, null, null, null);

        public DocxResolvedParagraphProperties Merge(DocxResolvedParagraphProperties other)
        {
            bool hasOtherBeforeSide = DocxParagraphSpacing.HasBeforeSpacingSide(other.Spacing);
            bool hasOtherAfterSide = DocxParagraphSpacing.HasAfterSpacingSide(other.Spacing);
            return new DocxResolvedParagraphProperties(
                other.Alignment ?? Alignment,
                other.AlignmentValue ?? AlignmentValue,
                hasOtherBeforeSide ? other.SpacingBeforePoints : other.SpacingBeforePoints ?? SpacingBeforePoints,
                hasOtherAfterSide ? other.SpacingAfterPoints : other.SpacingAfterPoints ?? SpacingAfterPoints,
                other.LineSpacingFactor ?? LineSpacingFactor,
                other.LineSpacingPoints ?? LineSpacingPoints,
                Spacing.Merge(other.Spacing),
                KeepRules.Merge(other.KeepRules),
                Indent.Merge(other.Indent),
                other.TabStops.Count != 0 ? other.TabStops : TabStops,
                other.SnapToGrid ?? SnapToGrid,
                other.SnapToGridValue ?? SnapToGridValue,
                other.PageBreakBefore ?? PageBreakBefore,
                other.PageBreakBeforeValue ?? PageBreakBeforeValue,
                other.WordWrap ?? WordWrap,
                other.WordWrapValue ?? WordWrapValue);
        }
    }
}