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
    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static string FormatNoteReferenceNumber(int value, string? format)
    {
        return format switch
        {
            "lowerRoman" => ToRomanNumeral(value).ToLowerInvariant(),
            "upperRoman" => ToRomanNumeral(value),
            "lowerLetter" => ToAlphabeticNumber(value, upper: false),
            "upperLetter" => ToAlphabeticNumber(value, upper: true),
            _ => value.ToString(CultureInfo.InvariantCulture)
        };
    }

    // Single caller; loop core shared via OoxNumbering, per-spec edge guard stays.
    private static string ToAlphabeticNumber(int value, bool upper)
    {
        if (value <= 0)
        {
            return value.ToString(CultureInfo.InvariantCulture);
        }

        return OoxNumbering.ToAlphabetic(value, upper);
    }

    // Single caller; loop core shared via OoxNumbering, per-spec edge guard stays.
    private static string ToRomanNumeral(int value)
    {
        if (value <= 0 || value > 3999)
        {
            return value.ToString(CultureInfo.InvariantCulture);
        }

        return OoxNumbering.ToRomanUpper(value);
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static DocxListLabel? CreateListLabel(XElement? paragraphProperties, DocxNumberingSet numbering, Dictionary<(string NumId, int Level), int> counters, DocxStyleSet styles, string? paragraphStyleId, DocxTableCellStyle? tableCellStyle)
    {
        XElement? numberingProperties = paragraphProperties?.Element(WordprocessingNamespace + "numPr");
        string? numId = (string?)numberingProperties?
            .Element(WordprocessingNamespace + "numId")
            ?.Attribute(WordprocessingNamespace + "val");
        int level = numberingProperties?
            .Element(WordprocessingNamespace + "ilvl")
            ?.Attribute(WordprocessingNamespace + "val") is { } levelAttribute
            ? int.Parse(levelAttribute.Value, CultureInfo.InvariantCulture)
            : 0;
        if (numId is null || !numbering.NumToAbstract.TryGetValue(numId, out string? abstractId))
        {
            return null;
        }

        DocxNumberingLevel? numberingLevel = numbering.LevelOverrides.TryGetValue((numId, level), out DocxNumberingLevel? concreteLevel)
            ? concreteLevel
            : numbering.Levels.TryGetValue((abstractId, level), out DocxNumberingLevel? abstractLevel)
                ? abstractLevel
                : null;
        if (numberingLevel is null)
        {
            return null;
        }

        DocxTextRunStyle labelStyle = ResolveListLabelStyle(numberingLevel.Style, paragraphProperties, styles, paragraphStyleId, tableCellStyle);
        if (numberingLevel.Format.Equals("bullet", StringComparison.OrdinalIgnoreCase))
        {
            string bulletText = string.IsNullOrEmpty(numberingLevel.Text) ? "\u2022" : numberingLevel.Text;
            return new DocxListLabel(bulletText, numberingLevel.Format, numberingLevel.Text, numberingLevel.Suffix, numId, level, numberingLevel.Indent, labelStyle);
        }

        var key = (numId, level);
        int start = numbering.StartOverrides.TryGetValue((numId, level), out int overriddenStart)
            ? overriddenStart
            : numberingLevel.Start;
        counters[key] = counters.TryGetValue(key, out int current) ? current + 1 : start;
        foreach (var resetKey in counters.Keys.Where(k => k.NumId == numId && k.Level > level).ToArray())
        {
            counters.Remove(resetKey);
        }

        string labelText = ResolveNumberingLevelText(numberingLevel.Text, numId, counters);
        return new DocxListLabel(labelText, numberingLevel.Format, numberingLevel.Text, numberingLevel.Suffix, numId, level, numberingLevel.Indent, labelStyle);
    }

    private const double WordListLabelFallbackMaxPoints = 12d;

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static DocxTextRunStyle ResolveListLabelStyle(DocxTextRunStyle levelStyle, XElement? paragraphProperties, DocxStyleSet styles, string? paragraphStyleId, DocxTableCellStyle? tableCellStyle)
    {
        // Office sizes bullets from the numbering level, then paragraph/table styles (style-less paragraphs
        // resolve through Normal), then unstyled table cells fall back to 12pt, then document defaults outright,
        // and otherwise flat 12pt (bullet-chain probes 2026-09-06: style 14pt wins over direct 10/18pt;
        // docDefaults 8/20pt win over direct body runs; unstyled table cells emit 12pt bullets at direct 9/10/14/18pt;
        // labsize/nosize probes 2026-09-07: unresolvable style and defaults emit flat 12pt labels at direct 9/14/18pt
        // for family-less and Symbol lvls alike, so direct size never refines the label;
        // cells-with-docDefaults untested, cascade order assumed).
        if (levelStyle.FontSize is not null)
        {
            return levelStyle;
        }

        XElement? markRunProperties = paragraphProperties?.Element(WordprocessingNamespace + "rPr");
        if (ReadRunProperties(markRunProperties).FontSize is { } markSize)
        {
            return levelStyle with { FontSize = markSize };
        }

        double? styleSize = tableCellStyle?.Run.FontSize;
        foreach (DocxStyle paragraphStyle in EnumerateStyleInheritance(paragraphStyleId ?? "Normal", styles.ParagraphStyles))
        {
            styleSize = paragraphStyle.Run.FontSize ?? styleSize;
        }

        if (styleSize is not null)
        {
            return levelStyle with { FontSize = styleSize };
        }

        if (tableCellStyle is not null)
        {
            return levelStyle with { FontSize = WordListLabelFallbackMaxPoints };
        }

        if (styles.RunDefaults.FontSize is { } defaultsSize)
        {
            return levelStyle with { FontSize = defaultsSize };
        }

        // No resolvable level, mark, style, cell, or default size: Word falls back to flat 12pt
        // (labsize/nosize probes 2026-09-07: styles-less and styles-without-size-or-defaults emit 12pt labels
        // at direct 9/14/18pt for family-less and Symbol lvls alike; direct size never refines the label).
        return levelStyle with { FontSize = WordListLabelFallbackMaxPoints };
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static string ResolveNumberingLevelText(string text, string numId, IReadOnlyDictionary<(string NumId, int Level), int> counters)
    {
        string resolved = text;
        for (int level = 0; level < 9; level++)
        {
            string token = "%" + (level + 1).ToString(CultureInfo.InvariantCulture);
            if (resolved.Contains(token, StringComparison.Ordinal))
            {
                string value = counters.TryGetValue((numId, level), out int counter)
                    ? counter.ToString(CultureInfo.InvariantCulture)
                    : "0";
                resolved = resolved.Replace(token, value, StringComparison.Ordinal);
            }
        }

        return resolved;
    }

    // Single caller; kept static: entry-point stage, not a local candidate.
    private static DocxNumberingSet LoadNumbering(OoxPackage package, string documentPartName, DocxFontCatalog fontCatalog, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxRelationship? numberingRelationship = package.GetRelationships(documentPartName, cancellationToken)
            .FirstOrDefault(r => !r.IsExternal && r.Type == NumberingRelationshipType && r.ResolvedTarget is not null);
        OoxPart? numberingPart = numberingRelationship?.ResolvedTarget is null
            ? package.Parts.FirstOrDefault(p => p.ContentType == NumberingContentType)
            : package.GetPart(numberingRelationship.ResolvedTarget);
        if (numberingPart is null)
        {
            return DocxNumberingSet.Empty;
        }

        using Stream stream = numberingPart.OpenRead();
        XDocument numberingXml = SafeXml.Load(stream, cancellationToken);
        var levels = new Dictionary<(string AbstractId, int Level), DocxNumberingLevel>();
        foreach (XElement abstractNum in numberingXml.Root?.Elements(WordprocessingNamespace + "abstractNum") ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();
            string? abstractId = (string?)abstractNum.Attribute(WordprocessingNamespace + "abstractNumId");
            if (abstractId is null)
            {
                continue;
            }

            foreach (XElement level in abstractNum.Elements(WordprocessingNamespace + "lvl"))
            {
                cancellationToken.ThrowIfCancellationRequested();
                int levelIndex = level.Attribute(WordprocessingNamespace + "ilvl") is { } ilvl
                    ? int.Parse(ilvl.Value, CultureInfo.InvariantCulture)
                    : 0;
                levels[(abstractId, levelIndex)] = ReadNumberingLevel(level, levelIndex, fontCatalog);
            }
        }

        var numToAbstract = new Dictionary<string, string>(StringComparer.Ordinal);
        var startOverrides = new Dictionary<(string NumId, int Level), int>();
        var levelOverrides = new Dictionary<(string NumId, int Level), DocxNumberingLevel>();
        foreach (XElement num in numberingXml.Root?.Elements(WordprocessingNamespace + "num") ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();
            string? numId = (string?)num.Attribute(WordprocessingNamespace + "numId");
            string? abstractId = (string?)num.Element(WordprocessingNamespace + "abstractNumId")?.Attribute(WordprocessingNamespace + "val");
            if (numId is not null && abstractId is not null)
            {
                numToAbstract[numId] = abstractId;
            }

            if (numId is null)
            {
                continue;
            }

            foreach (XElement overrideLevel in num.Elements(WordprocessingNamespace + "lvlOverride"))
            {
                cancellationToken.ThrowIfCancellationRequested();
                int levelIndex = overrideLevel.Attribute(WordprocessingNamespace + "ilvl") is { } ilvl
                    ? int.Parse(ilvl.Value, CultureInfo.InvariantCulture)
                    : 0;
                if (overrideLevel.Element(WordprocessingNamespace + "startOverride")?.Attribute(WordprocessingNamespace + "val") is { } startValue)
                {
                    startOverrides[(numId, levelIndex)] = int.Parse(startValue.Value, CultureInfo.InvariantCulture);
                }

                XElement? concreteLevel = overrideLevel.Element(WordprocessingNamespace + "lvl");
                if (concreteLevel is not null)
                {
                    levelOverrides[(numId, levelIndex)] = ReadNumberingLevel(concreteLevel, levelIndex, fontCatalog);
                }
            }
        }

        return new DocxNumberingSet(numToAbstract, levels, startOverrides, levelOverrides);
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static DocxNumberingLevel ReadNumberingLevel(XElement level, int levelIndex, DocxFontCatalog fontCatalog)
    {
        int start = level.Element(WordprocessingNamespace + "start")?.Attribute(WordprocessingNamespace + "val") is { } startValue
            ? int.Parse(startValue.Value, CultureInfo.InvariantCulture)
            : 1;
        string format = (string?)level.Element(WordprocessingNamespace + "numFmt")?.Attribute(WordprocessingNamespace + "val") ?? "decimal";
        string text = (string?)level.Element(WordprocessingNamespace + "lvlText")?.Attribute(WordprocessingNamespace + "val") ??
            (format.Equals("bullet", StringComparison.OrdinalIgnoreCase) ? "\u2022" : "%" + (levelIndex + 1) + ".");
        string suffix = (string?)level.Element(WordprocessingNamespace + "suff")?.Attribute(WordprocessingNamespace + "val") ?? "tab";
        DocxTextRunStyle style = ReadTextRunStyle(level.Element(WordprocessingNamespace + "rPr"));
        return new DocxNumberingLevel(format, ResolveNumberingSymbolText(text, style, fontCatalog), suffix, start, ReadNumberingIndent(level), style);
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static string ResolveNumberingSymbolText(string text, DocxTextRunStyle style, DocxFontCatalog fontCatalog)
    {
        return UsesSymbolCharset(style, fontCatalog) ? MapSymbolCharsetText(text) : text;
    }

    // Single caller; kept static: charset-probe pair kept together.
    private static bool UsesSymbolCharset(DocxTextRunStyle style, DocxFontCatalog fontCatalog)
    {
        string? family = FirstNonEmpty(style.Fonts.Ascii, style.Fonts.HighAnsi, style.FontFamily, style.Fonts.ComplexScript);
        if (family is null)
        {
            return false;
        }

        DocxFontTableEntry? entry = fontCatalog.Entries
            .FirstOrDefault(item => item.Name.Equals(family, StringComparison.OrdinalIgnoreCase));
        return entry?.CharsetValue is { } charset &&
            (charset.Equals("02", StringComparison.OrdinalIgnoreCase) || charset.Equals("2", StringComparison.OrdinalIgnoreCase));
    }

    // Single caller; kept static: charset-probe pair kept together.
    private static string MapSymbolCharsetText(string text)
    {
        Span<char> mapped = text.Length <= 256
            ? stackalloc char[text.Length]
            : new char[text.Length];
        for (int i = 0; i < text.Length; i++)
        {
            char ch = text[i];
            mapped[i] = ch <= 0x00FF ? (char)(0xF000 + ch) : ch;
        }

        return new string(mapped);
    }

    private sealed record DocxNumberingSet(
        IReadOnlyDictionary<string, string> NumToAbstract,
        IReadOnlyDictionary<(string AbstractId, int Level), DocxNumberingLevel> Levels,
        IReadOnlyDictionary<(string NumId, int Level), int> StartOverrides,
        IReadOnlyDictionary<(string NumId, int Level), DocxNumberingLevel> LevelOverrides)
    {
        public static DocxNumberingSet Empty { get; } = new(
            new Dictionary<string, string>(),
            new Dictionary<(string AbstractId, int Level), DocxNumberingLevel>(),
            new Dictionary<(string NumId, int Level), int>(),
            new Dictionary<(string NumId, int Level), DocxNumberingLevel>());
    }

    private sealed record DocxNumberingLevel(string Format, string Text, string Suffix, int Start, DocxNumberingIndent Indent, DocxTextRunStyle Style);

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static DocxNumberingIndent ReadNumberingIndent(XElement level)
    {
        XElement? indent = level
            .Element(WordprocessingNamespace + "pPr")
            ?.Element(WordprocessingNamespace + "ind");
        XElement? numberingTab = level
            .Element(WordprocessingNamespace + "pPr")
            ?.Element(WordprocessingNamespace + "tabs")
            ?.Elements(WordprocessingNamespace + "tab")
            .FirstOrDefault(tab => string.Equals(
                (string?)tab.Attribute(WordprocessingNamespace + "val"),
                "num",
                StringComparison.OrdinalIgnoreCase));
        return new DocxNumberingIndent(
            ReadLogicalStartTwips(indent),
            ReadLogicalEndTwips(indent),
            ReadTwipsAttribute(indent, WordprocessingNamespace + "firstLine"),
            ReadTwipsAttribute(indent, WordprocessingNamespace + "hanging"),
            ReadTwipsAttribute(numberingTab, WordprocessingNamespace + "pos"),
            ReadLogicalStartValue(indent),
            ReadLogicalEndValue(indent),
            (string?)indent?.Attribute(WordprocessingNamespace + "firstLine"),
            (string?)indent?.Attribute(WordprocessingNamespace + "hanging"),
            (string?)numberingTab?.Attribute(WordprocessingNamespace + "val"),
            (string?)numberingTab?.Attribute(WordprocessingNamespace + "pos"));
    }
}
