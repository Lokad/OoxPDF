using System.Globalization;
using System.Text;
using System.Xml.Linq;

using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static IReadOnlyList<TextRun> ReadTextRuns(XDocument slideXml, PptxDocument document, PptxTheme theme, int slideNumber, bool includePlaceholders, IReadOnlyList<XDocument> placeholderSources)
    {
        var runs = new List<TextRun>();
        var advanceEstimator = new TextAdvanceEstimator();
        foreach (XElement shape in slideXml.Descendants(PresentationNamespace + "sp"))
        {
            if (!includePlaceholders && IsPlaceholder(shape))
            {
                continue;
            }

            XElement? shapeProperties = shape.Element(PresentationNamespace + "spPr");
            XElement? textBody = shape.Element(PresentationNamespace + "txBody");
            XElement? inheritedPlaceholder = FindInheritedPlaceholderShape(shape, placeholderSources);
            XElement? inheritedTextBody = inheritedPlaceholder?.Element(PresentationNamespace + "txBody");
            ShapeBounds? bounds = shapeProperties is null ? null : ReadBounds(shapeProperties);
            bounds ??= inheritedPlaceholder?.Element(PresentationNamespace + "spPr") is { } inheritedProperties
                ? ReadBounds(inheritedProperties)
                : null;
            if (bounds is null || textBody is null)
            {
                continue;
            }

            bounds = ReadAncestorGroupTransform(shape).Apply(bounds.Value);
            double x = OoxUnits.EmuToPoints(bounds.Value.X);
            double yTop = OoxUnits.EmuToPoints(bounds.Value.Y);
            double width = OoxUnits.EmuToPoints(bounds.Value.Width);
            double height = OoxUnits.EmuToPoints(bounds.Value.Height);
            TextInsets insets = ReadTextInsets(textBody);
            double textX = x + insets.Left;
            double textWidth = Math.Max(1d, width - insets.Left - insets.Right);
            double textHeight = Math.Max(1d, height - insets.Top - insets.Bottom);
            double rotationCenterX = x + width / 2d;
            double rotationCenterY = document.SlideHeightPoints - yTop - height / 2d;
            bool clipsVerticalOverflow = ClipsVerticalOverflow(textBody);
            double textClipY = clipsVerticalOverflow
                ? document.SlideHeightPoints - yTop - insets.Top - textHeight
                : 0d;
            double textClipHeight = clipsVerticalOverflow
                ? textHeight
                : document.SlideHeightPoints;
            RgbColor? shapeFontColor = TryReadShapeFontColor(shape, theme, out RgbColor fontColor)
                ? fontColor
                : null;
            XElement? defaultParagraphProperties = MergeParagraphProperties(
                textBody.Element(DrawingNamespace + "lstStyle")?.Element(DrawingNamespace + "lvl1pPr"),
                inheritedTextBody?.Element(DrawingNamespace + "lstStyle")?.Element(DrawingNamespace + "lvl1pPr"),
                FindInheritedTextStyle(shape, placeholderSources));
            double verticalOffset = ReadVerticalAnchor(textBody) switch
            {
                TextVerticalAnchor.Top when inheritedTextBody is not null => ReadVerticalAnchor(inheritedTextBody) switch
                {
                    TextVerticalAnchor.Middle => Math.Max(0d, (textHeight - EstimateTextHeight(textBody, defaultParagraphProperties)) / 2d),
                    TextVerticalAnchor.Bottom => Math.Max(0d, textHeight - EstimateTextHeight(textBody, defaultParagraphProperties)),
                    _ => 0d
                },
                TextVerticalAnchor.Middle => Math.Max(0d, (textHeight - EstimateTextHeight(textBody, defaultParagraphProperties)) / 2d),
                TextVerticalAnchor.Bottom => Math.Max(0d, textHeight - EstimateTextHeight(textBody, defaultParagraphProperties)),
                _ => 0d
            };
            double cursorLineTop = document.SlideHeightPoints - yTop - insets.Top - verticalOffset;
            int autoNumberValue = 1;

            foreach (XElement paragraph in textBody.Elements(DrawingNamespace + "p"))
            {
                XElement? paragraphProperties = paragraph.Element(DrawingNamespace + "pPr");
                XElement? defaultRunProperties = paragraphProperties?.Element(DrawingNamespace + "defRPr") ??
                    defaultParagraphProperties?.Element(DrawingNamespace + "defRPr");
                double paragraphFontSize = ReadFirstParagraphFontSize(paragraph, defaultRunProperties);
                double spacingBefore = ReadParagraphSpacing(paragraphProperties, defaultParagraphProperties, "spcBef", paragraphFontSize);
                double spacingAfter = ReadParagraphSpacing(paragraphProperties, defaultParagraphProperties, "spcAft", paragraphFontSize);
                LineSpacing lineSpacing = ReadLineSpacing(paragraphProperties, defaultParagraphProperties);
                if (!ParagraphHasVisibleContent(paragraph))
                {
                    if (ParagraphHasLayoutContent(paragraph))
                    {
                        XElement? endRunProperties = paragraph.Element(DrawingNamespace + "endParaRPr");
                        double emptyFontSize = ReadFontSize(endRunProperties, defaultRunProperties);
                        double emptySpacingBefore = ReadParagraphSpacing(paragraphProperties, defaultParagraphProperties, "spcBef", emptyFontSize);
                        double emptySpacingAfter = ReadParagraphSpacing(paragraphProperties, defaultParagraphProperties, "spcAft", emptyFontSize);
                        cursorLineTop -= emptySpacingBefore + ReadParagraphAdvance(lineSpacing, emptyFontSize) + emptySpacingAfter;
                    }

                    continue;
                }

                TextAlignment alignment = ReadAlignment(paragraph, defaultParagraphProperties);
                string? bulletText = ReadBulletText(paragraphProperties, ref autoNumberValue);
                bool bulletPending = bulletText is not null;
                ParagraphIndent indent = ReadParagraphIndent(paragraphProperties);
                IReadOnlyList<double> tabStops = ReadTabStops(paragraphProperties);
                double bulletX = textX + Math.Max(0d, indent.MarginLeft + indent.Hanging);
                double paragraphTextX = bulletText is null
                    ? textX + Math.Max(0d, indent.MarginLeft + indent.Hanging)
                    : textX + Math.Max(0d, indent.MarginLeft);
                cursorLineTop -= spacingBefore;
                double cursorY = cursorLineTop - ReadFirstLineBaselineOffset(paragraph, defaultRunProperties, lineSpacing);
                double cursorX = paragraphTextX;
                double maxFontSize = paragraphFontSize;
                var paragraphRuns = new List<TextRun>();
                double paragraphEndX = paragraphTextX;
                foreach (XElement child in paragraph.Elements())
                {
                    if (child.Name == DrawingNamespace + "br")
                    {
                        AddAlignedParagraphRuns(runs, paragraphRuns, alignment, textX, textWidth, paragraphEndX);
                        paragraphRuns.Clear();
                        cursorLineTop -= ReadLineAdvance(lineSpacing, maxFontSize);
                        cursorY = double.NaN;
                        cursorX = paragraphTextX;
                        paragraphEndX = paragraphTextX;
                        maxFontSize = paragraphFontSize;
                        continue;
                    }

                    if (!IsTextRunElement(child))
                    {
                        continue;
                    }

                    XElement run = child;
                    XElement? runProperties = run.Element(DrawingNamespace + "rPr");
                    double nominalFontSize = ReadFontSize(runProperties, defaultRunProperties);
                    maxFontSize = Math.Max(maxFontSize, nominalFontSize);
                    double baselineOffset = ReadBaselineOffset(runProperties, defaultRunProperties, nominalFontSize);
                    if (double.IsNaN(cursorY))
                    {
                        cursorY = cursorLineTop - LineBaselineOffset(nominalFontSize, lineSpacing);
                    }

                    double fontSize = Math.Abs(baselineOffset) > 0.001d
                        ? nominalFontSize * 2d / 3d
                        : nominalFontSize;
                    double alpha = 1d;
                    RgbColor color;
                    if (TryReadSolidColorWithAlpha(runProperties, theme, out RgbColor runColor, out double runAlpha))
                    {
                        color = runColor;
                        alpha = runAlpha;
                    }
                    else if (TryReadSolidColorWithAlpha(defaultRunProperties, theme, out RgbColor defaultColor, out double defaultAlpha))
                    {
                        color = defaultColor;
                        alpha = defaultAlpha;
                    }
                    else
                    {
                        color = shapeFontColor ?? new RgbColor(0, 0, 0);
                    }
                    string? typeface = theme.ResolveTypeface((string?)(runProperties?
                        .Element(DrawingNamespace + "latin") ??
                        defaultRunProperties?.Element(DrawingNamespace + "latin"))
                        ?.Attribute("typeface"));
                    bool bold = ParseOptionalBoolAttribute(runProperties, "b") ||
                        (runProperties?.Attribute("b") is null && ParseOptionalBoolAttribute(defaultRunProperties, "b"));
                    bool italic = ParseOptionalBoolAttribute(runProperties, "i") ||
                        (runProperties?.Attribute("i") is null && ParseOptionalBoolAttribute(defaultRunProperties, "i"));
                    bool underline = ((string?)(runProperties?.Attribute("u") ?? defaultRunProperties?.Attribute("u"))) is { } underlineValue
                        && !underlineValue.Equals("none", StringComparison.OrdinalIgnoreCase);
                    bool strike = IsStrikeEnabled(runProperties, defaultRunProperties);
                    double characterSpacing = ReadCharacterSpacing(runProperties, defaultRunProperties);
                    RgbColor? highlight = TryReadHighlightColor(runProperties, out RgbColor highlightColor)
                        ? highlightColor
                        : null;
                    if (bulletPending)
                    {
                        BulletStyle bulletStyle = ReadBulletStyle(paragraphProperties, theme, fontSize, color, typeface);
                        double bulletWidth = Math.Max(1d, textWidth - (bulletX - textX));
                        paragraphRuns.Add(new TextRun(bulletText!, bulletX, cursorY, bulletWidth, textHeight, textX, textClipY, textWidth, textClipHeight, bulletStyle.FontSize, characterSpacing, 0d, bulletStyle.Color, 1d, null, bold, italic, underline, strike, alignment, bulletStyle.Typeface, bounds.Value.RotationDegrees, rotationCenterX, rotationCenterY));
                        paragraphEndX = Math.Max(paragraphEndX, bulletX + advanceEstimator.Measure(bulletText!, bulletStyle.FontSize, bulletStyle.Typeface, bold, italic, characterSpacing));
                        bulletPending = false;
                    }

                    string rawText = ReadTextElementText(run, slideNumber);
                    string[] tabParts = rawText.Split('\t');
                    for (int tabPartIndex = 0; tabPartIndex < tabParts.Length; tabPartIndex++)
                    {
                        if (tabPartIndex > 0)
                        {
                            cursorX = ResolveNextTabX(cursorX, paragraphTextX, tabStops, fontSize);
                            paragraphEndX = Math.Max(paragraphEndX, cursorX);
                        }

                        foreach (TextCapsFragment fragment in ApplyTextCaps(tabParts[tabPartIndex], runProperties, defaultRunProperties))
                        {
                            if (fragment.Text.Length == 0)
                            {
                                continue;
                            }

                            double fragmentFontSize = fontSize * fragment.FontScale;
                            foreach (string segment in SplitFlowSegments(fragment.Text))
                            {
                                string currentSegment = segment;
                                double segmentWidth = advanceEstimator.Measure(currentSegment, fragmentFontSize, typeface, bold, italic, characterSpacing);
                                bool overflowsLine = cursorX > paragraphTextX &&
                                    (cursorX + segmentWidth > textX + textWidth ||
                                        (characterSpacing > 0d && cursorX + segmentWidth > textX + textWidth - fragmentFontSize));
                                if (overflowsLine)
                                {
                                    AddAlignedParagraphRuns(runs, paragraphRuns, alignment, textX, textWidth, paragraphEndX);
                                    paragraphRuns.Clear();
                                    cursorLineTop -= ReadLineAdvance(lineSpacing, maxFontSize);
                                    cursorY = cursorLineTop - LineBaselineOffset(fragmentFontSize, lineSpacing);
                                    cursorX = paragraphTextX;
                                    paragraphEndX = paragraphTextX;
                                    maxFontSize = Math.Max(nominalFontSize, fragmentFontSize);
                                    currentSegment = currentSegment.TrimStart();
                                    segmentWidth = advanceEstimator.Measure(currentSegment, fragmentFontSize, typeface, bold, italic, characterSpacing);
                                }

                                if (currentSegment.Length == 0)
                                {
                                    continue;
                                }

                                paragraphRuns.Add(new TextRun(currentSegment, cursorX, cursorY, Math.Max(1d, segmentWidth), textHeight, textX, textClipY, textWidth, textClipHeight, fragmentFontSize, characterSpacing, baselineOffset, color, alpha, highlight, bold, italic, underline, strike, alignment, typeface, bounds.Value.RotationDegrees, rotationCenterX, rotationCenterY));
                                cursorX += segmentWidth;
                                paragraphEndX = Math.Max(paragraphEndX, cursorX);
                            }
                        }
                    }
                }

                AddAlignedParagraphRuns(runs, paragraphRuns, alignment, textX, textWidth, paragraphEndX);
                cursorLineTop -= ReadParagraphAdvance(lineSpacing, maxFontSize) + spacingAfter;
            }
        }

        return runs;
    }

    private static GroupTransform ReadAncestorGroupTransform(XElement shape)
    {
        GroupTransform transform = GroupTransform.Identity;
        foreach (XElement group in shape.Ancestors(PresentationNamespace + "grpSp").Reverse())
        {
            transform = transform.Combine(ReadGroupTransform(group));
        }

        return transform;
    }

    private static IReadOnlyList<TextRun> ReadTextRunsForShape(
        XElement shape,
        PptxDocument document,
        PptxTheme theme,
        int slideNumber,
        bool includePlaceholders,
        IReadOnlyList<XDocument> placeholderSources)
    {
        XElement current = new(shape);
        foreach (XElement group in shape.Ancestors(PresentationNamespace + "grpSp"))
        {
            var groupCopy = new XElement(PresentationNamespace + "grpSp");
            if (group.Element(PresentationNamespace + "grpSpPr") is { } properties)
            {
                groupCopy.Add(new XElement(properties));
            }

            groupCopy.Add(current);
            current = groupCopy;
        }

        var slide = new XDocument(
            new XElement(PresentationNamespace + "sld",
                new XElement(PresentationNamespace + "cSld",
                    new XElement(PresentationNamespace + "spTree", current))));
        return ReadTextRuns(slide, document, theme, slideNumber, includePlaceholders, placeholderSources);
    }

    private static XElement? FindInheritedPlaceholderShape(XElement shape, IReadOnlyList<XDocument> placeholderSources)
    {
        XElement? placeholder = shape
            .Element(PresentationNamespace + "nvSpPr")
            ?.Element(PresentationNamespace + "nvPr")
            ?.Element(PresentationNamespace + "ph");
        if (placeholder is null)
        {
            return null;
        }

        string? type = (string?)placeholder.Attribute("type");
        string? index = (string?)placeholder.Attribute("idx");
        foreach (XDocument source in placeholderSources.Reverse())
        {
            foreach (XElement candidate in source.Descendants(PresentationNamespace + "sp"))
            {
                XElement? candidatePlaceholder = candidate
                    .Element(PresentationNamespace + "nvSpPr")
                    ?.Element(PresentationNamespace + "nvPr")
                    ?.Element(PresentationNamespace + "ph");
                if (candidatePlaceholder is null)
                {
                    continue;
                }

                string? candidateType = (string?)candidatePlaceholder.Attribute("type");
                string? candidateIndex = (string?)candidatePlaceholder.Attribute("idx");
                bool indexMatches = index is not null && candidateIndex == index;
                bool typeMatches = index is null && type is not null && candidateType == type;
                if (!indexMatches && !typeMatches)
                {
                    continue;
                }

                return candidate;
            }
        }

        return null;
    }

    private static XElement? FindInheritedTextStyle(XElement shape, IReadOnlyList<XDocument> placeholderSources)
    {
        string? placeholderType = (string?)shape
            .Element(PresentationNamespace + "nvSpPr")
            ?.Element(PresentationNamespace + "nvPr")
            ?.Element(PresentationNamespace + "ph")
            ?.Attribute("type");
        string styleName = placeholderType switch
        {
            "title" or "ctrTitle" => "titleStyle",
            "body" or "subTitle" => "bodyStyle",
            _ => "otherStyle"
        };

        foreach (XDocument source in placeholderSources)
        {
            XElement? style = source.Root?
                .Element(PresentationNamespace + "txStyles")
                ?.Element(PresentationNamespace + styleName);
            XElement? level = style?.Element(DrawingNamespace + "lvl1pPr") ??
                style?.Element(DrawingNamespace + "defPPr");
            if (level is not null)
            {
                return level;
            }
        }

        return null;
    }

    private static IEnumerable<string> SplitFlowSegments(string text)
    {
        int index = 0;
        while (index < text.Length)
        {
            int start = index;
            while (index < text.Length && text[index] == ' ')
            {
                index++;
            }

            while (index < text.Length && text[index] != ' ')
            {
                index++;
                if (text[index - 1] == '-' && index < text.Length && text[index] != ' ')
                {
                    break;
                }
            }

            while (index < text.Length && text[index] == ' ')
            {
                index++;
            }

            if (index > start)
            {
                yield return text[start..index];
            }
        }
    }

    private static void AddAlignedParagraphRuns(List<TextRun> runs, List<TextRun> paragraphRuns, TextAlignment alignment, double textX, double textWidth, double paragraphEndX)
    {
        if (paragraphRuns.Count == 0)
        {
            return;
        }

        double paragraphWidth = Math.Max(0d, paragraphEndX - textX);
        double offset = alignment switch
        {
            TextAlignment.Center => Math.Max(0d, textWidth - paragraphWidth) / 2d,
            TextAlignment.Right => Math.Max(0d, textWidth - paragraphWidth),
            _ => 0d
        };

        foreach (TextRun run in paragraphRuns)
        {
            runs.Add(run with
            {
                X = run.X + offset,
                Width = run.Width,
                Alignment = TextAlignment.Left
            });
        }
    }

    private static XElement? MergeParagraphProperties(params XElement?[] sources)
    {
        XElement? merged = null;
        foreach (XElement? source in sources.Reverse())
        {
            if (source is null)
            {
                continue;
            }

            merged = merged is null
                ? new XElement(source)
                : OverlayParagraphProperties(source, merged);
        }

        return merged;
    }

    private static XElement OverlayParagraphProperties(XElement primary, XElement fallback)
    {
        XElement merged = new(primary);
        foreach (XAttribute attribute in fallback.Attributes())
        {
            if (merged.Attribute(attribute.Name) is null)
            {
                merged.Add(new XAttribute(attribute));
            }
        }

        MergeChildElement(merged, fallback, DrawingNamespace + "defRPr");
        return merged;
    }

    private static void MergeChildElement(XElement primaryParent, XElement fallbackParent, XName childName)
    {
        XElement? primary = primaryParent.Element(childName);
        XElement? fallback = fallbackParent.Element(childName);
        if (fallback is null)
        {
            return;
        }

        if (primary is null)
        {
            primaryParent.Add(new XElement(fallback));
            return;
        }

        foreach (XAttribute attribute in fallback.Attributes())
        {
            if (primary.Attribute(attribute.Name) is null)
            {
                primary.Add(new XAttribute(attribute));
            }
        }

        foreach (XElement child in fallback.Elements())
        {
            if (primary.Element(child.Name) is null)
            {
                primary.Add(new XElement(child));
            }
        }
    }

    private static bool ParagraphHasVisibleContent(XElement paragraph)
    {
        return paragraph.Elements().Any(child =>
            child.Name == DrawingNamespace + "r" ||
            child.Name == DrawingNamespace + "fld" ||
            child.Name == DrawingNamespace + "br");
    }

    private static bool ParagraphHasLayoutContent(XElement paragraph)
    {
        return paragraph.Element(DrawingNamespace + "pPr") is not null ||
            paragraph.Element(DrawingNamespace + "endParaRPr") is not null;
    }

    private static bool IsTextRunElement(XElement element)
    {
        return element.Name == DrawingNamespace + "r" ||
            element.Name == DrawingNamespace + "fld";
    }

    private static string ReadTextElementText(XElement element, int slideNumber)
    {
        if (slideNumber > 0 &&
            element.Name == DrawingNamespace + "fld" &&
            string.Equals((string?)element.Attribute("type"), "slidenum", StringComparison.OrdinalIgnoreCase))
        {
            return slideNumber.ToString(CultureInfo.InvariantCulture);
        }

        return NormalizeText((string?)element.Element(DrawingNamespace + "t") ?? string.Empty);
    }

    private static string NormalizeText(string text)
    {
        return text.Contains('\u00AD', StringComparison.Ordinal)
            ? text.Replace("\u00AD", string.Empty, StringComparison.Ordinal)
            : text;
    }

    private static double ReadFontSize(XElement? runProperties, XElement? defaultRunProperties)
    {
        return (runProperties?.Attribute("sz") ?? defaultRunProperties?.Attribute("sz")) is { } size
            ? int.Parse(size.Value, CultureInfo.InvariantCulture) / 100d
            : 18d;
    }

    private static double ReadCharacterSpacing(XElement? runProperties, XElement? defaultRunProperties)
    {
        return (runProperties?.Attribute("spc") ?? defaultRunProperties?.Attribute("spc")) is { } spacing
            ? int.Parse(spacing.Value, CultureInfo.InvariantCulture) / 100d
            : 0d;
    }

    private static double ReadBaselineOffset(XElement? runProperties, XElement? defaultRunProperties, double fontSize)
    {
        return (runProperties?.Attribute("baseline") ?? defaultRunProperties?.Attribute("baseline")) is { } baseline
            ? fontSize * int.Parse(baseline.Value, CultureInfo.InvariantCulture) / 100000d
            : 0d;
    }

    private static bool IsStrikeEnabled(XElement? runProperties, XElement? defaultRunProperties)
    {
        string? value = (string?)(runProperties?.Attribute("strike") ?? defaultRunProperties?.Attribute("strike"));
        return value is not null && !value.Equals("noStrike", StringComparison.OrdinalIgnoreCase);
    }

    private static IReadOnlyList<TextCapsFragment> ApplyTextCaps(string text, XElement? runProperties, XElement? defaultRunProperties)
    {
        string? value = (string?)(runProperties?.Attribute("cap") ?? defaultRunProperties?.Attribute("cap"));
        if (text.Length == 0)
        {
            return [];
        }

        if (value is "all")
        {
            return [new TextCapsFragment(text.ToUpperInvariant(), 1d)];
        }

        if (value is not "small")
        {
            return [new TextCapsFragment(text, 1d)];
        }

        var fragments = new List<TextCapsFragment>();
        var builder = new StringBuilder();
        bool? currentSmall = null;
        foreach (char character in text)
        {
            bool isSmall = char.IsLetter(character) && char.IsLower(character);
            if (currentSmall is not null && currentSmall != isSmall)
            {
                fragments.Add(new TextCapsFragment(builder.ToString(), currentSmall.Value ? 0.8d : 1d));
                builder.Clear();
            }

            currentSmall = isSmall;
            builder.Append(char.ToUpperInvariant(character));
        }

        if (builder.Length > 0 && currentSmall is not null)
        {
            fragments.Add(new TextCapsFragment(builder.ToString(), currentSmall.Value ? 0.8d : 1d));
        }

        return fragments;
    }

    private static TextInsets ReadTextInsets(XElement textBody)
    {
        XElement? bodyProperties = textBody.Element(DrawingNamespace + "bodyPr");
        return new TextInsets(
            ReadInset(bodyProperties, "lIns", 91440),
            ReadInset(bodyProperties, "rIns", 91440),
            ReadInset(bodyProperties, "tIns", 45720),
            ReadInset(bodyProperties, "bIns", 45720));
    }

    private static bool ClipsVerticalOverflow(XElement textBody)
    {
        string? overflow = (string?)textBody
            .Element(DrawingNamespace + "bodyPr")
            ?.Attribute("vertOverflow");
        return overflow?.Equals("clip", StringComparison.OrdinalIgnoreCase) == true;
    }

    private static double ReadInset(XElement? element, string attributeName, long defaultEmu)
    {
        long emu = element?.Attribute(attributeName) is { } attribute
            ? long.Parse(attribute.Value, CultureInfo.InvariantCulture)
            : defaultEmu;
        return OoxUnits.EmuToPoints(emu);
    }

    private static double ReadParagraphSpacing(
        XElement? paragraphProperties,
        XElement? defaultParagraphProperties,
        string elementName,
        double referenceFontSize)
    {
        XElement? spacing = paragraphProperties?.Element(DrawingNamespace + elementName) ??
            defaultParagraphProperties?.Element(DrawingNamespace + elementName);
        if (spacing?.Element(DrawingNamespace + "spcPts")?.Attribute("val") is { } points)
        {
            return int.Parse(points.Value, CultureInfo.InvariantCulture) / 100d;
        }

        if (spacing?.Element(DrawingNamespace + "spcPct")?.Attribute("val") is { } percent)
        {
            return referenceFontSize * Math.Max(0d, int.Parse(percent.Value, CultureInfo.InvariantCulture) / 100000d);
        }

        return 0d;
    }

    private static ParagraphIndent ReadParagraphIndent(XElement? paragraphProperties)
    {
        return new ParagraphIndent(
            ReadParagraphEmuAttribute(paragraphProperties, "marL"),
            ReadParagraphEmuAttribute(paragraphProperties, "indent"));
    }

    private static double ReadParagraphEmuAttribute(XElement? paragraphProperties, string attributeName)
    {
        return paragraphProperties?.Attribute(attributeName) is { } attribute
            ? OoxUnits.EmuToPoints(long.Parse(attribute.Value, CultureInfo.InvariantCulture))
            : 0d;
    }

    private static IReadOnlyList<double> ReadTabStops(XElement? paragraphProperties)
    {
        if (paragraphProperties?.Element(DrawingNamespace + "tabLst") is not { } tabList)
        {
            return Array.Empty<double>();
        }

        return tabList
            .Elements(DrawingNamespace + "tab")
            .Select(tab => tab.Attribute("pos") is { } position
                ? OoxUnits.EmuToPoints(long.Parse(position.Value, CultureInfo.InvariantCulture))
                : double.NaN)
            .Where(position => !double.IsNaN(position))
            .Order()
            .ToArray();
    }

    private static double ResolveNextTabX(double cursorX, double paragraphTextX, IReadOnlyList<double> tabStops, double fontSize)
    {
        double current = cursorX - paragraphTextX;
        foreach (double tabStop in tabStops)
        {
            if (tabStop > current + 0.01d)
            {
                return paragraphTextX + tabStop;
            }
        }

        return cursorX + fontSize * 2.2d;
    }

    private static LineSpacing ReadLineSpacing(XElement? paragraphProperties, XElement? defaultParagraphProperties)
    {
        XElement? spacing = paragraphProperties?.Element(DrawingNamespace + "lnSpc") ??
            defaultParagraphProperties?.Element(DrawingNamespace + "lnSpc");
        if (spacing?.Element(DrawingNamespace + "spcPts")?.Attribute("val") is { } points)
        {
            return LineSpacing.Absolute(Math.Max(0.1d, int.Parse(points.Value, CultureInfo.InvariantCulture) / 100d));
        }

        if (spacing?.Element(DrawingNamespace + "spcPct")?.Attribute("val") is { } percent)
        {
            return LineSpacing.Multiple(Math.Max(0.1d, int.Parse(percent.Value, CultureInfo.InvariantCulture) / 100000d), true);
        }

        return LineSpacing.Multiple(1d, false);
    }

    private static double ReadParagraphAdvance(LineSpacing lineSpacing, double fontSize)
    {
        return lineSpacing.IsExplicit
            ? lineSpacing.Resolve(fontSize)
            : fontSize * 1.2d;
    }

    private static double ReadLineAdvance(LineSpacing lineSpacing, double fontSize)
    {
        return lineSpacing.IsExplicit
            ? lineSpacing.Resolve(fontSize)
            : fontSize <= 12d ? fontSize : fontSize * 1.2d;
    }

    private static double ReadFirstLineBaselineOffset(XElement paragraph, XElement? defaultRunProperties, LineSpacing lineSpacing)
    {
        return LineBaselineOffset(ReadFirstParagraphFontSize(paragraph, defaultRunProperties), lineSpacing);
    }

    private static double ReadFirstParagraphFontSize(XElement paragraph, XElement? defaultRunProperties)
    {
        const double defaultFontSize = 18d;
        foreach (XElement child in paragraph.Elements())
        {
            if (child.Name == DrawingNamespace + "br")
            {
                return defaultFontSize;
            }

            if (!IsTextRunElement(child))
            {
                continue;
            }

            XElement? runProperties = child.Element(DrawingNamespace + "rPr");
            if (runProperties?.Attribute("sz") is { } size)
            {
                return int.Parse(size.Value, CultureInfo.InvariantCulture) / 100d;
            }

            if (defaultRunProperties?.Attribute("sz") is { } defaultSize)
            {
                return int.Parse(defaultSize.Value, CultureInfo.InvariantCulture) / 100d;
            }

            return defaultFontSize;
        }

        return defaultFontSize;
    }

    private static double LineBaselineOffset(double fontSize, LineSpacing lineSpacing)
    {
        if (lineSpacing.IsAbsolute)
        {
            return Math.Max(BaselineOffset(fontSize), lineSpacing.Value - fontSize * 0.374d);
        }

        return lineSpacing.IsExplicit
            ? lineSpacing.Resolve(fontSize) - fontSize * 0.234d
            : BaselineOffset(fontSize);
    }

    private static double BaselineOffset(double fontSize)
    {
        const double baselineOffsetFactor = 0.974d;
        return fontSize * baselineOffsetFactor;
    }

    private static string? ReadBulletText(XElement? paragraphProperties, ref int autoNumberValue)
    {
        if (paragraphProperties is null || paragraphProperties.Element(DrawingNamespace + "buNone") is not null)
        {
            return null;
        }

        if ((string?)paragraphProperties.Element(DrawingNamespace + "buChar")?.Attribute("char") is { } bullet)
        {
            return bullet;
        }

        XElement? autoNumber = paragraphProperties.Element(DrawingNamespace + "buAutoNum");
        if (autoNumber is null)
        {
            return null;
        }

        if (autoNumber.Attribute("startAt") is { } startAt &&
            int.TryParse(startAt.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int start) &&
            start > 0)
        {
            autoNumberValue = start;
        }

        string result = FormatAutoNumber(autoNumberValue, (string?)autoNumber.Attribute("type"));
        autoNumberValue++;
        return result;
    }

    private static string FormatAutoNumber(int value, string? type)
    {
        return type switch
        {
            "arabicParenBoth" => $"({value})",
            "arabicParenR" => $"{value})",
            "alphaLcPeriod" => $"{FormatAlphaNumber(value, upper: false)}.",
            "alphaUcPeriod" => $"{FormatAlphaNumber(value, upper: true)}.",
            "alphaLcParenR" => $"{FormatAlphaNumber(value, upper: false)})",
            "alphaUcParenR" => $"{FormatAlphaNumber(value, upper: true)})",
            "romanLcPeriod" => $"{FormatRomanNumber(value, upper: false)}.",
            "romanUcPeriod" => $"{FormatRomanNumber(value, upper: true)}.",
            "romanLcParenR" => $"{FormatRomanNumber(value, upper: false)})",
            "romanUcParenR" => $"{FormatRomanNumber(value, upper: true)})",
            _ => $"{value}."
        };
    }

    private static string FormatAlphaNumber(int value, bool upper)
    {
        if (value <= 0)
        {
            return string.Empty;
        }

        var builder = new StringBuilder();
        int current = value;
        while (current > 0)
        {
            current--;
            char letter = (char)((upper ? 'A' : 'a') + current % 26);
            builder.Insert(0, letter);
            current /= 26;
        }

        return builder.ToString();
    }

    private static string FormatRomanNumber(int value, bool upper)
    {
        if (value <= 0)
        {
            return string.Empty;
        }

        (int Value, string Numeral)[] numerals =
        [
            (1000, "M"),
            (900, "CM"),
            (500, "D"),
            (400, "CD"),
            (100, "C"),
            (90, "XC"),
            (50, "L"),
            (40, "XL"),
            (10, "X"),
            (9, "IX"),
            (5, "V"),
            (4, "IV"),
            (1, "I")
        ];

        var builder = new StringBuilder();
        int current = value;
        foreach ((int numeralValue, string numeral) in numerals)
        {
            while (current >= numeralValue)
            {
                builder.Append(numeral);
                current -= numeralValue;
            }
        }

        string result = builder.ToString();
        return upper ? result : result.ToLowerInvariant();
    }

    private static BulletStyle ReadBulletStyle(XElement? paragraphProperties, PptxTheme theme, double textFontSize, RgbColor textColor, string? textTypeface)
    {
        XElement? bulletFont = FindBulletProperty(paragraphProperties, "buFont");
        XElement? bulletColor = FindBulletProperty(paragraphProperties, "buClr");
        XElement? bulletSizePercent = FindBulletProperty(paragraphProperties, "buSzPct");
        XElement? bulletSizePoints = FindBulletProperty(paragraphProperties, "buSzPts");

        string? typeface = theme.ResolveTypeface((string?)bulletFont?.Attribute("typeface"));
        RgbColor color = bulletColor is not null &&
            TryReadSolidColor(bulletColor, theme, out RgbColor explicitColor)
                ? explicitColor
                : textColor;
        double fontSize = textFontSize;
        if (bulletSizePercent?.Attribute("val") is { } sizePercent)
        {
            fontSize = textFontSize * Math.Max(0.1d, int.Parse(sizePercent.Value, CultureInfo.InvariantCulture) / 100000d);
        }
        else if (bulletSizePoints?.Attribute("val") is { } sizePoints)
        {
            fontSize = Math.Max(0.1d, int.Parse(sizePoints.Value, CultureInfo.InvariantCulture) / 100d);
        }

        return new BulletStyle(fontSize, color, typeface ?? textTypeface);
    }

    private static XElement? FindBulletProperty(XElement? paragraphProperties, string localName)
    {
        if (paragraphProperties is null)
        {
            return null;
        }

        XName propertyName = DrawingNamespace + localName;
        XElement? marker = paragraphProperties
            .Elements()
            .FirstOrDefault(element => element.Name == DrawingNamespace + "buChar" ||
                element.Name == DrawingNamespace + "buAutoNum" ||
                element.Name == DrawingNamespace + "buBlip");
        IEnumerable<XElement> candidates = marker is null
            ? paragraphProperties.Elements()
            : paragraphProperties.Elements().TakeWhile(element => element != marker);
        return candidates.FirstOrDefault(element => element.Name == propertyName);
    }

    private static bool TryReadHighlightColor(XElement? runProperties, out RgbColor color)
    {
        XElement? highlight = runProperties?.Element(DrawingNamespace + "highlight");
        string? hex = (string?)highlight?.Element(DrawingNamespace + "srgbClr")?.Attribute("val");
        return RgbColor.TryParse(hex, out color);
    }

    private static TextVerticalAnchor ReadVerticalAnchor(XElement textBody)
    {
        string? anchor = (string?)textBody
            .Element(DrawingNamespace + "bodyPr")
            ?.Attribute("anchor");
        return anchor switch
        {
            "ctr" => TextVerticalAnchor.Middle,
            "b" => TextVerticalAnchor.Bottom,
            _ => TextVerticalAnchor.Top
        };
    }

    private static double EstimateTextHeight(XElement textBody, XElement? defaultParagraphProperties)
    {
        double height = 0d;
        foreach (XElement paragraph in textBody.Elements(DrawingNamespace + "p"))
        {
            XElement? paragraphProperties = paragraph.Element(DrawingNamespace + "pPr");
            XElement? defaultRunProperties = paragraphProperties?.Element(DrawingNamespace + "defRPr") ??
                defaultParagraphProperties?.Element(DrawingNamespace + "defRPr");
            LineSpacing lineSpacing = ReadLineSpacing(paragraphProperties, defaultParagraphProperties);
            double paragraphFontSize = ReadFirstParagraphFontSize(paragraph, defaultRunProperties);
            height += ReadParagraphSpacing(paragraphProperties, defaultParagraphProperties, "spcBef", paragraphFontSize);
            if (!ParagraphHasVisibleContent(paragraph))
            {
                if (ParagraphHasLayoutContent(paragraph))
                {
                    XElement? endRunProperties = paragraph.Element(DrawingNamespace + "endParaRPr");
                    height += ReadParagraphAdvance(lineSpacing, ReadFontSize(endRunProperties, defaultRunProperties));
                    double emptyFontSize = ReadFontSize(endRunProperties, defaultRunProperties);
                    height += ReadParagraphSpacing(paragraphProperties, defaultParagraphProperties, "spcAft", emptyFontSize);
                }

                continue;
            }

            double maxFontSize = 18d;
            bool hasLineContent = false;
            foreach (XElement child in paragraph.Elements())
            {
                if (child.Name == DrawingNamespace + "br")
                {
                    height += ReadLineAdvance(lineSpacing, maxFontSize);
                    maxFontSize = 18d;
                    hasLineContent = false;
                    continue;
                }

                if (!IsTextRunElement(child))
                {
                    continue;
                }

                XElement? runProperties = child.Element(DrawingNamespace + "rPr");
                double fontSize = ReadFontSize(runProperties, defaultRunProperties);
                maxFontSize = Math.Max(maxFontSize, fontSize);
                hasLineContent = true;
            }

            if (hasLineContent || maxFontSize > 0d)
            {
                height += ReadLineAdvance(lineSpacing, maxFontSize);
            }

            height += ReadParagraphSpacing(paragraphProperties, defaultParagraphProperties, "spcAft", paragraphFontSize);
        }

        return height;
    }

}
