using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    private static DocxSelectedStaticStory SelectStaticHeaderFooter(
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> bodyElementsByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> paragraphsByType,
        DocxPageSettings settings,
        int pageNumber)
    {
        // Office A/B (w52 title-page plus w53 even-pages probes, Word-COM rendered):
        // a selected-but-undefined first/even story renders empty on its pages; Word
        // does not fall back to the default story there.
        if (settings.TitlePage == true && pageNumber == 1)
        {
            return DocxBlockTraversal.TryGetStaticStoryBodyElements("first", bodyElementsByType, paragraphsByType, out IReadOnlyList<DocxBodyElement>? first)
                ? new DocxSelectedStaticStory(first, "first")
                : new DocxSelectedStaticStory([], null);
        }

        if (settings.EvenAndOddHeaders == true && pageNumber % 2 == 0)
        {
            return DocxBlockTraversal.TryGetStaticStoryBodyElements("even", bodyElementsByType, paragraphsByType, out IReadOnlyList<DocxBodyElement>? even)
                ? new DocxSelectedStaticStory(even, "even")
                : new DocxSelectedStaticStory([], null);
        }

        return DocxBlockTraversal.TryGetStaticStoryBodyElements("default", bodyElementsByType, paragraphsByType, out IReadOnlyList<DocxBodyElement>? defaults)
            ? new DocxSelectedStaticStory(defaults, "default")
            : new DocxSelectedStaticStory([], null);
    }

    private sealed record DocxSelectedStaticStory(IReadOnlyList<DocxBodyElement> BodyElements, string? VariantType);

    private static DocxSelectedStaticDrawings SelectStaticHeaderFooterDrawings(
        IReadOnlyDictionary<string, IReadOnlyList<DocxFloatingDrawing>> drawingsByType,
        DocxPageSettings settings,
        int pageNumber)
    {
        // Same no-fallback rule as text stories above (mechanism consistency; floating
        // drawings in undefined first/even stories are unprobed on their own).
        if (settings.TitlePage == true && pageNumber == 1)
        {
            return drawingsByType.TryGetValue("first", out IReadOnlyList<DocxFloatingDrawing>? first)
                ? new DocxSelectedStaticDrawings(first, "first")
                : new DocxSelectedStaticDrawings([], null);
        }

        if (settings.EvenAndOddHeaders == true && pageNumber % 2 == 0)
        {
            return drawingsByType.TryGetValue("even", out IReadOnlyList<DocxFloatingDrawing>? even)
                ? new DocxSelectedStaticDrawings(even, "even")
                : new DocxSelectedStaticDrawings([], null);
        }

        return drawingsByType.TryGetValue("default", out IReadOnlyList<DocxFloatingDrawing>? defaults)
            ? new DocxSelectedStaticDrawings(defaults, "default")
            : new DocxSelectedStaticDrawings([], null);
    }

    private sealed record DocxSelectedStaticDrawings(IReadOnlyList<DocxFloatingDrawing> Drawings, string? VariantType);

    private static double ResolveHeaderDistance(DocxLayoutPage page)
    {
        return page.PageSettings.HeaderDistancePoints ?? Math.Max(18d, page.MarginTop / 2d);
    }

    private static double ResolveFooterDistance(DocxLayoutPage page)
    {
        return page.PageSettings.FooterDistancePoints ?? Math.Max(18d, page.MarginBottom / 2d);
    }

    private static IReadOnlyDictionary<int, DocxEffectiveSectionSettings> BuildEffectiveSectionSettings(DocxDocument document, out DocxEffectiveSectionSettings finalSectionSettings)
    {
        var sectionSettingsByElementIndex = new Dictionary<int, DocxEffectiveSectionSettings>();
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> inheritedHeadersByType =
            new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase);
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> inheritedFootersByType =
            new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase);
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> inheritedHeaderBodyElementsByType =
            new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase);
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> inheritedFooterBodyElementsByType =
            new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase);

        for (int elementIndex = 0; elementIndex < document.BodyElements.Count; elementIndex++)
        {
            if (document.BodyElements[elementIndex] is not DocxSectionBreakElement sectionBreak)
            {
                continue;
            }

            DocxPageSettings effectiveSettings = ResolveEffectiveSectionSettings(
                sectionBreak.PageSettings,
                inheritedHeadersByType,
                inheritedFootersByType,
                inheritedHeaderBodyElementsByType,
                inheritedFooterBodyElementsByType);
            sectionSettingsByElementIndex[elementIndex] = new DocxEffectiveSectionSettings(
                effectiveSettings,
                CreateSectionLayoutProperties(sectionBreak));
            inheritedHeadersByType = effectiveSettings.HeaderParagraphsByType;
            inheritedFootersByType = effectiveSettings.FooterParagraphsByType;
            inheritedHeaderBodyElementsByType = effectiveSettings.HeaderBodyElementsByType;
            inheritedFooterBodyElementsByType = effectiveSettings.FooterBodyElementsByType;
        }

        finalSectionSettings = new DocxEffectiveSectionSettings(
            ResolveEffectiveSectionSettings(
                BuildFinalSectionSettings(),
                inheritedHeadersByType,
                inheritedFootersByType,
                inheritedHeaderBodyElementsByType,
                inheritedFooterBodyElementsByType),
            document.FinalSectionBreak is null
                ? new DocxSectionLayoutProperties(null, null, null, null, null, null, [])
                : CreateSectionLayoutProperties(document.FinalSectionBreak));
        return sectionSettingsByElementIndex;

        DocxPageSettings BuildFinalSectionSettings()
        {
            DocxPageSettings settings = document.PageSettings;
            IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> headersByType = settings.HeaderParagraphsByType.Count == 0 && document.HeaderParagraphsByType.Count > 0
                ? document.HeaderParagraphsByType
                : settings.HeaderParagraphsByType;
            IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> footersByType = settings.FooterParagraphsByType.Count == 0 && document.FooterParagraphsByType.Count > 0
                ? document.FooterParagraphsByType
                : settings.FooterParagraphsByType;
            IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> headerBodyElementsByType = settings.HeaderBodyElementsByType.Count == 0 && document.HeaderBodyElementsByType.Count > 0
                ? document.HeaderBodyElementsByType
                : settings.HeaderBodyElementsByType;
            IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> footerBodyElementsByType = settings.FooterBodyElementsByType.Count == 0 && document.FooterBodyElementsByType.Count > 0
                ? document.FooterBodyElementsByType
                : settings.FooterBodyElementsByType;

            if (headersByType.Count == 0 && document.HeaderParagraphs.Count > 0)
            {
                headersByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
                {
                    ["default"] = document.HeaderParagraphs
                };
                headerBodyElementsByType = ToStaticBodyElementsByType(headersByType);
            }

            if (footersByType.Count == 0 && document.FooterParagraphs.Count > 0)
            {
                footersByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
                {
                    ["default"] = document.FooterParagraphs
                };
                footerBodyElementsByType = ToStaticBodyElementsByType(footersByType);
            }

            return settings with
            {
                HeaderParagraphsByType = headersByType,
                FooterParagraphsByType = footersByType,
                HeaderBodyElementsByType = headerBodyElementsByType.Count == 0 ? ToStaticBodyElementsByType(headersByType) : headerBodyElementsByType,
                FooterBodyElementsByType = footerBodyElementsByType.Count == 0 ? ToStaticBodyElementsByType(footersByType) : footerBodyElementsByType
            };
        }
    }

    private static DocxSectionLayoutProperties CreateSectionLayoutProperties(DocxSectionBreakElement sectionBreak)
    {
        return new DocxSectionLayoutProperties(
            sectionBreak.TypeValue?.ToValueString(),
            sectionBreak.ColumnCountValue,
            sectionBreak.ColumnEqualWidthValue,
            sectionBreak.ColumnSpaceValue,
            ReadOptionalInt32Value(sectionBreak.ColumnCountValue),
            ReadOptionalTwipsValue(sectionBreak.ColumnSpaceValue),
            sectionBreak.ColumnDefinitions
                .Select(column => new DocxSectionColumnLayoutProperties(
                    column.WidthValue,
                    column.SpaceValue,
                    ReadOptionalTwipsValue(column.WidthValue),
                    ReadOptionalTwipsValue(column.SpaceValue)))
                .ToArray());
    }

    private static DocxPageSettings ResolveEffectiveSectionSettings(
        DocxPageSettings settings,
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> inheritedHeadersByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> inheritedFootersByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> inheritedHeaderBodyElementsByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> inheritedFooterBodyElementsByType)
    {
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> headersByType = MergeInheritedStaticParagraphs(inheritedHeadersByType, settings.HeaderParagraphsByType);
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> footersByType = MergeInheritedStaticParagraphs(inheritedFootersByType, settings.FooterParagraphsByType);
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> headerBodyElementsByType = MergeInheritedStaticBodyElements(inheritedHeaderBodyElementsByType, settings.HeaderBodyElementsByType);
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> footerBodyElementsByType = MergeInheritedStaticBodyElements(inheritedFooterBodyElementsByType, settings.FooterBodyElementsByType);

        return settings with
        {
            HeaderParagraphsByType = headersByType,
            FooterParagraphsByType = footersByType,
            HeaderBodyElementsByType = headerBodyElementsByType.Count == 0 ? ToStaticBodyElementsByType(headersByType) : headerBodyElementsByType,
            FooterBodyElementsByType = footerBodyElementsByType.Count == 0 ? ToStaticBodyElementsByType(footersByType) : footerBodyElementsByType
        };
    }

    private static IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> ToStaticBodyElementsByType(
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> paragraphsByType)
    {
        return paragraphsByType.ToDictionary(
            pair => pair.Key,
            pair => (IReadOnlyList<DocxBodyElement>)pair.Value.Select(DocxBodyElementFactory.CreateParagraph).Cast<DocxBodyElement>().ToArray(),
            StringComparer.OrdinalIgnoreCase);
    }

    private static IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> MergeInheritedStaticBodyElements(
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> inheritedBodyElementsByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> localBodyElementsByType)
    {
        if (inheritedBodyElementsByType.Count == 0)
        {
            return localBodyElementsByType;
        }

        if (localBodyElementsByType.Count == 0)
        {
            return inheritedBodyElementsByType;
        }

        var merged = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(inheritedBodyElementsByType, StringComparer.OrdinalIgnoreCase);
        foreach ((string type, IReadOnlyList<DocxBodyElement> bodyElements) in localBodyElementsByType)
        {
            merged[type] = bodyElements;
        }

        return merged;
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

    private static DocxEffectiveSectionSettings? FindSectionSettingsAtOrAfter(
        IReadOnlyList<DocxBodyElement> elements,
        int startIndex,
        IReadOnlyDictionary<int, DocxEffectiveSectionSettings> sectionSettingsByElementIndex)
    {
        for (int i = Math.Max(0, startIndex); i < elements.Count; i++)
        {
            if (elements[i] is DocxSectionBreakElement && sectionSettingsByElementIndex.TryGetValue(i, out DocxEffectiveSectionSettings? settings))
            {
                return settings;
            }
        }

        return null;
    }

    private static DocxPageGeometry ResolveSectionGeometry(
        DocxDocument document,
        DocxEffectiveSectionSettings section,
        bool reserveMarkupMargin,
        bool retuneReserve,
        double printScale,
        int pageNumber)
    {
        DocxPageSettings effectiveSettings = section.PageSettings;
        double width = ReadTwipsValue(effectiveSettings.WidthValue, document.PageWidthPoints);
        double height = ReadTwipsValue(effectiveSettings.HeightValue, document.PageHeightPoints);
        (width, height) = NormalizePageSize(width, height);
        if (effectiveSettings.OrientationValue?.Equals("landscape", StringComparison.OrdinalIgnoreCase) == true && height > width)
        {
            (width, height) = (height, width);
        }

        double marginLeft = ReadTwipsValue(effectiveSettings.MarginLeftValue, document.MarginLeftPoints);
        double marginRight = ReadTwipsValue(effectiveSettings.MarginRightValue, document.MarginRightPoints);
        double gutter = Math.Max(0d, ReadTwipsValue(effectiveSettings.GutterDistanceValue, effectiveSettings.GutterDistancePoints ?? 0d));
        // Odd-authored margins, captured before gutter and reserve: even mirrored pages
        // mirror the body (left) side onto the authored right while the right side keeps
        // odd-page geometry below.
        double oddAuthoredMarginLeft = marginLeft;
        double oddAuthoredMarginRight = marginRight;
        if (gutter > 0d)
        {
            if (ShouldApplyGutterToRightMargin(document, pageNumber))
            {
                marginRight += gutter;
            }
            else
            {
                marginLeft += gutter;
            }
        }

        double authoredMarginLeft = marginLeft;
        double authoredMarginRight = marginRight;
        if (reserveMarkupMargin)
        {
            // Office (mirrored-margin reference, Word-COM rendered): even mirrored pages
            // keep the balloon lane on the right like odd pages (balloons at 437.33 on both
            // pages), so the review reserve never mirrors left. The 2026-06 mirroring
            // assumption is retired; page gutters still mirror via ShouldApplyGutterToRightMargin.
            marginRight = ResolveReservedMarkupRightMargin(width, marginLeft, marginRight, printScale, retuneReserve);
        }

        double markupMarginReservePoints = Math.Max(0d, Math.Max(marginLeft - authoredMarginLeft, marginRight - authoredMarginRight));
        if (IsEvenMirroredPage(document, pageNumber))
        {
            // Office (mirrored-margin reference): even pages mirror the body (left) side onto
            // the authored right margin, while the right side keeps odd-page geometry (full
            // mirror without reserve, odd reserved geometry with reserve) so the review reserve
            // and balloon lane never mirror left.
            marginLeft = oddAuthoredMarginRight;
            if (!reserveMarkupMargin)
            {
                marginRight = oddAuthoredMarginLeft;
            }
        }
        double marginTop = ReadTwipsValue(effectiveSettings.MarginTopValue, document.MarginTopPoints);
        double marginBottom = ReadTwipsValue(effectiveSettings.MarginBottomValue, document.MarginBottomPoints);

        return new DocxPageGeometry(
            width,
            height,
            marginLeft,
            marginRight,
            markupMarginReservePoints,
            marginTop,
            marginBottom,
            effectiveSettings,
            section.SectionProperties,
            CreateColumnFrames(
                width,
                section.SectionProperties));

        IReadOnlyList<DocxLayoutColumnFrame> CreateColumnFrames(double pageWidth, DocxSectionLayoutProperties section)
        {
            double bodyWidth = Math.Max(1d, pageWidth - marginLeft - marginRight);
            int columnCount = Math.Max(1, section.ColumnCount ?? 1);
            if (columnCount == 1)
            {
                return [new DocxLayoutColumnFrame(0, marginLeft, bodyWidth, null)];
            }

            if (string.Equals(section.ColumnEqualWidthValue, "0", StringComparison.OrdinalIgnoreCase))
            {
                if (section.ColumnDefinitions.Count == 0)
                {
                    return [];
                }

                double x = marginLeft;
                var frames = new List<DocxLayoutColumnFrame>();
                for (int index = 0; index < section.ColumnDefinitions.Count; index++)
                {
                    DocxSectionColumnLayoutProperties column = section.ColumnDefinitions[index];
                    double customColumnWidth = Math.Max(1d, column.WidthPoints ?? 0d);
                    double? customGutter = index + 1 < section.ColumnDefinitions.Count
                        ? Math.Max(0d, column.SpacePoints ?? 0d)
                        : null;
                    frames.Add(new DocxLayoutColumnFrame(index, x, customColumnWidth, customGutter));
                    x += customColumnWidth + (customGutter ?? 0d);
                }

                return frames;
            }

            double columnGutter = Math.Max(0d, section.ColumnSpacePoints ?? 0d);
            double columnWidth = Math.Max(1d, (bodyWidth - columnGutter * (columnCount - 1)) / columnCount);
            return Enumerable.Range(0, columnCount)
                .Select(index => new DocxLayoutColumnFrame(
                    index,
                    marginLeft + index * (columnWidth + columnGutter),
                    columnWidth,
                    index + 1 < columnCount ? columnGutter : null))
                .ToArray();
        }
    }

    private static double ResolveReservedMarkupRightMargin(double pageWidth, double marginLeft, double marginRight, double printScale, bool retuneReserve)
    {
        double bodyWidth = Math.Max(1d, pageWidth - marginLeft - marginRight);
        if (bodyWidth <= MinimumMarkupBodyWidthPoints)
        {
            return marginRight;
        }

        double maxRightMargin = Math.Max(marginRight, pageWidth - marginLeft - MinimumMarkupBodyWidthPoints);
        double preferredMargin = PreferredMarkupMarginPoints;
        if (retuneReserve && Math.Abs(printScale - 1d) >= 0.000000001d)
        {
            // Uniform-scale break equivalence: scaled metrics break at the layout body, which
            // reads as layoutBody/scale in design space. Size the layout body to the authored
            // body times scale so breaks land where Word full-design layout breaks them.
            preferredMargin = Math.Max(marginRight, pageWidth - marginLeft - bodyWidth * printScale);
        }

        return Math.Max(marginRight, Math.Min(preferredMargin, maxRightMargin));
    }


    private static bool ShouldApplyGutterToRightMargin(DocxDocument document, int pageNumber)
    {
        return IsEvenMirroredPage(document, pageNumber);
    }

    private static bool IsEvenMirroredPage(DocxDocument document, int pageNumber)
    {
        return document.Settings.MirrorMargins == true && pageNumber % 2 == 0;
    }

    private static DocxLayoutColumnFrame ResolveActiveColumnFrame(DocxPageGeometry page, int activeColumnIndex)
    {
        if (page.ColumnFrames.Count == 0)
        {
            return new DocxLayoutColumnFrame(0, page.MarginLeft, page.BodyWidth, null);
        }

        int index = Math.Clamp(activeColumnIndex, 0, page.ColumnFrames.Count - 1);
        return page.ColumnFrames[index];
    }

    private static double ReadTwipsValue(string? value, double fallback)
    {
        return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long twips)
            ? OoxUnits.TwipsToPoints(twips)
            : fallback;
    }

    private static double? ReadOptionalTwipsValue(string? value)
    {
        return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long twips)
            ? OoxUnits.TwipsToPoints(twips)
            : null;
    }

    private static int? ReadOptionalInt32Value(string? value)
    {
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int result)
            ? result
            : null;
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
            sectionBreak.TypeValue is DocxSectionBreakType.NextPage or DocxSectionBreakType.OddPage or DocxSectionBreakType.EvenPage;
    }

    private static bool IsContinuousSectionBreak(DocxSectionBreakElement sectionBreak)
    {
        return sectionBreak.TypeValue == DocxSectionBreakType.Continuous;
    }

    private static bool ShouldInsertParityBlankPage(DocxSectionBreakElement sectionBreak, int nextPageNumber)
    {
        if (sectionBreak.TypeValue == DocxSectionBreakType.OddPage)
        {
            return nextPageNumber % 2 == 0;
        }

        if (sectionBreak.TypeValue == DocxSectionBreakType.EvenPage)
        {
            return nextPageNumber % 2 != 0;
        }

        return false;
    }
}
