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
    internal static string GetStaticStoryVariant(DocxPageSettings settings, DocxPageNumber pageNumber, int sectionFirstPageNumber = 1) =>
        IsStaticStoryVariantEnabled(settings, "first") && pageNumber.Physical == sectionFirstPageNumber ? "first" :
        IsStaticStoryVariantEnabled(settings, "even") &&
        pageNumber.Displayed % 2 == 0 ? "even" : "default";

    // Pre-layout ownership can exclude disabled variants without predicting a
    // section's page count. Placement applies first-page/parity selection below.
    internal static bool IsStaticStoryVariantEnabled(DocxPageSettings settings, string variant) =>
        string.Equals(variant, "default", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(variant, "first", StringComparison.OrdinalIgnoreCase) && settings.TitlePage == true ||
        string.Equals(variant, "even", StringComparison.OrdinalIgnoreCase) && settings.EvenAndOddHeaders == true;

    internal static bool IsStaticStoryVariantKnownDisabled(DocxPageSettings settings, string variant) =>
        (string.Equals(variant, "first", StringComparison.OrdinalIgnoreCase) ||
         string.Equals(variant, "even", StringComparison.OrdinalIgnoreCase)) &&
        !IsStaticStoryVariantEnabled(settings, variant);

    internal static (
        IReadOnlyList<DocxBodyElement> Header,
        IReadOnlyList<DocxBodyElement> Footer,
        IReadOnlyList<DocxFloatingDrawing> HeaderDrawings,
        IReadOnlyList<DocxFloatingDrawing> FooterDrawings)
        GetStaticFrameSources(DocxPageSettings settings, string variant) =>
        (DocxBlockTraversal.FindStaticStoryBodyElements(variant,
             settings.HeaderBodyElementsByType, settings.HeaderParagraphsByType) ?? [],
         DocxBlockTraversal.FindStaticStoryBodyElements(variant,
             settings.FooterBodyElementsByType, settings.FooterParagraphsByType) ?? [],
         settings.HeaderFloatingDrawingsByType.GetValueOrDefault(variant, []),
         settings.FooterFloatingDrawingsByType.GetValueOrDefault(variant, []));

    private static DocxSelectedStaticStory SelectStaticHeaderFooter(DocxPageSettings settings, bool header, string variant)
    {
        // Office A/B (w52 title-page plus w53 even-pages probes, Word-COM rendered):
        // a selected-but-undefined first/even story renders empty on its pages; Word
        // does not fall back to the default story there.
        IReadOnlyList<DocxBodyElement>? selected = DocxBlockTraversal.FindStaticStoryBodyElements(variant,
            header ? settings.HeaderBodyElementsByType : settings.FooterBodyElementsByType,
            header ? settings.HeaderParagraphsByType : settings.FooterParagraphsByType);
        return selected is null ? new DocxSelectedStaticStory([], null) : new DocxSelectedStaticStory(selected, variant);
    }

    private sealed record DocxSelectedStaticStory(IReadOnlyList<DocxBodyElement> BodyElements, string? VariantType);

    private static DocxSelectedStaticDrawings SelectStaticHeaderFooterDrawings(DocxPageSettings settings, bool header, string variant)
    {
        // Same no-fallback rule as text stories above (mechanism consistency; floating
        // drawings in undefined first/even stories are unprobed on their own).
        var drawingsByType = header ? settings.HeaderFloatingDrawingsByType : settings.FooterFloatingDrawingsByType;
        return drawingsByType.TryGetValue(variant, out IReadOnlyList<DocxFloatingDrawing>? selected)
            ? new DocxSelectedStaticDrawings(selected, variant)
            : new DocxSelectedStaticDrawings([], null);
    }

    private sealed record DocxSelectedStaticDrawings(IReadOnlyList<DocxFloatingDrawing> Drawings, string? VariantType);

    private static double ResolveHeaderDistance(DocxLayoutPage page)
    {
        return page.PageSettings.HeaderDistancePoints ?? DocxDefaults.HeaderFooterDistancePoints;
    }

    private static double ResolveFooterDistance(DocxLayoutPage page)
    {
        return page.PageSettings.FooterDistancePoints ?? DocxDefaults.HeaderFooterDistancePoints;
    }

    internal static IReadOnlyDictionary<int, DocxEffectiveSectionSettings> BuildEffectiveSectionSettings(DocxDocument document, out DocxEffectiveSectionSettings finalSectionSettings,
        CancellationToken cancellationToken = default)
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
        IReadOnlyDictionary<string, IReadOnlyList<DocxFloatingDrawing>> inheritedHeaderDrawingsByType =
            new Dictionary<string, IReadOnlyList<DocxFloatingDrawing>>(StringComparer.OrdinalIgnoreCase);
        IReadOnlyDictionary<string, IReadOnlyList<DocxFloatingDrawing>> inheritedFooterDrawingsByType =
            new Dictionary<string, IReadOnlyList<DocxFloatingDrawing>>(StringComparer.OrdinalIgnoreCase);

        for (int elementIndex = 0; elementIndex < document.BodyElements.Count; elementIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (document.BodyElements[elementIndex] is not DocxSectionBreakElement sectionBreak)
            {
                continue;
            }

            DocxPageSettings effectiveSettings = ResolveEffectiveSectionSettings(
                sectionBreak.PageSettings,
                inheritedHeadersByType,
                inheritedFootersByType,
                inheritedHeaderBodyElementsByType,
                inheritedFooterBodyElementsByType,
                inheritedHeaderDrawingsByType,
                inheritedFooterDrawingsByType);
            sectionSettingsByElementIndex[elementIndex] = new DocxEffectiveSectionSettings(
                effectiveSettings,
                CreateSectionLayoutProperties(sectionBreak));
            inheritedHeadersByType = effectiveSettings.HeaderParagraphsByType;
            inheritedFootersByType = effectiveSettings.FooterParagraphsByType;
            inheritedHeaderBodyElementsByType = effectiveSettings.HeaderBodyElementsByType;
            inheritedFooterBodyElementsByType = effectiveSettings.FooterBodyElementsByType;
            inheritedHeaderDrawingsByType = effectiveSettings.HeaderFloatingDrawingsByType;
            inheritedFooterDrawingsByType = effectiveSettings.FooterFloatingDrawingsByType;
        }

        finalSectionSettings = new DocxEffectiveSectionSettings(
            ResolveEffectiveSectionSettings(
                BuildFinalSectionSettings(),
                inheritedHeadersByType,
                inheritedFootersByType,
                inheritedHeaderBodyElementsByType,
                inheritedFooterBodyElementsByType,
                inheritedHeaderDrawingsByType,
                inheritedFooterDrawingsByType),
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

            // Paragraph-only inputs can supply a default. A body map already
            // owns its variants, including explicit empty stories and tables.
            if (headerBodyElementsByType.Count == 0 && headersByType.Count == 0 && document.HeaderParagraphs.Count > 0)
            {
                headersByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
                {
                    ["default"] = document.HeaderParagraphs
                };
                headerBodyElementsByType = ToStaticBodyElementsByType(headersByType);
            }

            if (footerBodyElementsByType.Count == 0 && footersByType.Count == 0 && document.FooterParagraphs.Count > 0)
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
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> inheritedFooterBodyElementsByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxFloatingDrawing>> inheritedHeaderDrawingsByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxFloatingDrawing>> inheritedFooterDrawingsByType)
    {
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> headersByType = MergeInheritedStaticItems(inheritedHeadersByType, settings.HeaderParagraphsByType);
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> footersByType = MergeInheritedStaticItems(inheritedFootersByType, settings.FooterParagraphsByType);
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> headerBodyElementsByType = MergeInheritedStaticItems(inheritedHeaderBodyElementsByType, settings.HeaderBodyElementsByType);
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> footerBodyElementsByType = MergeInheritedStaticItems(inheritedFooterBodyElementsByType, settings.FooterBodyElementsByType);

        return settings with
        {
            HeaderParagraphsByType = headersByType,
            FooterParagraphsByType = footersByType,
            HeaderBodyElementsByType = headerBodyElementsByType.Count == 0 ? ToStaticBodyElementsByType(headersByType) : headerBodyElementsByType,
            FooterBodyElementsByType = footerBodyElementsByType.Count == 0 ? ToStaticBodyElementsByType(footersByType) : footerBodyElementsByType,
            HeaderFloatingDrawingsByType = MergeInheritedStaticItems(inheritedHeaderDrawingsByType, settings.HeaderFloatingDrawingsByType),
            FooterFloatingDrawingsByType = MergeInheritedStaticItems(inheritedFooterDrawingsByType, settings.FooterFloatingDrawingsByType)
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

    // The reader supplies a key even for an explicitly empty story. Replacing
    // that list clears the previous owner; an absent key inherits it unchanged.
    private static IReadOnlyDictionary<string, IReadOnlyList<T>> MergeInheritedStaticItems<T>(
        IReadOnlyDictionary<string, IReadOnlyList<T>> inheritedByType,
        IReadOnlyDictionary<string, IReadOnlyList<T>> localByType)
    {
        if (inheritedByType.Count == 0)
        {
            return localByType;
        }

        if (localByType.Count == 0)
        {
            return inheritedByType;
        }

        var merged = new Dictionary<string, IReadOnlyList<T>>(inheritedByType, StringComparer.OrdinalIgnoreCase);
        foreach ((string type, IReadOnlyList<T> items) in localByType)
        {
            merged[type] = items;
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
        double width = ReadOptionalTwipsValue(effectiveSettings.WidthValue) ?? document.PageWidthPoints;
        double height = ReadOptionalTwipsValue(effectiveSettings.HeightValue) ?? document.PageHeightPoints;
        (width, height) = DocxPageSettings.NormalizeDimensions(width, height, effectiveSettings.OrientationValue);

        double marginLeft = ReadOptionalTwipsValue(effectiveSettings.MarginLeftValue) ?? document.MarginLeftPoints;
        double marginRight = ReadOptionalTwipsValue(effectiveSettings.MarginRightValue) ?? document.MarginRightPoints;
        double gutter = Math.Max(0d, ReadOptionalTwipsValue(effectiveSettings.GutterDistanceValue) ?? effectiveSettings.GutterDistancePoints ?? 0d);
        // Keep the authored right edge for the fixed review lane. Word mirrors
        // the body and inside gutter, while balloons remain at this odd-page edge.
        double oddAuthoredMarginLeft = marginLeft;
        double oddAuthoredMarginRight = marginRight;
        bool scaledMarkupReserve = reserveMarkupMargin && retuneReserve && Math.Abs(printScale - 1d) >= 0.000000001d;
        if (scaledMarkupReserve && IsEvenMirroredPage(document, pageNumber))
        {
            // Swap before charging the gutter and retuning the reserve so
            // identical odd/even content has the same scaled wrap width.
            (marginLeft, marginRight) = (marginRight, marginLeft);
        }
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
        if (!scaledMarkupReserve && IsEvenMirroredPage(document, pageNumber))
        {
            // Preserve the existing unscaled geometry profiles: body margins
            // fully mirror without reserve; the fixed reserve stays on the right.
            marginLeft = oddAuthoredMarginRight;
            if (!reserveMarkupMargin)
            {
                marginRight = oddAuthoredMarginLeft;
            }
        }
        double marginTop = ReadOptionalTwipsValue(effectiveSettings.MarginTopValue) ?? document.MarginTopPoints;
        double marginBottom = ReadOptionalTwipsValue(effectiveSettings.MarginBottomValue) ?? document.MarginBottomPoints;

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
                section.SectionProperties))
        {
            MarkupLaneDesignBodyEnd = scaledMarkupReserve ? width - oddAuthoredMarginRight : null
        };

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

    private static bool ShouldInsertParityBlankPage(DocxSectionBreakType? boundary, int nextPageNumber) => boundary switch
    {
        DocxSectionBreakType.OddPage => nextPageNumber % 2 == 0,
        DocxSectionBreakType.EvenPage => nextPageNumber % 2 != 0,
        _ => false
    };
}
