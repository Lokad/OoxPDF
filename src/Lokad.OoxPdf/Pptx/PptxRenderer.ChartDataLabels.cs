using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static ChartLayoutBox ResolveDataLabelBox(ChartPlotBox plotBox, ChartDataLabelOptions options, double x, double y, double width, double height)
    {
        ChartLayoutBox defaultBox = new(x, y, width, height);
        if (!options.Layout.HasLayout)
        {
            return defaultBox;
        }

        ChartFrameBox frame = new(plotBox.X, plotBox.Y, plotBox.Width, plotBox.Height);
        return TryBuildManualLayoutBox(options.Layout, frame, defaultBox, out ChartLayoutBox manualBox, clampToFrame: false, missingPositionModesAreFactor: true)
            ? manualBox
            : defaultBox;
    }

    private static double ResolveHorizontalBarDataLabelX(PptxSceneChartDataLabelPosition position, double barBase, double barEnd, double labelWidth)
    {
        position = ResolveChartDataLabelPosition(position);
        bool extendsRight = barEnd >= barBase;
        return position switch
        {
            PptxSceneChartDataLabelPosition.Center or PptxSceneChartDataLabelPosition.BestFit => (barBase + barEnd - labelWidth) / 2d,
            PptxSceneChartDataLabelPosition.InsideBase => extendsRight ? barBase + PptxChartMetricRules.BarDataLabelHorizontalGap : barBase - labelWidth - PptxChartMetricRules.BarDataLabelHorizontalGap,
            PptxSceneChartDataLabelPosition.InsideEnd or PptxSceneChartDataLabelPosition.Left or PptxSceneChartDataLabelPosition.Right => extendsRight ? barEnd - labelWidth - PptxChartMetricRules.BarDataLabelHorizontalGap : barEnd + PptxChartMetricRules.BarDataLabelHorizontalGap,
            PptxSceneChartDataLabelPosition.OutsideEnd or _ => extendsRight ? barEnd + PptxChartMetricRules.BarDataLabelHorizontalGap : barEnd - labelWidth - PptxChartMetricRules.BarDataLabelHorizontalGap
        };
    }

    private static double ResolveVerticalBarDataLabelY(PptxSceneChartDataLabelPosition position, double barBase, double barEnd, double labelHeight)
    {
        position = ResolveChartDataLabelPosition(position);
        bool extendsUp = barEnd >= barBase;
        return position switch
        {
            PptxSceneChartDataLabelPosition.Center or PptxSceneChartDataLabelPosition.BestFit => (barBase + barEnd - labelHeight) / 2d,
            PptxSceneChartDataLabelPosition.InsideBase => extendsUp ? barBase + PptxChartMetricRules.BarDataLabelVerticalGap : barBase - labelHeight - PptxChartMetricRules.BarDataLabelVerticalGap,
            PptxSceneChartDataLabelPosition.InsideEnd or PptxSceneChartDataLabelPosition.Top or PptxSceneChartDataLabelPosition.Bottom => extendsUp ? barEnd - labelHeight - PptxChartMetricRules.BarDataLabelVerticalGap : barEnd + PptxChartMetricRules.BarDataLabelVerticalGap,
            PptxSceneChartDataLabelPosition.OutsideEnd or _ => extendsUp ? barEnd + PptxChartMetricRules.BarDataLabelVerticalGap : barEnd - labelHeight - PptxChartMetricRules.BarDataLabelVerticalGap
        };
    }

    private static PptxSceneChartDataLabelPosition ResolveStackedBarDataLabelPosition(PptxSceneChartDataLabelPosition position, bool stacked)
    {
        return stacked && position == PptxSceneChartDataLabelPosition.Unknown
            ? PptxSceneChartDataLabelPosition.Center
            : position;
    }

    private static (double Start, double End) ResolveStackedBarSegmentValues(double[] positiveValues, double[] negativeValues, int category, double value)
    {
        if (value >= 0d)
        {
            double start = positiveValues[category];
            positiveValues[category] += value;
            return (start, positiveValues[category]);
        }

        double negativeStart = negativeValues[category];
        negativeValues[category] += value;
        return (negativeStart, negativeValues[category]);
    }

    private static (double X, double Y, TextAlignment Alignment) ResolveCategoryPointLabelPosition(
        PptxSceneChartDataLabelPosition position, double pointX, double pointY, double labelWidth,
        double ascent, double height, double markerRadius, double verticalExpansion)
    {
        double centeredBaseline = pointY + height / 2d - ascent;
        // The occupied frame keeps a physical clearance from the authored
        // marker extent. Font metrics determine baseline and height, not this gap.
        double horizontalGap = markerRadius + PptxChartMetricRules.CategoryPointLabelMarkerClearance;
        double verticalGap = horizontalGap + PptxChartMetricRules.DataLabelVerticalPadding;
        return position switch
        {
            PptxSceneChartDataLabelPosition.Top or PptxSceneChartDataLabelPosition.OutsideEnd =>
                (pointX - labelWidth / 2d, pointY + verticalGap + height - ascent + verticalExpansion, TextAlignment.Center),
            PptxSceneChartDataLabelPosition.Bottom =>
                (pointX - labelWidth / 2d, pointY - verticalGap - ascent - verticalExpansion, TextAlignment.Center),
            PptxSceneChartDataLabelPosition.Left =>
                (pointX - horizontalGap - labelWidth, centeredBaseline, TextAlignment.Right),
            PptxSceneChartDataLabelPosition.Right =>
                (pointX + horizontalGap, centeredBaseline, TextAlignment.Left),
            _ => (pointX - labelWidth / 2d, centeredBaseline, TextAlignment.Center)
        };
    }

    // Authored label sizes retain the existing approximate placement fallback.
    private static (double X, double Y, TextAlignment Alignment) ResolveLineDataLabelPosition(PptxSceneChartDataLabelPosition position, double pointX, double pointY, double labelWidth, double labelHeight)
    {
        position = ResolveChartDataLabelPosition(position);
        return position switch
        {
            PptxSceneChartDataLabelPosition.Bottom => (pointX - labelWidth / 2d, pointY - labelHeight * PptxChartMetricRules.LineDataLabelBelowOffsetFactor, TextAlignment.Center),
            PptxSceneChartDataLabelPosition.Left => (pointX - labelWidth - PptxChartMetricRules.LineDataLabelSideGap, pointY - labelHeight / 2d, TextAlignment.Right),
            PptxSceneChartDataLabelPosition.Right => (pointX + PptxChartMetricRules.LineDataLabelSideGap, pointY - labelHeight / 2d, TextAlignment.Left),
            PptxSceneChartDataLabelPosition.Center or PptxSceneChartDataLabelPosition.BestFit => (pointX - labelWidth / 2d, pointY - labelHeight / 2d, TextAlignment.Center),
            PptxSceneChartDataLabelPosition.Top or PptxSceneChartDataLabelPosition.OutsideEnd or _ => (pointX - labelWidth / 2d, pointY + labelHeight * PptxChartMetricRules.LineDataLabelAboveOffsetFactor, TextAlignment.Center)
        };
    }

    private static PptxSceneChartDataLabelPosition ResolveChartDataLabelPosition(PptxSceneChartDataLabelPosition position)
    {
        return position == PptxSceneChartDataLabelPosition.Unknown
            ? PptxSceneChartDataLabelPosition.OutsideEnd
            : position;
    }

    private static ChartTextStyle ResolveChartDataLabelTextStyle(PptxTheme theme, PptxColorMap colorMap, ChartDataLabelOptions options)
    {
        RgbColor fallbackColor = theme.TryResolveColor("tx1", colorMap, out RgbColor themeText)
            ? themeText
            : new RgbColor(0, 0, 0);
        PptxThemeTypefaceResolution typeface = ResolveChartThemeTypeface(theme);
        ChartTextStyle style = new(
            typeface.Typeface,
            PptxChartMetricRules.DataLabelFallbackFontSize,
            0d,
            fallbackColor,
            Alpha: 1d,
            Bold: false,
            Italic: false,
            Underline: false,
            Strike: false,
            typeface.RequestedTypeface,
            typeface.Typeface is null ? null : typeface.Source);
        return style.Merge(options.TextStyle);
    }

    private static ChartDataLabelOptions ResolveChartDataLabelOptions(ChartDataLabelOptions options, int index)
    {
        if (!options.Overrides.TryGetValue(index, out ChartDataLabelOverride dataLabel))
        {
            return options with { Overrides = EmptyChartDataLabelOverrides };
        }

        if (dataLabel.IsDeleted == true)
        {
            return ChartDataLabelOptions.None;
        }

        IReadOnlyDictionary<string, ChartBooleanOption> flags = ResolveChartDataLabelFlagOptions(options.FlagOptions, dataLabel.FlagOptions, overrideDefined: true);
        return options with
        {
            IsDeleted = false,
            ShowValue = ChartDataLabelFlagValue(flags, "showVal", labelsDefined: true),
            ShowPercent = ChartDataLabelFlagValue(flags, "showPercent", labelsDefined: true),
            ShowCategoryName = ChartDataLabelFlagValue(flags, "showCatName", labelsDefined: true),
            ShowSeriesName = ChartDataLabelFlagValue(flags, "showSerName", labelsDefined: true),
            ShowLeaderLines = dataLabel.ShowLeaderLines ?? options.ShowLeaderLines,
            ShowLegendKey = ChartDataLabelFlagValue(flags, "showLegendKey", labelsDefined: true),
            ShowBubbleSize = ChartDataLabelFlagValue(flags, "showBubbleSize", labelsDefined: true),
            LeaderLines = dataLabel.LeaderLines.IsDefined ? dataLabel.LeaderLines : options.LeaderLines,
            CustomText = string.IsNullOrEmpty(dataLabel.CustomText) ? options.CustomText : dataLabel.CustomText,
            CustomTextRuns = dataLabel.CustomTextRuns.Count == 0 ? options.CustomTextRuns : dataLabel.CustomTextRuns,
            PositionKind = string.IsNullOrEmpty(dataLabel.Position) ? options.PositionKind : dataLabel.PositionKind,
            Position = string.IsNullOrEmpty(dataLabel.Position) ? options.Position : dataLabel.Position,
            Separator = string.IsNullOrEmpty(dataLabel.Separator) ? options.Separator : dataLabel.Separator,
            NumberFormat = string.IsNullOrEmpty(dataLabel.NumberFormat) ? options.NumberFormat : dataLabel.NumberFormat,
            NumberFormatInfo = dataLabel.NumberFormatInfo.IsDefined ? dataLabel.NumberFormatInfo : options.NumberFormatInfo,
            Layout = dataLabel.Layout.HasLayout ? dataLabel.Layout : options.Layout,
            TextStyle = dataLabel.TextStyle,
            TextBodyProperties = IsChartTextBodyPropertiesEmpty(dataLabel.TextBodyProperties) ? options.TextBodyProperties : dataLabel.TextBodyProperties,
            ShapeStyle = dataLabel.ShapeStyle.IsEmpty ? options.ShapeStyle : dataLabel.ShapeStyle,
            FlagOptions = flags,
            Overrides = EmptyChartDataLabelOverrides
        };
    }

    private static ChartDataLabelOptions ResolveChartDataLabelOptionsForSeries(ChartDataLabelOptions plotOptions, IReadOnlyList<ChartDataLabelOptions> seriesOptions, int seriesIndex)
    {
        if (seriesIndex >= seriesOptions.Count || !seriesOptions[seriesIndex].IsDefined)
        {
            return plotOptions;
        }

        return MergeChartDataLabelOptions(plotOptions, seriesOptions[seriesIndex]);
    }

    private static ChartDataLabelOptions MergeChartDataLabelOptions(ChartDataLabelOptions baseOptions, ChartDataLabelOptions overrideOptions)
    {
        IReadOnlyDictionary<string, ChartBooleanOption> flags = ResolveChartDataLabelFlagOptions(baseOptions.FlagOptions, overrideOptions.FlagOptions, overrideOptions.IsDefined);
        bool labelsDefined = baseOptions.IsDefined || overrideOptions.IsDefined;
        return baseOptions with
        {
            IsDeleted = overrideOptions.IsDefined ? overrideOptions.IsDeleted : baseOptions.IsDeleted,
            ShowValue = ChartDataLabelFlagValue(flags, "showVal", labelsDefined),
            ShowPercent = ChartDataLabelFlagValue(flags, "showPercent", labelsDefined),
            ShowCategoryName = ChartDataLabelFlagValue(flags, "showCatName", labelsDefined),
            ShowSeriesName = ChartDataLabelFlagValue(flags, "showSerName", labelsDefined),
            ShowLeaderLines = ChartDataLabelFlagValue(flags, "showLeaderLines", labelsDefined),
            ShowLegendKey = ChartDataLabelFlagValue(flags, "showLegendKey", labelsDefined),
            ShowBubbleSize = ChartDataLabelFlagValue(flags, "showBubbleSize", labelsDefined),
            LeaderLines = overrideOptions.LeaderLines.IsDefined ? overrideOptions.LeaderLines : baseOptions.LeaderLines,
            CustomText = string.IsNullOrEmpty(overrideOptions.CustomText) ? baseOptions.CustomText : overrideOptions.CustomText,
            CustomTextRuns = overrideOptions.CustomTextRuns.Count == 0 ? baseOptions.CustomTextRuns : overrideOptions.CustomTextRuns,
            PositionKind = string.IsNullOrEmpty(overrideOptions.Position) ? baseOptions.PositionKind : overrideOptions.PositionKind,
            Position = string.IsNullOrEmpty(overrideOptions.Position) ? baseOptions.Position : overrideOptions.Position,
            Separator = string.IsNullOrEmpty(overrideOptions.Separator) ? baseOptions.Separator : overrideOptions.Separator,
            NumberFormat = string.IsNullOrEmpty(overrideOptions.NumberFormat) ? baseOptions.NumberFormat : overrideOptions.NumberFormat,
            NumberFormatInfo = overrideOptions.NumberFormatInfo.IsDefined ? overrideOptions.NumberFormatInfo : baseOptions.NumberFormatInfo,
            Layout = overrideOptions.Layout.HasLayout ? overrideOptions.Layout : baseOptions.Layout,
            // Each defined container already owns chart defaults plus its local
            // style. Missing properties request those defaults, not plot style.
            TextStyle = overrideOptions.TextStyle,
            TextBodyProperties = IsChartTextBodyPropertiesEmpty(overrideOptions.TextBodyProperties) ? baseOptions.TextBodyProperties : overrideOptions.TextBodyProperties,
            ShapeStyle = overrideOptions.ShapeStyle.IsEmpty ? baseOptions.ShapeStyle : overrideOptions.ShapeStyle,
            FlagOptions = flags,
            // A defined series owns its point declarations as well as its
            // component defaults; plot point requests do not leak into it.
            Overrides = overrideOptions.IsDefined ? overrideOptions.Overrides : baseOptions.Overrides,
            IsDefined = labelsDefined
        };
    }

    private static bool ChartDataLabelFlagValue(IReadOnlyDictionary<string, ChartBooleanOption> flags, string name, bool labelsDefined)
    {
        if (flags.TryGetValue(name, out ChartBooleanOption option) && option.IsDefined)
            return option.Value;

        // Office supplies components and keys for a defined label container.
        // Each defined plot/series/point container owns its component defaults;
        // explicit false in that active container survives. Styles still inherit.
        // The leader-line controls do not establish its default; keep that fallback.
        return labelsDefined && IsChartDataLabelComponentFlag(name);
    }

    private static bool IsChartDataLabelComponentFlag(string name) => name is
        "showVal" or "showPercent" or "showCatName" or "showSerName" or "showLegendKey" or "showBubbleSize";

    private static bool IsChartTextBodyPropertiesEmpty(PptxSceneChartTextBodyProperties properties)
    {
        return properties.RotationDegrees is null &&
            string.IsNullOrEmpty(properties.RotationValue) &&
            string.IsNullOrEmpty(properties.OrientationValue) &&
            string.IsNullOrEmpty(properties.VerticalOverflowValue) &&
            string.IsNullOrEmpty(properties.WrapValue);
    }

    // Radar axis labels route here alone (data labels use AddChartLabelRuns below);
    // Office applies no pair kerning to them (Power/Defense/Stamina read 0 to 4
    // against our GPOS pairs), so radar passes kerning off like the cartesian axes.
    private static TextRun CreateChartLabelRun(string text, double x, double y, double width, double height, ChartPlotBox plotBox, ChartTextStyle style, TextAlignment alignment, bool kerningEnabled = true)
    {
        return CreateChartTextRun(text, x, y, width, height, plotBox.X, plotBox.Y, plotBox.Width, plotBox.Height, style, alignment, kerningEnabled);
    }

    private static async ValueTask AddChartLabelRunsAsync(List<TextRun> runs, string text, ChartDataLabelOptions options, double x, double y, double width, double height, ChartPlotBox plotBox, ChartTextStyle style, TextAlignment alignment, PresentationFontResolver? fontResolver, List<ChartTextRunLink>? labelLinks, double descent = 0d)
    {
        ChartLayoutBox clipBox = ResolveDataLabelTextClipBox(plotBox, x, y - descent, width, height);
        if (options.CustomTextRuns.Count == 0)
        {
            runs.Add(CreateChartTextRun(text, x, y, width, height, clipBox.X, clipBox.Y, clipBox.Width, clipBox.Height, style, alignment));
            return;
        }

        await AddChartRichTextRunsAsync(runs, options.CustomTextRuns, text, x, y, width, height, clipBox.X, clipBox.Y, clipBox.Width, clipBox.Height, style, alignment, fontResolver, labelLinks).ConfigureAwait(false);
    }

    private static ChartLayoutBox ResolveDataLabelTextClipBox(ChartPlotBox plotBox, double x, double y, double width, double height)
    {
        // Labels paint in their own geometry, including automatic labels beyond
        // the data plot. Painting and link clipping use the same admitted region.
        double left = Math.Min(plotBox.X, x);
        double top = Math.Min(plotBox.Y, y);
        double right = Math.Max(plotBox.X + plotBox.Width, x + width);
        double bottom = Math.Max(plotBox.Y + plotBox.Height, y + height);
        return new ChartLayoutBox(left, top, Math.Max(1d, right - left), Math.Max(1d, bottom - top));
    }

    private static async ValueTask AddChartRichTextRunsAsync(List<TextRun> runs, IReadOnlyList<ChartTextRunOverride> richTextRuns, string fallbackText, double x, double y, double width, double height, double clipX, double clipY, double clipWidth, double clipHeight, ChartTextStyle style, TextAlignment alignment, PresentationFontResolver? fontResolver, List<ChartTextRunLink>? links = null)
    {
        if (richTextRuns.Count == 0)
        {
            runs.Add(CreateChartTextRun(fallbackText, x, y, width, height, clipX, clipY, clipWidth, clipHeight, style, alignment));
            return;
        }

        var textMeasurer = new ChartTextMeasurer(fontResolver);
        var measuredRuns = new List<ChartTextRunLayout>();
        foreach (ChartTextRunOverride run in richTextRuns)
        {
            if (string.IsNullOrEmpty(run.Text)) continue;
            ChartTextStyle runStyle = style.Merge(run.TextStyle);
            double runWidth = Math.Max(0d, await textMeasurer.MeasureAsync(run.Text, runStyle).ConfigureAwait(false));
            if (runWidth > 0d)
                measuredRuns.Add(new ChartTextRunLayout(run.Text, runStyle, runWidth,
                    HyperlinkClickId: run.HyperlinkClickId, HyperlinkClickAction: run.HyperlinkClickAction));
        }
        ChartTextRunLayout[] richRuns = measuredRuns.ToArray();
        if (richRuns.Length == 0)
        {
            runs.Add(CreateChartTextRun(fallbackText, x, y, width, height, clipX, clipY, clipWidth, clipHeight, style, alignment));
            return;
        }

        double totalWidth = richRuns.Sum(run => run.Width);
        double cursor = alignment switch
        {
            TextAlignment.Right => x + Math.Max(1d, width) - totalWidth,
            TextAlignment.Center => x + (Math.Max(1d, width) - totalWidth) / 2d,
            _ => x
        };

        foreach (ChartTextRunLayout run in richRuns)
        {
            double runWidth = Math.Max(0.1d, run.Width);
            TextRun richRun = CreateChartTextRun(run.Text, cursor, y, runWidth, height, clipX, clipY, clipWidth, clipHeight, run.Style, TextAlignment.Left) with { PreventCoalesce = true };
            runs.Add(richRun);
            if (links is not null && (!string.IsNullOrEmpty(run.HyperlinkClickId) || !string.IsNullOrEmpty(run.HyperlinkClickAction)))
            {
                links.Add(new ChartTextRunLink(richRun, run.HyperlinkClickId, run.HyperlinkClickAction));
            }
            cursor += runWidth;
        }
    }

    private static TextRun CreateChartTextRun(string text, double x, double y, double width, double height, double clipX, double clipY, double clipWidth, double clipHeight, ChartTextStyle style, TextAlignment alignment, bool kerningEnabled = true)
    {
        return new TextRun(
            text,
            x,
            y,
            Math.Max(1d, width),
            height,
            clipX,
            clipY,
            clipWidth,
            clipHeight,
            style.FontSize,
            style.CharacterSpacing,
            0d,
            style.Color,
            style.Alpha,
            null,
            Bold: style.Bold,
            Italic: style.Italic,
            Underline: style.Underline,
            Strike: style.Strike,
            KerningEnabled: kerningEnabled,
            alignment,
            FontFamily: style.FontFamily,
            RotationDegrees: 0d,
            RotationCenterX: 0d,
            RotationCenterY: 0d,
            FlipHorizontal: false,
            FlipVertical: false, PreventCoalesce: false, Outline: null, StrictClip: false);
    }

    private static ChartTextStyle ReadChartTextStyle(PptxTheme theme, XDocument chartXml, XElement? element, double fallbackFontSize)
    {
        return ReadChartTextStyle(theme, PptxColorMap.Default, chartXml, element, fallbackFontSize);
    }

    private static ChartTextStyle ReadChartTextStyle(PptxTheme theme, PptxColorMap colorMap, XDocument chartXml, XElement? element, double fallbackFontSize)
    {
        ChartTextStyle style = CreateDefaultChartTextStyle(theme, colorMap, fallbackFontSize);
        style = style.Merge(ToChartTextStyleOverride(PptxSceneBuilder.ReadChartTextStyleOverride(chartXml.Root, theme, colorMap)));
        style = style.Merge(ToChartTextStyleOverride(PptxSceneBuilder.ReadChartTextStyleOverride(element, theme, colorMap)));
        return style;
    }

    private static ChartTextStyle ReadSceneOrXmlChartTextStyle(PptxTheme theme, PptxSceneChart? sceneChart, PptxSceneChartAxis? sceneAxis, XDocument chartXml, XElement? element, double fallbackFontSize, string? chartStyleRole)
    {
        if (sceneChart is null)
        {
            return ReadChartTextStyle(theme, chartXml, element, fallbackFontSize);
        }

        ChartTextStyle style = CreateDefaultChartTextStyle(theme, sceneChart.ColorMap, fallbackFontSize);
        return style.Merge(ToChartTextStyleOverride(PptxSceneBuilder.ResolveChartAxisTextStyleOverride(sceneChart, sceneAxis, chartStyleRole)));
    }

    private static ChartTextStyle CreateDefaultChartTextStyle(PptxTheme theme, double fallbackFontSize)
    {
        return CreateDefaultChartTextStyle(theme, PptxColorMap.Default, fallbackFontSize);
    }

    private static ChartTextStyle CreateDefaultChartTextStyle(PptxTheme theme, PptxColorMap colorMap, double fallbackFontSize)
    {
        RgbColor fallbackColor = theme.TryResolveColor("tx1", colorMap, out RgbColor themeText)
            ? themeText
            : new RgbColor(0, 0, 0);
        PptxThemeTypefaceResolution typeface = ResolveChartThemeTypeface(theme);
        return new ChartTextStyle(
            typeface.Typeface,
            fallbackFontSize,
            0d,
            fallbackColor,
            Alpha: 1d,
            Bold: false,
            Italic: false,
            Underline: false,
            Strike: false,
            typeface.RequestedTypeface,
            typeface.Typeface is null ? null : typeface.Source);
    }

    private static ChartTextStyleOverride ToChartTextStyleOverride(PptxSceneChartTextStyleOverride style)
    {
        return new ChartTextStyleOverride(
            style.FontFamily,
            style.FontSize,
            style.CharacterSpacing,
            style.Color,
            style.Alpha,
            style.Bold,
            style.Italic,
            style.Underline,
            style.Strike,
            style.RequestedTypeface,
            style.TypefaceSource);
    }

    private static PptxThemeTypefaceResolution ResolveChartThemeTypeface(PptxTheme theme)
    {
        PptxThemeTypefaceResolution minorLatin = theme.ResolveTypefaceWithSource("+mn-lt");
        if (minorLatin.Typeface is not null)
        {
            return minorLatin;
        }

        PptxThemeTypefaceResolution majorLatin = theme.ResolveTypefaceWithSource("+mj-lt");
        return majorLatin.Typeface is null ? default : majorLatin;
    }

    private static ChartDataLabelOptions ReadChartDataLabelOptions(XElement chartElement, PptxTheme theme, PptxColorMap colorMap)
    {
        PptxSceneChartDataLabels labels = PptxSceneBuilder.ReadChartDataLabels(chartElement, theme, colorMap);
        return labels.IsDefined
            ? ToChartDataLabelOptions(labels, chartElement, ResolveDataLabelTextDefaults(null, chartElement, theme, colorMap),
                seriesScope: chartElement.Name == ChartNamespace + "ser")
            : ChartDataLabelOptions.None;
    }

    private static ChartDataLabelOptions ReadSceneOrXmlDataLabelOptions(PptxSceneChart? sceneChart, PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme, PptxColorMap colorMap)
    {
        return plot is null
            ? ReadChartDataLabelOptions(chartElement, theme, colorMap)
            : ToChartDataLabelOptions(plot.DataLabels, chartElement, ResolveDataLabelTextDefaults(sceneChart, chartElement, theme, colorMap),
                date1904: sceneChart?.Options.Date1904 == true);
    }

    private static IReadOnlyList<ChartDataLabelOptions> ReadSceneOrXmlSeriesDataLabelOptions(PptxSceneChart? sceneChart, PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme, PptxColorMap colorMap)
    {
        if (plot is not null)
        {
            PptxSceneChartTextStyleOverride defaults = ResolveDataLabelTextDefaults(sceneChart, chartElement, theme, colorMap);
            return plot.Series
                .Select(series => ToChartDataLabelOptions(series.DataLabels, chartElement, defaults,
                    date1904: sceneChart?.Options.Date1904 == true, seriesScope: true))
                .ToArray();
        }

        return chartElement
            .Elements(ChartNamespace + "ser")
            .Select(series => ReadChartDataLabelOptions(series, theme, colorMap))
            .ToArray();
    }

    private static PptxSceneChartTextStyleOverride ResolveDataLabelTextDefaults(PptxSceneChart? sceneChart, XElement chartElement, PptxTheme theme, PptxColorMap colorMap)
    {
        return sceneChart is not null
            ? PptxSceneBuilder.ResolveChartElementTextStyleOverride(sceneChart, default, "dataLabel")
            : PptxSceneBuilder.ReadChartTextStyleOverride(chartElement.AncestorsAndSelf(ChartNamespace + "chartSpace").FirstOrDefault(), theme, colorMap);
    }

    private static ChartDataLabelOptions ToChartDataLabelOptions(PptxSceneChartDataLabels labels, XElement chartElement,
        PptxSceneChartTextStyleOverride defaults, bool date1904 = false, bool seriesScope = false)
    {
        IReadOnlyDictionary<string, ChartBooleanOption> flags = ToChartDataLabelFlagOptions(labels);
        var options = new ChartDataLabelOptions(
            ChartDataLabelFlagValue(flags, "showVal", labels.IsDefined),
            ChartDataLabelFlagValue(flags, "showPercent", labels.IsDefined),
            ChartDataLabelFlagValue(flags, "showCatName", labels.IsDefined),
            ChartDataLabelFlagValue(flags, "showSerName", labels.IsDefined),
            ChartDataLabelFlagValue(flags, "showLeaderLines", labels.IsDefined),
            ChartDataLabelFlagValue(flags, "showLegendKey", labels.IsDefined),
            ChartDataLabelFlagValue(flags, "showBubbleSize", labels.IsDefined),
            ToChartDataLabelLeaderLines(labels.LeaderLines),
            string.Empty,
            [],
            labels.PositionKind,
            labels.Position,
            labels.Separator,
            labels.NumberFormat,
            ToChartNumberFormat(labels.NumberFormatInfo),
            labels.Layout,
            ToChartTextStyleOverride(PptxSceneBuilder.MergeChartTextStyleOverride(defaults, labels.TextStyle)),
            labels.TextBodyProperties,
            ToChartShapeStyle(labels.ShapeStyle),
            flags,
            EmptyChartDataLabelOverrides,
            labels.IsDefined,
            Date1904: date1904 || ResolveChartSpaceDate1904(chartElement),
            IsDeleted: labels.IsDeleted);
        // A disabled Cartesian component set suppresses its point requests.
        // Deletion hides default labels while retaining defined point scope.
        // Series and sector plots also have independent point-label scope.
        bool independentPoints = seriesScope || chartElement.Name == ChartNamespace + "pieChart" ||
            chartElement.Name == ChartNamespace + "doughnutChart";
        return independentPoints || options.HasVisibleContent || labels.IsDeleted == true
            ? options with { Overrides = ToChartDataLabelOverrides(labels.Overrides, defaults) }
            : options;
    }

    private static IReadOnlyDictionary<int, ChartDataLabelOverride> ToChartDataLabelOverrides(IReadOnlyList<PptxSceneChartDataLabelOverride> overrides,
        PptxSceneChartTextStyleOverride defaults)
    {
        if (overrides.Count == 0)
        {
            return EmptyChartDataLabelOverrides;
        }

        var result = new Dictionary<int, ChartDataLabelOverride>(overrides.Count);
        foreach (PptxSceneChartDataLabelOverride dataLabel in overrides)
        {
            result[dataLabel.Index] = new ChartDataLabelOverride(
                dataLabel.IsDeleted,
                dataLabel.ShowValue,
                dataLabel.ShowPercent,
                dataLabel.ShowCategoryName,
                dataLabel.ShowSeriesName,
                dataLabel.ShowLeaderLines,
                dataLabel.ShowLegendKey,
                dataLabel.ShowBubbleSize,
                ToChartDataLabelLeaderLines(dataLabel.LeaderLines),
                dataLabel.CustomText,
                ToChartTextRuns(dataLabel.CustomTextRuns),
                dataLabel.PositionKind,
                dataLabel.Position,
                dataLabel.Separator,
                dataLabel.NumberFormat,
                ToChartNumberFormat(dataLabel.NumberFormatInfo),
                dataLabel.Layout,
                // A defined point owns text defaults independently of its label
                // container. Resolve once from chart defaults and the point itself;
                // nullable source fields remain untouched in the scene model.
                ToChartTextStyleOverride(PptxSceneBuilder.MergeChartTextStyleOverride(defaults, dataLabel.TextStyle)),
                dataLabel.TextBodyProperties,
                ToChartShapeStyle(dataLabel.ShapeStyle),
                ToChartDataLabelOverrideFlagOptions(dataLabel));
        }

        return result;
    }

    private static ChartNumberFormat ToChartNumberFormat(PptxSceneChartNumberFormat numberFormat)
    {
        return numberFormat.IsDefined
            ? new ChartNumberFormat(numberFormat.IsDefined, numberFormat.FormatCode, numberFormat.SourceLinked, numberFormat.SourceLinkedValue)
            : default;
    }

    private static ChartDataLabelLeaderLines ToChartDataLabelLeaderLines(PptxSceneChartLeaderLines leaderLines)
    {
        return leaderLines.IsDefined
            ? new ChartDataLabelLeaderLines(IsDefined: true, ToChartSeriesStroke(leaderLines.Line, null))
            : ChartDataLabelLeaderLines.Empty;
    }

    private static IReadOnlyList<ChartTextRunOverride> ToChartTextRuns(IReadOnlyList<PptxSceneChartTextRun> runs)
    {
        return runs.Count == 0
            ? []
            : runs.Select(run => new ChartTextRunOverride(run.Text, ToChartTextStyleOverride(run.TextStyle), HyperlinkClickId: run.HyperlinkClickId, HyperlinkClickAction: run.HyperlinkClickAction)).ToArray();
    }

    private static IReadOnlyDictionary<string, ChartBooleanOption> ToChartDataLabelFlagOptions(PptxSceneChartDataLabels labels)
    {
        return new Dictionary<string, ChartBooleanOption>(StringComparer.Ordinal)
        {
            ["showVal"] = new(labels.ShowValue == true, labels.ShowValueValue, labels.ShowValue is not null),
            ["showPercent"] = new(labels.ShowPercent == true, labels.ShowPercentValue, labels.ShowPercent is not null),
            ["showCatName"] = new(labels.ShowCategoryName == true, labels.ShowCategoryNameValue, labels.ShowCategoryName is not null),
            ["showSerName"] = new(labels.ShowSeriesName == true, labels.ShowSeriesNameValue, labels.ShowSeriesName is not null),
            ["showLeaderLines"] = new(labels.ShowLeaderLines == true, labels.ShowLeaderLinesValue, labels.ShowLeaderLines is not null),
            ["showLegendKey"] = new(labels.ShowLegendKey == true, labels.ShowLegendKeyValue, labels.ShowLegendKey is not null),
            ["showBubbleSize"] = new(labels.ShowBubbleSize == true, labels.ShowBubbleSizeValue, labels.ShowBubbleSize is not null)
        };
    }

    private static IReadOnlyDictionary<string, ChartBooleanOption> ToChartDataLabelOverrideFlagOptions(PptxSceneChartDataLabelOverride label)
    {
        return new Dictionary<string, ChartBooleanOption>(StringComparer.Ordinal)
        {
            ["showVal"] = new(label.ShowValue == true, label.ShowValueValue, label.ShowValue is not null),
            ["showPercent"] = new(label.ShowPercent == true, label.ShowPercentValue, label.ShowPercent is not null),
            ["showCatName"] = new(label.ShowCategoryName == true, label.ShowCategoryNameValue, label.ShowCategoryName is not null),
            ["showSerName"] = new(label.ShowSeriesName == true, label.ShowSeriesNameValue, label.ShowSeriesName is not null),
            ["showLeaderLines"] = new(label.ShowLeaderLines == true, label.ShowLeaderLinesValue, label.ShowLeaderLines is not null),
            ["showLegendKey"] = new(label.ShowLegendKey == true, label.ShowLegendKeyValue, label.ShowLegendKey is not null),
            ["showBubbleSize"] = new(label.ShowBubbleSize == true, label.ShowBubbleSizeValue, label.ShowBubbleSize is not null)
        };
    }

    private static IReadOnlyDictionary<string, ChartBooleanOption> ResolveChartDataLabelFlagOptions(
        IReadOnlyDictionary<string, ChartBooleanOption> baseFlags,
        IReadOnlyDictionary<string, ChartBooleanOption> overrideFlags,
        bool overrideDefined)
    {
        if (overrideFlags.Count == 0)
        {
            return baseFlags;
        }

        var resolved = new Dictionary<string, ChartBooleanOption>(ChartDataLabelFlagNames.Length, StringComparer.Ordinal);
        foreach (string flagName in ChartDataLabelFlagNames)
        {
            if (overrideDefined && IsChartDataLabelComponentFlag(flagName))
            {
                resolved[flagName] = overrideFlags.TryGetValue(flagName, out ChartBooleanOption scopedOption)
                    ? scopedOption : new ChartBooleanOption(false, string.Empty, false);
            }
            else if (overrideFlags.TryGetValue(flagName, out ChartBooleanOption overrideOption) && overrideOption.IsDefined)
            {
                resolved[flagName] = overrideOption;
            }
            else if (baseFlags.TryGetValue(flagName, out ChartBooleanOption baseOption))
            {
                resolved[flagName] = baseOption;
            }
            else
            {
                resolved[flagName] = new ChartBooleanOption(false, string.Empty, false);
            }
        }

        return resolved;
    }

    // RV14: per-frame first-match category index built once; O(1) lookups replace
    // the per-label linear scan. Duplicates keep first-match precedence; sparse and
    // blank entries behave exactly like the scan (first HasText match or empty).
    private static Dictionary<int, string> BuildCategoryLabelIndex(ChartIndexedTextVector categoryLabels)
    {
        var index = new Dictionary<int, string>();
        foreach (ChartIndexedTextPoint point in categoryLabels.Points ?? [])
        {
            if (point.HasText && !index.ContainsKey(point.Index))
            {
                index[point.Index] = point.Text;
            }
        }

        return index;
    }

    private static string FormatCartesianDataLabel(
        double value,
        int seriesIndex,
        int categoryIndex,
        ChartIndexedNumberPoint point,
        ChartIndexedNumberPoint? workbookPoint,
        string? valueFormatCode,
        ChartDataLabelOptions options,
        IReadOnlyDictionary<int, string> categoryLabelIndex,
        IReadOnlyList<ChartSeriesNameRecord> seriesNames)
    {
        if (!string.IsNullOrWhiteSpace(options.CustomText))
        {
            return options.CustomText;
        }

        var parts = new List<string>(3);
        string seriesName = GetActiveSeriesName(seriesNames, seriesIndex);
        if (options.ShowSeriesName && !string.IsNullOrWhiteSpace(seriesName))
        {
            parts.Add(seriesName);
        }

        if (options.ShowCategoryName &&
            categoryLabelIndex.TryGetValue(categoryIndex, out string? categoryLabel) &&
            !string.IsNullOrWhiteSpace(categoryLabel))
        {
            parts.Add(categoryLabel);
        }

        if (options.ShowValue)
        {
            parts.Add(FormatChartDataLabelValue(value, options, workbookPoint ?? point, valueFormatCode));
        }

        return string.Join(GetChartDataLabelSeparator(options), parts);
    }

    private static string GetActiveSeriesName(IReadOnlyList<ChartSeriesNameRecord> seriesNames, int seriesIndex)
    {
        return seriesIndex >= 0 && seriesIndex < seriesNames.Count
            ? seriesNames[seriesIndex].ActiveName
            : string.Empty;
    }

    private static string FormatChartDataLabelValue(double value, ChartDataLabelOptions options, ChartIndexedNumberPoint? sourcePoint, string? sourceFormatCode)
    {
        if (ResolveSourceLinkedChartNumberFormatCode(options.NumberFormatInfo, sourcePoint) is { } sourceFormat)
        {
            return FormatChartNumber(value, sourceFormat, options.Date1904);
        }

        return FormatChartDataLabelValue(value, options.NumberFormatInfo, options.NumberFormat, sourceFormatCode, options.Date1904);
    }

    private static string FormatChartDataLabelValue(double value, ChartNumberFormat numberFormat, string legacyNumberFormat, string? sourceFormatCode, bool date1904 = false)
    {
        if (IsRenderableChartNumberFormat(numberFormat))
        {
            return FormatChartNumber(value, numberFormat.FormatCode, date1904);
        }

        return !string.IsNullOrWhiteSpace(legacyNumberFormat) &&
            !string.Equals(legacyNumberFormat, "General", StringComparison.OrdinalIgnoreCase)
            ? FormatChartNumber(value, legacyNumberFormat, date1904)
            : IsRenderableChartFormatCode(sourceFormatCode)
                ? FormatChartNumber(value, sourceFormatCode!, date1904)
            : FormatChartAxisLabel(value, null);
    }

    private static string? ResolveSourceLinkedChartNumberFormatCode(ChartNumberFormat numberFormat, ChartIndexedNumberPoint? sourcePoint)
    {
        if (sourcePoint is null)
        {
            return null;
        }

        ChartWorkbookRangeCell cell = sourcePoint.Value.WorkbookCell;
        return ResolveSourceLinkedChartNumberFormatCode(
            numberFormat,
            cell.StyleNumberFormatCode,
            cell.StyleAppliesNumberFormat);
    }

    private static string? ResolveSourceLinkedChartNumberFormatCode(
        ChartNumberFormat numberFormat,
        string? workbookFormatCode,
        bool? workbookAppliesNumberFormat)
    {
        // Date-like workbook codes flow through since date serial rendering is explicit
        // (S04); the shared formatter renders them with the chart date system.
        if (numberFormat.SourceLinked != true ||
            workbookAppliesNumberFormat == false ||
            !IsRenderableChartFormatCode(workbookFormatCode))
        {
            return null;
        }

        return workbookFormatCode;
    }

    private static string GetChartDataLabelSeparator(ChartDataLabelOptions options)
    {
        return string.IsNullOrEmpty(options.Separator) ? ", " : options.Separator;
    }
}
