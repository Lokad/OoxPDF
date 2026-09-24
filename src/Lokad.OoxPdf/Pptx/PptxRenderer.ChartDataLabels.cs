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
            return options;
        }

        if (dataLabel.IsDeleted == true)
        {
            return ChartDataLabelOptions.None;
        }

        ChartTextStyleOverride textStyle = new(
            dataLabel.TextStyle.FontFamily ?? options.TextStyle.FontFamily,
            dataLabel.TextStyle.FontSize ?? options.TextStyle.FontSize,
            dataLabel.TextStyle.CharacterSpacing ?? options.TextStyle.CharacterSpacing,
            dataLabel.TextStyle.Color ?? options.TextStyle.Color,
            dataLabel.TextStyle.Alpha ?? options.TextStyle.Alpha,
            dataLabel.TextStyle.Bold ?? options.TextStyle.Bold,
            dataLabel.TextStyle.Italic ?? options.TextStyle.Italic,
            dataLabel.TextStyle.Underline ?? options.TextStyle.Underline,
            dataLabel.TextStyle.Strike ?? options.TextStyle.Strike,
            dataLabel.TextStyle.FontFamily is null ? options.TextStyle.RequestedTypeface : dataLabel.TextStyle.RequestedTypeface,
            dataLabel.TextStyle.FontFamily is null ? options.TextStyle.TypefaceSource : dataLabel.TextStyle.TypefaceSource);
        return options with
        {
            ShowValue = dataLabel.ShowValue ?? options.ShowValue,
            ShowPercent = dataLabel.ShowPercent ?? options.ShowPercent,
            ShowCategoryName = dataLabel.ShowCategoryName ?? options.ShowCategoryName,
            ShowSeriesName = dataLabel.ShowSeriesName ?? options.ShowSeriesName,
            ShowLeaderLines = dataLabel.ShowLeaderLines ?? options.ShowLeaderLines,
            ShowLegendKey = dataLabel.ShowLegendKey ?? options.ShowLegendKey,
            ShowBubbleSize = dataLabel.ShowBubbleSize ?? options.ShowBubbleSize,
            LeaderLines = dataLabel.LeaderLines.IsDefined ? dataLabel.LeaderLines : options.LeaderLines,
            CustomText = string.IsNullOrEmpty(dataLabel.CustomText) ? options.CustomText : dataLabel.CustomText,
            CustomTextRuns = dataLabel.CustomTextRuns.Count == 0 ? options.CustomTextRuns : dataLabel.CustomTextRuns,
            PositionKind = string.IsNullOrEmpty(dataLabel.Position) ? options.PositionKind : dataLabel.PositionKind,
            Position = string.IsNullOrEmpty(dataLabel.Position) ? options.Position : dataLabel.Position,
            Separator = string.IsNullOrEmpty(dataLabel.Separator) ? options.Separator : dataLabel.Separator,
            NumberFormat = string.IsNullOrEmpty(dataLabel.NumberFormat) ? options.NumberFormat : dataLabel.NumberFormat,
            NumberFormatInfo = dataLabel.NumberFormatInfo.IsDefined ? dataLabel.NumberFormatInfo : options.NumberFormatInfo,
            Layout = dataLabel.Layout.HasLayout ? dataLabel.Layout : options.Layout,
            TextStyle = textStyle,
            TextBodyProperties = IsChartTextBodyPropertiesEmpty(dataLabel.TextBodyProperties) ? options.TextBodyProperties : dataLabel.TextBodyProperties,
            ShapeStyle = dataLabel.ShapeStyle.IsEmpty ? options.ShapeStyle : dataLabel.ShapeStyle,
            FlagOptions = ResolveChartDataLabelFlagOptions(options.FlagOptions, dataLabel.FlagOptions),
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
        IReadOnlyDictionary<string, ChartBooleanOption> flags = ResolveChartDataLabelFlagOptions(baseOptions.FlagOptions, overrideOptions.FlagOptions);
        ChartTextStyleOverride textStyle = new(
            overrideOptions.TextStyle.FontFamily ?? baseOptions.TextStyle.FontFamily,
            overrideOptions.TextStyle.FontSize ?? baseOptions.TextStyle.FontSize,
            overrideOptions.TextStyle.CharacterSpacing ?? baseOptions.TextStyle.CharacterSpacing,
            overrideOptions.TextStyle.Color ?? baseOptions.TextStyle.Color,
            overrideOptions.TextStyle.Alpha ?? baseOptions.TextStyle.Alpha,
            overrideOptions.TextStyle.Bold ?? baseOptions.TextStyle.Bold,
            overrideOptions.TextStyle.Italic ?? baseOptions.TextStyle.Italic,
            overrideOptions.TextStyle.Underline ?? baseOptions.TextStyle.Underline,
            overrideOptions.TextStyle.Strike ?? baseOptions.TextStyle.Strike,
            overrideOptions.TextStyle.FontFamily is null ? baseOptions.TextStyle.RequestedTypeface : overrideOptions.TextStyle.RequestedTypeface,
            overrideOptions.TextStyle.FontFamily is null ? baseOptions.TextStyle.TypefaceSource : overrideOptions.TextStyle.TypefaceSource);

        return baseOptions with
        {
            ShowValue = ChartDataLabelFlagValue(flags, "showVal"),
            ShowPercent = ChartDataLabelFlagValue(flags, "showPercent"),
            ShowCategoryName = ChartDataLabelFlagValue(flags, "showCatName"),
            ShowSeriesName = ChartDataLabelFlagValue(flags, "showSerName"),
            ShowLeaderLines = ChartDataLabelFlagValue(flags, "showLeaderLines"),
            ShowLegendKey = ChartDataLabelFlagValue(flags, "showLegendKey"),
            ShowBubbleSize = ChartDataLabelFlagValue(flags, "showBubbleSize"),
            LeaderLines = overrideOptions.LeaderLines.IsDefined ? overrideOptions.LeaderLines : baseOptions.LeaderLines,
            CustomText = string.IsNullOrEmpty(overrideOptions.CustomText) ? baseOptions.CustomText : overrideOptions.CustomText,
            CustomTextRuns = overrideOptions.CustomTextRuns.Count == 0 ? baseOptions.CustomTextRuns : overrideOptions.CustomTextRuns,
            PositionKind = string.IsNullOrEmpty(overrideOptions.Position) ? baseOptions.PositionKind : overrideOptions.PositionKind,
            Position = string.IsNullOrEmpty(overrideOptions.Position) ? baseOptions.Position : overrideOptions.Position,
            Separator = string.IsNullOrEmpty(overrideOptions.Separator) ? baseOptions.Separator : overrideOptions.Separator,
            NumberFormat = string.IsNullOrEmpty(overrideOptions.NumberFormat) ? baseOptions.NumberFormat : overrideOptions.NumberFormat,
            NumberFormatInfo = overrideOptions.NumberFormatInfo.IsDefined ? overrideOptions.NumberFormatInfo : baseOptions.NumberFormatInfo,
            Layout = overrideOptions.Layout.HasLayout ? overrideOptions.Layout : baseOptions.Layout,
            TextStyle = textStyle,
            TextBodyProperties = IsChartTextBodyPropertiesEmpty(overrideOptions.TextBodyProperties) ? baseOptions.TextBodyProperties : overrideOptions.TextBodyProperties,
            ShapeStyle = overrideOptions.ShapeStyle.IsEmpty ? baseOptions.ShapeStyle : overrideOptions.ShapeStyle,
            FlagOptions = flags,
            Overrides = MergeChartDataLabelOverrides(baseOptions.Overrides, overrideOptions.Overrides),
            IsDefined = baseOptions.IsDefined || overrideOptions.IsDefined
        };
    }

    private static bool ChartDataLabelFlagValue(IReadOnlyDictionary<string, ChartBooleanOption> flags, string name)
    {
        return flags.TryGetValue(name, out ChartBooleanOption option) && option.Value;
    }

    private static IReadOnlyDictionary<int, ChartDataLabelOverride> MergeChartDataLabelOverrides(IReadOnlyDictionary<int, ChartDataLabelOverride> baseOverrides, IReadOnlyDictionary<int, ChartDataLabelOverride> overrideOverrides)
    {
        if (baseOverrides.Count == 0)
        {
            return overrideOverrides;
        }

        if (overrideOverrides.Count == 0)
        {
            return baseOverrides;
        }

        var merged = new Dictionary<int, ChartDataLabelOverride>(baseOverrides);
        foreach (KeyValuePair<int, ChartDataLabelOverride> item in overrideOverrides)
        {
            merged[item.Key] = item.Value;
        }

        return merged;
    }

    private static bool IsChartTextBodyPropertiesEmpty(PptxSceneChartTextBodyProperties properties)
    {
        return properties.RotationDegrees is null &&
            string.IsNullOrEmpty(properties.RotationValue) &&
            string.IsNullOrEmpty(properties.OrientationValue) &&
            string.IsNullOrEmpty(properties.VerticalOverflowValue);
    }

    // Radar axis labels route here alone (data labels use AddChartLabelRuns below);
    // Office applies no pair kerning to them (Power/Defense/Stamina read 0 to 4
    // against our GPOS pairs), so radar passes kerning off like the cartesian axes.
    private static TextRun CreateChartLabelRun(string text, double x, double y, double width, double height, ChartPlotBox plotBox, ChartTextStyle style, TextAlignment alignment, bool kerningEnabled = true)
    {
        return CreateChartTextRun(text, x, y, width, height, plotBox.X, plotBox.Y, plotBox.Width, plotBox.Height, style, alignment, kerningEnabled);
    }

    private static void AddChartLabelRuns(List<TextRun> runs, string text, ChartDataLabelOptions options, double x, double y, double width, double height, ChartPlotBox plotBox, ChartTextStyle style, TextAlignment alignment, PresentationFontResolver? fontResolver, List<ChartTextRunLink>? labelLinks)
    {
        ChartLayoutBox clipBox = ResolveDataLabelTextClipBox(plotBox, options, x, y, width, height);
        if (options.CustomTextRuns.Count == 0)
        {
            runs.Add(CreateChartTextRun(text, x, y, width, height, clipBox.X, clipBox.Y, clipBox.Width, clipBox.Height, style, alignment));
            return;
        }

        AddChartRichTextRuns(runs, options.CustomTextRuns, text, x, y, width, height, clipBox.X, clipBox.Y, clipBox.Width, clipBox.Height, style, alignment, fontResolver, labelLinks);
    }

    private static ChartLayoutBox ResolveDataLabelTextClipBox(ChartPlotBox plotBox, ChartDataLabelOptions options, double x, double y, double width, double height)
    {
        if (!options.Layout.HasLayout)
        {
            return new ChartLayoutBox(plotBox.X, plotBox.Y, plotBox.Width, plotBox.Height);
        }

        double left = Math.Min(plotBox.X, x);
        double top = Math.Min(plotBox.Y, y);
        double right = Math.Max(plotBox.X + plotBox.Width, x + width);
        double bottom = Math.Max(plotBox.Y + plotBox.Height, y + height);
        return new ChartLayoutBox(left, top, Math.Max(1d, right - left), Math.Max(1d, bottom - top));
    }

    private static void AddChartRichTextRuns(List<TextRun> runs, IReadOnlyList<ChartTextRunOverride> richTextRuns, string fallbackText, double x, double y, double width, double height, double clipX, double clipY, double clipWidth, double clipHeight, ChartTextStyle style, TextAlignment alignment, PresentationFontResolver? fontResolver, List<ChartTextRunLink>? links = null)
    {
        if (richTextRuns.Count == 0)
        {
            runs.Add(CreateChartTextRun(fallbackText, x, y, width, height, clipX, clipY, clipWidth, clipHeight, style, alignment));
            return;
        }

        var textMeasurer = new ChartTextMeasurer(fontResolver);
        ChartTextRunLayout[] richRuns = richTextRuns
            .Where(run => !string.IsNullOrEmpty(run.Text))
            .Select(run =>
            {
                ChartTextStyle runStyle = style.Merge(run.TextStyle);
                return new ChartTextRunLayout(run.Text, runStyle, Math.Max(0d, textMeasurer.Measure(run.Text, runStyle)), HyperlinkClickId: run.HyperlinkClickId, HyperlinkClickAction: run.HyperlinkClickAction);
            })
            .Where(run => run.Width > 0d)
            .ToArray();
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
            ? ToChartDataLabelOptions(sceneChart: null, labels, chartElement)
            : ChartDataLabelOptions.None;
    }

    private static ChartDataLabelOptions ReadSceneOrXmlDataLabelOptions(PptxSceneChart? sceneChart, PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme, PptxColorMap colorMap)
    {
        return plot is null
            ? ReadChartDataLabelOptions(chartElement, theme, colorMap)
            : new ChartDataLabelOptions(
                plot.DataLabels.ShowValue == true,
                plot.DataLabels.ShowPercent == true,
                plot.DataLabels.ShowCategoryName == true,
                plot.DataLabels.ShowSeriesName == true,
                plot.DataLabels.ShowLeaderLines == true,
                plot.DataLabels.ShowLegendKey == true,
                plot.DataLabels.ShowBubbleSize == true,
                ToChartDataLabelLeaderLines(plot.DataLabels.LeaderLines),
                string.Empty,
                [],
                plot.DataLabels.PositionKind,
                plot.DataLabels.Position,
                plot.DataLabels.Separator,
                plot.DataLabels.NumberFormat,
                ToChartNumberFormat(plot.DataLabels.NumberFormatInfo),
                plot.DataLabels.Layout,
                ToChartDataLabelTextStyleOverride(sceneChart, plot.DataLabels),
                plot.DataLabels.TextBodyProperties,
                ToChartShapeStyle(plot.DataLabels.ShapeStyle),
                ToChartDataLabelFlagOptions(plot.DataLabels),
                ToChartDataLabelOverrides(plot.DataLabels.Overrides),
                plot.DataLabels.IsDefined,
                Date1904: sceneChart?.Options.Date1904 == true || ResolveChartSpaceDate1904(chartElement));
    }

    private static IReadOnlyList<ChartDataLabelOptions> ReadSceneOrXmlSeriesDataLabelOptions(PptxSceneChart? sceneChart, PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme, PptxColorMap colorMap)
    {
        if (plot is not null)
        {
            return plot.Series
                .Select(series => ToChartDataLabelOptions(sceneChart, series.DataLabels, chartElement))
                .ToArray();
        }

        return chartElement
            .Elements(ChartNamespace + "ser")
            .Select(series => ReadChartDataLabelOptions(series, theme, colorMap))
            .ToArray();
    }

    private static ChartDataLabelOptions ToChartDataLabelOptions(PptxSceneChart? sceneChart, PptxSceneChartDataLabels labels, XElement? chartElement = null)
    {
        return new ChartDataLabelOptions(
            labels.ShowValue == true,
            labels.ShowPercent == true,
            labels.ShowCategoryName == true,
            labels.ShowSeriesName == true,
            labels.ShowLeaderLines == true,
            labels.ShowLegendKey == true,
            labels.ShowBubbleSize == true,
            ToChartDataLabelLeaderLines(labels.LeaderLines),
            string.Empty,
            [],
            labels.PositionKind,
            labels.Position,
            labels.Separator,
            labels.NumberFormat,
            ToChartNumberFormat(labels.NumberFormatInfo),
            labels.Layout,
            ToChartDataLabelTextStyleOverride(sceneChart, labels),
            labels.TextBodyProperties,
            ToChartShapeStyle(labels.ShapeStyle),
            ToChartDataLabelFlagOptions(labels),
            ToChartDataLabelOverrides(labels.Overrides),
            labels.IsDefined,
            Date1904: sceneChart?.Options.Date1904 == true || ResolveChartSpaceDate1904(chartElement));
    }

    private static ChartTextStyleOverride ToChartDataLabelTextStyleOverride(PptxSceneChart? sceneChart, PptxSceneChartDataLabels labels)
    {
        ChartTextStyleOverride style = ChartTextStyleOverride.Empty;
        if (sceneChart is not null)
        {
            style = ToChartTextStyleOverride(PptxSceneBuilder.ResolveChartDataLabelTextStyleOverride(sceneChart, labels));
            return style;
        }

        return ToChartTextStyleOverride(labels.TextStyle);
    }

    private static IReadOnlyDictionary<int, ChartDataLabelOverride> ToChartDataLabelOverrides(IReadOnlyList<PptxSceneChartDataLabelOverride> overrides)
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
                ToChartTextStyleOverride(dataLabel.TextStyle),
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
        IReadOnlyDictionary<string, ChartBooleanOption> overrideFlags)
    {
        if (overrideFlags.Count == 0)
        {
            return baseFlags;
        }

        var resolved = new Dictionary<string, ChartBooleanOption>(ChartDataLabelFlagNames.Length, StringComparer.Ordinal);
        foreach (string flagName in ChartDataLabelFlagNames)
        {
            if (overrideFlags.TryGetValue(flagName, out ChartBooleanOption overrideOption) && overrideOption.IsDefined)
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

    private static string FormatChartDataLabelValue(double value, ChartDataLabelOptions options)
    {
        return FormatChartDataLabelValue(value, options.NumberFormatInfo, options.NumberFormat, sourceFormatCode: null, date1904: options.Date1904);
    }

    private static string FormatChartDataLabelValue(double value, ChartDataLabelOptions options, ChartIndexedNumberPoint? sourcePoint)
    {
        return FormatChartDataLabelValue(value, options, sourcePoint, sourceFormatCode: null);
    }

    private static string FormatChartDataLabelValue(double value, ChartDataLabelOptions options, ChartIndexedNumberPoint? sourcePoint, string? sourceFormatCode)
    {
        if (ResolveSourceLinkedChartNumberFormatCode(options.NumberFormatInfo, sourcePoint) is { } sourceFormat)
        {
            return FormatChartNumber(value, sourceFormat, options.Date1904);
        }

        return FormatChartDataLabelValue(value, options.NumberFormatInfo, options.NumberFormat, sourceFormatCode, options.Date1904);
    }

    private static string FormatChartDataLabelValue(double value, ChartNumberFormat numberFormat, string legacyNumberFormat, bool date1904 = false)
    {
        return FormatChartDataLabelValue(value, numberFormat, legacyNumberFormat, sourceFormatCode: null, date1904: date1904);
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