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
    private static IEnumerable<XElement> ReadChartCategoryAxes(XDocument chartXml)
    {
        return chartXml
            .Descendants(ChartNamespace + "plotArea")
            .FirstOrDefault()?
            .Elements()
            .Where(element => element.Name == ChartNamespace + "catAx" || element.Name == ChartNamespace + "dateAx") ??
            [];
    }

    private static ChartValueExtents ReadSceneBubbleChartValueAxisExtents(PptxSceneChartAxis? axis, ChartValueExtents dataExtents, double boundsTickTargetCount = PptxChartMetricRules.BubbleAxisBoundsTickTargetCount, bool preferUnitOneOverTwo = true)
    {
        // Measurement keeps the actual extrema. The existing point-chart
        // range fallback owns zero inclusion and its minimum unit-sized span.
        ChartValueExtents fallback = new(Math.Min(0d, dataExtents.Min), Math.Max(1d, dataExtents.Max));
        return ReadSceneChartValueAxisExtents(axis, fallback, useNearMaximumHeadroom: true,
            nearMaximumHeadroomRatio: PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio,
            boundsTickTargetCount: boundsTickTargetCount, preferUnitOneOverTwo: preferUnitOneOverTwo);
    }

    private static (ChartValueExtents Extents, ChartValueAxisRenderOptions Options) ReadChartPointValueAxisScale(ChartAxisSource source, ChartValueExtents data, ChartValueExtents occupied, bool bubble, bool horizontal,
        double? automaticTickTargetCount = null)
    {
        bool scatterX = !bubble && horizontal;
        double target = scatterX ? PptxChartMetricRules.ScatterXAxisNiceTickTargetCount : PptxChartMetricRules.BubbleAxisNiceTickTargetCount;
        ChartValueExtents extents = ReadSceneBubbleChartValueAxisExtents(source.SceneAxis, data,
            scatterX ? PptxChartMetricRules.ScatterXAxisNiceTickTargetCount : PptxChartMetricRules.BubbleAxisBoundsTickTargetCount, !scatterX);
        PptxSceneChartAxis? axis = source.SceneAxis;
        ChartAxisDeclaration? declaration = ReadChartAxisDeclaration(source.SceneAxis);
        double? automaticTarget = !bubble && !horizontal
            ? PptxChartMetricRules.AxisNiceTickTargetCount : automaticTickTargetCount;
        bool automaticUnit = !bubble && automaticTarget is { } count && count >= 2d && double.IsFinite(count) &&
            declaration is { MajorUnit: null } && occupied.Max > occupied.Min;
        // Authored units, inferred vertical units and measured source X plots
        // share bounds and emitted intervals. Other automatic X layouts and
        // bubbles keep their existing fallback.
        ChartValueExtents owned = default;
        double unit = 0d;
        bool resolvedIntervals = !bubble && declaration is { } declared &&
            (automaticUnit || declared.MajorUnit is not null) &&
            TryResolvePointAxisRange(occupied, declared.Minimum, declared.Maximum, declared.MajorUnit, automaticTarget ?? target, out owned, out unit);
        if (resolvedIntervals) extents = owned;
        ChartValueAxisRenderOptions options = ReadSceneChartBubbleValueAxisOptions(axis, extents, target);
        if (resolvedIntervals) options = options with { Units = options.Units with { MajorUnit = unit }, TickPolicy = ChartAxisTickPolicy.SourceMinimumIntervals };
        return (extents, options);
    }

    private readonly record struct ChartAxisDeclaration(double? Minimum, double? Maximum, double? MajorUnit);

    private static ChartAxisDeclaration? ReadChartAxisDeclaration(PptxSceneChartAxis? axis)
    {
        if (axis is not { HasScaling: true, HasLogarithmicScale: false }) return null;
        var minimum = (axis.Minimum, axis.MinimumValue);
        var maximum = (axis.Maximum, axis.MaximumValue);
        var major = (axis.MajorUnit, axis.MajorUnitValue);
        if (minimum.Item1 is null && !string.IsNullOrEmpty(minimum.Item2) ||
            maximum.Item1 is null && !string.IsNullOrEmpty(maximum.Item2) ||
            major.Item1 is null && !string.IsNullOrEmpty(major.Item2) ||
            minimum.Item1 is { } low && !double.IsFinite(low) || maximum.Item1 is { } high && !double.IsFinite(high) ||
            major.Item1 is { } unit && (!(unit > 0d) || !double.IsFinite(unit))) return null;
        return new(minimum.Item1, maximum.Item1, major.Item1);
    }

    private static ChartValueExtents IncludePointAxisZero(ChartValueExtents occupied, double? minimum, double? maximum)
    {
        double lower = minimum ?? occupied.Min, upper = maximum ?? occupied.Max;
        if (minimum is null && lower >= 0d && upper > 0d && (occupied.Min == occupied.Max || lower / upper < 5d / 6d)) lower = 0d;
        if (maximum is null && upper <= 0d && lower < 0d && (occupied.Min == occupied.Max || upper / lower < 5d / 6d)) upper = 0d;
        return new(lower, upper);
    }

    private static double? ReadScatterSourceIntervalTarget(PptxTheme theme, PptxSceneChart chart,
        ChartFrameBox frame, ChartAxisSource xAxis)
    {
        PptxSceneChartManualLayout source = chart.PlotAreaLayout;
        // Resolve the same authored inner width used for placement before caption
        // admission. Outer, automatic and unresolved modes retain their fallback.
        if (source.LayoutTargetKind != PptxSceneChartManualLayoutTarget.Inner ||
            !TryBuildManualPlotLayout(source, frame, GetDefaultChartPlotBox(frame), out ChartPlotLayout layout) ||
            !TryResolveCartesianPlotWidth(source, frame, layout.PlotBox, out var plot)) return null;
        ChartTextStyle style = ReadSceneChartTextStyle(theme, chart, xAxis.SceneAxis,
            PptxChartMetricRules.ValueAxisFallbackFontSize, "valueAxis");
        if (!(style.FontSize > 0d) || !double.IsFinite(style.FontSize)) return null;
        double slots = Math.Floor(plot.Width / (4d * style.FontSize));
        // Count the intervals between font-relative caption slots; dense and
        // oversized plots share the existing numeric-axis target ceiling.
        return double.IsFinite(slots) ? Math.Clamp(slots - 1d, 2d, PptxChartMetricRules.AxisNiceHorizontalValueTickTargetCount) : null;
    }

    private static bool TryResolvePointAxisRange(ChartValueExtents occupied, double? minimum, double? maximum,
        double? declaredUnit, double tickTargetCount, out ChartValueExtents result, out double unit, double headroomRatio = .05d)
    {
        result = default;
        unit = declaredUnit ?? 0d;
        if (!double.IsFinite(occupied.Min) || !double.IsFinite(occupied.Max) || occupied.Min > occupied.Max ||
            minimum is { } min && !double.IsFinite(min) || maximum is { } max && !double.IsFinite(max)) return false;
        if (minimum is { } low && maximum is { } high)
        {
            result = new(low, high);
            if (high > low && declaredUnit is null) unit = ChooseChartAxisMajorUnit(high - low, tickTargetCount);
            return high > low && unit > 0d && double.IsFinite(unit);
        }
        var (lower, upper) = IncludePointAxisZero(occupied, minimum, maximum);
        if (lower == 0d && upper == 0d && minimum is null && maximum is null)
        {
            if (declaredUnit is null) unit = ChooseChartAxisMajorUnit(1d, tickTargetCount);
            result = new(0d, Math.Ceiling(1d / unit) * unit);
            return double.IsFinite(result.Max) && result.Max > 0d;
        }
        double span = upper - lower;
        if (!(span > 0d) || !double.IsFinite(span)) return false;
        if (declaredUnit is null) unit = ChooseChartAxisMajorUnit(span, tickTargetCount);
        if (!(unit > 0d) || !double.IsFinite(unit)) return false;
        // Office fixed-unit controls retain zero for wide single-sign data;
        // scatter ends reserve five percent; bubbles have already admitted
        // physical radius padding and request no additional headroom.
        double paddedLower = minimum ?? (lower == 0d ? 0d : lower - headroomRatio * span);
        double paddedUpper = maximum ?? (upper == 0d ? 0d : upper + headroomRatio * span);
        if (minimum is { } origin) paddedUpper = origin + Math.Ceiling((paddedUpper - origin) / unit) * unit;
        else if (maximum is { } end) paddedLower = end - Math.Ceiling((end - paddedLower) / unit) * unit;
        else
        {
            paddedLower = Math.Floor(paddedLower / unit) * unit;
            paddedUpper = Math.Ceiling(paddedUpper / unit) * unit;
        }
        result = new(paddedLower, paddedUpper);
        return double.IsFinite(paddedLower) && double.IsFinite(paddedUpper) && paddedUpper > paddedLower;
    }

    private static ChartValueExtents ReadSceneChartValueAxisExtents(PptxSceneChartAxis? axis, ChartValueExtents fallback, bool useNearMaximumHeadroom, double nearMaximumHeadroomRatio, double boundsTickTargetCount = PptxChartMetricRules.AxisNiceTickTargetCount, bool preferUnitOneOverTwo = false)
    {
        if (axis is not { HasScaling: true }) return fallback;
        double? minimum = axis.Minimum;
        double? maximum = axis.Maximum;

        // Negative-only data uses the same outer rounding as positive data.
        // Reflect authored bounds too, so a partial bound keeps its owner.
        // Mixed-sign and degenerate inputs retain the existing rounding path.
        bool reflect = fallback.Max <= 0d && fallback.Min < 0d;
        ChartValueExtents data = reflect ? new ChartValueExtents(-fallback.Max, -fallback.Min) : fallback;
        if (reflect) (minimum, maximum) = (-maximum, -minimum);
        double min = minimum ?? GetNiceChartAxisMin(data.Min, data.Max);
        double max = maximum ?? GetNiceChartAxisMax(data.Max, min, boundsTickTargetCount, useNearMaximumHeadroom, nearMaximumHeadroomRatio, preferUnitOneOverTwo);
        if (max > min) return reflect ? new ChartValueExtents(-max, -min) : new ChartValueExtents(min, max);
        return fallback;
    }

    private static ChartValueExtents ReadPercentStackedAwareValueAxisExtents(PptxSceneChartAxis? axis, ChartValueExtents fallback, bool percentStacked, bool useNearMaximumHeadroom, double nearMaximumHeadroomRatio)
    {
        ChartValueExtents extents = ReadSceneChartValueAxisExtents(axis, fallback, useNearMaximumHeadroom, nearMaximumHeadroomRatio);
        if (!percentStacked)
        {
            return extents;
        }

        double min = axis?.Minimum is not null ? extents.Min : 0d;
        double max = axis?.Maximum is not null ? extents.Max : 1d;
        return max > min
            ? new ChartValueExtents(min, max)
            : extents;
    }

    private static ChartAxisUnits ReadSceneChartValueAxisUnits(PptxSceneChartAxis? axis)
    {
        return axis is null
            ? ChartAxisUnits.Empty
            : new ChartAxisUnits(axis.MajorUnit, axis.MinorUnit);
    }

    private static ChartAxisUnits ResolvePercentStackedAxisUnits(ChartAxisUnits axisUnits, bool percentStacked)
    {
        return percentStacked && axisUnits.MajorUnit is null
            ? axisUnits with { MajorUnit = 0.1d }
            : axisUnits;
    }

    private static ChartAxisUnits ResolveBubbleAxisUnits(ChartAxisUnits axisUnits, ChartValueExtents extents, double tickTargetCount = PptxChartMetricRules.BubbleAxisNiceTickTargetCount)
    {
        double span = extents.Max - extents.Min;
        return axisUnits.MajorUnit is null
            ? axisUnits with { MajorUnit = ChooseChartAxisMajorUnit(span > 0d ? span : 1d, tickTargetCount) }
            : axisUnits;
    }

    private static ChartAxisTickPolicy ReadCartesianChartAxisTickPolicy(PptxSceneChart? chart, PptxSceneChartAxis? axis, ChartValueExtents extents, bool percentStacked = false)
    {
        // Bounds were resolved once before measurement/emission. Recomputing
        // ceiling arithmetic here could shift a fractional origin a second time.
        if (!CanUseCartesianChartSourceIntervals(chart, axis)) return ChartAxisTickPolicy.ZeroAlignedWithEndpoints;
        return UsesAuthoredChartMinimumIntervals(axis?.Minimum, axis?.HasLogarithmicScale ?? false, extents) ||
            (axis?.Minimum is null && !percentStacked && CanOwnCartesianChartSourceRange(chart, axis) &&
                axis?.MajorUnit is { } unit && double.IsFinite(unit) && unit > 0d)
            ? ChartAxisTickPolicy.SourceMinimumIntervals
            : ChartAxisTickPolicy.ZeroAlignedWithEndpoints;
    }

    private static bool UsesAuthoredChartMinimumIntervals(double? minimum, bool logarithmic, ChartValueExtents extents) =>
        !logarithmic && minimum is { } value && double.IsFinite(value) && value == extents.Min;

    private static ChartValueExtents ReadCartesianChartValueAxisExtents(PptxSceneChart? chart, PptxSceneChartAxis? axis, ChartValueExtents fallback, bool percentStacked, bool useNearMaximumHeadroom, double nearMaximumHeadroomRatio)
    {
        ChartValueExtents extents = ReadPercentStackedAwareValueAxisExtents(axis, fallback, percentStacked,
            useNearMaximumHeadroom, nearMaximumHeadroomRatio);
        return !percentStacked && TryResolveCartesianChartSourceRange(chart, axis, fallback, extents, out ChartValueExtents owned)
            ? owned : extents;
    }

    private static bool TryResolveCartesianChartSourceRange(PptxSceneChart? chart, PptxSceneChartAxis? axis,
        ChartValueExtents data, ChartValueExtents extents, out ChartValueExtents owned)
    {
        owned = extents;
        if (!CanOwnCartesianChartSourceRange(chart, axis)) return false;
        if (axis!.MajorUnit is not { } unit || !double.IsFinite(unit) || unit <= 0d)
        {
            return axis.Minimum is { } minimum && double.IsFinite(minimum) && minimum == extents.Min;
        }

        // The authored major interval owns automatic limits too. Either fixed
        // bound anchors the lattice; otherwise zero does. Measure headroom in
        // that range, so translating an authored minimum and its data together
        // cannot change the interval count. Manual placement changes no scale.
        double origin = axis.Minimum ?? axis.Maximum ?? 0d;
        double lower = axis.Minimum ?? origin + Math.Floor((data.Min - origin) / unit) * unit;
        double upper = axis.Maximum ?? origin + Math.Ceiling((data.Max - origin) / unit) * unit;
        double span = upper - lower;
        // A nonzero fixed maximum owns the downward lattice without extra
        // headroom. Zero remains the baseline and retains lower headroom.
        // Source intervals use one shared padding decision across plot kinds;
        // automatic-unit scales retain their existing rounding policy above.
        bool extendLower = axis.Minimum is null && (axis.Maximum is null or 0d) && lower != 0d &&
            (lower >= data.Min || span > 0d && (upper - data.Min) / span >= PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
        bool extendUpper = axis.Maximum is null && (upper != 0d || lower == 0d) &&
            (upper <= data.Max || span > 0d && (data.Max - lower) / span >= PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
        if (extendLower) lower -= unit;
        if (extendUpper) upper += unit;
        if (!double.IsFinite(lower) || !double.IsFinite(upper) || upper <= lower)
        {
            return false;
        }
        owned = new ChartValueExtents(lower, upper);
        return true;
    }

    private static bool CanOwnCartesianChartSourceRange(PptxSceneChart? chart, PptxSceneChartAxis? axis) =>
        CanUseCartesianChartSourceIntervals(chart, axis) &&
        // Stacked horizontal intervals retain their existing bounds fallback.
        chart!.Plots[0] is not { PlotKind: PptxSceneChartPlotKind.Bar, BarDirectionKind: PptxSceneChartBarDirection.Bar,
            GroupingKind: not PptxSceneChartGrouping.Clustered };

    // A shared reciprocal source pair owns one linear scale across the supported
    // independent plots. Other compositions retain their existing fallback.
    private static bool HasSharedCartesianChartValueAxis(PptxSceneChart? chart, PptxSceneChartAxis? axis)
    {
        if (chart is null || chart.Plots.Count < 2 || axis is not { HasScaling: true, HasLogarithmicScale: false } ||
            chart.Axes.Count(candidate => candidate.AxisKind == PptxSceneChartAxisKind.Value) != 1 ||
            ReadScenePairedCategoryAxis(chart, axis) is not { } category) return false;
        PptxSceneChartBarDirection? direction = null;
        bool hasLines = false;
        foreach (var plot in chart.Plots)
        {
            if (plot.AxisIds.Count != 2 || !plot.AxisIds.Contains(axis.Id, StringComparer.Ordinal) ||
                !plot.AxisIds.Contains(category.Id, StringComparer.Ordinal)) return false;
            if (plot is { PlotKind: PptxSceneChartPlotKind.Bar,
                BarDirectionKind: PptxSceneChartBarDirection.Column or PptxSceneChartBarDirection.Bar,
                GroupingKind: PptxSceneChartGrouping.Clustered })
            {
                if (direction is not null && direction != plot.BarDirectionKind) return false;
                direction = plot.BarDirectionKind;
            }
            else if (plot is { PlotKind: PptxSceneChartPlotKind.Line, GroupingKind: PptxSceneChartGrouping.Standard }) hasLines = true;
            else return false;
        }
        return direction == PptxSceneChartBarDirection.Column ||
            direction == PptxSceneChartBarDirection.Bar && !hasLines;
    }

    private static bool CanUseCartesianChartSourceIntervals(PptxSceneChart? chart, PptxSceneChartAxis? axis) =>
        HasSharedCartesianChartValueAxis(chart, axis) || chart is not null && chart.Plots.Count == 1 &&
        (chart.Plots[0].PlotKind is PptxSceneChartPlotKind.Line or PptxSceneChartPlotKind.Area ||
            chart.Plots[0] is { PlotKind: PptxSceneChartPlotKind.Bar,
                BarDirectionKind: PptxSceneChartBarDirection.Column or PptxSceneChartBarDirection.Bar }) &&
        chart.Axes.Count(a => a.AxisKind == PptxSceneChartAxisKind.Value) == 1 &&
        axis is { HasScaling: true, HasLogarithmicScale: false };

    private sealed class ChartTextMeasurer
    {
        private readonly TextAdvanceEstimator estimator;
        private readonly bool kerningEnabled;

        // Office applies no pair kerning to chart axis tick and category labels
        // (adjustment sums read 0 across bar probes at 7 to 18pt); data labels,
        // legends, and titles keep the legacy behavior until role evidence lands (F03).
        public ChartTextMeasurer(PresentationFontResolver? fontResolver)
            : this(fontResolver, kerningEnabled: true)
        {
        }

        public ChartTextMeasurer(PresentationFontResolver? fontResolver, bool kerningEnabled)
        {
            estimator = new TextAdvanceEstimator(fontResolver, fontResolver?.CancellationToken ?? default);
            this.kerningEnabled = kerningEnabled;
        }

        public ValueTask PrepareAsync(string text, ChartTextStyle style) =>
            estimator.PrepareAsync(style.FontFamily, style.Bold, style.Italic, text);

        public async ValueTask PrepareAsync(IEnumerable<string> texts, ChartTextStyle style)
        {
            foreach (string text in texts)
                await PrepareAsync(text, style).ConfigureAwait(false);
        }

        public ValueTask<double> MeasureAsync(string text, ChartTextStyle style) =>
            MeasureAsync(text, style.FontSize, style.FontFamily, style.Bold, style.Italic, style.CharacterSpacing);

        public async ValueTask<double> MeasureAsync(string text, double fontSize, string? fontFamily, bool bold, bool italic, double characterSpacing)
        {
            await estimator.PrepareAsync(fontFamily, bold, italic, text).ConfigureAwait(false);
            return Measure(text, fontSize, fontFamily, bold, italic, characterSpacing);
        }

        public async ValueTask<double> MeasureNormalLineHeightAsync(ChartTextStyle style)
        {
            return await TryMeasureNormalLineHeightAsync(style).ConfigureAwait(false)
                ?? style.FontSize * PptxTextMetricRules.CssNormalLineHeightFallback;
        }

        public async ValueTask<double?> TryMeasureNormalLineHeightAsync(ChartTextStyle style)
        {
            await PrepareAsync(string.Empty, style).ConfigureAwait(false);
            return TryMeasureNormalLineHeight(style, out double height) ? height : null;
        }

        public async ValueTask<double?> TryMeasureWindowsLineHeightAsync(ChartTextStyle style)
        {
            await PrepareAsync(string.Empty, style).ConfigureAwait(false);
            var font = estimator.ResolveOpenTypeFont(style.FontFamily, style.Bold, style.Italic);
            if (font is null || font.UnitsPerEm == 0) return null;
            double ratio = ((double)font.Os2.WindowsAscender + font.Os2.WindowsDescender) / font.UnitsPerEm;
            return ratio > PptxTextMetricRules.MinimumFontLineBoxMetricRatio &&
                ratio <= PptxTextMetricRules.MaximumFontLineBoxMetricRatio ? style.FontSize * ratio : null;
        }

        public bool TryMeasureNormalLineHeight(ChartTextStyle style, out double height) =>
            TryMeasureLineBox(style, out _, out _, out height);

        public async ValueTask<(double Ascent, double Descent, double Height)> MeasureLineBoxAsync(ChartTextStyle style, bool horizontalHeader = false)
        {
            await PrepareAsync(string.Empty, style).ConfigureAwait(false);
            if (TryMeasureLineBox(style, out double ascent, out double descent, out double height, horizontalHeader))
                return (ascent, descent, height);
            height = style.FontSize * PptxTextMetricRules.CssNormalLineHeightFallback;
            return (height * 0.8d, height * 0.2d, height);
        }

        public async ValueTask<(double Ascent, double Height, bool HasMetrics)> MeasureTextLineBoxAsync(
            string text, ChartTextStyle style, bool horizontalHeader = true)
        {
            // Use the same embeddable per-glyph faces as advance measurement and
            // painting, including substitution of unsupported outline formats.
            double ascent = 0d, height = 0d;
            foreach (Rune rune in text.EnumerateRunes())
            {
                ResolvedGlyphFont? resolved = await estimator.ResolveGlyphFontAsync(
                    style.FontFamily, style.Bold, style.Italic, rune.Value).ConfigureAwait(false);
                if (resolved is not null && TryMeasureLineBox(resolved.Font, style,
                    out double glyphAscent, out _, out double glyphHeight, horizontalHeader))
                {
                    ascent = Math.Max(ascent, glyphAscent);
                    height = Math.Max(height, glyphHeight);
                }
            }
            if (height > 0d) return (ascent, height, true);
            var line = await MeasureLineBoxAsync(style, horizontalHeader).ConfigureAwait(false);
            return (line.Ascent, line.Height, false);
        }

        private bool TryMeasureLineBox(ChartTextStyle style, out double ascent, out double descent, out double height, bool horizontalHeader = false)
            => TryMeasureLineBox(estimator.ResolveOpenTypeFont(style.FontFamily, style.Bold, style.Italic),
                style, out ascent, out descent, out height, horizontalHeader);

        public async ValueTask<(double Bottom, double Top)?> TryMeasureInkBoundsAsync(string text, ChartTextStyle style)
        {
            double bottom = double.PositiveInfinity, top = double.NegativeInfinity;
            foreach (Rune rune in text.EnumerateRunes())
            {
                ResolvedGlyphFont? resolved = await estimator.ResolveGlyphFontAsync(
                    style.FontFamily, style.Bold, style.Italic, rune.Value).ConfigureAwait(false);
                if (resolved is null || !resolved.Font.TryGetGlyphBounds(resolved.Font.MapCodePoint(rune.Value), out var bounds))
                    return null;
                double scale = style.FontSize / resolved.Font.UnitsPerEm;
                bottom = Math.Min(bottom, bounds.YMin * scale);
                top = Math.Max(top, bounds.YMax * scale);
            }
            return double.IsFinite(bottom) && double.IsFinite(top) ? (bottom, top) : null;
        }

        private static bool TryMeasureLineBox(OpenTypeFont? font, ChartTextStyle style,
            out double ascent, out double descent, out double height, bool horizontalHeader)
        {
            ascent = descent = height = 0d;
            if (font is null || font.UnitsPerEm == 0)
            {
                return false;
            }

            // Native title controls use the horizontal-header line, including
            // its gap. Other chart text retains its existing selected OS/2 box.
            bool typographic = (font.Os2.SelectionFlags & 0x80) != 0;
            double ascentUnits = horizontalHeader
                ? font.Hhea.HorizontalAscender + font.Hhea.HorizontalLineGap
                : typographic ? font.Os2.TypographicAscender : font.Os2.WindowsAscender;
            double descentUnits = horizontalHeader ? -(double)font.Hhea.HorizontalDescender
                : typographic ? -(double)font.Os2.TypographicDescender : font.Os2.WindowsDescender;
            double units = horizontalHeader ? ascentUnits + descentUnits
                : typographic ? ascentUnits + descentUnits + font.Os2.TypographicLineGap : ascentUnits + descentUnits;
            double ratio = units / font.UnitsPerEm;
            if (ratio <= PptxTextMetricRules.MinimumFontLineBoxMetricRatio || ratio > PptxTextMetricRules.MaximumFontLineBoxMetricRatio)
            {
                return false;
            }

            height = style.FontSize * ratio;
            ascent = style.FontSize * Math.Max(0d, ascentUnits) / font.UnitsPerEm;
            descent = style.FontSize * Math.Max(0d, descentUnits) / font.UnitsPerEm;
            return true;
        }

        public double Measure(string text, ChartTextStyle style)
        {
            return Measure(text, style.FontSize, style.FontFamily, style.Bold, style.Italic, style.CharacterSpacing);
        }

        public double Measure(string text, double fontSize, string? fontFamily, bool bold, bool italic, double characterSpacing)
        {
            return estimator.Measure(
                text,
                fontSize,
                fontFamily,
                bold,
                italic,
                characterSpacing,
                kerningEnabled: kerningEnabled);
        }
    }

    private static int GetValueAxisSideSlot(PptxSceneChart? sceneChart, PptxSceneChartAxis? primarySceneAxis, PptxSceneChartAxis? secondarySceneAxis, bool defaultPrimaryRightSide, bool defaultSecondaryRightSide)
    {
        bool primaryRight = ResolveSceneValueAxisLabelsRightSide(sceneChart, primarySceneAxis, defaultPrimaryRightSide);
        bool secondaryRight = ResolveSceneValueAxisLabelsRightSide(sceneChart, secondarySceneAxis, defaultSecondaryRightSide);
        return primaryRight == secondaryRight ? 1 : 0;
    }

    private static bool IsSameChartAxis(ChartAxisSource first, ChartAxisSource second)
    {
        string? firstId = first.SceneAxis?.Id;
        string? secondId = second.SceneAxis?.Id;
        return !string.IsNullOrWhiteSpace(firstId) && string.Equals(firstId, secondId, StringComparison.Ordinal);
    }

    private static bool IsSceneOrXmlVisibleValueAxis(PptxSceneChartAxis? sceneAxis, XElement? axis)
    {
        if (sceneAxis is not null)
        {
            return sceneAxis.AxisKind == PptxSceneChartAxisKind.Value &&
                sceneAxis.IsDeleted != true;
        }

        return axis is not null &&
            !PptxSceneBuilder.IsOoxmlBooleanElementEnabled(axis.Element(ChartNamespace + "delete"));
    }

    private static string? ReadChartAxisId(XElement? axis)
    {
        return PptxSceneBuilder.ReadOptionalChartValueAttribute(axis?.Element(ChartNamespace + "axId"));
    }

    private static double? ReadSceneValueAxisCrossingValue(PptxSceneChartAxis? sceneAxis, ChartValueExtents extents)
    {
        if (sceneAxis?.CrossesAt is { } explicitCrossing)
        {
            return explicitCrossing;
        }

        PptxSceneChartAxisCrosses crosses = sceneAxis?.CrossesKind ?? PptxSceneChartAxisCrosses.AutoZero;
        if (crosses == PptxSceneChartAxisCrosses.Maximum)
        {
            return extents.Max;
        }

        if (crosses == PptxSceneChartAxisCrosses.Minimum)
        {
            return extents.Min;
        }

        if (extents.Min <= 0d && extents.Max >= 0d)
        {
            return 0d;
        }

        return extents.Min > 0d ? extents.Min : extents.Max;
    }

    private static bool ReadSceneValueAxisReversed(PptxSceneChartAxis? sceneAxis)
    {
        return sceneAxis?.OrientationKind == PptxSceneChartAxisOrientation.MaximumMinimum;
    }

    // Gallery style 2 axis-family default (S05): single-plot bar/column charts carrying
    // the built-in gallery style render unstyled axis-family strokes gray-0.537 at width
    // 1.0 (three cached Office refs); style-part charts keep their role-driven output and
    // multi-plot dual-axis charts keep legacy output for lack of axis-color evidence.
    // Modern Office applies the c14 Choice branch, not this fallback value, so a future
    // c14-capable normalization must revisit which branch drives the gray.
    private static ChartSeriesStroke? ResolveGalleryAxisFamilyDefault(string? galleryStyleId)
    {
        if (galleryStyleId != "2")
        {
            return null;
        }

        return new ChartSeriesStroke(new RgbColor(137, 137, 137), 1d, 1d);
    }

    private static ChartGridlineStyle ReadSceneChartGridlineStyle(PptxSceneChartAxis? sceneAxis)
    {
        return new ChartGridlineStyle(
            ToChartSeriesStroke(sceneAxis?.MajorGridlineLine ?? default, null) ??
                ToChartSeriesStroke(sceneAxis?.MajorGridlineStyleLine ?? default, null),
            ToChartSeriesStroke(sceneAxis?.MinorGridlineLine ?? default, null) ??
                ToChartSeriesStroke(sceneAxis?.MinorGridlineStyleLine ?? default, null));
    }

    private static ChartValueAxisRenderOptions ReadSceneChartValueAxisRenderOptions(
        PptxSceneChartAxis? sceneAxis,
        ChartValueExtents extents,
        bool percentStacked)
    {
        return new ChartValueAxisRenderOptions(
            ResolvePercentStackedAxisUnits(ReadSceneChartValueAxisUnits(sceneAxis), percentStacked),
            ReadSceneValueAxisReversed(sceneAxis),
            ReadSceneValueAxisCrossingValue(sceneAxis, extents),
            sceneAxis?.HasMajorGridlines ?? false,
            sceneAxis?.HasMinorGridlines ?? false,
            ReadSceneChartGridlineStyle(sceneAxis),
            sceneAxis?.HasMajorGridlineElement ?? false);
    }

    private static ChartValueAxisRenderOptions ReadSceneChartBubbleValueAxisOptions(
        PptxSceneChartAxis? sceneAxis,
        ChartValueExtents extents,
        double tickTargetCount = PptxChartMetricRules.BubbleAxisNiceTickTargetCount)
    {
        double? minimum = sceneAxis?.Minimum;
        // An authored minimum is the interval origin, independently of how the
        // other bound and unit were resolved. Rejected bounds keep the fallback.
        bool logarithmic = sceneAxis?.HasLogarithmicScale ?? false;
        ChartAxisTickPolicy tickPolicy = UsesAuthoredChartMinimumIntervals(minimum, logarithmic, extents)
            ? ChartAxisTickPolicy.SourceMinimumIntervals
            : ChartAxisTickPolicy.ZeroAlignedWithEndpoints;
        ChartValueAxisRenderOptions options = ReadSceneChartValueAxisRenderOptions(sceneAxis, extents, percentStacked: false);
        return options with
        {
            Units = ResolveBubbleAxisUnits(options.Units, extents, tickTargetCount),
            TickPolicy = tickPolicy
        };
    }

    private static ChartAxesStyle ReadSceneChartAxesStyle(PptxSceneChart sceneChart, PptxSceneChartPlot? plot)
    {
        ChartAxisSource valueAxisSource = ReadChartValueAxesForPlot(sceneChart, plot).FirstOrDefault();
        PptxSceneChartAxis? valueAxis = valueAxisSource.SceneAxis;
        PptxSceneChartAxis? categoryAxis = ReadChartCategoryAxisForPlot(sceneChart, plot).SceneAxis;
        PptxSceneChartAxis? secondaryValueAxis = ReadSceneSecondaryValueAxisForChart(sceneChart, valueAxisSource).SceneAxis;
        return new ChartAxesStyle(
            ToChartSeriesStroke(valueAxis?.Line ?? default, null),
            ToChartSeriesStroke(secondaryValueAxis?.Line ?? default, null),
            ToChartSeriesStroke(categoryAxis?.Line ?? default, null),
            ResolveSceneValueAxisRightSide(sceneChart, valueAxis, defaultRightSide: false),
            ResolveSceneValueAxisRightSide(sceneChart, secondaryValueAxis, defaultRightSide: true),
            ResolveSceneValueAxisBottomSide(valueAxis, defaultBottomSide: true),
            ResolveSceneCategoryAxisRightSide(categoryAxis, defaultRightSide: false),
            valueAxis?.IsDeleted != true,
            categoryAxis?.IsDeleted != true,
            ReadSceneChartAxisMajorTickMark(categoryAxis),
            ResolveSceneCategoryAxisTopSide(categoryAxis, defaultTopSide: false));
    }

    private static PptxSceneChartAxisTickMark ReadSceneChartAxisMajorTickMark(PptxSceneChartAxis? sceneAxis)
    {
        PptxSceneChartAxisTickMark tickMark = sceneAxis?.MajorTickMarkKind ?? PptxSceneChartAxisTickMark.Unknown;
        return tickMark == PptxSceneChartAxisTickMark.Unknown ? PptxSceneChartAxisTickMark.None : tickMark;
    }

    private static PptxSceneChartTickLabelPosition ResolveChartTickLabelPosition(PptxSceneChartTickLabelPosition position)
    {
        return position == PptxSceneChartTickLabelPosition.Unknown
            ? PptxSceneChartTickLabelPosition.NextTo
            : position;
    }

    private static bool ResolveSceneValueAxisRightSide(PptxSceneChart? chart, PptxSceneChartAxis? sceneAxis, bool defaultRightSide)
    {
        bool declaredSide = sceneAxis?.PositionKind switch
        {
            PptxSceneChartAxisPosition.Right => true,
            PptxSceneChartAxisPosition.Left => false,
            _ => defaultRightSide
        };
        // A paired ordinal axis can resolve a crossing to a physical edge.
        // Unpaired, date/numeric and interior crossings retain their declared fallback.
        PptxSceneChartAxis? category = ReadScenePairedCategoryAxis(chart, sceneAxis);
        if (category is null || sceneAxis?.CrossesAt is not null) return declaredSide;
        PptxSceneChartAxisCrosses crosses = string.IsNullOrEmpty(sceneAxis?.Crosses)
            ? PptxSceneChartAxisCrosses.AutoZero
            : sceneAxis?.CrossesKind ?? PptxSceneChartAxisCrosses.Unknown;
        bool reversed = ReadSceneValueAxisReversed(category);
        return crosses switch
        {
            PptxSceneChartAxisCrosses.Maximum => !reversed,
            PptxSceneChartAxisCrosses.Minimum or PptxSceneChartAxisCrosses.AutoZero => reversed,
            _ => declaredSide
        };
    }

    private static PptxSceneChartAxis? ReadScenePairedCategoryAxis(PptxSceneChart? chart, PptxSceneChartAxis? sceneAxis)
    {
        string? id = sceneAxis?.Id;
        string? crossId = sceneAxis?.CrossAxisId;
        if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(crossId)) return null;
        return chart?.Axes.FirstOrDefault(candidate => candidate.AxisKind == PptxSceneChartAxisKind.Category &&
            candidate.Id == crossId && candidate.CrossAxisId == id);
    }

    private static bool ResolveSceneValueAxisBottomSide(PptxSceneChartAxis? sceneAxis, bool defaultBottomSide)
    {
        return sceneAxis?.PositionKind switch
        {
            PptxSceneChartAxisPosition.Bottom => true,
            PptxSceneChartAxisPosition.Top => false,
            _ => defaultBottomSide
        };
    }

    private static bool ResolveSceneCategoryAxisRightSide(PptxSceneChartAxis? sceneAxis, bool defaultRightSide)
    {
        return sceneAxis?.PositionKind switch
        {
            PptxSceneChartAxisPosition.Right => true,
            PptxSceneChartAxisPosition.Left => false,
            _ => defaultRightSide
        };
    }

    private static bool ResolveSceneCategoryAxisTopSide(PptxSceneChartAxis? sceneAxis, bool defaultTopSide)
    {
        return sceneAxis?.PositionKind switch
        {
            PptxSceneChartAxisPosition.Top => true,
            PptxSceneChartAxisPosition.Bottom => false,
            _ => defaultTopSide
        };
    }

    private static bool IsSceneChartAxisLabelVisible(PptxSceneChartAxis? sceneAxis)
    {
        return sceneAxis?.IsDeleted != true &&
            ResolveChartTickLabelPosition(sceneAxis?.TickLabelPositionKind ?? PptxSceneChartTickLabelPosition.Unknown) !=
                PptxSceneChartTickLabelPosition.None;
    }

    private static bool ResolveSceneValueAxisLabelsRightSide(PptxSceneChart? chart, PptxSceneChartAxis? sceneAxis, bool defaultRightSide)
    {
        PptxSceneChartTickLabelPosition position = ResolveChartTickLabelPosition(
            sceneAxis?.TickLabelPositionKind ?? PptxSceneChartTickLabelPosition.Unknown);
        if (position is not (PptxSceneChartTickLabelPosition.Low or PptxSceneChartTickLabelPosition.High))
            return defaultRightSide;
        // Explicit label sides follow the paired ordinal direction. Missing,
        // date/numeric and unpaired axes retain their existing physical fallback.
        bool reversed = ReadSceneValueAxisReversed(ReadScenePairedCategoryAxis(chart, sceneAxis));
        return (position == PptxSceneChartTickLabelPosition.High) != reversed;
    }

    private static double ResolveSceneCategoryAxisLabelOffsetScale(PptxSceneChartAxis? sceneAxis)
    {
        int offset = sceneAxis?.LabelOffset ?? PptxChartMetricRules.CategoryAxisDefaultLabelOffset;
        offset = Math.Clamp(
            offset,
            PptxChartMetricRules.CategoryAxisMinimumLabelOffset,
            PptxChartMetricRules.CategoryAxisMaximumLabelOffset);
        return offset / (double)PptxChartMetricRules.CategoryAxisDefaultLabelOffset;
    }

    private static int ResolveSceneCategoryAxisTickLabelSkip(PptxSceneChartAxis? sceneAxis)
    {
        return Math.Max(PptxChartMetricRules.CategoryAxisDefaultTickLabelSkip,
            sceneAxis?.TickLabelSkip ?? PptxChartMetricRules.CategoryAxisDefaultTickLabelSkip);
    }

    private static bool ResolveSceneCategoryAxisLabelsOnTickMarks(PptxSceneChartAxis? sceneAxis)
    {
        return sceneAxis?.CrossBetweenKind == PptxSceneChartAxisCrossBetween.MidpointCategory;
    }
}
