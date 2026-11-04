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
    private static IReadOnlyList<XElement> ReadChartValueAxesForChart(XDocument chartXml, XElement chartElement)
    {
        string[] axisIds = chartElement
            .Elements(ChartNamespace + "axId")
            .Select(PptxSceneBuilder.ReadChartValueAttribute)
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .ToArray();
        if (axisIds.Length == 0)
        {
            return chartXml.Descendants(ChartNamespace + "valAx").ToArray();
        }

        var axes = new List<XElement>();
        foreach (string axisId in axisIds)
        {
            XElement? axis = chartXml
                .Descendants(ChartNamespace + "valAx")
                .FirstOrDefault(candidate => string.Equals(ReadChartAxisId(candidate), axisId, StringComparison.Ordinal));
            if (axis is not null)
            {
                axes.Add(axis);
            }
        }

        return axes;
    }

    private static IReadOnlyList<XElement> ReadChartCategoryAxesForChart(XDocument chartXml, XElement chartElement)
    {
        string[] axisIds = chartElement
            .Elements(ChartNamespace + "axId")
            .Select(PptxSceneBuilder.ReadChartValueAttribute)
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .ToArray();
        if (axisIds.Length == 0)
        {
            return ReadChartCategoryAxes(chartXml).ToArray();
        }

        var axes = new List<XElement>();
        foreach (string axisId in axisIds)
        {
            XElement? axis = ReadChartCategoryAxes(chartXml)
                .FirstOrDefault(candidate => string.Equals(ReadChartAxisId(candidate), axisId, StringComparison.Ordinal));
            if (axis is not null)
            {
                axes.Add(axis);
            }
        }

        return axes.Count == 0
            ? ReadChartCategoryAxes(chartXml).Take(1).ToArray()
            : axes;
    }

    private static IEnumerable<XElement> ReadChartCategoryAxes(XDocument chartXml)
    {
        return chartXml
            .Descendants(ChartNamespace + "plotArea")
            .FirstOrDefault()?
            .Elements()
            .Where(element => element.Name == ChartNamespace + "catAx" || element.Name == ChartNamespace + "dateAx") ??
            [];
    }

    private static ChartValueExtents ReadSceneOrXmlBubbleChartValueAxisExtents(PptxSceneChartAxis? axis, XElement? valueAxis, ChartValueExtents dataExtents, double boundsTickTargetCount = PptxChartMetricRules.BubbleAxisBoundsTickTargetCount, bool preferUnitOneOverTwo = true)
    {
        // Measurement keeps the actual extrema. The existing point-chart
        // range fallback owns zero inclusion and its minimum unit-sized span.
        ChartValueExtents fallback = new(Math.Min(0d, dataExtents.Min), Math.Max(1d, dataExtents.Max));
        return ReadSceneOrXmlChartValueAxisExtents(axis, valueAxis, fallback, useNearMaximumHeadroom: true,
            nearMaximumHeadroomRatio: PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio,
            boundsTickTargetCount: boundsTickTargetCount, preferUnitOneOverTwo: preferUnitOneOverTwo);
    }

    private static (ChartValueExtents Extents, ChartValueAxisRenderOptions Options) ReadChartPointValueAxisScale(
        PptxTheme theme, ChartAxisSource source, ChartValueExtents data, ChartValueExtents occupied, bool bubble, bool horizontal,
        double? automaticTickTargetCount = null)
    {
        bool scatterX = !bubble && horizontal;
        double target = scatterX ? PptxChartMetricRules.ScatterXAxisNiceTickTargetCount : PptxChartMetricRules.BubbleAxisNiceTickTargetCount;
        ChartValueExtents extents = ReadSceneOrXmlBubbleChartValueAxisExtents(source.SceneAxis, source.XmlAxis, data,
            scatterX ? PptxChartMetricRules.ScatterXAxisNiceTickTargetCount : PptxChartMetricRules.BubbleAxisBoundsTickTargetCount, !scatterX);
        PptxSceneChartAxis? axis = source.SceneAxis;
        XElement? xml = source.XmlAxis;
        ChartPointAxisDeclaration? declaration = ReadChartPointAxisDeclaration(source);
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
        ChartValueAxisRenderOptions options = ReadSceneOrXmlChartBubbleValueAxisOptions(axis, xml, theme, extents, target);
        if (resolvedIntervals) options = options with { Units = options.Units with { MajorUnit = unit }, TickPolicy = ChartAxisTickPolicy.SourceMinimumIntervals };
        return (extents, options);
    }

    private readonly record struct ChartPointAxisDeclaration(double? Minimum, double? Maximum, double? MajorUnit);

    private static ChartPointAxisDeclaration? ReadChartPointAxisDeclaration(ChartAxisSource source)
    {
        PptxSceneChartAxis? axis = source.SceneAxis;
        XElement? xml = source.XmlAxis;
        if (!(axis?.HasScaling ?? xml?.Element(ChartNamespace + "scaling") is not null) ||
            (axis?.HasLogarithmicScale ?? xml?.Element(ChartNamespace + "scaling")?.Element(ChartNamespace + "logBase") is not null)) return null;
        var minimum = axis is not null ? (axis.Minimum, axis.MinimumValue) :
            PptxSceneBuilder.ReadChartAxisScalingValueWithValue(xml!, "min");
        var maximum = axis is not null ? (axis.Maximum, axis.MaximumValue) :
            PptxSceneBuilder.ReadChartAxisScalingValueWithValue(xml!, "max");
        var major = axis is not null ? (axis.MajorUnit, axis.MajorUnitValue) :
            PptxSceneBuilder.ReadChartAxisUnitValueWithValue(xml!, "majorUnit");
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

    private static double? ReadScatterSourceIntervalTarget(PptxTheme theme, PptxSceneChart? chart, XDocument xml,
        ChartFrameBox frame, ChartAxisSource xAxis)
    {
        PptxSceneChartManualLayout source = chart?.PlotAreaLayout ?? PptxSceneBuilder.ReadChartPlotAreaManualLayout(xml);
        // Resolve the same authored inner width used for placement before caption
        // admission. Outer, automatic and unresolved modes retain their fallback.
        if (source.LayoutTargetKind != PptxSceneChartManualLayoutTarget.Inner ||
            !TryBuildManualPlotLayout(source, frame, GetDefaultChartPlotBox(frame), out ChartPlotLayout layout) ||
            !TryResolveCartesianPlotWidth(source, frame, layout.PlotBox, out var plot)) return null;
        ChartTextStyle style = ReadSceneOrXmlChartTextStyle(theme, chart, xAxis.SceneAxis, xml, xAxis.XmlAxis,
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

    private static ChartValueExtents ReadSceneOrXmlChartValueAxisExtents(PptxSceneChartAxis? axis, XElement? valueAxis, ChartValueExtents fallback, bool useNearMaximumHeadroom, double nearMaximumHeadroomRatio, double boundsTickTargetCount = PptxChartMetricRules.AxisNiceTickTargetCount, bool preferUnitOneOverTwo = false)
    {
        double? minimum, maximum;
        if (axis is null)
        {
            if (valueAxis?.Element(ChartNamespace + "scaling") is null) return fallback;
            minimum = PptxSceneBuilder.ReadChartAxisScalingValueWithValue(valueAxis, "min").Value;
            maximum = PptxSceneBuilder.ReadChartAxisScalingValueWithValue(valueAxis, "max").Value;
        }
        else
        {
            if (!axis.HasScaling) return fallback;
            minimum = axis.Minimum;
            maximum = axis.Maximum;
        }

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

    private static ChartValueExtents ReadPercentStackedAwareValueAxisExtents(PptxSceneChartAxis? axis, XElement? valueAxis, ChartValueExtents fallback, bool percentStacked, bool useNearMaximumHeadroom, double nearMaximumHeadroomRatio)
    {
        ChartValueExtents extents = ReadSceneOrXmlChartValueAxisExtents(axis, valueAxis, fallback, useNearMaximumHeadroom, nearMaximumHeadroomRatio);
        if (!percentStacked)
        {
            return extents;
        }

        double min = HasSceneOrXmlAxisScalingValue(axis, valueAxis, ChartAxisScalingBound.Minimum) ? extents.Min : 0d;
        double max = HasSceneOrXmlAxisScalingValue(axis, valueAxis, ChartAxisScalingBound.Maximum) ? extents.Max : 1d;
        return max > min
            ? new ChartValueExtents(min, max)
            : extents;
    }

    private static bool HasSceneOrXmlAxisScalingValue(PptxSceneChartAxis? axis, XElement? valueAxis, ChartAxisScalingBound bound)
    {
        if (axis is not null)
        {
            return bound == ChartAxisScalingBound.Minimum
                ? axis.Minimum is not null
                : axis.Maximum is not null;
        }

        string elementName = bound == ChartAxisScalingBound.Minimum ? "min" : "max";
        return valueAxis?
            .Element(ChartNamespace + "scaling")
            ?.Element(ChartNamespace + elementName) is not null;
    }

    private static ChartAxisUnits ReadChartValueAxisUnits(XElement? valueAxis)
    {
        if (valueAxis is null)
        {
            return ChartAxisUnits.Empty;
        }

        return new ChartAxisUnits(
            PptxSceneBuilder.ReadChartAxisUnitValueWithValue(valueAxis, "majorUnit").Value,
            PptxSceneBuilder.ReadChartAxisUnitValueWithValue(valueAxis, "minorUnit").Value);
    }

    private static ChartAxisUnits ReadSceneOrXmlChartValueAxisUnits(PptxSceneChartAxis? axis, XElement? valueAxis)
    {
        return axis is null
            ? ReadChartValueAxisUnits(valueAxis)
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

    private static ChartValueExtents ReadCartesianChartValueAxisExtents(PptxSceneChart? chart, PptxSceneChartAxis? axis,
        XElement? valueAxis, ChartValueExtents fallback, bool percentStacked, bool useNearMaximumHeadroom, double nearMaximumHeadroomRatio)
    {
        ChartValueExtents extents = ReadPercentStackedAwareValueAxisExtents(axis, valueAxis, fallback, percentStacked,
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
        // A fixed maximum owns the downward lattice without extra headroom.
        // Source intervals use one shared padding decision across plot kinds;
        // automatic-unit scales retain their existing rounding policy above.
        bool extendLower = axis.Minimum is null && axis.Maximum is null && lower != 0d &&
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
        chart!.Plots[0] is not { PlotKind: PptxSceneChartPlotKind.Bar, BarDirectionKind: PptxSceneChartBarDirection.Bar };

    // A standalone plot shares one resolved linear scale. Composed/secondary
    // scale ownership retains its existing fallback, independently of direction.
    private static bool CanUseCartesianChartSourceIntervals(PptxSceneChart? chart, PptxSceneChartAxis? axis) =>
        chart is not null && chart.Plots.Count == 1 &&
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

        public async ValueTask<(double Ascent, double Height)> MeasureTextLineBoxAsync(string text, ChartTextStyle style)
        {
            // Use the same embeddable per-glyph faces as advance measurement and
            // painting, including substitution of unsupported outline formats.
            double ascent = 0d, height = 0d;
            foreach (Rune rune in text.EnumerateRunes())
            {
                ResolvedGlyphFont? resolved = await estimator.ResolveGlyphFontAsync(
                    style.FontFamily, style.Bold, style.Italic, rune.Value).ConfigureAwait(false);
                if (resolved is not null && TryMeasureLineBox(resolved.Font, style,
                    out double glyphAscent, out _, out double glyphHeight, horizontalHeader: true))
                {
                    ascent = Math.Max(ascent, glyphAscent);
                    height = Math.Max(height, glyphHeight);
                }
            }
            if (height > 0d) return (ascent, height);
            var line = await MeasureLineBoxAsync(style, horizontalHeader: true).ConfigureAwait(false);
            return (line.Ascent, line.Height);
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

    private static int GetValueAxisSideSlot(PptxSceneChart? sceneChart, PptxSceneChartAxis? primarySceneAxis, XElement? primaryAxis, PptxSceneChartAxis? secondarySceneAxis, XElement secondaryAxis, bool defaultPrimaryRightSide, bool defaultSecondaryRightSide)
    {
        bool primaryRight = ResolveSceneOrXmlValueAxisLabelsRightSide(sceneChart, primarySceneAxis, primaryAxis, defaultPrimaryRightSide);
        bool secondaryRight = ResolveSceneOrXmlValueAxisLabelsRightSide(sceneChart, secondarySceneAxis, secondaryAxis, defaultSecondaryRightSide);
        return primaryRight == secondaryRight ? 1 : 0;
    }

    private static bool IsSameChartAxis(ChartAxisSource first, ChartAxisSource second)
    {
        string? firstId = first.SceneAxis?.Id ?? ReadChartAxisId(first.XmlAxis);
        string? secondId = second.SceneAxis?.Id ?? ReadChartAxisId(second.XmlAxis);
        return !string.IsNullOrWhiteSpace(firstId) && string.Equals(firstId, secondId, StringComparison.Ordinal);
    }

    private static XElement? ReadSecondaryRightValueAxis(XDocument chartXml)
    {
        return chartXml
            .Descendants(ChartNamespace + "valAx")
            .Where(IsRightValueAxis)
            .FirstOrDefault();
    }

    private static XElement? ReadSecondaryValueAxisForChart(XDocument chartXml, XElement? primaryValueAxis)
    {
        string? primaryAxisId = ReadChartAxisId(primaryValueAxis);
        if (string.IsNullOrWhiteSpace(primaryAxisId))
        {
            return ReadSecondaryRightValueAxis(chartXml);
        }

        return chartXml
            .Descendants(ChartNamespace + "valAx")
            .FirstOrDefault(axis =>
                !string.Equals(ReadChartAxisId(axis), primaryAxisId, StringComparison.Ordinal) &&
                !IsChartAxisDeleted(axis));
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

    private static bool IsRightValueAxis(XElement? axis)
    {
        return axis is not null &&
            PptxSceneBuilder.ParseChartAxisPosition(PptxSceneBuilder.ReadChartElementValue(axis, "axPos")) == PptxSceneChartAxisPosition.Right &&
            !PptxSceneBuilder.IsOoxmlBooleanElementEnabled(axis.Element(ChartNamespace + "delete"));
    }

    private static string? ReadChartAxisId(XElement? axis)
    {
        return PptxSceneBuilder.ReadOptionalChartValueAttribute(axis?.Element(ChartNamespace + "axId"));
    }

    private static double? ReadSceneOrXmlValueAxisCrossingValue(PptxSceneChartAxis? sceneAxis, XElement? valueAxis, ChartValueExtents extents)
    {
        double? crossesAt = sceneAxis is not null
            ? sceneAxis.CrossesAt
            : valueAxis is null
                ? null
                : PptxSceneBuilder.ReadChartElementDoubleWithValue(valueAxis, "crossesAt").Value;
        if (crossesAt is { } explicitCrossing)
        {
            return explicitCrossing;
        }

        PptxSceneChartAxisCrosses crosses = sceneAxis is not null
            ? ResolveChartAxisCrosses(sceneAxis.CrossesKind)
            : ResolveChartAxisCrosses(PptxSceneBuilder.ParseChartAxisCrosses(PptxSceneBuilder.ReadChartElementValue(valueAxis, "crosses")));
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

        PptxSceneChartAxisCrosses ResolveChartAxisCrosses(PptxSceneChartAxisCrosses crosses)
        {
            return crosses == PptxSceneChartAxisCrosses.Unknown
                ? PptxSceneChartAxisCrosses.AutoZero
                : crosses;
        }
    }

    private static bool ReadSceneOrXmlValueAxisReversed(PptxSceneChartAxis? sceneAxis, XElement? valueAxis)
    {
        if (sceneAxis is not null)
        {
            return ResolveChartAxisReversed(sceneAxis.OrientationKind);
        }

        return ResolveChartAxisReversed(PptxSceneBuilder.ParseChartAxisOrientation(
            PptxSceneBuilder.ReadChartElementValue(valueAxis?.Element(ChartNamespace + "scaling"), "orientation")));

        bool ResolveChartAxisReversed(PptxSceneChartAxisOrientation orientation)
        {
            return orientation == PptxSceneChartAxisOrientation.MaximumMinimum;
        }
    }

    private static bool ReadSceneOrXmlMajorGridlines(PptxSceneChartAxis? sceneAxis, XElement? axis)
    {
        return sceneAxis?.HasMajorGridlines ?? HasMajorGridlines();

        bool HasMajorGridlines()
        {
            return PptxSceneBuilder.IsChartGridlineVisible(axis?.Element(ChartNamespace + "majorGridlines"));
        }
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

    private static bool ReadSceneOrXmlMinorGridlines(PptxSceneChartAxis? sceneAxis, XElement? axis)
    {
        return sceneAxis?.HasMinorGridlines ?? HasMinorGridlines();

        bool HasMinorGridlines()
        {
            return PptxSceneBuilder.IsChartGridlineVisible(axis?.Element(ChartNamespace + "minorGridlines"));
        }
    }

    private static ChartGridlineStyle ReadSceneOrXmlChartGridlineStyle(PptxSceneChartAxis? sceneAxis, XElement? xmlAxis, PptxTheme theme)
    {
        return new ChartGridlineStyle(
            ReadSceneOrXmlChartGridlineStroke(sceneAxis, sceneAxis?.MajorGridlineLine ?? default, sceneAxis?.MajorGridlineStyleLine ?? default, xmlAxis?.Element(ChartNamespace + "majorGridlines"), theme),
            ReadSceneOrXmlChartGridlineStroke(sceneAxis, sceneAxis?.MinorGridlineLine ?? default, sceneAxis?.MinorGridlineStyleLine ?? default, xmlAxis?.Element(ChartNamespace + "minorGridlines"), theme));
    }

    private static ChartValueAxisRenderOptions ReadSceneOrXmlChartValueAxisRenderOptions(
        PptxSceneChartAxis? sceneAxis,
        XElement? xmlAxis,
        PptxTheme theme,
        ChartValueExtents extents,
        bool percentStacked)
    {
        return new ChartValueAxisRenderOptions(
            ResolvePercentStackedAxisUnits(ReadSceneOrXmlChartValueAxisUnits(sceneAxis, xmlAxis), percentStacked),
            ReadSceneOrXmlValueAxisReversed(sceneAxis, xmlAxis),
            ReadSceneOrXmlValueAxisCrossingValue(sceneAxis, xmlAxis, extents),
            ReadSceneOrXmlMajorGridlines(sceneAxis, xmlAxis),
            ReadSceneOrXmlMinorGridlines(sceneAxis, xmlAxis),
            ReadSceneOrXmlChartGridlineStyle(sceneAxis, xmlAxis, theme),
            sceneAxis?.HasMajorGridlineElement ?? xmlAxis?.Element(ChartNamespace + "majorGridlines") is not null);
    }

    private static ChartValueAxisRenderOptions ReadSceneOrXmlChartBubbleValueAxisOptions(
        PptxSceneChartAxis? sceneAxis,
        XElement? xmlAxis,
        PptxTheme theme,
        ChartValueExtents extents,
        double tickTargetCount = PptxChartMetricRules.BubbleAxisNiceTickTargetCount)
    {
        double? minimum = sceneAxis is null
            ? (xmlAxis is null ? null : PptxSceneBuilder.ReadChartAxisScalingValueWithValue(xmlAxis, "min").Value)
            : sceneAxis.Minimum;
        // An authored minimum is the interval origin, independently of how the
        // other bound and unit were resolved. Rejected bounds keep the fallback.
        bool logarithmic = sceneAxis?.HasLogarithmicScale ?? xmlAxis?.Element(ChartNamespace + "scaling")?.Element(ChartNamespace + "logBase") is not null;
        ChartAxisTickPolicy tickPolicy = UsesAuthoredChartMinimumIntervals(minimum, logarithmic, extents)
            ? ChartAxisTickPolicy.SourceMinimumIntervals
            : ChartAxisTickPolicy.ZeroAlignedWithEndpoints;
        ChartValueAxisRenderOptions options = ReadSceneOrXmlChartValueAxisRenderOptions(sceneAxis, xmlAxis, theme, extents, percentStacked: false);
        return options with
        {
            Units = ResolveBubbleAxisUnits(options.Units, extents, tickTargetCount),
            TickPolicy = tickPolicy
        };
    }

    private static ChartSeriesStroke? ReadSceneOrXmlChartGridlineStroke(PptxSceneChartAxis? sceneAxis, PptxSceneLineStyle sceneLine, PptxSceneLineStyle sceneStyleLine, XElement? gridlines, PptxTheme theme)
    {
        if (sceneAxis is not null)
        {
            return ToChartSeriesStroke(sceneLine, null) ?? ToChartSeriesStroke(sceneStyleLine, null);
        }

        return ReadChartGridlineStroke(gridlines, theme);
    }

    private static ChartSeriesStroke? ReadChartGridlineStroke(XElement? gridlines, PptxTheme theme)
    {
        return ToChartSeriesStroke(PptxSceneBuilder.ReadChartGridlineLine(gridlines, theme), null);
    }

    private static ChartAxesStyle ReadSceneOrXmlChartAxesStyle(PptxSceneChart? sceneChart, PptxSceneChartPlot? plot, XDocument chartXml, PptxTheme theme, XElement chartElement)
    {
        ChartAxisSource valueAxisSource = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, plot, chartXml, chartElement).FirstOrDefault();
        XElement? valueAxisElement = ResolveXmlValueAxisForSource(sceneChart, valueAxisSource, chartXml);
        ChartAxisSource categoryAxisSource = ReadSceneOrXmlChartCategoryAxisForPlot(sceneChart, plot, chartXml, chartElement);
        XElement? categoryAxisElement = categoryAxisSource.XmlAxis;
        ChartAxisSource secondaryValueAxisSource = ReadSceneOrXmlSecondaryValueAxisForChart(sceneChart, chartXml, valueAxisSource);
        XElement? secondaryValueAxisElement = secondaryValueAxisSource.XmlAxis;
        PptxSceneChartAxis? valueAxis = valueAxisSource.SceneAxis;
        PptxSceneChartAxis? categoryAxis = categoryAxisSource.SceneAxis;
        PptxSceneChartAxis? secondaryValueAxis = secondaryValueAxisSource.SceneAxis;
        return new ChartAxesStyle(
            ReadSceneOrXmlChartAxisStroke(valueAxis, valueAxisElement, theme),
            ReadSceneOrXmlChartAxisStroke(secondaryValueAxis, secondaryValueAxisElement, theme),
            ReadSceneOrXmlChartAxisStroke(categoryAxis, categoryAxisElement, theme),
            ResolveSceneOrXmlValueAxisRightSide(sceneChart, valueAxis, valueAxisElement, defaultRightSide: false),
            ResolveSceneOrXmlValueAxisRightSide(sceneChart, secondaryValueAxis, secondaryValueAxisElement, defaultRightSide: true),
            ResolveSceneOrXmlValueAxisBottomSide(valueAxis, valueAxisElement, defaultBottomSide: true),
            ResolveSceneOrXmlCategoryAxisRightSide(categoryAxis, categoryAxisElement, defaultRightSide: false),
            valueAxis is null ? !IsChartAxisDeleted(valueAxisElement) : valueAxis.IsDeleted != true,
            categoryAxis is null ? !IsChartAxisDeleted(categoryAxisElement) : categoryAxis.IsDeleted != true,
            ReadSceneOrXmlChartAxisMajorTickMark(categoryAxis, categoryAxisElement),
            ResolveSceneOrXmlCategoryAxisTopSide(categoryAxis, categoryAxisElement, defaultTopSide: false));
    }

    private static ChartSeriesStroke? ReadSceneOrXmlChartAxisStroke(PptxSceneChartAxis? sceneAxis, XElement? xmlAxis, PptxTheme theme)
    {
        return sceneAxis is not null
            ? ToChartSeriesStroke(sceneAxis.Line, null)
            : ReadChartAxisStroke(xmlAxis, theme);
    }

    private static PptxSceneChartAxisTickMark ReadSceneOrXmlChartAxisMajorTickMark(PptxSceneChartAxis? sceneAxis, XElement? xmlAxis)
    {
        if (sceneAxis is not null)
        {
            return ResolveChartAxisTickMark(sceneAxis.MajorTickMarkKind);
        }

        string majorTickMark = PptxSceneBuilder.ReadChartElementValue(xmlAxis, "majorTickMark");
        return ResolveChartAxisTickMark(PptxSceneBuilder.ParseChartAxisTickMark(
            string.IsNullOrEmpty(majorTickMark) ? "none" : majorTickMark));

        PptxSceneChartAxisTickMark ResolveChartAxisTickMark(PptxSceneChartAxisTickMark tickMark)
        {
            return tickMark == PptxSceneChartAxisTickMark.Unknown
                ? PptxSceneChartAxisTickMark.None
                : tickMark;
        }
    }

    private static PptxSceneChartTickLabelPosition ResolveChartTickLabelPosition(PptxSceneChartTickLabelPosition position)
    {
        return position == PptxSceneChartTickLabelPosition.Unknown
            ? PptxSceneChartTickLabelPosition.NextTo
            : position;
    }

    private static bool ResolveSceneOrXmlValueAxisRightSide(PptxSceneChart? chart, PptxSceneChartAxis? sceneAxis, XElement? axis, bool defaultRightSide)
    {
        PptxSceneChartAxisPosition position = sceneAxis?.PositionKind ??
            PptxSceneBuilder.ParseChartAxisPosition(PptxSceneBuilder.ReadChartElementValue(axis, "axPos"));
        bool declaredSide = position switch
        {
            PptxSceneChartAxisPosition.Right => true,
            PptxSceneChartAxisPosition.Left => false,
            _ => defaultRightSide
        };
        // A paired ordinal axis can resolve a crossing to a physical edge.
        // Unpaired, date/numeric and interior crossings retain their declared fallback.
        ChartAxisSource category = ReadSceneOrXmlPairedCategoryAxis(chart, sceneAxis, axis);
        if (category.SceneAxis is null && category.XmlAxis is null) return declaredSide;
        double? crossing = sceneAxis is not null ? sceneAxis.CrossesAt :
            PptxSceneBuilder.ReadChartElementDoubleWithValue(axis, "crossesAt").Value;
        if (crossing is not null) return declaredSide;
        string raw = sceneAxis?.Crosses ?? PptxSceneBuilder.ReadChartElementValue(axis, "crosses");
        PptxSceneChartAxisCrosses crosses = string.IsNullOrEmpty(raw) ? PptxSceneChartAxisCrosses.AutoZero :
            sceneAxis?.CrossesKind ?? PptxSceneBuilder.ParseChartAxisCrosses(raw);
        bool reversed = ReadSceneOrXmlValueAxisReversed(category.SceneAxis, category.XmlAxis);
        return crosses switch
        {
            PptxSceneChartAxisCrosses.Maximum => !reversed,
            PptxSceneChartAxisCrosses.Minimum or PptxSceneChartAxisCrosses.AutoZero => reversed,
            _ => declaredSide
        };
    }

    private static ChartAxisSource ReadSceneOrXmlPairedCategoryAxis(PptxSceneChart? chart, PptxSceneChartAxis? sceneAxis, XElement? axis)
    {
        string? id = sceneAxis?.Id ?? ReadChartAxisId(axis);
        string? crossId = sceneAxis?.CrossAxisId ?? PptxSceneBuilder.ReadChartElementValue(axis, "crossAx");
        if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(crossId)) return default;
        if (sceneAxis is not null)
            return new(chart?.Axes.FirstOrDefault(candidate => candidate.AxisKind == PptxSceneChartAxisKind.Category &&
                candidate.Id == crossId && candidate.CrossAxisId == id), null);
        // An authoritative scene must not acquire missing axes from fallback XML.
        if (chart is not null) return default;
        return new(null, axis?.Parent?.Elements(ChartNamespace + "catAx").FirstOrDefault(candidate =>
            ReadChartAxisId(candidate) == crossId && PptxSceneBuilder.ReadChartElementValue(candidate, "crossAx") == id));
    }

    private static bool ResolveSceneOrXmlValueAxisBottomSide(PptxSceneChartAxis? sceneAxis, XElement? axis, bool defaultBottomSide)
    {
        if (sceneAxis is not null)
        {
            return sceneAxis.PositionKind switch
            {
                PptxSceneChartAxisPosition.Bottom => true,
                PptxSceneChartAxisPosition.Top => false,
                _ => defaultBottomSide
            };
        }

        string position = PptxSceneBuilder.ReadChartElementValue(axis, "axPos");
        return PptxSceneBuilder.ParseChartAxisPosition(position) switch
        {
            PptxSceneChartAxisPosition.Bottom => true,
            PptxSceneChartAxisPosition.Top => false,
            _ => defaultBottomSide
        };
    }

    private static bool ResolveSceneOrXmlCategoryAxisRightSide(PptxSceneChartAxis? sceneAxis, XElement? axis, bool defaultRightSide)
    {
        if (sceneAxis is not null)
        {
            return sceneAxis.PositionKind switch
            {
                PptxSceneChartAxisPosition.Right => true,
                PptxSceneChartAxisPosition.Left => false,
                _ => defaultRightSide
            };
        }

        string position = PptxSceneBuilder.ReadChartElementValue(axis, "axPos");
        return PptxSceneBuilder.ParseChartAxisPosition(position) switch
        {
            PptxSceneChartAxisPosition.Right => true,
            PptxSceneChartAxisPosition.Left => false,
            _ => defaultRightSide
        };
    }

    private static bool ResolveSceneOrXmlCategoryAxisTopSide(PptxSceneChartAxis? sceneAxis, XElement? axis, bool defaultTopSide)
    {
        if (sceneAxis is not null)
        {
            return sceneAxis.PositionKind switch
            {
                PptxSceneChartAxisPosition.Top => true,
                PptxSceneChartAxisPosition.Bottom => false,
                _ => defaultTopSide
            };
        }

        string position = PptxSceneBuilder.ReadChartElementValue(axis, "axPos");
        return PptxSceneBuilder.ParseChartAxisPosition(position) switch
        {
            PptxSceneChartAxisPosition.Top => true,
            PptxSceneChartAxisPosition.Bottom => false,
            _ => defaultTopSide
        };
    }

    private static bool IsChartAxisDeleted(XElement? axis)
    {
        XElement? delete = axis?.Element(ChartNamespace + "delete");
        return PptxSceneBuilder.IsOoxmlBooleanElementEnabled(delete);
    }

    private static bool IsSceneOrXmlChartAxisLabelVisible(PptxSceneChartAxis? sceneAxis, XElement? axis)
    {
        if (sceneAxis is null)
        {
            return IsChartAxisLabelVisible();
        }

        return sceneAxis.IsDeleted != true &&
            ResolveChartTickLabelPosition(sceneAxis.TickLabelPositionKind) != PptxSceneChartTickLabelPosition.None;

        bool IsChartAxisLabelVisible()
        {
            if (IsChartAxisDeleted(axis))
            {
                return false;
            }

            string tickLabelPosition = PptxSceneBuilder.ReadChartElementValue(axis, "tickLblPos");
            return ResolveChartTickLabelPosition(PptxSceneBuilder.ParseChartTickLabelPosition(tickLabelPosition)) != PptxSceneChartTickLabelPosition.None;
        }
    }

    private static bool ResolveSceneOrXmlValueAxisLabelsRightSide(PptxSceneChart? chart, PptxSceneChartAxis? sceneAxis, XElement? axis, bool defaultRightSide)
    {
        PptxSceneChartTickLabelPosition position = ResolveChartTickLabelPosition(sceneAxis?.TickLabelPositionKind ??
            PptxSceneBuilder.ParseChartTickLabelPosition(PptxSceneBuilder.ReadChartElementValue(axis, "tickLblPos")));
        if (position is not (PptxSceneChartTickLabelPosition.Low or PptxSceneChartTickLabelPosition.High))
            return defaultRightSide;
        // Explicit label sides follow the paired ordinal direction. Missing,
        // date/numeric and unpaired axes retain their existing physical fallback.
        ChartAxisSource category = ReadSceneOrXmlPairedCategoryAxis(chart, sceneAxis, axis);
        bool reversed = ReadSceneOrXmlValueAxisReversed(category.SceneAxis, category.XmlAxis);
        return (position == PptxSceneChartTickLabelPosition.High) != reversed;
    }

    private static double ResolveSceneOrXmlCategoryAxisLabelOffsetScale(PptxSceneChartAxis? sceneAxis, XElement? axis)
    {
        int offset = sceneAxis is not null
            ? sceneAxis.LabelOffset ?? PptxChartMetricRules.CategoryAxisDefaultLabelOffset
            : PptxSceneBuilder.ReadChartElementIntWithValue(axis, "lblOffset").Value ?? PptxChartMetricRules.CategoryAxisDefaultLabelOffset;

        offset = Math.Clamp(
            offset,
            PptxChartMetricRules.CategoryAxisMinimumLabelOffset,
            PptxChartMetricRules.CategoryAxisMaximumLabelOffset);

        return offset / (double)PptxChartMetricRules.CategoryAxisDefaultLabelOffset;
    }

    private static int ResolveSceneOrXmlCategoryAxisTickLabelSkip(PptxSceneChartAxis? sceneAxis, XElement? axis)
    {
        int skip = sceneAxis is not null
            ? sceneAxis.TickLabelSkip ?? PptxChartMetricRules.CategoryAxisDefaultTickLabelSkip
            : PptxSceneBuilder.ReadChartElementIntWithValue(axis, "tickLblSkip").Value ?? PptxChartMetricRules.CategoryAxisDefaultTickLabelSkip;

        return Math.Max(PptxChartMetricRules.CategoryAxisDefaultTickLabelSkip, skip);
    }

    private static bool ResolveSceneOrXmlCategoryAxisLabelsOnTickMarks(PptxSceneChartAxis? sceneAxis, XElement? axis)
    {
        if (sceneAxis is not null)
        {
            return ResolveChartAxisCrossBetween(sceneAxis.CrossBetweenKind) == PptxSceneChartAxisCrossBetween.MidpointCategory;
        }

        string crossBetween = PptxSceneBuilder.ReadChartElementValue(axis, "crossBetween");
        return ResolveChartAxisCrossBetween(PptxSceneBuilder.ParseChartAxisCrossBetween(crossBetween)) == PptxSceneChartAxisCrossBetween.MidpointCategory;

        PptxSceneChartAxisCrossBetween ResolveChartAxisCrossBetween(PptxSceneChartAxisCrossBetween crossBetween)
        {
            return crossBetween == PptxSceneChartAxisCrossBetween.Unknown
                ? PptxSceneChartAxisCrossBetween.Between
                : crossBetween;
        }
    }

    private static ChartSeriesStroke? ReadChartAxisStroke(XElement? axis, PptxTheme theme)
    {
        return axis is null
            ? null
            : ToChartSeriesStroke(PptxSceneBuilder.ReadChartAxisLine(axis, theme), null);
    }

}
