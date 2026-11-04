using System.Xml.Linq;
using System.Text;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static double ChartCategoryLabelBaselineDistance(double fontSize, bool highSide,
        double offsetScale, double? adjustment)
    {
        double baseline = fontSize * PptxChartMetricRules.AxisLabelHeightFactor * (highSide
            ? PptxChartMetricRules.CategoryAxisVerticalTopSideOffsetFactor
            : PptxChartMetricRules.CategoryAxisVerticalTopOffsetFactor);
        return adjustment is { } offset ? baseline + offset : baseline * offsetScale;
    }

    private static async ValueTask<ChartLayout> PrepareCartesianCategoryFramesAsync(ChartLayout layout,
        PptxTheme theme, PptxSceneChart chart, XDocument xml, ChartAxisSource valueAxis,
        ChartValueExtents extents, ChartValueAxisRenderOptions options, double tickTarget,
        bool defaultRightSide, ChartAxisSource categoryAxis, IReadOnlyList<ChartIndexedTextPoint?> labels, double categoryRatio, bool categoryHighSide,
        PresentationFontResolver resolver, PptxRenderContext context, ChartLegendBox? legendBox = null, string? defaultNumberFormat = null)
    {
        bool canAdmit = layout.ManualPlotLayoutApplied &&
            chart.Axes.Count(axis => axis.AxisKind == PptxSceneChartAxisKind.Value) == 1 &&
            valueAxis.SceneAxis is { HasLogarithmicScale: false, CrossesAt: null } &&
            categoryAxis.SceneAxis?.AxisKind == PptxSceneChartAxisKind.Category &&
            !chart.Title.Layout.HasGeometryValues && !layout.Legend.Layout.HasGeometryValues;
        bool innerTarget = canAdmit && chart.PlotAreaLayout.LayoutTargetKind == PptxSceneChartManualLayoutTarget.Inner;
        bool outer = canAdmit && chart.PlotAreaLayout.LayoutTargetKind == PptxSceneChartManualLayoutTarget.Outer &&
            chart.Plots.Count == 1 && chart.PlotAreaLayout is
            { X: not null, Y: not null, Width: not null, Height: not null,
              WidthModeKind: PptxSceneChartManualLayoutMode.Factor, HeightModeKind: PptxSceneChartManualLayoutMode.Factor };
        ChartLayout original = layout;
        ChartPlotBox source = layout.PlotBox;
        double numericClearance = 0d, numericBand = 0d;
        bool labelsRight = ResolveSceneOrXmlValueAxisLabelsRightSide(chart, valueAxis.SceneAxis, valueAxis.XmlAxis, defaultRightSide);
        bool onTickMarks = ResolveSceneOrXmlCategoryAxisLabelsOnTickMarks(valueAxis.SceneAxis, valueAxis.XmlAxis);
        // Outer captions fit the area remaining after numeric admission. Inner
        // and automatic layouts retain their established preparation order.
        if (outer)
        {
            (numericClearance, numericBand) = await MeasureNumericFrameAsync().ConfigureAwait(false);
            source = ResolveCartesianPlotPosition(chart.PlotAreaLayout, layout.Frame,
                new(layout.PlotAreaBox, source, PptxSceneChartManualLayoutTarget.Outer)).PlotBox;
            if (source.Width - numericBand < 1d) outer = false;
            else layout = layout with { PlotBox = source with { X = source.X + (labelsRight ? 0d : numericBand),
                Width = source.Width - numericBand } };
        }
        bool categoryVisible = categoryAxis.SceneAxis is { } axis ? axis.IsDeleted != true : !IsChartAxisDeleted(categoryAxis.XmlAxis);
        if (categoryVisible)
            layout = await PrepareCategoryLabelLayoutAsync(layout, theme, chart, xml, categoryAxis, labels,
                onTickMarks, categoryHighSide, categoryRatio, resolver, legendBox,
                outer ? Math.Max(0d, source.Height - 2d * numericClearance) : null).ConfigureAwait(false);
        if (!innerTarget && !outer) return layout;
        if (layout.CategoryLabels?.Labels.Values.Any(parts => parts.Length != 1) == true)
            return outer ? await PrepareCategoryLabelLayoutAsync(original, theme, chart, xml, categoryAxis, labels,
                onTickMarks, categoryHighSide, categoryRatio, resolver, legendBox).ConfigureAwait(false) : layout;
        if (innerTarget)
        {
            if (layout.CategoryLabels is { RotationDegrees: not 0d }) return layout;
            (numericClearance, numericBand) = await MeasureNumericFrameAsync().ConfigureAwait(false);
            source = layout.PlotBox;
        }
        if (outer && layout.CategoryLabels is { RotationDegrees: not 0d } rotated)
        {
            // Mirroring an outer caption changes its anchor, not its occupied
            // clearance. Native side/direction/offset controls reserve equal
            // heights. Keep the existing inner/automatic baseline owners.
            // An authored offset is an increment in normal-line units, as on
            // the automatic path; it does not scale the full baseline distance.
            double offset = rotated.LineHeight * PptxChartMetricRules.CategoryAxisLabelOffsetNormalLineHeightRatio *
                (ResolveSceneOrXmlCategoryAxisLabelOffsetScale(categoryAxis.SceneAxis, categoryAxis.XmlAxis) - 1d);
            double gap = Math.Max(0d, ChartCategoryLabelBaselineDistance(rotated.Style.FontSize, highSide: false,
                offsetScale: 1d, adjustment: offset) - rotated.Ascent);
            layout = layout with { CategoryLabels = rotated with { AxisGap = gap } };
        }
        double categoryClearance = 0d, categoryBand = 0d;
        double captionRadians = -(layout.CategoryLabels?.RotationDegrees ?? 0d) * Math.PI / 180d;
        double captionSin = Math.Sin(captionRadians), captionCos = Math.Cos(captionRadians);
        bool rightCaptionAnchor = (captionRadians > 0d) != categoryHighSide;
        if (layout.CategoryLabels is { } category && category.Labels.Values.Any(parts => parts.Any(part => part.Text.Length > 0)))
        {
            categoryClearance = category.LineHeight / 2d;
            double baseline = ChartCategoryLabelBaselineDistance(category.Style.FontSize, categoryHighSide,
                ResolveSceneOrXmlCategoryAxisLabelOffsetScale(categoryAxis.SceneAxis, categoryAxis.XmlAxis),
                layout.AutomaticCategoryLabelOffsetAdjustment);
            if (category.RotationDegrees == 0d)
                categoryBand = baseline + (categoryHighSide ? category.Ascent : category.Descent) + PptxChartMetricRules.OuterPlotCaptionPadding;
            else
            {
                categoryBand = category.AxisGap + category.Labels.Values.Max(parts =>
                    (parts[0].Width - (rightCaptionAnchor ? parts[0].OverflowAdvance : 0d)) * Math.Abs(captionSin) + category.LineHeight * captionCos) +
                    PptxChartMetricRules.OuterPlotCaptionPadding;
            }
        }
        if (outer)
        {
            // A present category axis retains a minimum external band even
            // without labels, ticks or label offset. Independent native controls
            // follow the existing nominal tick distance; deleting it releases
            // the band. This is a layout reserve, separate from painted ink.
            if (categoryAxis.SceneAxis!.IsDeleted != true)
            {
                ChartTextStyle style = layout.CategoryLabels?.Style ?? ReadSceneOrXmlChartTextStyle(theme, chart,
                    categoryAxis.SceneAxis, xml, categoryAxis.XmlAxis, PptxChartMetricRules.CategoryAxisFallbackFontSize, "categoryAxis");
                categoryBand = Math.Max(categoryBand, style.FontSize * PptxChartMetricRules.ChartAxisMajorTickLengthFactor);
            }
            bool reversed = ReadSceneOrXmlValueAxisReversed(categoryAxis.SceneAxis, categoryAxis.XmlAxis);
            (double Left, double Right) RequiredBands(double width)
            {
                double left = labelsRight ? 0d : numericBand, right = labelsRight ? numericBand : 0d;
                if (labels.Count > 0 && layout.CategoryLabels is { } captions)
                    foreach (var (index, parts) in captions.Labels)
                    {
                        resolver.CancellationToken.ThrowIfCancellationRequested();
                        double ratio = ChartCategoryToPlotCoordinate(index, labels.Count, 0d, 1d, onTickMarks, reversed);
                        var frame = ChartCategoryCaptionHorizontalFrame(0d, parts[0].Width, captions.Ascent, captions.Descent,
                            captionSin, captionCos, rightCaptionAnchor);
                        left = Math.Max(left, -frame.Left + PptxChartMetricRules.DataLabelHorizontalPadding - ratio * width);
                        right = Math.Max(right, frame.Right + PptxChartMetricRules.DataLabelHorizontalPadding - (1d - ratio) * width);
                    }
                return (left, right);
            }
            var bands = AdmitChartCaptionBands(source.Width, source.Width, RequiredBands);
            if (bands is null)
                return await PrepareCategoryLabelLayoutAsync(original, theme, chart, xml, categoryAxis, labels,
                    onTickMarks, categoryHighSide, categoryRatio, resolver, legendBox).ConfigureAwait(false);
            var verticalBand = AdmitChartCaptionInterval(source.Y, source.Height, source.Y, source.Height,
                numericClearance, numericClearance, categoryBand, new(0d, categoryRatio, categoryHighSide, categoryHighSide));
            ChartPlotBox inner = new(source.X + bands.Value.Left, verticalBand.Start, bands.Value.Width, verticalBand.Length);
            // Plot fill, strokes, data, captions and annotations share the inner
            // data rectangle; the authored outer area only owns admission.
            return layout with { PlotBox = inner, PlotAreaBox = new(inner.X, inner.Y, inner.Width, inner.Height) };
        }
        var horizontal = AdmitChartCaptionInterval(layout.Frame.X, layout.Frame.Width, source.X, source.Width,
            0d, 0d, numericBand, new(0d, labelsRight ? 1d : 0d, labelsRight, labelsRight));
        // Each role activates its own constraint. An overflowing category endcap
        // admits the occupied frame as a unit. If only numeric endcaps overflow,
        // retain their admitted interval and allocate the category band inside
        // it. This also preserves the opposite edge for an interior crossing.
        (double Start, double Length) vertical = (source.Y, source.Height);
        var categoryPlacement = new ChartAxisPlacement(0d, categoryRatio, categoryHighSide, categoryHighSide);
        bool EndcapsOverflow(double clearance) => source.Y - clearance < layout.Frame.Y ||
            source.Y + source.Height + clearance > layout.Frame.Y + layout.Frame.Height;
        if (EndcapsOverflow(categoryClearance))
            vertical = AdmitChartCaptionInterval(layout.Frame.Y, layout.Frame.Height, source.Y, source.Height,
                numericClearance, numericClearance, categoryBand, categoryPlacement);
        else if (EndcapsOverflow(numericClearance))
        {
            var numeric = AdmitChartCaptionInterval(layout.Frame.Y, layout.Frame.Height, source.Y, source.Height,
                numericClearance, numericClearance, 0d, categoryPlacement);
            double insideBand = Math.Max(0d, categoryHighSide
                ? numeric.Start + numeric.Length + categoryBand - layout.Frame.Y - layout.Frame.Height
                : layout.Frame.Y + categoryBand - numeric.Start);
            vertical = AdmitChartCaptionInterval(numeric.Start, numeric.Length, numeric.Start, numeric.Length,
                0d, 0d, insideBand, categoryPlacement);
        }
        ChartPlotBox plot = source with { X = horizontal.Start, Width = horizontal.Length, Y = vertical.Start, Height = vertical.Length };
        return layout with
        {
            PlotBox = plot,
            PlotAreaBox = layout.PlotAreaBox with { X = layout.PlotAreaBox.X + plot.X - source.X,
                Y = layout.PlotAreaBox.Y + plot.Y - source.Y, Width = Math.Max(1d, layout.PlotAreaBox.Width + plot.Width - source.Width),
                Height = Math.Max(1d, layout.PlotAreaBox.Height + plot.Height - source.Height) }
        };

        async ValueTask<(double Clearance, double Band)> MeasureNumericFrameAsync()
        {
            if (!IsSceneOrXmlChartAxisLabelVisible(valueAxis.SceneAxis, valueAxis.XmlAxis)) return (0d, 0d);
            var measure = new ChartTextMeasurer(resolver, kerningEnabled: false);
            ChartTextStyle style = ReadSceneOrXmlChartTextStyle(theme, chart, valueAxis.SceneAxis, xml, valueAxis.XmlAxis,
                PptxChartMetricRules.ValueAxisFallbackFontSize, "valueAxis");
            double clearance = (await measure.MeasureLineBoxAsync(style).ConfigureAwait(false)).Height / 2d;
            var captions = await MeasureChartValueAxisCaptionsAsync(measure, style, valueAxis, extents,
                options.Units, tickTarget, options.TickPolicy, defaultNumberFormat, reserveEndpoints: true, context).ConfigureAwait(false);
            return (clearance, captions.MaximumWidth + Math.Max(3d, style.FontSize * PptxChartMetricRules.ValueAxisLabelSideGapFactor) +
                PptxChartMetricRules.OuterPlotCaptionPadding);
        }
    }

    private static async ValueTask<(ChartLayout Layout, ChartLegendBox? Box)> AdmitCartesianRightLegendAsync(
        ChartLayout layout, PptxSceneChart chart, ChartAxisSource valueAxis,
        IReadOnlyList<ChartIndexedTextPoint?> labels, bool onTickMarks, bool reversed,
        IReadOnlyList<ChartLegendEntry> entries, ChartTextStyle legendStyle,
        PresentationFontResolver resolver, ChartLegendPlacement placement)
    {
        if (layout.ManualPlotLayoutApplied || !string.IsNullOrWhiteSpace(layout.Title) ||
            !layout.Legend.Visible || layout.Legend.Overlay || layout.Legend.PositionKind != PptxSceneChartLegendPosition.Right ||
            layout.Legend.Layout.HasGeometryValues || chart.Plots.Count != 1 ||
            chart.Axes.Count(axis => axis.AxisKind == PptxSceneChartAxisKind.Value) != 1 ||
            valueAxis.SceneAxis is not { HasLogarithmicScale: false } ||
            ReadSceneOrXmlPairedCategoryAxis(chart, valueAxis.SceneAxis, valueAxis.XmlAxis).SceneAxis is null ||
            layout.CategoryLabels is not { RotationDegrees: 0d } captions ||
            captions.Labels.Values.Any(parts => parts.Length != 1) || labels.Count == 0 || entries.Count == 0)
            return (layout, null);

        // Office controls establish horizontal, unwrapped caption admission.
        // Rotated/wrapped content retains its existing layout and legend path.
        // Preserve the already measured legend position while the inner plot
        // admits numeric and category fragments. Side stroke boxes otherwise
        // follow the plot right edge during painting.
        var measure = new ChartTextMeasurer(resolver, kerningEnabled: false);
        await measure.PrepareAsync(entries.Select(entry => entry.Name), legendStyle).ConfigureAwait(false);
        ChartLegendBox legend = ResolveChartLegendBox(layout.Frame, layout.PlotBox, entries,
            layout.Legend, legendStyle, measure, placement);
        double margin = PptxChartMetricRules.ColumnPlotRightReserve;
        double outerLeft = layout.Frame.X + margin;
        double outerRight = layout.PlotBox.X + layout.PlotBox.Width;
        if (placement == ChartLegendPlacement.AreaRightLegend)
        {
            // Share the measured allocation used by bubble side legends. The
            // inherited fixed area block remains only on the fallback path.
            if (await MeasureSideLegendFrameAsync(measure, legendStyle, entries).ConfigureAwait(false) is not { } frame) return (layout, null);
            outerRight = layout.Frame.X + layout.Frame.Width - margin -
                PptxChartMetricRules.ChartContentFrameInset - frame.Width;
            legend = PrepareAreaSideLegend(legend, layout.Frame, entries.Count, legendStyle, frame.LineHeight);
        }
        double available = outerRight - outerLeft;
        double numericBand = IsSceneOrXmlChartAxisLabelVisible(valueAxis.SceneAxis, valueAxis.XmlAxis)
            ? Math.Max(0d, layout.PlotBox.X - outerLeft) : 0d;
        bool numericRight = ResolveSceneOrXmlValueAxisLabelsRightSide(chart, valueAxis.SceneAxis, valueAxis.XmlAxis,
            ResolveSceneOrXmlValueAxisRightSide(chart, valueAxis.SceneAxis, valueAxis.XmlAxis, defaultRightSide: false));
        double numericLeft = numericRight ? 0d : numericBand;
        double numericRightBand = numericRight ? numericBand : 0d;
        if (available - numericLeft - numericRightBand < 1d) return (layout, null);

        (double Left, double Right) RequiredBands(double width)
        {
            double left = numericLeft, right = numericRightBand;
            foreach (var (index, parts) in captions.Labels)
            {
                resolver.CancellationToken.ThrowIfCancellationRequested();
                double ratio = ChartCategoryToPlotCoordinate(index, labels.Count, 0d, 1d, onTickMarks, reversed);
                double half = parts[0].Width / 2d;
                left = Math.Max(left, half - ratio * width);
                right = Math.Max(right, half - (1d - ratio) * width);
                // Both ordinal range ends reserve their endpoint advance;
                // for a sole category the first and last source are the same.
                if (onTickMarks && index == labels.Count - 1)
                {
                    left = Math.Max(left, reversed ? half : 0d);
                    right = Math.Max(right, reversed ? 0d : half);
                }
            }
            return (left, right);
        }

        var admitted = AdmitChartCaptionBands(available, available - numericLeft - numericRightBand, RequiredBands);
        if (admitted is null) return (layout, null);
        double lower = admitted.Value.Width;
        double slotWidth = lower / labels.Count * Math.Min(captions.TickLabelSkip, labels.Count) *
            PptxChartMetricRules.CategoryAxisVerticalWidthFactor;
        if (captions.Labels.Values.Any(parts => parts[0].Width > slotWidth)) return (layout, null);
        ChartPlotBox plot = layout.PlotBox with { X = outerLeft + admitted.Value.Left, Width = lower };
        return (layout with { PlotBox = plot, PlotAreaBox = new(plot.X, plot.Y, plot.Width, plot.Height) }, legend);
    }

    private static (double Left, double Width, double Right)? AdmitChartCaptionBands(double available, double maximumWidth,
        Func<double, (double Left, double Right)> requiredBands)
    {
        // The intersection of these linear frame constraints is convex. Zero
        // width supplies a feasible lower bound; overfull fragments keep the
        // existing fallback. Bisection finds the widest admitted inner plot.
        var minimum = requiredBands(0d);
        if (minimum.Left + minimum.Right > available) return null;
        double lower = 0d, upper = maximumWidth;
        var upperBands = requiredBands(upper);
        if (upper + upperBands.Left + upperBands.Right > available)
        {
            for (int step = 0; step < 40; step++)
            {
                double width = (lower + upper) / 2d;
                var bands = requiredBands(width);
                if (width + bands.Left + bands.Right <= available) lower = width;
                else upper = width;
            }
        }
        else lower = upper;
        if (lower < 1d) return null;
        var result = requiredBands(lower);
        return (result.Left, lower, result.Right);
    }

    private static (double Pivot, double Left, double Right) ChartCategoryCaptionHorizontalFrame(double center,
        double width, double ascent, double descent, double sin, double cos, bool rightAnchor)
    {
        if (sin == 0d && cos == 1d) return (center, center - width / 2d, center + width / 2d);
        double pivot = center + (ascent - descent) * sin / 2d;
        double left = pivot + (rightAnchor ? -width * cos : 0d) + (sin >= 0d ? -ascent * sin : descent * sin);
        double right = pivot + (rightAnchor ? 0d : width * cos) + (sin >= 0d ? descent * sin : -ascent * sin);
        return (pivot, left, right);
    }

    private static ChartLayout AdmitObliqueCategoryCaptionFrames(ChartLayout layout, int count, bool onTickMarks,
        bool reversed, bool highSide, CancellationToken cancellationToken)
    {
        // Horizontal captions keep their wrapping owner; vertical captions have
        // no horizontal advance. Manual plots retain authored geometry.
        if (layout.ManualPlotLayoutApplied || layout.CategoryLabels is not { } captions ||
            captions.RotationDegrees == 0d || Math.Abs(captions.RotationDegrees) == 90d) return layout;
        double radians = -captions.RotationDegrees * Math.PI / 180d;
        double sin = Math.Sin(radians), cos = Math.Cos(radians);
        bool rightAnchor = (radians > 0d) != highSide;
        double outerLeft = layout.Frame.X + PptxChartMetricRules.ChartContentFrameInset;
        double outerRight = layout.Frame.X + layout.Frame.Width - PptxChartMetricRules.ChartContentFrameInset;
        double numericLeft = Math.Max(0d, layout.PlotBox.X - outerLeft);
        double numericRight = Math.Max(0d, outerRight - layout.PlotBox.X - layout.PlotBox.Width);
        (double Left, double Right) RequiredBands(double width)
        {
            double left = numericLeft, right = numericRight;
            foreach (var (index, parts) in captions.Labels)
            {
                cancellationToken.ThrowIfCancellationRequested();
                double ratio = ChartCategoryToPlotCoordinate(index, count, 0d, 1d, onTickMarks, reversed);
                foreach (var part in parts)
                {
                    if (part.Text.Length == 0) continue;
                    var frame = ChartCategoryCaptionHorizontalFrame(0d, part.Width, captions.Ascent, captions.Descent, sin, cos, rightAnchor);
                    left = Math.Max(left, -frame.Left - ratio * width);
                    right = Math.Max(right, frame.Right - (1d - ratio) * width);
                }
            }
            return (left, right);
        }
        double available = outerRight - outerLeft;
        var original = RequiredBands(layout.PlotBox.Width);
        if (layout.PlotBox.Width + original.Left + original.Right <= available) return layout;
        var fitted = AdmitChartCaptionBands(available, layout.PlotBox.Width, RequiredBands);
        if (fitted is null) return layout;
        double x = outerLeft + fitted.Value.Left;
        double shift = x - layout.PlotBox.X;
        double shrink = layout.PlotBox.Width - fitted.Value.Width;
        return layout with
        {
            PlotBox = layout.PlotBox with { X = x, Width = fitted.Value.Width },
            PlotAreaBox = layout.PlotAreaBox with { X = layout.PlotAreaBox.X + shift, Width = Math.Max(1d, layout.PlotAreaBox.Width - shrink) }
        };
    }

    private static bool CanOwnColumnCategoryDirection(PptxSceneChart chart, ChartAxisSource category, ChartAxisSource value) =>
        chart.Plots.Count == 1 && chart.Axes.Count(axis => axis.AxisKind == PptxSceneChartAxisKind.Value) == 1 &&
        category.SceneAxis is { AxisKind: PptxSceneChartAxisKind.Category } ordinal &&
        value.SceneAxis is { AxisKind: PptxSceneChartAxisKind.Value, HasLogarithmicScale: false } numeric &&
        !string.IsNullOrEmpty(ordinal.Id) && !string.IsNullOrEmpty(numeric.Id) &&
        ordinal.CrossAxisId == numeric.Id && numeric.CrossAxisId == ordinal.Id;

    private static ChartAxisPlacement? ReadHorizontalBarCategoryPlacement(PptxSceneChart chart,
        ChartAxisSource category, ChartAxisSource value, ChartValueExtents extents, bool reversed)
    {
        // A reciprocal ordinal/numeric pair owns the source crossing. Keep
        // unpaired and unsupported composed scales on their declared fallback.
        if (chart.Plots.Count != 1 && !HasSharedCartesianChartValueAxis(chart, value.SceneAxis) ||
            chart.Axes.Count(a => a.AxisKind == PptxSceneChartAxisKind.Value) != 1 ||
            category.SceneAxis is not { AxisKind: PptxSceneChartAxisKind.Category } ordinal ||
            value.SceneAxis is not { HasScaling: true, HasLogarithmicScale: false } numeric ||
            string.IsNullOrEmpty(ordinal.Id) || string.IsNullOrEmpty(numeric.Id) ||
            ordinal.CrossAxisId != numeric.Id || numeric.CrossAxisId != ordinal.Id) return null;
        return ReadChartAxisPlacement(category, extents, reversed, horizontal: false);
    }

    private static ChartAxisPlacement? ReadChartAxisPlacement(ChartAxisSource axis, ChartValueExtents extents, bool reversed, bool horizontal = true)
    {
        if (axis.SceneAxis is null && axis.XmlAxis is null) return null;
        if (ReadSceneOrXmlValueAxisCrossingValue(axis.SceneAxis, axis.XmlAxis, extents) is not { } crossing || !double.IsFinite(crossing))
            return null;
        double? explicitCrossing = axis.SceneAxis is not null ? axis.SceneAxis.CrossesAt :
            PptxSceneBuilder.ReadChartElementDoubleWithValue(axis.XmlAxis, "crossesAt").Value;
        PptxSceneChartAxisCrosses crosses = axis.SceneAxis?.CrossesKind ??
            PptxSceneBuilder.ParseChartAxisCrosses(PptxSceneBuilder.ReadChartElementValue(axis.XmlAxis, "crosses"));
        bool maximum = explicitCrossing is null && crosses == PptxSceneChartAxisCrosses.Maximum;
        bool numeric = axis.SceneAxis is { } scene ? scene.AxisKind == PptxSceneChartAxisKind.Value
            : axis.XmlAxis?.Name == ChartNamespace + "valAx";
        PptxSceneChartAxisCrossBetween between = axis.SceneAxis?.CrossBetweenKind ??
            PptxSceneBuilder.ParseChartAxisCrossBetween(PptxSceneBuilder.ReadChartElementValue(axis.XmlAxis, "crossBetween"));
        // On numeric pairs, explicit between uses half a value unit, even when
        // the source range, data or major unit changes. Omitted/midpoint modes
        // keep the numeric crossing; min/max enums still own their range edge.
        bool edge = explicitCrossing is null && crosses is PptxSceneChartAxisCrosses.Minimum or PptxSceneChartAxisCrosses.Maximum;
        if (numeric && between == PptxSceneChartAxisCrossBetween.Between && !edge) crossing -= .5d;
        double ratio = GetChartValuePlotRatio(extents, crossing, reversed);
        // Office admits reversed upper numeric next-to captions to the low band;
        // the stroke keeps its numeric coordinate and explicit low/high still win.
        if (explicitCrossing is not null && reversed && ratio == 0d) ratio = 1d;
        var labels = ResolveChartAxisLabelPlacement(axis.SceneAxis, axis.XmlAxis, ratio, reversed, maximum);
        bool axisTop = explicitCrossing is not null
            ? horizontal
                ? ResolveSceneOrXmlCategoryAxisTopSide(axis.SceneAxis, axis.XmlAxis, defaultTopSide: false)
                : ResolveSceneOrXmlCategoryAxisRightSide(axis.SceneAxis, axis.XmlAxis, defaultRightSide: false)
            : maximum != reversed;
        return new(crossing, labels.LabelPlotRatio, labels.LabelsHighSide, axisTop);
    }

    private static (double Top, double Bottom) ResolveCartesianCategoryReserves(double frameHeight,
        double headerReserve, double categoryReserve, double numericReserve, ChartAxisPlacement placement)
    {
        double top = headerReserve, bottom = categoryReserve;
        double ratio = placement.LabelPlotRatio;
        if (placement.LabelsHighSide)
        {
            bottom = numericReserve;
            double titleReserve = headerReserve - numericReserve;
            top = ratio > 0d ? Math.Max(headerReserve,
                (titleReserve + categoryReserve - (1d - ratio) * (frameHeight - bottom)) / ratio) : headerReserve;
        }
        else if (ratio > 0d)
        {
            bottom = ratio < 1d ? Math.Max(numericReserve,
                (categoryReserve - ratio * (frameHeight - top)) / (1d - ratio)) : numericReserve;
        }
        return (top, bottom);
    }

    private static (double LabelPlotRatio, bool LabelsHighSide) ResolveChartAxisLabelPlacement(
        PptxSceneChartAxis? axis, XElement? xml, double crossingRatio, bool reversed, bool maximumCrossing = false)
    {
        PptxSceneChartTickLabelPosition position = ResolveChartTickLabelPosition(axis is not null
            ? axis.TickLabelPositionKind
            : PptxSceneBuilder.ParseChartTickLabelPosition(PptxSceneBuilder.ReadChartElementValue(xml, "tickLblPos")));
        // Low/high refer to value direction, so reversing the scale swaps both
        // the physical edge and the label side. An authored maximum crossing
        // selects the high side for next-to captions; a numeric crossesAt at
        // the same bound keeps the ordinary next-to side.
        double ratio = position == PptxSceneChartTickLabelPosition.Low ? (reversed ? 1d : 0d) :
            position == PptxSceneChartTickLabelPosition.High ? (reversed ? 0d : 1d) : crossingRatio;
        bool highSide = position == PptxSceneChartTickLabelPosition.High ||
            position == PptxSceneChartTickLabelPosition.NextTo && maximumCrossing;
        return (ratio, highSide != reversed);
    }

    // Resolve spacing before placing the plot, then paint these same measured
    // labels. Automatic angles use the outer plot strip; authored angles use
    // the inner plot. Office PDF/API controls distinguish these two owners.
    private static async ValueTask<ChartLayout> PrepareCategoryLabelLayoutAsync(
        ChartLayout layout, PptxTheme theme, PptxSceneChart chart, XDocument xml,
        ChartAxisSource axis, IReadOnlyList<ChartIndexedTextPoint?> labels,
        bool onTickMarks, bool topSide, double axisRatio,
        PresentationFontResolver resolver, ChartLegendBox? legendBox = null, double? availableCaptionHeight = null)
    {
        if (labels.Count == 0 || !IsSceneOrXmlChartAxisLabelVisible(axis.SceneAxis, axis.XmlAxis))
            return layout;
        ChartTextStyle style = ReadSceneOrXmlChartTextStyle(theme, chart, axis.SceneAxis, xml, axis.XmlAxis,
            PptxChartMetricRules.CategoryAxisFallbackFontSize, "categoryAxis");
        PptxSceneChartTextBodyProperties body = axis.SceneAxis?.TextBodyProperties ?? PptxSceneBuilder.ReadChartTextBodyProperties(axis.XmlAxis);
        // Empty axis body properties inherit Office's automatic orientation.
        // An authored horizontal orientation with omitted rotation means zero.
        bool automaticAngle = body.RotationValue == "-60000000" ||
            string.IsNullOrEmpty(body.RotationValue) && string.IsNullOrEmpty(body.OrientationValue);
        double angle = body.RotationDegrees is >= -90d and <= 90d ? body.RotationDegrees.Value : 0d;
        string skipSpelling = axis.SceneAxis?.TickLabelSkipValue ?? PptxSceneBuilder.ReadChartElementIntWithValue(axis.XmlAxis, "tickLblSkip").RawValue;
        bool automaticSkip = string.IsNullOrEmpty(skipSpelling);
        int skip = ResolveSceneOrXmlCategoryAxisTickLabelSkip(axis.SceneAxis, axis.XmlAxis);
        var measure = new ChartTextMeasurer(resolver, kerningEnabled: false);
        // Office's chart font getter quantizes hundredths of a point to twips.
        // Keep emission at its existing size; this is the spacing measurement.
        ChartTextStyle policyStyle = style with { FontSize = Math.Round(style.FontSize * 20d, MidpointRounding.AwayFromZero) / 20d };
        var line = await measure.MeasureLineBoxAsync(style).ConfigureAwait(false);
        // Tick spacing includes the horizontal-header line gap. Selecting the
        // OS/2 ink box instead admits too many Arial/Times labels (Office controls).
        double policyHeight = (await measure.MeasureLineBoxAsync(policyStyle, horizontalHeader: true).ConfigureAwait(false)).Height;
        double widestWord = 0d;
        foreach (var label in labels)
        {
            resolver.CancellationToken.ThrowIfCancellationRequested();
            if (string.IsNullOrWhiteSpace(label?.Text)) continue;
            foreach (string word in label.Text.Split(' ', StringSplitOptions.RemoveEmptyEntries))
                widestWord = Math.Max(widestWord, await measure.MeasureAsync(word, policyStyle).ConfigureAwait(false));
        }
        double availableWidth = automaticAngle && !layout.ManualPlotLayoutApplied
            ? Math.Max(1d, layout.Frame.Width - 2d * PptxChartMetricRules.ColumnPlotRightReserve)
            : layout.PlotBox.Width;
        double pitch = availableWidth / Math.Max(1, onTickMarks ? labels.Count - 1 : labels.Count);
        int RequiredSkip(double degrees)
        {
            double span = degrees == 0d ? widestWord : policyHeight / Math.Sin(Math.Abs(degrees) * Math.PI / 180d);
            return (int)Math.Clamp(Math.Ceiling(span / pitch), 1d, int.MaxValue);
        }
        if (automaticAngle)
        {
            angle = 0d;
            int needed = RequiredSkip(angle);
            foreach (double candidate in new[] { -45d, -90d })
            {
                int candidateSkip = RequiredSkip(candidate);
                if (automaticSkip ? candidateSkip < needed : needed > skip)
                {
                    angle = candidate;
                    needed = candidateSkip;
                }
            }
        }
        if (automaticSkip) skip = RequiredSkip(angle);

        double innerPitch = layout.PlotBox.Width / labels.Count;
        double flatWidth = Math.Max(1d, innerPitch * Math.Min(skip, labels.Count) * PptxChartMetricRules.CategoryAxisVerticalWidthFactor);
        var measured = new Dictionary<int, ChartCategoryTextLine[]>();
        double widest = 0d;
        int maxLines = 1;
        for (int i = 0; i < labels.Count; i++)
        {
            resolver.CancellationToken.ThrowIfCancellationRequested();
            if (i % skip != 0 || string.IsNullOrWhiteSpace(labels[i]?.Text)) continue;
            string text = labels[i]!.Text;
            double width = await measure.MeasureAsync(text, style).ConfigureAwait(false);
            var lines = new List<ChartCategoryTextLine>();
            if (angle == 0d && automaticSkip && width > flatWidth)
            {
                string current = string.Empty;
                foreach (string word in text.Split(' ', StringSplitOptions.RemoveEmptyEntries))
                {
                    string joined = current.Length == 0 ? word : current + " " + word;
                    double joinedWidth = await measure.MeasureAsync(joined, style).ConfigureAwait(false);
                    if (current.Length > 0 && joinedWidth > flatWidth)
                    {
                        lines.Add(new(current, await measure.MeasureAsync(current, style).ConfigureAwait(false)));
                        current = word;
                    }
                    else current = joined;
                }
                if (current.Length > 0) lines.Add(new(current, await measure.MeasureAsync(current, style).ConfigureAwait(false)));
            }
            else lines.Add(new(text, width));
            measured.Add(i, lines.ToArray());
            maxLines = Math.Max(maxLines, lines.Count);
            widest = Math.Max(widest, width);
        }
        double sin = Math.Sin(Math.Abs(angle) * Math.PI / 180d);
        double cos = Math.Cos(angle * Math.PI / 180d);
        double requested = angle == 0d ? (maxLines - 1) * line.Height : Math.Max(0d, widest * sin + line.Height * cos - line.Height);
        // The legacy fallback limits caption height against the inner plot.
        // Measured overflow frames below have their own fitting container.
        bool atEdge = topSide ? axisRatio == 1d : axisRatio == 0d;
        double freeEdge = topSide ? layout.Frame.Y + layout.Frame.Height : layout.Frame.Y;
        if (legendBox is { } legend && layout.Legend.Visible && !layout.Legend.Overlay &&
            layout.Legend.PositionKind == (topSide ? PptxSceneChartLegendPosition.Top : PptxSceneChartLegendPosition.Bottom))
            freeEdge = topSide ? legend.ClipY : legend.ClipY + legend.ClipHeight;
        double axisY = layout.PlotBox.Y + layout.PlotBox.Height * axisRatio;
        double existingExtra = layout.ManualPlotLayoutApplied || !atEdge
            ? Math.Max(0d, (topSide ? freeEdge - axisY : axisY - freeEdge) - ComputeCartesianLabelStripBottomReserve(style.FontSize)) : 0d;
        double admitted = layout.ManualPlotLayoutApplied || !atEdge ? 0d :
            Math.Min(requested, Math.Max(0d, (layout.PlotBox.Height - line.Height) / 2d));
        double availableExtra = admitted + existingExtra;
        if (availableCaptionHeight is { } captionHeight)
            availableExtra = Math.Max(0d, captionHeight - line.Height);
        // Reuse the header owner's clearance rather than measuring its fonts
        // or rebuilding its eligibility here. The outer fitting budget restores
        // the default flat strip and removes the two minimum end margins.
        double fittingBudget = layout.PlotBox.Height + ComputeCartesianLabelStripBottomReserve(style.FontSize) -
            2d * PptxChartMetricRules.HorizontalBarPlotTopReserve + (layout.AutomaticCaptionFittingClearance ?? 0d);
        bool overflowFrame = angle != 0d && body.VerticalOverflowValue == "ellipsis" &&
            !layout.ManualPlotLayoutApplied && atEdge && layout.AutomaticCaptionFittingClearance is not null && fittingBudget > 0d;
        const double fittingPixel = 72d / 96d;
        ChartTextStyle fittingStyle = style with
        { FontSize = OfficePdfTextEmissionProfile.FontSize(style.FontSize), CharacterSpacing = 0d };
        if (angle != 0d && (overflowFrame || availableExtra < requested))
        {
            // Bootstrap controls distinguish the fitting grid from the source
            // advances used for geometry and painting. This remains an
            // approximation of Office's text fitting, not a font hinting engine.
            double fittingHeight = Math.Ceiling(style.FontSize * PptxChartMetricRules.AxisLabelHeightFactor / fittingPixel) * fittingPixel;
            double cap = Math.Max(0d, overflowFrame ? (fittingBudget / 2d - fittingHeight * cos) / sin
                : (line.Height + availableExtra - line.Height * cos) / sin);
            for (int index = 0; index < labels.Count; index++)
            {
                if (measured.TryGetValue(index, out var parts)) measured[index] = [await FitAsync(parts[0].Text, cap).ConfigureAwait(false)];
            }
        }
        else if (angle == 0d && availableExtra < requested)
        {
            int admittedLines = 1 + (int)Math.Floor(availableExtra / line.Height);
            for (int index = 0; index < labels.Count; index++)
            {
                if (!measured.TryGetValue(index, out var parts) || parts.Length <= admittedLines) continue;
                var fitted = parts.Take(admittedLines).ToArray();
                fitted[^1] = await FitAsync(string.Join(" ", parts.Skip(admittedLines - 1).Select(part => part.Text)), flatWidth).ConfigureAwait(false);
                measured[index] = fitted;
            }
        }
        if (admitted > 0d || overflowFrame)
        {
            // Fitting uses a capacity; placement reserves the content that
            // actually fits. Discrete truncation or wrapping can leave part of
            // that capacity unused. Return it to the plot rather than making
            // unchanged captions consume half of every increase in frame height.
            // Explicit clipping still owns its original capacity: the native
            // clip control changes visible glyphs as the frame grows. Its clip
            // geometry must be established before reclaiming that space.
            double occupied = admitted;
            if (body.VerticalOverflowValue == "ellipsis")
            {
                occupied = 0d;
                foreach (var parts in measured.Values)
                {
                    resolver.CancellationToken.ThrowIfCancellationRequested();
                    occupied = Math.Max(occupied, angle == 0d
                        ? (parts.Length - 1) * line.Height
                        : (parts[0].Width - parts[0].OverflowAdvance) * sin + line.Height * cos - line.Height);
                }
            }
            admitted = Math.Min(overflowFrame ? Math.Max(0d, layout.PlotBox.Height - 1d) : admitted, Math.Max(0d, occupied));
            double shift = topSide ? 0d : admitted;
            layout = layout with
            {
                PlotBox = layout.PlotBox with { Y = layout.PlotBox.Y + shift, Height = layout.PlotBox.Height - admitted },
                PlotAreaBox = layout.PlotAreaBox with { Y = layout.PlotAreaBox.Y + shift, Height = Math.Max(1d, layout.PlotAreaBox.Height - admitted) }
            };
        }
        double axisGap = Math.Max(0d, ChartCategoryLabelBaselineDistance(style.FontSize, topSide,
            ResolveSceneOrXmlCategoryAxisLabelOffsetScale(axis.SceneAxis, axis.XmlAxis),
            layout.AutomaticCategoryLabelOffsetAdjustment) - (topSide ? line.Descent : line.Ascent));
        layout = layout with { CategoryLabels = new(style, angle, skip, line.Ascent, line.Descent, line.Height, axisGap, measured) };
        return AdmitObliqueCategoryCaptionFrames(layout, labels.Count, onTickMarks,
            ReadSceneOrXmlValueAxisReversed(axis.SceneAxis, axis.XmlAxis), topSide, resolver.CancellationToken);

        async ValueTask<ChartCategoryTextLine> FitAsync(string text, double cap)
        {
            double fullWidth = await measure.MeasureAsync(text, style).ConfigureAwait(false);
            double fittingWidth = overflowFrame ? await FitWidthAsync(text).ConfigureAwait(false) : fullWidth;
            if (fittingWidth <= cap) return new(text, fullWidth);
            string suffix = body.VerticalOverflowValue == "ellipsis" ? "\u2026" : string.Empty;
            if (await FitWidthAsync(suffix).ConfigureAwait(false) > cap) return new(string.Empty, 0d);
            int low = 0, high = text.Length, best = 0;
            while (low <= high)
            {
                resolver.CancellationToken.ThrowIfCancellationRequested();
                int middle = low + (high - low) / 2;
                int cut = middle > 0 && char.IsHighSurrogate(text[middle - 1]) ? middle - 1 : middle;
                if (await FitWidthAsync(text[..cut] + suffix).ConfigureAwait(false) <= cap)
                {
                    best = cut;
                    low = middle + 1;
                }
                else high = middle - 1;
            }
            string fitted = text[..best] + suffix;
            double width = await measure.MeasureAsync(fitted, style).ConfigureAwait(false);
            double overflow = overflowFrame ? width - await measure.MeasureAsync(text[..best], style).ConfigureAwait(false) : 0d;
            return new(fitted, width, Math.Max(0d, overflow));
        }

        async ValueTask<double> FitWidthAsync(string text)
        {
            if (!overflowFrame) return await measure.MeasureAsync(text, style).ConfigureAwait(false);
            double width = 0d;
            int count = 0;
            foreach (Rune rune in text.EnumerateRunes())
            {
                resolver.CancellationToken.ThrowIfCancellationRequested();
                double advance = await measure.MeasureAsync(rune.ToString(), fittingStyle).ConfigureAwait(false);
                width += Math.Round(advance / fittingPixel, MidpointRounding.AwayFromZero) * fittingPixel;
                if (count++ > 0) width += style.CharacterSpacing;
            }
            return width;
        }
    }
}
