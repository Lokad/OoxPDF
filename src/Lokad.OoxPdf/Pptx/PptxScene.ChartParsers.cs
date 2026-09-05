using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxSceneBuilder
{
    internal static PptxSceneChartMarkerSymbol ParseChartMarkerSymbol(string? symbol)
    {
        return symbol switch
        {
            "circle" => PptxSceneChartMarkerSymbol.Circle,
            "dash" => PptxSceneChartMarkerSymbol.Dash,
            "diamond" => PptxSceneChartMarkerSymbol.Diamond,
            "dot" => PptxSceneChartMarkerSymbol.Dot,
            "none" => PptxSceneChartMarkerSymbol.None,
            "plus" => PptxSceneChartMarkerSymbol.Plus,
            "square" => PptxSceneChartMarkerSymbol.Square,
            "star" => PptxSceneChartMarkerSymbol.Star,
            "triangle" => PptxSceneChartMarkerSymbol.Triangle,
            "x" => PptxSceneChartMarkerSymbol.X,
            _ when symbol?.Equals("circle", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartMarkerSymbol.Circle,
            _ when symbol?.Equals("dash", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartMarkerSymbol.Dash,
            _ when symbol?.Equals("diamond", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartMarkerSymbol.Diamond,
            _ when symbol?.Equals("dot", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartMarkerSymbol.Dot,
            _ when symbol?.Equals("none", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartMarkerSymbol.None,
            _ when symbol?.Equals("plus", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartMarkerSymbol.Plus,
            _ when symbol?.Equals("square", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartMarkerSymbol.Square,
            _ when symbol?.Equals("star", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartMarkerSymbol.Star,
            _ when symbol?.Equals("triangle", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartMarkerSymbol.Triangle,
            _ when symbol?.Equals("x", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartMarkerSymbol.X,
            _ => PptxSceneChartMarkerSymbol.Unknown
        };
    }

    internal static PptxSceneChartDataSourceReferenceKind ParseChartDataSourceReferenceKind(string? referenceKind)
    {
        return referenceKind switch
        {
            "strRef" => PptxSceneChartDataSourceReferenceKind.StringReference,
            "numRef" => PptxSceneChartDataSourceReferenceKind.NumberReference,
            "multiLvlStrRef" => PptxSceneChartDataSourceReferenceKind.MultiLevelStringReference,
            _ => PptxSceneChartDataSourceReferenceKind.Unknown
        };
    }

    internal static PptxSceneChartDataSourceCacheKind ParseChartDataSourceCacheKind(string? cacheKind)
    {
        return cacheKind switch
        {
            "strCache" => PptxSceneChartDataSourceCacheKind.StringCache,
            "numCache" => PptxSceneChartDataSourceCacheKind.NumberCache,
            "multiLvlStrCache" => PptxSceneChartDataSourceCacheKind.MultiLevelStringCache,
            _ => PptxSceneChartDataSourceCacheKind.Unknown
        };
    }

    internal static PptxSceneChartDisplayBlanksAs ParseChartDisplayBlanksAs(string? displayBlanksAs)
    {
        return displayBlanksAs switch
        {
            "gap" => PptxSceneChartDisplayBlanksAs.Gap,
            "span" => PptxSceneChartDisplayBlanksAs.Span,
            "zero" => PptxSceneChartDisplayBlanksAs.Zero,
            _ when displayBlanksAs?.Equals("gap", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartDisplayBlanksAs.Gap,
            _ when displayBlanksAs?.Equals("span", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartDisplayBlanksAs.Span,
            _ when displayBlanksAs?.Equals("zero", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartDisplayBlanksAs.Zero,
            _ => PptxSceneChartDisplayBlanksAs.Unknown
        };
    }

    internal static PptxSceneChartDataLabelPosition ParseChartDataLabelPosition(string? position)
    {
        return position switch
        {
            "bestFit" => PptxSceneChartDataLabelPosition.BestFit,
            "b" => PptxSceneChartDataLabelPosition.Bottom,
            "ctr" => PptxSceneChartDataLabelPosition.Center,
            "inBase" => PptxSceneChartDataLabelPosition.InsideBase,
            "inEnd" => PptxSceneChartDataLabelPosition.InsideEnd,
            "l" => PptxSceneChartDataLabelPosition.Left,
            "outEnd" => PptxSceneChartDataLabelPosition.OutsideEnd,
            "r" => PptxSceneChartDataLabelPosition.Right,
            "t" => PptxSceneChartDataLabelPosition.Top,
            _ when position?.Equals("bestFit", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartDataLabelPosition.BestFit,
            _ when position?.Equals("b", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartDataLabelPosition.Bottom,
            _ when position?.Equals("ctr", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartDataLabelPosition.Center,
            _ when position?.Equals("inBase", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartDataLabelPosition.InsideBase,
            _ when position?.Equals("inEnd", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartDataLabelPosition.InsideEnd,
            _ when position?.Equals("l", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartDataLabelPosition.Left,
            _ when position?.Equals("outEnd", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartDataLabelPosition.OutsideEnd,
            _ when position?.Equals("r", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartDataLabelPosition.Right,
            _ when position?.Equals("t", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartDataLabelPosition.Top,
            _ => PptxSceneChartDataLabelPosition.Unknown
        };
    }

    internal static PptxSceneChartLegendPosition ParseChartLegendPosition(string? position)
    {
        return position switch
        {
            "b" => PptxSceneChartLegendPosition.Bottom,
            "l" => PptxSceneChartLegendPosition.Left,
            "r" => PptxSceneChartLegendPosition.Right,
            "t" => PptxSceneChartLegendPosition.Top,
            "tr" => PptxSceneChartLegendPosition.TopRight,
            _ when position?.Equals("b", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartLegendPosition.Bottom,
            _ when position?.Equals("l", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartLegendPosition.Left,
            _ when position?.Equals("r", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartLegendPosition.Right,
            _ when position?.Equals("t", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartLegendPosition.Top,
            _ when position?.Equals("tr", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartLegendPosition.TopRight,
            _ => PptxSceneChartLegendPosition.Unknown
        };
    }

    internal static PptxSceneChartPlotKind ParseChartPlotKind(string? kind)
    {
        return kind switch
        {
            "areaChart" => PptxSceneChartPlotKind.Area,
            "barChart" => PptxSceneChartPlotKind.Bar,
            "bubbleChart" => PptxSceneChartPlotKind.Bubble,
            "doughnutChart" => PptxSceneChartPlotKind.Doughnut,
            "lineChart" => PptxSceneChartPlotKind.Line,
            "pieChart" => PptxSceneChartPlotKind.Pie,
            "radarChart" => PptxSceneChartPlotKind.Radar,
            "scatterChart" => PptxSceneChartPlotKind.Scatter,
            _ when kind?.Equals("areaChart", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartPlotKind.Area,
            _ when kind?.Equals("barChart", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartPlotKind.Bar,
            _ when kind?.Equals("bubbleChart", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartPlotKind.Bubble,
            _ when kind?.Equals("doughnutChart", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartPlotKind.Doughnut,
            _ when kind?.Equals("lineChart", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartPlotKind.Line,
            _ when kind?.Equals("pieChart", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartPlotKind.Pie,
            _ when kind?.Equals("radarChart", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartPlotKind.Radar,
            _ when kind?.Equals("scatterChart", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartPlotKind.Scatter,
            _ => PptxSceneChartPlotKind.Unknown
        };
    }

    internal static string? GetChartPlotElementName(PptxSceneChartPlotKind kind)
    {
        return kind switch
        {
            PptxSceneChartPlotKind.Area => "areaChart",
            PptxSceneChartPlotKind.Bar => "barChart",
            PptxSceneChartPlotKind.Bubble => "bubbleChart",
            PptxSceneChartPlotKind.Doughnut => "doughnutChart",
            PptxSceneChartPlotKind.Line => "lineChart",
            PptxSceneChartPlotKind.Pie => "pieChart",
            PptxSceneChartPlotKind.Radar => "radarChart",
            PptxSceneChartPlotKind.Scatter => "scatterChart",
            _ => null
        };
    }

    internal static PptxSceneChartGrouping ParseChartGrouping(string? grouping)
    {
        return grouping switch
        {
            "clustered" => PptxSceneChartGrouping.Clustered,
            "percentStacked" => PptxSceneChartGrouping.PercentStacked,
            "stacked" => PptxSceneChartGrouping.Stacked,
            "standard" => PptxSceneChartGrouping.Standard,
            _ when grouping?.Equals("clustered", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartGrouping.Clustered,
            _ when grouping?.Equals("percentStacked", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartGrouping.PercentStacked,
            _ when grouping?.Equals("stacked", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartGrouping.Stacked,
            _ when grouping?.Equals("standard", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartGrouping.Standard,
            _ => PptxSceneChartGrouping.Unknown
        };
    }

    internal static PptxSceneChartBarDirection ParseChartBarDirection(string? direction)
    {
        return direction switch
        {
            "bar" => PptxSceneChartBarDirection.Bar,
            "col" => PptxSceneChartBarDirection.Column,
            _ when direction?.Equals("bar", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartBarDirection.Bar,
            _ when direction?.Equals("col", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartBarDirection.Column,
            _ => PptxSceneChartBarDirection.Unknown
        };
    }

    internal static PptxSceneChartScatterStyle ParseChartScatterStyle(string? style)
    {
        return style switch
        {
            "line" => PptxSceneChartScatterStyle.Line,
            "lineMarker" => PptxSceneChartScatterStyle.LineMarker,
            "marker" => PptxSceneChartScatterStyle.Marker,
            "none" => PptxSceneChartScatterStyle.None,
            "smooth" => PptxSceneChartScatterStyle.Smooth,
            "smoothMarker" => PptxSceneChartScatterStyle.SmoothMarker,
            _ when style?.Equals("line", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartScatterStyle.Line,
            _ when style?.Equals("lineMarker", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartScatterStyle.LineMarker,
            _ when style?.Equals("marker", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartScatterStyle.Marker,
            _ when style?.Equals("none", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartScatterStyle.None,
            _ when style?.Equals("smooth", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartScatterStyle.Smooth,
            _ when style?.Equals("smoothMarker", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartScatterStyle.SmoothMarker,
            _ => PptxSceneChartScatterStyle.Unknown
        };
    }

    internal static PptxSceneChartRadarStyle ParseChartRadarStyle(string? style)
    {
        return style switch
        {
            "filled" => PptxSceneChartRadarStyle.Filled,
            "marker" => PptxSceneChartRadarStyle.Marker,
            "standard" => PptxSceneChartRadarStyle.Standard,
            _ when style?.Equals("filled", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartRadarStyle.Filled,
            _ when style?.Equals("marker", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartRadarStyle.Marker,
            _ when style?.Equals("standard", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartRadarStyle.Standard,
            _ => PptxSceneChartRadarStyle.Unknown
        };
    }

    internal static PptxSceneChartAxisPosition ParseChartAxisPosition(string? position)
    {
        return position switch
        {
            "b" => PptxSceneChartAxisPosition.Bottom,
            "l" => PptxSceneChartAxisPosition.Left,
            "r" => PptxSceneChartAxisPosition.Right,
            "t" => PptxSceneChartAxisPosition.Top,
            _ when position?.Equals("b", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisPosition.Bottom,
            _ when position?.Equals("l", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisPosition.Left,
            _ when position?.Equals("r", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisPosition.Right,
            _ when position?.Equals("t", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisPosition.Top,
            _ => PptxSceneChartAxisPosition.Unknown
        };
    }

    internal static PptxSceneChartAxisKind ParseChartAxisKind(string? kind)
    {
        return kind switch
        {
            "catAx" => PptxSceneChartAxisKind.Category,
            "dateAx" => PptxSceneChartAxisKind.Date,
            "serAx" => PptxSceneChartAxisKind.Series,
            "valAx" => PptxSceneChartAxisKind.Value,
            _ when kind?.Equals("catAx", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisKind.Category,
            _ when kind?.Equals("dateAx", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisKind.Date,
            _ when kind?.Equals("serAx", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisKind.Series,
            _ when kind?.Equals("valAx", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisKind.Value,
            _ => PptxSceneChartAxisKind.Unknown
        };
    }

    internal static PptxSceneChartManualLayoutTarget ParseChartManualLayoutTarget(string? target)
    {
        return target switch
        {
            "inner" => PptxSceneChartManualLayoutTarget.Inner,
            "outer" => PptxSceneChartManualLayoutTarget.Outer,
            _ when target?.Equals("inner", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartManualLayoutTarget.Inner,
            _ when target?.Equals("outer", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartManualLayoutTarget.Outer,
            _ => PptxSceneChartManualLayoutTarget.Unknown
        };
    }

    internal static PptxSceneChartManualLayoutMode ParseChartManualLayoutMode(string? mode)
    {
        return mode switch
        {
            "edge" => PptxSceneChartManualLayoutMode.Edge,
            "factor" => PptxSceneChartManualLayoutMode.Factor,
            _ when mode?.Equals("edge", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartManualLayoutMode.Edge,
            _ when mode?.Equals("factor", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartManualLayoutMode.Factor,
            _ => PptxSceneChartManualLayoutMode.Unknown
        };
    }

    internal static PptxSceneChartTickLabelPosition ParseChartTickLabelPosition(string? position)
    {
        return position switch
        {
            "high" => PptxSceneChartTickLabelPosition.High,
            "low" => PptxSceneChartTickLabelPosition.Low,
            "nextTo" => PptxSceneChartTickLabelPosition.NextTo,
            "none" => PptxSceneChartTickLabelPosition.None,
            _ when position?.Equals("high", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartTickLabelPosition.High,
            _ when position?.Equals("low", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartTickLabelPosition.Low,
            _ when position?.Equals("nextTo", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartTickLabelPosition.NextTo,
            _ when position?.Equals("none", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartTickLabelPosition.None,
            _ => PptxSceneChartTickLabelPosition.Unknown
        };
    }

    internal static PptxSceneChartAxisCrosses ParseChartAxisCrosses(string? crosses)
    {
        return crosses switch
        {
            "autoZero" => PptxSceneChartAxisCrosses.AutoZero,
            "max" => PptxSceneChartAxisCrosses.Maximum,
            "min" => PptxSceneChartAxisCrosses.Minimum,
            _ when crosses?.Equals("autoZero", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisCrosses.AutoZero,
            _ when crosses?.Equals("max", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisCrosses.Maximum,
            _ when crosses?.Equals("min", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisCrosses.Minimum,
            _ => PptxSceneChartAxisCrosses.Unknown
        };
    }

    internal static PptxSceneChartAxisCrossBetween ParseChartAxisCrossBetween(string? crossBetween)
    {
        return crossBetween switch
        {
            "between" => PptxSceneChartAxisCrossBetween.Between,
            "midCat" => PptxSceneChartAxisCrossBetween.MidpointCategory,
            _ when crossBetween?.Equals("between", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisCrossBetween.Between,
            _ when crossBetween?.Equals("midCat", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisCrossBetween.MidpointCategory,
            _ => PptxSceneChartAxisCrossBetween.Unknown
        };
    }

    internal static PptxSceneChartAxisOrientation ParseChartAxisOrientation(string? orientation)
    {
        return orientation switch
        {
            "minMax" => PptxSceneChartAxisOrientation.MinimumMaximum,
            "maxMin" => PptxSceneChartAxisOrientation.MaximumMinimum,
            _ when orientation?.Equals("minMax", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisOrientation.MinimumMaximum,
            _ when orientation?.Equals("maxMin", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisOrientation.MaximumMinimum,
            _ => PptxSceneChartAxisOrientation.Unknown
        };
    }

    internal static PptxSceneChartAxisTickMark ParseChartAxisTickMark(string? tickMark)
    {
        return tickMark switch
        {
            "cross" => PptxSceneChartAxisTickMark.Cross,
            "in" => PptxSceneChartAxisTickMark.Inside,
            "none" => PptxSceneChartAxisTickMark.None,
            "out" => PptxSceneChartAxisTickMark.Outside,
            _ when tickMark?.Equals("cross", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisTickMark.Cross,
            _ when tickMark?.Equals("in", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisTickMark.Inside,
            _ when tickMark?.Equals("none", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisTickMark.None,
            _ when tickMark?.Equals("out", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisTickMark.Outside,
            _ => PptxSceneChartAxisTickMark.Unknown
        };
    }
}
