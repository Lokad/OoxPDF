using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxChartAxesTests
{
    public static void PptxChartSceneAxesUseSceneOwnedXmlEvidence()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:lineChart>
                <c:catAx><c:axId val="10"/><c:tickLblPos val="low"/></c:catAx>
                <c:valAx><c:axId val="20"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        TestAssert.Equal(2, chart.Axes.Count);

        XDocument mismatchedChartXml = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart><c:axId val="10"/><c:axId val="20"/></c:lineChart>
                <c:catAx><c:axId val="10"/><c:tickLblPos val="high"/></c:catAx>
                <c:valAx><c:axId val="20"/><c:majorUnit val="99"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);
        XNamespace chartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedPlot = mismatchedChartXml.Descendants(chartNamespace + "lineChart").First();
        System.Reflection.MethodInfo readValueAxes = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartValueAxesForPlot",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected renderer value-axis bridge.");
        System.Reflection.MethodInfo readCategoryAxis = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartCategoryAxisForPlot",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected renderer category-axis bridge.");

        object sceneValueAxis = ((System.Collections.IEnumerable)(readValueAxes.Invoke(null, [chart, chart.Plots[0], mismatchedChartXml, mismatchedPlot]) ?? throw new InvalidOperationException("Expected scene value-axis collection."))).Cast<object>().Single();
        object xmlOnlyValueAxis = ((System.Collections.IEnumerable)(readValueAxes.Invoke(null, [null, null, mismatchedChartXml, mismatchedPlot]) ?? throw new InvalidOperationException("Expected XML-only value-axis collection."))).Cast<object>().Single();
        object sceneCategoryAxis = readCategoryAxis.Invoke(null, [chart, chart.Plots[0], mismatchedChartXml, mismatchedPlot]) ?? throw new InvalidOperationException("Expected scene category-axis source.");
        object xmlOnlyCategoryAxis = readCategoryAxis.Invoke(null, [null, null, mismatchedChartXml, mismatchedPlot]) ?? throw new InvalidOperationException("Expected XML-only category-axis source.");

        XElement sceneValueXmlAxis = (XElement?)sceneValueAxis.GetType().GetProperty("XmlAxis")?.GetValue(sceneValueAxis) ?? throw new InvalidOperationException("Expected scene value-axis XML evidence.");
        XElement xmlOnlyValueXmlAxis = (XElement?)xmlOnlyValueAxis.GetType().GetProperty("XmlAxis")?.GetValue(xmlOnlyValueAxis) ?? throw new InvalidOperationException("Expected XML-only value-axis XML evidence.");
        XElement sceneCategoryXmlAxis = (XElement?)sceneCategoryAxis.GetType().GetProperty("XmlAxis")?.GetValue(sceneCategoryAxis) ?? throw new InvalidOperationException("Expected scene category-axis XML evidence.");
        XElement xmlOnlyCategoryXmlAxis = (XElement?)xmlOnlyCategoryAxis.GetType().GetProperty("XmlAxis")?.GetValue(xmlOnlyCategoryAxis) ?? throw new InvalidOperationException("Expected XML-only category-axis XML evidence.");

        TestAssert.True(sceneValueXmlAxis.Element(chartNamespace + "majorUnit") is null, "Expected scene value axis to keep scene-owned XML evidence.");
        TestAssert.Equal("99", xmlOnlyValueXmlAxis.Element(chartNamespace + "majorUnit")?.Attribute("val")?.Value ?? string.Empty);
        TestAssert.Equal("low", sceneCategoryXmlAxis.Element(chartNamespace + "tickLblPos")?.Attribute("val")?.Value ?? string.Empty);
        TestAssert.Equal("high", xmlOnlyCategoryXmlAxis.Element(chartNamespace + "tickLblPos")?.Attribute("val")?.Value ?? string.Empty);
    }

    public static void PptxChartValueAxisScaleFallbackUsesSceneAuthoritativeAbsence()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:lineChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        TestAssert.Equal(0, chart.Axes.Count);

        XDocument mismatchedChartXml = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:valAx><c:axId val="20"/><c:scaling><c:min val="-100"/><c:max val="100"/></c:scaling></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);

        Type axisSourceType = typeof(PptxRenderer).GetNestedType(
            "ChartAxisSource",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart axis-source type.");
        object emptyAxisSource = Activator.CreateInstance(axisSourceType, [null, null]) ?? throw new InvalidOperationException("Expected empty axis source.");
        System.Reflection.MethodInfo resolveValueAxis = typeof(PptxRenderer).GetMethod(
            "ResolveXmlValueAxisForSource",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected value-axis fallback resolver.");

        object? sceneAxis = resolveValueAxis.Invoke(null, [chart, emptyAxisSource, mismatchedChartXml]);
        object? xmlOnlyAxis = resolveValueAxis.Invoke(null, [null, emptyAxisSource, mismatchedChartXml]);

        TestAssert.True(sceneAxis is null, "Expected scene-backed missing value axis not to be repaired from fallback XML for scale/layout estimation.");
        TestAssert.True(xmlOnlyAxis is XElement, "Expected XML-only value-axis scale lookup to keep reading XML.");
    }

    public static void PptxChartUnknownDataLabelPositionResolvesThroughExplicitDefault()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:dLbls><c:showVal val="1"/><c:dLblPos val="bogus"/></c:dLbls>
                  <c:ser><c:tx><c:v>Line Series</c:v></c:tx><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:lineChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        TestAssert.Equal(PptxSceneChartDataLabelPosition.Unknown, chart.Plots[0].DataLabels.PositionKind);
        TestAssert.Equal("bogus", chart.Plots[0].DataLabels.Position);

        System.Reflection.MethodInfo resolvePosition = typeof(PptxRenderer).GetMethod(
            "ResolveChartDataLabelPosition",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected data-label position resolver.");
        object position = resolvePosition.Invoke(null, [PptxSceneChartDataLabelPosition.Unknown]) ?? throw new InvalidOperationException("Expected resolved data-label position.");

        TestAssert.Equal(PptxSceneChartDataLabelPosition.OutsideEnd, (PptxSceneChartDataLabelPosition)position);
    }

    public static void PptxChartUnknownGroupingUsesSceneAuthoritativeDefault()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:grouping val="bogus"/>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:lineChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        TestAssert.Equal(PptxSceneChartGrouping.Unknown, chart.Plots[0].GroupingKind);
        TestAssert.Equal("bogus", chart.Plots[0].Grouping);

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:lineChart><c:grouping val="stacked"/></c:lineChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "lineChart").Single();
        System.Reflection.MethodInfo readGrouping = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartGrouping",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart grouping resolver.");
        object grouping = readGrouping.Invoke(null, [chart.Plots[0], mismatchedXmlFallback, PptxSceneChartGrouping.Standard]) ?? throw new InvalidOperationException("Expected resolved grouping.");

        TestAssert.Equal(PptxSceneChartGrouping.Standard, (PptxSceneChartGrouping)grouping);
    }

    public static void PptxChartUnknownBarDirectionUsesSceneAuthoritativeDefault()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:barChart>
                  <c:barDir val="bogus"/>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:barChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        TestAssert.Equal(PptxSceneChartBarDirection.Unknown, chart.Plots[0].BarDirectionKind);
        TestAssert.Equal("bogus", chart.Plots[0].BarDirection);

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:barChart><c:barDir val="bar"/></c:barChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "barChart").Single();
        System.Reflection.MethodInfo readBarDirection = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartBarDirection",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart bar-direction resolver.");
        object barDirection = readBarDirection.Invoke(null, [chart.Plots[0], mismatchedXmlFallback]) ?? throw new InvalidOperationException("Expected resolved bar direction.");

        TestAssert.Equal(PptxSceneChartBarDirection.Column, (PptxSceneChartBarDirection)barDirection);
    }

    public static void PptxChartBarOptionsUseSceneAuthoritativeDefaults()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:barChart>
                  <c:barDir val="bogus"/>
                  <c:grouping val="bogus"/>
                  <c:gapWidth val="futureGap"/>
                  <c:overlap val="futureOverlap"/>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:barChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartPlot plot = chart.Plots[0];
        TestAssert.Equal(PptxSceneChartGrouping.Unknown, plot.GroupingKind);
        TestAssert.Equal(PptxSceneChartBarDirection.Unknown, plot.BarDirectionKind);
        TestAssert.True(plot.VaryColors is null, "Expected missing varyColors metadata to remain distinct from the effective renderer default.");
        TestAssert.Equal(string.Empty, plot.VaryColorsValue);
        TestAssert.Equal(null, plot.GapWidth);
        TestAssert.Equal(null, plot.Overlap);

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:barChart>
                <c:barDir val="bar"/>
                <c:grouping val="stacked"/>
                <c:varyColors val="0"/>
                <c:gapWidth val="300"/>
                <c:overlap val="75"/>
              </c:barChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "barChart").Single();
        System.Reflection.MethodInfo readOptions = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartBarOptions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart bar-option resolver.");
        object options = readOptions.Invoke(null, [plot, mismatchedXmlFallback, PptxSceneChartGrouping.Clustered]) ?? throw new InvalidOperationException("Expected resolved bar options.");
        Type optionsType = options.GetType();

        TestAssert.Equal(PptxSceneChartGrouping.Clustered, (PptxSceneChartGrouping)(optionsType.GetProperty("Grouping")?.GetValue(options) ?? default(PptxSceneChartGrouping)));
        TestAssert.Equal(PptxSceneChartBarDirection.Column, (PptxSceneChartBarDirection)(optionsType.GetProperty("BarDirection")?.GetValue(options) ?? default(PptxSceneChartBarDirection)));
        object varyColors = optionsType.GetProperty("VaryColors")?.GetValue(options) ?? throw new InvalidOperationException("Expected varyColors option.");
        TestAssert.True(PptxTests.ChartBooleanOptionValue(varyColors), "Expected missing effective scene varyColors to use the scene default, not the XML fallback.");
        TestAssert.True(!PptxTests.ChartBooleanOptionIsDefined(varyColors), "Expected missing scene varyColors state to remain distinguishable from an explicit true token.");
        TestAssert.Equal(string.Empty, PptxTests.ChartBooleanOptionRawValue(varyColors));
        TestAssert.Equal(150d, (double)(optionsType.GetProperty("GapWidth")?.GetValue(options) ?? double.NaN));
        TestAssert.Equal(0d, (double)(optionsType.GetProperty("Overlap")?.GetValue(options) ?? double.NaN));
    }

    public static void PptxChartMultiValueAxisStripFactorDistinguishesOppositeSides()
    {
        System.Reflection.MethodInfo readStripFactor = typeof(PptxRenderer).GetMethod(
            "GetMultiValueAxisStripFactor",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart axis strip-factor resolver.");

        double singleLeft = (double)(readStripFactor.Invoke(null, [1, false]) ?? double.NaN);
        double singleRight = (double)(readStripFactor.Invoke(null, [1, true]) ?? double.NaN);
        double multipleLeft = (double)(readStripFactor.Invoke(null, [2, false]) ?? double.NaN);
        double multipleRight = (double)(readStripFactor.Invoke(null, [2, true]) ?? double.NaN);
        Type metricRules = typeof(PptxRenderer).GetNestedType(
            "PptxChartMetricRules",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart metric rules.");
        double primaryFactor = (double)(metricRules.GetField("BarMultiValueAxisPrimaryStripFactor")?.GetRawConstantValue() ?? double.NaN);
        double secondaryFactor = (double)(metricRules.GetField("BarMultiValueAxisSecondaryStripFactor")?.GetRawConstantValue() ?? double.NaN);

        TestAssert.Equal(1d, singleLeft);
        TestAssert.Equal(1d, singleRight);
        TestAssert.Equal(primaryFactor, multipleLeft);
        TestAssert.Equal(secondaryFactor, multipleRight);
    }

    public static void PptxChartVerticalValueAxisAutoTicksUseOfficeDenseDefault()
    {
        Type extentsType = typeof(PptxRenderer).GetNestedType(
            "ChartValueExtents",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart value extents.");
        object extents = Activator.CreateInstance(
            extentsType,
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic,
            binder: null,
            args: [0d, 50d],
            culture: CultureInfo.InvariantCulture) ?? throw new InvalidOperationException("Expected chart value extents instance.");
        System.Reflection.MethodInfo readTargetCount = typeof(PptxRenderer).GetMethod(
            "GetValueAxisAutoTickTargetCount",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart tick target resolver.");
        System.Reflection.MethodInfo readTickValues = typeof(PptxRenderer).GetMethods(
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)
            .Single(method => method.Name == "GetChartAxisTickValues" && method.GetParameters().Length == 4);

        double targetCount = (double)(readTargetCount.Invoke(null, [false, true, false]) ?? double.NaN);
        object values = readTickValues.Invoke(null, [extents, null, true, targetCount]) ?? throw new InvalidOperationException("Expected chart tick values.");
        string tickList = string.Join(
            ",",
            ((System.Collections.IEnumerable)values).Cast<object>().Select(value => Convert.ToDouble(value, CultureInfo.InvariantCulture).ToString("0.########", CultureInfo.InvariantCulture)));

        TestAssert.Equal("0,5,10,15,20,25,30,35,40,45,50", tickList);
    }

    public static void PptxChartLineOptionsUseSceneAuthoritativeDefaults()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart>
                <c:dispBlanksAs val="futureBlankPolicy"/>
                <c:plotArea>
                  <c:lineChart>
                    <c:grouping val="bogus"/>
                    <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  </c:lineChart>
                </c:plotArea>
              </c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartPlot plot = chart.Plots[0];
        TestAssert.Equal(PptxSceneChartGrouping.Unknown, plot.GroupingKind);
        TestAssert.Equal(PptxSceneChartDisplayBlanksAs.Unknown, chart.Options.DisplayBlanksAsKind);
        TestAssert.True(plot.Series[0].Smooth is null, "Expected missing smooth metadata to remain distinct from the effective renderer default.");

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XDocument mismatchedChartXml = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart>
                <c:dispBlanksAs val="span"/>
                <c:plotArea><c:lineChart>
                  <c:grouping val="stacked"/>
                  <c:ser><c:smooth val="1"/></c:ser>
                </c:lineChart></c:plotArea>
              </c:chart>
            </c:chartSpace>
            """);
        XElement mismatchedXmlFallback = mismatchedChartXml.Descendants(c + "lineChart").Single();
        System.Reflection.MethodInfo readOptions = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartLineOptions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart line-option resolver.");
        object options = readOptions.Invoke(null, [chart, plot, mismatchedChartXml, mismatchedXmlFallback, PptxSceneChartGrouping.Standard]) ?? throw new InvalidOperationException("Expected resolved line options.");
        Type optionsType = options.GetType();
        object[] smoothSeries = (((System.Collections.IEnumerable?)optionsType.GetProperty("SmoothSeries")?.GetValue(options)) ?? throw new InvalidOperationException("Expected smooth series options.")).Cast<object>().ToArray();

        TestAssert.Equal(PptxSceneChartGrouping.Standard, (PptxSceneChartGrouping)(optionsType.GetProperty("Grouping")?.GetValue(options) ?? default(PptxSceneChartGrouping)));
        TestAssert.True((bool)(optionsType.GetProperty("Stacked")?.GetValue(options) ?? true) == false, "Expected unknown scene grouping to use the supplied scene default, not fallback XML stacked.");
        TestAssert.True((bool)(optionsType.GetProperty("PercentStacked")?.GetValue(options) ?? true) == false, "Expected unknown scene grouping to stay non-percent-stacked.");
        TestAssert.True(smoothSeries.Length == 1 && !PptxTests.ChartBooleanOptionValue(smoothSeries[0]), "Expected missing scene smooth flag to use the scene default, not fallback XML smooth=1.");
        TestAssert.True(!PptxTests.ChartBooleanOptionIsDefined(smoothSeries[0]), "Expected missing scene smooth state to remain distinguishable from an explicit false token.");
        TestAssert.Equal(PptxSceneChartDisplayBlanksAs.Gap, (PptxSceneChartDisplayBlanksAs)(optionsType.GetProperty("DisplayBlanksAs")?.GetValue(options) ?? default(PptxSceneChartDisplayBlanksAs)));
    }

    public static void PptxChartAreaOptionsUseSceneAuthoritativeDefaults()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart>
                <c:dispBlanksAs val="futureBlankPolicy"/>
                <c:plotArea>
                  <c:areaChart>
                    <c:grouping val="bogus"/>
                    <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  </c:areaChart>
                </c:plotArea>
              </c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartPlot plot = chart.Plots[0];
        TestAssert.Equal(PptxSceneChartGrouping.Unknown, plot.GroupingKind);
        TestAssert.Equal(PptxSceneChartDisplayBlanksAs.Unknown, chart.Options.DisplayBlanksAsKind);

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XDocument mismatchedChartXml = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart>
                <c:dispBlanksAs val="span"/>
                <c:plotArea><c:areaChart><c:grouping val="percentStacked"/></c:areaChart></c:plotArea>
              </c:chart>
            </c:chartSpace>
            """);
        XElement mismatchedXmlFallback = mismatchedChartXml.Descendants(c + "areaChart").Single();
        System.Reflection.MethodInfo readOptions = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartAreaOptions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart area-option resolver.");
        object options = readOptions.Invoke(null, [chart, plot, mismatchedChartXml, mismatchedXmlFallback, PptxSceneChartGrouping.Standard]) ?? throw new InvalidOperationException("Expected resolved area options.");
        Type optionsType = options.GetType();

        TestAssert.Equal(PptxSceneChartGrouping.Standard, (PptxSceneChartGrouping)(optionsType.GetProperty("Grouping")?.GetValue(options) ?? default(PptxSceneChartGrouping)));
        TestAssert.True((bool)(optionsType.GetProperty("Stacked")?.GetValue(options) ?? true) == false, "Expected unknown scene grouping to use the supplied scene default, not fallback XML percentStacked.");
        TestAssert.True((bool)(optionsType.GetProperty("PercentStacked")?.GetValue(options) ?? true) == false, "Expected unknown scene grouping to stay non-percent-stacked.");
        TestAssert.Equal(PptxSceneChartDisplayBlanksAs.Gap, (PptxSceneChartDisplayBlanksAs)(optionsType.GetProperty("DisplayBlanksAs")?.GetValue(options) ?? default(PptxSceneChartDisplayBlanksAs)));
    }

    public static void PptxChartCategoryAxisLabelPolicyUsesSceneAuthoritativeDefaults()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:barChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:barChart>
                <c:catAx>
                  <c:axId val="10"/><c:axPos val="b"/><c:lblOffset val="futureOffset"/><c:tickLblSkip val="futureSkip"/><c:crossAx val="20"/>
                </c:catAx>
                <c:valAx><c:axId val="20"/><c:axPos val="l"/><c:crossAx val="10"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartAxis sceneAxis = chart.Axes.Single(axis => axis.AxisKind == PptxSceneChartAxisKind.Category);
        TestAssert.True(sceneAxis.LabelOffset is null, "Expected unparseable scene label offset to stay distinct from an authored value.");
        TestAssert.True(sceneAxis.TickLabelSkip is null, "Expected unparseable scene tick-label skip to stay distinct from an authored value.");

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedAxis = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:catAx><c:axId val="10"/><c:lblOffset val="200"/><c:tickLblSkip val="3"/></c:catAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "catAx").Single();
        System.Reflection.MethodInfo resolveLabelOffset = typeof(PptxRenderer).GetMethod(
            "ResolveSceneOrXmlCategoryAxisLabelOffsetScale",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected category-axis label-offset resolver.");
        System.Reflection.MethodInfo resolveTickLabelSkip = typeof(PptxRenderer).GetMethod(
            "ResolveSceneOrXmlCategoryAxisTickLabelSkip",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected category-axis tick-label-skip resolver.");

        double sceneOffsetScale = (double)resolveLabelOffset.Invoke(null, [sceneAxis, mismatchedAxis])!;
        int sceneSkip = (int)resolveTickLabelSkip.Invoke(null, [sceneAxis, mismatchedAxis])!;
        double xmlOnlyOffsetScale = (double)resolveLabelOffset.Invoke(null, [null, mismatchedAxis])!;
        int xmlOnlySkip = (int)resolveTickLabelSkip.Invoke(null, [null, mismatchedAxis])!;

        TestAssert.Equal(1d, sceneOffsetScale);
        TestAssert.Equal(1, sceneSkip);
        TestAssert.Equal(2d, xmlOnlyOffsetScale);
        TestAssert.Equal(3, xmlOnlySkip);
    }

    public static void PptxChartScatterOptionsUseSceneAuthoritativeDefaults()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:scatterChart>
                  <c:scatterStyle val="bogus"/>
                  <c:ser><c:xVal><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:xVal><c:yVal><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:yVal></c:ser>
                </c:scatterChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartPlot plot = chart.Plots[0];
        TestAssert.Equal(PptxSceneChartScatterStyle.Unknown, plot.ScatterStyleKind);
        TestAssert.True(plot.Series[0].Smooth is null, "Expected missing scatter smooth metadata to remain distinct from the effective renderer default.");

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:scatterChart>
                <c:scatterStyle val="lineMarker"/>
                <c:ser><c:smooth val="1"/></c:ser>
              </c:scatterChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "scatterChart").Single();
        System.Reflection.MethodInfo readOptions = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartScatterOptions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart scatter-option resolver.");
        object options = readOptions.Invoke(null, [plot, mismatchedXmlFallback]) ?? throw new InvalidOperationException("Expected resolved scatter options.");
        Type optionsType = options.GetType();
        object[] smoothSeries = (((System.Collections.IEnumerable?)optionsType.GetProperty("SmoothSeries")?.GetValue(options)) ?? throw new InvalidOperationException("Expected smooth series options.")).Cast<object>().ToArray();

        TestAssert.Equal(PptxSceneChartScatterStyle.Unknown, (PptxSceneChartScatterStyle)(optionsType.GetProperty("ScatterStyle")?.GetValue(options) ?? default(PptxSceneChartScatterStyle)));
        TestAssert.True((bool)(optionsType.GetProperty("ConnectLines")?.GetValue(options) ?? true) == false, "Expected unknown scene scatterStyle to keep the no-line default, not fallback XML lineMarker.");
        TestAssert.True(smoothSeries.Length == 1 && !PptxTests.ChartBooleanOptionValue(smoothSeries[0]), "Expected missing scene scatter smooth flag to use the scene default, not fallback XML smooth=1.");
        TestAssert.True(!PptxTests.ChartBooleanOptionIsDefined(smoothSeries[0]), "Expected missing scene scatter smooth state to remain distinguishable from an explicit false token.");
    }

    public static void PptxChartSmoothOptionsPreserveRawBooleanState()
    {
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement chartElement = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:scatterChart>
                <c:scatterStyle val="lineMarker"/>
                <c:ser><c:smooth/></c:ser>
                <c:ser><c:smooth val="0"/></c:ser>
                <c:ser/>
              </c:scatterChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "scatterChart").Single();
        System.Reflection.MethodInfo readOptions = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartScatterOptions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart scatter-option resolver.");
        object options = readOptions.Invoke(null, [null, chartElement]) ?? throw new InvalidOperationException("Expected resolved scatter options.");
        Type optionsType = options.GetType();
        object[] smoothSeries = (((System.Collections.IEnumerable?)optionsType.GetProperty("SmoothSeries")?.GetValue(options)) ?? throw new InvalidOperationException("Expected smooth series options.")).Cast<object>().ToArray();

        TestAssert.Equal(3, smoothSeries.Length);
        TestAssert.True(PptxTests.ChartBooleanOptionValue(smoothSeries[0]), "Expected shorthand smooth element to resolve through the shared OOXML boolean parser.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(smoothSeries[0]), "Expected shorthand smooth element presence to remain explicit.");
        TestAssert.Equal(string.Empty, PptxTests.ChartBooleanOptionRawValue(smoothSeries[0]));
        TestAssert.True(!PptxTests.ChartBooleanOptionValue(smoothSeries[1]), "Expected explicit smooth=0 to resolve false.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(smoothSeries[1]), "Expected explicit smooth=0 presence to remain explicit.");
        TestAssert.Equal("0", PptxTests.ChartBooleanOptionRawValue(smoothSeries[1]));
        TestAssert.True(!PptxTests.ChartBooleanOptionValue(smoothSeries[2]), "Expected missing smooth element to use the renderer default false.");
        TestAssert.True(!PptxTests.ChartBooleanOptionIsDefined(smoothSeries[2]), "Expected missing smooth element to remain distinct from explicit false.");
        TestAssert.Equal(string.Empty, PptxTests.ChartBooleanOptionRawValue(smoothSeries[2]));
    }

    public static void PptxChartVaryColorsOptionsPreserveRawBooleanState()
    {
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        System.Reflection.MethodInfo readOptions = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartBarOptions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart bar-option resolver.");

        object shorthand = ReadBarOptions("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:barChart><c:varyColors/></c:barChart></c:plotArea></c:chart>
            </c:chartSpace>
            """);
        object explicitFalse = ReadBarOptions("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:barChart><c:varyColors val="0"/></c:barChart></c:plotArea></c:chart>
            </c:chartSpace>
            """);
        object missing = ReadBarOptions("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:barChart/></c:plotArea></c:chart>
            </c:chartSpace>
            """);

        object shorthandVaryColors = shorthand.GetType().GetProperty("VaryColors")?.GetValue(shorthand) ?? throw new InvalidOperationException("Expected shorthand varyColors option.");
        TestAssert.True(PptxTests.ChartBooleanOptionValue(shorthandVaryColors), "Expected shorthand varyColors element to resolve through the shared OOXML boolean parser.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(shorthandVaryColors), "Expected shorthand varyColors element presence to remain explicit.");
        TestAssert.Equal(string.Empty, PptxTests.ChartBooleanOptionRawValue(shorthandVaryColors));

        object explicitFalseVaryColors = explicitFalse.GetType().GetProperty("VaryColors")?.GetValue(explicitFalse) ?? throw new InvalidOperationException("Expected explicit false varyColors option.");
        TestAssert.True(!PptxTests.ChartBooleanOptionValue(explicitFalseVaryColors), "Expected explicit varyColors=0 to resolve false.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(explicitFalseVaryColors), "Expected explicit varyColors=0 presence to remain explicit.");
        TestAssert.Equal("0", PptxTests.ChartBooleanOptionRawValue(explicitFalseVaryColors));

        object missingVaryColors = missing.GetType().GetProperty("VaryColors")?.GetValue(missing) ?? throw new InvalidOperationException("Expected missing varyColors option.");
        TestAssert.True(PptxTests.ChartBooleanOptionValue(missingVaryColors), "Expected missing varyColors element to use the renderer default true.");
        TestAssert.True(!PptxTests.ChartBooleanOptionIsDefined(missingVaryColors), "Expected missing varyColors element to remain distinct from explicit true.");
        TestAssert.Equal(string.Empty, PptxTests.ChartBooleanOptionRawValue(missingVaryColors));

        object ReadBarOptions(string xml)
        {
            XElement chartElement = XDocument.Parse(xml).Descendants(c + "barChart").Single();
            return readOptions.Invoke(null, [null, chartElement, PptxSceneChartGrouping.Clustered]) ?? throw new InvalidOperationException("Expected resolved bar options.");
        }
    }

    public static void PptxChartDataLabelOptionsPreserveRawBooleanState()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea><c:barChart>
                <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                <c:dLbls>
                  <c:txPr><a:bodyPr rot="3600000"/><a:lstStyle/><a:p/></c:txPr>
                  <c:leaderLines><c:spPr><a:ln w="38100"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:ln></c:spPr></c:leaderLines>
                  <c:showVal/>
                  <c:showPercent val="0"/>
                  <c:dLbl><c:idx val="0"/><c:txPr><a:bodyPr rot="-1200000"/><a:lstStyle/><a:p/></c:txPr><c:leaderLines><c:spPr><a:ln w="25400"><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:ln></c:spPr></c:leaderLines><c:showVal val="0"/><c:showLegendKey/></c:dLbl>
                </c:dLbls>
              </c:barChart></c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartPlot plot = chart.Plots[0];

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea><c:barChart>
                <c:dLbls><c:txPr><a:bodyPr rot="600000"/><a:lstStyle/><a:p/></c:txPr><c:leaderLines><c:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></a:ln></c:spPr></c:leaderLines><c:showVal val="0"/><c:showLegendKey val="1"/><c:dLbl><c:idx val="0"/><c:txPr><a:bodyPr rot="2400000"/><a:lstStyle/><a:p/></c:txPr><c:leaderLines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="FEDCBA"/></a:solidFill></a:ln></c:spPr></c:leaderLines><c:showVal/><c:showLegendKey val="0"/></c:dLbl></c:dLbls>
              </c:barChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "barChart").Single();
        System.Reflection.MethodInfo readOptions = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlDataLabelOptions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected data-label option resolver.");
        System.Reflection.MethodInfo resolveOptions = typeof(PptxRenderer).GetMethod(
            "ResolveChartDataLabelOptions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected data-label override resolver.");
        System.Reflection.MethodInfo readSeriesOptions = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlSeriesDataLabelOptions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected series data-label option resolver.");
        System.Reflection.MethodInfo resolveSeriesOptions = typeof(PptxRenderer).GetMethod(
            "ResolveChartDataLabelOptionsForSeries",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected series data-label merge resolver.");
        object sceneOptions = readOptions.Invoke(null, [chart, plot, mismatchedXmlFallback, PptxTheme.Empty, PptxColorMap.Default]) ?? throw new InvalidOperationException("Expected scene-backed label options.");

        object sceneShowValue = PptxTests.ChartDataLabelFlagOption(sceneOptions, "showVal");
        TestAssert.True(PptxTests.ChartBooleanOptionValue(sceneShowValue), "Expected scene-backed shorthand showVal to resolve true.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(sceneShowValue), "Expected scene-backed shorthand showVal presence to remain explicit.");
        TestAssert.Equal(string.Empty, PptxTests.ChartBooleanOptionRawValue(sceneShowValue));
        object sceneShowPercent = PptxTests.ChartDataLabelFlagOption(sceneOptions, "showPercent");
        TestAssert.True(!PptxTests.ChartBooleanOptionValue(sceneShowPercent), "Expected scene-backed showPercent=0 to resolve false.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(sceneShowPercent), "Expected scene-backed showPercent=0 presence to remain explicit.");
        TestAssert.Equal("0", PptxTests.ChartBooleanOptionRawValue(sceneShowPercent));
        object sceneShowLegendKey = PptxTests.ChartDataLabelFlagOption(sceneOptions, "showLegendKey");
        TestAssert.True(!PptxTests.ChartBooleanOptionValue(sceneShowLegendKey), "Expected missing scene showLegendKey to use the renderer default false, not fallback XML.");
        TestAssert.True(!PptxTests.ChartBooleanOptionIsDefined(sceneShowLegendKey), "Expected missing scene showLegendKey to remain distinct from explicit false.");
        object sceneLeaderLines = PptxTests.ChartDataLabelLeaderLines(sceneOptions);
        TestAssert.True(PptxTests.ChartDataLabelLeaderLinesIsDefined(sceneLeaderLines), "Expected scene-backed data-label leader-line source to remain defined.");
        TestAssert.Equal(new RgbColor(17, 34, 51), PptxTests.ChartSeriesStrokeColor(PptxTests.ChartDataLabelLeaderLinesStroke(sceneLeaderLines)));
        TestAssert.Equal(3d, PptxTests.ChartSeriesStrokeWidth(PptxTests.ChartDataLabelLeaderLinesStroke(sceneLeaderLines)));
        object sceneBody = PptxTests.ChartDataLabelTextBodyProperties(sceneOptions);
        TestAssert.Equal(60d, PptxTests.ChartTextBodyRotationDegrees(sceneBody) ?? 0d);
        TestAssert.Equal("3600000", PptxTests.ChartTextBodyRotationValue(sceneBody));
        object sceneOverride = PptxTests.ChartDataLabelOverride(sceneOptions, 0);
        object sceneOverrideShowValue = PptxTests.ChartDataLabelFlagOption(sceneOverride, "showVal");
        TestAssert.True(!PptxTests.ChartBooleanOptionValue(sceneOverrideShowValue), "Expected scene-backed override showVal=0 to resolve false.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(sceneOverrideShowValue), "Expected scene-backed override showVal=0 presence to remain explicit.");
        TestAssert.Equal("0", PptxTests.ChartBooleanOptionRawValue(sceneOverrideShowValue));
        object sceneOverrideShowLegendKey = PptxTests.ChartDataLabelFlagOption(sceneOverride, "showLegendKey");
        TestAssert.True(PptxTests.ChartBooleanOptionValue(sceneOverrideShowLegendKey), "Expected scene-backed override shorthand showLegendKey to resolve true.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(sceneOverrideShowLegendKey), "Expected scene-backed override shorthand showLegendKey presence to remain explicit.");
        object sceneOverrideLeaderLines = PptxTests.ChartDataLabelLeaderLines(sceneOverride);
        TestAssert.True(PptxTests.ChartDataLabelLeaderLinesIsDefined(sceneOverrideLeaderLines), "Expected scene-backed data-label override leader-line source to remain defined.");
        TestAssert.Equal(new RgbColor(68, 85, 102), PptxTests.ChartSeriesStrokeColor(PptxTests.ChartDataLabelLeaderLinesStroke(sceneOverrideLeaderLines)));
        TestAssert.Equal(2d, PptxTests.ChartSeriesStrokeWidth(PptxTests.ChartDataLabelLeaderLinesStroke(sceneOverrideLeaderLines)));
        object sceneOverrideBody = PptxTests.ChartDataLabelTextBodyProperties(sceneOverride);
        TestAssert.Equal(-20d, PptxTests.ChartTextBodyRotationDegrees(sceneOverrideBody) ?? 0d);
        TestAssert.Equal("-1200000", PptxTests.ChartTextBodyRotationValue(sceneOverrideBody));
        object resolvedSceneOptions = resolveOptions.Invoke(null, [sceneOptions, 0]) ?? throw new InvalidOperationException("Expected resolved scene-backed label options.");
        object resolvedSceneShowValue = PptxTests.ChartDataLabelFlagOption(resolvedSceneOptions, "showVal");
        TestAssert.True(!PptxTests.ChartBooleanOptionValue(resolvedSceneShowValue), "Expected resolved scene-backed override showVal=0 to replace base shorthand showVal.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(resolvedSceneShowValue), "Expected resolved scene-backed override showVal=0 presence to remain explicit.");
        TestAssert.Equal("0", PptxTests.ChartBooleanOptionRawValue(resolvedSceneShowValue));
        object resolvedSceneShowPercent = PptxTests.ChartDataLabelFlagOption(resolvedSceneOptions, "showPercent");
        TestAssert.True(!PptxTests.ChartBooleanOptionValue(resolvedSceneShowPercent), "Expected resolved scene-backed missing override showPercent to inherit base showPercent=0.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(resolvedSceneShowPercent), "Expected resolved scene-backed missing override showPercent to retain base explicit state.");
        object resolvedSceneLeaderLines = PptxTests.ChartDataLabelLeaderLines(resolvedSceneOptions);
        TestAssert.True(PptxTests.ChartDataLabelLeaderLinesIsDefined(resolvedSceneLeaderLines), "Expected resolved scene-backed leader-line source to remain defined.");
        TestAssert.Equal(new RgbColor(68, 85, 102), PptxTests.ChartSeriesStrokeColor(PptxTests.ChartDataLabelLeaderLinesStroke(resolvedSceneLeaderLines)));
        TestAssert.Equal(2d, PptxTests.ChartSeriesStrokeWidth(PptxTests.ChartDataLabelLeaderLinesStroke(resolvedSceneLeaderLines)));
        object resolvedSceneBody = PptxTests.ChartDataLabelTextBodyProperties(resolvedSceneOptions);
        TestAssert.Equal(-20d, PptxTests.ChartTextBodyRotationDegrees(resolvedSceneBody) ?? 0d);
        TestAssert.Equal("-1200000", PptxTests.ChartTextBodyRotationValue(resolvedSceneBody));

        object xmlOptions = readOptions.Invoke(null, [null, null, mismatchedXmlFallback, PptxTheme.Empty, PptxColorMap.Default]) ?? throw new InvalidOperationException("Expected XML-backed label options.");
        object xmlShowValue = PptxTests.ChartDataLabelFlagOption(xmlOptions, "showVal");
        TestAssert.True(!PptxTests.ChartBooleanOptionValue(xmlShowValue), "Expected XML showVal=0 to resolve false.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(xmlShowValue), "Expected XML showVal=0 presence to remain explicit.");
        TestAssert.Equal("0", PptxTests.ChartBooleanOptionRawValue(xmlShowValue));
        object xmlShowLegendKey = PptxTests.ChartDataLabelFlagOption(xmlOptions, "showLegendKey");
        TestAssert.True(PptxTests.ChartBooleanOptionValue(xmlShowLegendKey), "Expected XML showLegendKey=1 to resolve true.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(xmlShowLegendKey), "Expected XML showLegendKey=1 presence to remain explicit.");
        TestAssert.Equal("1", PptxTests.ChartBooleanOptionRawValue(xmlShowLegendKey));
        object xmlLeaderLines = PptxTests.ChartDataLabelLeaderLines(xmlOptions);
        TestAssert.True(PptxTests.ChartDataLabelLeaderLinesIsDefined(xmlLeaderLines), "Expected XML-backed data-label leader-line source to remain defined.");
        TestAssert.Equal(new RgbColor(171, 205, 239), PptxTests.ChartSeriesStrokeColor(PptxTests.ChartDataLabelLeaderLinesStroke(xmlLeaderLines)));
        TestAssert.Equal(1d, PptxTests.ChartSeriesStrokeWidth(PptxTests.ChartDataLabelLeaderLinesStroke(xmlLeaderLines)));
        object xmlBody = PptxTests.ChartDataLabelTextBodyProperties(xmlOptions);
        TestAssert.Equal(10d, PptxTests.ChartTextBodyRotationDegrees(xmlBody) ?? 0d);
        TestAssert.Equal("600000", PptxTests.ChartTextBodyRotationValue(xmlBody));
        object xmlOverride = PptxTests.ChartDataLabelOverride(xmlOptions, 0);
        object xmlOverrideShowValue = PptxTests.ChartDataLabelFlagOption(xmlOverride, "showVal");
        TestAssert.True(PptxTests.ChartBooleanOptionValue(xmlOverrideShowValue), "Expected XML override shorthand showVal to resolve true.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(xmlOverrideShowValue), "Expected XML override shorthand showVal presence to remain explicit.");
        object xmlOverrideShowLegendKey = PptxTests.ChartDataLabelFlagOption(xmlOverride, "showLegendKey");
        TestAssert.True(!PptxTests.ChartBooleanOptionValue(xmlOverrideShowLegendKey), "Expected XML override showLegendKey=0 to resolve false.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(xmlOverrideShowLegendKey), "Expected XML override showLegendKey=0 presence to remain explicit.");
        TestAssert.Equal("0", PptxTests.ChartBooleanOptionRawValue(xmlOverrideShowLegendKey));
        object xmlOverrideLeaderLines = PptxTests.ChartDataLabelLeaderLines(xmlOverride);
        TestAssert.True(PptxTests.ChartDataLabelLeaderLinesIsDefined(xmlOverrideLeaderLines), "Expected XML-backed data-label override leader-line source to remain defined.");
        TestAssert.Equal(new RgbColor(254, 220, 186), PptxTests.ChartSeriesStrokeColor(PptxTests.ChartDataLabelLeaderLinesStroke(xmlOverrideLeaderLines)));
        TestAssert.Equal(0.5d, PptxTests.ChartSeriesStrokeWidth(PptxTests.ChartDataLabelLeaderLinesStroke(xmlOverrideLeaderLines)));
        object xmlOverrideBody = PptxTests.ChartDataLabelTextBodyProperties(xmlOverride);
        TestAssert.Equal(40d, PptxTests.ChartTextBodyRotationDegrees(xmlOverrideBody) ?? 0d);
        TestAssert.Equal("2400000", PptxTests.ChartTextBodyRotationValue(xmlOverrideBody));
        object resolvedXmlOptions = resolveOptions.Invoke(null, [xmlOptions, 0]) ?? throw new InvalidOperationException("Expected resolved XML-backed label options.");
        object resolvedXmlShowLegendKey = PptxTests.ChartDataLabelFlagOption(resolvedXmlOptions, "showLegendKey");
        TestAssert.True(!PptxTests.ChartBooleanOptionValue(resolvedXmlShowLegendKey), "Expected resolved XML override showLegendKey=0 to replace base showLegendKey=1.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(resolvedXmlShowLegendKey), "Expected resolved XML override showLegendKey=0 presence to remain explicit.");
        TestAssert.Equal("0", PptxTests.ChartBooleanOptionRawValue(resolvedXmlShowLegendKey));
        object resolvedXmlLeaderLines = PptxTests.ChartDataLabelLeaderLines(resolvedXmlOptions);
        TestAssert.True(PptxTests.ChartDataLabelLeaderLinesIsDefined(resolvedXmlLeaderLines), "Expected resolved XML-backed leader-line source to remain defined.");
        TestAssert.Equal(new RgbColor(254, 220, 186), PptxTests.ChartSeriesStrokeColor(PptxTests.ChartDataLabelLeaderLinesStroke(resolvedXmlLeaderLines)));
        TestAssert.Equal(0.5d, PptxTests.ChartSeriesStrokeWidth(PptxTests.ChartDataLabelLeaderLinesStroke(resolvedXmlLeaderLines)));
        object resolvedXmlBody = PptxTests.ChartDataLabelTextBodyProperties(resolvedXmlOptions);
        TestAssert.Equal(40d, PptxTests.ChartTextBodyRotationDegrees(resolvedXmlBody) ?? 0d);
        TestAssert.Equal("2400000", PptxTests.ChartTextBodyRotationValue(resolvedXmlBody));

        XElement xmlWithSeriesLabels = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:barChart>
                <c:dLbls><c:showVal/><c:showPercent val="0"/><c:separator>, </c:separator></c:dLbls>
                <c:ser>
                  <c:dLbls><c:showLegendKey/><c:separator>; </c:separator></c:dLbls>
                  <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val>
                </c:ser>
              </c:barChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "barChart").Single();
        object plotOptions = readOptions.Invoke(null, [null, null, xmlWithSeriesLabels, PptxTheme.Empty, PptxColorMap.Default]) ?? throw new InvalidOperationException("Expected plot label options.");
        object seriesOptions = readSeriesOptions.Invoke(null, [null, null, xmlWithSeriesLabels, PptxTheme.Empty, PptxColorMap.Default]) ?? throw new InvalidOperationException("Expected series label options.");
        object resolvedSeriesOptions = resolveSeriesOptions.Invoke(null, [plotOptions, seriesOptions, 0]) ?? throw new InvalidOperationException("Expected resolved series label options.");
        object seriesShowValue = PptxTests.ChartDataLabelFlagOption(resolvedSeriesOptions, "showVal");
        TestAssert.True(PptxTests.ChartBooleanOptionValue(seriesShowValue), "Expected missing series showVal to inherit the plot-level explicit showVal.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(seriesShowValue), "Expected inherited plot-level showVal presence to remain explicit.");
        object seriesShowPercent = PptxTests.ChartDataLabelFlagOption(resolvedSeriesOptions, "showPercent");
        TestAssert.True(!PptxTests.ChartBooleanOptionValue(seriesShowPercent), "Expected missing series showPercent to inherit the plot-level explicit false.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(seriesShowPercent), "Expected inherited plot-level showPercent presence to remain explicit.");
        object seriesShowLegendKey = PptxTests.ChartDataLabelFlagOption(resolvedSeriesOptions, "showLegendKey");
        TestAssert.True(PptxTests.ChartBooleanOptionValue(seriesShowLegendKey), "Expected series-level showLegendKey to override the plot-level default.");
        TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(seriesShowLegendKey), "Expected series-level showLegendKey presence to remain explicit.");
        TestAssert.Equal("; ", (string?)resolvedSeriesOptions.GetType().GetProperty("Separator")?.GetValue(resolvedSeriesOptions) ?? string.Empty);
    }

    public static void PptxChartRadarOptionsUseSceneAuthoritativeDefaults()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:radarChart>
                  <c:radarStyle val="bogus"/>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt><c:pt idx="2"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:radarChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartPlot plot = chart.Plots[0];
        TestAssert.Equal(PptxSceneChartRadarStyle.Unknown, plot.RadarStyleKind);

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:radarChart><c:radarStyle val="filled"/></c:radarChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "radarChart").Single();
        System.Reflection.MethodInfo readOptions = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartRadarOptions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart radar-option resolver.");
        object options = readOptions.Invoke(null, [plot, mismatchedXmlFallback]) ?? throw new InvalidOperationException("Expected resolved radar options.");
        Type optionsType = options.GetType();

        TestAssert.Equal(PptxSceneChartRadarStyle.Standard, (PptxSceneChartRadarStyle)(optionsType.GetProperty("RadarStyle")?.GetValue(options) ?? default(PptxSceneChartRadarStyle)));
    }

    public static void PptxChartPolarPointOptionsUseSceneAuthoritativeDefaults()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:pieChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:pieChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartPlot plot = chart.Plots[0];
        TestAssert.True(plot.FirstSliceAngle is null, "Expected missing firstSliceAng metadata to remain distinct from the effective renderer default.");
        TestAssert.True(plot.Series[0].PointStyles.Count == 0, "Expected no scene point styles.");

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:pieChart>
                <c:ser><c:dPt><c:idx val="0"/><c:explosion val="25"/></c:dPt></c:ser>
                <c:firstSliceAng val="270"/>
              </c:pieChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "pieChart").Single();
        System.Reflection.MethodInfo readOptions = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartPolarPointOptions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart polar-point option resolver.");
        object options = readOptions.Invoke(null, [plot, mismatchedXmlFallback, PptxTheme.Empty, PptxColorMap.Default, null]) ?? throw new InvalidOperationException("Expected resolved polar point options.");
        Type optionsType = options.GetType();
        object pointExplosions = optionsType.GetProperty("PointExplosions")?.GetValue(options) ?? throw new InvalidOperationException("Expected point explosion options.");

        TestAssert.True(((System.Collections.IEnumerable)pointExplosions).Cast<object>().Count() == 0, "Expected missing scene point explosions to stay empty, not import fallback XML dPt explosion.");
        TestAssert.Equal(0d, (double)(optionsType.GetProperty("FirstSliceAngle")?.GetValue(options) ?? double.NaN));
    }

    public static void PptxChartDoughnutOptionsUseSceneAuthoritativeDefaults()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:doughnutChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:doughnutChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartPlot plot = chart.Plots[0];
        TestAssert.True(plot.HoleSize is null, "Expected missing holeSize metadata to remain distinct from the effective renderer default.");

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:doughnutChart>
                <c:holeSize val="10"/>
                <c:firstSliceAng val="270"/>
              </c:doughnutChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "doughnutChart").Single();
        System.Reflection.MethodInfo readOptions = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartDoughnutOptions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart doughnut option resolver.");
        object options = readOptions.Invoke(null, [plot, mismatchedXmlFallback, PptxTheme.Empty, PptxColorMap.Default, null]) ?? throw new InvalidOperationException("Expected resolved doughnut options.");
        Type optionsType = options.GetType();
        object polarPoints = optionsType.GetProperty("PolarPoints")?.GetValue(options) ?? throw new InvalidOperationException("Expected polar point options.");

        TestAssert.Equal(0.56d, (double)(optionsType.GetProperty("HoleSize")?.GetValue(options) ?? double.NaN));
        TestAssert.Equal(0d, (double)(polarPoints.GetType().GetProperty("FirstSliceAngle")?.GetValue(polarPoints) ?? double.NaN));
    }

    public static void PptxChartUnknownScatterStyleKeepsSceneLineConnectionDefault()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:scatterChart>
                  <c:scatterStyle val="bogus"/>
                  <c:ser><c:xVal><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:xVal><c:yVal><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:yVal></c:ser>
                </c:scatterChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        TestAssert.Equal(PptxSceneChartScatterStyle.Unknown, chart.Plots[0].ScatterStyleKind);
        TestAssert.Equal("bogus", chart.Plots[0].ScatterStyle);

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:scatterChart><c:scatterStyle val="lineMarker"/></c:scatterChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "scatterChart").Single();
        System.Reflection.MethodInfo readScatterStyle = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartScatterStyle",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart scatter-style resolver.");
        object scatterStyle = readScatterStyle.Invoke(null, [chart.Plots[0], mismatchedXmlFallback]) ?? throw new InvalidOperationException("Expected resolved scatter style.");
        System.Reflection.MethodInfo resolveLineConnection = typeof(PptxRenderer).GetMethod(
            "ResolveChartScatterLineConnection",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected scatter line-connection resolver.");
        object connectLines = resolveLineConnection.Invoke(null, [scatterStyle]) ?? throw new InvalidOperationException("Expected resolved scatter line-connection state.");

        TestAssert.Equal(PptxSceneChartScatterStyle.Unknown, (PptxSceneChartScatterStyle)scatterStyle);
        TestAssert.True((bool)connectLines == false, "Expected unknown scatterStyle to keep the scene-owned no-line default instead of borrowing XML fallback lineMarker.");
    }

    public static void PptxChartUnknownRadarStyleUsesSceneAuthoritativeDefault()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:radarChart>
                  <c:radarStyle val="bogus"/>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt><c:pt idx="2"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:radarChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        TestAssert.Equal(PptxSceneChartRadarStyle.Unknown, chart.Plots[0].RadarStyleKind);
        TestAssert.Equal("bogus", chart.Plots[0].RadarStyle);

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:radarChart><c:radarStyle val="filled"/></c:radarChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "radarChart").Single();
        System.Reflection.MethodInfo readRadarStyle = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartRadarStyle",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart radar-style resolver.");
        object radarStyle = readRadarStyle.Invoke(null, [chart.Plots[0], mismatchedXmlFallback]) ?? throw new InvalidOperationException("Expected resolved radar style.");

        TestAssert.Equal(PptxSceneChartRadarStyle.Standard, (PptxSceneChartRadarStyle)radarStyle);
    }

    public static void PptxChartUnknownAxisOrientationUsesSceneAuthoritativeDefault()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:lineChart>
                <c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="20"/></c:catAx>
                <c:valAx><c:axId val="20"/><c:scaling><c:orientation val="bogus"/></c:scaling><c:axPos val="l"/><c:crossAx val="10"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartAxis valueAxis = chart.Axes.First(axis => axis.AxisKind == PptxSceneChartAxisKind.Value);
        TestAssert.Equal(PptxSceneChartAxisOrientation.Unknown, valueAxis.OrientationKind);
        TestAssert.Equal("bogus", valueAxis.Orientation);

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:valAx><c:scaling><c:orientation val="maxMin"/></c:scaling></c:valAx></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "valAx").Single();
        System.Reflection.MethodInfo readAxisReversed = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlValueAxisReversed",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected value-axis orientation resolver.");
        object reversed = readAxisReversed.Invoke(null, [valueAxis, mismatchedXmlFallback]) ?? throw new InvalidOperationException("Expected resolved axis orientation.");

        TestAssert.True(!(bool)reversed, "Unknown scene-owned value-axis orientation should resolve to the Office default minMax, not fall back to mismatched XML.");
    }

    public static void PptxChartUnknownAxisCrossesUsesSceneAuthoritativeDefault()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:lineChart>
                <c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="20"/></c:catAx>
                <c:valAx><c:axId val="20"/><c:scaling><c:min val="-5"/><c:max val="5"/></c:scaling><c:axPos val="l"/><c:crossAx val="10"/><c:crosses val="bogus"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartAxis valueAxis = chart.Axes.First(axis => axis.AxisKind == PptxSceneChartAxisKind.Value);
        TestAssert.Equal(PptxSceneChartAxisCrosses.Unknown, valueAxis.CrossesKind);
        TestAssert.Equal("bogus", valueAxis.Crosses);

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:valAx><c:crosses val="max"/></c:valAx></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "valAx").Single();
        Type extentsType = typeof(PptxRenderer).GetNestedType(
            "ChartValueExtents",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart value extents type.");
        object extents = Activator.CreateInstance(extentsType, [-5d, 5d]) ?? throw new InvalidOperationException("Expected chart value extents.");
        System.Reflection.MethodInfo readCrossingValue = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlValueAxisCrossingValue",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected value-axis crossing resolver.");
        object crossingValue = readCrossingValue.Invoke(null, [valueAxis, mismatchedXmlFallback, extents]) ?? throw new InvalidOperationException("Expected resolved axis crossing.");

        TestAssert.Equal(0d, (double)crossingValue);
    }

    public static void PptxChartInvalidAxisCrossesAtUsesSceneAuthoritativeCrosses()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:lineChart>
                <c:catAx><c:axId val="10"/><c:axPos val="b"/><c:crossAx val="20"/></c:catAx>
                <c:valAx><c:axId val="20"/><c:scaling><c:min val="-5"/><c:max val="5"/></c:scaling><c:axPos val="l"/><c:crossAx val="10"/><c:crosses val="max"/><c:crossesAt val="futureCross"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartAxis valueAxis = chart.Axes.First(axis => axis.AxisKind == PptxSceneChartAxisKind.Value);
        TestAssert.Equal(PptxSceneChartAxisCrosses.Maximum, valueAxis.CrossesKind);
        TestAssert.True(valueAxis.CrossesAt is null, "Expected unparseable scene crossesAt to stay distinct from an authored value.");
        TestAssert.Equal("futureCross", valueAxis.CrossesAtValue);

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:valAx><c:crossesAt val="2"/></c:valAx></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "valAx").Single();
        Type extentsType = typeof(PptxRenderer).GetNestedType(
            "ChartValueExtents",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart value extents type.");
        object extents = Activator.CreateInstance(extentsType, [-5d, 5d]) ?? throw new InvalidOperationException("Expected chart value extents.");
        System.Reflection.MethodInfo readCrossingValue = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlValueAxisCrossingValue",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected value-axis crossing resolver.");
        object sceneCrossingValue = readCrossingValue.Invoke(null, [valueAxis, mismatchedXmlFallback, extents]) ?? throw new InvalidOperationException("Expected resolved scene axis crossing.");
        object xmlOnlyCrossingValue = readCrossingValue.Invoke(null, [null, mismatchedXmlFallback, extents]) ?? throw new InvalidOperationException("Expected resolved XML-only axis crossing.");

        TestAssert.Equal(5d, (double)sceneCrossingValue);
        TestAssert.Equal(2d, (double)xmlOnlyCrossingValue);
    }

    public static void PptxChartUnknownAxisCrossBetweenUsesSceneAuthoritativeDefault()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:lineChart>
                <c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="20"/></c:catAx>
                <c:valAx><c:axId val="20"/><c:scaling><c:min val="0"/><c:max val="5"/></c:scaling><c:axPos val="l"/><c:crossAx val="10"/><c:crossBetween val="bogus"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartAxis valueAxis = chart.Axes.First(axis => axis.AxisKind == PptxSceneChartAxisKind.Value);
        TestAssert.Equal(PptxSceneChartAxisCrossBetween.Unknown, valueAxis.CrossBetweenKind);
        TestAssert.Equal("bogus", valueAxis.CrossBetween);

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:valAx><c:crossBetween val="midCat"/></c:valAx></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "valAx").Single();
        System.Reflection.MethodInfo labelsOnTickMarks = typeof(PptxRenderer).GetMethod(
            "ResolveSceneOrXmlCategoryAxisLabelsOnTickMarks",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected category-axis label placement resolver.");
        object onTickMarks = labelsOnTickMarks.Invoke(null, [valueAxis, mismatchedXmlFallback]) ?? throw new InvalidOperationException("Expected resolved category-axis label placement.");

        TestAssert.True(!(bool)onTickMarks, "Unknown scene-owned crossBetween should resolve to the Office default between, not fall back to mismatched XML.");
    }

    public static void PptxChartUnknownTickLabelPositionUsesSceneAuthoritativeDefault()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:lineChart>
                <c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="20"/><c:tickLblPos val="bogus"/></c:catAx>
                <c:valAx><c:axId val="20"/><c:scaling><c:min val="0"/><c:max val="5"/></c:scaling><c:axPos val="l"/><c:crossAx val="10"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartAxis categoryAxis = chart.Axes.First(axis => axis.AxisKind == PptxSceneChartAxisKind.Category);
        TestAssert.Equal(PptxSceneChartTickLabelPosition.Unknown, categoryAxis.TickLabelPositionKind);
        TestAssert.Equal("bogus", categoryAxis.TickLabelPosition);

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:catAx><c:tickLblPos val="none"/></c:catAx></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "catAx").Single();
        System.Reflection.MethodInfo labelVisible = typeof(PptxRenderer).GetMethod(
            "IsSceneOrXmlChartAxisLabelVisible",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected axis label visibility resolver.");
        object visible = labelVisible.Invoke(null, [categoryAxis, mismatchedXmlFallback]) ?? throw new InvalidOperationException("Expected resolved axis label visibility.");

        TestAssert.True((bool)visible, "Unknown scene-owned tickLblPos should resolve to the Office default nextTo/visible, not fall back to mismatched XML.");
    }

    public static void PptxChartMissingMajorTickMarkResolvesThroughExplicitDefault()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:lineChart>
                <c:catAx><c:axId val="10"/><c:axPos val="b"/><c:crossAx val="20"/></c:catAx>
                <c:valAx><c:axId val="20"/><c:axPos val="l"/><c:crossAx val="10"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartAxis categoryAxis = chart.Axes.First(axis => axis.AxisKind == PptxSceneChartAxisKind.Category);
        TestAssert.Equal(PptxSceneChartAxisTickMark.Unknown, categoryAxis.MajorTickMarkKind);
        TestAssert.Equal(string.Empty, categoryAxis.MajorTickMark);

        System.Reflection.MethodInfo readMajorTickMark = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartAxisMajorTickMark",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart major tick-mark resolver.");
        object tickMark = readMajorTickMark.Invoke(null, [categoryAxis, null]) ?? throw new InvalidOperationException("Expected resolved major tick mark.");

        TestAssert.Equal(PptxSceneChartAxisTickMark.None, (PptxSceneChartAxisTickMark)tickMark);
    }

    public static void PptxChartValueAxisRenderOptionsUseSceneAuthoritativeDefaults()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:lineChart>
                <c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="20"/></c:catAx>
                <c:valAx><c:axId val="20"/><c:scaling><c:min val="-5"/><c:max val="5"/><c:orientation val="bogus"/></c:scaling><c:axPos val="l"/><c:crossAx val="10"/><c:crosses val="bogus"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartAxis valueAxis = chart.Axes.First(axis => axis.AxisKind == PptxSceneChartAxisKind.Value);

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea><c:valAx>
                <c:scaling><c:orientation val="maxMin"/></c:scaling>
                <c:crosses val="max"/>
                <c:majorUnit val="0.25"/>
                <c:majorGridlines><c:spPr><a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>
              </c:valAx></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "valAx").Single();
        Type extentsType = typeof(PptxRenderer).GetNestedType(
            "ChartValueExtents",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart value extents type.");
        object extents = Activator.CreateInstance(extentsType, [-5d, 5d]) ?? throw new InvalidOperationException("Expected chart value extents.");
        System.Reflection.MethodInfo readOptions = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartValueAxisRenderOptions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart value-axis option resolver.");
        object options = readOptions.Invoke(null, [valueAxis, mismatchedXmlFallback, PptxTheme.Empty, extents, true]) ?? throw new InvalidOperationException("Expected value-axis render options.");
        Type optionsType = options.GetType();

        object units = optionsType.GetProperty("Units")?.GetValue(options) ?? throw new InvalidOperationException("Expected axis units.");
        TestAssert.Equal(0.1d, (double)(units.GetType().GetProperty("MajorUnit")?.GetValue(units) ?? double.NaN));
        TestAssert.True((bool)(optionsType.GetProperty("Reversed")?.GetValue(options) ?? true) == false, "Expected unknown scene orientation to use the default minMax, not fallback XML maxMin.");
        TestAssert.Equal(0d, (double)(optionsType.GetProperty("CrossingValue")?.GetValue(options) ?? double.NaN));
        TestAssert.True((bool)(optionsType.GetProperty("MajorGridlines")?.GetValue(options) ?? true) == false, "Expected missing scene gridlines to stay missing, not import fallback XML gridlines.");
    }

    public static void PptxChartAxisNumberFormatUsesSceneAuthoritativeAbsence()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>1.25</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:lineChart>
                <c:catAx><c:axId val="10"/><c:axPos val="b"/><c:crossAx val="20"/></c:catAx>
                <c:valAx><c:axId val="20"/><c:axPos val="l"/><c:crossAx val="10"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartAxis valueAxis = chart.Axes.First(axis => axis.AxisKind == PptxSceneChartAxisKind.Value);
        TestAssert.True(valueAxis.NumberFormatInfo.IsDefined == false, "Expected scene value axis to carry an authoritative missing number format.");

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:valAx><c:axId val="20"/><c:numFmt formatCode="0%" sourceLinked="0"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "valAx").Single();
        System.Reflection.MethodInfo formatAxisLabel = typeof(PptxRenderer).GetMethod(
            "FormatSceneOrXmlChartAxisLabel",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart axis label formatter.");

        string sceneLabel = (string?)formatAxisLabel.Invoke(null, [1.25d, valueAxis, mismatchedXmlFallback, null]) ?? string.Empty;
        string xmlOnlyLabel = (string?)formatAxisLabel.Invoke(null, [1.25d, null, mismatchedXmlFallback, null]) ?? string.Empty;

        TestAssert.Equal("1.25", sceneLabel);
        TestAssert.Equal("125%", xmlOnlyLabel);
    }

    public static void PptxChartAxisTextStyleUsesSceneAuthoritativeMissingAxis()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:lineChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XDocument mismatchedXml = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea>
                <c:valAx>
                  <c:txPr><a:p><a:pPr><a:defRPr sz="4000" b="1"/></a:pPr></a:p></c:txPr>
                </c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);
        XElement mismatchedAxis = mismatchedXml.Descendants(c + "valAx").Single();

        System.Reflection.MethodInfo readStyle = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartTextStyle",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart text-style resolver.");

        object sceneStyle = readStyle.Invoke(null, [PptxTheme.Empty, chart, null, mismatchedXml, mismatchedAxis, 10d, null]) ?? throw new InvalidOperationException("Expected scene text style.");
        object xmlStyle = readStyle.Invoke(null, [PptxTheme.Empty, null, null, mismatchedXml, mismatchedAxis, 10d, null]) ?? throw new InvalidOperationException("Expected XML text style.");

        TestAssert.Equal(10d, (double)(sceneStyle.GetType().GetProperty("FontSize")?.GetValue(sceneStyle) ?? double.NaN));
        TestAssert.True((bool)(sceneStyle.GetType().GetProperty("Bold")?.GetValue(sceneStyle) ?? true) == false, "Expected missing scene axis to keep scene defaults, not import fallback XML text properties.");
        TestAssert.Equal(40d, (double)(xmlStyle.GetType().GetProperty("FontSize")?.GetValue(xmlStyle) ?? double.NaN));
        TestAssert.True((bool)(xmlStyle.GetType().GetProperty("Bold")?.GetValue(xmlStyle) ?? false), "Expected XML-only compatibility to keep reading axis text properties.");
    }

    public static void PptxChartTextStyleCarriesTypefaceSourceThroughRendererStyle()
    {
        const string chartXml = """
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:txPr><a:p><a:pPr><a:defRPr sz="1100" spc="125"><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr>
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:lineChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """;
        PptxSceneChart chart = PptxTests.BuildSingleChartScene(chartXml) ?? throw new InvalidOperationException("Expected chart scene.");
        XDocument document = XDocument.Parse(chartXml);
        System.Reflection.MethodInfo readStyle = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartTextStyle",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart text-style resolver.");

        object style = readStyle.Invoke(null, [PptxTheme.Empty, chart, null, document, document.Root, 10d, null]) ?? throw new InvalidOperationException("Expected chart text style.");

        TestAssert.Equal("Arial", (string?)style.GetType().GetProperty("FontFamily")?.GetValue(style) ?? string.Empty);
        TestAssert.Equal(1.25d, (double)(style.GetType().GetProperty("CharacterSpacing")?.GetValue(style) ?? double.NaN));
        TestAssert.Equal("Arial", (string?)style.GetType().GetProperty("RequestedTypeface")?.GetValue(style) ?? string.Empty);
        TestAssert.Equal(PptxThemeTypefaceSource.Direct, (PptxThemeTypefaceSource?)style.GetType().GetProperty("TypefaceSource")?.GetValue(style) ?? default);
    }

    public static void PptxChartUnknownDisplayBlanksAsUsesSceneAuthoritativeDefault()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:lineChart>
              </c:plotArea><c:dispBlanksAs val="bogus"/></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        TestAssert.Equal(PptxSceneChartDisplayBlanksAs.Unknown, chart.Options.DisplayBlanksAsKind);
        TestAssert.Equal("bogus", chart.Options.DisplayBlanksAs);

        XDocument mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:dispBlanksAs val="span"/></c:chart>
            </c:chartSpace>
            """);
        System.Reflection.MethodInfo readDisplayBlanksAs = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartDisplayBlanksAs",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected display-blanks resolver.");
        object displayBlanksAs = readDisplayBlanksAs.Invoke(null, [chart, mismatchedXmlFallback]) ?? throw new InvalidOperationException("Expected resolved display-blanks value.");

        TestAssert.Equal(PptxSceneChartDisplayBlanksAs.Gap, (PptxSceneChartDisplayBlanksAs)displayBlanksAs);
    }

    public static void PptxChartTitlePolarDetectionUsesScenePlotKinds()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:pieChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:pieChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");

        XDocument mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);
        System.Reflection.MethodInfo hasPolarChart = typeof(PptxRenderer).GetMethod(
            "HasSceneOrXmlPolarChart",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected polar chart resolver.");
        object sceneBacked = hasPolarChart.Invoke(null, [chart, mismatchedXmlFallback]) ?? throw new InvalidOperationException("Expected scene-backed polar decision.");
        object xmlBacked = hasPolarChart.Invoke(null, [null, mismatchedXmlFallback]) ?? throw new InvalidOperationException("Expected XML-backed polar decision.");

        TestAssert.True((bool)sceneBacked, "Expected scene-backed title placement to use scene-owned polar plot kinds, not fallback XML.");
        TestAssert.True(!(bool)xmlBacked, "Expected XML-only title placement to keep the existing XML fallback behavior.");
    }

    public static void PptxChartPlotElementSelectionUsesSceneSource()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:barChart>
                  <c:barDir val="bar"/>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:barChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");

        XDocument mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);
        System.Reflection.MethodInfo readPlotElement = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlFirstChartPlotElement",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected plot element resolver.");
        object sceneBacked = readPlotElement.Invoke(null, [chart, mismatchedXmlFallback, PptxSceneChartPlotKind.Bar]) ?? throw new InvalidOperationException("Expected scene-backed bar plot.");
        object? xmlBacked = readPlotElement.Invoke(null, [null, mismatchedXmlFallback, PptxSceneChartPlotKind.Bar]);

        TestAssert.Equal("barChart", ((XElement)sceneBacked).Name.LocalName);
        TestAssert.True(xmlBacked is null, "Expected XML-only fallback to report the missing bar plot instead of using scene data.");
    }

    public static void PptxChartPlotElementsSelectionUsesSceneSources()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:barChart>
                  <c:barDir val="col"/>
                  <c:axId val="10"/><c:axId val="20"/>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:barChart>
                <c:barChart>
                  <c:barDir val="bar"/>
                  <c:axId val="10"/><c:axId val="30"/>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>B</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:barChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");

        XDocument mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);
        System.Reflection.MethodInfo readPlotElements = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartPlotElements",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected plot elements resolver.");
        object sceneBacked = readPlotElements.Invoke(null, [chart, mismatchedXmlFallback, PptxSceneChartPlotKind.Bar]) ?? throw new InvalidOperationException("Expected scene-backed bar plots.");
        object xmlBacked = readPlotElements.Invoke(null, [null, mismatchedXmlFallback, PptxSceneChartPlotKind.Bar]) ?? throw new InvalidOperationException("Expected XML-backed plot list.");

        IReadOnlyList<XElement> scenePlots = (IReadOnlyList<XElement>)sceneBacked;
        IReadOnlyList<XElement> xmlPlots = (IReadOnlyList<XElement>)xmlBacked;
        XNamespace chartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        TestAssert.Equal(2, scenePlots.Count);
        TestAssert.Equal("col", scenePlots[0].Element(chartNamespace + "barDir")?.Attribute("val")?.Value ?? string.Empty);
        TestAssert.Equal("bar", scenePlots[1].Element(chartNamespace + "barDir")?.Attribute("val")?.Value ?? string.Empty);
        TestAssert.Equal(0, xmlPlots.Count);
    }

    public static void PptxChartSecondaryRightValueAxisSelectionUsesSceneSource()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:barChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:barChart>
                <c:barChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>B</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>30</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="10"/><c:axId val="30"/>
                </c:barChart>
                <c:catAx><c:axId val="10"/><c:axPos val="b"/><c:crossAx val="20"/></c:catAx>
                <c:valAx><c:axId val="20"/><c:axPos val="l"/><c:crossAx val="10"/></c:valAx>
                <c:valAx><c:axId val="30"/><c:axPos val="r"/><c:crossAx val="10"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");

        XDocument mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:valAx><c:axId val="20"/><c:axPos val="r"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);
        System.Reflection.MethodInfo readSecondaryAxis = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlSecondaryRightValueAxis",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected secondary value-axis resolver.");
        object sceneBacked = readSecondaryAxis.Invoke(null, [chart, mismatchedXmlFallback]) ?? throw new InvalidOperationException("Expected scene-backed secondary value axis.");
        object xmlBacked = readSecondaryAxis.Invoke(null, [null, mismatchedXmlFallback]) ?? throw new InvalidOperationException("Expected XML-backed secondary value axis.");

        PptxSceneChartAxis? sceneAxis = (PptxSceneChartAxis?)sceneBacked.GetType().GetProperty("SceneAxis")?.GetValue(sceneBacked);
        XElement? sceneXmlAxis = (XElement?)sceneBacked.GetType().GetProperty("XmlAxis")?.GetValue(sceneBacked);
        PptxSceneChartAxis? xmlSceneAxis = (PptxSceneChartAxis?)xmlBacked.GetType().GetProperty("SceneAxis")?.GetValue(xmlBacked);
        XElement? xmlAxis = (XElement?)xmlBacked.GetType().GetProperty("XmlAxis")?.GetValue(xmlBacked);

        XNamespace chartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        TestAssert.Equal("30", sceneAxis?.Id ?? string.Empty);
        TestAssert.Equal("30", sceneXmlAxis?.Element(chartNamespace + "axId")?.Attribute("val")?.Value ?? string.Empty);
        TestAssert.True(xmlSceneAxis is null, "Expected XML-only secondary axis selection to remain XML-backed.");
        TestAssert.Equal("20", xmlAxis?.Element(chartNamespace + "axId")?.Attribute("val")?.Value ?? string.Empty);
    }

    public static void PptxChartRotatedTextClipMapsIntoTransformedSpace()
    {
        Type textRunType = typeof(PptxRenderer).GetNestedType(
            "TextRun",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart text run.");
        Type alignmentType = typeof(PptxRenderer).GetNestedType(
            "TextAlignment",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected text alignment.");
        System.Reflection.MethodInfo inverseClip = typeof(PptxRenderer).GetMethod(
            "InverseTransformClip",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected rotated-clip bridge.");
        object center = Enum.ToObject(alignmentType, 1);
        object? plainRun = Activator.CreateInstance(textRunType, ["", 0d, 0d, 0d, 0d, 0d, 0d, 0d, 0d, 12d, 0d, 0d, new RgbColor(0, 0, 0), 1d, null, false, false, false, false, true, center, null, 0d, 0d, 0d, false, false, false, null, false, 0]);
        object? rotatedRun = Activator.CreateInstance(textRunType, ["", 0d, 0d, 0d, 0d, 0d, 0d, 0d, 0d, 12d, 0d, 0d, new RgbColor(0, 0, 0), 1d, null, false, false, false, false, true, center, null, -90d, 100d, 200d, false, false, false, null, false, 0]);
        (double plainX, double plainY, double plainW, double plainH) = ((double, double, double, double))(inverseClip.Invoke(null, [plainRun, 10d, 20d, 30d, 40d]) ?? throw new InvalidOperationException("Expected clip."));
        (double rotX, double rotY, double rotW, double rotH) = ((double, double, double, double))(inverseClip.Invoke(null, [rotatedRun, 10d, 20d, 30d, 40d]) ?? throw new InvalidOperationException("Expected clip."));
        TestAssert.Equal((10d, 20d, 30d, 40d), (plainX, plainY, plainW, plainH));
        TestAssert.True(Math.Abs(rotX - -80d) < 1e-6 && Math.Abs(rotY - 260d) < 1e-6 && Math.Abs(rotW - 40d) < 1e-6 && Math.Abs(rotH - 30d) < 1e-6, "Expected a -90-degree clip to map frame bounds into rotated space, covering text past the frame edge.");
    }

    public static void PptxChartSecondaryValueAxisStyleSelectionUsesSceneSource()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:barChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:barChart>
                <c:barChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>B</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>30</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="10"/><c:axId val="30"/>
                </c:barChart>
                <c:catAx><c:axId val="10"/><c:axPos val="b"/><c:crossAx val="20"/></c:catAx>
                <c:valAx><c:axId val="20"/><c:axPos val="l"/><c:crossAx val="10"/></c:valAx>
                <c:valAx><c:axId val="30"/><c:axPos val="r"/><c:crossAx val="10"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");

        XDocument mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:barChart><c:axId val="10"/><c:axId val="20"/></c:barChart>
                <c:catAx><c:axId val="10"/></c:catAx>
                <c:valAx><c:axId val="20"/><c:axPos val="l"/></c:valAx>
                <c:valAx><c:axId val="99"/><c:axPos val="r"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);
        XNamespace chartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedPlot = mismatchedXmlFallback.Descendants(chartNamespace + "barChart").First();
        System.Reflection.MethodInfo readValueAxes = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartValueAxesForPlot",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected renderer value-axis bridge.");
        System.Reflection.MethodInfo readSecondaryAxis = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlSecondaryValueAxisForChart",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected secondary style value-axis resolver.");

        object scenePrimaryAxis = ((System.Collections.IEnumerable)(readValueAxes.Invoke(null, [chart, chart.Plots[0], mismatchedXmlFallback, mismatchedPlot]) ?? throw new InvalidOperationException("Expected scene value axes."))).Cast<object>().First();
        object xmlPrimaryAxis = ((System.Collections.IEnumerable)(readValueAxes.Invoke(null, [null, null, mismatchedXmlFallback, mismatchedPlot]) ?? throw new InvalidOperationException("Expected XML value axes."))).Cast<object>().First();
        object sceneBacked = readSecondaryAxis.Invoke(null, [chart, mismatchedXmlFallback, scenePrimaryAxis]) ?? throw new InvalidOperationException("Expected scene-backed secondary style axis.");
        object xmlBacked = readSecondaryAxis.Invoke(null, [null, mismatchedXmlFallback, xmlPrimaryAxis]) ?? throw new InvalidOperationException("Expected XML-backed secondary style axis.");

        PptxSceneChartAxis? sceneAxis = (PptxSceneChartAxis?)sceneBacked.GetType().GetProperty("SceneAxis")?.GetValue(sceneBacked);
        XElement? sceneXmlAxis = (XElement?)sceneBacked.GetType().GetProperty("XmlAxis")?.GetValue(sceneBacked);
        PptxSceneChartAxis? xmlSceneAxis = (PptxSceneChartAxis?)xmlBacked.GetType().GetProperty("SceneAxis")?.GetValue(xmlBacked);
        XElement? xmlAxis = (XElement?)xmlBacked.GetType().GetProperty("XmlAxis")?.GetValue(xmlBacked);

        TestAssert.Equal("30", sceneAxis?.Id ?? string.Empty);
        TestAssert.Equal("30", sceneXmlAxis?.Element(chartNamespace + "axId")?.Attribute("val")?.Value ?? string.Empty);
        TestAssert.True(xmlSceneAxis is null, "Expected XML-only secondary style axis selection to remain XML-backed.");
        TestAssert.Equal("99", xmlAxis?.Element(chartNamespace + "axId")?.Attribute("val")?.Value ?? string.Empty);
    }

    public static void PptxChartSeriesStylesUseSceneSource()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea><c:barChart>
                <c:ser>
                  <c:spPr>
                    <a:solidFill><a:srgbClr val="112233"/></a:solidFill>
                    <a:ln w="25400"><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:ln>
                  </c:spPr>
                  <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val>
                </c:ser>
              </c:barChart></c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartPlot plot = chart.Plots[0];

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea><c:barChart>
                <c:ser>
                  <c:spPr>
                    <a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill>
                    <a:ln w="12700"><a:solidFill><a:srgbClr val="FEDCBA"/></a:solidFill></a:ln>
                  </c:spPr>
                </c:ser>
              </c:barChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "barChart").Single();
        System.Reflection.MethodInfo readFills = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlSeriesFills",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected series fill resolver.");
        System.Reflection.MethodInfo readStrokes = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlSeriesStrokes",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected series stroke resolver.");

        object sceneFill = (((System.Collections.IEnumerable?)readFills.Invoke(null, [plot, mismatchedXmlFallback, PptxTheme.Empty, PptxColorMap.Default])) ?? throw new InvalidOperationException("Expected scene fills.")).Cast<object>().Single();
        object sceneStroke = (((System.Collections.IEnumerable?)readStrokes.Invoke(null, [plot, mismatchedXmlFallback, PptxTheme.Empty, PptxColorMap.Default, null])) ?? throw new InvalidOperationException("Expected scene strokes.")).Cast<object>().Single();
        object xmlFill = (((System.Collections.IEnumerable?)readFills.Invoke(null, [null, mismatchedXmlFallback, PptxTheme.Empty, PptxColorMap.Default])) ?? throw new InvalidOperationException("Expected XML fills.")).Cast<object>().Single();
        object xmlStroke = (((System.Collections.IEnumerable?)readStrokes.Invoke(null, [null, mismatchedXmlFallback, PptxTheme.Empty, PptxColorMap.Default, null])) ?? throw new InvalidOperationException("Expected XML strokes.")).Cast<object>().Single();

        TestAssert.Equal(new RgbColor(17, 34, 51), PptxTests.ChartSeriesFillColor(sceneFill));
        TestAssert.Equal(new RgbColor(68, 85, 102), PptxTests.ChartSeriesStrokeColor(sceneStroke));
        TestAssert.Equal(2d, PptxTests.ChartSeriesStrokeWidth(sceneStroke));
        TestAssert.Equal(new RgbColor(171, 205, 239), PptxTests.ChartSeriesFillColor(xmlFill));
        TestAssert.Equal(new RgbColor(254, 220, 186), PptxTests.ChartSeriesStrokeColor(xmlStroke));
        TestAssert.Equal(1d, PptxTests.ChartSeriesStrokeWidth(xmlStroke));
    }

    public static void PptxChartSeriesPointStylesUseSceneSource()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea><c:barChart>
                <c:ser>
                  <c:dPt>
                    <c:idx val="1"/>
                    <c:spPr>
                      <a:solidFill><a:srgbClr val="123456"/></a:solidFill>
                      <a:ln w="38100"><a:solidFill><a:srgbClr val="654321"/></a:solidFill></a:ln>
                    </c:spPr>
                  </c:dPt>
                  <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt></c:numLit></c:val>
                </c:ser>
              </c:barChart></c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartPlot plot = chart.Plots[0];

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea><c:barChart>
                <c:ser>
                  <c:dPt>
                    <c:idx val="0"/>
                    <c:spPr>
                      <a:solidFill><a:srgbClr val="A0B0C0"/></a:solidFill>
                      <a:ln w="12700"><a:solidFill><a:srgbClr val="0C0B0A"/></a:solidFill></a:ln>
                    </c:spPr>
                  </c:dPt>
                </c:ser>
              </c:barChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "barChart").Single();
        System.Reflection.MethodInfo readPointFills = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlSeriesPointFills",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected series point fill resolver.");
        System.Reflection.MethodInfo readPointStrokes = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlSeriesPointStrokes",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected series point stroke resolver.");

        object scenePointFills = (((System.Collections.IEnumerable?)readPointFills.Invoke(null, [plot, mismatchedXmlFallback, PptxTheme.Empty, PptxColorMap.Default])) ?? throw new InvalidOperationException("Expected scene point fills.")).Cast<object>().Single();
        object scenePointStrokes = (((System.Collections.IEnumerable?)readPointStrokes.Invoke(null, [plot, mismatchedXmlFallback, PptxTheme.Empty, PptxColorMap.Default])) ?? throw new InvalidOperationException("Expected scene point strokes.")).Cast<object>().Single();
        object xmlPointFills = (((System.Collections.IEnumerable?)readPointFills.Invoke(null, [null, mismatchedXmlFallback, PptxTheme.Empty, PptxColorMap.Default])) ?? throw new InvalidOperationException("Expected XML point fills.")).Cast<object>().Single();
        object xmlPointStrokes = (((System.Collections.IEnumerable?)readPointStrokes.Invoke(null, [null, mismatchedXmlFallback, PptxTheme.Empty, PptxColorMap.Default])) ?? throw new InvalidOperationException("Expected XML point strokes.")).Cast<object>().Single();

        object sceneFill = PptxTests.ChartDictionaryValue(scenePointFills, 1);
        object sceneStroke = PptxTests.ChartDictionaryValue(scenePointStrokes, 1);
        object xmlFill = PptxTests.ChartDictionaryValue(xmlPointFills, 0);
        object xmlStroke = PptxTests.ChartDictionaryValue(xmlPointStrokes, 0);

        TestAssert.Equal(new RgbColor(18, 52, 86), PptxTests.ChartSeriesFillColor(sceneFill));
        TestAssert.Equal(new RgbColor(101, 67, 33), PptxTests.ChartSeriesStrokeColor(sceneStroke));
        TestAssert.Equal(3d, PptxTests.ChartSeriesStrokeWidth(sceneStroke));
        TestAssert.Equal(new RgbColor(160, 176, 192), PptxTests.ChartSeriesFillColor(xmlFill));
        TestAssert.Equal(new RgbColor(12, 11, 10), PptxTests.ChartSeriesStrokeColor(xmlStroke));
        TestAssert.Equal(1d, PptxTests.ChartSeriesStrokeWidth(xmlStroke));
    }
}
