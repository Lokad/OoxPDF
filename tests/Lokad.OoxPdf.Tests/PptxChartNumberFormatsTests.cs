using System.Xml.Linq;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxChartNumberFormatsTests
{
    public static void PptxChartNumberScalingCommasDivideDisplayValue()
    {
        TestAssert.Equal("1,500", FormatChartNumber(1500000d, "#,##0,"));
        TestAssert.Equal("2.5", FormatChartNumber(2500000d, "0.0,,"));
        TestAssert.Equal("0", FormatChartNumber(499d, "#,##0,"));
    }

    public static void PptxChartNumberEscapedPercentRendersLiterally()
    {
        TestAssert.Equal("1%", FormatChartNumber(0.5d, "0\\%"));
        TestAssert.Equal("0.3%", FormatChartNumber(0.256d, "0.0\\%"));
    }

    public static void PptxChartNumberNegativeSectionSelectsSecondPattern()
    {
        TestAssert.Equal("-12.3", FormatChartNumber(-12.345d, "0.0;-0.0"));
        TestAssert.Equal("12.3", FormatChartNumber(12.345d, "0.0;-0.0"));
        TestAssert.Equal("-1,235", FormatChartNumber(-1234.6d, "#,##0;-#,##0"));
        TestAssert.Equal("(1,235)", FormatChartNumber(-1234.6d, "#,##0;(#,##0)"));
    }

    public static void PptxChartNumberZeroSectionSelectsThirdPattern()
    {
        TestAssert.Equal("0.0", FormatChartNumber(0d, "0.0;-0.0;0.0"));
        TestAssert.Equal("0.0", FormatChartNumber(0d, "0.0;-0.0"));
    }

    public static void PptxChartNumberCommonFormatsKeepExistingBehavior()
    {
        TestAssert.Equal("$1,235", FormatChartNumber(1234.5d, "$#,##0"));
        TestAssert.Equal("13%", FormatChartNumber(0.126d, "0%"));
        TestAssert.Equal("1,234.57", FormatChartNumber(1234.567d, "#,##0.00"));
        TestAssert.Equal("-12.3", FormatChartNumber(-12.345d, "0.0"));

    }

    public static void PptxChartNumberScientificNotationRendersMantissaAndExponent()
    {
        TestAssert.Equal("1.2E+04", FormatChartNumber(12345d, "0.0E+00"));
        TestAssert.Equal("1.23E-03", FormatChartNumber(0.00123d, "0.00E+00"));
        TestAssert.Equal("12.3E+03", FormatChartNumber(12345d, "00.0E+00"));
        TestAssert.Equal("5E+00", FormatChartNumber(5d, "0E+00"));
        TestAssert.Equal("1.0E+03", FormatChartNumber(999.9d, "0.0E+00"));
        TestAssert.Equal("-5.0E-01", FormatChartNumber(-0.5d, "0.0E+00"));
        TestAssert.Equal("0.0E+00", FormatChartNumber(0d, "0.0E+00"));
        TestAssert.Equal("1.2e+04", FormatChartNumber(12345d, "0.0e+00"));
    }

    public static void PptxChartNumberFractionsRenderWholeAndNumeratorDenominator()
    {
        TestAssert.Equal("1 1/2", FormatChartNumber(1.5d, "# ?/?"));
        TestAssert.Equal("1/2", FormatChartNumber(0.5d, "# ?/?"));
        TestAssert.Equal("1/3", FormatChartNumber(1d / 3d, "# ??/??"));
        TestAssert.Equal("2 1/4", FormatChartNumber(2.25d, "# ?/?"));
        TestAssert.Equal("0", FormatChartNumber(0d, "# ?/?"));
        TestAssert.Equal("2", FormatChartNumber(2d, "?/?"));
        TestAssert.Equal("-1 1/2", FormatChartNumber(-1.5d, "# ?/?"));
        TestAssert.Equal("-1 1/2", FormatChartNumber(-1.5d, "# ?/?;-# ?/?"));
    }

    public static void PptxChartNumberAccountingSkipsCushionAndFillRuns()
    {
        TestAssert.Equal("$1,235", FormatChartNumber(1234.5d, "_($* #,##0_)"));
        TestAssert.Equal("($1,235)", FormatChartNumber(-1234.5d, "_($* #,##0_);_($* (#,##0)"));
        TestAssert.Equal("$1,234.57", FormatChartNumber(1234.567d, "_($* #,##0.00_)"));
    }

    public static void PptxChartNumberDatesRenderCalendarParts()
    {
        TestAssert.Equal("3/15/23", FormatChartNumber(45000d, "m/d/yy"));
        TestAssert.Equal("2023-03-15", FormatChartNumber(45000d, "yyyy-mm-dd"));
        TestAssert.Equal("15-Mar-23", FormatChartNumber(45000d, "d-mmm-yy"));
        TestAssert.Equal("Wednesday, March 15, 2023", FormatChartNumber(45000d, "dddd, mmmm d, yyyy"));
        TestAssert.Equal("12:00", FormatChartNumber(0.5d, "h:mm"));
        TestAssert.Equal("12:00:00", FormatChartNumber(0.5d, "h:mm:ss"));
        TestAssert.Equal("3/15/23 18:00", FormatChartNumber(45000.75d, "m/d/yy h:mm"));
        TestAssert.Equal("1/1/23", FormatChartNumber(44927d, "m/d/yy"));
        TestAssert.Equal("27:00", FormatChartNumber(1.125d, "[h]:mm"));
        TestAssert.Equal("45000", FormatChartNumber(45000d, "0"));
        TestAssert.Equal("-5", FormatChartNumber(-5d, "m/d/yy"));
    }

    public static void PptxChartNumberDatesOutsideCalendarRangeFallBackToNumeric()
    {
        foreach (bool date1904 in new[] { false, true })
        {
            DateOnly origin = date1904 ? new DateOnly(1904, 1, 1) : new DateOnly(1899, 12, 30);
            int maximum = DateOnly.MaxValue.DayNumber - origin.DayNumber;
            TestAssert.Equal("9999-12-31", FormatChartNumber(maximum, "yyyy-mm-dd", date1904));
            TestAssert.Equal("9999-12-31 12:00", FormatChartNumber(maximum + 0.5d, "yyyy-mm-dd h:mm", date1904));
            TestAssert.Equal((maximum + 1).ToString(System.Globalization.CultureInfo.InvariantCulture),
                FormatChartNumber(maximum + 1d, "yyyy-mm-dd", date1904));
            TestAssert.Equal("4000000", FormatChartNumber(4000000d, "yyyy-mm-dd", date1904));
            TestAssert.Equal("2147483647", FormatChartNumber(int.MaxValue, "yyyy-mm-dd", date1904));
            TestAssert.Equal("-5", FormatChartNumber(-5d, "yyyy-mm-dd", date1904));
        }
    }

    public static void PptxChartNumberDateSystemsResolveSerials()
    {
        TestAssert.Equal("1/1/04", FormatChartNumber(0d, "m/d/yy", date1904: true));
        TestAssert.Equal("1/1/05", FormatChartNumber(366d, "m/d/yy", date1904: true));
        TestAssert.Equal("3/1/00", FormatChartNumber(60d, "m/d/yy", date1904: false));
    }

    public static void PptxChartDateFallbackPreservesPlotAndLabels()
    {
        foreach (bool date1904 in new[] { false, true })
        {
            var entries = new Dictionary<string, string> {
                ["[Content_Types].xml"] = PptxTests.BasicContentTypes().Replace("</Types>", "<Override PartName=\"/ppt/charts/chart1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawingml.chart+xml\"/></Types>"),
                ["_rels/.rels"] = PptxTests.PackageRelationship(),
                ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
                ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
                ["ppt/slides/_rels/slide1.xml.rels"] = """
                    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="chart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>
                    """,
                ["ppt/slides/slide1.xml"] = """
                    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree>
                    <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="3657600"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="chart"/></a:graphicData></a:graphic></p:graphicFrame>
                    </p:spTree></p:cSld></p:sld>
                    """,
                ["ppt/charts/chart1.xml"] = $$"""
                    <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:date1904 val="{{(date1904 ? 1 : 0)}}"/><c:chart><c:plotArea>
                    <c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/>
                    <c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                    <c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>4000000</c:v></c:pt><c:pt idx="1"><c:v>5000000</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val="10"/><c:axId val="20"/></c:barChart>
                    <c:catAx><c:axId val="10"/><c:axPos val="b"/><c:crossAx val="20"/><c:tickLblPos val="nextTo"/></c:catAx>
                    <c:valAx><c:axId val="20"/><c:scaling><c:max val="6000000"/><c:min val="0"/></c:scaling><c:axPos val="l"/><c:numFmt formatCode="yyyy-mm-dd" sourceLinked="0"/><c:tickLblPos val="nextTo"/><c:crossAx val="10"/><c:majorUnit val="1000000"/></c:valAx>
                    </c:plotArea></c:chart></c:chartSpace>
                    """
            };
            using var input = TestFixtures.CreateZipPackage(entries);
            using var output = new MemoryStream();
            var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
            OoxPdfConverter.Convert(input, output, new OoxPdfOptions {
                InputKind = OoxPdfInputKind.Pptx, FontResolver = new PortableRenderingTests.PinnedResolver(), DiagnosticSink = diagnostics.Add
            });
            string pdf = System.Text.Encoding.Latin1.GetString(output.ToArray());
            TestAssert.True(!diagnostics.Any(d => d.Id == "PPTX_NODE_RENDER_FAILED"), "A date-format fallback must retain its chart.");
            TestAssert.True(DocxTests.CountPdfTextShows(pdf) >= 9, "Date fallback must preserve category and axis labels.");
            TestAssert.True(PptxTests.CountOccurrences(pdf, " re f") >= 3, "Date fallback must preserve the plot's bars.");
        }
    }

    public static void PptxChartAxisDate1904ResolvesFromChartSpace()
    {
        System.Xml.Linq.XDocument chart = System.Xml.Linq.XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:date1904 val="1"/>
              <c:chart><c:plotArea>
                <c:valAx><c:axId val="20"/><c:numFmt formatCode="m/d/yy" sourceLinked="0"/><c:axPos val="l"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);
        System.Xml.Linq.XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        System.Xml.Linq.XElement axis = chart.Descendants(c + "valAx").Single();
        System.Reflection.MethodInfo format = typeof(PptxRenderer).GetMethods(System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)
            .Single(m => m.Name == "FormatSceneOrXmlChartAxisLabel" && m.GetParameters().Length == 4);

        TestAssert.Equal("1/1/04", format.Invoke(null, [0d, null, axis, null]) as string ?? throw new System.InvalidOperationException("Expected formatted axis label."));
    }

    public static void PptxChartNumberConditionalSectionsSelectFirstMatch()
    {
        TestAssert.Equal("150.0", FormatChartNumber(150d, "[>=100]0.0;[Red]0"));
        TestAssert.Equal("50", FormatChartNumber(50d, "[>=100]0.0;[Red]0"));
        TestAssert.Equal("", FormatChartNumber(25d, "[>=100]0.0;[>50]0.0"));
        TestAssert.Equal("5", FormatChartNumber(5d, "[=5]0;[=6]0"));
        TestAssert.Equal("", FormatChartNumber(7d, "[=5]0;[=6]0"));
        TestAssert.Equal("2", FormatChartNumber(2d, "[<>5]0;[=5]0"));
    }

    public static void PptxChartDataLabelDatesHonorChartSpaceDateSystem()
    {
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement dated = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:date1904 val="1"/>
              <c:chart><c:plotArea><c:barChart>
                <c:dLbls><c:numFmt formatCode="m/d/yy" sourceLinked="0"/><c:showVal val="1"/></c:dLbls>
              </c:barChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "barChart").Single();
        XElement undated = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea><c:barChart>
                <c:dLbls><c:numFmt formatCode="m/d/yy" sourceLinked="0"/><c:showVal val="1"/></c:dLbls>
              </c:barChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "barChart").Single();
        System.Reflection.MethodInfo readOptions = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlDataLabelOptions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new System.InvalidOperationException("Expected data-label option resolver.");
        System.Reflection.MethodInfo formatLabel = typeof(PptxRenderer).GetMethods(System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)
            .Single(m => m.Name == "FormatChartDataLabelValue" && m.GetParameters().Length == 4);
        object datedOptions = readOptions.Invoke(null, new object?[] { null, null, dated, PptxTheme.Empty, PptxColorMap.Default }) ?? throw new System.InvalidOperationException("Expected dated label options.");
        object undatedOptions = readOptions.Invoke(null, new object?[] { null, null, undated, PptxTheme.Empty, PptxColorMap.Default }) ?? throw new System.InvalidOperationException("Expected undated label options.");
        TestAssert.Equal("1/1/04", formatLabel.Invoke(null, new object?[] { 0d, datedOptions, null, null }) as string ?? throw new System.InvalidOperationException("Expected formatted dated data label."));
        TestAssert.Equal("1/0/00", formatLabel.Invoke(null, new object?[] { 0d, undatedOptions, null, null }) as string ?? throw new System.InvalidOperationException("Expected formatted undated data label."));
    }

    public static void PptxChartNumberLiteralsEmitAroundValue()
    {
        TestAssert.Equal("5 kg", FormatChartNumber(5d, "0\" kg\""));
        TestAssert.Equal("Total: 5", FormatChartNumber(5d, "\"Total: \"0"));
        TestAssert.Equal("5 kg", FormatChartNumber(5d, "0\\ kg"));
        TestAssert.Equal("5%", FormatChartNumber(5d, "0\"%\""));
        TestAssert.Equal("abc", FormatChartNumber(5d, "\"abc\""));
        TestAssert.Equal("5th", FormatChartNumber(5d, "0th"));
        TestAssert.Equal("5", FormatChartNumber(5d, "0 items"));
    }

    public static void PptxChartNumberUnsupportedConstructsIdentifyFallbackGaps()
    {
        TestAssert.Equal("", GetUnsupportedChartNumberConstructs("[$-409]0"));
        TestAssert.Equal("", GetUnsupportedChartNumberConstructs("[$-407]0.00"));
        TestAssert.Equal("locale", GetUnsupportedChartNumberConstructs("[$-999]0"));
        TestAssert.Equal("locale", GetUnsupportedChartNumberConstructs("[$-F400]0"));
        TestAssert.Equal("color", GetUnsupportedChartNumberConstructs("[Red]0.0"));
        TestAssert.Equal("color", GetUnsupportedChartNumberConstructs("[Red]0;[Blue]0.0"));
        TestAssert.Equal("native-digits", GetUnsupportedChartNumberConstructs("[DBNum1]0"));
        TestAssert.Equal("bang-escape", GetUnsupportedChartNumberConstructs("0!-"));
        TestAssert.Equal("text-placeholder", GetUnsupportedChartNumberConstructs("0@"));
        TestAssert.Equal("", GetUnsupportedChartNumberConstructs("[>=100]0.0"));
        TestAssert.Equal("", GetUnsupportedChartNumberConstructs("0.0"));
        TestAssert.Equal("", GetUnsupportedChartNumberConstructs("[h]:mm"));
        TestAssert.Equal("", GetUnsupportedChartNumberConstructs("m/d/yy"));
        TestAssert.Equal("", GetUnsupportedChartNumberConstructs("0\" kg\""));
    }

    private static string GetUnsupportedChartNumberConstructs(string formatCode)
    {
        System.Reflection.MethodInfo get = typeof(PptxRenderer).GetMethod(
            "GetUnsupportedChartNumberFormatConstructs",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected unsupported-construct bridge.");
        System.Collections.Generic.IReadOnlyList<string> constructs = (System.Collections.Generic.IReadOnlyList<string>)(get.Invoke(null, [formatCode]) ?? throw new InvalidOperationException("Expected unsupported constructs."));
        return string.Join(",", constructs);
    }

    public static void PptxChartNumberLocaleCurrencyResolvesSymbol()
    {
        TestAssert.Equal("\u20AC5,00", FormatChartNumber(5d, "[$\u20AC-407]0.00"));
        TestAssert.Equal("\u00A35", FormatChartNumber(5d, "[$\u00A3-809]0"));
        TestAssert.Equal("5", FormatChartNumber(5d, "[$-409]0"));
        TestAssert.Equal("\u20AC3/15/23", FormatChartNumber(45000d, "[$\u20AC-407]m/d/yy"));
    }

    public static void PptxChartNumberLocaleSeparatorsFollowCalibratedLcids()
    {
        TestAssert.Equal("1,234.56", FormatChartNumber(1234.56d, "[$-409]#,##0.00"));
        TestAssert.Equal("1.234,56", FormatChartNumber(1234.56d, "[$-407]#,##0.00"));
        TestAssert.Equal("1.234,56", FormatChartNumber(1234.56d, "[$-410]#,##0.00"));
        TestAssert.Equal("1,234.56", FormatChartNumber(1234.56d, "[$-411]#,##0.00"));
        TestAssert.Equal("1,234", FormatChartNumber(1234d, "[$-809]#,##0"));
        TestAssert.Equal("500,0%", FormatChartNumber(5d, "[$-407]0.0%"));
        TestAssert.Equal("1,2E+03", FormatChartNumber(1234d, "[$-407]0.0E+00"));
        TestAssert.Equal("\u20AC1.234,56", FormatChartNumber(1234.56d, "[$\u20AC-410]#,##0.00"));
        TestAssert.Equal("\u00A35", FormatChartNumber(5d, "[$\u00A3-809]0"));
        TestAssert.Equal("5", FormatChartNumber(5d, "[$-409]0"));
    }

    public static void PptxChartNumberGeneralFallbackKeepsFullPrecision()
    {
        TestAssert.Equal("6.438276616", FormatChartAxisLabel(6.438276615812609d));
        TestAssert.Equal("1.25", FormatChartAxisLabel(1.25d));
        TestAssert.Equal("0.125", FormatChartAxisLabel(0.125d));
        TestAssert.Equal("100", FormatChartAxisLabel(100d));
        TestAssert.Equal("2.5", FormatChartAxisLabel(2.5d));
    }

    private static string FormatChartNumber(double value, string formatCode, bool date1904 = false)
    {
        System.Reflection.MethodInfo format = typeof(PptxRenderer).GetMethod(
            "FormatChartNumber",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart number-format bridge.");
        return (string)(format.Invoke(null, [value, formatCode, date1904]) ?? throw new InvalidOperationException("Expected formatted chart number."));
    }

    private static string FormatChartAxisLabel(double value)
    {
        System.Reflection.MethodInfo format = typeof(PptxRenderer).GetMethod(
            "FormatChartAxisLabel",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected chart axis-label bridge.");
        return (string)(format.Invoke(null, [value, null]) ?? throw new InvalidOperationException("Expected formatted axis label."));
    }
}
