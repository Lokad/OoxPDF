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

internal static class PptxTests
{
    internal static string ShortHash(string text)
    {
        byte[] hash = SHA256.HashData(Encoding.UTF8.GetBytes(text));
        return Convert.ToHexString(hash, 0, 4);
    }

    internal static double AveragePositiveAdjustment(IReadOnlyList<PptxTextGlyphLayoutSnapshot> glyphs)
    {
        double[] values = glyphs.Select(glyph => glyph.AdjustmentBefore).Where(value => value > 0d).ToArray();
        return values.Length == 0 ? 0d : values.Average();
    }

    internal static double AverageNegativeAdjustment(IReadOnlyList<PptxTextGlyphLayoutSnapshot> glyphs)
    {
        double[] values = glyphs.Select(glyph => glyph.AdjustmentBefore).Where(value => value < 0d).ToArray();
        return values.Length == 0 ? 0d : values.Average();
    }

    internal static int CountFilledDecorationPaths(string pdf, string rgbPattern)
    {
        return Regex.Matches(
            pdf,
            $@"{rgbPattern} rg\s+(?:-?[0-9.]+ -?[0-9.]+ m\s+)(?:-?[0-9.]+ -?[0-9.]+ l\s+){{5}}h\s+f\*").Count;
    }

    internal static string RenderLineChartWithDisplayBlanksAs(string displayBlanksAs)
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:graphicFrame><p:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="3657600"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/charts/chart1.xml"] = TestFixtures.Utf8($$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:lineChart>
                  <c:grouping val="standard"/>
                  <c:marker val="0"/>
                  <c:ser>
                    <c:spPr><a:ln w="25400"><a:solidFill><a:srgbClr val="12AB34"/></a:solidFill></a:ln></c:spPr>
                    <c:marker><c:symbol val="none"/></c:marker>
                    <c:cat><c:strLit><c:ptCount val="3"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strLit></c:cat>
                    <c:val><c:numLit><c:ptCount val="3"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="2"><c:v>20</c:v></c:pt></c:numLit></c:val>
                  </c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="20"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/><c:min val="0"/><c:max val="20"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea><c:dispBlanksAs val="{{displayBlanksAs}}"/></c:chart></c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var collector = new DiagnosticCollector();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = collector.Add });

        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_CHART_STATIC_FALLBACK"), $"Line chart displayBlanksAs={displayBlanksAs} should render without static fallback diagnostics.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"), $"Line chart displayBlanksAs={displayBlanksAs} should not emit unsupported chart diagnostics.");
        return File.ReadAllText(output, Encoding.ASCII);
    }

    internal static int CountLineChartSeriesSegments(string pdf, string strokeColorOperation)
    {
        int colorIndex = pdf.IndexOf(strokeColorOperation, StringComparison.Ordinal);
        TestAssert.True(colorIndex >= 0, $"Expected series stroke color operation '{strokeColorOperation}'.");
        int nextColorIndex = pdf.IndexOf(" RG", colorIndex + strokeColorOperation.Length, StringComparison.Ordinal);
        string seriesScope = nextColorIndex < 0 ? pdf[colorIndex..] : pdf[colorIndex..nextColorIndex];
        return Regex.Matches(seriesScope, @"[0-9.]+ [0-9.]+ l").Count;
    }

    internal static IReadOnlyList<OoxPdfDiagnostic> ConvertSingleChartAndCollectDiagnostics(string chartXml)
    {
        string contentTypes = BasicContentTypes().Replace(
            "</Types>",
            """
              <Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
            </Types>
            """);
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = contentTypes,
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/_rels/slide1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
                </Relationships>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:nvGraphicFramePr><p:cNvPr id="2" name="Chart"/><p:nvPr/></p:nvGraphicFramePr>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="2286000"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """,
            ["ppt/charts/chart1.xml"] = chartXml
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
        return diagnostics;
    }

    internal static PptxSceneChart? BuildSingleChartScene(string chartXml)
    {
        return BuildSingleChartPackageScene(chartXml).Slides[0].SlideNodes[0].Chart;
    }

    internal static PptxSceneNodeSnapshot BuildSingleChartSceneSnapshot(string chartXml)
    {
        (PptxDocument document, OoxPackage package) = BuildSingleChartPackage(chartXml);
        PptxSceneSnapshot scene = PptxRenderer.InspectScene(document, package);
        return scene.Slides[0].SlideNodes[0];
    }

    internal static PptxScene BuildSingleChartPackageScene(string chartXml)
    {
        (PptxDocument document, OoxPackage package) = BuildSingleChartPackage(chartXml);
        return new PptxSceneBuilder().Build(document, package, CancellationToken.None);
    }

    internal static (PptxDocument Document, OoxPackage Package) BuildSingleChartPackage(
        string chartXml,
        IReadOnlyDictionary<string, byte[]>? additionalParts = null)
    {
        var parts = new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="2743200"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/charts/chart1.xml"] = TestFixtures.Utf8(chartXml)
        };
        if (additionalParts is not null)
        {
            foreach (KeyValuePair<string, byte[]> part in additionalParts)
            {
                parts[part.Key] = part.Value;
            }
        }

        string input = TestFixtures.WriteTempPackage(".pptx", parts);

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        return (document, package);
    }

    internal static bool ChartBooleanOptionValue(object option)
    {
        return (bool)(option.GetType().GetProperty("Value")?.GetValue(option)
            ?? throw new InvalidOperationException("Expected chart boolean option value."));
    }

    internal static string ChartBooleanOptionRawValue(object option)
    {
        return (string)(option.GetType().GetProperty("RawValue")?.GetValue(option)
            ?? throw new InvalidOperationException("Expected chart boolean option raw value."));
    }

    internal static bool ChartBooleanOptionIsDefined(object option)
    {
        return (bool)(option.GetType().GetProperty("IsDefined")?.GetValue(option)
            ?? throw new InvalidOperationException("Expected chart boolean option defined state."));
    }

    internal static object ChartDataLabelFlagOption(object options, string flagName)
    {
        object flags = options.GetType().GetProperty("FlagOptions")?.GetValue(options)
            ?? throw new InvalidOperationException("Expected chart data-label flag options.");
        foreach (object entry in (System.Collections.IEnumerable)flags)
        {
            string key = (string)(entry.GetType().GetProperty("Key")?.GetValue(entry)
                ?? throw new InvalidOperationException("Expected chart data-label flag key."));
            if (key == flagName)
            {
                return entry.GetType().GetProperty("Value")?.GetValue(entry)
                    ?? throw new InvalidOperationException("Expected chart data-label flag value.");
            }
        }

        throw new InvalidOperationException($"Expected chart data-label flag '{flagName}'.");
    }

    internal static object ChartDataLabelOverride(object options, int index)
    {
        object overrides = options.GetType().GetProperty("Overrides")?.GetValue(options)
            ?? throw new InvalidOperationException("Expected chart data-label overrides.");
        foreach (object entry in (System.Collections.IEnumerable)overrides)
        {
            int key = (int)(entry.GetType().GetProperty("Key")?.GetValue(entry)
                ?? throw new InvalidOperationException("Expected chart data-label override key."));
            if (key == index)
            {
                return entry.GetType().GetProperty("Value")?.GetValue(entry)
                    ?? throw new InvalidOperationException("Expected chart data-label override value.");
            }
        }

        throw new InvalidOperationException($"Expected chart data-label override '{index}'.");
    }

    internal static object ChartDataLabelTextBodyProperties(object options)
    {
        return ChartTextBodyProperties(options);
    }

    internal static object ChartDataLabelLeaderLines(object options)
    {
        return options.GetType().GetProperty("LeaderLines")?.GetValue(options)
            ?? throw new InvalidOperationException("Expected chart data-label leader lines.");
    }

    internal static bool ChartDataLabelLeaderLinesIsDefined(object leaderLines)
    {
        return (bool)(leaderLines.GetType().GetProperty("IsDefined")?.GetValue(leaderLines)
            ?? throw new InvalidOperationException("Expected chart data-label leader-line defined state."));
    }

    internal static object ChartDataLabelLeaderLinesStroke(object leaderLines)
    {
        return leaderLines.GetType().GetProperty("Stroke")?.GetValue(leaderLines)
            ?? throw new InvalidOperationException("Expected chart data-label leader-line stroke.");
    }

    internal static object ChartTextBodyProperties(object options)
    {
        return options.GetType().GetProperty("TextBodyProperties")?.GetValue(options)
            ?? throw new InvalidOperationException("Expected chart text body properties.");
    }

    internal static double? ChartTextBodyRotationDegrees(object properties)
    {
        return (double?)properties.GetType().GetProperty("RotationDegrees")?.GetValue(properties);
    }

    internal static string ChartTextBodyRotationValue(object properties)
    {
        return (string?)properties.GetType().GetProperty("RotationValue")?.GetValue(properties) ?? string.Empty;
    }

    internal static string ChartMarkerStyleSymbol(object marker)
    {
        return (string)(marker.GetType().GetProperty("Symbol")?.GetValue(marker)
            ?? throw new InvalidOperationException("Expected chart marker symbol."));
    }

    internal static string? ChartMarkerStyleSizeValue(object marker)
    {
        return (string?)marker.GetType().GetProperty("SizeValue")?.GetValue(marker);
    }

    internal static double ChartMarkerStyleSize(object marker)
    {
        return (double)(marker.GetType().GetProperty("Size")?.GetValue(marker)
            ?? throw new InvalidOperationException("Expected chart marker size."));
    }

    internal static bool ChartMarkerStyleIsDefined(object marker)
    {
        return (bool)(marker.GetType().GetProperty("IsDefined")?.GetValue(marker)
            ?? throw new InvalidOperationException("Expected chart marker defined state."));
    }

    internal static object ChartMarkerStyleFill(object marker)
    {
        return marker.GetType().GetProperty("Fill")?.GetValue(marker)
            ?? throw new InvalidOperationException("Expected chart marker fill.");
    }

    internal static object ChartMarkerStyleStroke(object marker)
    {
        return marker.GetType().GetProperty("Stroke")?.GetValue(marker)
            ?? throw new InvalidOperationException("Expected chart marker stroke.");
    }

    internal static RgbColor ChartSeriesFillColor(object fill)
    {
        return (RgbColor)(fill.GetType().GetProperty("Color")?.GetValue(fill)
            ?? throw new InvalidOperationException("Expected chart series fill color."));
    }

    internal static RgbColor ChartSeriesStrokeColor(object stroke)
    {
        return (RgbColor)(stroke.GetType().GetProperty("Color")?.GetValue(stroke)
            ?? throw new InvalidOperationException("Expected chart series stroke color."));
    }

    internal static double ChartSeriesStrokeWidth(object stroke)
    {
        return (double)(stroke.GetType().GetProperty("Width")?.GetValue(stroke)
            ?? throw new InvalidOperationException("Expected chart series stroke width."));
    }

    internal static object ChartDictionaryValue(object dictionary, int key)
    {
        foreach (object entry in (System.Collections.IEnumerable)dictionary)
        {
            int candidateKey = (int)(entry.GetType().GetProperty("Key")?.GetValue(entry)
                ?? throw new InvalidOperationException("Expected chart dictionary key."));
            if (candidateKey == key)
            {
                return entry.GetType().GetProperty("Value")?.GetValue(entry)
                    ?? throw new InvalidOperationException("Expected chart dictionary value.");
            }
        }

        throw new InvalidOperationException($"Expected chart dictionary key '{key}'.");
    }

    internal static string ReadPdfDecodedAscii(string path)
    {
        byte[] bytes = File.ReadAllBytes(path);
        string pdf = Encoding.ASCII.GetString(bytes);
        var text = new StringBuilder(pdf);
        int searchStart = 0;
        while (true)
        {
            int streamMarker = pdf.IndexOf("stream", searchStart, StringComparison.Ordinal);
            if (streamMarker < 0)
            {
                break;
            }

            int streamStart = streamMarker + "stream".Length;
            if (streamStart < bytes.Length && bytes[streamStart] == (byte)'\r')
            {
                streamStart++;
            }

            if (streamStart < bytes.Length && bytes[streamStart] == (byte)'\n')
            {
                streamStart++;
            }

            int streamEnd = pdf.IndexOf("endstream", streamStart, StringComparison.Ordinal);
            if (streamEnd < 0)
            {
                break;
            }

            int dictionaryStart = Math.Max(0, pdf.LastIndexOf("<<", streamMarker, StringComparison.Ordinal));
            string dictionary = pdf.Substring(dictionaryStart, streamMarker - dictionaryStart);
            if (dictionary.Contains("/FlateDecode", StringComparison.Ordinal))
            {
                using var input = new MemoryStream(bytes, streamStart, streamEnd - streamStart);
                using var deflate = new ZLibStream(input, CompressionMode.Decompress);
                using var output = new MemoryStream();
                deflate.CopyTo(output);
                text.Append(Encoding.ASCII.GetString(output.ToArray()));
            }

            searchStart = streamEnd + "endstream".Length;
        }

        return text.ToString();
    }

    internal static List<byte[]> ReadPdfDeviceRgbImageStreams(string path, int width, int height)
    {
        byte[] bytes = File.ReadAllBytes(path);
        string pdf = Encoding.ASCII.GetString(bytes);
        var streams = new List<byte[]>();
        int searchStart = 0;
        while (true)
        {
            int streamMarker = pdf.IndexOf("stream", searchStart, StringComparison.Ordinal);
            if (streamMarker < 0)
            {
                break;
            }

            int streamStart = streamMarker + "stream".Length;
            if (streamStart < bytes.Length && bytes[streamStart] == (byte)'\r')
            {
                streamStart++;
            }

            if (streamStart < bytes.Length && bytes[streamStart] == (byte)'\n')
            {
                streamStart++;
            }

            int streamEnd = pdf.IndexOf("endstream", streamStart, StringComparison.Ordinal);
            if (streamEnd < 0)
            {
                break;
            }

            int dictionaryStart = Math.Max(0, pdf.LastIndexOf("<<", streamMarker, StringComparison.Ordinal));
            string dictionary = pdf.Substring(dictionaryStart, streamMarker - dictionaryStart);
            if (dictionary.Contains("/Subtype /Image", StringComparison.Ordinal) &&
                dictionary.Contains("/ColorSpace /DeviceRGB", StringComparison.Ordinal) &&
                dictionary.Contains(FormattableString.Invariant($"/Width {width} /Height {height}"), StringComparison.Ordinal) &&
                dictionary.Contains("/Filter /FlateDecode", StringComparison.Ordinal))
            {
                using var input = new MemoryStream(bytes, streamStart, streamEnd - streamStart);
                using var deflate = new ZLibStream(input, CompressionMode.Decompress);
                using var output = new MemoryStream();
                deflate.CopyTo(output);
                streams.Add(output.ToArray());
            }

            searchStart = streamEnd + "endstream".Length;
        }

        return streams;
    }

    internal static double ReadOnlyTextBaseline(string pdf)
    {
        MatchCollection matches = Regex.Matches(pdf, @"1 0 0 1 [0-9.]+ (?<y>[0-9.]+) Tm");
        TestAssert.Equal(1, matches.Count);
        return double.Parse(matches[0].Groups["y"].Value, CultureInfo.InvariantCulture);
    }

    internal static int CountTextMatrices(string pdf)
    {
        return Regex.Matches(pdf, @"1 0 0 1 [0-9.]+ [0-9.]+ Tm").Count;
    }

    internal static byte[] EmbeddedChartWorkbook()
    {
        using MemoryStream stream = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
                  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
                  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
                  <Override PartName="/xl/tables/table2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
                </Relationships>
                """,
            ["xl/_rels/workbook.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
                  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                </Relationships>
                """,
            ["xl/worksheets/_rels/sheet1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table2.xml"/>
                </Relationships>
                """,
            ["xl/workbook.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <workbookPr date1904="1"/>
                  <sheets><sheet name="Sheet1" sheetId="1" state="hidden" r:id="rId1"/></sheets>
                  <definedNames>
                    <definedName name="SalesLabels">Sheet1!$A$2:$A$4</definedName>
                    <definedName name="SalesValues">Sheet1!$B$2:$B$4</definedName>
                    <definedName name="SalesUnionValues">Sheet1!$B$2:$B$2,$B$4:$B$4</definedName>
                    <definedName name="SheetLocalValues" localSheetId="0">Sheet1!$B$2:$B$4</definedName>
                  </definedNames>
                  <calcPr calcId="191029" calcMode="auto" fullCalcOnLoad="1" forceFullCalc="0"/>
                </workbook>
                """,
            ["xl/sharedStrings.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <si><r><t xml:space="preserve">No</t></r><r><t>rth</t></r></si>
                  <si><t>South</t></si>
                  <si><t>West</t></si>
                  <si><t>Share</t></si>
                  <si><t>Quoted ] Amount</t></si>
                </sst>
                """,
            ["xl/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <numFmts count="1"><numFmt numFmtId="165" formatCode="m/d/yy"/></numFmts>
                  <cellXfs count="6">
                    <xf numFmtId="0"/>
                    <xf numFmtId="1"/>
                    <xf numFmtId="2"/>
                    <xf numFmtId="14"/>
                    <xf numFmtId="49"/>
                    <xf numFmtId="165" applyNumberFormat="1"/>
                  </cellXfs>
                </styleSheet>
                """,
            ["xl/worksheets/sheet1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <cols><col min="2" max="2" hidden="1"/></cols>
                  <sheetData>
                    <row r="1"><c r="B1" t="s"><v>3</v></c><c r="C1" t="s"><v>4</v></c></row>
                    <row r="2"><c r="A2" t="s"><v>0</v></c><c r="B2" s="5"><v>8.2</v></c><c r="C2" s="4"><f t="array" ref="C2:C2" ca="1">NA()</f></c></row>
                    <row r="3" hidden="1"><c r="A3" t="s"><v>1</v></c><c r="B3"><v>3.2</v></c><c r="C3"><v>7.5</v></c></row>
                    <row r="4"><c r="A4" t="inlineStr"><is><t>West</t></is></c><c r="B4"><f t="shared" ref="B4:B4" si="7">B2-B3</f><v>1.4</v></c><c r="C4"><v>9.1</v></c></row>
                  </sheetData>
                  <tableParts count="2"><tablePart r:id="rId1"/><tablePart r:id="rId2"/></tableParts>
                </worksheet>
                """,
            ["xl/tables/table1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                       id="1" name="SalesTable" displayName="Sales_Table" ref="A1:C4" headerRowCount="1" totalsRowShown="0">
                  <autoFilter ref="A1:C4">
                    <filterColumn colId="0" showButton="1"><filters><filter val="North"/></filters></filterColumn>
                    <filterColumn colId="1"><customFilters><customFilter operator="greaterThan" val="3"/></customFilters></filterColumn>
                  </autoFilter>
                  <tableColumns count="3">
                    <tableColumn id="1" name="Region" totalsRowFunction="none"/>
                    <tableColumn id="2" name="Amount"><calculatedColumnFormula>B2</calculatedColumnFormula></tableColumn>
                    <tableColumn id="3" name="Quoted ] Amount"/>
                  </tableColumns>
                </table>
                """,
            ["xl/tables/table2.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                       id="2" name="SalesTotalsTable" displayName="Sales_Totals_Table" ref="A1:B4" headerRowCount="1" totalsRowCount="1" totalsRowShown="1">
                  <tableColumns count="2">
                    <tableColumn id="1" name="Region"/>
                    <tableColumn id="2" name="Amount" totalsRowFunction="sum"><totalsRowFormula>SUBTOTAL(109,[Amount])</totalsRowFormula></tableColumn>
                  </tableColumns>
                </table>
                """
        });
        return stream.ToArray();
    }

    internal static string RenderSyntheticMathSpacingProbe(bool highlightMiddleRun)
    {
        string middleRunProperties = highlightMiddleRun
            ? """<a:rPr sz="2000" b="1" i="1"><a:latin typeface="Cambria Math"/><a:highlight><a:srgbClr val="FFFF00"/></a:highlight></a:rPr>"""
            : """<a:rPr sz="2000" b="1" i="1"><a:latin typeface="Cambria Math"/></a:rPr>""";
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = $"""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr wrap="square"/><a:lstStyle/>
                      <a:p>
                        <a:r><a:rPr sz="2000" b="1" i="1"><a:latin typeface="Cambria Math"/></a:rPr><a:t>Alpha beta gamma </a:t></a:r>
                        <a:r>{middleRunProperties}<a:t>focus</a:t></a:r>
                        <a:r><a:rPr sz="2000" b="1" i="1"><a:latin typeface="Cambria Math"/></a:rPr><a:t> delta epsilon</a:t></a:r>
                      </a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        return File.ReadAllText(output, Encoding.ASCII);
    }

    internal static string BasicContentTypes()
    {
        return """
            <?xml version="1.0" encoding="UTF-8"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
              <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
            </Types>
            """;
    }

    internal static string PackageRelationship()
    {
        return """
            <?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
            </Relationships>
            """;
    }

    internal static string PresentationRelationship()
    {
        return """
            <?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
            </Relationships>
            """;
    }

    internal static string BasicPresentation()
    {
        return """
            <?xml version="1.0" encoding="UTF-8"?>
            <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <p:sldSz cx="9144000" cy="6858000"/>
              <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
            </p:presentation>
            """;
    }

    internal static string InheritedShapePart(string color, int x, int y)
    {
        return $$"""
            <?xml version="1.0" encoding="UTF-8"?>
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree><p:sp>
                <p:spPr>
                  <a:xfrm><a:off x="{{x}}" y="{{y}}"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                  <a:prstGeom prst="rect"/>
                  <a:solidFill><a:srgbClr val="{{color}}"/></a:solidFill>
                </p:spPr>
              </p:sp></p:spTree></p:cSld>
            </p:sld>
            """;
    }

    internal static int CountOccurrences(string text, string value)
    {
        int count = 0;
        int start = 0;
        while (true)
        {
            int index = text.IndexOf(value, start, StringComparison.Ordinal);
            if (index < 0)
            {
                return count;
            }

            count++;
            start = index + value.Length;
        }
    }

    internal static void AssertContainsTextMatrixAtX(string pdf, double x)
    {
        string xText = FormatPdfNumber(x);
        TestAssert.True(Regex.IsMatch(pdf, $@"1 0 0 1 {Regex.Escape(xText)} [0-9.]+ Tm"), $"Expected a text matrix at x={xText}.");
    }

    internal static void AssertDoesNotContainTextMatrixAtX(string pdf, double x, string message)
    {
        string xText = FormatPdfNumber(x);
        TestAssert.True(!Regex.IsMatch(pdf, $@"1 0 0 1 {Regex.Escape(xText)} [0-9.]+ Tm"), message);
    }

    internal static string FormatPdfNumber(double value)
    {
        return value.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);
    }
}
