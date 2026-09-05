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

internal static class PptxTablesTests
{
    public static void PptxSyntheticTableKeepsAveragePdfSpacingResidualInPositioningArray()
    {
        string calibri = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "calibri.ttf");
        if (!File.Exists(calibri))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
                        <a:tblGrid><a:gridCol w="1828800"/></a:tblGrid>
                        <a:tr h="914400"><a:tc>
                          <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"><a:latin typeface="Calibri"/></a:rPr><a:t>$120K</a:t></a:r></a:p></a:txBody>
                          <a:tcPr/>
                        </a:tc></a:tr>
                      </a:tbl></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxTextGlyphRunSnapshot glyphRun = PptxRenderer.InspectTextGlyphRuns(document, package, 0).Single();
        TestAssert.True(glyphRun.TableRowIndex == 0, "Expected the inspected run to come from the table cell.");
        TestAssert.True(Math.Abs(glyphRun.PdfCharacterSpacing) < 0.001d, $"Expected table text to keep residual PDF spacing in TJ positioning instead of promoting it to Tc, got {glyphRun.PdfCharacterSpacing.ToString("0.###", CultureInfo.InvariantCulture)}.");

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 Tc", pdf);
        TestAssert.Contains(" TJ", pdf);
    }

    public static void PptxSyntheticTableRendersGridAndText()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="1828800"/></p:xfrm>
                      <a:graphic>
                        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                          <a:tbl>
                            <a:tblGrid><a:gridCol w="1828800"/><a:gridCol w="1828800"/></a:tblGrid>
                            <a:tr h="914400">
                              <a:tc>
                                <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>One</a:t></a:r></a:p></a:txBody>
                                <a:tcPr>
                                  <a:solidFill><a:srgbClr val="D9EAD3"/></a:solidFill>
                                  <a:lnL><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnL>
                                  <a:lnR><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnR>
                                  <a:lnT><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnT>
                                  <a:lnB><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnB>
                                </a:tcPr>
                              </a:tc>
                              <a:tc>
                                <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>Two</a:t></a:r></a:p></a:txBody>
                                <a:tcPr/>
                              </a:tc>
                            </a:tr>
                            <a:tr h="914400">
                              <a:tc>
                                <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>Three</a:t></a:r></a:p></a:txBody>
                                <a:tcPr/>
                              </a:tc>
                              <a:tc>
                                <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>Four</a:t></a:r></a:p></a:txBody>
                                <a:tcPr><a:solidFill><a:srgbClr val="FCE5CD"/></a:solidFill></a:tcPr>
                              </a:tc>
                            </a:tr>
                          </a:tbl>
                        </a:graphicData>
                      </a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.851 0.918 0.827 rg", pdf);
        TestAssert.Contains("72 396 144 72 re f*", pdf);
        TestAssert.Contains("0 G", pdf);
        TestAssert.Contains("/Subtype /Type0", pdf);
        TestAssert.Contains(" TJ", pdf);
    }

    public static void PptxSyntheticTableExpandsSlackRowsFromCellText()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="2286000"/></p:xfrm>
                      <a:graphic>
                        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                          <a:tbl>
                            <a:tblGrid><a:gridCol w="1828800"/></a:tblGrid>
                            <a:tr h="127000">
                              <a:tc>
                                <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"/><a:t>Alpha Beta Gamma Delta Epsilon Zeta</a:t></a:r></a:p></a:txBody>
                                <a:tcPr><a:solidFill><a:srgbClr val="D9EAD3"/></a:solidFill></a:tcPr>
                              </a:tc>
                            </a:tr>
                            <a:tr h="1016000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Second</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>
                            <a:tr h="1016000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Third</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>
                          </a:tbl>
                        </a:graphicData>
                      </a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        Match headerFill = Regex.Matches(pdf, @"72 (?<y>[0-9.]+) 144 (?<height>[0-9.]+) re f")
            .Cast<Match>()
            .FirstOrDefault(match => double.Parse(match.Groups["height"].Value, CultureInfo.InvariantCulture) > 20d) ?? Match.Empty;

        TestAssert.True(headerFill.Success, "Expected text minimum height to expand a slack table's first row beyond its declared proportional height.");
        double headerY = double.Parse(headerFill.Groups["y"].Value, CultureInfo.InvariantCulture);
        double headerHeight = double.Parse(headerFill.Groups["height"].Value, CultureInfo.InvariantCulture);
        TestAssert.True(headerY < 448d, $"Expected expanded first row to move the lower edge well below the proportional 457.412pt row boundary, got y={headerY.ToString("0.###", CultureInfo.InvariantCulture)}.");
        TestAssert.True(headerHeight > 20d, $"Expected expanded first row to exceed 20pt, got {headerHeight.ToString("0.###", CultureInfo.InvariantCulture)}pt.");
        TestAssert.DoesNotContain("72 457.412 144 10.588 re f", pdf);
    }

    public static void PptxSyntheticTableKeepsDeclaredRowsWithoutMaterialSlack()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="2095500"/></p:xfrm>
                      <a:graphic>
                        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                          <a:tbl>
                            <a:tblGrid><a:gridCol w="1828800"/></a:tblGrid>
                            <a:tr h="127000">
                              <a:tc>
                                <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"/><a:t>Alpha Beta Gamma Delta Epsilon Zeta</a:t></a:r></a:p></a:txBody>
                                <a:tcPr><a:solidFill><a:srgbClr val="D9EAD3"/></a:solidFill></a:tcPr>
                              </a:tc>
                            </a:tr>
                            <a:tr h="1016000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Second</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>
                            <a:tr h="1016000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Third</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>
                          </a:tbl>
                        </a:graphicData>
                      </a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        Match headerFill = Regex.Matches(pdf, @"72 (?<y>[0-9.]+) 144 (?<height>[0-9.]+) re f")
            .Cast<Match>()
            .FirstOrDefault() ?? Match.Empty;

        TestAssert.True(headerFill.Success, "Expected the first table row fill rectangle to be emitted.");
        double headerY = double.Parse(headerFill.Groups["y"].Value, CultureInfo.InvariantCulture);
        double headerHeight = double.Parse(headerFill.Groups["height"].Value, CultureInfo.InvariantCulture);
        TestAssert.True(headerY > 455d, $"Expected low-slack table to keep the declared proportional first row boundary, got y={headerY.ToString("0.###", CultureInfo.InvariantCulture)}.");
        TestAssert.True(headerHeight < 20d, $"Expected low-slack table to keep the compact declared first row height, got {headerHeight.ToString("0.###", CultureInfo.InvariantCulture)}pt.");
        TestAssert.DoesNotContain("72 420", pdf);
    }

    public static void PptxSyntheticTableKeepsDeclaredRowsForSmallPositiveUniqueSlack()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="1929136"/></p:xfrm>
                      <a:graphic>
                        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                          <a:tbl>
                            <a:tblGrid><a:gridCol w="1828800"/></a:tblGrid>
                            <a:tr h="914400">
                              <a:tc>
                                <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1200"/><a:t>First</a:t></a:r></a:p></a:txBody>
                                <a:tcPr><a:solidFill><a:srgbClr val="D9EAD3"/></a:solidFill></a:tcPr>
                              </a:tc>
                            </a:tr>
                            <a:tr h="965200"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1200"/><a:t>Second</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>
                          </a:tbl>
                        </a:graphicData>
                      </a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        Match headerFill = Regex.Matches(pdf, @"72 (?<y>[0-9.]+) 144 (?<height>[0-9.]+) re f")
            .Cast<Match>()
            .FirstOrDefault() ?? Match.Empty;

        TestAssert.True(headerFill.Success, "Expected the first table row fill rectangle to be emitted.");
        double headerY = double.Parse(headerFill.Groups["y"].Value, CultureInfo.InvariantCulture);
        double headerHeight = double.Parse(headerFill.Groups["height"].Value, CultureInfo.InvariantCulture);
        TestAssert.True(Math.Abs(headerY - 396d) < 0.01d, $"Expected small positive unique slack to keep the declared first-row lower edge at 396pt, got y={headerY.ToString("0.###", CultureInfo.InvariantCulture)}.");
        TestAssert.True(Math.Abs(headerHeight - 72d) < 0.01d, $"Expected small positive unique slack to keep the declared first-row height at 72pt, got {headerHeight.ToString("0.###", CultureInfo.InvariantCulture)}pt.");
        TestAssert.DoesNotContain("72 394.", pdf);
    }

    public static void PptxSyntheticTableKeepsDeclaredRowsForModerateSlack()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="1955800"/></p:xfrm>
                      <a:graphic>
                        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                          <a:tbl>
                            <a:tblGrid><a:gridCol w="1828800"/></a:tblGrid>
                            <a:tr h="914400">
                              <a:tc>
                                <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1200"/><a:t>First</a:t></a:r></a:p></a:txBody>
                                <a:tcPr><a:solidFill><a:srgbClr val="D9EAD3"/></a:solidFill></a:tcPr>
                              </a:tc>
                            </a:tr>
                            <a:tr h="914400">
                              <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1200"/><a:t>Second</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                            </a:tr>
                          </a:tbl>
                        </a:graphicData>
                      </a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        Match headerFill = Regex.Matches(pdf, @"72 (?<y>[0-9.]+) 144 (?<height>[0-9.]+) re f")
            .Cast<Match>()
            .FirstOrDefault() ?? Match.Empty;

        TestAssert.True(headerFill.Success, "Expected the first table row fill rectangle to be emitted.");
        double headerY = double.Parse(headerFill.Groups["y"].Value, CultureInfo.InvariantCulture);
        double headerHeight = double.Parse(headerFill.Groups["height"].Value, CultureInfo.InvariantCulture);
        TestAssert.True(headerY > 394d, $"Expected moderate slack to keep the declared first row boundary near 396pt, got y={headerY.ToString("0.###", CultureInfo.InvariantCulture)}.");
        TestAssert.True(headerHeight < 74d, $"Expected moderate slack to keep the declared first row height near 72pt, got {headerHeight.ToString("0.###", CultureInfo.InvariantCulture)}pt.");
        TestAssert.DoesNotContain("72 390.", pdf);
    }

    public static void PptxSyntheticTableKeepsOverflowingContentMinimumRowsUnscaled()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="508000"/></p:xfrm>
                      <a:graphic>
                        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                          <a:tbl>
                            <a:tblGrid><a:gridCol w="914400"/></a:tblGrid>
                            <a:tr h="63500">
                              <a:tc>
                                <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Alpha Beta Gamma Delta</a:t></a:r></a:p></a:txBody>
                                <a:tcPr/>
                              </a:tc>
                            </a:tr>
                            <a:tr h="63500">
                              <a:tc>
                                <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Epsilon Zeta Eta Theta</a:t></a:r></a:p></a:txBody>
                                <a:tcPr/>
                              </a:tc>
                            </a:tr>
                          </a:tbl>
                        </a:graphicData>
                      </a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxTextFrameModelSnapshot[] tableFrames = PptxRenderer.InspectTableTextFrameModels(document, package, 0).ToArray();

        TestAssert.Equal(2, tableFrames.Length);
        double totalRenderedRows = tableFrames.Sum(frame => frame.FrameHeight);
        TestAssert.True(totalRenderedRows > 40.01d, $"Expected overflowing content-minimum rows to exceed the 40pt graphicFrame instead of being scaled down, got {totalRenderedRows.ToString("0.###", CultureInfo.InvariantCulture)}pt.");
        TestAssert.True(tableFrames.All(frame => frame.FrameHeight > 20d), "Expected each row to preserve its content minimum rather than the declared 20pt row height.");
    }

    public static void PptxSyntheticTableTextFrameModelPreservesCellPropertySources()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></p:xfrm>
                      <a:graphic>
                        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                          <a:tbl>
                            <a:tblGrid><a:gridCol w="1828800"/><a:gridCol w="1828800"/></a:tblGrid>
                            <a:tr h="914400">
                              <a:tc>
                                <a:txBody><a:bodyPr tIns="45720"/><a:lstStyle/><a:p><a:r><a:t>Explicit</a:t></a:r></a:p></a:txBody>
                                <a:tcPr marL="182880" anchor="ctr"/>
                              </a:tc>
                              <a:tc>
                                <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Default</a:t></a:r></a:p></a:txBody>
                                <a:tcPr/>
                              </a:tc>
                            </a:tr>
                          </a:tbl>
                        </a:graphicData>
                      </a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextFrameModelSnapshot[] tableFrames = PptxRenderer.InspectTableTextFrameModels(document, package, 0).ToArray();
        TestAssert.Equal(2, tableFrames.Length);
        PptxTextFrameModelSnapshot explicitCell = tableFrames.Single(frame => frame.Paragraphs.Any(paragraph => paragraph.Runs.Any(run => run.Text == "Explicit")));
        PptxTextFrameModelSnapshot defaultCell = tableFrames.Single(frame => frame.Paragraphs.Any(paragraph => paragraph.Runs.Any(run => run.Text == "Default")));
        TestAssert.Equal("TableCellProperties", explicitCell.InsetLeftSource);
        TestAssert.Equal("182880", explicitCell.InsetLeftValue ?? string.Empty);
        TestAssert.Equal("DirectBodyPr", explicitCell.InsetTopSource);
        TestAssert.Equal("45720", explicitCell.InsetTopValue ?? string.Empty);
        TestAssert.Equal("TableCellProperties", explicitCell.VerticalAnchorSource);
        TestAssert.Equal("ctr", explicitCell.VerticalAnchorValue ?? string.Empty);
        TestAssert.True(Math.Abs((explicitCell.TableDeclaredRowHeight ?? 0d) - 72d) < 0.001d, $"Expected table frame inspection to preserve the 72pt declared row height, got {(explicitCell.TableDeclaredRowHeight ?? 0d).ToString("0.###", CultureInfo.InvariantCulture)}pt.");
        TestAssert.True(Math.Abs((explicitCell.TableDeclaredHeight ?? 0d) - 72d) < 0.001d, $"Expected table frame inspection to preserve the 72pt declared table height, got {(explicitCell.TableDeclaredHeight ?? 0d).ToString("0.###", CultureInfo.InvariantCulture)}pt.");
        TestAssert.True(Math.Abs((explicitCell.TableHeightSlackFactor ?? 0d) - 1d) < 0.001d, $"Expected no table-height slack in the one-row fixture, got {(explicitCell.TableHeightSlackFactor ?? 0d).ToString("0.###", CultureInfo.InvariantCulture)}.");
        TestAssert.True(Math.Abs(explicitCell.TextWrapWidth - explicitCell.TextWidth) < 0.001d, $"Expected explicit table-cell wrap width to match the inset text width, got wrap={explicitCell.TextWrapWidth.ToString("0.###", CultureInfo.InvariantCulture)} width={explicitCell.TextWidth.ToString("0.###", CultureInfo.InvariantCulture)}.");
        TestAssert.Equal("DefaultValue", defaultCell.VerticalAnchorSource);
        TestAssert.Equal("t", defaultCell.VerticalAnchorValue ?? string.Empty);
        TestAssert.True(Math.Abs((defaultCell.TableDeclaredRowSpanHeight ?? 0d) - 72d) < 0.001d, $"Expected table frame inspection to preserve the 72pt declared row-span height, got {(defaultCell.TableDeclaredRowSpanHeight ?? 0d).ToString("0.###", CultureInfo.InvariantCulture)}pt.");
        TestAssert.True(Math.Abs(defaultCell.TextWrapWidth - defaultCell.TextWidth) < 0.001d, $"Expected default table-cell wrap width to match the inset text width, got wrap={defaultCell.TextWrapWidth.ToString("0.###", CultureInfo.InvariantCulture)} width={defaultCell.TextWidth.ToString("0.###", CultureInfo.InvariantCulture)}.");
    }

    public static void PptxSyntheticUnsupportedTableStyleEmitsDiagnosticUntilCascadeExists()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(PptxTests.BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                        <a:tbl>
                          <a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{11111111-1111-1111-1111-111111111111}</a:tableStyleId></a:tblPr>
                          <a:tblGrid><a:gridCol w="1828800"/></a:tblGrid>
                          <a:tr h="914400">
                            <a:tc>
                              <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Unsupported style</a:t></a:r></a:p></a:txBody>
                              <a:tcPr><a:solidFill><a:srgbClr val="D9EAD3"/></a:solidFill></a:tcPr>
                            </a:tc>
                          </a:tr>
                        </a:tbl>
                      </a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.851 0.918 0.827 rg", pdf);
        TestAssert.True(diagnostics.Any(d =>
            d.Id == "PPTX_UNSUPPORTED_TABLE_STYLE" &&
            d.PartName == "/ppt/slides/slide1.xml" &&
            d.Feature == "table style" &&
            d.Fallback == "DefaultStyle"),
            "Unknown table styles should stay diagnostic-covered until the Office table-style cascade exists.");
    }

    public static void PptxSyntheticTableStyleMediumStyle2Accent6RendersFills()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(PptxTests.BasicContentTypes().Replace(
                "</Types>",
                "  <Override PartName=\"/ppt/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>\r\n</Types>",
                StringComparison.Ordinal)),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
                </Relationships>
                """),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/theme/theme1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="TableTheme">
                  <a:themeElements>
                    <a:clrScheme name="TableTheme">
                      <a:dk1><a:srgbClr val="000000"/></a:dk1>
                      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
                      <a:accent6><a:srgbClr val="336699"/></a:accent6>
                    </a:clrScheme>
                    <a:fontScheme name="TableTheme"><a:majorFont/><a:minorFont/></a:fontScheme>
                  </a:themeElements>
                </a:theme>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="1828800"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                        <a:tbl>
                          <a:tblPr firstRow="1" firstCol="1" lastRow="1" lastCol="1" bandRow="1"><a:tableStyleId>{93296810-A885-4BE3-A3E7-6D5BEEA58F35}</a:tableStyleId></a:tblPr>
                          <a:tblGrid><a:gridCol w="914400"/><a:gridCol w="914400"/><a:gridCol w="914400"/></a:tblGrid>
                          <a:tr h="609600">
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>H1</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>H2</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>H3</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                          </a:tr>
                          <a:tr h="609600">
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>A</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>B</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>C</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                          </a:tr>
                          <a:tr h="609600">
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>G</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>H</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>I</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                          </a:tr>
                          <a:tr h="609600">
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>D</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>E</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>F</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                          </a:tr>
                        </a:tbl>
                      </a:graphicData></a:graphic>
                    </p:graphicFrame>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="4114800" y="914400"/><a:ext cx="1828800" cy="1828800"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                        <a:tbl>
                          <a:tblPr firstRow="1"><a:bandCol/><a:tableStyleId>{68D230F3-CF80-4859-8CE7-A43EE81993B5}</a:tableStyleId></a:tblPr>
                          <a:tblGrid><a:gridCol w="914400"/><a:gridCol w="914400"/></a:tblGrid>
                          <a:tr h="914400">
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>L1</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>L2</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                          </a:tr>
                          <a:tr h="914400">
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>L3</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>L4</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                          </a:tr>
                        </a:tbl>
                      </a:graphicData></a:graphic>
                    </p:graphicFrame>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="6400800" y="914400"/><a:ext cx="1371600" cy="1828800"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                        <a:tbl>
                          <a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{AF606853-7671-496A-8E4F-DF71F8EC918B}</a:tableStyleId></a:tblPr>
                          <a:tblGrid><a:gridCol w="1371600"/></a:tblGrid>
                          <a:tr h="914400">
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>D1</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                          </a:tr>
                          <a:tr h="914400">
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>D2</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                          </a:tr>
                        </a:tbl>
                      </a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.2 0.4 0.6 rg", pdf);
        TestAssert.Contains("0.761 0.827 0.89 rg", pdf);
        TestAssert.Contains("0.886 0.918 0.949 rg", pdf);
        TestAssert.Contains("1 g", pdf);
        TestAssert.True(Regex.Matches(pdf, "0\\.2 0\\.4 0\\.6 rg").Count >= 3, "Expected header cells and first-column body cell to use the accent fill.");
        TestAssert.Contains("/GS40000F100000S gs", pdf);
        TestAssert.Contains("0.078 0.161 0.239 rg", pdf);
    }

    public static void PptxSyntheticTableWrapsCellTextToColumnWidth()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
                        <a:tblGrid><a:gridCol w="914400"/></a:tblGrid>
                        <a:tr h="914400">
                          <a:tc>
                            <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>Alpha Beta Gamma</a:t></a:r></a:p></a:txBody>
                            <a:tcPr/>
                          </a:tc>
                        </a:tr>
                      </a:tbl></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        int lineStarts = Regex.Matches(pdf, $@"1 0 0 1 {Regex.Escape(PptxTests.FormatPdfNumber(79.2d))} [0-9.]+ Tm").Count;
        TestAssert.True(lineStarts >= 2, "Expected narrow table-cell text to wrap onto multiple lines at the cell text inset.");
        TestAssert.Contains("72 396 72 72 re W* n", pdf);
        TestAssert.DoesNotContain("79.2 396 57.6 72 re W* n", pdf);
    }

    public static void PptxSyntheticTableKeepsSlide6HeaderOnOneLine()
    {
        string cambria = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc");
        if (!File.Exists(cambria))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="2453469" cy="700198"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
                        <a:tblGrid><a:gridCol w="2453469"/></a:tblGrid>
                        <a:tr h="700198">
                          <a:tc>
                            <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:buNone/></a:pPr><a:r><a:rPr lang="en-US" sz="1200" b="1"><a:latin typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/><a:ea typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/></a:rPr><a:t>Recurring decision Lokad automates</a:t></a:r></a:p></a:txBody>
                            <a:tcPr marL="46789" marR="46789" marT="23394" marB="23394" anchor="ctr"/>
                          </a:tc>
                        </a:tr>
                      </a:tbl></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });

        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        string[] matrices = Regex.Matches(pdf, @"1 0 0 1 [0-9.]+ [0-9.]+ Tm")
            .Select(match => match.Value)
            .ToArray();
        int lineBaselines = matrices
            .Select(matrix => Regex.Match(matrix, @"1 0 0 1 [0-9.]+ (?<y>[0-9.]+) Tm").Groups["y"].Value)
            .Distinct(StringComparer.Ordinal)
            .Count();
        TestAssert.True(lineBaselines == 1, $"Expected one rendered baseline for the table header; got {lineBaselines}. Matrices: {string.Join(" | ", matrices)}");
        TestAssert.True(Regex.IsMatch(pdf, @"1 0 0 1 75\.684 [0-9.]+ Tm"), $"The centered table header should measure and render at the table-cell inset. Matrices: {string.Join(" | ", matrices)}");
    }

    public static void PptxSyntheticCenteredTableCellUsesOfficeWrapSlack()
    {
        string cambria = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc");
        if (!File.Exists(cambria))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="4572000" y="914400"/><a:ext cx="1585000" cy="609600"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
                        <a:tblGrid><a:gridCol w="1585000"/></a:tblGrid>
                        <a:tr h="609600">
                          <a:tc>
                            <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="ctr"><a:buNone/></a:pPr><a:r><a:rPr lang="en-US" sz="1200"><a:latin typeface="Cambria Math"/><a:ea typeface="Cambria Math"/></a:rPr><a:t>Forecasting, Allocation, Planning</a:t></a:r></a:p></a:txBody>
                            <a:tcPr marL="36000" marR="36000" marT="18000" marB="18000" anchor="ctr"/>
                          </a:tc>
                        </a:tr>
                      </a:tbl></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });

        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        string[] matrices = Regex.Matches(pdf, @"1 0 0 1 [0-9.]+ [0-9.]+ Tm")
            .Select(match => match.Value)
            .ToArray();
        int lineBaselines = matrices
            .Select(matrix => Regex.Match(matrix, @"1 0 0 1 [0-9.]+ (?<y>[0-9.]+) Tm").Groups["y"].Value)
            .Distinct(StringComparer.Ordinal)
            .Count();
        TestAssert.True(lineBaselines == 2, $"Expected centered table text to keep the first two words on one line; got {lineBaselines}. Matrices: {string.Join(" | ", matrices)}");
    }

    public static void PptxSyntheticTableDistributesSmallPositiveSlackToShortestRows()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="4521200"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
                        <a:tblGrid><a:gridCol w="1828800"/></a:tblGrid>
                        <a:tr h="508000">
                          <a:tc>
                            <a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody>
                            <a:tcPr><a:solidFill><a:srgbClr val="D9EAD3"/></a:solidFill></a:tcPr>
                          </a:tc>
                        </a:tr>
                        <a:tr h="508000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc></a:tr>
                        <a:tr h="609600"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc></a:tr>
                        <a:tr h="889000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc></a:tr>
                        <a:tr h="762000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc></a:tr>
                        <a:tr h="1168400"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc></a:tr>
                      </a:tbl></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("72 425 144 43 re f", pdf);
        TestAssert.DoesNotContain("72 427.314 144 40.686 re f", pdf);
    }

    public static void PptxSyntheticTableIgnoresLeadingEmptyCellParagraph()
    {
        string cambria = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc");
        if (!File.Exists(cambria))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="1585169" cy="481032"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
                        <a:tblGrid><a:gridCol w="1585169"/></a:tblGrid>
                        <a:tr h="481032">
                          <a:tc>
                            <a:txBody>
                              <a:bodyPr/><a:lstStyle/>
                              <a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US" sz="1200" b="0"><a:latin typeface="Cambria Math"/><a:ea typeface="Cambria Math"/></a:endParaRPr></a:p>
                              <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1200" b="0"><a:latin typeface="Cambria Math"/><a:ea typeface="Cambria Math"/></a:rPr><a:t>Long-term Scheduling, Production</a:t></a:r></a:p>
                            </a:txBody>
                            <a:tcPr marL="35667" marR="35667" marT="17833" marB="17833" anchor="ctr"/>
                          </a:tc>
                        </a:tr>
                      </a:tbl></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
            PptxSceneTableCell cell = new PptxSceneBuilder().Build(document, package, CancellationToken.None).Slides[0].SlideNodes[0].Table?.Rows[0].Cells[0]
                ?? throw new InvalidOperationException("Expected table cell scene node.");

            TestAssert.Equal(1, cell.LeadingEmptyTextParagraphCount);
            TestAssert.Equal(2, cell.TextBody?.Elements(XName.Get("p", "http://schemas.openxmlformats.org/drawingml/2006/main")).Count() ?? 0);
            TestAssert.Equal(1, cell.LayoutTextBody?.Elements(XName.Get("p", "http://schemas.openxmlformats.org/drawingml/2006/main")).Count() ?? 0);
        }

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        string[] baselines = Regex.Matches(pdf, @"1 0 0 1 [0-9.]+ (?<y>[0-9.]+) Tm")
            .Select(match => match.Groups["y"].Value)
            .Distinct(StringComparer.Ordinal)
            .ToArray();
        TestAssert.Equal(2, baselines.Length);
    }

    public static void PptxSyntheticTableCentersTextByContentHeight()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="0"/><a:ext cx="1828800" cy="914400"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
                        <a:tblGrid><a:gridCol w="1828800"/></a:tblGrid>
                        <a:tr h="914400">
                          <a:tc>
                            <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1200"><a:latin typeface="Arial"/></a:rPr><a:t>Centered</a:t></a:r></a:p></a:txBody>
                            <a:tcPr anchor="ctr" marL="0" marR="0" marT="0" marB="0"/>
                          </a:tc>
                        </a:tr>
                      </a:tbl></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        string[] matrices = Regex.Matches(pdf, @"1 0 0 1 [0-9.]+ [0-9.]+ Tm")
            .Select(match => match.Value)
            .ToArray();
        TestAssert.True(Regex.IsMatch(pdf, @"1 0 0 1 72 499\.667 Tm"), $"Centered table-cell text should account for its line height before vertical anchoring. Matrices: {string.Join(" | ", matrices)}");
    }

    public static void PptxSyntheticTableStyleBoldContributesToCenteredTextHeight()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(PptxTests.BasicContentTypes().Replace(
                "</Types>",
                "  <Override PartName=\"/ppt/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>\r\n</Types>",
                StringComparison.Ordinal)),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
                </Relationships>
                """),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/theme/theme1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="TableTheme">
                  <a:themeElements>
                    <a:clrScheme name="TableTheme">
                      <a:dk1><a:srgbClr val="000000"/></a:dk1>
                      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
                      <a:accent6><a:srgbClr val="336699"/></a:accent6>
                    </a:clrScheme>
                    <a:fontScheme name="TableTheme"><a:majorFont/><a:minorFont/></a:fontScheme>
                  </a:themeElements>
                </a:theme>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="0"/><a:ext cx="1066800" cy="914400"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
                        <a:tblPr firstRow="1"><a:tableStyleId>{93296810-A885-4BE3-A3E7-6D5BEEA58F35}</a:tableStyleId></a:tblPr>
                        <a:tblGrid><a:gridCol w="533400"/><a:gridCol w="533400"/></a:tblGrid>
                        <a:tr h="914400">
                          <a:tc>
                            <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1200"><a:latin typeface="Arial"/></a:rPr><a:t>Next up</a:t></a:r></a:p></a:txBody>
                            <a:tcPr anchor="ctr" marL="0" marR="0" marT="0" marB="0"/>
                          </a:tc>
                          <a:tc>
                            <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1200" b="0"><a:latin typeface="Arial"/></a:rPr><a:t>Next up</a:t></a:r></a:p></a:txBody>
                            <a:tcPr anchor="ctr" marL="0" marR="0" marT="0" marB="0"/>
                          </a:tc>
                        </a:tr>
                      </a:tbl></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """)
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextGlyphRunSnapshot[] glyphRuns = PptxRenderer.InspectTextGlyphRuns(document, package, 0)
            .Where(run => run.Text.Contains("Next", StringComparison.Ordinal) || run.Text.Contains("up", StringComparison.Ordinal))
            .ToArray();
        PptxTextGlyphRunSnapshot[] styleBoldRuns = glyphRuns.Where(run => run.X < 100d).ToArray();
        PptxTextGlyphRunSnapshot[] directNonBoldRuns = glyphRuns.Where(run => run.X > 100d).ToArray();
        string diagnostics = string.Join(" | ", glyphRuns.Select(run => $"{run.Text}@{run.X:0.###},{run.BaselineY:0.###},w{run.Width:0.###},line{run.LineIndex}"));

        TestAssert.True(styleBoldRuns.Select(run => run.LineIndex).Distinct().Count() == 2, "Expected inherited table-style bold to affect wrapping. Runs: " + diagnostics);
        TestAssert.True(directNonBoldRuns.Select(run => run.LineIndex).Distinct().Count() == 1, "Expected explicit b=\"0\" to keep the comparison cell on one line. Runs: " + diagnostics);
        TestAssert.True(styleBoldRuns.Max(run => run.BaselineY) > directNonBoldRuns.Max(run => run.BaselineY) + 2d, "Expected table-style bold to feed the centered text-height estimate. Runs: " + diagnostics);
    }

    public static void PptxSyntheticTableMergedCellsSuppressInteriorGrid()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="1828800"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
                        <a:tblGrid><a:gridCol w="1828800"/><a:gridCol w="1828800"/></a:tblGrid>
                        <a:tr h="914400">
                          <a:tc gridSpan="2">
                            <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>Wide</a:t></a:r></a:p></a:txBody>
                            <a:tcPr><a:solidFill><a:srgbClr val="D9EAD3"/></a:solidFill></a:tcPr>
                          </a:tc>
                          <a:tc hMerge="1"><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc>
                        </a:tr>
                        <a:tr h="914400">
                          <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>Left</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                          <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>Right</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                        </a:tr>
                      </a:tbl></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("72 396 288 72 re f", pdf);
        TestAssert.Contains("216 396.5 m 216 324 l S", pdf);
        TestAssert.DoesNotContain("216 468.5 m 216 396.5 l S", pdf);
    }

    public static void PptxSyntheticExplicitTableBordersClampAtOuterEdges()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
                        <a:tblGrid><a:gridCol w="914400"/><a:gridCol w="914400"/></a:tblGrid>
                        <a:tr h="914400">
                          <a:tc>
                            <a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody>
                            <a:tcPr>
                              <a:lnL><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnL>
                              <a:lnB><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnB>
                            </a:tcPr>
                          </a:tc>
                          <a:tc>
                            <a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody>
                            <a:tcPr>
                              <a:lnB><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnB>
                            </a:tcPr>
                          </a:tc>
                        </a:tr>
                      </a:tbl></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("72 396 m 216 396 l S", pdf);
        TestAssert.Contains("72 396 m 72 468 l S", pdf);
        TestAssert.DoesNotContain("71.5 396 m 216.5 396 l S", pdf);
        TestAssert.DoesNotContain("72 395.5 m 72 468.5 l S", pdf);
    }
}
