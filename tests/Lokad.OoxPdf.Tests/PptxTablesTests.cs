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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(calibri))");
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
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

    public static void PptxSyntheticTableKeepsDeclaredRowsUnderLargeFrameSlack()
    {
        // Office keeps declared row heights when the frame carries large slack
        // (probed at 1.2x-3x frame-to-declared ratios); rows must not stretch to fill
        // the frame. Single row declared at 72pt inside a 144pt frame stays 72pt.
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
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                        <a:tbl>
                          <a:tblGrid><a:gridCol w="1828800"/><a:gridCol w="1828800"/></a:tblGrid>
                          <a:tr h="914400">
                            <a:tc>
                              <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>One</a:t></a:r></a:p></a:txBody>
                              <a:tcPr><a:solidFill><a:srgbClr val="D9EAD3"/></a:solidFill></a:tcPr>
                            </a:tc>
                            <a:tc>
                              <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>Two</a:t></a:r></a:p></a:txBody>
                              <a:tcPr><a:solidFill><a:srgbClr val="D9EAD3"/></a:solidFill></a:tcPr>
                            </a:tc>
                          </a:tr>
                        </a:tbl>
                      </a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        Match declaredFill = Regex.Matches(pdf, "([0-9.]+) ([0-9.]+) ([0-9.]+) ([0-9.]+) re f").Cast<Match>().FirstOrDefault(match => double.Parse(match.Groups[3].Value, System.Globalization.CultureInfo.InvariantCulture) > 100d && double.Parse(match.Groups[4].Value, System.Globalization.CultureInfo.InvariantCulture) >= 70d && double.Parse(match.Groups[4].Value, System.Globalization.CultureInfo.InvariantCulture) <= 74d) ?? Match.Empty;
        TestAssert.True(declaredFill.Success, "Expected a cell fill rect at declared 72pt row height under 2x frame slack.");
    }

    public static void PptxSyntheticTableKeepsDeclaredRowsWithoutMaterialSlack()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
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
        using var input = TestFixtures.CreateZipPackage(new Dictionary<string, byte[]>
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
        using var package = OoxPackage.Open(input, default);
        PdfPage page = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
            .RenderPagesAsync(new PptxReader().Read(package, default), package, null, default)
            .ToBlockingEnumerable().Single();

        AssertCellFills(page.Content, new RgbColor(51, 102, 153),
        [
            "72 420 72 48 re f*", "144 420 72 48 re f*", "216 420 72 48 re f*",
            "72 372 72 48 re f*", "216 372 72 48 re f*",
            "72 324 72 48 re f*", "216 324 72 48 re f*",
            "72 276 72 48 re f*", "144 276 72 48 re f*", "216 276 72 48 re f*",
            "324 396 72 72 re f*", "396 396 72 72 re f*", "324 324 72 72 re f*"
        ]);
        AssertCellFills(page.Content, new RgbColor(205, 211, 222), ["144 372 72 48 re f*"]);
        AssertCellFills(page.Content, new RgbColor(232, 234, 239), ["144 324 72 48 re f*"]);
        AssertCellFills(page.Content, new RgbColor(39, 80, 121), ["504 324 108 72 re f*"]);

        PdfExtGStateResource translucent = page.ExtGStates.Single(state => state.FillAlpha == 0.4d && state.StrokeAlpha == 1d);
        string content = page.Content.Replace("\r\n", "\n", StringComparison.Ordinal);
        TestAssert.Contains($"q\n/{translucent.ResourceName} gs\n0.2 0.4 0.6 rg\n324 324 72 72 re f*\nQ", content);
        var textPaint = ReadCellTextPaint(page);
        TestAssert.Equal("A B C D D1 D2 E F G H H1 H2 H3 I L1 L2 L3 L4",
            string.Join(" ", textPaint.Keys.Order(StringComparer.Ordinal)));
        TestAssert.Equal("1 g", textPaint["D1"]);
        TestAssert.Equal("1 g", textPaint["D2"]);
    }

    public static void PptxSyntheticTableStyleDarkStyle1Accent6RendersWhiteBodyText()
    {
        PdfPage page = RenderDarkStyleAccent6Table("""firstRow="1" bandRow="1" """,
            ["HDR"], ["BOD1"], ["BOD2"], ["BOD3"]);
        var textPaint = ReadCellTextPaint(page);
        TestAssert.Equal(4, textPaint.Count);
        foreach (string label in new[] { "HDR", "BOD1", "BOD2", "BOD3" })
            TestAssert.Equal("1 g", textPaint[label]);
    }

    public static void PptxSyntheticTableStyleDarkStyle1Accent6RendersLinearColumnShades()
    {
        PdfPage page = RenderDarkStyleAccent6Table("""firstCol="1" lastCol="1" """,
            ["T1", "T2", "T3"], ["T4", "T5", "T6"]);
        var shade = new RgbColor(39, 80, 121);
        string[] outerCells = ["72 420 72 48 re f*", "216 420 72 48 re f*", "72 372 72 48 re f*", "216 372 72 48 re f*"];
        AssertCellFills(page.Content, shade, outerCells);

        // Exchange the first row's middle and right colors while keeping the
        // number of rectangles in each color exactly the same.
        string separator = Environment.NewLine;
        string medium = "0.2 0.4 0.6 rg" + separator;
        string dark = "0.153 0.314 0.475 rg" + separator;
        string middleCell = "144 420 72 48 re f*";
        string rightCell = outerCells[1];
        TestAssert.Contains(medium + middleCell, page.Content);
        TestAssert.Contains(dark + rightCell, page.Content);
        string swapped = page.Content.Replace(medium + middleCell, dark + middleCell, StringComparison.Ordinal)
            .Replace(dark + rightCell, medium + rightCell, StringComparison.Ordinal);
        TestAssert.Throws<InvalidOperationException>(() => AssertCellFills(swapped, shade, outerCells));
    }

    public static void PptxSyntheticTableStyleDarkStyle1Accent6RendersShadedLastRow()
    {
        PdfPage page = RenderDarkStyleAccent6Table("""lastRow="1" """,
            ["U1", "U2"], ["U3", "U4"], ["L1", "L2"]);
        AssertCellFills(page.Content, new RgbColor(30, 65, 100), ["72 324 72 48 re f*", "144 324 72 48 re f*"]);
    }

    private static void AssertCellFills(string content, RgbColor color, string[] expected)
    {
        var paths = new List<string>();
        RgbColor[] colors = PptxChartPaletteTests.ReadColoredPaints(content, paths.Add).ToArray();
        var actual = paths.Where((path, index) => colors[index] == color &&
            path.EndsWith(" re f*", StringComparison.Ordinal)).Order(StringComparer.Ordinal);
        TestAssert.Equal(string.Join(", ", expected.Order(StringComparer.Ordinal)), string.Join(", ", actual));
    }

    private static Dictionary<string, string> ReadCellTextPaint(PdfPage page) =>
        Regex.Matches(page.Content, @"(?ms)^BT\r?\n(?<fill>[^\r\n]+)\r?\n.*?^ET\r?$")
            .ToDictionary(block => string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(
                new PdfPage(page.Width, page.Height, block.Value, page.Fonts)).Select(show => show.Text)),
                block => block.Groups["fill"].Value, StringComparer.Ordinal);

    private static PdfPage RenderDarkStyleAccent6Table(string tableFlags, params string[][] rows)
    {
        int columns = rows[0].Length;
        int columnWidth = columns == 1 ? 1371600 : 914400;
        const int rowHeight = 609600;
        string gridColumn = FormattableString.Invariant($"""<a:gridCol w="{columnWidth}"/>""");
        string grid = string.Concat(Enumerable.Repeat(gridColumn, columns));
        var tableRows = new StringBuilder();
        foreach (string[] row in rows)
        {
            tableRows.Append(FormattableString.Invariant($"""<a:tr h="{rowHeight}">"""));
            foreach (string label in row)
            {
                tableRows.Append($"""
                    <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>{label}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                    """);
            }
            tableRows.Append("</a:tr>");
        }

        var entries = new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes().Replace("</Types>",
                """<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>""",
                StringComparison.Ordinal),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship().Replace("</Relationships>", """
                <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>
                """, StringComparison.Ordinal),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/theme/theme1.xml"] = """
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
                """,
            ["ppt/slides/slide1.xml"] = FormattableString.Invariant($$"""
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:graphicFrame>
                    <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="{{columnWidth * columns}}" cy="{{rowHeight * rows.Length}}"/></p:xfrm>
                    <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
                      <a:tblPr {{tableFlags}}><a:tableStyleId>{AF606853-7671-496A-8E4F-DF71F8EC918B}</a:tableStyleId></a:tblPr>
                      <a:tblGrid>{{grid}}</a:tblGrid>
                      {{tableRows}}
                    </a:tbl></a:graphicData></a:graphic>
                  </p:graphicFrame></p:spTree></p:cSld>
                </p:sld>
                """)
        };

        using var input = TestFixtures.CreateZipPackage(entries);
        using var package = OoxPackage.Open(input, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver())
            .RenderPagesAsync(new PptxReader().Read(package, default), package, null, default)
            .ToBlockingEnumerable().Single();
    }

    public static void PptxSyntheticTableWrapsCellTextToColumnWidth()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(cambria))");
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(cambria))");
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

    public static void PptxTableCellLayoutReusesSourceAndTrimsOnlyLeadingEmptyParagraphs()
    {
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XElement emptyParagraph = new(a + "p", new XElement(a + "endParaRPr", new XAttribute("sz", "1200")));
        var empty = new XElement(a + "txBody", new XElement(a + "bodyPr"), emptyParagraph, emptyParagraph);
        var ordinary = new XElement(a + "txBody", new XElement(a + "bodyPr"),
            new XElement(a + "p", new XElement(a + "r", new XElement(a + "t", "Ordinary text"))), emptyParagraph);
        foreach (XElement source in new[] { ordinary, empty })
        {
            PptxSceneTableCell cell = ReadCell(source);
            TestAssert.True(ReferenceEquals(cell.TextBody, cell.LayoutTextBody),
                "Ordinary and entirely empty cells should share their untrimmed source body.");
        }

        foreach (XElement content in new[]
        {
            new XElement(a + "r", new XElement(a + "t", "Retained text")),
            new XElement(a + "fld", new XAttribute("type", "slidenum"), new XElement(a + "t", "3")),
            new XElement(a + "br")
        })
        {
            var effective = new XElement(a + "txBody", new XElement(a + "bodyPr"),
                new XElement(a + "p", content), emptyParagraph);
            var source = new XElement(effective);
            source.Element(a + "p")!.AddBeforeSelf(emptyParagraph);

            PptxSceneTableCell cell = ReadCell(source);
            TestAssert.True(XElement.DeepEquals(effective, cell.LayoutTextBody),
                $"Layout should retain the {content.Name.LocalName} paragraph and its trailing paragraph.");
            TestAssert.True(ReferenceEquals(source, cell.TextBody) &&
                XElement.DeepEquals(emptyParagraph, source.Elements(a + "p").First()),
                "Trimming the layout view must preserve the authored leading paragraph.");
        }

        PptxSceneTableCell ReadCell(XElement body) => PptxSceneBuilder.ReadTableCell(
            new XElement(a + "tc", body), PptxTheme.Empty, PptxColorMap.Default, default, 0, 0, 1, 1);
    }

    public static void PptxSyntheticTableIgnoresLeadingEmptyCellParagraph()
    {
        string cambria = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc");
        if (!File.Exists(cambria))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(cambria))");
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
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
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
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
