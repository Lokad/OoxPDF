using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxImagesTests
{
    public static void DocxSyntheticInlinePngRendersImageXObject()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """),
            ["word/_rels/document.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
                </Relationships>
                """),
            ["word/document.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p>
                      <w:r>
                        <w:drawing>
                          <wp:inline>
                            <wp:extent cx="1828800" cy="914400"/>
                            <a:graphic>
                              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:pic><pic:blipFill><a:blip r:embed="rIdImage1"/></pic:blipFill></pic:pic>
                              </a:graphicData>
                            </a:graphic>
                          </wp:inline>
                        </w:drawing>
                      </w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """),
            ["word/media/image1.png"] = TestFixtures.CreateRgbPng(2, 1, [255, 0, 0, 0, 0, 255])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Subtype /Image", pdf);
        TestAssert.Contains("/XObject", pdf);
        TestAssert.Contains("/Im1 Do", pdf);
        TestAssert.Contains("/Width 2 /Height 1", pdf);
    }

    public static void DocxSyntheticVmlInlinePngRendersImageXObject()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """),
            ["word/_rels/document.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
                </Relationships>
                """),
            ["word/document.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                            xmlns:v="urn:schemas-microsoft-com:vml">
                  <w:body>
                    <w:p>
                      <w:r>
                        <w:pict>
                          <v:shape id="vml-image" style="width:72pt;height:36pt">
                            <v:imagedata r:id="rIdImage1"/>
                          </v:shape>
                        </w:pict>
                      </w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """),
            ["word/media/image1.png"] = TestFixtures.CreateRgbPng(2, 1, [32, 64, 96, 128, 160, 192])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.DoesNotContain("DOCX_UNSUPPORTED_VML", ids);
        TestAssert.Contains("/Subtype /Image", pdf);
        TestAssert.Contains("/Im1 Do", pdf);
        TestAssert.Contains("/Width 2 /Height 1", pdf);
    }

    public static void DocxEmissionObservesMidConversionCancellation()
    {
        // Q01: cancelling while DOCX emission is underway must fail fast with no
        // partial PDF. The sink cancels when the mid-document unsupported image
        // reports during emission; trailing paragraphs guarantee a checkpoint fires.
        var body = new StringBuilder();
        for (int i = 0; i < 150; i++)
        {
            body.Append("<w:p><w:r><w:t>para ").Append(i).Append("</w:t></w:r></w:p>");
        }

        body.Append("<w:p><w:r><w:drawing><wp:inline><wp:extent cx=\"914400\" cy=\"914400\"/><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic><pic:blipFill><a:blip r:embed=\"rIdImage1\"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>");
        for (int i = 0; i < 150; i++)
        {
            body.Append("<w:p><w:r><w:t>tail ").Append(i).Append("</w:t></w:r></w:p>");
        }

        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """),
            ["word/_rels/document.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
                </Relationships>
                """),
            ["word/document.xml"] = TestFixtures.Utf8(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""
                + " xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\""
                + " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""
                + " xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\""
                + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                + "<w:body>" + body.ToString()
                + "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>"
                + "</w:body></w:document>"),
            ["word/media/image1.png"] = TestFixtures.CreateUnsupportedHighBitDepthPng()
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        using var cancellation = new CancellationTokenSource();
        var options = new OoxPdfOptions
        {
            DiagnosticSink = diagnostic =>
            {
                if (diagnostic.Id == "IMAGE_UNSUPPORTED_FORMAT")
                {
                    cancellation.Cancel();
                }
            }
        };
        TestAssert.Throws<OperationCanceledException>(
            () => { OoxPdfConverter.Convert(input, output, options, cancellation.Token); });
        TestAssert.True(!File.Exists(output), "Mid-emission cancellation must not publish a partial PDF.");
    }

    public static void DocxUnsupportedPngImageEmitsDiagnostic()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """),
            ["word/_rels/document.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
                </Relationships>
                """),
            ["word/document.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdImage1"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """),
            ["word/media/image1.png"] = TestFixtures.CreateUnsupportedHighBitDepthPng()
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var collector = new DiagnosticCollector();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = collector.Add });

        TestAssert.True(File.Exists(output), "Unsupported DOCX image should not fail the whole conversion.");
        TestAssert.True(collector.Diagnostics.Any(d => d.Id == "IMAGE_UNSUPPORTED_FORMAT" && d.Severity == OoxPdfSeverity.Error && d.PartName == "/word/media/image1.png"), "Unsupported DOCX image should emit a release-blocking diagnostic.");
    }

    public static void DocxMarkupModesFilterRevisedImagesAndFloatingDrawings()
    {
        string input = DocxTests.WriteImageRevisionProbeDocx();

        DocxDocument finalDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
        DocxDocument originalDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Original);
        DocxDocument allDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);

        TestAssert.Equal(1, finalDocument.Paragraphs.Single().Images.Count);
        TestAssert.Equal("/word/media/inline-inserted.png", finalDocument.Paragraphs.Single().Images.Single().PartName ?? string.Empty);
        TestAssert.Equal(1, finalDocument.FloatingDrawings.Count);
        TestAssert.Equal("/word/media/floating-inserted.png", finalDocument.FloatingDrawings.Single().Image?.PartName ?? string.Empty);
        TestAssert.Equal(DocxRevisionKind.Insertion, finalDocument.Paragraphs.Single().Images.Single().Revisions.Single().Kind);
        TestAssert.Equal(DocxRevisionKind.Insertion, finalDocument.FloatingDrawings.Single().Revisions.Single().Kind);

        TestAssert.Equal(1, originalDocument.Paragraphs.Single().Images.Count);
        TestAssert.Equal("/word/media/inline-deleted.png", originalDocument.Paragraphs.Single().Images.Single().PartName ?? string.Empty);
        TestAssert.Equal(1, originalDocument.FloatingDrawings.Count);
        TestAssert.Equal("/word/media/floating-deleted.png", originalDocument.FloatingDrawings.Single().Image?.PartName ?? string.Empty);
        TestAssert.Equal(DocxRevisionKind.Deletion, originalDocument.Paragraphs.Single().Images.Single().Revisions.Single().Kind);
        TestAssert.Equal(DocxRevisionKind.Deletion, originalDocument.FloatingDrawings.Single().Revisions.Single().Kind);

        TestAssert.Equal(2, allDocument.Paragraphs.Single().Images.Count);
        TestAssert.Equal(2, allDocument.FloatingDrawings.Count);
        TestAssert.True(allDocument.Paragraphs.Single().Images.Any(image => image.Revisions.Single().Kind == DocxRevisionKind.Insertion) && allDocument.Paragraphs.Single().Images.Any(image => image.Revisions.Single().Kind == DocxRevisionKind.Deletion), "All-markup inline images should retain inserted and deleted image provenance.");
        TestAssert.True(allDocument.FloatingDrawings.Any(drawing => drawing.Revisions.Single().Kind == DocxRevisionKind.Insertion) && allDocument.FloatingDrawings.Any(drawing => drawing.Revisions.Single().Kind == DocxRevisionKind.Deletion), "All-markup floating drawings should retain inserted and deleted drawing provenance.");
    }

    public static void DocxMarkupModesFilterMovedAndCommentedImagesAndDrawings()
    {
        string input = DocxTests.WriteMovedCommentedImageProbeDocx();

        DocxDocument finalDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
        DocxDocument originalDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Original);
        DocxDocument allDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);

        AssertInlineImageParts(finalDocument, ["/word/media/body-commented.png", "/word/media/inline-move-to.png"], ["/word/media/inline-move-from.png"]);
        AssertInlineImageParts(originalDocument, ["/word/media/body-commented.png", "/word/media/inline-move-from.png"], ["/word/media/inline-move-to.png"]);
        AssertInlineImageParts(allDocument, ["/word/media/body-commented.png", "/word/media/inline-move-from.png", "/word/media/inline-move-to.png"], []);
        TestAssert.True(allDocument.Paragraphs.Single().CommentRanges.Any(range => range.Id == "1") && allDocument.Paragraphs.Single().InlineReferences.Any(reference => reference.Kind == DocxRelatedStoryKind.Comment && reference.Id == "1"), "Commented inline images should keep their body comment anchor metadata.");

        TestAssert.True(finalDocument.FloatingDrawings.Single().Image?.PartName == "/word/media/floating-move-to.png", "Final view should keep moved-to floating drawings.");
        TestAssert.True(originalDocument.FloatingDrawings.Single().Image?.PartName == "/word/media/floating-move-from.png", "Original view should keep moved-from floating drawings.");
        TestAssert.True(allDocument.FloatingDrawings.Any(drawing => drawing.Image?.PartName == "/word/media/floating-move-from.png" && drawing.Revisions.Single().Kind == DocxRevisionKind.MoveFrom), "All-markup should retain moved-from floating drawing provenance.");
        TestAssert.True(allDocument.FloatingDrawings.Any(drawing => drawing.Image?.PartName == "/word/media/floating-move-to.png" && drawing.Revisions.Single().Kind == DocxRevisionKind.MoveTo && drawing.BehindDocumentValue == "1"), "All-markup should retain moved-to behind-document floating drawing provenance.");

        DocxRelatedStory commentStory = allDocument.RelatedStories.Single(story => story.Kind == DocxRelatedStoryKind.Comment && story.Id == "1");
        DocxFloatingDrawing commentDrawing = commentStory.FloatingDrawings.Single();
        TestAssert.True(commentDrawing.ImageRelationshipId == "rIdCommentAnchor" && commentDrawing.Image?.PartName == "/word/media/comment-anchor.png", "Comment story anchored drawings should resolve through comment-part relationships.");
        TestAssert.True(commentDrawing.Revisions.Single().Kind == DocxRevisionKind.MoveTo && commentDrawing.BehindDocumentValue == "1", "Revised behind-document anchored drawings inside comment stories should retain revision and geometry provenance.");

        static void AssertInlineImageParts(DocxDocument document, IReadOnlyList<string> expected, IReadOnlyList<string> unexpected)
        {
            string[] parts = document.Paragraphs.Single().Images.Select(image => image.PartName ?? string.Empty).ToArray();
            foreach (string part in expected)
            {
                TestAssert.True(parts.Contains(part, StringComparer.Ordinal), "Expected inline image part was not visible in the selected markup view: " + part);
            }

            foreach (string part in unexpected)
            {
                TestAssert.True(!parts.Contains(part, StringComparer.Ordinal), "Unexpected inline image part survived markup filtering: " + part);
            }
        }
    }

    public static void DocxLayoutSnapshotReportsInlineImageSourceBlockIndexes()
    {
        var image = new DocxInlineImage(24d, 18d, "image/png", [0x89, 0x50, 0x4E, 0x47], "word/media/image1.png");
        DocxParagraph imageParagraph = DocxTests.CreateDocxLayoutParagraph(string.Empty, 10d, 12d) with
        {
            Runs = [],
            Images = [image]
        };
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(imageParagraph)], []);

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None));

        TestAssert.Equal(1, snapshot.Pages.Count);
        TestAssert.Equal(1, snapshot.Pages[0].InlineImageCount);
        TestAssert.Equal(1, snapshot.Pages[0].SourceBlockCount);
        DocxLayoutItemSnapshot imageItem = snapshot.Pages[0].Items.Single(item => item.Kind == "InlineImage");
        TestAssert.Equal(0, imageItem.SourceBlockIndex ?? -1);
        DocxLayoutSourceBlockSnapshot sourceBlock = snapshot.SourceBlocks.Single();
        TestAssert.Equal(0, sourceBlock.SourceBlockIndex);
        TestAssert.Equal("InlineImage", sourceBlock.Kind);
        TestAssert.Equal(1, sourceBlock.InlineImageCount);
        TestAssert.Equal(0, sourceBlock.TextLineCount);
        TestAssert.Equal(0, sourceBlock.TextLength);
        TestAssert.Equal(0, sourceBlock.TableRowCount);
        TestAssert.True(sourceBlock.VerticalTop > sourceBlock.VerticalBottom, "Inline image source block bounds should include the image rectangle.");
    }

    public static void DocxRendererRendersPlacedFootnoteFloatingImages()
    {
        DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body footnote marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "42",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 5)
            ]
        };
        DocxParagraph footnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Footnote anchored image", 10d, 12d);
        var footnoteImage = new DocxInlineImage(18d, 9d, "image/png", TestFixtures.CreateRgbPng(1, 1, [96, 48, 24]), "/word/media/footnote-anchor.png");
        var footnoteDrawing = new DocxFloatingDrawing(
            DistanceTopValue: "0",
            DistanceBottomValue: "0",
            DistanceLeftValue: "0",
            DistanceRightValue: "0",
            SimplePositionValue: "0",
            RelativeHeightValue: "0",
            BehindDocumentValue: "0",
            LockedValue: null,
            LayoutInCellValue: null,
            AllowOverlapValue: null,
            ExtentCxValue: "228600",
            ExtentCyValue: "114300",
            HorizontalRelativeFromValue: "column",
            HorizontalAlignValue: null,
            HorizontalOffsetValue: "457200",
            VerticalRelativeFromValue: "paragraph",
            VerticalAlignValue: null,
            VerticalOffsetValue: "0",
            WrapKind: DocxFloatingWrapKind.None,
            WrapTextValue: null,
            ImageRelationshipId: "rIdFootnoteImage1",
            Image: footnoteImage,
            SourceParagraphIndex: 0,
            SourceBlockIndex: 0);
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "42",
            [new DocxParagraphElement(footnoteParagraph)],
            [],
            [], null)
        {
            FloatingDrawings = [footnoteDrawing]
        };
        DocxDocument document = new(
            200d,
            200d,
            10d,
            10d,
            20d,
            20d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(body)],
            [body],
            [])
        {
            RelatedStories = [footnoteStory]
        };

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxPlacedRelatedStoryLayout placedStory = layout.Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .Single(story => story.StoryLayout.Story.Id == "42");
        DocxFloatingDrawingLayout placedDrawing = placedStory.FloatingDrawings.Single();
        TestAssert.True(placedDrawing.PlacedX >= document.MarginLeftPoints, "Placed footnote floating drawings should be translated from note-local coordinates into the page margin frame.");

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        TestAssert.Equal(1, page.Images.Count);
        TestAssert.Contains("/Im1 Do", page.Content);
    }

    public static void DocxRelatedStoryLayoutOwnsInlineImages()
    {
        DocxTextRun bodyRun = new("Body", 12d, "000000", false, false, false, null, null);
        DocxParagraph bodyParagraph = new(
            [bodyRun],
            [],
            StyleId: null,
            DocxTextAlignment.Left,
            AlignmentValue: null,
            SpacingBeforePoints: 0d,
            SpacingAfterPoints: 0d,
            LineSpacingFactor: 1d,
            LineSpacingPoints: null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            ListLabel: null);
        DocxInlineImage storyImage = new(48d, 24d, "image/png", [1, 2, 3], "/word/media/comment.png");
        DocxParagraph imageParagraph = new(
            [],
            [storyImage],
            StyleId: null,
            DocxTextAlignment.Center,
            AlignmentValue: "center",
            SpacingBeforePoints: 0d,
            SpacingAfterPoints: 0d,
            LineSpacingFactor: 1d,
            LineSpacingPoints: null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            ListLabel: null);
        DocxRelatedStory commentStory = new(
            DocxRelatedStoryKind.Comment,
            "/word/comments.xml",
            "9",
            [new DocxParagraphElement(imageParagraph)],
            [],
            [], null);
        DocxDocument document = new DocxDocument(612d, 792d)
        {
            BodyElements = [new DocxParagraphElement(bodyParagraph)],
            RelatedStories = [commentStory]
        };

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None));
        DocxRelatedStoryLayoutSnapshot storySnapshot = snapshot.RelatedStories.Single();

        TestAssert.True(storySnapshot.Kind == "Comment" && storySnapshot.PartName == "/word/comments.xml" && storySnapshot.Id == "9", "Related-story layout snapshots should preserve the owning story identity.");
        TestAssert.True(storySnapshot.BlockCount == 1 && storySnapshot.ParagraphCount == 1 && storySnapshot.TableCount == 0, "Related-story layout snapshots should derive counts from the body block stream, not stale parallel inventories.");
        TestAssert.True(storySnapshot.TextLineCount == 0 && storySnapshot.InlineImageCount == 1 && storySnapshot.ContentHeight >= 24d, "Related-story inline images should be owned by story layout instead of only contributing anonymous height.");
        DocxLayoutItemSnapshot imageItem = storySnapshot.Items.Single(item => item.Kind == "InlineImage");
        TestAssert.True(imageItem.SourceBlockIndex == 0 && imageItem.SourceParagraphIndex == 0 && imageItem.Width == 48d && imageItem.Height == 24d, "Related-story image item snapshots should carry private-safe source coordinates and geometry.");
        DocxRelatedStorySourceBlockSnapshot imageBlock = storySnapshot.SourceBlocks.Single();
        TestAssert.True(imageBlock.Kind == "InlineImage" && imageBlock.InlineImageCount == 1 && imageBlock.ItemCount == 1 && imageBlock.ConsumedHeight >= 24d, "Related-story source-block snapshots should summarize image-only story blocks without page ownership.");
    }
    public static void DocxReaderReadsFloatingTextBoxContentInsets()
    {
        // Office A/B (w6a1d-tbx reference): Word insets textbox content by the bodyPr
        // values, so the reader must retain them on the drawing model.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                  <w:body>
                    <w:p>
                      <w:r>
                        <w:drawing>
                          <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
                            <wp:simplePos x="0" y="0"/>
                            <wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionH>
                            <wp:positionV relativeFrom="page"><wp:posOffset>1828800</wp:posOffset></wp:positionV>
                            <wp:extent cx="3200400" cy="1097280"/>
                            <wp:wrapNone/>
                            <wp:docPr id="1" name="TextBox 1"/>
                            <a:graphic>
                              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                                <wps:wsp>
                                  <wps:cNvSpPr txBox="1"/>
                                  <wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>
                                  <wps:txbx>
                                    <w:txbxContent>
                                      <w:p><w:r><w:t>Box text</w:t></w:r></w:p>
                                    </w:txbxContent>
                                  </wps:txbx>
                                  <wps:bodyPr lIns="182880" tIns="91440" rIns="182880" bIns="91440"/>
                                </wps:wsp>
                              </a:graphicData>
                            </a:graphic>
                          </wp:anchor>
                        </w:drawing>
                      </w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);
        DocxFloatingDrawing drawing = document.FloatingDrawings.Single();
        TestAssert.Equal("182880", drawing.TextBoxInsetLeftValue ?? "?");
        TestAssert.Equal("91440", drawing.TextBoxInsetTopValue ?? "?");
        TestAssert.Equal("182880", drawing.TextBoxInsetRightValue ?? "?");
        TestAssert.Equal("91440", drawing.TextBoxInsetBottomValue ?? "?");
    }

    public static void DocxReaderLeavesMissingTextBoxContentInsetsNull()
    {
        // The text-box fixture carries no bodyPr insets; the model stays null so layout
        // applies the Office defaults.
        string input = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "Cases",
            "docx-markup-comment-text-box.docx"));
        DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);
        DocxFloatingDrawing drawing = document.FloatingDrawings.Single();
        TestAssert.True(
            drawing.TextBoxInsetLeftValue is null &&
            drawing.TextBoxInsetTopValue is null &&
            drawing.TextBoxInsetRightValue is null &&
            drawing.TextBoxInsetBottomValue is null,
            "Missing bodyPr insets should stay null on the drawing model.");
    }

    private static DocxDocument WriteFloatingTextBoxInsetLayoutDocument(string text)
    {
        DocxParagraph bodyParagraph = DocxTests.CreateDocxLayoutParagraph("Anchor", 10d, 12d);
        DocxParagraph textBoxParagraph = DocxTests.CreateDocxLayoutParagraph(text, 10d, 12d);
        DocxFloatingDrawing floatingDrawing = DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(textBoxParagraph)]);
        return new DocxDocument(
            612d,
            792d,
            72d,
            72d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [floatingDrawing],
            [],
            [],
            [new DocxParagraphElement(bodyParagraph)],
            [],
            [])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };
    }

    public static void DocxWordCompatibleFloatingTextBoxWrapsInsideInsetContentWidth()
    {
        // Office defaults inset the content box (0.1in sides); the crafted text fits the
        // full extent but wraps inside the inset width (END-to-END against pre-insets layout).
        // Textbox content scales with the page like body text (tbxrev probe refutes the
        // unscaled-content reading of the s=1 fixture).
        DocxDocument document = WriteFloatingTextBoxInsetLayoutDocument(new string('X', 21) + " " + new string('Y', 2));
        var scaled = new DocxTests.ScaledLayoutTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), 0.842391d, 0.7936d);
        DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .Create(document, scaled, CancellationToken.None)
            .FloatingDrawings.Single().TextBoxLayout?.TextLines.ToArray() ?? [];

        TestAssert.Equal(2, lines.Length);
        TestAssert.True(lines[0].Text.StartsWith(new string('X', 21), StringComparison.Ordinal), "First inset line should hold the leading run. Text=[" + lines[0].Text + "]");
        TestAssert.True(lines[1].Text.EndsWith("Y", StringComparison.Ordinal), "Inset overflow should wrap to a second line. Text=[" + lines[1].Text + "]");
    }

    public static void DocxWordCompatibleFloatingTextBoxMeasuresContentInDesignSpace()
    {
        // Office A/B (tbxrev probe, Word-COM rendered): scaled pages lay floating-textbox
        // content out in design space (12pt advances in the unscaled content box) and map it
        // uniformly to emission space, so the layout must measure it raw, not scaled.
        DocxDocument document = WriteFloatingTextBoxInsetLayoutDocument(new string('X', 18));
        var scaled = new DocxTests.ScaledLayoutTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), 0.842391d, 0.7936d);
        DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .Create(document, scaled, CancellationToken.None, new DocxTests.FamilyWidthTextMeasurer())
            .FloatingDrawings.Single().TextBoxLayout?.TextLines.ToArray() ?? [];

        TestAssert.Equal(1, lines.Length);
        TestAssert.True(Math.Abs(lines[0].Width - 90d) < 0.001d, "Textbox content should measure raw design advances (18 chars at 5pt), not scaled. Width=" + lines[0].Width.ToString(CultureInfo.InvariantCulture));
    }

    public static void DocxWordCompatibleFloatingTextBoxKeepsScaledMeasurementAtUnitScale()
    {
        // The design-space fork only engages on scaled pages; unit-scale geometry keeps the
        // legacy single-measurer path even when a raw measurer is threaded alongside.
        DocxDocument document = WriteFloatingTextBoxInsetLayoutDocument(new string('X', 18));
        var scaled = new DocxTests.ScaledLayoutTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), 0.842391d, 0.7936d);
        DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, 1d)
            .Create(document, scaled, CancellationToken.None, new DocxTests.FamilyWidthTextMeasurer())
            .FloatingDrawings.Single().TextBoxLayout?.TextLines.ToArray() ?? [];

        TestAssert.Equal(1, lines.Length);
        TestAssert.True(Math.Abs(lines[0].Width - 90d * 0.842391d) < 0.001d, "Unit-scale pages should keep scaled measurement. Width=" + lines[0].Width.ToString(CultureInfo.InvariantCulture));
    }

    public static void DocxWordCompatibleFloatingTextBoxEmissionMapMatchesOfficeProbe()
    {
        // Office A/B (tbxrev probe, Word-COM rendered at s=0.758834): design content origin
        // X 79.2 lands at 60.1 (origin scale) and design baseline 633.12 at 576.1
        // (page-center scale), while the body first-pin shift strands them at 61.8/557.7.
        DocxMarkupContext context = DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup) with
        {
            WordCompatiblePrintScale = 0.758834d,
            WordCompatibleTextXOffset = -17.3639d,
            WordCompatibleTextYOffset = 75.42d
        };
        TestAssert.True(DocxRenderer.FloatingTextBoxEmissionMap.TryCreate(context, 792d, out DocxRenderer.FloatingTextBoxEmissionMap map), "Scaled WC pages should engage the floating-textbox emission map.");
        double emissionX = map.MapEmissionX(79.2d);
        double emissionY = map.MapEmissionY(633.12d);
        TestAssert.True(Math.Abs(emissionX - 60.1d) < 0.05d, "Mapped content X should match Word 60.1. X=" + emissionX.ToString(CultureInfo.InvariantCulture));
        TestAssert.True(Math.Abs(emissionY - 576.1d) < 0.3d, "Mapped baseline Y should match Word 576.1. Y=" + emissionY.ToString(CultureInfo.InvariantCulture));
        TestAssert.True(Math.Abs(map.PrecompensateX(79.2d) + context.WordCompatibleTextXOffset - emissionX) < 0.000000001d, "Pre-compensated X must round-trip through the body X offset.");
        TestAssert.True(Math.Abs(map.PrecompensateY(633.12d) - context.WordCompatibleTextYOffset - emissionY) < 0.000000001d, "Pre-compensated Y must round-trip through the body Y offset.");
        DocxMarkupContext unitScale = context with { WordCompatiblePrintScale = 1d, WordCompatibleTextXOffset = 0d, WordCompatibleTextYOffset = 0d };
        TestAssert.True(!DocxRenderer.FloatingTextBoxEmissionMap.TryCreate(unitScale, 792d, out _), "Unit-scale pages should keep the legacy path.");
        DocxMarkupContext preserve = DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout) with { WordCompatiblePrintScale = 0.758834d };
        TestAssert.True(!DocxRenderer.FloatingTextBoxEmissionMap.TryCreate(preserve, 792d, out _), "Preserve geometry should keep the legacy path.");
    }

    public static void DocxWordCompatibleFloatingTextBoxPrecompensatedEmissionMatchesCenterScale()
    {
        // Layout (design space) plus the emission map must reproduce the uniform page-center
        // scale independently of the body first-pin shift. The box sits ~400pt below the
        // body pin so the test reproduces the reported divergence class (pin strands text
        // low, under its own clip, growing with distance from the first baseline).
        const double printScale = 0.758834d;
        DocxParagraph bodyParagraph = DocxTests.CreateDocxLayoutParagraph("Anchor", 10d, 12d);
        DocxParagraph textBoxParagraph = DocxTests.CreateDocxLayoutParagraph(new string('X', 18), 10d, 12d);
        DocxFloatingDrawing lowDrawing = DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(textBoxParagraph)]) with
        {
            VerticalOffsetValue = "6249600"
        };
        DocxDocument document = new(
            612d,
            792d,
            72d,
            72d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [lowDrawing],
            [],
            [],
            [new DocxParagraphElement(bodyParagraph)],
            [],
            [])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };
        var scaled = new DocxTests.ScaledLayoutTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), printScale, 0.7936d);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, printScale)
            .Create(document, scaled, CancellationToken.None, new DocxTests.FamilyWidthTextMeasurer());
        DocxFloatingDrawingLayout drawing = layout.FloatingDrawings.Single();
        double placedX = drawing.PlacedX ?? throw new InvalidOperationException("Probe drawing should place horizontally.");
        double placedTop = drawing.PlacedTop ?? throw new InvalidOperationException("Probe drawing should place vertically.");
        DocxTextLineLayout line = drawing.TextBoxLayout?.TextLines.Single() ?? throw new InvalidOperationException("Probe drawing should lay out one textbox line.");
        double pageHeight = layout.Pages[0].Height;
        double firstBaseline = layout.Pages[0].Items.OfType<DocxTextLineLayout>().Max(item => item.BaselineY);
        double xOffset = -document.MarginLeftPoints * (1d - printScale);
        double yOffset = (firstBaseline - pageHeight / 2d) * (1d - printScale);
        DocxMarkupContext context = DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup) with
        {
            WordCompatiblePrintScale = printScale,
            WordCompatibleTextXOffset = xOffset,
            WordCompatibleTextYOffset = yOffset
        };
        TestAssert.True(DocxRenderer.FloatingTextBoxEmissionMap.TryCreate(context, pageHeight, out DocxRenderer.FloatingTextBoxEmissionMap map), "Scaled WC pages should engage the floating-textbox emission map.");
        DocxLayoutEngine.ResolveTextBoxContentInsets(drawing.Drawing, out double insetLeft, out double insetTop, out _, out _);
        double designX = placedX + insetLeft + line.X;
        double designY = placedTop - insetTop + line.BaselineY;
        double expectedX = designX * printScale;
        double expectedY = designY * printScale + (pageHeight / 2d) * (1d - printScale);
        DocxTextLineLayout mapped = map.PrecompensateLine(line, placedX + insetLeft, placedTop - insetTop);
        TestAssert.True(Math.Abs(mapped.X + xOffset - expectedX) < 0.000000001d, "Pre-compensated X must land center-scaled through the body offset.");
        TestAssert.True(Math.Abs(mapped.BaselineY - yOffset - expectedY) < 0.000000001d, "Pre-compensated Y must land center-scaled through the body offset.");
        TestAssert.True(expectedY - (designY - yOffset) > 10d, "Center-scale must sit well above the first-pin prediction that stranded textbox text under its clip.");
    }

    public static void DocxReaderReadsInlineTextBoxContent()
    {
        // Inline DrawingML textboxes must parse like floating ones (extent, bodyPr insets,
        // txbxContent elements); today the reader drops them (no blip, no image).
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                  <w:body>
                    <w:p>
                      <w:r>
                        <w:drawing>
                          <wp:inline>
                            <wp:extent cx="3200400" cy="1097280"/>
                            <wp:docPr id="1" name="TextBox 1"/>
                            <a:graphic>
                              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                                <wps:wsp>
                                  <wps:cNvSpPr txBox="1"/>
                                  <wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>
                                  <wps:txbx>
                                    <w:txbxContent>
                                      <w:p><w:r><w:t>Box text</w:t></w:r></w:p>
                                    </w:txbxContent>
                                  </wps:txbx>
                                  <wps:bodyPr lIns="182880" tIns="91440" rIns="182880" bIns="91440"/>
                                </wps:wsp>
                              </a:graphicData>
                            </a:graphic>
                          </wp:inline>
                        </w:drawing>
                      </w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);
        DocxParagraph paragraph = document.Paragraphs.Single();
        TestAssert.True(paragraph.Images.Count == 0, "An inline textbox carries no blip and must not parse as a picture.");
        DocxInlineTextBox textBox = paragraph.InlineTextBoxes.Single();
        TestAssert.Equal("3200400", textBox.ExtentCxValue ?? "?");
        TestAssert.Equal("1097280", textBox.ExtentCyValue ?? "?");
        TestAssert.Equal("182880", textBox.TextBoxInsetLeftValue ?? "?");
        TestAssert.Equal("91440", textBox.TextBoxInsetTopValue ?? "?");
        TestAssert.Equal("182880", textBox.TextBoxInsetRightValue ?? "?");
        TestAssert.Equal("91440", textBox.TextBoxInsetBottomValue ?? "?");
        DocxParagraph boxParagraph = ((DocxParagraphElement)textBox.BodyElements.Single()).Paragraph;
        TestAssert.Equal("Box text", string.Concat(boxParagraph.Runs.Select(run => run.Text)));
    }

    public static void DocxWordCompatibleInlineTextBoxLaysOutScaledContentInFlow()
    {
        // Office A/B (w6-inline probe, Word-COM rendered): inline boxes join the scaled
        // body flow uniformly (box 191.1 equals 252 times s, content 9.1pt), so the layout
        // measures content scaled in the scaled content box and places it in flow.
        const double printScale = 0.758834d;
        DocxParagraph textBoxParagraph = DocxTests.CreateDocxLayoutParagraph(new string('X', 18), 10d, 12d);
        var textBox = new DocxInlineTextBox("1371600", "457200", "91440", "45720", "91440", "45720")
        {
            BodyElements = [new DocxParagraphElement(textBoxParagraph)]
        };
        DocxParagraph hostParagraph = DocxTests.CreateDocxLayoutParagraph("Anchor", 10d, 12d) with
        {
            InlineTextBoxes = [textBox]
        };
        DocxDocument document = new(
            612d,
            792d,
            72d,
            72d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(hostParagraph)],
            [],
            [])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };
        var scaled = new DocxTests.ScaledLayoutTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), printScale, 0.7936d);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, printScale)
            .Create(document, scaled, CancellationToken.None);
        DocxInlineTextBoxLayout box = layout.Pages[0].Items.OfType<DocxInlineTextBoxLayout>().Single();
        TestAssert.True(Math.Abs(box.BoxWidth - 108d * printScale) < 0.001d, "Inline box width should scale uniformly with the page. Width=" + box.BoxWidth.ToString(CultureInfo.InvariantCulture));
        DocxTextLineLayout line = box.TextLines.Single();
        TestAssert.True(Math.Abs(line.Width - 90d * printScale) < 0.001d, "Inline box content should measure scaled like body text. Width=" + line.Width.ToString(CultureInfo.InvariantCulture));
        TestAssert.True(Math.Abs(line.X - (box.BoxX + 7.2d * printScale)) < 0.001d, "Inline box content should start at the scaled inset. X=" + line.X.ToString(CultureInfo.InvariantCulture));
    }

    public static void DocxWordCompatibleTableCellInlineTextBoxLaysOutScaledContent()
    {
        // Office A/B (w6-tablebox probe, Word-COM rendered): cell boxes join the scaled
        // flow uniformly (content 9.1pt, box revision ballooned, box comment suppressed),
        // so cell layout measures box content scaled in the scaled content box.
        const double printScale = 0.758834d;
        DocxParagraph textBoxParagraph = DocxTests.CreateDocxLayoutParagraph(new string('X', 18), 10d, 12d);
        var textBox = new DocxInlineTextBox("1371600", "457200", "91440", "45720", "91440", "45720")
        {
            BodyElements = [new DocxParagraphElement(textBoxParagraph)]
        };
        DocxParagraph hostParagraph = DocxTests.CreateDocxLayoutParagraph("Cell ", 10d, 12d) with
        {
            InlineTextBoxes = [textBox]
        };
        var cell = new DocxTableCell(string.Empty, [hostParagraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [200d], [new DocxTableRow([cell], null)]);
        DocxDocument document = new(
            612d,
            792d,
            72d,
            72d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxTableElement(table)],
            [],
            [table])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };
        var scaled = new DocxTests.ScaledLayoutTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), printScale, 0.7936d);
        DocxTableCellLayout cellLayout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, printScale)
            .Create(document, scaled, CancellationToken.None)
            .Pages.Single()
            .Items.OfType<DocxTableRowLayout>()
            .Single().Cells.Single();
        DocxInlineTextBoxLayout box = cellLayout.InlineTextBoxes.Single();
        DocxTextLineLayout line = box.TextLines.Single();
        TestAssert.True(Math.Abs(line.Width - 90d * printScale) < 0.001d, "Cell box content should measure scaled like body text. Width=" + line.Width.ToString(CultureInfo.InvariantCulture));
        TestAssert.True(Math.Abs(line.X - (box.BoxX + 7.2d * printScale)) < 0.001d, "Cell box content should start at the scaled inset. X=" + line.X.ToString(CultureInfo.InvariantCulture));
        TestAssert.True(box.BoxHeight + 0.001d >= 36d * printScale, "Cell box height should cover the scaled extent. Height=" + box.BoxHeight.ToString(CultureInfo.InvariantCulture));
        DocxTextLineLayout firstCellLine = cellLayout.TextLines.First();
        TestAssert.True(Math.Abs((firstCellLine.BaselineY - box.BoxTop) - 12d) < 0.001d, "Cell box block should start exactly one line below the host text line (block-level, like pictures).");
    }
    public static void DocxCellTextBoxAfterTextTucksToBaseline()
    {
        // Office A/B (w6-celltucksize probes, Word-COM rendered at two lane-fit scales):
        // a box-only cell paragraph after text tucks its top below the last text baseline
        // following a font-size affine law, instead of stacking a full line plus spacing.
        const double printScale = 0.758834d;
        const double fontSize = 12d;
        DocxParagraph textParagraph = DocxTests.CreateDocxLayoutParagraph("Cell text", fontSize, fontSize);
        var textBox = new DocxInlineTextBox("2286000", "1097280")
        {
            BodyElements = [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Box text", 10d, 12d))]
        };
        DocxParagraph boxParagraph = new DocxParagraph([], [], null, DocxTextAlignment.Left, null, 0d, 0d, 1d, null, DocxParagraphSpacing.Empty, DocxParagraphKeepRules.Empty, null) with
        {
            InlineTextBoxes = [textBox]
        };
        var cell = new DocxTableCell(string.Empty, [textParagraph, boxParagraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [200d], [new DocxTableRow([cell], null)]);
        DocxDocument document = new(
            612d,
            792d,
            72d,
            72d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxTableElement(table)],
            [],
            [table])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };
        var scaled = new DocxTests.ScaledLayoutTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), printScale, 0.7936d);
        DocxTableCellLayout cellLayout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, printScale)
            .Create(document, scaled, CancellationToken.None)
            .Pages.Single()
            .Items.OfType<DocxTableRowLayout>()
            .Single().Cells.Single();
        DocxTextLineLayout textLine = cellLayout.TextLines.Single();
        DocxInlineTextBoxLayout box = cellLayout.InlineTextBoxes.Single();
        double expectedTop = textLine.BaselineY - printScale * (0.470d * fontSize + 3.65d);
        TestAssert.True(Math.Abs(box.BoxTop - expectedTop) < 0.05d, "Cell box tops should tuck below the text baseline by the Office font-size law. BoxTop=" + box.BoxTop.ToString(CultureInfo.InvariantCulture) + " Expected=" + expectedTop.ToString(CultureInfo.InvariantCulture));
    }

    public static void DocxCellTextBoxAfterTextTuckGrowsWithFontSize()
    {
        // Companion slope probe (w6-celltucksize 16pt row): the tuck grows with the
        // preceding text size, so larger text carries a deeper box top.
        const double printScale = 0.758834d;
        const double fontSize = 16d;
        DocxParagraph textParagraph = DocxTests.CreateDocxLayoutParagraph("Cell text", fontSize, fontSize);
        var textBox = new DocxInlineTextBox("2286000", "1097280")
        {
            BodyElements = [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Box text", 10d, 12d))]
        };
        DocxParagraph boxParagraph = new DocxParagraph([], [], null, DocxTextAlignment.Left, null, 0d, 0d, 1d, null, DocxParagraphSpacing.Empty, DocxParagraphKeepRules.Empty, null) with
        {
            InlineTextBoxes = [textBox]
        };
        var cell = new DocxTableCell(string.Empty, [textParagraph, boxParagraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [200d], [new DocxTableRow([cell], null)]);
        DocxDocument document = new(
            612d,
            792d,
            72d,
            72d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxTableElement(table)],
            [],
            [table])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };
        var scaled = new DocxTests.ScaledLayoutTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), printScale, 0.7936d);
        DocxTableCellLayout cellLayout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, printScale)
            .Create(document, scaled, CancellationToken.None)
            .Pages.Single()
            .Items.OfType<DocxTableRowLayout>()
            .Single().Cells.Single();
        DocxTextLineLayout textLine = cellLayout.TextLines.Single();
        DocxInlineTextBoxLayout box = cellLayout.InlineTextBoxes.Single();
        double expectedTop = textLine.BaselineY - printScale * (0.470d * fontSize + 3.65d);
        TestAssert.True(Math.Abs(box.BoxTop - expectedTop) < 0.05d, "Cell box tucks should grow with the preceding text size. BoxTop=" + box.BoxTop.ToString(CultureInfo.InvariantCulture) + " Expected=" + expectedTop.ToString(CultureInfo.InvariantCulture));
    }

    public static void DocxCellTextBoxAfterTextHonorsSpacingExcess()
    {
        // Office A/B (w6-celltuckspacing probes, Word-COM rendered): explicit spacing
        // between text and a following box honors only the excess over the 8pt default,
        // so after-8 matches the default tuck while after-24 adds 16pt.
        const double printScale = 0.758834d;
        const double fontSize = 12d;
        DocxParagraph textParagraph = DocxTests.CreateDocxLayoutParagraph("Cell text", fontSize, fontSize) with
        {
            SpacingAfterPoints = 24d
        };
        var textBox = new DocxInlineTextBox("2286000", "1097280")
        {
            BodyElements = [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Box text", 10d, 12d))]
        };
        DocxParagraph boxParagraph = new DocxParagraph([], [], null, DocxTextAlignment.Left, null, 0d, 0d, 1d, null, DocxParagraphSpacing.Empty, DocxParagraphKeepRules.Empty, null) with
        {
            InlineTextBoxes = [textBox]
        };
        var cell = new DocxTableCell(string.Empty, [textParagraph, boxParagraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [200d], [new DocxTableRow([cell], null)]);
        DocxDocument document = new(
            612d,
            792d,
            72d,
            72d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxTableElement(table)],
            [],
            [table])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };
        var scaled = new DocxTests.ScaledLayoutTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), printScale, 0.7936d);
        DocxTableCellLayout cellLayout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, printScale)
            .Create(document, scaled, CancellationToken.None)
            .Pages.Single()
            .Items.OfType<DocxTableRowLayout>()
            .Single().Cells.Single();
        DocxTextLineLayout textLine = cellLayout.TextLines.Single();
        DocxInlineTextBoxLayout box = cellLayout.InlineTextBoxes.Single();
        double expectedTop = textLine.BaselineY - printScale * (0.470d * fontSize + 3.65d) - (24d - 8d) * printScale;
        TestAssert.True(Math.Abs(box.BoxTop - expectedTop) < 0.05d, "Cell boxes should honor spacing above the 8pt default on top of the tuck. BoxTop=" + box.BoxTop.ToString(CultureInfo.InvariantCulture) + " Expected=" + expectedTop.ToString(CultureInfo.InvariantCulture));
    }

    public static void DocxStaticHeaderInlineTextBoxLaysOutFileGeometryContent()
    {
        // Static stories lay out inline boxes like body flow (scaled extents and insets
        // with the ambient measurer); at unit scale the file geometry passes through.
        DocxParagraph textBoxParagraph = DocxTests.CreateDocxLayoutParagraph(new string('X', 18), 10d, 12d);
        var textBox = new DocxInlineTextBox("1371600", "457200", "91440", "45720", "91440", "45720")
        {
            BodyElements = [new DocxParagraphElement(textBoxParagraph)]
        };
        DocxParagraph hostParagraph = DocxTests.CreateDocxLayoutParagraph("Header host", 10d, 12d) with
        {
            InlineTextBoxes = [textBox]
        };
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [new DocxParagraphElement(hostParagraph)]
            }
        };
        DocxDocument document = new(
            612d,
            792d,
            72d,
            72d,
            72d,
            72d,
            settings,
            [],
            [],
            [],
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Body", 10d, 12d))],
            [],
            []);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        DocxInlineTextBoxLayout box = layout.Pages[0].StaticInlineTextBoxes.Single();
        TestAssert.True(Math.Abs(box.BoxWidth - 108d) < 0.001d, "Static box width should keep file geometry at unit scale. Width=" + box.BoxWidth.ToString(CultureInfo.InvariantCulture));
        DocxTextLineLayout line = box.TextLines.Single();
        TestAssert.True(Math.Abs(line.Width - 90d) < 0.001d, "Static box content should measure raw at unit scale. Width=" + line.Width.ToString(CultureInfo.InvariantCulture));
        TestAssert.True(Math.Abs(line.X - (box.BoxX + 7.2d)) < 0.001d, "Static box content should start at the file inset. X=" + line.X.ToString(CultureInfo.InvariantCulture));
    }

    public static void FloatingDrawingPageIndexMatchesLegacyFiltering()
    {
        // R12: the once-per-render page index returns exactly the legacy per-page
        // sequences: document order for All, stable z-order per layer, null anchors
        // excluded, missing pages empty. Markers keep value-equal records distinct.
        DocxFloatingDrawingLayout a = MakePageDrawing("A", 0, behind: false, z: "2");
        DocxFloatingDrawingLayout b = MakePageDrawing("B", 0, behind: true, z: "5");
        DocxFloatingDrawingLayout c = MakePageDrawing("C", 0, behind: true, z: "5");
        DocxFloatingDrawingLayout d = MakePageDrawing("D", 0, behind: false, z: "1");
        DocxFloatingDrawingLayout e = MakePageDrawing("E", 1, behind: true, z: "9");
        DocxFloatingDrawingLayout f = MakePageDrawing("F", null, behind: false, z: "0");
        var floating = new List<DocxFloatingDrawingLayout> { a, b, c, d, e, f };
        DocxFloatingDrawingLayout g = MakePageDrawing("G", 0, behind: false, z: "7");
        DocxFloatingDrawingLayout h = MakePageDrawing("H", 2, behind: true, z: "3");
        var staticDrawings = new List<DocxFloatingDrawingLayout> { g, h };

        DocxRenderer.FloatingDrawingPageIndex floatingIndex = DocxRenderer.FloatingDrawingPageIndex.Build(floating, CancellationToken.None);
        DocxRenderer.FloatingDrawingPageIndex staticIndex = DocxRenderer.FloatingDrawingPageIndex.Build(staticDrawings, CancellationToken.None);
        var pair = new DocxRenderer.FloatingDrawingPageIndex.PageIndexPair(floatingIndex, staticIndex);

        TestAssert.True(SequenceEqual([a, b, c, d], floatingIndex.Get(0).All), "Page 0 All must keep document order.");
        TestAssert.True(SequenceEqual([b, c], floatingIndex.Get(0).Behind), "Page 0 Behind must sort stably by z-order.");
        TestAssert.True(SequenceEqual([d, a], floatingIndex.Get(0).Ahead), "Page 0 Ahead must sort stably by z-order.");
        TestAssert.True(SequenceEqual([e], floatingIndex.Get(1).All), "Page 1 All must hold its drawing.");
        TestAssert.True(SequenceEqual([e], floatingIndex.Get(1).Behind), "Page 1 Behind must hold its drawing.");
        TestAssert.True(SequenceEqual([], floatingIndex.Get(1).Ahead), "Page 1 Ahead must be empty.");
        TestAssert.True(SequenceEqual([], floatingIndex.Get(99).All), "Missing pages must yield empty sets.");
        TestAssert.True(SequenceEqual([], floatingIndex.Get(99).Behind), "Missing pages must yield empty sets.");
        TestAssert.True(SequenceEqual([g], staticIndex.Get(0).All), "Static page 0 must hold its drawing.");
        TestAssert.True(SequenceEqual([h], staticIndex.Get(2).Behind), "Static page 2 must hold its drawing.");
        TestAssert.True(SequenceEqual([a, b, c, d, g], pair.PageAll(0)), "PageAll must concatenate floating then static in document order.");
        TestAssert.True(SequenceEqual([e], pair.PageAll(1)), "PageAll must hold single-list pages.");
        TestAssert.True(SequenceEqual([h], pair.PageAll(2)), "PageAll must hold static-only pages.");
        TestAssert.True(SequenceEqual([], pair.PageAll(99)), "PageAll must be empty for missing pages.");
    }

    private static DocxFloatingDrawingLayout MakePageDrawing(string marker, int? page, bool behind, string z)
    {
        var drawing = new DocxFloatingDrawing(
            DistanceTopValue: null,
            DistanceBottomValue: null,
            DistanceLeftValue: null,
            DistanceRightValue: null,
            SimplePositionValue: null,
            RelativeHeightValue: z,
            BehindDocumentValue: behind ? "1" : null,
            LockedValue: null,
            LayoutInCellValue: null,
            AllowOverlapValue: null,
            ExtentCxValue: null,
            ExtentCyValue: null,
            HorizontalRelativeFromValue: null,
            HorizontalAlignValue: marker,
            HorizontalOffsetValue: null,
            VerticalRelativeFromValue: null,
            VerticalAlignValue: null,
            VerticalOffsetValue: null,
            WrapKind: null,
            WrapTextValue: null,
            ImageRelationshipId: null,
            Image: null,
            SourceParagraphIndex: null,
            SourceBlockIndex: null);
        return new DocxFloatingDrawingLayout(
            Drawing: drawing,
            PageStartIndex: null,
            PageEndIndex: null,
            AnchorPageIndex: page,
            AnchorColumnIndex: null,
            AnchorBlockVerticalTop: null,
            AnchorBlockVerticalBottom: null,
            ExtentWidthPoints: null,
            ExtentHeightPoints: null,
            HorizontalOffsetPoints: null,
            VerticalOffsetPoints: null,
            DistanceTopPoints: null,
            DistanceBottomPoints: null,
            DistanceLeftPoints: null,
            DistanceRightPoints: null,
            HorizontalReferenceX: null,
            HorizontalReferenceWidth: null,
            VerticalReferenceTop: null,
            VerticalReferenceBottom: null,
            PlacedX: null,
            PlacedTop: null,
            HorizontalPlacementSource: null,
            VerticalPlacementSource: null,
            WrapExclusionX: null,
            WrapExclusionTop: null,
            WrapExclusionWidth: null,
            WrapExclusionHeight: null,
            Story: null,
            TextBoxLayout: null);
    }

    private static bool SequenceEqual(IReadOnlyList<DocxFloatingDrawingLayout> expected, IReadOnlyList<DocxFloatingDrawingLayout> actual)
    {
        if (expected.Count != actual.Count)
        {
            return false;
        }

        for (int i = 0; i < expected.Count; i++)
        {
            if (!ReferenceEquals(expected[i], actual[i]))
            {
                return false;
            }
        }

        return true;
    }

    // RV05: inline images must carry their source run affinity so layout can place
    // them at run position instead of after paragraph text.
    public static void DocxInlineImageCarriesSourceRunAffinity()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Default Extension="xml" ContentType="application/xml"/>
                <Default Extension="png" ContentType="image/png"/>
                <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """),
            ["word/_rels/document.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
                </Relationships>
                """),
            ["word/document.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p>
                      <w:r><w:t>BEFORE</w:t></w:r>
                      <w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdImage1"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>
                      <w:r><w:t>AFTER</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """),
            ["word/media/image1.png"] = TestFixtures.CreateRgbPng(2, 1, [255, 0, 0, 0, 0, 255])
        });
    
        DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
    
        DocxParagraph paragraph = document.Paragraphs.Single();
        TestAssert.Equal("BEFOREAFTER", string.Concat(paragraph.Runs.Select(run => run.Text)));
        TestAssert.Equal(1, paragraph.Images.Count);
        TestAssert.Equal(1, paragraph.Images.Single().SourceRunIndex);
    }
}
