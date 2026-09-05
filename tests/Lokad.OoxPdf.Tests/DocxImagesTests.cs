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
}
