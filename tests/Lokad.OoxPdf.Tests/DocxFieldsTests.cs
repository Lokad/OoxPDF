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

internal static class DocxFieldsTests
{
    public static void DocxReaderPreservesFieldReferencesStructurally()
    {
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
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:fldSimple w:instr=" PAGE "/>
                      <w:fldSimple w:instr=" NUMPAGES "/>
                      <w:fldSimple w:instr=" PAGEREF Target "><w:r><w:t>TargetValue</w:t></w:r></w:fldSimple>
                      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                      <w:r><w:instrText> PAGE </w:instrText></w:r>
                      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                      <w:r><w:t>1</w:t></w:r>
                      <w:r><w:fldChar w:fldCharType="end"/></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxParagraph paragraph = document.Paragraphs.Single();

        TestAssert.Equal(4, paragraph.FieldReferences.Count);
        TestAssert.Equal(2, paragraph.FieldReferences.Count(field => field.Kind == DocxFieldKind.Page));
        TestAssert.Equal(1, paragraph.FieldReferences.Count(field => field.Kind == DocxFieldKind.NumPages));
        TestAssert.Equal(1, paragraph.FieldReferences.Count(field => field.Kind == DocxFieldKind.Other));
        TestAssert.Equal(3, paragraph.FieldReferences.Count(field => field.SourceKind == DocxFieldSourceKind.Simple));
        TestAssert.Equal(1, paragraph.FieldReferences.Count(field => field.SourceKind == DocxFieldSourceKind.ComplexInstruction));
        TestAssert.Equal(3, paragraph.FieldReferences.Count(field => field.UsesPlaceholder));
        TestAssert.Equal(2, paragraph.FieldReferences.Count(field => field.HasCachedResult));
        TestAssert.Equal(1, paragraph.FieldReferences.Count(field => field.HasSeparate));
        TestAssert.Equal(2, paragraph.FieldReferences.Count(field => field.Placeholder == "{PAGE}"));
        TestAssert.Equal(1, paragraph.FieldReferences.Count(field => field.Placeholder == "{NUMPAGES}"));
        TestAssert.True(paragraph.FieldReferences.Single(field => field.Instruction?.Contains("PAGEREF", StringComparison.Ordinal) == true).Placeholder is null, "PAGEREF should not be treated as a PAGE placeholder.");
        DocxFieldReference simplePage = paragraph.FieldReferences.First(field => field.Kind == DocxFieldKind.Page && field.SourceKind == DocxFieldSourceKind.Simple);
        DocxFieldReference simpleNumPages = paragraph.FieldReferences.Single(field => field.Kind == DocxFieldKind.NumPages);
        DocxFieldReference pageRef = paragraph.FieldReferences.Single(field => field.Kind == DocxFieldKind.Other);
        DocxFieldReference complexPage = paragraph.FieldReferences.Single(field => field.Kind == DocxFieldKind.Page && field.SourceKind == DocxFieldSourceKind.ComplexInstruction);
        TestAssert.Equal(0, simplePage.TextRunIndex);
        TestAssert.Equal(1, simplePage.TextRunCount);
        TestAssert.Equal(6, simplePage.TextLength);
        TestAssert.Equal(1, simpleNumPages.TextRunIndex);
        TestAssert.Equal(1, simpleNumPages.TextRunCount);
        TestAssert.Equal(10, simpleNumPages.TextLength);
        TestAssert.Equal(2, pageRef.TextRunIndex);
        TestAssert.Equal(1, pageRef.TextRunCount);
        TestAssert.Equal(11, pageRef.TextLength);
        TestAssert.Equal(3, complexPage.TextRunIndex);
        TestAssert.Equal(1, complexPage.TextRunCount);
        TestAssert.Equal(6, complexPage.TextLength);
        TestAssert.True(complexPage.UsesPlaceholder, "Complex PAGE fields should render the dynamic page placeholder.");
        TestAssert.True(complexPage.HasCachedResult, "Complex PAGE fields should record that a cached source result was present.");
        TestAssert.True(!complexPage.RendersCachedResult, "Complex PAGE fields should not report that the cached page number was rendered.");
        TestAssert.Equal(1, complexPage.InstructionRunCount);
        TestAssert.Equal(1, complexPage.ResultRunCount);
        TestAssert.Equal("{PAGE}{NUMPAGES}TargetValue{PAGE}", string.Concat(paragraph.Runs.Select(run => run.Text)));

        DocxStructureSnapshot snapshot = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectStructure(document);
        DocxStructureBlockSnapshot block = snapshot.Blocks.Single(block => block.Kind == "Paragraph");
        DocxStructureStorySnapshot bodyStory = snapshot.Stories.Single(story => story.Kind == "Body");
        TestAssert.Equal(4, snapshot.FieldReferenceCount);
        TestAssert.Equal(2, snapshot.PageFieldReferenceCount);
        TestAssert.Equal(1, snapshot.NumPagesFieldReferenceCount);
        TestAssert.Equal(1, snapshot.OtherFieldReferenceCount);
        TestAssert.Equal(1, snapshot.ComplexFieldReferenceCount);
        TestAssert.Equal(2, snapshot.CachedResultFieldReferenceCount);
        TestAssert.Equal(1, snapshot.RenderedCachedResultFieldReferenceCount);
        TestAssert.Equal(3, snapshot.PlaceholderFieldReferenceCount);
        TestAssert.Equal(3, snapshot.DynamicFieldReferenceCount);
        TestAssert.Equal(3, snapshot.DynamicPlaceholderFieldReferenceCount);
        TestAssert.Equal(0, snapshot.DynamicComplexWithoutCachedResultFieldReferenceCount);
        TestAssert.Equal(1, snapshot.DynamicCachedResultNotRenderedFieldReferenceCount);
        TestAssert.Equal(0, snapshot.NestedFieldReferenceCount);
        TestAssert.Equal(4, block.FieldReferenceCount);
        TestAssert.Equal(2, block.PageFieldReferenceCount);
        TestAssert.Equal(1, block.NumPagesFieldReferenceCount);
        TestAssert.Equal(1, block.OtherFieldReferenceCount);
        TestAssert.Equal(1, block.ComplexFieldReferenceCount);
        TestAssert.Equal(2, block.CachedResultFieldReferenceCount);
        TestAssert.Equal(4, bodyStory.FieldReferenceCount);
        TestAssert.Equal(1, bodyStory.ComplexFieldReferenceCount);
    }

    public static void DocxReaderPreservesDynamicFieldReferencesInsideNoteStories()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
                  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
                  <Relationship Id="rIdEndnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:t>Markers</w:t></w:r>
                      <w:r><w:footnoteReference w:id="5"/></w:r>
                      <w:r><w:endnoteReference w:id="7"/></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
            ["word/footnotes.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:footnote w:id="5">
                    <w:p>
                      <w:r><w:t>Foot </w:t></w:r>
                      <w:fldSimple w:instr=" PAGE "/>
                      <w:r><w:t> of </w:t></w:r>
                      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                      <w:r><w:instrText> NUMPAGES </w:instrText></w:r>
                      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                      <w:r><w:t>9</w:t></w:r>
                      <w:r><w:fldChar w:fldCharType="end"/></w:r>
                    </w:p>
                  </w:footnote>
                </w:footnotes>
                """,
            ["word/endnotes.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:endnote w:id="7">
                    <w:p>
                      <w:r><w:t>End </w:t></w:r>
                      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                      <w:r><w:instrText> PAGE </w:instrText></w:r>
                      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                      <w:r><w:t>3</w:t></w:r>
                      <w:r><w:fldChar w:fldCharType="end"/></w:r>
                      <w:r><w:t> of </w:t></w:r>
                      <w:fldSimple w:instr=" NUMPAGES "/>
                    </w:p>
                  </w:endnote>
                </w:endnotes>
                """
        });

        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxRelatedStory footnoteStory = document.RelatedStories.Single(story => story.Kind == DocxRelatedStoryKind.Footnote && story.Id == "5");
        DocxRelatedStory endnoteStory = document.RelatedStories.Single(story => story.Kind == DocxRelatedStoryKind.Endnote && story.Id == "7");
        DocxParagraph footnote = footnoteStory.Paragraphs.Single();
        DocxParagraph endnote = endnoteStory.Paragraphs.Single();

        TestAssert.Equal("Foot {PAGE} of {NUMPAGES}", string.Concat(footnote.Runs.Select(run => run.Text)));
        TestAssert.Equal("End {PAGE} of {NUMPAGES}", string.Concat(endnote.Runs.Select(run => run.Text)));
        TestAssert.Equal(2, footnote.FieldReferences.Count);
        TestAssert.Equal(2, endnote.FieldReferences.Count);
        DocxFieldReference footnotePage = footnote.FieldReferences.Single(field => field.Kind == DocxFieldKind.Page);
        DocxFieldReference footnoteNumPages = footnote.FieldReferences.Single(field => field.Kind == DocxFieldKind.NumPages);
        DocxFieldReference endnotePage = endnote.FieldReferences.Single(field => field.Kind == DocxFieldKind.Page);
        DocxFieldReference endnoteNumPages = endnote.FieldReferences.Single(field => field.Kind == DocxFieldKind.NumPages);
        TestAssert.True(footnotePage.SourceKind == DocxFieldSourceKind.Simple && footnotePage.UsesPlaceholder && !footnotePage.HasCachedResult, "Simple PAGE fields in footnotes should remain dynamic placeholders.");
        TestAssert.True(footnoteNumPages.SourceKind == DocxFieldSourceKind.ComplexInstruction && footnoteNumPages.HasSeparate && footnoteNumPages.HasCachedResult && !footnoteNumPages.RendersCachedResult && footnoteNumPages.UsesPlaceholder, "Complex NUMPAGES fields in footnotes should keep metadata without rendering stale cached results.");
        TestAssert.True(endnotePage.SourceKind == DocxFieldSourceKind.ComplexInstruction && endnotePage.HasSeparate && endnotePage.HasCachedResult && !endnotePage.RendersCachedResult && endnotePage.UsesPlaceholder, "Complex PAGE fields in endnotes should keep metadata without rendering stale cached results.");
        TestAssert.True(endnoteNumPages.SourceKind == DocxFieldSourceKind.Simple && endnoteNumPages.UsesPlaceholder && !endnoteNumPages.HasCachedResult, "Simple NUMPAGES fields in endnotes should remain dynamic placeholders.");

        DocxStructureSnapshot snapshot = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectStructure(document);
        TestAssert.Equal(4, snapshot.FieldReferenceCount);
        TestAssert.Equal(2, snapshot.PageFieldReferenceCount);
        TestAssert.Equal(2, snapshot.NumPagesFieldReferenceCount);
        TestAssert.Equal(2, snapshot.ComplexFieldReferenceCount);
        TestAssert.Equal(2, snapshot.CachedResultFieldReferenceCount);
        TestAssert.Equal(0, snapshot.RenderedCachedResultFieldReferenceCount);
        TestAssert.Equal(4, snapshot.PlaceholderFieldReferenceCount);
        TestAssert.Equal(4, snapshot.DynamicFieldReferenceCount);
        TestAssert.Equal(4, snapshot.DynamicPlaceholderFieldReferenceCount);
        TestAssert.Equal(2, snapshot.DynamicCachedResultNotRenderedFieldReferenceCount);
        DocxStructureStorySnapshot footnoteSnapshot = snapshot.Stories.Single(story => story.Kind == "Footnote" && story.VariantType == "5");
        DocxStructureStorySnapshot endnoteSnapshot = snapshot.Stories.Single(story => story.Kind == "Endnote" && story.VariantType == "7");
        TestAssert.True(footnoteSnapshot.FieldReferenceCount == 2 && footnoteSnapshot.ComplexFieldReferenceCount == 1 && footnoteSnapshot.PlaceholderFieldReferenceCount == 2, "Footnote story snapshots should expose dynamic field ownership.");
        TestAssert.True(endnoteSnapshot.FieldReferenceCount == 2 && endnoteSnapshot.ComplexFieldReferenceCount == 1 && endnoteSnapshot.PlaceholderFieldReferenceCount == 2, "Endnote story snapshots should expose dynamic field ownership.");
    }

    public static void DocxMarkupNoteLinksFieldsFixturePreservesPlacedNoteLinksAndFields()
    {
        string input = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "Cases",
            "docx-markup-note-links-fields.docx"));
        DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);
        DocxRelatedStory footnoteStory = document.RelatedStories.Single(story => story.Kind == DocxRelatedStoryKind.Footnote && story.Id == "5");
        DocxRelatedStory endnoteStory = document.RelatedStories.Single(story => story.Kind == DocxRelatedStoryKind.Endnote && story.Id == "7");
        DocxParagraph footnote = footnoteStory.Paragraphs.Single();
        DocxParagraph endnote = endnoteStory.Paragraphs.Single();

        TestAssert.Equal(2, footnote.FieldReferences.Count);
        TestAssert.Equal(2, endnote.FieldReferences.Count);
        TestAssert.Equal("https://example.invalid/markup-footnote", footnote.Hyperlinks.Single().Target ?? string.Empty);
        TestAssert.Equal("https://example.invalid/markup-endnote", endnote.Hyperlinks.Single().Target ?? string.Empty);
        TestAssert.True(footnote.FieldReferences.Any(field => field.Kind == DocxFieldKind.Page && field.SourceKind == DocxFieldSourceKind.Simple && field.UsesPlaceholder), "The fixture should keep a dynamic footnote PAGE placeholder.");
        TestAssert.True(footnote.FieldReferences.Any(field => field.Kind == DocxFieldKind.NumPages && field.SourceKind == DocxFieldSourceKind.ComplexInstruction && field.HasCachedResult && !field.RendersCachedResult), "The fixture should keep complex footnote NUMPAGES metadata.");
        TestAssert.True(endnote.FieldReferences.Any(field => field.Kind == DocxFieldKind.Page && field.SourceKind == DocxFieldSourceKind.ComplexInstruction && field.HasCachedResult && !field.RendersCachedResult), "The fixture should keep complex endnote PAGE metadata.");
        TestAssert.True(endnote.FieldReferences.Any(field => field.Kind == DocxFieldKind.NumPages && field.SourceKind == DocxFieldSourceKind.Simple && field.UsesPlaceholder), "The fixture should keep a dynamic endnote NUMPAGES placeholder.");

        var renderer = new DocxRenderer(
            fontResolver: null,
            markupMode: OoxPdfDocxMarkupMode.AllMarkup,
            markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        DocxTextEmissionLineSnapshot[] noteLines = renderer.InspectTextEmission(document).Lines
            .Where(line => line.StoryKind == "Footnote" || line.StoryKind == "Endnote")
            .ToArray();
        PdfLinkAnnotation[] annotations = renderer.RenderBlankPages(document, null, CancellationToken.None)
            .SelectMany(page => page.Annotations)
            .ToArray();

        TestAssert.True(noteLines.Any(line => line.StoryKind == "Footnote" && line.PageIndex == 0), "The fixture should place footnote fields on the marker page.");
        TestAssert.True(noteLines.Any(line => line.StoryKind == "Endnote" && line.PageIndex == 1), "The fixture should place endnote fields on the document-end page after pagination.");
        TestAssert.True(
            annotations.Any(annotation => annotation.Uri == "https://example.invalid/markup-footnote" && annotation.X >= document.MarginLeftPoints && annotation.Width > 0d),
            "The fixture should render an external footnote hyperlink annotation in page coordinates.");
        TestAssert.True(
            annotations.Any(annotation => annotation.Uri == "https://example.invalid/markup-endnote" && annotation.X >= document.MarginLeftPoints && annotation.Width > 0d),
            "The fixture should render an external endnote hyperlink annotation in page coordinates.");
    }

    public static void DocxComplexFieldWithCachedResultDoesNotEmitUnsupportedDiagnostic()
    {
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
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                      <w:r><w:instrText> DATE \@ &quot;yyyy&quot; </w:instrText></w:r>
                      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                      <w:r><w:t>2026</w:t></w:r>
                      <w:r><w:fldChar w:fldCharType="end"/></w:r>
                      <w:r><w:t> After</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxTextRun[] runs = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final).Paragraphs[0].Runs.ToArray();

        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.DoesNotContain("DOCX_UNSUPPORTED_COMPLEX_FIELD", ids);
        TestAssert.Equal("Before ", runs[0].Text);
        TestAssert.Equal("2026", runs[1].Text);
        TestAssert.Equal(" After", runs[2].Text);
    }

    public static void DocxCrossReferenceComplexFieldCachedResultInsideHyperlinkDoesNotEmitUnsupportedDiagnostic()
    {
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
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/field" TargetMode="External"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:hyperlink r:id="rIdLink">
                        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                        <w:r><w:instrText> PAGEREF Target \h </w:instrText></w:r>
                        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                        <w:r><w:t>section 2</w:t></w:r>
                        <w:r><w:fldChar w:fldCharType="end"/></w:r>
                      </w:hyperlink>
                      <w:r><w:t> After</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxParagraph paragraph = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final).Paragraphs[0];

        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.DoesNotContain("DOCX_UNSUPPORTED_COMPLEX_FIELD", ids);
        TestAssert.Equal("Before ", paragraph.Runs[0].Text);
        TestAssert.Equal("section 2", paragraph.Runs[1].Text);
        TestAssert.Equal(" After", paragraph.Runs[2].Text);
        TestAssert.Equal(1, paragraph.Hyperlinks.Count);
        TestAssert.Equal(1, paragraph.Hyperlinks[0].TextRunStartIndex);
        TestAssert.Equal(1, paragraph.Hyperlinks[0].TextRunCount);
        DocxFieldReference field = paragraph.FieldReferences.Single();
        TestAssert.Equal(DocxFieldKind.Other, field.Kind);
        TestAssert.True(field.HasCachedResult, "Cross-reference fields should record cached result availability.");
        TestAssert.True(field.RendersCachedResult, "Cross-reference fields should render their cached result.");
        TestAssert.True(!field.UsesPlaceholder, "Cross-reference fields should not use PAGE placeholders.");
        TestAssert.Equal(1, field.TextRunCount);
        TestAssert.Equal("section 2".Length, field.TextLength);
    }

    public static void DocxNestedComplexFieldsUseCachedResultsAndInspectionCounters()
    {
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
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                      <w:r><w:instrText> REF Outer </w:instrText></w:r>
                      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                      <w:r><w:t>Outer </w:t></w:r>
                      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                      <w:r><w:instrText> DATE \@ &quot;yyyy&quot; </w:instrText></w:r>
                      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                      <w:r><w:t>2026</w:t></w:r>
                      <w:r><w:fldChar w:fldCharType="end"/></w:r>
                      <w:r><w:t> Done</w:t></w:r>
                      <w:r><w:fldChar w:fldCharType="end"/></w:r>
                      <w:r><w:t> After</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxParagraph paragraph = document.Paragraphs.Single();

        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.DoesNotContain("DOCX_UNSUPPORTED_COMPLEX_FIELD", ids);
        TestAssert.Equal("Before Outer 2026 Done After", string.Concat(paragraph.Runs.Select(run => run.Text)));
        TestAssert.Equal(2, paragraph.FieldReferences.Count);

        DocxFieldReference outer = paragraph.FieldReferences.Single(field => field.Instruction?.Contains("REF Outer", StringComparison.Ordinal) == true);
        DocxFieldReference inner = paragraph.FieldReferences.Single(field => field.Instruction?.Contains("DATE", StringComparison.Ordinal) == true);
        TestAssert.True(outer.HasCachedResult && outer.RendersCachedResult, "Outer field should render its cached result span.");
        TestAssert.True(inner.HasCachedResult && inner.RendersCachedResult, "Nested field should render its cached result span.");
        TestAssert.Equal(0, outer.NestingDepth);
        TestAssert.Equal(1, inner.NestingDepth);
        TestAssert.Equal(3, outer.TextRunCount);
        TestAssert.Equal("Outer 2026 Done".Length, outer.TextLength);
        TestAssert.Equal(1, inner.TextRunCount);
        TestAssert.Equal("2026".Length, inner.TextLength);

        DocxStructureSnapshot snapshot = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectStructure(document);
        TestAssert.Equal(2, snapshot.ComplexFieldReferenceCount);
        TestAssert.Equal(2, snapshot.CachedResultFieldReferenceCount);
        TestAssert.Equal(2, snapshot.RenderedCachedResultFieldReferenceCount);
        TestAssert.Equal(0, snapshot.PlaceholderFieldReferenceCount);
        TestAssert.Equal(1, snapshot.NestedFieldReferenceCount);
    }

    public static void DocxComplexFieldsWithCachedResultsInsideDeletedAndMovedFromContentDoNotEmitUnsupportedDiagnostic()
    {
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
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:del w:id="41" w:author="A" w:date="2026-06-10T00:00:00Z">
                        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                        <w:r><w:instrText> REF DeletedTarget </w:instrText></w:r>
                        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                        <w:r><w:delText>deleted-field </w:delText></w:r>
                        <w:r><w:fldChar w:fldCharType="end"/></w:r>
                      </w:del>
                      <w:moveFrom w:id="42" w:author="B" w:date="2026-06-10T00:00:00Z">
                        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                        <w:r><w:instrText> REF MoveTarget </w:instrText></w:r>
                        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                        <w:r><w:t>moved-field</w:t></w:r>
                        <w:r><w:fldChar w:fldCharType="end"/></w:r>
                      </w:moveFrom>
                      <w:r><w:t> After</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            DocxMarkupMode = OoxPdfDocxMarkupMode.AllMarkup,
            DiagnosticSink = diagnostics.Add
        });
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxParagraph paragraph = new DocxReader().Read(package, null, CancellationToken.None, markupMode: OoxPdfDocxMarkupMode.AllMarkup).Paragraphs.Single();

        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.DoesNotContain("DOCX_UNSUPPORTED_COMPLEX_FIELD", ids);
        TestAssert.Equal("Before deleted-field moved-field After", string.Concat(paragraph.Runs.Select(run => run.Text)));
        TestAssert.Equal(2, paragraph.FieldReferences.Count);
        TestAssert.True(paragraph.FieldReferences.All(field => field.HasCachedResult && field.RendersCachedResult), "Cached complex fields in review containers should render their stored results.");
        TestAssert.True(paragraph.FieldReferences.All(field => field.Kind == DocxFieldKind.Other), "Cached REF fields should not be treated as dynamic PAGE placeholders.");
        TestAssert.True(paragraph.Runs.Any(run => run.Revision?.Kind == DocxRevisionKind.Deletion), "Deleted field result runs should keep deletion provenance.");
        TestAssert.True(paragraph.Runs.Any(run => run.Revision?.Kind == DocxRevisionKind.MoveFrom), "Moved-from field result runs should keep move-from provenance.");
    }

    public static void DocxComplexFieldWithCachedResultInsideInlineContentControlDoesNotEmitUnsupportedDiagnostic()
    {
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
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:sdt>
                        <w:sdtPr><w:alias w:val="Synthetic content control"/></w:sdtPr>
                        <w:sdtContent>
                          <w:r><w:t>control </w:t></w:r>
                          <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                          <w:r><w:instrText> REF ControlTarget </w:instrText></w:r>
                          <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                          <w:r><w:t>cached-ref</w:t></w:r>
                          <w:r><w:fldChar w:fldCharType="end"/></w:r>
                        </w:sdtContent>
                      </w:sdt>
                      <w:r><w:t> After</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxParagraph paragraph = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final).Paragraphs.Single();

        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.DoesNotContain("DOCX_UNSUPPORTED_COMPLEX_FIELD", ids);
        TestAssert.Equal("Before control cached-ref After", string.Concat(paragraph.Runs.Select(run => run.Text)));
        DocxFieldReference field = paragraph.FieldReferences.Single();
        TestAssert.Equal(DocxFieldKind.Other, field.Kind);
        TestAssert.True(field.HasCachedResult && field.RendersCachedResult, "Cached complex fields inside inline content controls should render their stored result.");
        TestAssert.Equal(1, field.TextRunCount);
        TestAssert.Equal("cached-ref".Length, field.TextLength);
    }

    public static void DocxComplexFieldSpanningInlineContentControlUsesCachedResult()
    {
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
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                      <w:r><w:instrText> REF ControlTarget </w:instrText></w:r>
                      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                      <w:sdt>
                        <w:sdtPr><w:tag w:val="synthetic-field-result"/></w:sdtPr>
                        <w:sdtContent>
                          <w:r><w:t>cached-ref</w:t></w:r>
                        </w:sdtContent>
                      </w:sdt>
                      <w:r><w:fldChar w:fldCharType="end"/></w:r>
                      <w:r><w:t> After</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxParagraph paragraph = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final).Paragraphs.Single();

        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.DoesNotContain("DOCX_UNSUPPORTED_COMPLEX_FIELD", ids);
        TestAssert.Equal("Before cached-ref After", string.Concat(paragraph.Runs.Select(run => run.Text)));
        DocxFieldReference field = paragraph.FieldReferences.Single();
        TestAssert.True(field.HasCachedResult && field.RendersCachedResult, "Complex fields spanning inline content controls should keep their cached result span.");
        TestAssert.Equal(1, field.TextRunCount);
        TestAssert.Equal("cached-ref".Length, field.TextLength);
    }

    public static void DocxMultiBlockComplexFieldWithCachedResultDoesNotEmitUnsupportedDiagnostic()
    {
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
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                      <w:r><w:instrText> TOC \o &quot;1-2&quot; </w:instrText></w:r>
                      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                      <w:r><w:t>Entry one</w:t></w:r>
                    </w:p>
                    <w:p>
                      <w:r><w:t>Entry two</w:t></w:r>
                      <w:r><w:fldChar w:fldCharType="end"/></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.DoesNotContain("DOCX_UNSUPPORTED_COMPLEX_FIELD", ids);
        TestAssert.Equal("Entry one|Entry two", string.Join("|", document.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text)))));
    }

    public static void DocxMalformedComplexFieldWithCachedResultKeepsRenderedResultMetadata()
    {
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
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                      <w:r><w:instrText> REF BrokenTarget </w:instrText></w:r>
                      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                      <w:r><w:t>cached-ref</w:t></w:r>
                      <w:r><w:t> After</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxParagraph paragraph = document.Paragraphs.Single();

        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.DoesNotContain("DOCX_UNSUPPORTED_COMPLEX_FIELD", ids);
        TestAssert.Equal("Before cached-ref After", string.Concat(paragraph.Runs.Select(run => run.Text)));
        DocxFieldReference field = paragraph.FieldReferences.Single();
        TestAssert.Equal(DocxFieldKind.Other, field.Kind);
        TestAssert.True(field.HasSeparate, "Malformed cached fields should record the separate marker.");
        TestAssert.True(field.HasCachedResult && field.RendersCachedResult, "Malformed cached fields should retain rendered-result metadata even without an end marker.");
        TestAssert.Equal(2, field.TextRunCount);
        TestAssert.Equal("cached-ref After".Length, field.TextLength);
    }

    public static void DocxReaderPreservesInlineReferencesInsideRunContainers()
    {
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
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:t>Base</w:t></w:r>
                      <w:ins>
                        <w:r><w:t>Ins</w:t><w:footnoteReference w:id="5" w:customMarkFollows="1"/></w:r>
                      </w:ins>
                      <w:hyperlink w:anchor="Target">
                        <w:r><w:commentReference w:id="6"/><w:t>Link</w:t></w:r>
                      </w:hyperlink>
                      <w:fldSimple w:instr=" REF Target ">
                        <w:r><w:t>Field</w:t><w:endnoteReference w:id="7"/></w:r>
                      </w:fldSimple>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxParagraph paragraph = document.Paragraphs.Single();

        TestAssert.Equal(3, paragraph.InlineReferences.Count);
        DocxInlineReference footnote = paragraph.InlineReferences.Single(reference => reference.Kind == DocxRelatedStoryKind.Footnote);
        DocxInlineReference comment = paragraph.InlineReferences.Single(reference => reference.Kind == DocxRelatedStoryKind.Comment);
        DocxInlineReference endnote = paragraph.InlineReferences.Single(reference => reference.Kind == DocxRelatedStoryKind.Endnote);
        TestAssert.True(footnote.Id == "5" && footnote.CustomMarkFollowsValue == "1", "Inserted footnote marker metadata should survive run-container parsing.");
        TestAssert.True(comment.Id == "6" && comment.CustomMarkFollowsValue is null, "Hyperlink comment marker metadata should survive run-container parsing.");
        TestAssert.True(endnote.Id == "7" && endnote.CustomMarkFollowsValue is null, "Simple-field endnote marker metadata should survive run-container parsing.");
        TestAssert.True(footnote.DisplayText is null, "Footnote references with customMarkFollows should not synthesize an automatic marker.");
        TestAssert.True(comment.DisplayText is null, "Comment references should stay structural until comment display rules are modeled.");
        TestAssert.Equal("1", endnote.DisplayText ?? string.Empty);
        TestAssert.Equal(1, footnote.SourceRunIndex);
        TestAssert.Equal(1, footnote.RunChildIndex);
        TestAssert.Equal(3, footnote.TextOffsetInRun);
        TestAssert.Equal(2, comment.SourceRunIndex);
        TestAssert.Equal(0, comment.RunChildIndex);
        TestAssert.Equal(0, comment.TextOffsetInRun);
        TestAssert.Equal(3, endnote.SourceRunIndex);
        TestAssert.Equal(1, endnote.RunChildIndex);
        TestAssert.Equal(5, endnote.TextOffsetInRun);

        DocxStructureSnapshot snapshot = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectStructure(document);
        DocxStructureBlockSnapshot block = snapshot.Blocks.Single(block => block.Kind == "Paragraph");
        TestAssert.Equal(3, snapshot.InlineReferenceCount);
        TestAssert.Equal(3, snapshot.AnchoredInlineReferenceCount);
        TestAssert.Equal(1, block.CommentReferenceCount);
        TestAssert.Equal(1, block.FootnoteReferenceCount);
        TestAssert.Equal(1, block.EndnoteReferenceCount);
    }

    public static void DocxReaderEmitsAutomaticFootnoteAndEndnoteMarkersAsSuperscriptRuns()
    {
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
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:t>Before</w:t><w:footnoteReference w:id="2"/><w:t>Middle</w:t><w:endnoteReference w:id="3"/><w:t>After</w:t></w:r>
                    </w:p>
                    <w:p>
                      <w:r><w:t>Next</w:t><w:footnoteReference w:id="4"/></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxParagraph first = document.Paragraphs[0];
        TestAssert.Equal(5, first.Runs.Count);
        TestAssert.Equal("Before", first.Runs[0].Text);
        TestAssert.Equal("1", first.Runs[1].Text);
        TestAssert.Equal("superscript", first.Runs[1].VerticalAlignmentValue ?? string.Empty);
        TestAssert.Equal("Middle", first.Runs[2].Text);
        TestAssert.Equal("1", first.Runs[3].Text);
        TestAssert.Equal("superscript", first.Runs[3].VerticalAlignmentValue ?? string.Empty);
        TestAssert.Equal("After", first.Runs[4].Text);
        TestAssert.Equal("1", first.InlineReferences.Single(reference => reference.Kind == DocxRelatedStoryKind.Footnote).DisplayText ?? string.Empty);
        TestAssert.Equal("1", first.InlineReferences.Single(reference => reference.Kind == DocxRelatedStoryKind.Endnote).DisplayText ?? string.Empty);

        DocxParagraph second = document.Paragraphs[1];
        TestAssert.Equal("2", second.Runs[1].Text);
        TestAssert.Equal("2", second.InlineReferences.Single().DisplayText ?? string.Empty);

        DocxStructureSnapshot snapshot = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectStructure(document);
        DocxStructureInlineReferenceSnapshot[] references = snapshot.InlineReferences.ToArray();
        TestAssert.Equal("1", references.Single(reference => reference.Kind == "Footnote" && reference.Id == "2").DisplayText ?? string.Empty);
        TestAssert.Equal("1", references.Single(reference => reference.Kind == "Endnote").DisplayText ?? string.Empty);
        TestAssert.Equal("2", references.Single(reference => reference.Kind == "Footnote" && reference.Id == "4").DisplayText ?? string.Empty);
    }

    public static void DocxReaderPreservesBookmarkAnchorsForInternalHyperlinks()
    {
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
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:bookmarkStart w:id="7" w:name="Target"/>
                      <w:r><w:t>Target</w:t></w:r>
                      <w:bookmarkEnd w:id="7"/>
                      <w:hyperlink w:anchor="Target"><w:r><w:t>Jump</w:t></w:r></w:hyperlink>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxParagraph paragraph = document.Paragraphs.Single();

        DocxBookmarkAnchor bookmark = paragraph.BookmarkAnchors.Single();
        TestAssert.Equal("7", bookmark.Id ?? string.Empty);
        TestAssert.Equal("Target", bookmark.Name ?? string.Empty);
        TestAssert.Equal(1, bookmark.SourceRunIndex);
        TestAssert.Equal(1, bookmark.TextRunIndex);
        TestAssert.Equal(7, bookmark.TextOffset);

        DocxHyperlinkSpan link = paragraph.Hyperlinks.Single();
        TestAssert.Equal("Target", link.Anchor ?? string.Empty);
        TestAssert.Equal(2, link.TextRunStartIndex);
        TestAssert.Equal(1, link.TextRunCount);

        DocxStructureSnapshot snapshot = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectStructure(document);
        DocxStructureBlockSnapshot block = snapshot.Blocks.Single(block => block.Kind == "Paragraph");
        DocxStructureStorySnapshot bodyStory = snapshot.Stories.Single(story => story.Kind == "Body");
        TestAssert.Equal(1, snapshot.BookmarkAnchorCount);
        TestAssert.Equal(1, block.BookmarkAnchorCount);
        TestAssert.Equal(1, bodyStory.BookmarkAnchorCount);
        TestAssert.Equal(1, snapshot.InternalHyperlinkCount);
    }

    public static void DocxReaderPreservesBookmarkAnchorsInsideHyperlinks()
    {
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
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:t>Before </w:t></w:r>
                      <w:hyperlink w:anchor="OuterTarget">
                        <w:bookmarkStart w:id="11" w:name="InnerTarget"/>
                        <w:r><w:t>Linked</w:t></w:r>
                      </w:hyperlink>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxParagraph paragraph = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final).Paragraphs.Single();

        DocxBookmarkAnchor bookmark = paragraph.BookmarkAnchors.Single();
        TestAssert.Equal("InnerTarget", bookmark.Name ?? string.Empty);
        TestAssert.Equal(1, bookmark.SourceRunIndex);
        TestAssert.Equal(1, bookmark.TextRunIndex);
        TestAssert.Equal(7, bookmark.TextOffset);
        TestAssert.Equal("OuterTarget", paragraph.Hyperlinks.Single().Anchor ?? string.Empty);
    }
    public static void DocxLiteralFieldLikeTextSurvivesWhileActualFieldsResolve()
    {
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
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:r><w:t>Literal {PAGE} and {NUMPAGES}</w:t></w:r></w:p>
                    <w:p><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
        });

        DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
        TestAssert.Equal(2, document.Paragraphs.Count);
        DocxParagraph literal = document.Paragraphs[0];
        TestAssert.True(literal.Runs.All(run => run.FieldKind is null), "Ordinary text containing braces must not be typed as a field.");
        TestAssert.True(literal.FieldReferences.Count == 0, "Literal braces must not produce field references.");
        DocxParagraph field = document.Paragraphs[1];
        TestAssert.True(field.Runs.Any(run => run.FieldKind == DocxFieldKind.Page && run.Text == "{PAGE}"), "Actual PAGE fields must retain typed identity.");

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxTextLineLayout[] lines = layout.Pages[0].Items.OfType<DocxTextLineLayout>().ToArray();
        TestAssert.True(lines.Any(line => line.Text.Contains("Literal {PAGE} and {NUMPAGES}", StringComparison.Ordinal)), "Literal field-like text must survive layout, got: " + string.Join("|", lines.Select(line => line.Text)));
        TestAssert.True(lines.All(line => !line.Text.Equals("Literal 1 and 1", StringComparison.Ordinal)), "Literals must not be substituted as fields.");
        TestAssert.True(lines.Any(line => line.Text == "1"), "Actual PAGE field must resolve to the page number.");
    }
    public static void DocxReaderSizesRunlessFieldPlaceholderThroughStyleCascade()
    {
        // Office A/B 2026-09-07 (probe-fieldsize): styles-less empty PAGE fields compute at 12pt.
        // Placeholders with no runs resolve like runs (style chain, then document defaults, then flat 12pt).
        string bare = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:fldSimple w:instr=" PAGE "/></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
        });
        string styled = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>""",
            ["word/styles.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:style></w:styles>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:fldSimple w:instr=" PAGE "/></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
        });
        DocxDocument bareDocument = DocxTests.ReadDocx(bare, OoxPdfDocxMarkupMode.Final);
        TestAssert.Equal(12d, bareDocument.Paragraphs.Single().Runs.Single(run => run.Text == "{PAGE}").FontSize);
        DocxDocument styledDocument = DocxTests.ReadDocx(styled, OoxPdfDocxMarkupMode.Final);
        TestAssert.Equal(11d, styledDocument.Paragraphs.Single().Runs.Single(run => run.Text == "{PAGE}").FontSize);
    }
}
