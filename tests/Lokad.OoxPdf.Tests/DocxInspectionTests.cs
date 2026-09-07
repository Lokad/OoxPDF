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

internal static class DocxInspectionTests
{
    public static void DocxMarkupInspectionSnapshotsExposePrivateSafeRevisionCounts()
    {
        string input = DocxTests.WriteTrackedChangeModeProbeDocx();
        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, markupMode: OoxPdfDocxMarkupMode.AllMarkup);
        var renderer = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);

        DocxStructureSnapshot structure = renderer.InspectStructure(document);
        DocxLayoutSnapshot layout = renderer.InspectLayout(document);
        DocxTextEmissionSnapshot textEmission = renderer.InspectTextEmission(document);

        TestAssert.Equal("AllMarkup", structure.MarkupMode);
        TestAssert.Equal(4, structure.RevisionCount);
        TestAssert.Equal(1, structure.InsertionRevisionCount);
        TestAssert.Equal(1, structure.DeletionRevisionCount);
        TestAssert.Equal(1, structure.MoveFromRevisionCount);
        TestAssert.Equal(1, structure.MoveToRevisionCount);
        TestAssert.Equal(0, structure.OtherRevisionCount);

        DocxStructureBlockSnapshot block = structure.Blocks.Single(block => block.Kind == "Paragraph");
        TestAssert.Equal(4, block.RevisionCount);
        TestAssert.Equal("AllMarkup", layout.MarkupMode);
        TestAssert.True(layout.RevisionCount >= 4, "Layout inspection should expose visible revision ownership without document text.");
        TestAssert.True(layout.InsertionRevisionCount >= 1 && layout.DeletionRevisionCount >= 1, "Layout inspection should preserve revision kind counts.");
        TestAssert.True(layout.Pages.Single().RevisionCount >= 4, "Page layout inspection should expose page-local revision counts.");

        TestAssert.Equal("AllMarkup", textEmission.MarkupMode);
        TestAssert.True(textEmission.RevisionSegmentCount >= 4, "Text emission inspection should expose revised emitted segments without their text.");
        TestAssert.True(textEmission.InsertionRevisionSegmentCount >= 1 && textEmission.DeletionRevisionSegmentCount >= 1, "Text emission inspection should preserve revision kind counts.");
        TestAssert.True(textEmission.Lines.Any(line => line.RevisionSegmentCount >= 4), "Line inspection should expose line-local revision segment counts.");
        TestAssert.True(textEmission.Lines.SelectMany(line => line.Segments).Any(segment => segment.RevisionKind == "MoveTo"), "Segment inspection should expose private-safe revision kinds.");
    }

    public static void DocxMarkupInspectionSnapshotsExposePrivateSafeCommentCounts()
    {
        string input = DocxTests.WriteCommentMarkerProbeDocx();
        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, markupMode: OoxPdfDocxMarkupMode.SimpleMarkup);
        var renderer = new DocxRenderer(null, OoxPdfDocxMarkupMode.SimpleMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);

        DocxStructureSnapshot structure = renderer.InspectStructure(document);
        DocxLayoutSnapshot layout = renderer.InspectLayout(document);
        DocxTextEmissionSnapshot textEmission = renderer.InspectTextEmission(document);

        TestAssert.Equal("SimpleMarkup", structure.MarkupMode);
        TestAssert.Equal(1, structure.CommentReferenceCount);
        TestAssert.Equal(1, structure.Blocks.Single(block => block.Kind == "Paragraph").CommentReferenceCount);
        TestAssert.Equal("SimpleMarkup", layout.MarkupMode);
        TestAssert.True(layout.CommentReferenceCount >= 1, "Layout inspection should expose visible comment-reference ownership without comment text.");
        TestAssert.Equal("SimpleMarkup", textEmission.MarkupMode);
        TestAssert.True(textEmission.CommentReferenceCount >= 1, "Text emission inspection should expose comment-reference ownership without comment text.");
    }

    public static void DocxMarkupInspectionClassifiesCommentStoryAnchors()
    {
        string input = DocxTests.WriteCommentAnchorAccountingProbeDocx();
        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, markupMode: OoxPdfDocxMarkupMode.Final);

        TestAssert.Equal("1|2", string.Join("|", document.PackageCommentAnchorIds));
        TestAssert.Equal("2", string.Join("|", document.HiddenCommentAnchorIds));

        DocxStructureSnapshot structure = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectStructure(document);
        TestAssert.Equal(2, structure.PackageCommentAnchorIdCount);
        TestAssert.Equal(1, structure.HiddenCommentAnchorIdCount);
        TestAssert.Equal(1, structure.ResolvedCommentStoryAnchorCount);
        TestAssert.Equal(1, structure.HiddenCommentStoryAnchorCount);
        TestAssert.Equal(1, structure.OrphanedCommentStoryAnchorCount);
        TestAssert.Equal(1, structure.UnsupportedCommentStoryAnchorCount);

        DocxStructureCommentStoryAnchorSnapshot visible = structure.CommentStoryAnchors!.Single(anchor => anchor.Id == "1");
        DocxStructureCommentStoryAnchorSnapshot hidden = structure.CommentStoryAnchors!.Single(anchor => anchor.Id == "2");
        DocxStructureCommentStoryAnchorSnapshot orphaned = structure.CommentStoryAnchors!.Single(anchor => anchor.Id == "3");
        DocxStructureCommentStoryAnchorSnapshot unsupported = structure.CommentStoryAnchors!.Single(anchor => anchor.Status == "Unsupported");
        TestAssert.True(visible.Status == "Visible" && visible.VisibleInlineReferenceCount == 1 && visible.HasPackageAnchor, "Visible comment bodies should be tied to a visible anchor.");
        TestAssert.True(hidden.Status == "HiddenByMarkupMode" && hidden.HasHiddenAnchor && hidden.VisibleInlineReferenceCount == 0, "Comment bodies anchored only in filtered revisions should be classified as hidden by the selected markup mode.");
        TestAssert.True(orphaned.Status == "Orphaned" && !orphaned.HasPackageAnchor, "Comment bodies without package anchors should be classified as orphaned without exposing body text.");
        TestAssert.True(unsupported.Id is null && !unsupported.HasPackageAnchor && unsupported.VisibleInlineReferenceCount == 0, "Comment bodies without usable ids should be classified as unsupported.");
    }

    public static void DocxMarkupInspectionClassifiesDeletedCommentAnchorsByMode()
    {
        string input = DocxTests.WriteCommentAnchorAccountingProbeDocx();
        DocxStructureCommentStoryAnchorSnapshot FinalAnchor(string id, OoxPdfDocxMarkupMode mode)
        {
            using FileStream stream = File.OpenRead(input);
            DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, markupMode: mode);
            return new DocxRenderer(null, mode, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .InspectStructure(document)
                .CommentStoryAnchors!
                .Single(anchor => anchor.Id == id);
        }

        DocxStructureCommentStoryAnchorSnapshot finalDeletedAnchor = FinalAnchor("2", OoxPdfDocxMarkupMode.Final);
        DocxStructureCommentStoryAnchorSnapshot simpleDeletedAnchor = FinalAnchor("2", OoxPdfDocxMarkupMode.SimpleMarkup);
        DocxStructureCommentStoryAnchorSnapshot originalDeletedAnchor = FinalAnchor("2", OoxPdfDocxMarkupMode.Original);
        DocxStructureCommentStoryAnchorSnapshot allDeletedAnchor = FinalAnchor("2", OoxPdfDocxMarkupMode.AllMarkup);

        TestAssert.True(finalDeletedAnchor.Status == "HiddenByMarkupMode" && finalDeletedAnchor.VisibleInlineReferenceCount == 0, "Final view should classify comments anchored only in deleted content as hidden by mode.");
        TestAssert.True(simpleDeletedAnchor.Status == "HiddenByMarkupMode" && simpleDeletedAnchor.VisibleInlineReferenceCount == 0, "Simple markup should classify comments anchored only in deleted content as hidden by mode.");
        TestAssert.True(originalDeletedAnchor.Status == "Visible" && originalDeletedAnchor.VisibleInlineReferenceCount == 1, "Original view should classify comments anchored in deleted content as visible.");
        TestAssert.True(allDeletedAnchor.Status == "Visible" && allDeletedAnchor.VisibleInlineReferenceCount == 1, "All markup should classify comments anchored in deleted content as visible.");
    }

    public static void DocxMarkupRenderingTreatsDeletedCommentAnchorsByMode()
    {
        string input = DocxTests.WriteCommentAnchorAccountingProbeDocx();

        (int StructureReferences, int LayoutReferences, int TextReferences, int RenderedCommentCandidates) Counts(OoxPdfDocxMarkupMode mode)
        {
            using FileStream stream = File.OpenRead(input);
            DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, markupMode: mode);
            var renderer = new DocxRenderer(null, mode, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            return (
                renderer.InspectStructure(document).CommentReferenceCount,
                renderer.InspectLayout(document).CommentReferenceCount,
                renderer.InspectTextEmission(document).CommentReferenceCount,
                renderer.InspectMarkupBalloons(document).Sum(placement => placement.CommentCandidateCount));
        }

        TestAssert.Equal((1, 1, 1, 0), Counts(OoxPdfDocxMarkupMode.Final));
        TestAssert.Equal((1, 1, 1, 0), Counts(OoxPdfDocxMarkupMode.SimpleMarkup));
        TestAssert.Equal((2, 2, 2, 0), Counts(OoxPdfDocxMarkupMode.Original));
        TestAssert.Equal((2, 2, 2, 2), Counts(OoxPdfDocxMarkupMode.AllMarkup));
    }

    public static void DocxReaderPreservesCommentMetadataAndThreading()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
                  <Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.ms-word.commentsExtended+xml"/>
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
                  <Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
                  <Relationship Id="rIdCommentsExtended" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:commentRangeStart w:id="1"/>
                      <w:r><w:t>Reviewed text</w:t></w:r>
                      <w:commentRangeEnd w:id="1"/>
                      <w:r><w:commentReference w:id="1"/></w:r>
                    </w:p>
                  </w:body>
                </w:document>
                """,
            ["word/comments.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
                  <w:comment w:id="1" w:author="Reviewer One" w:initials="RO" w:date="2024-01-02T03:04:05Z">
                    <w:p w14:paraId="11111111"><w:r><w:t>Parent comment</w:t></w:r></w:p>
                  </w:comment>
                  <w:comment w:id="2" w:author="Reviewer Two" w:initials="RT" w:date="2024-01-03T03:04:05Z">
                    <w:p w14:paraId="22222222"><w:r><w:t>Reply comment</w:t></w:r></w:p>
                  </w:comment>
                  <w:comment w:id="3" w:author="Reviewer Three" w:initials="R3" w:date="2024-01-04T03:04:05Z">
                    <w:p w14:paraId="33333333"><w:r><w:t>Second reply</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """,
            ["word/commentsExtended.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
                  <w15:commentEx w15:paraId="11111111" w15:done="1"/>
                  <w15:commentEx w15:paraId="22222222" w15:paraIdParent="11111111" w15:done="0"/>
                  <w15:commentEx w15:paraId="33333333" w15:paraIdParent="11111111" w15:done="0"/>
                </w15:commentsEx>
                """
        });
        using FileStream stream = File.OpenRead(input);

        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, markupMode: OoxPdfDocxMarkupMode.AllMarkup);

        TestAssert.Equal(3, document.RelatedStories.Count(story => story.Kind == DocxRelatedStoryKind.Comment));
        DocxRelatedStory parent = document.RelatedStories.Single(story => story.Kind == DocxRelatedStoryKind.Comment && story.Id == "1");
        DocxRelatedStory reply = document.RelatedStories.Single(story => story.Kind == DocxRelatedStoryKind.Comment && story.Id == "2");
        DocxRelatedStory secondReply = document.RelatedStories.Single(story => story.Kind == DocxRelatedStoryKind.Comment && story.Id == "3");
        TestAssert.True(parent.CommentMetadata?.Author == "Reviewer One" && parent.CommentMetadata.Initials == "RO" && parent.CommentMetadata.Date == "2024-01-02T03:04:05Z", "Classic comment author, initials, and date metadata should survive reading.");
        TestAssert.True(parent.CommentMetadata?.ParagraphId == "11111111" && parent.CommentMetadata.ParentParagraphId is null && parent.CommentMetadata.ParentCommentId is null && parent.CommentMetadata.IsResolved == true, "Parent comment extension metadata should preserve paragraph id and resolved state.");
        TestAssert.True(reply.CommentMetadata?.Author == "Reviewer Two" && reply.CommentMetadata.Initials == "RT" && reply.CommentMetadata.Date == "2024-01-03T03:04:05Z", "Reply comment classic metadata should survive reading.");
        TestAssert.True(reply.CommentMetadata?.ParagraphId == "22222222" && reply.CommentMetadata.ParentParagraphId == "11111111" && reply.CommentMetadata.ParentCommentId == "1" && reply.CommentMetadata.IsResolved == false, "Reply comment extension metadata should resolve the parent comment id through the parent paragraph id.");
        TestAssert.True(secondReply.CommentMetadata?.ParagraphId == "33333333" && secondReply.CommentMetadata.ParentParagraphId == "11111111" && secondReply.CommentMetadata.ParentCommentId == "1" && secondReply.CommentMetadata.IsResolved == false, "Second reply metadata should resolve to the same parent comment thread.");

        DocxStructureSnapshot structure = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectStructure(document);
        DocxStructureStorySnapshot parentSnapshot = structure.Stories.Single(story => story.Kind == "Comment" && story.VariantType == "1");
        DocxStructureStorySnapshot replySnapshot = structure.Stories.Single(story => story.Kind == "Comment" && story.VariantType == "2");
        DocxStructureStorySnapshot secondReplySnapshot = structure.Stories.Single(story => story.Kind == "Comment" && story.VariantType == "3");
        TestAssert.True(parentSnapshot.HasCommentAuthor && parentSnapshot.HasCommentInitials && parentSnapshot.HasCommentDate && parentSnapshot.CommentParagraphId == "11111111" && parentSnapshot.CommentResolved == true, "Structure snapshots should expose private-safe parent comment metadata presence and resolved state.");
        TestAssert.True(replySnapshot.HasCommentAuthor && replySnapshot.HasCommentInitials && replySnapshot.HasCommentDate && replySnapshot.CommentParagraphId == "22222222" && replySnapshot.CommentParentParagraphId == "11111111" && replySnapshot.CommentParentId == "1" && replySnapshot.CommentResolved == false, "Structure snapshots should expose private-safe threaded reply ownership.");
        TestAssert.True(secondReplySnapshot.HasCommentAuthor && secondReplySnapshot.CommentParagraphId == "33333333" && secondReplySnapshot.CommentParentParagraphId == "11111111" && secondReplySnapshot.CommentParentId == "1", "Structure snapshots should expose private-safe ownership for every reply in the thread.");

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxRelatedStoryLayout parentLayout = layout.RelatedStories.Single(story => story.Story.Id == "1");
        DocxRelatedStoryLayout replyLayout = layout.RelatedStories.Single(story => story.Story.Id == "2");
        DocxRelatedStoryLayout secondReplyLayout = layout.RelatedStories.Single(story => story.Story.Id == "3");
        TestAssert.Equal("Reviewer One (RO) #1 2024-01-02 resolved", DocxRenderer.BuildCommentBalloonTitle(parent, "1"));
        TestAssert.Equal("Reply Reviewer Two (RT) #2 2024-01-03 open", DocxRenderer.BuildCommentBalloonTitle(reply, "2"));
        TestAssert.Equal("Reply Reviewer Three (R3) #3 2024-01-04 open", DocxRenderer.BuildCommentBalloonTitle(secondReply, "3"));
        string threadedPreview = DocxRenderer.BuildCommentBalloonPreview(parentLayout, [replyLayout, secondReplyLayout]);
        TestAssert.Contains("2 replies", threadedPreview);
        TestAssert.Contains("Reply: Reply comment", threadedPreview);
        TestAssert.Contains("Reply: Second reply", threadedPreview);
        string wordCompatibleThreadedPreview = DocxRenderer.BuildWordCompatibleCommentBalloonPreview(parentLayout, [replyLayout, secondReplyLayout]);
        TestAssert.Contains("2024-01-02 resolved Parent comment", wordCompatibleThreadedPreview);
        TestAssert.Contains("2 replies", wordCompatibleThreadedPreview);
        TestAssert.Contains("Reply 2024-01-03 open Reply comment", wordCompatibleThreadedPreview);
        TestAssert.Contains("Reply 2024-01-04 open Second reply", wordCompatibleThreadedPreview);

        DocxMarkupBalloonPlacementSnapshot parentBalloon = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .InspectMarkupBalloons(document)
            .Single(placement => placement.Kind == "Comment");
        TestAssert.True(
            parentBalloon.CommentWithDateCount == 3 &&
            parentBalloon.CommentResolvedCount == 1 &&
            parentBalloon.CommentOpenCount == 2 &&
            parentBalloon.CommentReplyCount == 2,
            "Comment balloon placement snapshots should expose private-safe thread metadata counts for dated, resolved/open, and reply comments.");
    }

    public static void DocxInlineReferencesRetainRevisionProvenance()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
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
                  <Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:t>Text</w:t></w:r>
                      <w:ins w:id="9" w:author="A" w:date="2026-06-04T00:00:00Z">
                        <w:r><w:commentReference w:id="1"/></w:r>
                      </w:ins>
                    </w:p>
                  </w:body>
                </w:document>
                """,
            ["word/comments.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:comment w:id="1"><w:p><w:r><w:t>Comment</w:t></w:r></w:p></w:comment>
                </w:comments>
                """
        });
        using FileStream stream = File.OpenRead(input);

        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, markupMode: OoxPdfDocxMarkupMode.AllMarkup);

        DocxInlineReference reference = document.Paragraphs.Single().InlineReferences.Single();
        TestAssert.True(reference.Revision?.Kind == DocxRevisionKind.Insertion && reference.Revision.Author == "A" && reference.Revision.Date == "2026-06-04T00:00:00Z" && reference.Revision.SourceElement == "ins", "Inline references inside revision containers should retain direct revision metadata.");
        TestAssert.Equal(1, reference.Revisions.Count);

        DocxStructureInlineReferenceSnapshot snapshot = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .InspectStructure(document)
            .InlineReferences
            .Single();
        TestAssert.True(snapshot.RevisionCount == 1 && snapshot.RevisionKind == "Insertion" && snapshot.RevisionSourceElement == "ins", "Structure snapshots should expose private-safe inline-reference revision provenance.");
    }

    public static void DocxMarkupDiagnosticsApproximateCommentsInMarkupModes()
    {
        string input = DocxTests.WriteCommentMarkerProbeDocx();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            DocxMarkupMode = OoxPdfDocxMarkupMode.SimpleMarkup,
            DiagnosticSink = diagnostics.Add
        });

        TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_COMMENTS" && d.Fallback == "Approximated"), "Markup modes should report comments as approximated instead of broadly unsupported.");
        TestAssert.True(!diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_COMMENTS"), "Markup modes should not emit the broad unsupported comment diagnostic.");
    }

    public static void DocxMarkupDiagnosticsApproximateTrackedChangesInMarkupModes()
    {
        string input = DocxTests.WriteTrackedChangeModeProbeDocx();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            DocxMarkupMode = OoxPdfDocxMarkupMode.AllMarkup,
            DiagnosticSink = diagnostics.Add
        });

        TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_TRACKED_CHANGES" && d.Fallback == "Approximated"), "Markup modes should report tracked changes as approximated instead of broadly unsupported.");
        TestAssert.True(!diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_TRACKED_CHANGES"), "Markup modes should not emit the broad unsupported tracked-change diagnostic.");
    }

    public static void DocxMarkupDiagnosticsCoverFormattingRevisions()
    {
        string input = DocxTests.WriteFormattingRevisionProbeDocx();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            DocxMarkupMode = OoxPdfDocxMarkupMode.AllMarkup,
            DiagnosticSink = diagnostics.Add
        });

        TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_FORMATTING_REVISIONS" && d.Fallback == "Approximated"), "Markup modes should report formatting revisions as approximated while full formatting balloons are pending.");
        TestAssert.True(!diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_FORMATTING_REVISIONS"), "Markup modes should not emit the unsupported formatting-revision diagnostic.");
    }

    public static void DocxFormattingRevisionBalloonLabelsPrioritizeVisibleProperties()
    {
        var runRevision = new DocxRevisionInfo(
            DocxRevisionKind.RunPropertiesChange,
            "1",
            "Reviewer",
            "2026-06-10T00:00:00Z",
            "rPrChange",
            DocxRevisionPropertyFamily.Run,
            ["bCs", "highlight", "iCs", "rFonts", "szCs", "sz", "color"]);
        var paragraphRevision = new DocxRevisionInfo(
            DocxRevisionKind.ParagraphPropertiesChange,
            "2",
            "Reviewer",
            "2026-06-10T00:00:00Z",
            "pPrChange",
            DocxRevisionPropertyFamily.Paragraph,
            ["keepNext", "pStyle", "tabs", "spacing", "jc"]);

        TestAssert.Equal(
            "Formatted run: font, font size, color, highlight, +3 more",
            DocxRenderer.BuildRevisionBalloonPreview([runRevision]));
        TestAssert.Equal(
            "Formatted paragraph: paragraph style, alignment, spacing, tab stops, +1 more",
            DocxRenderer.BuildRevisionBalloonPreview([paragraphRevision]));
    }

    public static void DocxFormattingRevisionBalloonLabelsGroupSameFamilyProperties()
    {
        var paragraphStyleRevision = new DocxRevisionInfo(
            DocxRevisionKind.ParagraphPropertiesChange,
            "1",
            "Reviewer",
            "2026-06-10T00:00:00Z",
            "pPrChange",
            DocxRevisionPropertyFamily.Paragraph,
            ["pStyle", "spacing"]);
        var paragraphLayoutRevision = new DocxRevisionInfo(
            DocxRevisionKind.ParagraphPropertiesChange,
            "2",
            "Reviewer",
            "2026-06-10T00:00:00Z",
            "pPrChange",
            DocxRevisionPropertyFamily.Paragraph,
            ["tabs", "jc", "keepNext"]);
        var runFontRevision = new DocxRevisionInfo(
            DocxRevisionKind.RunPropertiesChange,
            "3",
            "Reviewer",
            "2026-06-10T00:00:00Z",
            "rPrChange",
            DocxRevisionPropertyFamily.Run,
            ["rFonts"]);
        var runColorRevision = new DocxRevisionInfo(
            DocxRevisionKind.RunPropertiesChange,
            "4",
            "Reviewer",
            "2026-06-10T00:00:00Z",
            "rPrChange",
            DocxRevisionPropertyFamily.Run,
            ["highlight", "color"]);

        TestAssert.Equal(
            "Formatted paragraph: paragraph style, alignment, spacing, tab stops, +1 more, Formatted run: font, color, highlight",
            DocxRenderer.BuildRevisionBalloonPreview([paragraphLayoutRevision, runColorRevision, paragraphStyleRevision, runFontRevision]));
    }

    public static void DocxFormattingRevisionBalloonLabelsOrderFamiliesByScopeSeverity()
    {
        var runRevision = new DocxRevisionInfo(DocxRevisionKind.RunPropertiesChange, "1", "Reviewer", "2026-06-10T00:00:00Z", "rPrChange", DocxRevisionPropertyFamily.Run, ["color"]);
        var paragraphRevision = new DocxRevisionInfo(DocxRevisionKind.ParagraphPropertiesChange, "2", "Reviewer", "2026-06-10T00:00:00Z", "pPrChange", DocxRevisionPropertyFamily.Paragraph, ["jc"]);
        var cellRevision = new DocxRevisionInfo(DocxRevisionKind.TableCellPropertiesChange, "3", "Reviewer", "2026-06-10T00:00:00Z", "tcPrChange", DocxRevisionPropertyFamily.Cell, ["tcW"]);
        var rowRevision = new DocxRevisionInfo(DocxRevisionKind.TableRowPropertiesChange, "4", "Reviewer", "2026-06-10T00:00:00Z", "trPrChange", DocxRevisionPropertyFamily.Row, ["trHeight"]);
        var tableRevision = new DocxRevisionInfo(DocxRevisionKind.TablePropertiesChange, "5", "Reviewer", "2026-06-10T00:00:00Z", "tblPrChange", DocxRevisionPropertyFamily.Table, ["tblW"]);
        var sectionRevision = new DocxRevisionInfo(DocxRevisionKind.SectionPropertiesChange, "6", "Reviewer", "2026-06-10T00:00:00Z", "sectPrChange", DocxRevisionPropertyFamily.Section, ["pgMar"]);

        TestAssert.Equal(
            "Formatted section: page margins, Formatted table: table width, Formatted row: row height, Formatted cell: cell width, Formatted paragraph: alignment, Formatted run: color",
            DocxRenderer.BuildRevisionBalloonPreview([runRevision, cellRevision, sectionRevision, paragraphRevision, tableRevision, rowRevision]));
    }

    public static void DocxFormattingRevisionBalloonLabelsNameCommonWordProperties()
    {
        var runRevision = new DocxRevisionInfo(
            DocxRevisionKind.RunPropertiesChange,
            "1",
            "Reviewer",
            "2026-06-10T00:00:00Z",
            "rPrChange",
            DocxRevisionPropertyFamily.Run,
            ["iCs", "webHidden", "outline", "rtl", "vertAlign", "kern"]);
        var paragraphRevision = new DocxRevisionInfo(
            DocxRevisionKind.ParagraphPropertiesChange,
            "2",
            "Reviewer",
            "2026-06-10T00:00:00Z",
            "pPrChange",
            DocxRevisionPropertyFamily.Paragraph,
            ["textAlignment", "outlineLvl", "pageBreakBefore", "pBdr", "bidi", "framePr", "wordWrap"]);
        var tableRevision = new DocxRevisionInfo(
            DocxRevisionKind.TablePropertiesChange,
            "3",
            "Reviewer",
            "2026-06-10T00:00:00Z",
            "tblPrChange",
            DocxRevisionPropertyFamily.Table,
            ["bidiVisual", "tblCellSpacing", "shd", "jc", "tblLook"]);
        var rowRevision = new DocxRevisionInfo(
            DocxRevisionKind.TableRowPropertiesChange,
            "4",
            "Reviewer",
            "2026-06-10T00:00:00Z",
            "trPrChange",
            DocxRevisionPropertyFamily.Row,
            ["tblCellSpacing", "jc", "cantSplit"]);
        var cellRevision = new DocxRevisionInfo(
            DocxRevisionKind.TableCellPropertiesChange,
            "5",
            "Reviewer",
            "2026-06-10T00:00:00Z",
            "tcPrChange",
            DocxRevisionPropertyFamily.Cell,
            ["noWrap", "textDirection", "shd", "tcFitText"]);
        var sectionRevision = new DocxRevisionInfo(
            DocxRevisionKind.SectionPropertiesChange,
            "6",
            "Reviewer",
            "2026-06-10T00:00:00Z",
            "sectPrChange",
            DocxRevisionPropertyFamily.Section,
            ["lnNumType", "footnotePr", "pgNumType", "docGrid", "endnotePr"]);

        TestAssert.Equal(
            "Formatted run: vertical position, right-to-left, kerning, outline, +2 more",
            DocxRenderer.BuildRevisionBalloonPreview([runRevision]));
        TestAssert.Equal(
            "Formatted paragraph: page break before, outline level, borders, word wrap, +3 more",
            DocxRenderer.BuildRevisionBalloonPreview([paragraphRevision]));
        TestAssert.Equal(
            "Formatted table: table look, alignment, cell spacing, shading, +1 more",
            DocxRenderer.BuildRevisionBalloonPreview([tableRevision]));
        TestAssert.Equal(
            "Formatted row: row split, alignment, cell spacing",
            DocxRenderer.BuildRevisionBalloonPreview([rowRevision]));
        TestAssert.Equal(
            "Formatted cell: shading, text direction, fit text, no wrap",
            DocxRenderer.BuildRevisionBalloonPreview([cellRevision]));
        TestAssert.Equal(
            "Formatted section: page numbering, document grid, line numbering, footnotes, +1 more",
            DocxRenderer.BuildRevisionBalloonPreview([sectionRevision]));
    }

    public static void DocxFormattingRevisionBalloonLabelsNameAdditionalWordProperties()
    {
        var runRevision = new DocxRevisionInfo(
            DocxRevisionKind.RunPropertiesChange,
            "1",
            "Reviewer",
            "2026-06-10T00:00:00Z",
            "rPrChange",
            DocxRevisionPropertyFamily.Run,
            ["oMath", "noProof", "snapToGrid", "spacing", "eastAsianLayout"]);
        var paragraphRevision = new DocxRevisionInfo(
            DocxRevisionKind.ParagraphPropertiesChange,
            "2",
            "Reviewer",
            "2026-06-10T00:00:00Z",
            "pPrChange",
            DocxRevisionPropertyFamily.Paragraph,
            ["mirrorIndents", "suppressAutoHyphens", "autoSpaceDE", "suppressOverlap"]);
        var tableRevision = new DocxRevisionInfo(
            DocxRevisionKind.TablePropertiesChange,
            "3",
            "Reviewer",
            "2026-06-10T00:00:00Z",
            "tblPrChange",
            DocxRevisionPropertyFamily.Table,
            ["tblpPr", "tblOverlap", "tblCaption", "tblStyleColBandSize"]);
        var rowRevision = new DocxRevisionInfo(
            DocxRevisionKind.TableRowPropertiesChange,
            "4",
            "Reviewer",
            "2026-06-10T00:00:00Z",
            "trPrChange",
            DocxRevisionPropertyFamily.Row,
            ["gridBefore", "gridAfter", "wBefore", "tblPrEx"]);
        var cellRevision = new DocxRevisionInfo(
            DocxRevisionKind.TableCellPropertiesChange,
            "5",
            "Reviewer",
            "2026-06-10T00:00:00Z",
            "tcPrChange",
            DocxRevisionPropertyFamily.Cell,
            ["hideMark", "hMerge", "cellIns", "cellMerge"]);
        var sectionRevision = new DocxRevisionInfo(
            DocxRevisionKind.SectionPropertiesChange,
            "6",
            "Reviewer",
            "2026-06-10T00:00:00Z",
            "sectPrChange",
            DocxRevisionPropertyFamily.Section,
            ["pgBorders", "titlePg", "vAlign", "rtlGutter"]);
        var fallbackRevision = new DocxRevisionInfo(
            DocxRevisionKind.ParagraphPropertiesChange,
            "7",
            "Reviewer",
            "2026-06-10T00:00:00Z",
            "pPrChange",
            DocxRevisionPropertyFamily.Paragraph,
            ["customXmlPr"]);
        var numberingRevision = new DocxRevisionInfo(
            DocxRevisionKind.ParagraphPropertiesChange,
            "8",
            "Reviewer",
            "2026-06-10T00:00:00Z",
            "pPrChange",
            DocxRevisionPropertyFamily.Paragraph,
            ["numberingChange", "numId", "ilvl", "numPr"]);

        TestAssert.Equal(
            "Formatted run: East Asian layout, character spacing, snap to grid, proofing, +1 more",
            DocxRenderer.BuildRevisionBalloonPreview([runRevision]));
        TestAssert.Equal(
            "Formatted paragraph: East Asian auto spacing, auto hyphenation, mirror indents, overlap suppression",
            DocxRenderer.BuildRevisionBalloonPreview([paragraphRevision]));
        TestAssert.Equal(
            "Formatted table: floating table position, table overlap, table caption, column band size",
            DocxRenderer.BuildRevisionBalloonPreview([tableRevision]));
        TestAssert.Equal(
            "Formatted row: grid before, grid after, width before, table property exceptions",
            DocxRenderer.BuildRevisionBalloonPreview([rowRevision]));
        TestAssert.Equal(
            "Formatted cell: end mark, horizontal merge, inserted cell, cell merge",
            DocxRenderer.BuildRevisionBalloonPreview([cellRevision]));
        TestAssert.Equal(
            "Formatted section: page borders, different first page, vertical alignment, right-to-left gutter",
            DocxRenderer.BuildRevisionBalloonPreview([sectionRevision]));
        TestAssert.Equal(
            "Formatted paragraph: custom xml properties",
            DocxRenderer.BuildRevisionBalloonPreview([fallbackRevision]));
        TestAssert.Equal(
            "Formatted paragraph: numbering, list level, numbering id, numbering change",
            DocxRenderer.BuildRevisionBalloonPreview([numberingRevision]));
    }

    public static void DocxMarkupInspectionSnapshotsExposeFormattingRevisionProvenance()
    {
        string input = DocxTests.WriteFormattingRevisionProbeDocx();
        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, markupMode: OoxPdfDocxMarkupMode.AllMarkup);
        DocxParagraph paragraph = document.Paragraphs.Single();

        TestAssert.True(paragraph.Revisions.Any(revision => revision.Kind == DocxRevisionKind.ParagraphPropertiesChange && revision.SourceElement == "pPrChange"), "Paragraph formatting revisions should retain private-safe provenance.");
        TestAssert.True(paragraph.Revisions.Any(revision => revision.Kind == DocxRevisionKind.RunPropertiesChange && revision.SourceElement == "rPrChange"), "Run formatting revisions should be lifted into paragraph-level revision counts.");
        TestAssert.True(paragraph.Runs.Single(run => run.Text == "Formatting revision").Revisions.Any(revision => revision.Kind == DocxRevisionKind.RunPropertiesChange), "Runs should retain their own formatting-revision provenance.");
        TestAssert.True(paragraph.Revisions.Any(revision => revision.PropertyChangeFamily == DocxRevisionPropertyFamily.Paragraph && revision.PropertyElementNames.Contains("jc")), "Paragraph formatting revisions should expose private-safe changed property names.");
        TestAssert.True(paragraph.Revisions.Any(revision => revision.PropertyChangeFamily == DocxRevisionPropertyFamily.Run && revision.PropertyElementNames.Contains("b")), "Run formatting revisions should expose private-safe changed property names.");
        string revisionPreview = DocxRenderer.BuildRevisionBalloonPreview(paragraph.Revisions);
        TestAssert.Contains("Formatted paragraph: alignment", revisionPreview);
        TestAssert.Contains("Formatted run: color, bold", revisionPreview);

        DocxStructureSnapshot structure = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectStructure(document);
        TestAssert.Equal(3, structure.RevisionCount);
        TestAssert.Equal(3, structure.OtherRevisionCount);
        TestAssert.Equal(3, structure.FormattingRevisionCount);
        TestAssert.Equal(1, structure.RunFormattingRevisionCount);
        TestAssert.Equal(1, structure.ParagraphFormattingRevisionCount);
        TestAssert.Equal(1, structure.SectionFormattingRevisionCount);
        TestAssert.True((structure.FormattingRevisionProperties ?? []).Any(property => property.Family == "Paragraph" && property.PropertyElementName == "jc" && property.Count == 1), "Formatting revision property snapshots should count paragraph property names without text.");
        TestAssert.True((structure.FormattingRevisionProperties ?? []).Any(property => property.Family == "Section" && property.PropertyElementName == "pgMar" && property.Count == 1), "Formatting revision property snapshots should count section property names without text.");
        TestAssert.Equal(2, structure.Blocks.Single(block => block.Kind == "Paragraph").OtherRevisionCount);
    }

    public static void DocxMarkupInspectionSnapshotsCountStoryFormattingRevisions()
    {
        var bodyRevision = new DocxRevisionInfo(
            DocxRevisionKind.ParagraphPropertiesChange,
            "1",
            "Reviewer",
            "2026-06-10T00:00:00Z",
            "pPrChange",
            DocxRevisionPropertyFamily.Paragraph,
            ["jc"]);
        var headerRevision = new DocxRevisionInfo(
            DocxRevisionKind.RunPropertiesChange,
            "2",
            "Reviewer",
            "2026-06-10T00:00:00Z",
            "rPrChange",
            DocxRevisionPropertyFamily.Run,
            ["highlight"]);
        var textBoxRevision = new DocxRevisionInfo(
            DocxRevisionKind.RunPropertiesChange,
            "3",
            "Reviewer",
            "2026-06-10T00:00:00Z",
            "rPrChange",
            DocxRevisionPropertyFamily.Run,
            ["b"]);
        DocxParagraph bodyParagraph = DocxTests.CreateDocxLayoutParagraph("Body", 10d, 12d) with
        {
            Revisions = [bodyRevision]
        };
        DocxParagraph headerParagraph = DocxTests.CreateDocxLayoutParagraph("Header", 10d, 12d) with
        {
            Revisions = [headerRevision]
        };
        DocxParagraph textBoxParagraph = DocxTests.CreateDocxLayoutParagraph("Text box", 10d, 12d) with
        {
            Revisions = [textBoxRevision]
        };
        DocxPageSettings pageSettings = DocxPageSettings.Empty with
        {
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [new DocxParagraphElement(headerParagraph)]
            }
        };
        DocxDocument document = new(
            300d,
            300d,
            30d,
            90d,
            30d,
            30d,
            pageSettings,
            [DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(textBoxParagraph)])],
            [],
            [],
            [new DocxParagraphElement(bodyParagraph)],
            [],
            [])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        DocxStructureSnapshot structure = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectStructure(document);

        TestAssert.Equal(3, structure.FormattingRevisionCount);
        TestAssert.Equal(2, structure.RunFormattingRevisionCount);
        TestAssert.Equal(1, structure.ParagraphFormattingRevisionCount);
        TestAssert.True((structure.FormattingRevisionProperties ?? []).Any(property => property.Family == "Run" && property.PropertyElementName == "highlight" && property.Count == 1), "Static-story formatting revision properties should be counted without text.");
        TestAssert.True((structure.FormattingRevisionProperties ?? []).Any(property => property.Family == "Run" && property.PropertyElementName == "b" && property.Count == 1), "Floating text-box formatting revision properties should be counted without text.");
    }

    public static void DocxSupportedBodyKeepRulesDoNotEmitUnsupportedKeepDiagnostic()
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
                      <w:pPr>
                        <w:keepNext/>
                        <w:keepLines/>
                        <w:widowControl/>
                      </w:pPr>
                      <w:r><w:t>Kept body paragraph</w:t></w:r>
                    </w:p>
                    <w:p><w:r><w:t>Following paragraph</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(!diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_PARAGRAPH_KEEP_RULE"), "Body paragraph keep/widow rules are parsed and consumed by page layout, so they should not emit stale unsupported diagnostics.");
    }

    public static void DocxReaderPreservesFootnoteStoryTypesWithoutPlacingSeparators()
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
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:t>Body before</w:t></w:r>
                      <w:r><w:footnoteReference w:id="0"/></w:r>
                      <w:r><w:t> body middle </w:t></w:r>
                      <w:r><w:footnoteReference w:id="2"/></w:r>
                      <w:r><w:t> body after</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
            ["word/footnotes.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:footnote w:type="separator" w:id="-1">
                    <w:p><w:r><w:t>Separator body</w:t></w:r></w:p>
                  </w:footnote>
                  <w:footnote w:type="continuationSeparator" w:id="0">
                    <w:p><w:r><w:t>Continuation body</w:t></w:r></w:p>
                  </w:footnote>
                  <w:footnote w:id="2">
                    <w:p><w:r><w:t>Normal footnote body</w:t></w:r></w:p>
                  </w:footnote>
                </w:footnotes>
                """
        });

        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        TestAssert.Equal(3, document.RelatedStories.Count);
        TestAssert.Equal("separator", document.RelatedStories.Single(story => story.Id == "-1").Type?.ToValueString() ?? string.Empty);
        TestAssert.Equal("continuationSeparator", document.RelatedStories.Single(story => story.Id == "0").Type?.ToValueString() ?? string.Empty);
        TestAssert.True(document.RelatedStories.Single(story => story.Id == "2").Type is null, "Normal note bodies without w:type should remain normal rather than receiving an inferred type token.");

        DocxLayoutSnapshot layoutSnapshot = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectLayout(document);
        TestAssert.Equal("separator", layoutSnapshot.RelatedStories.Single(story => story.Id == "-1").Type ?? string.Empty);
        TestAssert.Equal("continuationSeparator", layoutSnapshot.RelatedStories.Single(story => story.Id == "0").Type ?? string.Empty);
        DocxLayoutPageSnapshot footnotePage = layoutSnapshot.Pages.Single(page => page.PlacedFootnoteStoryCount == 1);
        DocxPlacedRelatedStoryLayoutSnapshot[] placedStories = footnotePage.PlacedRelatedStories.ToArray();
        TestAssert.Equal(2, placedStories.Length);
        DocxPlacedRelatedStoryLayoutSnapshot separatorStory = placedStories.Single(story => story.Type == "separator");
        DocxPlacedRelatedStoryLayoutSnapshot placedStory = placedStories.Single(story => story.Id == "2");
        TestAssert.True(separatorStory.SourceBlockIndex == 0 && separatorStory.TopY > placedStory.TopY && separatorStory.SeparatorY is null, "Structural footnote separator stories should be placed above normal note bodies without drawing the generic separator rectangle.");
        TestAssert.True(placedStory.Type is null && placedStory.SeparatorY is null, "Only normal footnote stories should count as placed note bodies once separator stories are structural.");
        TestAssert.True(!placedStories.Any(story => story.Type == "continuationSeparator"), "Continuation separator stories should remain unplaced until actual multi-page note continuation is modeled.");
    }

    public static void DocxUnsupportedStoryDiagnosticsPreferRelatedPartNames()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
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
                  <Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
                  <Relationship Id="rIdFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
                  <Relationship Id="rIdEndnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
                </Relationships>
                """,
            ["word/_rels/comments.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdCommentLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/comment" TargetMode="External"/>
                  <Relationship Id="rIdCommentImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/comment.png"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:commentRangeStart w:id="1"/>
                      <w:r><w:t>Referenced story bodies</w:t></w:r>
                      <w:r><w:commentReference w:id="1"/></w:r>
                      <w:r><w:t>Before</w:t><w:footnoteReference w:id="2" w:customMarkFollows="1"/><w:t>After</w:t></w:r>
                      <w:r><w:endnoteReference w:id="3"/></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
            ["word/comments.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:comment w:id="1">
                    <w:p>
                      <w:r><w:t>Comment body</w:t></w:r>
                      <w:r>
                        <w:drawing>
                          <wp:anchor distT="0" distB="0" distL="0" distR="0" behindDoc="1">
                            <wp:extent cx="914400" cy="457200"/>
                            <wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionH>
                            <wp:positionV relativeFrom="paragraph"><wp:posOffset>228600</wp:posOffset></wp:positionV>
                            <wp:wrapNone/>
                            <a:graphic>
                              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:pic><pic:blipFill><a:blip r:embed="rIdCommentImage"/></pic:blipFill></pic:pic>
                              </a:graphicData>
                            </a:graphic>
                          </wp:anchor>
                        </w:drawing>
                      </w:r>
                    </w:p>
                    <w:p><w:hyperlink r:id="rIdCommentLink"><w:r><w:t>Comment link</w:t></w:r></w:hyperlink></w:p>
                    <w:tbl><w:tr><w:tc><w:p><w:r><w:t>Comment table</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
                  </w:comment>
                </w:comments>
                """,
            ["word/footnotes.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:footnote w:id="2"><w:p><w:r><w:t>Footnote body</w:t></w:r></w:p></w:footnote>
                </w:footnotes>
                """,
            ["word/endnotes.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:endnote w:id="3"><w:p><w:r><w:t>Endnote body</w:t></w:r></w:p></w:endnote>
                </w:endnotes>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_COMMENTS" && d.PartName == "/word/comments.xml"), "Comments diagnostics should point to comments.xml when the story body exists.");
        TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_FOOTNOTE" && d.PartName == "/word/footnotes.xml" && d.Fallback == "Approximated"), "Footnote diagnostics should point to footnotes.xml and report approximation when the story body exists.");
        TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_ENDNOTE" && d.PartName == "/word/endnotes.xml" && d.Fallback == "Approximated"), "Endnote diagnostics should point to endnotes.xml and report approximation when the story body exists.");
        TestAssert.True(!diagnostics.Any(d =>
            (d.Id == "DOCX_UNSUPPORTED_COMMENTS" || d.Id == "DOCX_APPROXIMATED_FOOTNOTE" || d.Id == "DOCX_APPROXIMATED_ENDNOTE") &&
            d.PartName == "/word/document.xml"), "Story-body diagnostics should not be flattened to document.xml when the related body part exists.");

        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxParagraph referenceParagraph = document.Paragraphs.Single();
        TestAssert.Equal(3, referenceParagraph.InlineReferences.Count);
        DocxInlineReference commentReference = referenceParagraph.InlineReferences.Single(reference => reference.Kind == DocxRelatedStoryKind.Comment);
        DocxInlineReference footnoteReference = referenceParagraph.InlineReferences.Single(reference => reference.Kind == DocxRelatedStoryKind.Footnote);
        DocxInlineReference endnoteReference = referenceParagraph.InlineReferences.Single(reference => reference.Kind == DocxRelatedStoryKind.Endnote);
        TestAssert.True(commentReference.Id == "1" && commentReference.CustomMarkFollowsValue is null, "Comment reference markers should be preserved as inline DOCX structure.");
        TestAssert.True(footnoteReference.Id == "2" && footnoteReference.CustomMarkFollowsValue == "1", "Footnote reference markers should preserve custom mark flags.");
        TestAssert.True(endnoteReference.Id == "3" && endnoteReference.CustomMarkFollowsValue is null, "Endnote reference markers should be preserved as inline DOCX structure.");
        TestAssert.Equal(1, commentReference.SourceRunIndex);
        TestAssert.Equal(0, commentReference.RunChildIndex);
        TestAssert.Equal(0, commentReference.TextOffsetInRun);
        DocxCommentRange commentRange = referenceParagraph.CommentRanges.Single();
        TestAssert.True(commentRange.Id == "1" && commentRange.StartSourceRunIndex == 0 && commentRange.StartTextOffset == 0 && commentRange.ReferenceSourceRunIndex == 1 && commentRange.ReferenceTextOffset == 0, "Comment range starts and marker anchors should be preserved without exposing body text.");
        TestAssert.Equal(2, footnoteReference.SourceRunIndex);
        TestAssert.Equal(1, footnoteReference.RunChildIndex);
        TestAssert.Equal(6, footnoteReference.TextOffsetInRun);
        TestAssert.Equal(3, endnoteReference.SourceRunIndex);
        TestAssert.Equal(0, endnoteReference.RunChildIndex);
        TestAssert.Equal(0, endnoteReference.TextOffsetInRun);
        TestAssert.Equal(3, document.RelatedStories.Count);
        DocxRelatedStory commentStory = document.RelatedStories.Single(story => story.Kind == DocxRelatedStoryKind.Comment);
        TestAssert.True(commentStory.PartName == "/word/comments.xml" && commentStory.Id == "1" && commentStory.BodyElements.Count == 3 && commentStory.Paragraphs.Count == 2 && commentStory.Tables.Count == 1, "Comment bodies should be preserved as related DOCX stories.");
        TestAssert.Equal(1, commentStory.Paragraphs.Sum(paragraph => paragraph.Hyperlinks.Count));
        TestAssert.Equal("https://example.invalid/comment", commentStory.Paragraphs.SelectMany(paragraph => paragraph.Hyperlinks).Single().Target ?? string.Empty);
        DocxFloatingDrawing commentDrawing = commentStory.FloatingDrawings.Single();
        TestAssert.True(commentDrawing.ImageRelationshipId == "rIdCommentImage" && commentDrawing.SourceParagraphIndex == 0 && commentDrawing.SourceBlockIndex == 0, "Related-story anchored drawings should preserve their owning paragraph, block, and part-local image relationship.");
        TestAssert.True(commentDrawing.HorizontalRelativeFromValue == "page" && commentDrawing.VerticalRelativeFromValue == "paragraph" && commentDrawing.BehindDocumentValue == "1", "Related-story anchored drawing geometry tokens should be preserved structurally before story placement is modeled.");
        TestAssert.True(document.RelatedStories.Any(story => story.Kind == DocxRelatedStoryKind.Footnote && story.PartName == "/word/footnotes.xml" && story.Id == "2" && story.Paragraphs.Count == 1), "Footnote bodies should be preserved as related DOCX stories.");
        TestAssert.True(document.RelatedStories.Any(story => story.Kind == DocxRelatedStoryKind.Endnote && story.PartName == "/word/endnotes.xml" && story.Id == "3" && story.Paragraphs.Count == 1), "Endnote bodies should be preserved as related DOCX stories.");

        DocxStructureSnapshot snapshot = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectStructure(document);
        DocxStructureBlockSnapshot referenceBlock = snapshot.Blocks.Single(block => block.Kind == "Paragraph");
        TestAssert.Equal(3, snapshot.InlineReferenceCount);
        TestAssert.Equal(3, snapshot.AnchoredInlineReferenceCount);
        TestAssert.Equal(3, snapshot.ResolvedInlineReferenceCount);
        TestAssert.Equal(6, snapshot.MaxInlineReferenceTextOffsetInRun);
        TestAssert.Equal(3, referenceBlock.InlineReferenceCount);
        TestAssert.Equal(3, referenceBlock.AnchoredInlineReferenceCount);
        TestAssert.Equal(3, referenceBlock.ResolvedInlineReferenceCount);
        TestAssert.Equal(6, referenceBlock.MaxInlineReferenceTextOffsetInRun);
        TestAssert.Equal(1, referenceBlock.CommentReferenceCount);
        TestAssert.Equal(1, referenceBlock.FootnoteReferenceCount);
        TestAssert.Equal(1, referenceBlock.EndnoteReferenceCount);
        TestAssert.True(snapshot.Stories.Any(story => story.Kind == "Body" && story.InlineReferenceCount == 3 && story.ResolvedInlineReferenceCount == 3), "Structure snapshots should expose body inline story-reference ownership.");
        TestAssert.True(snapshot.Stories.Any(story => story.Kind == "Comment" && story.Scope == "/word/comments.xml" && story.VariantType == "1" && story.BlockCount == 3 && story.ParagraphCount == 2 && story.TableCount == 1 && story.TextLength == 37 && story.HyperlinkCount == 1 && story.ExternalHyperlinkCount == 1 && story.FloatingDrawingCount == 1), "Structure snapshots should expose comment story ownership, hyperlinks, anchored drawings, and table metrics.");
        TestAssert.True(snapshot.Stories.Any(story => story.Kind == "Footnote" && story.Scope == "/word/footnotes.xml" && story.VariantType == "2" && story.TextLength == 13), "Structure snapshots should expose footnote story text metrics.");
        TestAssert.True(snapshot.Stories.Any(story => story.Kind == "Endnote" && story.Scope == "/word/endnotes.xml" && story.VariantType == "3" && story.TextLength == 12), "Structure snapshots should expose endnote story text metrics.");
        TestAssert.Equal(3, snapshot.InlineReferences.Count);
        DocxStructureCommentRangeSnapshot commentRangeSnapshot = snapshot.CommentRanges!.Single();
        TestAssert.True(commentRangeSnapshot.SourceBlockIndex == 0 && commentRangeSnapshot.SourceBlockKind == "Paragraph" && commentRangeSnapshot.Id == "1" && commentRangeSnapshot.StartSourceRunIndex == 0 && commentRangeSnapshot.ReferenceSourceRunIndex == 1, "Comment range snapshots should preserve private-safe source coordinates.");
        TestAssert.True(commentRangeSnapshot.ResolvedStoryPartName == "/word/comments.xml" && commentRangeSnapshot.ResolvedStoryId == "1" && commentRangeSnapshot.ResolvedStoryBlockCount == 3 && commentRangeSnapshot.ResolvedStoryTextLength == 37, "Comment range snapshots should resolve to the comment story body.");
        DocxStructureInlineReferenceSnapshot commentReferenceSnapshot = snapshot.InlineReferences.Single(reference => reference.Kind == "Comment");
        TestAssert.True(commentReferenceSnapshot.SourceBlockIndex == 0 && commentReferenceSnapshot.SourceBlockKind == "Paragraph" && commentReferenceSnapshot.SourceRunIndex == 1 && commentReferenceSnapshot.TextOffsetInRun == 0, "Comment reference snapshot should preserve the source marker coordinates.");
        TestAssert.True(commentReferenceSnapshot.ResolvedStoryKind == "Comment" && commentReferenceSnapshot.ResolvedStoryPartName == "/word/comments.xml" && commentReferenceSnapshot.ResolvedStoryId == "1" && commentReferenceSnapshot.ResolvedStoryBlockCount == 3 && commentReferenceSnapshot.ResolvedStoryTextLength == 37, "Comment reference snapshot should resolve to the comment story body.");
        DocxStructureInlineReferenceSnapshot footnoteReferenceSnapshot = snapshot.InlineReferences.Single(reference => reference.Kind == "Footnote");
        TestAssert.True(footnoteReferenceSnapshot.CustomMarkFollowsValue == "1" && footnoteReferenceSnapshot.SourceRunIndex == 2 && footnoteReferenceSnapshot.TextOffsetInRun == 6, "Footnote reference snapshot should preserve custom mark and source marker offsets.");
        TestAssert.True(footnoteReferenceSnapshot.ResolvedStoryKind == "Footnote" && footnoteReferenceSnapshot.ResolvedStoryPartName == "/word/footnotes.xml" && footnoteReferenceSnapshot.ResolvedStoryId == "2" && footnoteReferenceSnapshot.ResolvedStoryTextLength == 13, "Footnote reference snapshot should resolve to the footnote story body.");
        DocxStructureInlineReferenceSnapshot endnoteReferenceSnapshot = snapshot.InlineReferences.Single(reference => reference.Kind == "Endnote");
        TestAssert.True(endnoteReferenceSnapshot.SourceRunIndex == 3 && endnoteReferenceSnapshot.TextOffsetInRun == 0, "Endnote reference snapshot should preserve source marker offsets.");
        TestAssert.True(endnoteReferenceSnapshot.ResolvedStoryKind == "Endnote" && endnoteReferenceSnapshot.ResolvedStoryPartName == "/word/endnotes.xml" && endnoteReferenceSnapshot.ResolvedStoryId == "3" && endnoteReferenceSnapshot.ResolvedStoryTextLength == 12, "Endnote reference snapshot should resolve to the endnote story body.");

        DocxLayoutSnapshot layoutSnapshot = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectLayout(document);
        TestAssert.Equal(3, layoutSnapshot.RelatedStories.Count);
        DocxRelatedStoryLayoutSnapshot commentLayout = layoutSnapshot.RelatedStories.Single(story => story.Kind == "Comment");
        TestAssert.True(commentLayout.PartName == "/word/comments.xml" && commentLayout.Id == "1" && commentLayout.BlockCount == 3 && commentLayout.ParagraphCount == 2 && commentLayout.TableCount == 1, "Related-story layout snapshots should preserve comment story ownership without flattening it into body layout.");
        TestAssert.True(commentLayout.TextLineCount >= 2 && commentLayout.TableCellTextLineCount >= 1 && commentLayout.TableRowCount == 1 && commentLayout.FloatingDrawingCount == 1 && commentLayout.TextLength == 37 && commentLayout.ContentHeight > 0d, "Comment story layout should measure paragraph text and table rows while preserving unpaged anchored drawing ownership.");
        TestAssert.True(commentLayout.Items.Count(item => item.Kind == "TextLine") >= 2 && commentLayout.Items.Count(item => item.Kind == "TableRow") == 1 && commentLayout.TableRows.Count == 1, "Related-story snapshots should expose private-safe item and table-row ownership for future story placement.");
        TestAssert.True(commentLayout.SourceBlocks.Count == 3 && commentLayout.SourceBlocks.Any(block => block.Kind == "Table" && block.TableRowCount == 1) && commentLayout.SourceBlocks.Count(block => block.Kind == "Paragraph") == 2, "Related-story snapshots should expose private-safe source-block summaries without assigning fake page indexes.");
        TestAssert.True(layoutSnapshot.RelatedStories.Any(story => story.Kind == "Footnote" && story.PartName == "/word/footnotes.xml" && story.Id == "2" && story.TextLineCount >= 1 && story.TableRowCount == 0 && story.ContentHeight > 0d), "Footnote story layout should be measured as related-story content.");
        TestAssert.True(layoutSnapshot.RelatedStories.Any(story => story.Kind == "Endnote" && story.PartName == "/word/endnotes.xml" && story.Id == "3" && story.TextLineCount >= 1 && story.TableRowCount == 0 && story.ContentHeight > 0d), "Endnote story layout should be measured as related-story content.");
        TestAssert.True(layoutSnapshot.Pages.Sum(page => page.PlacedRelatedStoryCount) == 2 && layoutSnapshot.Pages.Sum(page => page.PlacedFootnoteStoryCount) == 1 && layoutSnapshot.Pages.Sum(page => page.PlacedEndnoteStoryCount) == 1, "Resolved footnote story bodies should be placed on the marker page while endnotes become document-end page-owned story content.");
        DocxLayoutPageSnapshot footnotePage = layoutSnapshot.Pages.Single(page => page.PlacedFootnoteStoryCount == 1);
        DocxLayoutPageSnapshot endnotePage = layoutSnapshot.Pages.Single(page => page.PlacedEndnoteStoryCount == 1);
        TestAssert.True(footnotePage.PlacedRelatedStories.Any(story => story.Kind == "Footnote" && story.SourceBlockIndex == 0) && endnotePage.PlacedRelatedStories.Any(story => story.Kind == "Endnote" && story.SourceBlockIndex == -1), "Placed related-story snapshots should expose marker-owned footnotes and document-end endnotes without flattening them into body items.");
        TestAssert.True(footnotePage.PlacedRelatedItems.Any(item => item.Kind == "PlacedTextLine") && endnotePage.PlacedRelatedItems.Any(item => item.Kind == "PlacedTextLine"), "Placed related-story item snapshots should still expose page-owned note text lines for PDF-flow diagnostics.");
        TestAssert.True(layoutSnapshot.Pages.SelectMany(page => page.PlacedRelatedItems.Select(item => (Page: page, Item: item))).All(pair => pair.Item.Y >= pair.Page.MarginBottom), "Placed related-story items should be shifted from the unpaged story canvas into the page note area.");

        DocxFontPlan fontPlan = DocxFontPlan.Create(document, new MapFontResolver([], "Fallback"), CancellationToken.None);
        TestAssert.True(fontPlan.Runs.Any(run => run.Run.Text == "Comment body"), "Related story runs should participate in DOCX font planning.");
        TestAssert.True(fontPlan.Runs.Any(run => run.Run.Text == "Comment link"), "Related story hyperlink runs should participate in DOCX font planning.");
        TestAssert.True(fontPlan.Runs.Any(run => run.Run.Text == "Comment table"), "Related story table runs should participate in DOCX font planning.");
        TestAssert.True(fontPlan.Runs.Any(run => run.Run.Text == "Footnote body"), "Footnote runs should participate in DOCX font planning.");
        TestAssert.True(fontPlan.Runs.Any(run => run.Run.Text == "Endnote body"), "Endnote runs should participate in DOCX font planning.");
    }

    public static void DocxLayoutWrapsPlacedFootnoteUsingOwningSectionWidth()
    {
        DocxPageSettings wideFirstSection = new(
            "12240",
            "12000",
            null,
            "1440",
            "1440",
            "1440",
            "1440",
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null);
        DocxPageSettings narrowFinalSection = new(
            "4400",
            "6000",
            null,
            "200",
            "200",
            "200",
            "200",
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null);
        DocxParagraph anchor = DocxTests.CreateDocxLayoutParagraph("Narrow note marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "23",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 7)
            ]
        };
        DocxParagraph footnoteParagraph = DocxTests.CreateDocxLayoutParagraph("alpha beta gamma delta epsilon zeta eta theta", 10d, 12d);
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "23",
            [new DocxParagraphElement(footnoteParagraph)],
            [],
            [], null);
        var document = new DocxDocument(
            612d,
            792d,
            72d,
            72d,
            72d,
            72d,
            narrowFinalSection,
            [],
            [],
            [],
            [
                new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Wide first section", 10d, 12d)),
                new DocxSectionBreakElement(wideFirstSection, DocxSectionBreakType.NextPage, null, null, null, []),
                new DocxParagraphElement(anchor)
            ],
            [],
            [])
        {
            RelatedStories = [footnoteStory]
        };

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxPlacedRelatedStoryLayout placedStory = layout.Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .Single(story => story.StoryLayout.Story.Id == "23");
        DocxTextLineLayout[] lines = placedStory.TextLines.ToArray();

        TestAssert.True(Math.Abs(placedStory.Width - 200d) < 0.001d, "The placed footnote should inherit the narrow final section body width.");
        TestAssert.True(lines.Length >= 2, "The placed footnote should wrap using the owning section width, not the wider first section width.");
        TestAssert.True(lines.All(line => line.Width <= placedStory.Width + 0.001d), "Placed footnote lines should not exceed the owning section body width.");
    }

    public static void DocxLayoutSplitsSingleSectEndEndnoteStoryBeforeFollowingSection()
    {
        DocxParagraph firstSectionParagraph = DocxTests.CreateDocxLayoutParagraph("first section long endnote marker", 10d, 18d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Endnote,
                    "24",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 6)
            ]
        };
        DocxParagraph secondSectionParagraph = DocxTests.CreateDocxLayoutParagraph("second section body", 10d, 12d);
        DocxParagraph endnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Endnote split body", 10d, 18d);
        var endnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Endnote,
            "/word/endnotes.xml",
            "24",
            Enumerable.Range(0, 8).Select(_ => new DocxParagraphElement(endnoteParagraph)).Cast<DocxBodyElement>().ToArray(),
            [],
            [], null);
        DocxPageSettings firstSectionSettings = DocxPageSettings.Empty with
        {
            EndnoteReferenceSettings = DocxNoteReferenceSettings.Empty with { PositionValue = "sectEnd" }
        };
        var document = new DocxDocument(
            220d,
            112d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [
                new DocxParagraphElement(firstSectionParagraph),
                new DocxSectionBreakElement(firstSectionSettings, DocxSectionBreakType.NextPage, null, null, null, []),
                new DocxParagraphElement(secondSectionParagraph)
            ],
            [firstSectionParagraph, secondSectionParagraph],
            [])
        {
            RelatedStories = [endnoteStory]
        };

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None));
        int secondSectionPageIndex = snapshot.Pages
            .Select((page, pageIndex) => (page, pageIndex))
            .First(item => item.page.Items.Any(pageItem => pageItem.SourceBlockIndex == 2))
            .pageIndex;
        var placedSlices = snapshot.Pages
            .SelectMany((page, pageIndex) => page.PlacedRelatedStories
                .Where(story => story.Kind == "Endnote" && story.Id == "24")
                .Select(story => (story, pageIndex)))
            .ToArray();

        TestAssert.True(placedSlices.Length >= 2, "A single overlong sectEnd endnote story should be split into multiple placed slices instead of clipping away the tail.");
        TestAssert.True(placedSlices.All(slice => slice.pageIndex < secondSectionPageIndex), "All section-end endnote slices must stay before the following section page.");
        TestAssert.Equal(0d, placedSlices[0].story.ContentTopOffset);
        TestAssert.True(placedSlices[^1].story.ContentTopOffset > 0d, "Continuation slices should carry a nonzero source-story top offset.");
        TestAssert.True(placedSlices[^1].story.ContentHeight > placedSlices[^1].story.Height, "Continuation slices should retain the full unpaged story height for diagnostics.");
    }

    public static void DocxTextEmissionSplitsUncoveredGlyphsAcrossFallbackFonts()
    {
        var resolver = new WindowsFontResolver();
        FontFaceResolution symbolResolution = resolver.Resolve(new FontRequest("Symbol"));
        if (symbolResolution.IsFallback)
        {
            return;
        }

        OpenTypeFont? symbolFont = FontProgramLoader.Load(symbolResolution, CancellationToken.None);
        if (symbolFont is null || symbolFont.MapCodePoint(0xF0B7) == 0)
        {
            return;
        }

        string body = """
            <?xml version="1.0" encoding="UTF-8"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">a[BULLET]b</w:t></w:r></w:p>
                <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
              </w:body>
            </w:document>
            """.Replace("[BULLET]", ((char)0xF0B7).ToString());
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
            ["word/document.xml"] = body
        });
        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, markupMode: OoxPdfDocxMarkupMode.Final);
        DocxFontPlan plan = DocxFontPlan.Create(document, resolver, CancellationToken.None);
        DocxResolvedRunTypeface resolved = plan.Runs.Single(run => run.Run.Text.IndexOf((char)0xF0B7) >= 0);
        OpenTypeFont? primaryFont = resolved.Resolution is FontFaceResolution primaryResolution ? FontProgramLoader.Load(primaryResolution, CancellationToken.None) : null;
        if (primaryFont is null || primaryFont.MapCodePoint(0xF0B7) != 0)
        {
            return;
        }

        var renderer = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        DocxTextEmissionSnapshot emission = renderer.InspectTextEmission(document);
        DocxTextEmissionLineSnapshot line = emission.Lines.Single();
        DocxTextEmissionSegmentSnapshot[] segments = line.Segments.Where(snapshot => !snapshot.IsTerminalLineSpace).ToArray();
        TestAssert.Equal(3, segments.Length);
        TestAssert.Equal(3, segments.Sum(snapshot => snapshot.TextLength));
        TestAssert.Equal(2, segments.Select(snapshot => snapshot.FontResourceName).Distinct().Count());
    }
}
