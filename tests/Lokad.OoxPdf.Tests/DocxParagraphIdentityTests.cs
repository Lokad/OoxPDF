using System.Runtime.CompilerServices;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxParagraphIdentityTests
{
    public static void DocxParagraphIdentityRetainsCollidingRevisionOwners()
    {
        var (first,second)=FindCollision(Revision("First paragraph", "First reviewer"), Revision("Second paragraph", "Second reviewer"));
        TestAssert.Equal(RuntimeHelpers.GetHashCode(first),RuntimeHelpers.GetHashCode(second));
        var balloons=Renderer().InspectMarkupBalloons(Document(first,second));
        TestAssert.Equal(2,balloons.Count);
        TestAssert.Equal(2,balloons.Sum(balloon=>balloon.RevisionCandidateCount));
    }

    public static void DocxParagraphIdentityDistinguishesEqualValuesAndRepeatedOwners()
    {
        DocxParagraph first=Revision("Identical paragraph", "Reviewer");
        DocxParagraph second=first with {};
        TestAssert.True(first.Equals(second) && !ReferenceEquals(first,second),"Equal values are still distinct source owners.");
        TestAssert.Equal(2,Renderer().InspectMarkupBalloons(Document(first,second)).Count);
        // The same source owner appears on two lines after wrapping, and must
        // continue to contribute only one balloon on this page.
        var document=Document(first,second);
        first=first with { Runs=[first.Runs[0] with {Text="A long revised paragraph that wraps across several lines in the narrow body column."}] };
        document=document with { BodyElements=[new DocxParagraphElement(first)], MarginRightPoints=400 };
        TestAssert.True(Renderer().InspectLayout(document).Pages.Single().Items.Count(item=>item.Kind=="TextLine")>1,"Control must actually wrap.");
        TestAssert.Equal(1,Renderer().InspectMarkupBalloons(document).Count);
    }

    public static void DocxParagraphIdentityKeepsCommentIdsWithinTheirOwners()
    {
        var (first,second)=FindCollision(Comment("First paragraph"),Comment("Second paragraph"));
        var balloons=Renderer().InspectMarkupBalloons(Document(first,second));
        TestAssert.Equal(2,balloons.Count);
        TestAssert.Equal(2,balloons.Sum(balloon=>balloon.CommentCandidateCount));
        // Repeated references to the same comment on the same owner coalesce.
        first=first with {InlineReferences=[first.InlineReferences[0],first.InlineReferences[0]]};
        TestAssert.Equal(1,Renderer().InspectMarkupBalloons(Document(first)).Count);
        first=first with {InlineReferences=[first.InlineReferences[0],first.InlineReferences[0] with {Id="2"}]};
        TestAssert.Equal(2,Renderer().InspectMarkupBalloons(Document(first)).Sum(balloon=>balloon.CommentCandidateCount));
    }

    public static void DocxParagraphIdentityDoesNotConfuseBodyAndTextboxComments()
    {
        var (textbox,body)=FindCollision(Comment("Textbox comment"),Comment("Body comment"));
        var drawing=DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(textbox)]);
        var document=Document(body) with {FloatingDrawings=[drawing]};
        var renderer=Renderer(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        var balloons=renderer.InspectMarkupBalloons(document);
        TestAssert.Equal(1,balloons.Sum(balloon=>balloon.CommentCandidateCount));
        TestAssert.Equal(1,renderer.InspectLayout(document).FloatingDrawings.Count);
    }

    private static DocxRenderer Renderer(OoxPdfDocxMarkupGeometryMode geometry=OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout) =>
        new(new PortableRenderingTests.PinnedResolver(),OoxPdfDocxMarkupMode.AllMarkup,geometry);

    private static DocxParagraph Revision(string text,string author) => Paragraph(text) with {
        Revisions=[new DocxRevisionInfo(DocxRevisionKind.RunPropertiesChange,"1",author,"2026-10-20T12:00:00Z","rPrChange",null,["b"])]
    };
    private static DocxParagraph Comment(string text) => Paragraph(text) with {
        InlineReferences=[new DocxInlineReference(DocxRelatedStoryKind.Comment,"1",CustomMarkFollowsValue:null,DisplayText:null,SourceRunIndex:0,RunChildIndex:0,TextOffsetInRun:0)]
    };
    private static DocxParagraph Paragraph(string text)
    {
        DocxParagraph paragraph=DocxTests.CreateDocxLayoutParagraph(text,12,20);
        return paragraph with {Runs=[paragraph.Runs[0] with {FontFamily="Arial"}],SpacingAfterPoints=40};
    }
    private static DocxDocument Document(params DocxParagraph[] paragraphs) => new(
        612,792,72,180,72,72,DocxPageSettings.Empty,[],[],[],
        paragraphs.Select(paragraph=>(DocxBodyElement)new DocxParagraphElement(paragraph)).ToArray(),[],[])
        {
            MarkupMode=OoxPdfDocxMarkupMode.AllMarkup,
            RelatedStories=paragraphs.SelectMany(paragraph=>paragraph.InlineReferences).Select(reference=>reference.Id)
                .OfType<string>().Distinct(StringComparer.Ordinal)
                .Select(id=>new DocxRelatedStory(DocxRelatedStoryKind.Comment,"/word/comments.xml",id,
                    [new DocxParagraphElement(Paragraph("Comment body"))],[],[],null)).ToArray(),
            PackageCommentAnchorIds=paragraphs.SelectMany(paragraph=>paragraph.InlineReferences).Select(reference=>reference.Id).OfType<string>().ToArray()
        };

    private static (DocxParagraph First,DocxParagraph Second) FindCollision(DocxParagraph first,DocxParagraph second)
    {
        // Exercise real runtime collisions with unmodified production code.
        // Keep candidates alive; address reuse is not a hash collision.
        var seen=new Dictionary<int,(DocxParagraph Paragraph,bool First)>();
        for(int index=0;index<200000;index++) {
            bool isFirst=index%2==0;
            DocxParagraph paragraph=(isFirst ? first : second) with {};
            int hash=RuntimeHelpers.GetHashCode(paragraph);
            if(seen.TryGetValue(hash,out var previous) && previous.First!=isFirst)
                return isFirst ? (paragraph,previous.Paragraph) : (previous.Paragraph,paragraph);
            seen[hash]=(paragraph,isFirst);
        }
        throw new InvalidOperationException("Expected a distinct-owner runtime-hash collision within the bounded probe.");
    }
}
