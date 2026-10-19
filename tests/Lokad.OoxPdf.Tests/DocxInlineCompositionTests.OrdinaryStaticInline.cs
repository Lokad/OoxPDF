using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxOrdinaryStaticInlineFramesPreserveSourceAndContainerOwnership()
    {
        foreach (string kind in new[] { "header", "footer" })
        foreach (string content in new[] { "box", "picture" })
        {
            var document = ReadMutation("InlineBoxMath/static-inline-ordinary-" + kind + "-" + content,
                OoxPdfDocxMarkupMode.Final, false, false,
                partOverride: "word/static-inline-ordinary-" + kind + "-" + content + ".xml");
            var source = (kind == "header" ? document.HeaderBodyElementsByType : document.FooterBodyElementsByType)
                ["default"].OfType<DocxParagraphElement>().Single().Paragraph;
            TestAssert.Equal(0, source.Equations.Count);
            TestAssert.True(DocxInlineContent.HasSupportedSourceOwner(source), "The static source retains its parsed ordered owners.");
            var page = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(document, new DocxTests.FamilyWidthTextMeasurer(), default).Pages.Single();
            var box = page.StaticInlineTextBoxes.Single();
            TestAssert.True(ReferenceEquals(source.InlineTextBoxes.Single(), box.TextBox), "The complete frame keeps its source owner.");
            TestAssert.True(box.Story?.Kind == (kind == "header" ? DocxStoryKind.Header : DocxStoryKind.Footer),
                "The placed frame retains its selected static story.");
            TestAssert.Equal(content == "picture" ? 1 : 0, page.StaticInlineImages.Count);
            var lines = page.StaticTextLines.Where(line => ReferenceEquals(line.SourceParagraph, source)).ToArray();
            TestAssert.True(lines.Length != 0 && lines.All(line => line.InlineFrame is not null),
                "The container places the shared measured frame.");
            TestAssert.True(box.BoxTop <= page.Height && box.BoxTop - box.BoxHeight >= 0d,
                "The complete static frame remains inside its page.");
        }
    }

    public static void DocxOrdinaryStaticInlineOriginsMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent ordinary static origins use installed Windows fonts.");
        const string golden = "static-inline-ordinary-office.json";
        CheckInlineBoxMath(null, golden);
        CheckInlineBoxProseOrigins(golden);
        CheckInlineBoxFollowingOrigins(golden);
    }
}
