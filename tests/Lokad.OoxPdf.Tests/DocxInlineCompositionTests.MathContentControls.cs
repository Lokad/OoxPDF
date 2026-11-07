using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxMathContentControlsPreserveBodyAndCellPaintAcrossViews()
    {
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        foreach (string fixture in new[] { "CompositionProbes/math-inner-bookmark-outer",
            "CompositionProbes/math-picture-first-table-single-cell" })
        {
            var renderer = MathContentControlRenderer(mode);
            var direct = renderer.RenderPagesAsync(Read(false), null, default).ToBlockingEnumerable().ToArray();
            foreach (string scope in new[] { "row", "numerator" })
            {
                var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
                var controlled = renderer.RenderPagesAsync(Read(true, scope), diagnostics.Add, default)
                    .ToBlockingEnumerable().ToArray();
                TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"),
                    "Stored supported math remains eligible inside the content control.");
                AssertMathContentControlPages(direct, controlled);
            }

            DocxDocument Read(bool control, string scope = "row") => ReadMutation(fixture, mode, false, false,
                (xml, w, m) =>
                {
                    if (!control) return;
                    var numerator = xml.Descendants(m + "num").First();
                    var owner = scope == "row" ? numerator.Ancestors(m + "oMath").First() : numerator;
                    var children = owner.Elements().Where(child =>
                        !(child.Name.Namespace == m && child.Name.LocalName.EndsWith("Pr", StringComparison.Ordinal))).ToArray();
                    foreach (var child in children) child.Remove();
                    owner.Add(MathContentControl(children, w));
                });
        }
    }

    public static void DocxMathContentControlsPreserveActualRevisionAndBookmarkOwners()
    {
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        foreach (string kind in new[] { "ins", "del" })
        foreach (bool control in new[] { false, true })
        {
            var renderer = MathContentControlRenderer(mode);
            var source = ReadMathRevisionBoundary(mode, kind, control, "bookmark");
            var direct = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            foreach (string placement in new[] { "inside", "outside" })
            {
                var document = ReadMathRevisionBoundary(mode, kind, control, "bookmark", placement);
                var paragraph = document.Paragraphs.First();
                TestAssert.True(DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, default),
                    "A content control cannot replace or invent the actual mathematical revision owner.");
                TestAssert.Equal(source.Paragraphs.First().BookmarkAnchors.Count, paragraph.BookmarkAnchors.Count);
                AssertMathContentControlPages(direct,
                    renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray());
            }
        }
    }

    private static DocxRenderer MathContentControlRenderer(OoxPdfDocxMarkupMode mode) =>
        new(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
            mode == OoxPdfDocxMarkupMode.AllMarkup ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);

    private static void AssertMathContentControlPages(IReadOnlyList<Lokad.OoxPdf.Pdf.PdfPage> before,
        IReadOnlyList<Lokad.OoxPdf.Pdf.PdfPage> after)
    {
        TestAssert.Equal(before.Count, after.Count);
        for (int index = 0; index < before.Count; index++)
        {
            TestAssert.Equal(before[index].Content, after[index].Content);
            TestAssert.Equal(before[index].Fonts.Count, after[index].Fonts.Count);
            TestAssert.Equal(before[index].Images.Count, after[index].Images.Count);
            TestAssert.Equal(before[index].Annotations.Count, after[index].Annotations.Count);
            for (int link = 0; link < before[index].Annotations.Count; link++)
            {
                TestAssert.Equal(before[index].Annotations[link], after[index].Annotations[link]);
            }
        }
    }

    private static XElement MathContentControl(IEnumerable<XElement> children, XNamespace w) =>
        new(w + "sdt", new XElement(w + "sdtPr", new XElement(w + "id", new XAttribute(w + "val", "19131")),
            new XElement(w + "tag", new XAttribute(w + "val", "PublicMathControl"))),
            new XElement(w + "sdtContent", children));
}
