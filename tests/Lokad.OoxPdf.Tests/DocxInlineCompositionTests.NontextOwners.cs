using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxNontextRunRevisionsRetainActualOwnersAcrossPayloads()
    {
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        foreach (string kind in new[] { "hint", "trimmed", "picture", "box", "empty" })
        {
            var document = ReadRevisedNontextRun("body", kind, mode, revised: true);
            var paragraph = DocxBlockTraversal.EnumerateSourceParagraphs(document)
                .First(source => source.Revisions.Any(revision => revision.Id == "923"));
            var revision = paragraph.Revisions.Single(item => item.Id == "923");
            var owner = paragraph.NonRenderingRunOwners.Single();
            TestAssert.True(owner.SourceRunIndex >= 0 && owner.Revisions.Any(item => ReferenceEquals(item, revision)),
                "A parsed run retains its actual history independently of its nontext child kind.");
            TestAssert.True(DocxInlineContent.HasSupportedSourceOwner(paragraph, hasPageInputs: true),
                "Nontext source ownership cannot veto neighboring supported content: " + kind + "/" + mode);
            foreach (var missing in new[]
            {
                paragraph with { NonRenderingRunOwners = [] },
                paragraph with { NonRenderingRunOwners = [owner with { SourceRunIndex = -1 }] },
                paragraph with { NonRenderingRunOwners = [owner with { Revisions = [revision with { }] }] }
            })
                TestAssert.True(!DocxInlineContent.HasSupportedParagraphRevisionOwners(missing, default),
                    "Unowned, unknown-position and value-equal histories retain their fallback.");
        }
    }

    public static void DocxNontextRunRevisionsKeepMathPicturesAndBoxesAcrossStories()
    {
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (string story in new[] { "body", "cell", "header", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string kind in new[] { "hint", "trimmed", "picture", "box", "empty" })
        {
            var renderer = new DocxRenderer(resolver, mode, CachedMathGeometry(mode));
            var expected = renderer.RenderPagesAsync(ReadRevisedNontextRun(story, kind, mode, revised: false),
                null, default).ToBlockingEnumerable().ToArray();
            var document = ReadRevisedNontextRun(story, kind, mode, revised: true);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var actual = renderer.RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            TestAssert.True(NontextOwnerMath(expected).Length > 0, "The source control contains supported mathematics.");
            TestAssert.Equal(expected.Length, actual.Length);
            TestAssert.Equal(NontextOwnerMath(expected), NontextOwnerMath(actual));
            TestAssert.Equal(expected.Sum(page => page.Images.Count), actual.Sum(page => page.Images.Count));
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Run history cannot remove supported math: " + story + "/" + kind + "/" + mode);
            var expectedLinks = expected.SelectMany(page => page.Annotations).ToArray();
            var links = actual.SelectMany(page => page.Annotations).ToArray();
            TestAssert.True(expectedLinks.Where(link => link.IsUri).Select(link => link.Uri).Distinct().Order()
                .SequenceEqual(links.Where(link => link.IsUri).Select(link => link.Uri).Distinct().Order()),
                "Every source URI target survives nontext-run ownership.");
            TestAssert.True(links.Count(link => link.IsDestination) >= expectedLinks.Count(link => link.IsDestination),
                "Source bookmark links retain their painted fragments.");
            TestAssert.True(links.All(link => double.IsFinite(link.X + link.Y + link.Width + link.Height) &&
                link.Width > 0d && link.Height > 0d && (link.Destination is not { } target ||
                    target.PageIndex >= 0 && target.PageIndex < actual.Length)),
                "Navigation retains finite rectangles and admitted destination pages.");
            if (kind != "hint") continue;
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(document, null,
                cancellation.Token).ToBlockingEnumerable().ToArray());
        }
    }

    private static string NontextOwnerMath(IEnumerable<PdfPage> pages) => string.Concat(pages.SelectMany(page =>
        DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).Where(show => page.Fonts.Any(font =>
            font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(font.Font, show.Font))))
        .Select(show => show.Text)).Normalize(NormalizationForm.FormKC);

    private static DocxDocument ReadRevisedNontextRun(string story, string kind, OoxPdfDocxMarkupMode mode, bool revised)
    {
        bool box = kind == "box";
        string fixture = box ? "InlineBoxMath/" + (story == "body" ? "inline-box-body-equation-inline-box-valid" : "inline-box-math-" + story)
            : "CompositionProbes/math-composition-unnumbered" + (story == "body" ? "" : "-" + story);
        string? part = story is "header" or "footer" ? box ? "word/inline-box-" + story + ".xml" : "word/" + story + "1.xml" : null;
        return ReadMutation(fixture, mode, true, true, (xml, w, _) =>
        {
            var drawing = box ? xml.Descendants(w + "drawing").First(element => element.Descendants(w + "txbxContent").Any())
                : xml.Descendants(w + "drawing").Single();
            var source = drawing.Parent!;
            if (kind is not ("picture" or "box"))
            {
                var inserted = new XElement(w + "r", kind switch
                {
                    "hint" => new XElement(w + "lastRenderedPageBreak"),
                    "trimmed" => new XElement(w + "t", "   "),
                    _ => null
                });
                source.AddBeforeSelf(inserted);
                source = inserted;
            }
            if (!revised) return;
            var properties = source.Element(w + "rPr");
            if (properties is null) { properties = new XElement(w + "rPr"); source.AddFirst(properties); }
            TestAssert.True(properties.Element(w + "rPrChange") is null && properties.Element(w + "b") is null,
                "The source control owns no existing formatting history or explicit bold.");
            var historical = new XElement(properties);
            historical.Add(new XElement(w + "b"));
            properties.Add(new XElement(w + "rPrChange", new XAttribute(w + "id", "923"),
                new XAttribute(w + "author", "Public formatter"), new XAttribute(w + "date", "2026-01-01T00:00:00Z"), historical));
        }, partOverride: part);
    }
}
