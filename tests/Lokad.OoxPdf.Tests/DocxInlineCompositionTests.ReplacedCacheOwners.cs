using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxReplacedCacheRevisionsRetainActualNonpaintingOwners()
    {
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        foreach (string instruction in new[] { "PAGE", "NUMPAGES", "REF StoredValue" })
        foreach (bool locked in new[] { false, true })
        foreach (string shape in new[] { "plain", "ordered", "field" })
        {
            var document = ReadRevisedComputedCache("body", instruction, mode, shape, revised: true, locked);
            var paragraph = DocxBlockTraversal.EnumerateBodyParagraphs(document)
                .Single(source => source.FieldReferences.Any(field => field.Instruction!.Trim() == instruction));
            bool replaced = !locked && instruction != "REF StoredValue";
            var revision = paragraph.Revisions.Single(item => item.Id == "713");
            TestAssert.Equal(replaced ? 1 : 0, paragraph.NonRenderingRunOwners.Count);
            TestAssert.True(DocxInlineContent.HasSupportedSourceOwner(paragraph, hasPageInputs: true),
                "Replaced and stored caches retain their own actual revision owner: " + instruction + "/" + shape + "/" + mode);
            if (replaced)
            {
                var owner = paragraph.NonRenderingRunOwners.Single();
                TestAssert.True(owner.SourceRunIndex >= 0 && owner.Revisions.Any(item => ReferenceEquals(item, revision)),
                    "Multiple cached text children retain one real nonpainting source run.");
                TestAssert.True(!DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph with { NonRenderingRunOwners = [] }, default),
                    "The actual replaced source is required; aggregate metadata alone cannot qualify.");
                TestAssert.True(!DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph with {
                    NonRenderingRunOwners = [owner with { Revisions = [revision with { }] }] }, default),
                    "A value-equal revision copy cannot invent ownership.");
            }
            else
            {
                TestAssert.True(paragraph.Runs.Any(run => run.Text.Contains("sta", StringComparison.Ordinal) &&
                    run.Revisions.Any(item => ReferenceEquals(item, revision))),
                    "Stored and locked results keep painting through their existing resolved runs.");
            }
        }
    }

    public static void DocxReplacedCacheRevisionsKeepNeighborMathPicturesAndNavigationAcrossStories()
    {
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (string story in new[] { "body", "cell", "header", "footer" })
        foreach (string instruction in new[] { "PAGE", "NUMPAGES" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(resolver, mode, CachedMathGeometry(mode));
            string MathText(IEnumerable<PdfPage> pages) => string.Concat(pages.SelectMany(page =>
                DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).Where(show => page.Fonts.Any(font =>
                    font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(font.Font, show.Font))))
                .Select(show => show.Text)).Normalize(NormalizationForm.FormKC);
            foreach (string shape in new[] { "plain", "ordered", "field" })
            {
                var expected = renderer.RenderPagesAsync(ReadRevisedComputedCache(story, instruction, mode, shape, revised: false),
                    null, default).ToBlockingEnumerable().ToArray();
                TestAssert.True(MathText(expected).Length > 0 && expected.Sum(page => page.Images.Count) == 1,
                    "The matching source-shape control contains its supported formulas and picture.");
                var document = ReadRevisedComputedCache(story, instruction, mode, shape, revised: true);
                var diagnostics = new List<OoxPdfDiagnostic>();
                var actual = renderer.RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
                TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"),
                    "A replaced cache's formatting history cannot discard neighboring formulas: " + story + "/" + instruction + "/" + shape + "/" + mode);
                TestAssert.Equal(expected.Length, actual.Length);
                TestAssert.Equal(MathText(expected), MathText(actual));
                TestAssert.Equal(1, actual.Sum(page => page.Images.Count));
                // Adding a real formatting revision can open the Review lane,
                // changing physical rectangles and hyperlink line fragments.
                // Check retained targets and valid placement independently of
                // the existing source-pinned Office geometry assertions.
                var expectedLinks = expected.SelectMany(page => page.Annotations).ToArray();
                var links = actual.SelectMany(page => page.Annotations).ToArray();
                TestAssert.True(expectedLinks.Where(link => link.IsUri).Select(link => link.Uri).Distinct()
                    .Order().SequenceEqual(links.Where(link => link.IsUri).Select(link => link.Uri).Distinct().Order()),
                    "Every source URI target survives cache replacement.");
                TestAssert.True(links.Count(link => link.IsDestination) >= expectedLinks.Count(link => link.IsDestination),
                    "Source bookmark links retain their painted fragments.");
                TestAssert.True(links.All(link => double.IsFinite(link.X + link.Y + link.Width + link.Height) &&
                    link.Width > 0d && link.Height > 0d && (link.Destination is not { } target ||
                        target.PageIndex >= 0 && target.PageIndex < actual.Length &&
                        double.IsFinite(target.Left ?? 0d) && double.IsFinite(target.Top ?? 0d))),
                    "Retained navigation has finite geometry and admitted destination pages.");
                using var cancellation = new CancellationTokenSource();
                cancellation.Cancel();
                TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(document, null,
                    cancellation.Token).ToBlockingEnumerable().ToArray());
            }
        }
    }

    private static DocxDocument ReadRevisedComputedCache(string story, string instruction,
        OoxPdfDocxMarkupMode mode, string shape, bool revised, bool locked = false) =>
        ReadMutation("CompositionProbes/math-composition-unnumbered" + (story == "body" ? "" : "-" + story),
            mode, true, true, (xml, w, _) =>
            {
                var drawing = xml.Descendants(w + "drawing").Single();
                var value = new XElement(w + "r",
                    revised ? new XElement(w + "rPr", new XElement(w + "rPrChange", new XAttribute(w + "id", "713"),
                        new XAttribute(w + "author", "Public formatter"), new XAttribute(w + "date", "2026-01-01T00:00:00Z"),
                        new XElement(w + "rPr", new XElement(w + "b")))) : null,
                    new XElement(w + "t", "sta"), new XElement(w + "t", "le"));
                object result = shape == "ordered" ? value : new[] { value, new XElement(w + "r", new XElement(drawing)) };
                if (shape == "ordered") value.Add(new XElement(drawing));
                var field = CachedMathField(w, result, complex: true, instruction, locked);
                if (shape == "field")
                {
                    value.Add(new XElement(field[^1].Element(w + "fldChar")!));
                    field = field[..^1];
                }
                drawing.Parent!.ReplaceWith(field);
            }, partOverride: story is "header" or "footer" ? "word/" + story + "1.xml" : null);
}
