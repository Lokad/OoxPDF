using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxFieldCodeRevisionsRetainActualNonpaintingOwners()
    {
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        foreach (string instruction in new[] { "PAGE", "NUMPAGES", "REF StoredValue" })
        foreach (bool locked in new[] { false, true })
        foreach (bool splitInstruction in new[] { false, true })
        foreach (string slot in new[] { "begin", "instruction", "separate", "end" })
        {
            var document = ReadRevisedFieldCode("body", instruction, mode, slot, revised: true, locked, splitInstruction);
            var paragraph = DocxBlockTraversal.EnumerateBodyParagraphs(document)
                .Single(source => source.FieldReferences.Any(field => field.Instruction!.Trim() == instruction));
            var revision = paragraph.Revisions.Single(item => item.Id == "917");
            var owner = paragraph.NonRenderingRunOwners.Single();
            TestAssert.True(owner.SourceRunIndex >= 0 && owner.Revisions.Any(item => ReferenceEquals(item, revision)),
                "A nonpainting field-code source retains its own actual formatting history.");
            TestAssert.True(DocxInlineContent.HasSupportedSourceOwner(paragraph, hasPageInputs: true),
                "Code revisions cannot orphan independently supported content: " + instruction + "/" + slot + "/" + mode);
            bool hasPaintedOwner = paragraph.Runs.Any(run => run.Revisions.Any(item => ReferenceEquals(item, revision)));
            if (!hasPaintedOwner)
            {
                TestAssert.True(!DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph with { NonRenderingRunOwners = [] }, default),
                    "Aggregate field metadata cannot substitute for its source run.");
                TestAssert.True(!DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph with
                {
                    NonRenderingRunOwners = [owner with { Revisions = [revision with { }] }]
                }, default), "A value-equal history cannot invent a source owner.");
            }
        }
    }

    public static void DocxFieldCodeRevisionsKeepComposedContentAcrossStories()
    {
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (string story in new[] { "body", "cell", "header", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool splitInstruction in new[] { false, true })
        foreach (string slot in new[] { "begin", "instruction", "separate", "end" })
        {
            var renderer = new DocxRenderer(resolver, mode, CachedMathGeometry(mode));
            string MathText(IEnumerable<PdfPage> pages) => string.Concat(pages.SelectMany(page =>
                DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).Where(show => page.Fonts.Any(font =>
                    font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(font.Font, show.Font))))
                .Select(show => show.Text)).Normalize(NormalizationForm.FormKC);
            var expected = renderer.RenderPagesAsync(ReadRevisedFieldCode(story, "PAGE", mode, slot,
                revised: false, splitInstruction: splitInstruction), null, default).ToBlockingEnumerable().ToArray();
            var document = ReadRevisedFieldCode(story, "PAGE", mode, slot, revised: true, splitInstruction: splitInstruction);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var actual = renderer.RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            TestAssert.True(MathText(expected).Length > 0, "The matching source control contains mathematics.");
            TestAssert.Equal(expected.Length, actual.Length);
            TestAssert.Equal(MathText(expected), MathText(actual));
            TestAssert.Equal(1, actual.Sum(page => page.Images.Count));
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Nonpainting field history cannot remove formulas: " + story + "/" + slot + "/" + splitInstruction + "/" + mode);
            var expectedLinks = expected.SelectMany(page => page.Annotations).ToArray();
            var links = actual.SelectMany(page => page.Annotations).ToArray();
            TestAssert.True(expectedLinks.Where(link => link.IsUri).Select(link => link.Uri).Distinct().Order()
                .SequenceEqual(links.Where(link => link.IsUri).Select(link => link.Uri).Distinct().Order()),
                "All source URI targets survive field-code formatting history.");
            TestAssert.True(links.Count(link => link.IsDestination) >= expectedLinks.Count(link => link.IsDestination),
                "Source bookmark links retain their painted fragments.");
            TestAssert.True(links.All(link => double.IsFinite(link.X + link.Y + link.Width + link.Height) &&
                link.Width > 0d && link.Height > 0d && (link.Destination is not { } target ||
                    target.PageIndex >= 0 && target.PageIndex < actual.Length)),
                "Navigation keeps finite rectangles and admitted destination pages.");
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(document, null,
                cancellation.Token).ToBlockingEnumerable().ToArray());
        }
    }

    private static DocxDocument ReadRevisedFieldCode(string story, string instruction, OoxPdfDocxMarkupMode mode,
        string slot, bool revised, bool locked = false, bool splitInstruction = false) =>
        ReadMutation("CompositionProbes/math-composition-unnumbered" + (story == "body" ? "" : "-" + story),
            mode, true, true, (xml, w, _) =>
            {
                var drawing = xml.Descendants(w + "drawing").Single();
                var field = CachedMathField(w, new XElement(w + "r", new XElement(drawing)), complex: true, instruction, locked);
                if (splitInstruction)
                {
                    int cut = Math.Min(2, instruction.Length - 1);
                    var first = new XElement(w + "r", new XElement(w + "instrText", " " + instruction[..cut]));
                    var second = new XElement(w + "r", new XElement(w + "instrText", instruction[cut..]));
                    field = [field[0], first, second, .. field[2..]];
                }
                var source = slot == "instruction" ? field.First(element => element.Element(w + "instrText") is not null)
                    : field.Single(element => (string?)element.Element(w + "fldChar")?.Attribute(w + "fldCharType") == slot);
                if (revised)
                    source.AddFirst(new XElement(w + "rPr", new XElement(w + "rPrChange", new XAttribute(w + "id", "917"),
                        new XAttribute(w + "author", "Public formatter"), new XAttribute(w + "date", "2026-01-01T00:00:00Z"),
                        new XElement(w + "rPr", new XElement(w + "b")))));
                drawing.Parent!.ReplaceWith(field);
            }, partOverride: story is "header" or "footer" ? "word/" + story + "1.xml" : null);
}
