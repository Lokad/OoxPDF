using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxComputedFieldBoxesRetainSourceOwnersVisibilityAndMixedDrawingOrder()
    {
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        foreach (string revision in new[] { "none", "ins", "del" })
        foreach (bool hidden in new[] { false, true })
        foreach (bool locked in new[] { false, true })
        foreach (bool pictureFirst in new[] { false, true })
        {
            var document = ReadComputedFieldBox("body", "PAGE", mode, "simple", hidden, locked, revision,
                mixedPicture: true, pictureFirst: pictureFirst);
            var paragraphs = DocxBlockTraversal.EnumerateSourceParagraphs(document)
                .Distinct<DocxParagraph>(ReferenceEqualityComparer.Instance).ToArray();
            bool included = revision == "none" || (revision == "ins" ? mode != OoxPdfDocxMarkupMode.Original :
                mode is OoxPdfDocxMarkupMode.Original or OoxPdfDocxMarkupMode.AllMarkup);
            var parent = paragraphs.SingleOrDefault(p => p.FieldReferences.Count != 0);
            if (!included)
            {
                TestAssert.True(parent is null, "An excluded field wrapper cannot publish a box or counter.");
                continue;
            }
            var paragraph = TestAssert.NotNull(parent);
            TestAssert.Equal(hidden ? 0 : 1, paragraph.InlineTextBoxes.Count);
            TestAssert.Equal(1, paragraph.Images.Count);
            TestAssert.Equal(!locked, paragraph.FieldReferences.Single().UsesPlaceholder);
            TestAssert.True(DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, default),
                "Cached boxes and pictures keep their actual visible revision references.");
            var items = TestAssert.NotNull(paragraph.InlineContent);
            TestAssert.Equal(paragraph.InlineTextBoxes.Count, items.OfType<DocxInlineBox>().Count());
            TestAssert.Equal(1, items.OfType<DocxInlinePicture>().Count());
            if (hidden) continue;
            var box = paragraph.InlineTextBoxes.Single();
            TestAssert.True(ReferenceEquals(box, items.OfType<DocxInlineBox>().Single().Box),
                "Publication retains the actual child-story owner.");
            int boxIndex = items.ToList().FindIndex(item => item is DocxInlineBox);
            int pictureIndex = items.ToList().FindIndex(item => item is DocxInlinePicture);
            TestAssert.Equal(pictureFirst, pictureIndex < boxIndex);
            if (!locked)
            {
                var counter = paragraph.Runs.Single(run => run.FieldKind == DocxFieldKind.Page);
                TestAssert.Equal(counter.SourceRunIndex, box.SourceRunIndex);
                TestAssert.Equal(counter.SourceRunIndex, paragraph.Images.Single().SourceRunIndex);
                TestAssert.Equal(pictureFirst ? 1 : 0, box.SourceOrderInRun);
                TestAssert.Equal(pictureFirst ? 0 : 1, paragraph.Images.Single().SourceOrderInRun);
            }
        }
    }

    public static void DocxComputedFieldBoxesKeepContentAcrossStoriesAndCacheRunBoundaries()
    {
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (string story in new[] { "body", "cell", "header", "footer" })
        foreach (string kind in new[] { "PAGE", "NUMPAGES" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(resolver, mode, CachedMathGeometry(mode));
            var direct = renderer.RenderPagesAsync(ReadComputedFieldBox(story, kind, mode, "complex"),
                null, default).ToBlockingEnumerable().ToArray();
            TestAssert.True(direct.Length != 0, "The existing complex field retains its supported box.");
            foreach (string shape in new[] { "simple", "joined", "wrapped" })
            {
                var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
                var source = ReadComputedFieldBox(story, kind, mode, shape);
                TestAssert.Equal(1, DocxBlockTraversal.EnumerateSourceParagraphs(source)
                    .Distinct<DocxParagraph>(ReferenceEqualityComparer.Instance).Sum(p => p.InlineTextBoxes.Count));
                var actual = renderer.RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                    "Counter cache encoding cannot remove supported box or sibling mathematics: " + story + "/" + kind + "/" + shape + "/" + mode);
                AssertMathContentControlPages(direct, actual);
            }
            var removed = renderer.RenderPagesAsync(ReadComputedFieldBox(story, kind, mode, "removed"),
                null, default).ToBlockingEnumerable().ToArray();
            var hiddenDiagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
            var hiddenPages = renderer.RenderPagesAsync(ReadComputedFieldBox(story, kind, mode, "simple", hidden: true),
                hiddenDiagnostics.Add, default).ToBlockingEnumerable().ToArray();
            AssertMathContentControlPages(removed, hiddenPages);
            TestAssert.True(hiddenDiagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Intentionally hidden cached box equations retain nonprinting ownership.");
        }
    }

    private static DocxDocument ReadComputedFieldBox(string story, string kind, OoxPdfDocxMarkupMode mode,
        string shape, bool hidden = false, bool locked = false, string revision = "none",
        bool mixedPicture = false, bool pictureFirst = false)
    {
        string fixture = mixedPicture ? "inline-box-picture-body" : story switch
        {
            "cell" => "inline-box-math-cell",
            "header" => "inline-box-math-header",
            "footer" => "inline-box-math-footer",
            _ => "inline-box-body-equation-inline-box-valid"
        };
        return ReadMutation("InlineBoxMath/" + fixture, mode, false, false, (xml, w, _) =>
        {
            var run = xml.Descendants(w + "txbxContent").Single().Ancestors(w + "r").First();
            var box = new XElement(run);
            if (hidden)
            {
                var properties = box.Element(w + "rPr");
                if (properties is null) { properties = new(w + "rPr"); box.AddFirst(properties); }
                properties.Elements(w + "vanish").Remove();
                properties.Add(new XElement(w + "vanish"));
            }
            var value = new XElement(w + "r", new XElement(w + "t", locked ? "1" : "7"));
            var content = new List<XElement> { value };
            XElement? picture = null;
            if (mixedPicture)
            {
                picture = new XElement(w + "r", new XElement(run.Descendants(w + "drawing")
                    .First(drawing => !drawing.Descendants(w + "txbxContent").Any())));
                if (pictureFirst) content.Add(picture);
            }
            if (shape != "removed") content.Add(box);
            if (picture is not null && !pictureFirst) content.Add(picture);
            if (shape == "joined")
                content = [new XElement(w + "r", content.SelectMany(r => r.Elements()).Select(e => new XElement(e)))];
            if (shape == "complex")
            {
                run.ReplaceWith(new XElement(w + "r", new XElement(w + "fldChar", new XAttribute(w + "fldCharType", "begin"))),
                    new XElement(w + "r", new XElement(w + "instrText", " " + kind + " ")),
                    new XElement(w + "r", new XElement(w + "fldChar", new XAttribute(w + "fldCharType", "separate"))),
                    content, new XElement(w + "r", new XElement(w + "fldChar", new XAttribute(w + "fldCharType", "end"))));
                return;
            }
            var field = new XElement(w + "fldSimple", new XAttribute(w + "instr", " " + kind + " "),
                locked ? new XAttribute(w + "fldLock", "true") : null);
            if (shape == "wrapped") field.Add(new XElement(w + "sdt", new XElement(w + "sdtPr"), new XElement(w + "sdtContent", content)));
            else field.Add(content);
            run.ReplaceWith(revision == "none" ? field : new XElement(w + revision,
                new XAttribute(w + "id", "932"), new XAttribute(w + "author", "Public reviewer"), field));
        }, partOverride: story is "header" or "footer" ? "word/inline-box-" + story + ".xml" : null);
    }
}
