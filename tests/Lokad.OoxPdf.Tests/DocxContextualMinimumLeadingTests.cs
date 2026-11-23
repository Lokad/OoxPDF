using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using static Lokad.OoxPdf.Tests.DocxContextualLeadingFixture;
using static Lokad.OoxPdf.Tests.DocxMinimumBodyLineLeadingTests;

namespace Lokad.OoxPdf.Tests;

internal static class DocxContextualMinimumLeadingTests
{
    public static void DocxContextualMinimumKeepsResolvedStylesAndAutomaticNeighbors()
    {
        foreach (string styleMode in new[] { "implicit", "named", "mixed-default" })
        {
            var source = ReadContextual(false, (xml, styles) =>
            {
                if (styleMode == "implicit") return;
                AddStyle(styles, "PublicContextualA", styleMode == "mixed-default");
                var paragraphs = Body(xml);
                for (int i = 0; i < paragraphs.Length; i++)
                    if (styleMode == "named" || i % 2 != 0)
                        paragraphs[i].Element(Word + "pPr")!.AddFirst(new XElement(Word + "pStyle", new XAttribute(Word + "val", "PublicContextualA")));
            });
            TestAssert.True(source.Paragraphs.All(p => p.Spacing.ContextualSpacing == true && p.Spacing.BeforeValue == "80" && p.Spacing.AfterValue == "80"),
                "Raw contextual flags and numeric tokens remain available.");
            var original = CreatePreserveLayout(source);
            TestAssert.True(DocxMinimumBodyLineLeading.TryCreateLayout(source, original, out var adjusted), "Resolved contextual identities admit plain minimum leading.");
            var before = Lines(original); var after = Lines(adjusted);
            TestAssert.Equal(before.Length, after.Length);
            for (int i = 0; i < before.Length; i++)
            {
                if (before[i].LineHeightSource != DocxLineHeightSource.AtLeastLineSpacing)
                {
                    TestAssert.True(ReferenceEquals(before[i], after[i]), "Automatic neighbors retain their entire original lines.");
                    continue;
                }
                Near(before[i].LineHeight!.Value - before[i].SingleLineHeight!.Value, before[i].BaselineY - after[i].BaselineY, "Minimum excess moves its visible baseline.");
                Near(before[i].BaselineY + before[i].FirstLineInsetPoints!.Value, after[i].BaselineY + after[i].FirstLineInsetPoints!.Value,
                    "Contextual gap resolution and the authored paragraph top remain fixed.");
                TestAssert.Equal(before[i].AppliedBeforeSpacing, after[i].AppliedBeforeSpacing);
                TestAssert.Equal(before[i].PendingAfterSpacing, after[i].PendingAfterSpacing);
                TestAssert.True(ReferenceEquals(before[i].Segments, after[i].Segments), "Horizontal glyph plans retain identity.");
            }
            foreach (var mode in new[] { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
                TestAssert.Equal(0, Renderer(mode).RenderPagesAsync(source, null, CancellationToken.None).ToBlockingEnumerable().Sum(p => p.Annotations.Count));
        }
    }

    public static void DocxContextualMinimumActualNoteLinksRetainContextualGapOwnership()
    {
        foreach (var spec in new[] { (Mode: "same", Height: 20d), (Mode: "different", Height: 24d),
            (Mode: "predecessor-only", Height: 28d), (Mode: "current-only", Height: 20d) })
        {
            var source = ReadContextual(true, (xml, styles) =>
            {
                var paragraphs = Body(xml);
                if (spec.Mode == "different") { AddStyle(styles, "PublicContextualA"); AddStyle(styles, "PublicContextualB"); }
                for (int i = 0; i < paragraphs.Length; i++)
                {
                    var pr = paragraphs[i].Element(Word + "pPr")!;
                    if (spec.Mode == "different") pr.AddFirst(new XElement(Word + "pStyle", new XAttribute(Word + "val", i % 2 == 0 ? "PublicContextualA" : "PublicContextualB")));
                    if (spec.Mode is "predecessor-only" or "current-only")
                        pr.Element(Word + "contextualSpacing")!.SetAttributeValue(Word + "val", i % 2 == (spec.Mode == "current-only" ? 1 : 0) ? "1" : "0");
                }
                if (spec.Mode is "predecessor-only" or "current-only") Spacing(paragraphs[1]).SetAttributeValue(Word + "before", "160");
            });
            var original = CreatePreserveLayout(source);
            TestAssert.True(DocxMinimumBodyLineLeading.TryCreateLayout(source, original, out var adjusted), "Complete contextual note owners admit the baseline correction.");
            TestAssert.True(DocxPageFootnoteNumbering.HasStableCompleteNoteOwners(source, original, adjusted), "Every actual source marker and complete note retains its page.");
            var pages = Renderer(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderPagesAsync(source, null, CancellationToken.None).ToBlockingEnumerable().ToArray();
            var links = pages.SelectMany(p => p.Annotations).ToArray();
            TestAssert.Equal(4, links.Length);
            Near(spec.Height, links[0].Height, "The actual link owns the minimum slot and only the remaining contextual gaps.");
            TestAssert.Equal(4, links.Select(a => a.Destination).Distinct().Count());
            var preserve = Renderer(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderPagesAsync(source, null, CancellationToken.None).ToBlockingEnumerable().SelectMany(p => p.Annotations).ToArray();
            Near(20, preserve[0].Height, "Preserve keeps its original slot hit area.");
        }
    }

    public static void DocxContextualMinimumDistinguishesMissingDefaultsFromLocalExclusions()
    {
        foreach (bool notes in new[] { false, true })
        {
            foreach (bool removeStyles in new[] { true, false })
            {
                var source = ReadContextual(notes, (_, styles) =>
                {
                    if (removeStyles) return;
                    AddStyle(styles, "PublicContextualA", true);
                    AddStyle(styles, "PublicContextualB", true);
                }, removeStyles: removeStyles);
                var original = CreatePreserveLayout(source);
                TestAssert.True(!DocxMinimumBodyLineLeading.TryCreateLayout(source, original, out var fallback) &&
                    ReferenceEquals(original, fallback), "Without a resolved default style, no contextual line is corrected.");
            }

            foreach (bool unknownStyle in new[] { true, false })
            {
                var source = ReadContextual(notes, (xml, _) =>
                {
                    var first = Body(xml)[0];
                    if (unknownStyle)
                        first.Element(Word + "pPr")!.AddFirst(new XElement(Word + "pStyle", new XAttribute(Word + "val", "PublicMissing")));
                    else
                        Spacing(first).SetAttributeValue(Word + "beforeAutospacing", "1");
                });
                var original = CreatePreserveLayout(source);
                var excluded = Lines(original).Single(line => ReferenceEquals(line.SourceParagraph, source.Paragraphs[0]));
                var neighbor = Lines(original).Single(line => ReferenceEquals(line.SourceParagraph, source.Paragraphs[2]));
                double excess = neighbor.LineHeight!.Value - neighbor.SingleLineHeight!.Value;
                TestAssert.True(excess > 0, "The adjacent minimum-height paragraph has measurable excess.");
                TestAssert.True(DocxMinimumBodyLineLeading.TryCreateLayout(source, original, out var adjusted),
                    "A local exclusion leaves the adjacent minimum-height paragraph eligible.");
                var lines = Lines(adjusted);
                TestAssert.True(lines.Any(line => ReferenceEquals(line, excluded)),
                    "The excluded minimum-height paragraph keeps its original line.");
                var corrected = lines.Single(line => ReferenceEquals(line.SourceParagraph, neighbor.SourceParagraph));
                Near(excess, neighbor.BaselineY - corrected.BaselineY, "The adjacent baseline still moves by the excess.");
            }
        }
    }

    public static void DocxContextualMinimumInheritedFlagsKeepPreserveCapacityAndOwners()
    {
        foreach (bool notes in new[] { false, true })
        {
            var source = ReadContextual(notes, (xml, styles) =>
            {
                var defaults = styles.Root!.Element(Word + "docDefaults")!;
                defaults.Add(new XElement(Word + "pPrDefault", new XElement(Word + "pPr", new XElement(Word + "contextualSpacing"))));
                var paragraphs = Body(xml);
                foreach (var p in paragraphs) p.Element(Word + "pPr")!.Element(Word + "contextualSpacing")!.Remove();
                for (int i = 0; i < 40; i++)
                {
                    var spacer = new XElement(paragraphs[0]); spacer.Descendants(Word + "t").First().Value = $"Public contextual capacity spacer {i:00}.";
                    paragraphs[4].AddBeforeSelf(spacer);
                }
            });
            TestAssert.True(source.Paragraphs.All(p => p.Spacing.ContextualSpacing == true && p.ContextualSpacingDefaultStyleId == "Normal"), "Document-default flags retain a loaded implicit Normal identity.");
            var original = CreatePreserveLayout(source);
            TestAssert.True(original.Pages.Count > 1, "The synthetic contextual frame exercises page boundaries.");
            TestAssert.True(DocxMinimumBodyLineLeading.TryCreateLayout(source, original, out var adjusted), "The complete inherited frame qualifies without repagination.");
            TestAssert.Equal(original.Pages.Count, adjusted.Pages.Count);
            foreach (var pair in original.Pages.Zip(adjusted.Pages))
                TestAssert.True(pair.First.Items.OfType<DocxTextLineLayout>().Select(l => l.SourceBlockIndex).SequenceEqual(pair.Second.Items.OfType<DocxTextLineLayout>().Select(l => l.SourceBlockIndex)), "Each actual body owner retains its page.");
            var first = Lines(original)[0];
            var preserve = Renderer(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectTextEmission(source).Lines.First(l => l.StoryKind == "Body" && l.SourceBlockIndex == first.SourceBlockIndex);
            var word = Renderer(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).InspectTextEmission(source).Lines.First(l => l.StoryKind == "Body" && l.SourceBlockIndex == first.SourceBlockIndex);
            Near(first.LineHeight!.Value - first.SingleLineHeight!.Value, preserve.Segments.First().BaselineY - word.Segments.First().BaselineY, "Actual requested Word emission changes while Preserve keeps the original baseline.");
            var pages = Renderer(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderPagesAsync(source, null, CancellationToken.None).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(notes ? 4 : 0, pages.Sum(p => p.Annotations.Count));
            if (notes)
            {
                TestAssert.True(DocxPageFootnoteNumbering.HasStableCompleteNoteOwners(source, original, adjusted), "Complete inherited note ownership survives capacity boundaries.");
                TestAssert.Equal("105", source.Paragraphs.First(p => p.InlineReferences.Count != 0).InlineReferences[0].DisplayText);
                for (int i = 0; i < pages.Length; i++) TestAssert.True(pages[i].Annotations.All(a => a.Destination!.Value.PageIndex == i), "Each actual note link targets its owner page.");
            }
        }
    }

    private static void Near(double expected, double actual, string message) => TestAssert.True(Math.Abs(expected - actual) < .00001d, message);
    private static void AddStyle(XDocument styles, string id, bool isDefault = false)
    {
        var style = new XElement(Word + "style", new XAttribute(Word + "type", "paragraph"), new XAttribute(Word + "styleId", id), new XElement(Word + "name", new XAttribute(Word + "val", id)));
        if (isDefault) style.SetAttributeValue(Word + "default", "1");
        styles.Root!.Add(style);
    }
}
