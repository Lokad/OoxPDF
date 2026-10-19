using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionNoteTablesKeepPreparedSourceAndMeasuredCellOwners()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadMutation("CompositionProbes/mixed-note-" + kind + "-table", mode, false, false);
            var story = source.RelatedStories.Single(story => story.Id == "37");
            var paragraph = DocxBlockTraversal.EnumerateBodyParagraphs(story).Single(paragraph => paragraph.Equations.Count != 0);
            var owner = DocxMathParagraphPolicy.EnumerateStaticPreparationCandidates(source, default)
                .Single(item => ReferenceEquals(item.Paragraph, paragraph));
            TestAssert.True(!owner.HasPageInputs, "Note cells do not acquire static/body page-counter inputs.");
            TestAssert.Equal(2, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
            var math = paragraph.Equations.ToDictionary(equation => equation,
                equation => new DocxPreparedMathParagraph(equation, new(18d, 9d, 4d, []), 12d, 16d));
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: math);
            var measured = layout.RelatedStories.Single(item => ReferenceEquals(item.Story, story));
            var lines = measured.TableRows.SelectMany(row => row.Cells).SelectMany(cell => cell.TextLines).ToArray();
            TestAssert.Equal(2, lines.Sum(line => line.InlineMath.Count));
            TestAssert.True(lines.Where(line => line.InlineMath.Count != 0).All(line => ReferenceEquals(line.SourceParagraph, paragraph)),
                "Shared table measurement keeps the actual note paragraph instead of reconstructing its source.");
            var placed = layout.Pages.SelectMany(page => page.PlacedRelatedStories).Single(item => ReferenceEquals(item.StoryLayout.Story, story));
            TestAssert.Equal(2, placed.TableRows.SelectMany(row => row.Cells).SelectMany(cell => cell.TextLines).Sum(line => line.InlineMath.Count));
            TestAssert.True(!DocxPageFootnoteNumbering.IsEligibleDocument(source),
                "Adding measured table math does not admit complex-note typography/page-counter corrections.");
        }
    }

    public static void DocxInlineCompositionNoteTablesRetainComplexCellAndStoryFallbacks()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (string variant in new[] { "merge-restart", "merge-continuation", "merge-unknown", "vertical", "nested-vertical", "height", "unknown-story" })
        {
            var source = ReadMutation("CompositionProbes/mixed-note-" + kind + "-table", OoxPdfDocxMarkupMode.Final,
                false, false, (xml, w, _) =>
                {
                    var story = xml.Root!.Elements(w + kind).Single(story => (string?)story.Attribute(w + "id") == "37");
                    var table = story.Element(w + "tbl")!;
                    var cell = table.Descendants(w + "tc").Single();
                    if (variant == "unknown-story") story.SetAttributeValue(w + "type", "unknown");
                    else if (variant == "nested-vertical")
                    {
                        var paragraph = cell.Element(w + "p")!;
                        paragraph.ReplaceWith(new XElement(w + "tbl", new XElement(w + "tr", new XElement(w + "tc",
                            new XElement(w + "tcPr", new XElement(w + "textDirection", new XAttribute(w + "val", "tbRl"))), paragraph))));
                    }
                    else if (variant == "height")
                    {
                        var row = table.Element(w + "tr")!;
                        row.Element(w + "trPr")?.Remove();
                        row.AddFirst(new XElement(w + "trPr", new XElement(w + "trHeight",
                            new XAttribute(w + "val", "320"), new XAttribute(w + "hRule", "unknown"))));
                    }
                    else cell.Element(w + "tcPr")!.Add(variant switch
                    {
                        "merge-restart" => new XElement(w + "vMerge", new XAttribute(w + "val", "restart")),
                        "merge-continuation" => new XElement(w + "vMerge"),
                        "merge-unknown" => new XElement(w + "vMerge", new XAttribute(w + "val", "unknown")),
                        _ => new XElement(w + "textDirection", new XAttribute(w + "val", "tbRl"))
                    });
                }, partOverride: "word/" + kind + "s.xml");
            var candidates = DocxMathParagraphPolicy.EnumerateStaticPreparationCandidates(source, default).ToArray();
            TestAssert.Equal(variant == "merge-restart" ? 2 : 0,
                candidates.Sum(candidate => candidate.Paragraph.Equations.Count));
        }
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var control = ReadMutation("CompositionProbes/mixed-note-footnote-table", OoxPdfDocxMarkupMode.Final, false, false);
        TestAssert.Throws<OperationCanceledException>(() =>
            DocxMathParagraphPolicy.EnumerateStaticPreparationCandidates(control, cancellation.Token).ToArray());
    }

    public static void DocxInlineCompositionNoteTablesKeepIndependentContentThroughSourceSplitting()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native note-table controls use installed Windows fonts.");
        CheckNoteTableComposition(null);
    }

    internal static void CheckNoteTableComposition(IFontResolver? resolver)
    {
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var golden = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "mixed-note-containers-office.json")));
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(Path.Combine(folder, id + ".docx")))));
            var mode = row.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            string kind = id.Contains("footnote", StringComparison.Ordinal) ? "footnote" : "endnote";
            var source = ReadMutation("CompositionProbes/" + id, mode, false, false);
            var transformed = ReadMutation("CompositionProbes/" + id, mode, true, true, partOverride: "word/" + kind + "s.xml");
            var diagnostics = new List<OoxPdfDiagnostic>();
            var renderer = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var before = renderer.RenderBlankPages(source, diagnostics.Add, default).ToArray();
            var after = renderer.RenderBlankPages(transformed, diagnostics.Add, default).ToArray();
            TestAssert.Equal(before.Length, after.Length);
            foreach (var pages in new[] { before, after })
            {
                string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
                TestAssert.Equal(row.GetProperty("ExpectedEquations").GetInt32(), text.Split("ratetime", StringSplitOptions.None).Length - 1);
                foreach (string prose in new[] { "Public note calculation", "Alpha", "Delta" }) TestAssert.Contains(prose, text);
                TestAssert.Equal(row.GetProperty("ImagePlacements").GetInt32(), pages.Sum(page => page.Images.Sum(image =>
                    Regex.Matches(page.Content, "/" + Regex.Escape(image.ResourceName) + @"\s+Do\b").Count)));
                var links = pages.SelectMany(page => page.Annotations).ToArray();
                TestAssert.Equal(4, links.Length);
                TestAssert.Equal(4, links.Select(link => link.Destination).Distinct().Count());
                TestAssert.True(links.All(link => link.Destination is not null && link.Width > 0d && link.Height > 0d),
                    "Every existing source note retains a distinct complete destination.");
            }
            for (int page = 0; page < before.Length; page++)
            {
                TestAssert.Equal(DocxPageFootnoteNumberingTests.DecodeText(before[page]), DocxPageFootnoteNumberingTests.DecodeText(after[page]));
                TestAssert.Equal(before[page].Annotations.Count, after[page].Annotations.Count);
                for (int link = 0; link < before[page].Annotations.Count; link++)
                    TestAssert.Equal(before[page].Annotations[link], after[page].Annotations[link]);
            }
            TestAssert.True(diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Simple note-table placement and neutral annotations must retain supported mathematical content.");
        }
    }
}
