using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxParagraphMarkCaptionsShareResolvedRunFormattingAndFallbacks()
    {
        using var gold = ReadParagraphMarkCaptionGolden();
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            var paragraph = ParagraphMarkCaptionOwner(ReadParagraphMarkCaptionSource(row, mode));
            var mark = paragraph.ParagraphMarkRun!;
            var revision = mark.Revisions.Single();
            var change = TestAssert.NotNull(revision.RunFormattingChange);
            TestAssert.True(change.Bold == true || change.Italic == true, "The existing resolver retains the authored boolean change.");
            TestAssert.True(paragraph.Revisions.Any(owner => ReferenceEquals(owner, revision)), "The typed snapshot is shared before source publication.");
            TestAssert.Equal(row.GetProperty("Body").GetString(), DocxRenderer.BuildWordCompatibleRunFormattingPreview(paragraph, default));
            TestAssert.True(DocxRenderer.BuildWordCompatibleRunFormattingPreview(paragraph with
                { Revisions = [revision with { }] }, default) is null, "A copied aggregate cannot substitute for the actual mark owner.");
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxRenderer.BuildWordCompatibleRunFormattingPreview(paragraph, canceled.Token));
        }
        foreach (string kind in new[] { "unknown-value", "size", "character-style", "duplicate", "mixed-role" })
        {
            var source = ReadMutation("CompositionProbes/math-mark-property-body", OoxPdfDocxMarkupMode.AllMarkup, false, false, (xml, w, m) =>
            {
                var change = xml.Descendants(w + "rPrChange").Single();
                var mark = change.Parent!;
                switch (kind)
                {
                    case "unknown-value": mark.Element(w + "b")!.SetAttributeValue(w + "val", "unknown"); break;
                    case "size":
                        mark.AddFirst(new XElement(w + "sz", new XAttribute(w + "val", "22")));
                        change.Element(w + "rPr")!.Add(new XElement(w + "sz", new XAttribute(w + "val", "20"))); break;
                    case "character-style": mark.AddFirst(new XElement(w + "rStyle", new XAttribute(w + "val", "DefaultParagraphFont"))); break;
                    case "duplicate": change.AddAfterSelf(new XElement(change)); break;
                    case "mixed-role": mark.AddFirst(new XElement(w + "ins", new XAttribute(w + "id", "993"))); break;
                }
            });
            TestAssert.True(DocxRenderer.BuildWordCompatibleRunFormattingPreview(ParagraphMarkCaptionOwner(source), default) is null,
                "Uncertain or unsupported source history retains the generic description: " + kind);
        }
        var off = ReadMutation("CompositionProbes/math-mark-property-body", OoxPdfDocxMarkupMode.AllMarkup, false, false, (xml, w, m) =>
        {
            var change = xml.Descendants(w + "rPrChange").Single();
            change.Parent!.Element(w + "b")!.SetAttributeValue(w + "val", "0");
            change.Element(w + "rPr")!.Element(w + "b")!.SetAttributeValue(w + "val", "1");
        });
        TestAssert.Equal("Font: Not Bold", DocxRenderer.BuildWordCompatibleRunFormattingPreview(ParagraphMarkCaptionOwner(off), default));
        var combined = ReadMutation("CompositionProbes/math-mark-property-body", OoxPdfDocxMarkupMode.AllMarkup, false, false, (xml, w, m) =>
        {
            var change = xml.Descendants(w + "rPrChange").Single();
            change.Parent!.AddFirst(new XElement(w + "i"));
            change.Element(w + "rPr")!.Add(new XElement(w + "i", new XAttribute(w + "val", "0")));
        });
        TestAssert.Equal("Font: Bold, Italic", DocxRenderer.BuildWordCompatibleRunFormattingPreview(ParagraphMarkCaptionOwner(combined), default));
    }

    public static void DocxParagraphMarkCaptionsPreservePaintedMathLinksAndSourceIdentity()
    {
        using var gold = ReadParagraphMarkCaptionGolden();
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var document = ReadParagraphMarkCaptionSource(row, OoxPdfDocxMarkupMode.AllMarkup);
            var paragraph = ParagraphMarkCaptionOwner(document);
            var mark = paragraph.ParagraphMarkRun;
            var page = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderPagesAsync(document, null, default).ToBlockingEnumerable().Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).ToArray();
            var title = shows.Single(show => show.Text.Trim() == row.GetProperty("Title").GetString());
            var body = shows.Single(show => show.Text.Trim() == row.GetProperty("Body").GetString());
            TestAssert.True(Math.Abs(title.Y - body.Y) < .001d && title.X < body.X, "The supported description remains one ordered caption row.");
            var math = shows.Where(show => page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) &&
                ReferenceEquals(font.Font, show.Font))).ToArray();
            TestAssert.Equal(18, math.Length);
            TestAssert.Equal("x+2ya+b=cx+2ya+b=c", string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
            TestAssert.True(page.Annotations.Count != 0, "Caption selection must retain the composite document's source links.");
            TestAssert.True(ReferenceEquals(mark, ParagraphMarkCaptionOwner(document).ParagraphMarkRun), "Painting cannot replace the actual source mark.");
        }
    }

    public static void DocxParagraphMarkCaptionsMatchIndependentWordTextAndHorizontalOrigins()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Independent caption origins require installed Word-compatible fonts.");
        using var gold = ReadParagraphMarkCaptionGolden();
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var page = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderPagesAsync(ReadParagraphMarkCaptionSource(row, OoxPdfDocxMarkupMode.AllMarkup), null, default).ToBlockingEnumerable().Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).ToArray();
            var title = shows.Single(show => show.Text.Trim() == row.GetProperty("Title").GetString());
            var body = shows.Single(show => show.Text.Trim() == row.GetProperty("Body").GetString());
            TestAssert.True(Math.Abs(title.X - row.GetProperty("TitleX").GetDouble()) < 1d &&
                Math.Abs(body.X - row.GetProperty("BodyX").GetDouble()) < 1d,
                "All caption sources retain the existing one-point Word horizontal-origin gate: " + row.GetProperty("Input").GetString());
        }
    }

    private static JsonDocument ReadParagraphMarkCaptionGolden() =>
        JsonDocument.Parse(File.ReadAllText(NumberingMarkCase("math-mark-captions-office.json")));

    private static DocxDocument ReadParagraphMarkCaptionSource(JsonElement row, OoxPdfDocxMarkupMode mode)
    {
        string path = NumberingMarkCase(row.GetProperty("Input").GetString()!);
        TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
        return DocxTests.ReadDocx(path, mode);
    }

    private static DocxParagraph ParagraphMarkCaptionOwner(DocxDocument document) => DocxBlockTraversal.EnumerateSourceParagraphs(document)
        .Distinct<DocxParagraph>(ReferenceEqualityComparer.Instance).Single(p => p.ParagraphMarkRun?.Revisions.Any(r => r.Id == "992") == true);
}
