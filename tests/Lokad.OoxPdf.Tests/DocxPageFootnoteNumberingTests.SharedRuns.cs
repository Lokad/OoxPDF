using System.IO.Compression;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxPageFootnoteNumberingTests
{
    public static void DocxSharedRunFootnoteNumbersKeepIndependentGeneratedOwners()
    {
        foreach (bool authoredDigits in new[] { false, true })
        foreach (bool ordered in new[] { false, true })
        {
            DocxDocument source = ReadSharedMarkerFixture(merge: true, authoredDigits);
            if (ordered)
            {
                var replacements = new Dictionary<DocxParagraph, DocxParagraph>(ReferenceEqualityComparer.Instance);
                foreach (var p in source.Paragraphs)
                    replacements.Add(p, p with { InlineContent = DocxInlineContent.FromPrefix(p.Runs, p.Images, default) });
                var publication = new DocxParagraphReplacements(replacements, default);
                source = source with { BodyElements = publication.ReplaceElements(source.BodyElements),
                    FallbackParagraphs = source.FallbackParagraphs.Select(publication.Replace).ToArray() };
            }
            var paragraph = source.Paragraphs.First(p => p.InlineReferences.Count != 0);
            var pair = paragraph.InlineReferences;
            TestAssert.Equal(pair[0].SourceRunIndex, pair[1].SourceRunIndex);
            TestAssert.Equal(authoredDigits ? 3 : 0, pair[1].TextOffsetInRun - pair[0].TextOffsetInRun);
            TestAssert.True(!ReferenceEquals(pair[0].DisplayRun, pair[1].DisplayRun), "XML coordinates do not identify generated marker runs.");
            TestAssert.True(DocxPageFootnoteNumbering.TryCreateDocument(source, CreateLayout(source), out var numbered, out var owners,
                ownerScope: DocxFootnoteOwnerScope.AllSourceReferences), "Distinct measured marker owners can restart even inside one XML run.");
            TestAssert.Equal("09,10,09,10", string.Join(',', numbered.Paragraphs.SelectMany(p => p.InlineReferences).Select(r => r.DisplayText)));
            TestAssert.Equal("1,1,2,2", string.Join(',', new[] { "37", "4", "73", "8" }.Select(id => owners[id])));
            foreach (var (before, after) in source.Paragraphs.Zip(numbered.Paragraphs))
            {
                foreach (var run in before.Runs.Where(r => !r.IsAutomaticNoteReferenceMarker))
                    TestAssert.True(after.Runs.Any(r => ReferenceEquals(r, run)), "Authored digits and text keep their exact source owners.");
                foreach (var (oldReference, reference) in before.InlineReferences.Zip(after.InlineReferences))
                {
                    TestAssert.Equal(oldReference.SourceRunIndex, reference.SourceRunIndex);
                    TestAssert.Equal(oldReference.RunChildIndex, reference.RunChildIndex);
                    TestAssert.Equal(oldReference.TextOffsetInRun, reference.TextOffsetInRun);
                    TestAssert.True(reference.DisplayRun is { } marker && marker.Text == reference.DisplayText &&
                        after.Runs.Any(r => ReferenceEquals(r, marker)) && (!ordered ||
                        after.InlineContent!.OfType<DocxInlineText>().Count(item => ReferenceEquals(item.Run, marker) &&
                            item.Start == 0 && item.Length == marker.Text.Length) == 1),
                        "Rewritten labels rebind both references and ordered inline content to the same complete marker.");
                }
            }
            TestAssert.Equal("105,106,107,108", string.Join(',', source.Paragraphs.SelectMany(p => p.InlineReferences).Select(r => r.DisplayText)));
            TestAssert.True(DocxPageFootnoteNumbering.HasStableNumberingPages(numbered, CreateLayout(numbered), owners,
                ownerScope: DocxFootnoteOwnerScope.AllSourceReferences), "Each replacement preserves its independently measured page.");
        }
    }

    public static void DocxSharedRunFootnoteNumbersRenderLikeSeparateMarkerRuns()
    {
        foreach (var geometry in new[] { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
        {
            PdfPage[][] variants = new[] { false, true }.Select(merge =>
                new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, geometry)
                    .RenderPagesAsync(ReadSharedMarkerFixture(merge), null, CancellationToken.None).ToBlockingEnumerable().ToArray()).ToArray();
            TestAssert.Equal(2, variants[0].Length); TestAssert.Equal(2, variants[1].Length);
            for (int i = 0; i < 2; i++)
            {
                TestAssert.Equal(variants[0][i].Content, variants[1][i].Content);
                TestAssert.Equal(2, variants[1][i].Annotations.Count);
                TestAssert.Equal(2, variants[1][i].Annotations.Select(a => a.Destination).Distinct().Count());
                TestAssert.True(variants[0][i].Annotations.SequenceEqual(variants[1][i].Annotations),
                    "Merging identically formatted marker runs preserves painted content and destination geometry.");
                string text = DecodeText(variants[1][i]);
                TestAssert.Contains("0910", text);
                TestAssert.Contains("09 Public destination", text); TestAssert.Contains("10 Public destination", text);
            }
        }
    }

    public static void DocxSharedRunFootnoteNumbersRejectMissingDuplicateAndForeignMeasuredOwners()
    {
        DocxDocument source = ReadSharedMarkerFixture(merge: true);
        DocxLayout layout = CreateLayout(source);
        var paragraph = source.Paragraphs.First(p => p.InlineReferences.Count != 0);
        var reference = paragraph.InlineReferences[0];
        var line = layout.Pages[0].Items.OfType<DocxTextLineLayout>().Single(l => ReferenceEquals(l.SourceParagraph, paragraph));
        var marker = line.Segments.Single(s => ReferenceEquals(s.StyleRun, reference.DisplayRun));
        IReadOnlyList<DocxTextSegmentLayout>[] variants = [
            line.Segments.Where(s => !ReferenceEquals(s, marker)).ToArray(),
            line.Segments.Append(marker).ToArray(),
            line.Segments.Select(s => ReferenceEquals(s, marker) ? s with { StyleRun = s.StyleRun with { } } : s).ToArray()
        ];
        foreach (var segments in variants)
        {
            var pages = layout.Pages.ToArray();
            pages[0] = pages[0] with { Items = pages[0].Items.Select(i => ReferenceEquals(i, line) ? line with { Segments = segments } : i).ToArray() };
            TestAssert.True(!DocxPageFootnoteNumbering.TryCreateDocument(source, layout with { Pages = pages }, out var fallback, out _,
                ownerScope: DocxFootnoteOwnerScope.AllSourceReferences), "Exactly one segment must belong to this marker, regardless of equal text or source coordinates.");
            TestAssert.True(ReferenceEquals(source, fallback), "Rejected ownership preserves the original document.");
        }
    }

    private static DocxDocument ReadSharedMarkerFixture(bool merge, bool authoredDigits = false)
    {
        string seed = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "note-each-page-paragraph-break.docx"));
        var parts = new Dictionary<string, string>();
        using (ZipArchive archive = ZipFile.OpenRead(seed))
            foreach (var entry in archive.Entries) { using var reader = new StreamReader(entry.Open()); parts[entry.FullName] = reader.ReadToEnd(); }
        XDocument xml = XDocument.Parse(parts["word/document.xml"]);
        XElement[] markers = xml.Descendants(Word + "footnoteReference").ToArray();
        for (int i = 0; i < markers.Length; i += 2)
        {
            XElement first = markers[i].Parent!, second = markers[i + 1].Parent!;
            TestAssert.Equal(first.Element(Word + "rPr")?.ToString() ?? "", second.Element(Word + "rPr")?.ToString() ?? "");
            second.Remove(); first.AddAfterSelf(second);
            if (!merge || i != 0) continue;
            if (authoredDigits) first.Add(new XElement(Word + "t", "105"));
            first.Add(second.Elements().Where(e => e.Name != Word + "rPr").Select(e => new XElement(e)));
            second.Remove();
        }
        // The already page-restarted fixture supplies all note/section settings.
        // Only marker placement in document.xml changes; every other part stays.
        parts["word/document.xml"] = xml.ToString();
        string input = TestFixtures.WriteTempPackage(".docx", parts);
        try { return DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup); }
        finally { File.Delete(input); }
    }
}
