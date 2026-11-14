using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static class DocxDeletedParagraphMarkTests
{
    public static void DocxDeletedParagraphMarksRetainCompleteOrderedSources()
    {
        foreach (bool chained in new[] { false, true })
        foreach (bool formatted in new[] { false, true })
        foreach (string container in new[] { "body", "cell", "header" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.SimpleMarkup,
            OoxPdfDocxMarkupMode.Original, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = Read(chained, mode, container, formatted);
            var paragraphs = container == "header" ? document.HeaderParagraphs :
                DocxBlockTraversal.EnumerateBodyParagraphs(document).Where(p => p.Equations.Count != 0 || p.Images.Count != 0 ||
                    p.Runs.Any(run => run.Text.Length != 0)).ToArray();
            bool merged = mode is OoxPdfDocxMarkupMode.Final or OoxPdfDocxMarkupMode.SimpleMarkup;
            TestAssert.Equal(merged ? 1 : chained ? 3 : 2, paragraphs.Count);
            TestAssert.Equal(2, paragraphs.Sum(p => p.Equations.Count));
            TestAssert.Equal(1, paragraphs.Sum(p => p.Images.Count));
            TestAssert.True(paragraphs.SelectMany(p => p.Images).All(image => image.PartName is not null && image.Bytes is { Length: > 0 }),
                "Each container resolves its picture through its own relationships.");
            if (!merged)
            {
                TestAssert.True(paragraphs.All(p => p.ConsumedParagraphMarks.Count == 0),
                    "Original/AllMarkup retain the source paragraph marks.");
                continue;
            }
            var paragraph = paragraphs.Single();
            var items = TestAssert.NotNull(paragraph.InlineContent);
            TestAssert.Equal(chained ? 2 : 1, paragraph.ConsumedParagraphMarks.Count);
            TestAssert.True(!paragraph.HasDeletedParagraphMark && paragraph.Revisions.Count == paragraph.ConsumedParagraphMarks.Count * (formatted ? 2 : 1),
                "Consume the marks while retaining their actual revision metadata.");
            TestAssert.True(DocxInlineContent.HasSupportedSourceOwner(paragraph), "Merged inline content has complete source owners.");
            TestAssert.True(items.OfType<DocxInlineEquation>().Select(item => item.Equation)
                .SequenceEqual(paragraph.Equations, ReferenceEqualityComparer.Instance), "Source and equation collections share ordered identities.");
            TestAssert.True(ReferenceEquals(items.OfType<DocxInlinePicture>().Single().Image, paragraph.Images.Single()),
                "Picture placement retains the shifted source image.");
            TestAssert.True(items.OfType<DocxInlineText>().All(item => paragraph.Runs.Any(run => ReferenceEquals(run, item.Run))),
                "Ordered text and paragraph runs share the same shifted objects.");
            TestAssert.True(DocxInlineContent.HasValidHyperlinkSourceRanges(paragraph, default), "Shift measured link boundaries with their source.");
            var linked = TestAssert.NotNull(DocxInlineContent.GetHyperlinkedEquations(paragraph, default));
            TestAssert.Equal(1, linked.Count);
            TestAssert.True(linked.Contains(paragraph.Equations.Last()), "The last formula retains its link after the join.");
            TestAssert.True(paragraph.Images.Single().SourceRunIndex >= 0 &&
                paragraph.Equations.Last().SourceRunIndex > paragraph.Equations.First().SourceRunIndex,
                "Object-only source runs contribute to join offsets.");
            TestAssert.True(paragraph.BookmarkAnchors.All(anchor => anchor.InlineItemIndex is >= 0),
                "Reconstructed terminal prose and object bookmarks retain ordered boundaries.");
        }
    }

    public static void DocxDeletedParagraphMarksRequireConsumedRevisionOwners()
    {
        var paragraph = Read(true, OoxPdfDocxMarkupMode.Final).Paragraphs.Single();
        TestAssert.True(DocxInlineContent.HasSupportedSourceOwner(paragraph), "Actual consumed marks own deletion metadata.");
        foreach (var replacement in new[] {
            paragraph with { ConsumedParagraphMarks = [] },
            paragraph with { ConsumedParagraphMarks = paragraph.ConsumedParagraphMarks.Select(mark =>
                mark with { Revisions = mark.Revisions.Select(revision => revision with { }).ToArray() }).ToArray() },
            paragraph with { ConsumedParagraphMarks = paragraph.ConsumedParagraphMarks.Select(mark => mark with { SourceRunIndex = 0 }).ToArray() },
            paragraph with { ConsumedParagraphMarks = paragraph.ConsumedParagraphMarks.Select(mark => mark with { Text = "visible" }).ToArray() }
        })
            TestAssert.True(!DocxInlineContent.HasSupportedSourceOwner(replacement),
                "Missing, value-equal, misplaced or visible objects cannot stand in for consumed source marks.");
        var original = Read(false, OoxPdfDocxMarkupMode.AllMarkup).Paragraphs.First();
        TestAssert.True(original.HasDeletedParagraphMark && DocxInlineContent.HasSupportedSourceOwner(original),
            "An actual unconsumed boundary owns its deletion without removing inline content.");
        TestAssert.True(!DocxInlineContent.HasSupportedSourceOwner(original with { ParagraphMarkRun = null }),
            "Removing the actual unconsumed mark loses its metadata owner.");
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxInlineContent.HasSupportedSourceOwner(paragraph, token: cancellation.Token));
    }

    public static void DocxDeletedParagraphMarkConversionRetainsMathPicturesAndLinks()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.SimpleMarkup,
            OoxPdfDocxMarkupMode.Original, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
            var renderer = new DocxRenderer(resolver, mode, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var pages = renderer.RenderPagesAsync(Read(false, mode), diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            TestAssert.True(text.Contains("x+2", StringComparison.Ordinal) && text.Contains("u+v=w", StringComparison.Ordinal),
                "Both source formulas survive paragraph-mark normalization: " + mode + "; text=" + text);
            TestAssert.True(text.Contains("PublicPrefix", StringComparison.Ordinal) && text.Contains("PublicSuffix", StringComparison.Ordinal),
                "Supported prose survives with the mathematics.");
            TestAssert.Equal(1, pages.Sum(page => page.Images.Count));
            TestAssert.True(pages.Sum(page => page.Annotations.Count(annotation => annotation.Destination is not null)) >= 1,
                "The formula's internal destination survives placement.");
            TestAssert.True(!diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_EQUATION"), "No recovered formula is silently omitted.");
            using var input = Input(false); using var output = new MemoryStream();
            OoxPdfConverter.Convert(input, output, new OoxPdfOptions {
                InputKind = OoxPdfInputKind.Docx, FontResolver = resolver, DocxMarkupMode = mode,
                FixedCreationDate = DateTimeOffset.UnixEpoch, DiagnosticSink = diagnostics.Add });
            string pdf = Encoding.Latin1.GetString(output.ToArray());
            TestAssert.True(pdf.StartsWith("%PDF-", StringComparison.Ordinal) && pdf.Contains("/Subtype /Image", StringComparison.Ordinal) &&
                pdf.Contains("/Annots", StringComparison.Ordinal), "Public conversion publishes the compound PDF.");
            TestAssert.True(!diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_EQUATION"), "Public conversion uses the same retained source.");
        }
    }

    private static DocxDocument Read(bool chained, OoxPdfDocxMarkupMode mode, string container = "body", bool formatted = false)
    {
        using var input = Input(chained, container, formatted); using var package = OoxPackage.Open(input, default);
        return new DocxReader().Read(package, null, default, mode);
    }

    private static MemoryStream Input(bool chained, string container = "body", bool formatted = false)
    {
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes/body.docx"));
        using var source = ZipFile.OpenRead(path);
        var parts = source.Entries.ToDictionary(entry => entry.FullName, entry => {
            using var stream = entry.Open(); using var copy = new MemoryStream(); stream.CopyTo(copy); return copy.ToArray(); });
        using var xmlStream = new MemoryStream(parts["word/document.xml"]); var xml = XDocument.Load(xmlStream);
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        var body = xml.Root!.Element(w + "body")!; var paragraph = body.Element(w + "p")!;
        var lastMath = paragraph.Elements(m + "oMath").Last();
        // Keep this formula distinct from the seed's header/footer mathematics.
        lastMath.Descendants(m + "t").Single().Value = "u+v=w";
        lastMath.ReplaceWith(new XElement(w + "hyperlink", new XAttribute(w + "anchor", "MixedObjects"), new XElement(lastMath)));
        var children = paragraph.Elements().Select(element => new XElement(element)).ToArray();
        var prefix = new XElement(w + "p", Text("PublicPrefix "));
        var groups = new List<XElement> { prefix };
        if (chained)
        {
            int drawing = Array.FindIndex(children, child => child.Descendants(w + "drawing").Any());
            prefix.Add(children.Take(drawing));
            groups.Add(new XElement(w + "p", children[drawing]));
            groups.Add(new XElement(w + "p", children.Skip(drawing + 1)));
        }
        else groups.Add(new XElement(w + "p", children));
        prefix.Add(new XElement(w + "bookmarkStart", new XAttribute(w + "id", "99"), new XAttribute(w + "name", "PrefixEnd")),
            new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "99")));
        groups.Last().Add(Text(" PublicSuffix"));
        for (int index = 0; index < groups.Count - 1; index++)
        {
            var mark = new XElement(w + "rPr", new XElement(w + "del",
                new XAttribute(w + "id", index + 1), new XAttribute(w + "author", "Public reviewer")));
            if (formatted)
                mark.Add(new XElement(w + "b"), new XElement(w + "rPrChange", new XAttribute(w + "id", 990 + index),
                    new XAttribute(w + "author", "Public reviewer"), new XElement(w + "rPr", new XElement(w + "b", new XAttribute(w + "val", "0")))));
            groups[index].AddFirst(new XElement(w + "pPr", mark));
        }
        if (container == "cell")
            paragraph.ReplaceWith(new XElement(w + "tbl", new XElement(w + "tblGrid", new XElement(w + "gridCol", new XAttribute(w + "w", "8640"))),
                new XElement(w + "tr", new XElement(w + "tc", groups))));
        else if (container == "header")
        {
            using var header = new MemoryStream(); new XDocument(new XElement(w + "hdr", groups)).Save(header);
            parts["word/header1.xml"] = header.ToArray();
            using var relationships = new MemoryStream(parts["word/_rels/document.xml.rels"]);
            var documentRelationships = XDocument.Load(relationships);
            XNamespace rel = "http://schemas.openxmlformats.org/package/2006/relationships";
            using var headerRelationships = new MemoryStream();
            new XDocument(new XElement(rel + "Relationships", documentRelationships.Root!.Elements(rel + "Relationship")
                .Where(relationship => ((string?)relationship.Attribute("Type"))?.EndsWith("/image", StringComparison.Ordinal) == true)
                .Select(relationship => new XElement(relationship)))).Save(headerRelationships);
            parts["word/_rels/header1.xml.rels"] = headerRelationships.ToArray();
            paragraph.ReplaceWith(new XElement(w + "p", Text("Public body")));
        }
        else paragraph.ReplaceWith(groups);
        using var updated = new MemoryStream(); xml.Save(updated, SaveOptions.DisableFormatting);
        parts["word/document.xml"] = updated.ToArray();
        return TestFixtures.CreateZipPackage(parts);

        XElement Text(string value) => new(w + "r", new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), value));
    }
}
