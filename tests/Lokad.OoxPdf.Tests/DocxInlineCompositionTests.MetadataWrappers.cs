using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineMetadataWrappersKeepResolvedParentFrames()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Parent frame comparison uses installed Windows fonts.");
        CheckInlineMetadataParentFrames(null);
    }

    internal static void CheckInlineMetadataParentFrames(IFontResolver? resolver)
    {
        string directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "InlineBoxMath"));
        using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "inline-metadata-wrappers.json")));
        foreach (var item in manifest.RootElement.EnumerateArray())
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string source = Path.GetFileNameWithoutExtension(item.GetProperty("SourceInput").GetString())!;
            var parent = ReadMutation("InlineBoxMath/" + source, mode, false, false);
            var wrapped = ReadMutation("InlineBoxMath/" + item.GetProperty("Id").GetString(), mode, false, false);
            var renderer = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final ?
                OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var before = renderer.RenderBlankPages(parent, null, default).ToArray();
            var after = renderer.RenderBlankPages(wrapped, null, default).ToArray();
            TestAssert.Equal(before.Length, after.Length);
            for (int index = 0; index < before.Length; index++)
            {
                TestAssert.Equal(before[index].Content, after[index].Content);
                TestAssert.Equal(before[index].Images.Count, after[index].Images.Count);
                TestAssert.Equal(before[index].Fonts.Count, after[index].Fonts.Count);
            }
        }
        foreach (string part in new[] { "header", "footer" })
        {
            var mode = OoxPdfDocxMarkupMode.Final;
            string fixture = "InlineBoxMath/inline-sibling-" + part;
            string sourcePart = "word/inline-box-" + part + ".xml";
            var before = ReadMutation(fixture, mode, false, false, partOverride: sourcePart);
            var after = ReadMutation(fixture, mode, false, false, (xml, w, _) =>
            {
                var paragraph = xml.Descendants(w + "p").First(p => p.Elements().Any(e => e.Name.LocalName == "oMath"));
                var children = paragraph.Elements().Where(e => e.Name != w + "pPr").ToArray();
                foreach (var child in children) child.Remove();
                paragraph.Add(MetadataWrapper(w, "customXml", MetadataWrapper(w, "smartTag", children)));
            }, sourcePart);
            var renderer = new DocxRenderer(resolver, mode, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            TestAssert.Equal(renderer.RenderBlankPages(before, null, default).Single().Content,
                renderer.RenderBlankPages(after, null, default).Single().Content);
        }
    }

    public static void DocxInlineMetadataWrappersMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent wrapper geometry uses installed Windows fonts.");
        const string golden = "inline-metadata-wrappers-office.json";
        CheckInlineBoxMath(null, golden);
        CheckInlineBoxWordOrigins(golden);
        CheckInlineBoxProseOrigins(golden);
        CheckInlineBoxFollowingOrigins(golden);
    }

    public static void DocxInlineMetadataWrappersKeepRevisionFieldAndLinkOwners()
    {
        foreach (string kind in new[] { "customXml", "smartTag" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupMode.Original })
        {
            var document = ReadMutation("Composition/report", mode, false, false, (xml, w, _) =>
            {
                XElement Run(string text) => new(w + "r", new XElement(w + "t", text));
                var link = new XElement(w + "hyperlink", new XAttribute(w + "anchor", "PublicTarget"),
                    new XElement(w + "bookmarkStart", new XAttribute(w + "id", "501"), new XAttribute(w + "name", "PublicTarget")),
                    Run("Visible"), new XElement(w + "fldSimple", new XAttribute(w + "instr", " REF PublicTarget "), Run("Cached")),
                    new XElement(w + "ins", new XAttribute(w + "id", "502"), new XAttribute(w + "author", "Public"),
                        new XAttribute(w + "date", "2026-01-01T00:00:00Z"), MetadataWrapper(w, kind, Run("Inserted"))),
                    new XElement(w + "del", new XAttribute(w + "id", "503"), new XAttribute(w + "author", "Public"),
                        new XAttribute(w + "date", "2026-01-01T00:00:00Z"), MetadataWrapper(w, kind,
                            new XElement(w + "r", new XElement(w + "delText", "Deleted")))),
                    new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "501")));
                ReplaceMetadataTestBody(xml, w, new XElement(w + "p", MetadataWrapper(w, kind, link)));
            });
            var paragraph = document.Paragraphs.Single();
            string expected = mode switch
            {
                OoxPdfDocxMarkupMode.Final => "VisibleCachedInserted",
                OoxPdfDocxMarkupMode.Original => "VisibleCachedDeleted",
                _ => "VisibleCachedInsertedDeleted"
            };
            TestAssert.Equal(expected, string.Concat(paragraph.Runs.Select(run => run.Text)));
            TestAssert.Equal(1, paragraph.BookmarkAnchors.Count);
            TestAssert.Equal(1, paragraph.Hyperlinks.Count);
            TestAssert.True(DocxInlineContent.HasValidHyperlinkSourceRanges(paragraph, default),
                "Metadata wrappers preserve the actual inline hyperlink owner range.");
            TestAssert.True(paragraph.FieldReferences.Single() is { HasCachedResult: true, RendersCachedResult: true },
                "Cached fields preserve their source result and field policy.");
        }
    }

    public static void DocxInlineMetadataWrappersKeepHardBreakBoundaries()
    {
        foreach (string kind in new[] { "customXml", "smartTag" })
        foreach (string type in new[] { "page", "column" })
        {
            var document = ReadMutation("Composition/report", OoxPdfDocxMarkupMode.Final, false, false, (xml, w, _) =>
                ReplaceMetadataTestBody(xml, w, new XElement(w + "p", MetadataWrapper(w, kind,
                    new XElement(w + "r", new XElement(w + "t", "Before"),
                        new XElement(w + "br", new XAttribute(w + "type", type)), new XElement(w + "t", "After"))))));
            TestAssert.Equal(3, document.BodyElements.Count);
            TestAssert.Equal("Before", string.Concat(((DocxParagraphElement)document.BodyElements[0]).Paragraph.Runs.Select(run => run.Text)));
            TestAssert.True(type == "page" ? document.BodyElements[1] is DocxPageBreakElement : document.BodyElements[1] is DocxManualBreakElement,
                "A hard break inside metadata remains an explicit body boundary.");
            TestAssert.Equal("After", string.Concat(((DocxParagraphElement)document.BodyElements[2]).Paragraph.Runs.Select(run => run.Text)));
        }
    }

    public static void DocxInlineMetadataWrappersKeepCancellationAndUnknownFallback()
    {
        var document = ReadMutation("Composition/report", OoxPdfDocxMarkupMode.Final, false, false, (xml, w, _) =>
            ReplaceMetadataTestBody(xml, w, new XElement(w + "p", new XElement(w + "r", new XElement(w + "t", "Known")),
                new XElement(XName.Get("customXml", "urn:future"), new XElement(w + "r", new XElement(w + "t", "Unqualified"))))));
        TestAssert.Equal("Known", document.Paragraphs.Single().Runs.Single().Text);
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "InlineBoxMath",
            "inline-wrapper-customxml-body-mixed.docx"));
        using var stream = File.OpenRead(path);
        var package = OoxPackage.Open(stream, default);
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => new DocxReader().Read(package, null, cancelled.Token, OoxPdfDocxMarkupMode.Final));
    }

    private static XElement MetadataWrapper(XNamespace w, string kind, params object[] children) =>
        new(w + kind, new XAttribute(w + "uri", "urn:lokad:public-composition"), new XAttribute(w + "element", "calculation"), children);

    private static void ReplaceMetadataTestBody(XDocument document, XNamespace w, XElement paragraph)
    {
        var body = document.Root!.Element(w + "body")!;
        var section = body.Element(w + "sectPr");
        body.ReplaceNodes(paragraph, section is null ? null : new XElement(section));
    }
}
