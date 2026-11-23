using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineMetadataWrappersKeepResolvedParentFrames()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Parent frame comparison uses installed Windows fonts.");
        CheckInlineMetadataParentFrames(null);
    }

    internal static void CheckMetadataParentFrames(IFontResolver? resolver, string manifestFile)
    {
        string directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "InlineBoxMath"));
        using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, manifestFile)));
        foreach (var item in manifest.RootElement.EnumerateArray())
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string source = Path.GetFileNameWithoutExtension(item.GetProperty("SourceInput").GetString())!;
            var parent = ReadMutation("InlineBoxMath/" + source, mode, false, false);
            var wrapped = ReadMutation("InlineBoxMath/" + item.GetProperty("Id").GetString(), mode, false, false);
            var renderer = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final ?
                OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var before = renderer.RenderPagesAsync(parent, null, default).ToBlockingEnumerable().ToArray();
            var after = renderer.RenderPagesAsync(wrapped, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(before.Length, after.Length);
            for (int index = 0; index < before.Length; index++)
            {
                TestAssert.Equal(before[index].Content, after[index].Content);
                TestAssert.Equal(before[index].Images.Count, after[index].Images.Count);
                TestAssert.Equal(before[index].Fonts.Count, after[index].Fonts.Count);
            }
        }
    }

    internal static void CheckInlineMetadataParentFrames(IFontResolver? resolver)
    {
        CheckMetadataParentFrames(resolver, "inline-metadata-wrappers.json");
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
            TestAssert.Equal(renderer.RenderPagesAsync(before, null, default).ToBlockingEnumerable().Single().Content,
                renderer.RenderPagesAsync(after, null, default).ToBlockingEnumerable().Single().Content);
        }
    }

    public static void DocxInlineMetadataWrappersMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent wrapper geometry uses installed Windows fonts.");
        const string golden = "inline-metadata-wrappers-office.json";
        CheckInlineBoxReference(golden, InlineBoxChecks.All);
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
        foreach (bool splitOwner in new[] { false, true })
        {
            var document = ReadMutation("Composition/report", OoxPdfDocxMarkupMode.Final, false, false, (xml, w, _) =>
            {
                XElement before = new(w + "t", "Before");
                XElement after = new(w + "t", "After");
                XElement lineBreak = new(w + "br", new XAttribute(w + "type", type));
                XElement[] content = splitOwner
                    ? [MetadataWrapper(w, kind, new XElement(w + "r", before)), new XElement(w + "r", lineBreak, after)]
                    : [MetadataWrapper(w, kind, new XElement(w + "r", before, lineBreak, after))];
                ReplaceMetadataTestBody(xml, w, new XElement(w + "p", content));
            });
            TestAssert.Equal(3, document.BodyElements.Count);
            TestAssert.Equal("Before", string.Concat(((DocxParagraphElement)document.BodyElements[0]).Paragraph.Runs.Select(run => run.Text)));
            TestAssert.True(type == "page" ? document.BodyElements[1] is DocxPageBreakElement : document.BodyElements[1] is DocxManualBreakElement,
                "A hard break inside metadata remains an explicit body boundary.");
            TestAssert.Equal("After", string.Concat(((DocxParagraphElement)document.BodyElements[2]).Paragraph.Runs.Select(run => run.Text)));
        }
    }

    public static void DocxInlineMetadataLeadingBreaksKeepBookmarkOnThePaintedTarget()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupMode.Original })
        foreach (string type in new[] { "page", "column" })
        {
            var geometry = mode == OoxPdfDocxMarkupMode.AllMarkup ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup :
                OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout;
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode, geometry);
            PdfPage[] control = Render("direct");
            int targetPage = type == "page" ? 1 : 0;
            var link = control[0].Annotations.Single();
            TestAssert.Equal(targetPage + 1, control.Length);
            TestAssert.True(link.Destination is { } destination && destination.PageIndex == targetPage &&
                (type == "page" || destination.Left > 200d), "The direct bookmark reaches the continuation page or column.");
            TestAssert.Contains("AUTHORED AFTER BREAK", string.Concat(
                DocxCommentsTests.ReadEmbeddedGlyphTextShows(control[targetPage]).Select(show => show.Text)));

            foreach (string fieldOwner in new[] { "empty-field", "hint-field", "locked-field", "nested-hint-field" })
            {
                CheckNonprinting(Render(fieldOwner));
            }
            PdfPage[] storedField = Render("visible-field", retainLeadingFragment: true);
            TestAssert.Contains("VISIBLE STORED FIELD", string.Concat(
                DocxCommentsTests.ReadEmbeddedGlyphTextShows(storedField[0]).Select(show => show.Text)));
            TestAssert.True(storedField[0].Annotations.Single().Destination is { PageIndex: 0, Left: < 200d },
                "A visible stored result retains the bookmark before the break.");

            if (mode == OoxPdfDocxMarkupMode.Original)
            {
                CheckNonprinting(Render("hidden-break"));
                CheckNonprinting(Render("excluded-insertion"));
                CheckNonprinting(Render("excluded-move", includeDrawings: true));
                CheckNonprinting(Render("closed-move-range"));
                CheckNonprinting(Render("excluded-move-range"));
                PdfPage[] openRange = Render("open-move-range", retainLeadingFragment: true);
                TestAssert.True(openRange[0].Annotations.Single().Destination is { PageIndex: 0, Left: < 200d },
                    "An open move range keeps its authored owner.");
                PdfPage[] moved = Render("selected-move-range", retainLeadingFragment: true);
                TestAssert.Contains("VISIBLE MOVE", string.Concat(
                    DocxCommentsTests.ReadEmbeddedGlyphTextShows(moved[0]).Select(show => show.Text)));
                TestAssert.True(moved[0].Annotations.Single().Destination is { PageIndex: 0, Left: < 200d },
                    "A move range containing selected ink retains its leading fragment.");
                PdfPage[] deletion = Render("selected-deletion", retainLeadingFragment: true);
                TestAssert.Contains("VISIBLE DELETION", string.Concat(
                    DocxCommentsTests.ReadEmbeddedGlyphTextShows(deletion[0]).Select(show => show.Text)));
                TestAssert.True(deletion[0].Annotations.Single().Destination is { PageIndex: 0, Left: < 200d },
                    "Original retains the pre-break fragment of a visible deletion.");
                continue;
            }

            foreach (string wrapper in new[] { "sdt", "customXml", "smartTag", "nested", "proof", "permission",
                "run-hint", "preceding-hint", "nested-preceding-hint", "adjacent-sdt-hint",
                "adjacent-custom-xml-hint", "nested-adjacent-hint", "adjacent-hyperlink-hint",
                "nested-hyperlink-hint", "empty-revision", "editing-only-revision" })
            {
                CheckNonprinting(Render(wrapper));
            }

            PdfPage[] field = Render("page-field-hint", retainLeadingFragment: true);
            TestAssert.Equal(targetPage + 1, field.Length);
            TestAssert.True(DocxCommentsTests.ReadEmbeddedGlyphTextShows(field[0]).Any(show => show.Text.Trim() == "1"),
                "The PAGE field still paints its result in the pre-break fragment.");
            TestAssert.True(field[0].Annotations.Single().Destination is { PageIndex: 0, Left: < 200d },
                "The bookmark belongs to the generated content before the break.");

            PdfPage[] visible = Render("visible-hyperlink-hint", retainLeadingFragment: true);
            TestAssert.Contains("VISIBLE BEFORE BREAK", string.Concat(
                DocxCommentsTests.ReadEmbeddedGlyphTextShows(visible[0]).Select(show => show.Text)));
            TestAssert.Equal(2, visible[0].Annotations.Count);
            TestAssert.True(visible[0].Annotations.All(annotation => annotation.Destination is { PageIndex: 0, Left: < 200d }),
                "Both links follow the bookmark to its visible pre-break fragment.");

            PdfPage[] revision = Render("visible-revision", retainLeadingFragment: true);
            TestAssert.Contains("VISIBLE REVISION", string.Concat(
                DocxCommentsTests.ReadEmbeddedGlyphTextShows(revision[0]).Select(show => show.Text)));
            TestAssert.True(revision[0].Annotations.Single().Destination is { PageIndex: 0, Left: < 200d },
                "The bookmark reaches the visible insertion before the break.");

            PdfPage[] revisionRun = Render("empty-revision-run", retainLeadingFragment: true);
            TestAssert.True(revisionRun[0].Annotations.Single().Destination is { PageIndex: 0, Left: < 200d },
                "An authored run retains its explicit revision owner.");

            void CheckNonprinting(PdfPage[] pages)
            {
                TestAssert.True(control.Select(page => page.Content).SequenceEqual(pages.Select(page => page.Content)),
                    "Nonprinting content preserves the authored pagination and paint.");
                TestAssert.Equal(link, pages[0].Annotations.Single());
            }

            PdfPage[] Render(string wrapper, bool retainLeadingFragment = false, bool includeDrawings = false)
            {
                var document = ReadMutation("Composition/report", mode, false, false, (xml, w, _) =>
                {
                    var body = xml.Root!.Element(w + "body")!;
                    var section = new XElement(body.Element(w + "sectPr")!);
                    if (type == "column")
                    {
                        section.SetElementValue(w + "cols", null);
                        section.Add(new XElement(w + "cols", new XAttribute(w + "num", 2)));
                    }

                    var run = new XElement(w + "r", new XElement(w + "br", new XAttribute(w + "type", type)),
                        new XElement(w + "t", "AUTHORED AFTER BREAK"));
                    XElement FloatingDrawing(int height) => new(w + "drawing",
                        new XElement(OoxNamespaces.WordprocessingDrawingNamespace + "anchor", new XAttribute("relativeHeight", height)));
                    if (includeDrawings) run.Add(FloatingDrawing(2));
                    XElement Hint() => new(w + "lastRenderedPageBreak");
                    XElement HintRun() => new(w + "r", Hint());
                    XElement Link(params XElement[] items) => new(w + "hyperlink",
                        new XAttribute(w + "anchor", "destination"), items);
                    XElement Revision(string kind, params XElement[] items) => new(w + kind,
                        new XAttribute(w + "id", 71), new XAttribute(w + "author", "Public probe"), items);
                    XElement MoveStart(string kind) => new(w + (kind + "RangeStart"), new XAttribute(w + "id", 79),
                        new XAttribute(w + "name", "PublicMove"), new XAttribute(w + "author", "Public probe"));
                    XElement MoveEnd(string kind) => new(w + (kind + "RangeEnd"), new XAttribute(w + "id", 79));
                    XElement ContentControl(params XElement[] items) => new(w + "sdt",
                        new XElement(w + "sdtPr", new XElement(w + "alias", new XAttribute(w + "val", "Break probe"))),
                        new XElement(w + "sdtContent", items));
                    XElement StoredField(string instruction, params XElement[] items) =>
                        new(w + "fldSimple", new XAttribute(w + "instr", instruction), items);
                    XElement Metadata(string kind, XElement item) => MetadataWrapper(w, kind, new XElement(w + (kind + "Pr")), item);
                    if (wrapper == "run-hint")
                    {
                        run.AddFirst(Hint());
                    }
                    XElement[] content = wrapper switch
                    {
                        "direct" or "run-hint" => [run],
                        "preceding-hint" => [HintRun(), run],
                        "nested-preceding-hint" => [ContentControl(MetadataWrapper(w, "customXml",
                            new XElement(w + "customXmlPr"), HintRun(), run))],
                        "adjacent-sdt-hint" => [ContentControl(HintRun()), run],
                        "adjacent-custom-xml-hint" => [Metadata("customXml", HintRun()), run],
                        "nested-adjacent-hint" => [ContentControl(Metadata("smartTag", HintRun()), run)],
                        "adjacent-hyperlink-hint" => [Link(HintRun()), run],
                        "nested-hyperlink-hint" => [ContentControl(Link(HintRun())), run],
                        "empty-revision" => [Revision("ins"), run],
                        "editing-only-revision" => [Revision("ins", ContentControl(
                            new XElement(w + "proofErr", new XAttribute(w + "type", "spellStart")))), run],
                        "visible-revision" => [Revision("ins", new XElement(w + "r", new XElement(w + "t", "VISIBLE REVISION"))), run],
                        "empty-revision-run" => [Revision("ins", ContentControl(new XElement(w + "r"))), run],
                        "empty-field" => [StoredField("REF EmptyReference"), run],
                        "hint-field" => [StoredField("REF EmptyReference", HintRun()), run],
                        "locked-field" => [StoredField("PAGE", HintRun()), run],
                        "nested-hint-field" => [ContentControl(StoredField("REF EmptyReference", HintRun())), run],
                        "visible-field" => [StoredField("REF VisibleReference",
                            new XElement(w + "r", Hint(), new XElement(w + "t", "VISIBLE STORED FIELD"))), run],
                        "hidden-break" => [ContentControl(new XElement(w + "r", new XElement(w + "rPr", new XElement(w + "vanish")),
                            new XElement(w + "br", new XAttribute(w + "type", type)))), run],
                        "excluded-insertion" => [Revision("ins", new XElement(w + "r", new XElement(w + "t", "EXCLUDED INSERTION"))), run],
                        "excluded-move" => [ContentControl(Revision("moveTo", new XElement(w + "r",
                            new XElement(w + "t", "EXCLUDED MOVE"), new XElement(w + "br", new XAttribute(w + "type", type)),
                            FloatingDrawing(1)))), run],
                        "closed-move-range" => [ContentControl(MoveStart("moveTo"), MoveEnd("moveTo")), run],
                        "excluded-move-range" => [MoveStart("moveTo"),
                            ContentControl(Revision("moveTo", new XElement(w + "r", new XElement(w + "t", "EXCLUDED MOVE")))),
                            MoveEnd("moveTo"), run],
                        "open-move-range" => [MoveStart("moveTo"), run],
                        "selected-move-range" => [MoveStart("moveFrom"),
                            Revision("moveFrom", new XElement(w + "r", new XElement(w + "delText", "VISIBLE MOVE"))),
                            MoveEnd("moveFrom"), run],
                        "selected-deletion" => [Revision("del", new XElement(w + "r", new XElement(w + "delText", "VISIBLE DELETION"))), run],
                        "page-field-hint" => [new XElement(w + "fldSimple", new XAttribute(w + "instr", "PAGE"), HintRun()), run],
                        "visible-hyperlink-hint" => [Link(new XElement(w + "r", Hint(), new XElement(w + "t", "VISIBLE BEFORE BREAK"))), run],
                        "sdt" => [ContentControl(run)],
                        "nested" => [ContentControl(Metadata("customXml", Metadata("smartTag", run)))],
                        "proof" => [new XElement(w + "proofErr", new XAttribute(w + "type", "spellStart")),
                            run, new XElement(w + "proofErr", new XAttribute(w + "type", "spellEnd"))],
                        "permission" => [ContentControl(
                            new XElement(w + "permStart", new XAttribute(w + "id", 50), new XAttribute(w + "edGrp", "everyone")),
                            run, new XElement(w + "permEnd", new XAttribute(w + "id", 50)))],
                        _ => [Metadata(wrapper, run)]
                    };
                    if (wrapper == "locked-field")
                        content[0].SetAttributeValue(w + "fldLock", "true");
                    body.ReplaceNodes(
                        new XElement(w + "p", new XElement(w + "hyperlink", new XAttribute(w + "anchor", "destination"),
                            new XElement(w + "r", new XElement(w + "t", "JUMP TO BREAK DESTINATION")))),
                        new XElement(w + "bookmarkStart", new XAttribute(w + "id", 131), new XAttribute(w + "name", "destination")),
                        new XElement(w + "p", content),
                        new XElement(w + "bookmarkEnd", new XAttribute(w + "id", 131)),
                        section);
                });
                TestAssert.Equal(retainLeadingFragment ? 4 : 3, document.BodyElements.Count);
                if (wrapper is "empty-field" or "hint-field" or "locked-field" or "nested-hint-field")
                {
                    var field = document.Paragraphs.Last().FieldReferences.Single();
                    TestAssert.Equal(DocxFieldSourceKind.Simple, field.SourceKind);
                    TestAssert.Equal(wrapper == "locked-field" ? DocxFieldKind.Page : DocxFieldKind.Other, field.Kind);
                    TestAssert.Equal(wrapper == "locked-field" ? "PAGE" : "REF EmptyReference", field.Instruction);
                    TestAssert.True(!field.RendersCachedResult && !field.UsesPlaceholder,
                        "The empty stored result retains its field metadata on the continuation.");
                }
                if (wrapper is "closed-move-range" or "excluded-move-range")
                {
                    DocxRevisionRange range = document.Paragraphs.Last().RevisionRanges.Single();
                    TestAssert.Equal("79", range.Id);
                    TestAssert.Equal("PublicMove", range.Name);
                    TestAssert.True(!range.HasContentSource && range.StartSourceRunIndex is not null,
                        "The closed nonprinting range remains on the continuation at its original source boundary.");
                }
                if (includeDrawings)
                {
                    DocxFloatingDrawing drawing = document.FloatingDrawings.Single();
                    TestAssert.Equal("2", drawing.RelativeHeightValue);
                    TestAssert.Equal((int?)2, drawing.LayoutBlockIndex);
                }
                return renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
            }
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
