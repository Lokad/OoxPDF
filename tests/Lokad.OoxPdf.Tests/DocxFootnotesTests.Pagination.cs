using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxFootnotesTests
{

    public static void DocxScaledFootnoteExactAdvanceMatchesIndependentOffice()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("The Word control uses installed Arial metrics.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        string input = Path.Combine(cases, "scaled-note-body.docx");
        using var reference = System.Text.Json.JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "scaled-note-office-advances.json")));
        TestAssert.Equal(reference.RootElement.GetProperty("InputSha256").GetString(),
            Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(input))));
        var document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);
        var emission = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .InspectTextEmission(document).Lines.Where(line => line.ContainerStoryKind == "Body" && line.SourceBlockIndex == 1).ToArray();
        TestAssert.Equal(3, emission.Length);
        var gaps = reference.RootElement.GetProperty("FirstParagraphBaselineGapsPoints");
        for (int index = 0; index < 2; index++)
            TestAssert.True(Math.Abs(emission[index].Segments[0].BaselineY - emission[index + 1].Segments[0].BaselineY - gaps[index].GetDouble()) < .05d,
                "An automatic superscript must not change the independently measured printed exact-line advance.");
    }

    public static void DocxScaledFootnoteAdmissionMatchesIndependentPageMembership()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("The Office controls use installed Arial metrics.");
        foreach (string container in new[] { "body", "cell" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxTests.ReadDocx(Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
                "CompositionProbes", "scaled-note-" + container + ".docx")), mode);
            var pages = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.True(pages.Length == 1, $"{container} / {mode}: expected one page, got {pages.Length}.");
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText));
            for (int section = 1; section <= 6; section++)
            for (int line = 1; line <= 3; line++)
                TestAssert.True(text.Contains($"Section {section} line {line}.", StringComparison.Ordinal), "Body admission preserves all source lines.");
            TestAssert.True(text.Contains("Public calculation note.", StringComparison.Ordinal), "Note content remains beside its reference page.");
        }
    }

    public static void DocxAtomicFootnoteHeadIgnoresFutureContinuationRule()
    {
        var story = new DocxRelatedStory(DocxRelatedStoryKind.Footnote, "/word/footnotes.xml", "1", [], [], [], null);
        var run = new DocxTextRun("Note", 10d, null, false, false, false, null, null);
        var line = new DocxTextLineLayout("Note", run, 10d, 0d, -10d, 100d, [], null, null, null, null, 12d,
            null, null, false, null, null, null, null, null, null, null, null, null, null, null, null, false);
        var layout = new DocxRelatedStoryLayout(story, 0, [line], [], [], [], 12d);
        var separator = new DocxRelatedStoryLayout(story, 0, [line with { Text = " " }], [], [], [], 8d);
        var location = new DocxLayoutEngine.DocxInlineReferenceLocation(0,
            DocxTests.CreateDocxLayoutParagraph("Anchor", 10d, 12d),
            new DocxInlineReference(DocxRelatedStoryKind.Footnote, "1", null, "1", 0, 0, 0));
        var head = new DocxLayoutEngine.InFlightRelatedStory(layout, separator, separator, 0, location);
        var noFutureRule = new DocxLayoutEngine.InFlightRelatedStory(layout, separator, null, 0, location);
        double headHeight = DocxLayoutEngine.FootnoteRemainderHeight(head, [12d], null);
        TestAssert.Equal(DocxLayoutEngine.FootnoteRemainderHeight(noFutureRule, [12d], null), headHeight);
        head.SeparatorPlaced = true;
        TestAssert.True(DocxLayoutEngine.FootnoteRemainderHeight(head, [12d], null) > 12d,
            "A continued one-line tail must still charge its present continuation rule.");
        var split = new DocxLayoutEngine.InFlightRelatedStory(layout with { TextLines = [line, line] }, separator, separator, 0, location);
        TestAssert.True(DocxLayoutEngine.FootnoteRemainderHeight(split, [12d, 12d], null) > headHeight + 12d,
            "Potentially split multi-line notes retain their qualified conservative reserve.");
    }

    public static void DocxAtomicFootnotesMatchIndependentPageAdmission()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("The Office controls use installed Arial metrics.");
        foreach (string container in new[] { "body", "cell" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxTests.ReadDocx(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes",
                "complete-note-" + container + ".docx"), mode);
            var pages = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(1, pages.Length);
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText));
            for (int section = 1; section <= 6; section++)
            for (int line = 1; line <= 3; line++)
                TestAssert.True(text.Contains($"Section {section} line {line}.", StringComparison.Ordinal), "Every admitted line remains visible in the PDF stream.");
            TestAssert.True(text.Contains("Public calculation note.", StringComparison.Ordinal), "The complete note survives page admission.");
            TestAssert.Equal(1, pages.Single().Annotations.Count(annotation => annotation.Destination is not null));
        }
    }

    public static void DocxSimpleSectionPaginationUsesFollowingStartType()
    {
        foreach ((string? preceding, string? following, int expected) in new[]
        {
            ((string?)"nextPage", (string?)"continuous", 1),
            ((string?)"continuous", (string?)"nextPage", 2),
            ((string?)null, (string?)"continuous", 1),
            ((string?)"continuous", (string?)null, 2)
        })
        {
            DocxDocument document = ReadSectionNumberingFixture("footnote", (xml, _, parts) =>
            {
                SetSectionTypes(xml, preceding, following);
                SetNotePresence(xml, parts, "footnote", false, false);
            });
            TestAssert.Equal(expected, RenderSectionPages(document).Length);
            TestAssert.Equal(preceding, document.BodyElements.OfType<DocxSectionBreakElement>().Single().TypeValue?.ToValueString());
            TestAssert.Equal(following, document.FinalSectionBreak!.TypeValue?.ToValueString());
        }
    }

    public static void DocxSimpleSectionPaginationSeparatesPageBottomFootnotes()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach ((bool first, bool second) in new[] { (true, true), (true, false), (false, true), (false, false) })
        foreach (string? position in kind == "footnote" ? new string?[] { "pageBottom", "beneathText", null } : ["docEnd"])
        {
            DocxDocument document = ReadSectionNumberingFixture(kind, (xml, _, parts) =>
            {
                SetSectionTypes(xml, "continuous", "continuous");
                SetNotePresence(xml, parts, kind, first, second);
                foreach (XElement section in xml.Descendants(NoteWord + "sectPr"))
                {
                    XElement props = section.Element(NoteWord + kind + "Pr")!;
                    props.Elements(NoteWord + "pos").Remove();
                    if (position is not null) props.AddFirst(new XElement(NoteWord + "pos", new XAttribute(NoteWord + "val", position)));
                }
            });
            PdfPage[] pages = RenderSectionPages(document);
            bool splits = kind == "footnote" && first && position != "beneathText";
            TestAssert.Equal(splits ? 2 : 1, pages.Length);
            PdfLinkAnnotation[] links = pages.SelectMany(p => p.Annotations).ToArray();
            TestAssert.Equal((first ? 2 : 0) + (second ? 2 : 0), links.Length);
            TestAssert.True(links.All(l => l.IsDestination && l.Width > 0 && l.Height > 0), "Visible note markers remain linked.");
            if (splits && second)
            {
                TestAssert.Equal(2, pages[0].Annotations.Count);
                TestAssert.Equal(2, pages[1].Annotations.Count);
                TestAssert.True(pages[0].Annotations.All(l => l.Destination!.Value.PageIndex == 0) &&
                    pages[1].Annotations.All(l => l.Destination!.Value.PageIndex == 1), "Notes follow the section's source page.");
            }
        }
    }

    public static void DocxSimpleSectionPaginationIgnoresNotesAlreadyPlacedOnEarlierPages()
    {
        DocxDocument document = ReadSectionNumberingFixture("footnote", (xml, _, parts) =>
        {
            SetSectionTypes(xml, "continuous", "continuous");
            XElement[] paragraphs = xml.Root!.Element(NoteWord + "body")!.Elements(NoteWord + "p").ToArray();
            paragraphs[0].Add(new XElement(NoteWord + "r", new XElement(NoteWord + "br", new XAttribute(NoteWord + "type", "page"))));
            paragraphs[1].Descendants(NoteWord + "footnoteReference").Single().Parent!.Remove();
            XDocument notes = XDocument.Parse(parts["word/footnotes.xml"]);
            notes.Root!.Elements(NoteWord + "footnote").Single(n => (string?)n.Attribute(NoteWord + "id") == "4").Remove();
            parts["word/footnotes.xml"] = notes.ToString();
        });
        PdfPage[] pages = RenderSectionPages(document);
        TestAssert.Equal(2, pages.Length);
        TestAssert.Equal(1, pages[0].Annotations.Count);
        TestAssert.Equal(2, pages[1].Annotations.Count);
        TestAssert.True(pages[1].Annotations.All(l => l.Destination!.Value.PageIndex == 1), "An earlier note must not insert a third page.");
    }

    public static void DocxSimpleSectionPaginationRetainsExcludedFlow()
    {
        int variant = 0;
        foreach (Action<XDocument, XDocument, Dictionary<string, string>> modify in new Action<XDocument, XDocument, Dictionary<string, string>>[]
        {
            (xml, _, _) =>
            {
                XElement body = xml.Root!.Element(NoteWord + "body")!;
                XElement paragraph = body.Elements(NoteWord + "p").First();
                paragraph.Remove();
                body.AddFirst(new XElement(NoteWord + "tbl", new XElement(NoteWord + "tr", new XElement(NoteWord + "tc", paragraph))));
            },
            (xml, _, _) =>
            {
                foreach (XElement section in xml.Descendants(NoteWord + "sectPr"))
                    section.Add(new XElement(NoteWord + "cols", new XAttribute(NoteWord + "num", "2"), new XAttribute(NoteWord + "space", "720")));
            },
            (xml, _, _) => xml.Descendants(NoteWord + "sectPr").Last().Element(NoteWord + "pgMar")!.SetAttributeValue(NoteWord + "left", "1800"),
            (xml, _, _) => xml.Root!.Element(NoteWord + "body")!.Elements(NoteWord + "p").Skip(2).Remove(),
            (xml, _, _) => xml.Root!.Element(NoteWord + "body")!.Elements(NoteWord + "p").First().Add(
                new XElement(NoteWord + "ins", new XElement(NoteWord + "r", new XElement(NoteWord + "t", "tracked"))))
        })
        {
            DocxDocument document = ReadSectionNumberingFixture("footnote", (xml, settings, parts) =>
            {
                SetSectionTypes(xml, "continuous", "continuous");
                modify(xml, settings, parts);
            });
            PdfPage[] pages = RenderSectionPages(document);
            // Atomic heads no longer reserve unused future continuation rules.
            // Column-flow fallback must remain; the old second page was capacity
            // overhead, rather than an independently qualified section boundary.
            TestAssert.True(pages.Length == 1, $"Excluded variant {variant} must retain its admitted source page; got {pages.Length}.");
            if (variant == 1)
            {
                var layout = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup,
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).InspectLayout(document);
                TestAssert.True(layout.Pages.All(page => page.ColumnFrameCount == 2), "The excluded case must retain two-column placement.");
                TestAssert.Equal(4, pages.Single().Annotations.Count);
                TestAssert.True(pages.Single().Annotations.All(link => link.Destination is { PageIndex: 0 }),
                    "All four source markers retain note destinations on their admitted page.");
            }
            variant++;
        }
        DocxDocument ordinary = ReadSectionNumberingFixture("footnote", (xml, _, _) => SetSectionTypes(xml, "continuous", "continuous"));
        var scaled = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, wordCompatiblePrintScale: 0.5d);
        TestAssert.Equal(1, scaled.Create(ordinary, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None).Pages.Count);
    }

    public static void DocxSimpleSectionPaginationPreservesLabelsAndRepeatedDestinations()
    {
        foreach (OoxPdfDocxMarkupGeometryMode geometry in new[]
            { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
        {
            DocxDocument document = ReadSectionNumberingFixture("footnote", (xml, _, _) =>
            {
                SetSectionTypes(xml, "continuous", "continuous");
                SetNoteSection(xml.Descendants(NoteWord + "sectPr").Last(), "footnote", "lowerRoman", "4", "eachSect");
            });
            CheckNoteLabels(document, ["iv", "v", "iv", "v"]);
            PdfPage[] pages = RenderSectionPages(document, geometry);
            PdfLinkAnnotation[] links = pages.SelectMany(p => p.Annotations).ToArray();
            TestAssert.Equal(2, pages.Length);
            TestAssert.Equal(4, links.Length);
            TestAssert.Equal(4, links.Select(l => l.Destination).Distinct().Count());
            TestAssert.Equal(105, document.Settings.FootnoteReferenceSettings.NumberStart!.Value);
            TestAssert.Equal(DocxSectionBreakType.Continuous, document.BodyElements.OfType<DocxSectionBreakElement>().Single().TypeValue!.Value);
        }
    }

    private static PdfPage[] RenderSectionPages(DocxDocument document,
        OoxPdfDocxMarkupGeometryMode geometry = OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup) =>
        new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, geometry)
            .RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().ToArray();

    private static void SetSectionTypes(XDocument xml, string? preceding, string? following)
    {
        XElement[] sections = xml.Descendants(NoteWord + "sectPr").ToArray();
        for (int i = 0; i < sections.Length; i++)
        {
            sections[i].Elements(NoteWord + "type").Remove();
            string? value = i == 0 ? preceding : following;
            if (value is not null) sections[i].Add(new XElement(NoteWord + "type", new XAttribute(NoteWord + "val", value)));
        }
    }

    private static void SetNotePresence(XDocument xml, Dictionary<string, string> parts, string kind, bool first, bool second)
    {
        XElement[] references = xml.Descendants(NoteWord + kind + "Reference").ToArray();
        XDocument notes = XDocument.Parse(parts["word/" + kind + "s.xml"]);
        for (int i = 0; i < references.Length; i++)
        {
            if (i < 2 ? first : second) continue;
            string id = (string)references[i].Attribute(NoteWord + "id")!;
            references[i].Parent!.Remove();
            notes.Root!.Elements(NoteWord + kind).Single(n => (string?)n.Attribute(NoteWord + "id") == id).Remove();
        }
        parts["word/" + kind + "s.xml"] = notes.ToString();
    }
}
