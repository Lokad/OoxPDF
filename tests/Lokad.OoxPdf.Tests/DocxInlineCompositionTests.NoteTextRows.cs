using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxNoteTextRowsRetainLocalMetricsContentAndSourceOwners()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string content in new[] { "plain", "picture", "math", "mixed" })
        {
            var control = RenderTextRowMetrics(content, "ordinary", mode);
            foreach (bool endnote in new[] { false, true })
            foreach (bool sameRun in new[] { false, true })
            {
                var source = ReadNoteTextRows(content, "ordinary", mode, endnote, sameRun);
                var paragraph = source.Paragraphs[0]; var reference = paragraph.InlineReferences.Single();
                string label = endnote ? "i" : "1";
                TestAssert.Equal(label, reference.DisplayText);
                var marker = TestAssert.NotNull(reference.DisplayRun);
                TestAssert.True(paragraph.Runs.Any(run => ReferenceEquals(run, marker)), "The actual marker owns its generated run.");
                if (sameRun) TestAssert.True(paragraph.Runs.Any(run => !ReferenceEquals(run, marker) &&
                    run.SourceRunIndex == marker.SourceRunIndex && run.Text.Contains("Small row 01", StringComparison.Ordinal)),
                    "The automatic marker and ordinary text share an OOXML source run but have separate resolved owners.");
                var actual = RenderTextRowMetricDocument(source, mode);
                TestAssert.Equal(1, actual.Length);
                AssertTextRowOwnersAndPitches(control, actual);
                string native = string.Concat(actual.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page,
                    true, includeTextTransforms: true, honorActualText: false,
                    fontResourceFilter: name => !name.StartsWith("Eq", StringComparison.Ordinal))).Select(show => show.Text));
                for (int row = 1; row <= 7; row++) TestAssert.Equal(1, Regex.Matches(native, $"Small row {row:00}").Count);
                TestAssert.Equal(1, Regex.Matches(native, "Source note").Count);
                string expected = "Largeheading" + label +
                    string.Concat(Enumerable.Range(1, 7).Select(row => $"Smallrow{row:00}")) + label + "Sourcenote";
                TestAssert.Equal(expected, string.Concat(native.Normalize(System.Text.NormalizationForm.FormKC).Where(ch => !char.IsWhiteSpace(ch))));
                TestAssert.Equal(content is "picture" or "mixed" ? 1 : 0,
                    actual.Sum(page => Regex.Matches(page.Content, @"/\w+\s+Do\b").Count));
                string math = string.Concat(actual.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page,
                    true, includeTextTransforms: true, honorActualText: false).Where(show => page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) &&
                        ReferenceEquals(font.Font, show.Font)))).Select(show => show.Text)).Normalize(System.Text.NormalizationForm.FormKC);
                TestAssert.Equal(content is "math" or "mixed" ? "x" : "", math);
                if (content is "math" or "mixed")
                {
                    var shows = actual.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page,
                        true, includeTextTransforms: true, honorActualText: false)).ToArray();
                    var markerShow = shows.First(show => show.Text == label);
                    var mathShow = actual.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page,
                        true, includeTextTransforms: true, honorActualText: false,
                        fontResourceFilter: name => name.StartsWith("Eq", StringComparison.Ordinal))).Single();
                    TestAssert.True(mathShow.X < markerShow.X,
                        "The equation remains before its note marker in visible source order, independently of PDF paint-operation order.");
                }
                TestAssert.True(actual.Sum(page => page.Annotations.Count(annotation => annotation.Destination is not null)) >= 1,
                    "The actual note marker retains its painted navigation.");
                TestAssert.True(ReferenceEquals(reference, paragraph.InlineReferences.Single()) && ReferenceEquals(marker, reference.DisplayRun),
                    "Measurement does not replace original note metadata or its source run.");
            }
        }
    }

    public static void DocxNoteTextRowsShareNoteReservationKeepAndWidowAdmission()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string admission in new[] { "keepLines", "keepNext", "widow" })
        {
            var control = RenderTextRowMetricDocument(ReadNoteTextRows("plain", admission, mode, false, false), mode);
            var actual = RenderTextRowMetricDocument(ReadNoteTextRows("picture", admission, mode, false, true), mode);
            TestAssert.Equal(2, actual.Length);
            AssertTextRowOwnersAndPitches(control, actual);
            int headingPage = Array.FindIndex(actual, page => DocxPageFootnoteNumberingTests.DecodeText(page).Contains("Large heading", StringComparison.Ordinal));
            int notePage = Array.FindIndex(actual, page => DocxPageFootnoteNumberingTests.DecodeText(page).Contains("Source note", StringComparison.Ordinal));
            TestAssert.Equal(admission == "widow" ? 0 : 1, headingPage);
            TestAssert.Equal(headingPage, notePage);
            if (admission == "keepNext") TestAssert.Contains("Afterword", DocxPageFootnoteNumberingTests.DecodeText(actual[headingPage]));
        }
    }

    public static void DocxNoteTextRowsKeepMarkerAndUnknownOwnerFrameFallbacks()
    {
        foreach (string variant in new[] { "ordinary", "exact", "unknown-source", "missing-run" })
        {
            var original = ReadNoteTextRows("picture", "ordinary", OoxPdfDocxMarkupMode.Final, false, true);
            var paragraph = original.Paragraphs[0]; var reference = paragraph.InlineReferences.Single();
            if (variant == "exact") paragraph = paragraph with { LineSpacingPoints = 12d,
                Spacing = paragraph.Spacing with { LineValue = "240", LineRuleValue = "exact" } };
            if (variant is "unknown-source" or "missing-run") paragraph = paragraph with { InlineReferences =
                [variant == "unknown-source" ? reference with { SourceRunIndex = -1 } : reference with { DisplayRun = null }] };
            var source = original with { BodyElements = [new DocxParagraphElement(paragraph)] };
            var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
            var plan = DocxFontPlan.Create(source, resolver, default);
            var measurer = DocxFontPlanTextMeasurer.CreateAsync(plan, resolver.Resolve(new FontRequest("DejaVu Sans", false, false)),
                default, resolver).GetAwaiter().GetResult();
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(source, measurer, default);
            var lines = layout.Pages.SelectMany(page => page.Items).OfType<DocxTextLineLayout>().ToArray();
            var rows = lines.Where(line => line.Text.StartsWith("Small row", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(7, rows.Length);
            TestAssert.True(lines.First(line => line.Text.Contains("Large heading", StringComparison.Ordinal)).FontSize == 36d,
                "The marker-bearing first row retains its paragraph frame.");
            if (variant == "ordinary") TestAssert.True(rows.All(row => row.FontSize == 11d), "Only proven unmarked rows take local metrics.");
            else if (variant == "exact") TestAssert.True(rows.All(row => row.LineHeight == 12d), "An authored exact slot remains exact.");
            else TestAssert.True(rows.All(row => row.FontSize == 36d), "Missing or ambiguous marker ownership retains the conservative paragraph frame.");
        }
    }

    private static DocxDocument ReadNoteTextRows(string content, string admission, OoxPdfDocxMarkupMode mode, bool endnote, bool sameRun,
        int markerRow = 0, bool beforeText = false, bool splitRows = false) =>
        ReadTextRowMetrics(content, admission, mode, (xml, w, _) =>
        {
            var paragraph = xml.Descendants(w + "p").Single(p => p.Descendants(w + "t").Any(text => text.Value == "Large heading"));
            var firstBreak = paragraph.Elements(w + "r").First(run => run.Element(w + "br") is not null);
            var small = firstBreak.ElementsAfterSelf(w + "r").First();
            var properties = new XElement(small.Element(w + "rPr")!);
            if (!sameRun) properties.Element(w + "sz")!.SetAttributeValue(w + "val", "16");
            var marker = new XElement(w + "r", properties,
                new XElement(w + (endnote ? "endnoteReference" : "footnoteReference"), new XAttribute(w + "id", "1")));
            if (markerRow == 0)
            {
                firstBreak.AddBeforeSelf(marker);
                if (sameRun) { marker.Add(new XElement(w + "br"), new XElement(small.Element(w + "t")!)); firstBreak.Remove(); small.Remove(); }
            }
            else
            {
                var text = paragraph.Elements(w + "r").Elements(w + "t").Single(text => text.Value == $"Small row {markerRow:00}");
                if (sameRun)
                {
                    if (beforeText) text.AddBeforeSelf(marker.Elements().Single(child => child.Name != w + "rPr"));
                    else text.AddAfterSelf(marker.Elements().Single(child => child.Name != w + "rPr"));
                }
                else if (beforeText) text.Parent!.AddBeforeSelf(marker);
                else text.Parent!.AddAfterSelf(marker);
            }
            if (splitRows)
            {
                var propertiesTemplate = paragraph.Element(w + "pPr")!;
                var rows = new List<XElement>(); var row = new XElement(w + "p", new XElement(propertiesTemplate));
                foreach (var child in paragraph.Elements().Where(child => child.Name != w + "pPr"))
                {
                    if (child.Name == w + "r" && child.Element(w + "br") is not null)
                    { rows.Add(row); row = new XElement(w + "p", new XElement(propertiesTemplate)); }
                    else row.Add(new XElement(child));
                }
                rows.Add(row); paragraph.ReplaceWith(rows);
            }
        }, zip =>
        {
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var entry = zip.GetEntry("word/footnotes.xml")!; XDocument notes;
            using (var stream = entry.Open()) notes = XDocument.Load(stream);
            var note = notes.Root!.Elements(w + "footnote").Single(item => (string?)item.Attribute(w + "id") == "1");
            XElement Run(XElement child) => new(w + "r", new XElement(w + "rPr",
                new XElement(w + "rFonts", new XAttribute(w + "ascii", "DejaVu Sans"), new XAttribute(w + "hAnsi", "DejaVu Sans")),
                new XElement(w + "sz", new XAttribute(w + "val", "16"))), child);
            note.ReplaceNodes(new XElement(w + "p", new XElement(w + "pPr", new XElement(w + "spacing", new XAttribute(w + "before", "0"),
                new XAttribute(w + "after", "0"), new XAttribute(w + "line", "240"), new XAttribute(w + "lineRule", "auto"))),
                Run(new XElement(w + "footnoteRef")), Run(new XElement(w + "t", "Source note"))));
            if (endnote)
            {
                foreach (var node in notes.Descendants().Where(node => node.Name == w + "footnotes" || node.Name == w + "footnote" || node.Name == w + "footnoteRef"))
                    node.Name = w + node.Name.LocalName.Replace("footnote", "endnote", StringComparison.Ordinal);
                foreach (string part in new[] { "word/_rels/document.xml.rels", "[Content_Types].xml" })
                {
                    var item = zip.GetEntry(part)!; XDocument xml;
                    using (var stream = item.Open()) xml = XDocument.Load(stream);
                    foreach (var attribute in xml.Descendants().Attributes().Where(attribute => attribute.Value.Contains("footnotes", StringComparison.Ordinal)))
                        attribute.Value = attribute.Value.Replace("footnotes", "endnotes", StringComparison.Ordinal);
                    item.Delete(); using var output = zip.CreateEntry(part).Open(); xml.Save(output);
                }
            }
            entry.Delete(); using var target = zip.CreateEntry(endnote ? "word/endnotes.xml" : "word/footnotes.xml").Open(); notes.Save(target);
        });
}
