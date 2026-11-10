using System.Text.RegularExpressions;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    public static void LayoutRelatedStoryLookupSeparatesCoincidentReviewBodyWidths()
    {
        foreach (bool keep in new[] { false, true })
        foreach (bool widerFirst in new[] { false, true })
        foreach (var geometry in new[] { OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin,
            OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
            CheckRelatedStoryWidthComposition(keep, widerFirst, geometry, pageField: false);
    }

    public static void LayoutRelatedStoryLookupPreservesUnreservedWidthsAndPageBinding()
    {
        foreach (bool keep in new[] { false, true })
        foreach (var geometry in new[] { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout,
            OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
            CheckRelatedStoryWidthComposition(keep, widerFirst: true, geometry, pageField: true);
    }

    private static void CheckRelatedStoryWidthComposition(bool keep, bool widerFirst,
        OoxPdfDocxMarkupGeometryMode geometry, bool pageField)
    {
        var measurer = LayoutAdmissionMeasurer();
        DocxLayout? earlier = null;
        string[]? expected = null;
        foreach (bool prior in new[] { false, true })
        {
            using var input = RelatedStoryCollisionInput(prior, keep, widerFirst, pageField);
            using var package = OoxPackage.Open(input, default);
            var source = new DocxReader().Read(package, null, default, OoxPdfDocxMarkupMode.Final);
            var layout = new DocxLayoutEngine(geometry, wordCompatiblePrintScale: 1d).Create(source, measurer, default);
            var note = layout.Pages.SelectMany((page, index) => page.PlacedRelatedStories
                .Where(note => note.StoryLayout.Story.IsNormalStoryType && note.StoryLayout.Story.Id == "2")
                .Select(note => (Page: page, Number: index + 1, Note: note))).Single();
            var groups = note.Note.TextLines.Select(line => string.Join(" ", Regex.Matches(line.Text, "NoteWord[0-9]{3}")
                .Select(match => match.Value))).Where(words => words.Length > 0).ToArray();
            TestAssert.Equal(80, groups.Sum(group => group.Split(' ').Length));
            double width = note.Page.Width - note.Page.MarginLeft - note.Page.MarginRight + note.Page.MarkupMarginReservePoints;
            TestAssert.True(note.Note.TextLines.All(line => line.Width <= width + .001d),
                "Each note must wrap to its own design width, independent of earlier sections.");
            if (pageField)
                TestAssert.True(string.Concat(note.Note.TextLines.Select(line => line.Text)).Contains("RefPage" + note.Number),
                    "Actual reference-page binding survives source-width cache reuse.");
            if (!prior) { expected = groups; earlier = layout; }
            else
            {
                TestAssert.True(expected!.SequenceEqual(groups),
                    $"Adding an earlier section cannot change later note wrapping ({geometry}, keep={keep}, widerFirst={widerFirst}).");
                if (geometry != OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                    TestAssert.Equal(1, layout.Pages.Select(page => page.Width - page.MarginLeft - page.MarginRight).Distinct().Count());
                TestAssert.True(earlier!.Pages.Count < layout.Pages.Count, "The earlier section genuinely adds a physical page.");
            }
        }
    }

    private static MemoryStream RelatedStoryCollisionInput(bool prior, bool keep, bool widerFirst, bool pageField)
    {
        const string w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        string Section(int right) => $"<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:right=\"{right}\"/></w:sectPr>";
        string Lead() => keep ? "<w:p><w:pPr><w:keepNext/></w:pPr><w:r><w:t>Lead.</w:t></w:r></w:p>" : "";
        string body = (prior ? Lead() + $"<w:p><w:pPr>{Section(widerFirst ? 1440 : 2880)}</w:pPr><w:r><w:t>Earlier body.</w:t></w:r><w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>" : "") +
            Lead() + "<w:p><w:r><w:t>Later body.</w:t></w:r><w:r><w:footnoteReference w:id=\"2\"/></w:r></w:p>" + Section(widerFirst ? 2880 : 1440);
        string words = string.Join(" ", Enumerable.Range(0, 80).Select(index => "NoteWord" + index.ToString("D3")));
        string field = pageField ? "<w:r><w:t> RefPage</w:t></w:r><w:fldSimple w:instr=\"PAGE\"><w:r><w:t>9</w:t></w:r></w:fldSimple>" : "";
        return TestFixtures.CreateZipPackage(new Dictionary<string, string> {
            ["[Content_Types].xml"] = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"xml\" ContentType=\"application/xml\"/><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>",
            ["word/document.xml"] = $"<w:document xmlns:w=\"{w}\"><w:body>{body}</w:body></w:document>",
            ["word/footnotes.xml"] = $"<w:footnotes xmlns:w=\"{w}\"><w:footnote w:id=\"1\"><w:p><w:r><w:footnoteRef/><w:t>Earlier note.</w:t></w:r></w:p></w:footnote><w:footnote w:id=\"2\"><w:p><w:r><w:footnoteRef/><w:t>{words}</w:t></w:r>{field}</w:p></w:footnote></w:footnotes>",
            ["word/_rels/document.xml.rels"] = $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"n\" Type=\"{rel}/footnotes\" Target=\"footnotes.xml\"/></Relationships>"
        });
    }
}
