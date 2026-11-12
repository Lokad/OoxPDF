using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxAutomaticNoteReferencesKeepDrawingSourceOrder()
    {
        foreach (bool endnote in new[] { false, true })
        foreach (bool pictureFirst in new[] { false, true })
        foreach (bool complex in new[] { false, true })
        {
            var document = ReadNotePictureBoundary(endnote, pictureFirst, complex, joined: true, OoxPdfDocxMarkupMode.Final);
            var paragraph = document.Paragraphs.Single(p => p.Images.Count != 0);
            var reference = paragraph.InlineReferences.Single();
            var marker = TestAssert.NotNull(reference.DisplayRun);
            var items = TestAssert.NotNull(paragraph.InlineContent).ToList();
            int mark = items.FindIndex(item => item is DocxInlineText text && ReferenceEquals(text.Run, marker));
            int picture = items.FindIndex(item => item is DocxInlinePicture);
            TestAssert.True(mark >= 0 && picture >= 0 && pictureFirst == (picture < mark),
                "An automatic note reference retains its actual source order at a zero-width text boundary.");
            TestAssert.Equal(pictureFirst ? 1 : 0, marker.SourceDrawingIndexInRun);
            TestAssert.Equal(complex ? 5 : 0, marker.SourceTextOffsetInRun);
            TestAssert.Equal(1, items.OfType<DocxInlineEquation>().Count());
        }
    }

    public static void DocxAutomaticNotePictureRunBoundariesPreserveVisibleContent()
    {
        foreach (bool endnote in new[] { false, true })
        foreach (bool pictureFirst in new[] { false, true })
        foreach (bool complex in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
            var renderer = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var before = renderer.RenderPagesAsync(ReadNotePictureBoundary(endnote, pictureFirst, complex, false, mode), null, default)
                .ToBlockingEnumerable().ToArray();
            var after = renderer.RenderPagesAsync(ReadNotePictureBoundary(endnote, pictureFirst, complex, true, mode), null, default)
                .ToBlockingEnumerable().ToArray();
            TestAssert.Equal(before.Length, after.Length);
            for (int page = 0; page < before.Length; page++)
            {
                var a = DocxCommentsTests.ReadEmbeddedGlyphTextShows(before[page], true).Select(g => (g.Text, g.X, g.Y, g.Size));
                var b = DocxCommentsTests.ReadEmbeddedGlyphTextShows(after[page], true).Select(g => (g.Text, g.X, g.Y, g.Size));
                TestAssert.True(a.SequenceEqual(b), "Joining a note reference and picture preserves every glyph and its geometry.");
                TestAssert.Equal(before[page].Images.Count, after[page].Images.Count);
                string Pictures(string content) => string.Join(";", Regex.Matches(content,
                    @"[-\d.]+ 0 0 [-\d.]+ [-\d.]+ [-\d.]+ cm\s+/Im\d+ Do").Select(match => match.Value));
                TestAssert.Equal(Pictures(before[page].Content), Pictures(after[page].Content));
                TestAssert.True(before[page].Annotations.SequenceEqual(after[page].Annotations),
                    "Automatic note destinations and hit areas follow the same painted marker.");
            }
        }
    }

    private static DocxDocument ReadNotePictureBoundary(bool endnote, bool pictureFirst, bool complex,
        bool joined, OoxPdfDocxMarkupMode mode) => ReadMutation(endnote ? "Composition/composed-static-cell-report" : "Composition/report",
        mode, false, false, (xml, w, math) =>
        {
            var body = xml.Root!.Element(w + "body")!;
            var section = new XElement(body.Element(w + "sectPr")!);
            var marker = new XElement(body.Descendants(w + (endnote ? "endnoteReference" : "footnoteReference")).Single());
            var picture = new XElement(body.Descendants(w + "drawing").Single());
            var equation = new XElement(body.Descendants(math + "oMath").First());
            var markRun = new XElement(w + "r", marker);
            var pictureRun = new XElement(w + "r", picture);
            var runs = new List<XElement>();
            if (complex) runs.Add(new XElement(w + "r",
                new XElement(w + "fldChar", new XAttribute(w + "fldCharType", "begin")), new XElement(w + "instrText", " REF NoteLiteral "),
                new XElement(w + "fldChar", new XAttribute(w + "fldCharType", "separate")),
                new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), "Lead "),
                new XElement(w + "fldChar", new XAttribute(w + "fldCharType", "end"))));
            runs.AddRange(pictureFirst ? [pictureRun, markRun] : new[] { markRun, pictureRun });
            runs.Add(new XElement(w + "r", new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), " tail ")));
            var paragraph = new XElement(w + "p");
            if (joined) paragraph.Add(new XElement(w + "r", runs.SelectMany(run => run.Elements()).Select(node => new XElement(node))));
            else paragraph.Add(runs);
            paragraph.Add(equation, new XElement(w + "r", new XElement(w + "t", " end.")));
            body.RemoveNodes();
            body.Add(paragraph, section);
        });
}
