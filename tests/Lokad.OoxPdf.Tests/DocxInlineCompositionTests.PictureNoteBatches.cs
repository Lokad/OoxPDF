using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionNotePicturesFollowSourceLinesThroughContinuation()
    {
        foreach (bool additional in new[] { false, true })
        {
            var source = ReadMutation("CompositionProbes/long-note-inline-picture-report", OoxPdfDocxMarkupMode.Final, false, false,
                (xml, w, _) =>
                {
                    if (!additional) return;
                    var note = xml.Root!.Elements(w + "footnote").Single(n => (string?)n.Attribute(w + "id") == "2");
                    var run = note.Descendants(w + "drawing").Single().Ancestors(w + "r").First();
                    var clone = new XElement(run);
                    XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
                    clone.Descendants(wp + "docPr").Single().SetAttributeValue("id", "502");
                    note.Elements(w + "p").Skip(1).First().Add(clone);
                }, partOverride: "word/footnotes.xml");
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default);
            var canonical = layout.RelatedStories.Single(story => story.Story.Kind == DocxRelatedStoryKind.Footnote && story.Story.Id == "2");
            var picture = canonical.InlineImages.Single(image => image.SourceParagraphIndex == 0);
            TestAssert.Equal(0, picture.SourceLineIndex);
            int owner = canonical.TextLines.ToList().FindIndex(line => line.SourceBlockIndex == picture.SourceBlockIndex &&
                line.SourceParagraphIndex == picture.SourceParagraphIndex && line.SourceLineIndex == picture.SourceLineIndex);
            TestAssert.Equal(0, owner);
            var tail = DocxLayoutEngine.NarrowStoryTextLinesForOffset(canonical, owner + 1);
            TestAssert.Equal(0, tail.TextLines[0].SourceLineIndex);
            TestAssert.True(tail.TextLines[0].SourceParagraphIndex != picture.SourceParagraphIndex,
                "Paragraph-local line indices restart; the complete source owner decides picture membership.");
            TestAssert.Equal(additional ? 1 : 0, tail.InlineImages.Count);
            TestAssert.Equal(additional ? 2 : 1, canonical.InlineImages.Count);
            TestAssert.True(Math.Abs(canonical.ContentHeight - tail.ContentHeight - picture.LineFrame!.Value.Height) < .000001d,
                "Consuming the picture's line charges its measured occupied height and after-space exactly once.");
            var placed = layout.Pages.SelectMany(page => page.PlacedRelatedStories.Where(story =>
                story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && story.StoryLayout.Story.Id == "2")).ToArray();
            TestAssert.Equal(additional ? 2 : 1, placed.Sum(part => part.InlineImages.Count));
            TestAssert.True(canonical.TextLines.Select(line => (line.Text, line.SourceParagraph, line.SourceLineIndex)).SequenceEqual(
                placed.SelectMany(part => part.StoryLayout.TextLines).Select(line => (line.Text, line.SourceParagraph, line.SourceLineIndex))),
                "Image-bearing note batches retain every canonical text line and source paragraph once.");
            foreach (int? recordedLine in new int?[] { null, 999 })
            {
                var legacy = canonical with { InlineImages = canonical.InlineImages.Select(image =>
                    ReferenceEquals(image, picture) ? image with { SourceLineIndex = recordedLine } : image).ToArray() };
                TestAssert.Equal(canonical.InlineImages.Count, DocxLayoutEngine.NarrowStoryTextLinesForOffset(legacy, owner + 1).InlineImages.Count);
            }
        }
    }

    public static void DocxInlineCompositionPictureNoteBatchesMatchIndependentWordFrames()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent picture-note frames require installed Windows fonts.");
        CheckPictureNoteBatches(null);
    }

    public static void DocxInlineCompositionWrappedPictureNoteFramesMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent picture-note frames require installed Windows fonts.");
        CheckContinuationPictureFrameControls(null);
    }

    internal static void CheckContinuationPictureFrameControls(IFontResolver? resolver)
    {
        foreach (string fixture in new[] { "wrapped-note-pictures-report", "wrapped-single-picture-note-report", "two-note-pictures-report" })
            CheckPictureNoteBatches(resolver, fixture);
    }

    internal static void CheckPictureNoteBatches(IFontResolver? resolver, string fixture = "long-note-inline-picture-report")
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath(fixture[..^"-report".Length] + "-office.json")));
        int expectedImages = gold.RootElement.GetProperty("ImagePlacements").GetInt32();
        int extraMath = gold.RootElement.TryGetProperty("ExtraMathematicalGlyphs", out var extra) ? extra.GetInt32() : 0;
        TestAssert.Equal(gold.RootElement.GetProperty("InputSha256").GetString(),
            Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(NestedMathCasePath(fixture + ".docx")))));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            List<(int Page, double X, double Y, double Width, double Height)>? originalPictures = null;
            foreach (bool transformed in new[] { false, true })
            {
                var document = ReadMutation("CompositionProbes/" + fixture, mode, false, transformed,
                    (xml, w, _) =>
                    {
                        if (!transformed) return;
                        int id = 901;
                        foreach (var run in xml.Descendants(w + "drawing").Select(drawing => drawing.Ancestors(w + "r").First()).ToArray())
                        {
                            run.AddBeforeSelf(new XElement(w + "bookmarkStart", new XAttribute(w + "id", id), new XAttribute(w + "name", "PublicNotePicture" + id)));
                            run.AddAfterSelf(new XElement(w + "bookmarkEnd", new XAttribute(w + "id", id++)));
                        }
                    }, partOverride: "word/footnotes.xml");
                var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
                var pages = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(document, diagnostics.Add, default).ToArray();
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "An inline picture cannot remove existing supported formulas.");
                var pictures = new List<(int Page, double X, double Y, double Width, double Height)>();
                var seen = new List<int>();
                var mathOrigins = new List<(int Page, string Text, double X, double Y)>();
                var noteOrigins = new List<(int Number, int Page, double Y)>();
                int heads = 0, math = 0, secondHeadPage = -1;
                double scale = final ? 1d : DocxRenderer.ResolveWordCompatiblePrintScale(document,
                    DocxMarkupContext.FromMode(mode, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup));
                for (int index = 0; index < pages.Length; index++)
                {
                    var page = pages[index];
                    foreach (var image in page.Images)
                    {
                        var matches = Regex.Matches(page.Content,
                            @"(?m)^(?<w>[-\d.]+) 0 0 (?<h>[-\d.]+) (?<x>[-\d.]+) (?<y>[-\d.]+) cm\s*/" + Regex.Escape(image.ResourceName) + @"\s+Do\b");
                        foreach (Match match in matches) pictures.Add((index + 1, Number(match, "x"), Number(match, "y"), Number(match, "w"), Number(match, "h")));
                    }
                    var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray();
                    var equationShows = shows.Where(show => page.Fonts.Any(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(resource.Font, show.Font))).ToArray();
                    math += equationShows.Length;
                    mathOrigins.AddRange(equationShows.Select(show => (index + 1, show.Text, show.X, show.Y)));
                    var lines = shows.GroupBy(show => show.Y).OrderByDescending(group => group.Key).Select(group => group.OrderBy(show => show.X).ToArray()).ToArray();
                    var text = lines.Select(line => string.Concat(line.Select(show => show.Text))).ToArray();
                    for (int line = 0; line < text.Length; line++)
                        if (Regex.Match(text[line], @"^Public continuation note line (\d\d):") is { Success: true } numbered)
                            noteOrigins.Add((int.Parse(numbered.Groups[1].Value, CultureInfo.InvariantCulture), index + 1, lines[line][0].Y));
                    int before = seen.Count;
                    foreach (Match match in Regex.Matches(string.Join(" ", text), @"Public\s+continuation\s+note\s+line\s+(\d\d):")) seen.Add(int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture));
                    var note = Enumerable.Range(0, lines.Length).Where(i => text[i].StartsWith("All quantities", StringComparison.Ordinal) ||
                        text[i].StartsWith("Second public continuation note.", StringComparison.Ordinal) || text[i].StartsWith("Public continuation note line", StringComparison.Ordinal)).ToArray();
                    if (note.Length == 0) continue;
                    var footer = lines[Array.FindIndex(text, t => t.StartsWith("Public synthetic composition portfolio", StringComparison.Ordinal))];
                    TestAssert.True(note.SelectMany(i => lines[i]).SelectMany(show => OccupiedFooterInk(show, scale)).Min() > footer.SelectMany(show => OccupiedFooterInk(show, scale)).Max(),
                        "Every picture-note paragraph clears actual footer ink.");
                    foreach (var expected in row.GetProperty("Heads").EnumerateArray())
                    {
                        int line = Array.FindIndex(text, t => t.StartsWith(expected.GetProperty("Prefix").GetString()!, StringComparison.Ordinal));
                        if (line < 0) continue;
                        heads++;
                        if (text[line].StartsWith("Second public", StringComparison.Ordinal)) secondHeadPage = index + 1;
                        var target = pages.SelectMany(p => p.Annotations).Single(link => link.Destination?.PageIndex == index).Destination!.Value;
                        if (resolver is null)
                        {
                            TestAssert.Equal(expected.GetProperty("Page").GetInt32(), index + 1);
                            TestAssert.True(Math.Abs(lines[line][0].Y - expected.GetProperty("Y").GetDouble()) < 1d &&
                                target.Top is { } top && Math.Abs(top - expected.GetProperty("DestinationY").GetDouble()) < 1d,
                                "Image-bearing heads and automatic targets retain independent one-point Word origins.");
                        }
                    }
                    if (resolver is null)
                    {
                        var expected = row.GetProperty("NotePages").EnumerateArray().Single(p => p.GetProperty("Page").GetInt32() == index + 1);
                        TestAssert.True(seen.Skip(before).SequenceEqual(expected.GetProperty("Numbers").EnumerateArray().Select(n => n.GetInt32())),
                            "Shared picture-note batches retain Word's independent source-paragraph page membership.");
                    }
                }
                TestAssert.Equal(2, heads);
                TestAssert.True(seen.SequenceEqual(Enumerable.Range(1, 64)), "Every original source paragraph survives once and in order.");
                TestAssert.Equal(22 + 5 * pages.Length + extraMath, math);
                TestAssert.Equal(2, pages.Sum(page => page.Annotations.Count(link => link.Destination is not null)));
                TestAssert.Equal(expectedImages, pictures.Count);
                TestAssert.Equal(secondHeadPage, pictures[1].Page);
                TestAssert.Equal(expectedImages, pages.Sum(page => page.Images.Sum(image => Regex.Matches(page.Content, "/" + Regex.Escape(image.ResourceName) + @"\s+Do\b").Count)));
                if (originalPictures is not null)
                    for (int i = 0; i < pictures.Count; i++)
                    {
                        var a = originalPictures[i]; var z = pictures[i];
                        TestAssert.True(a.Page == z.Page && Math.Abs(a.X - z.X) < .002d && Math.Abs(a.Y - z.Y) < .002d && Math.Abs(a.Width - z.Width) < .002d && Math.Abs(a.Height - z.Height) < .002d,
                            "Bookmarks and unchanged-format run splitting retain every picture's geometry.");
                    }
                originalPictures = pictures;
                if (resolver is not null) continue;
                if (row.TryGetProperty("NoteLineAnchors", out var noteAnchors))
                    foreach (var anchor in noteAnchors.EnumerateArray())
                        TestAssert.True(noteOrigins.Any(origin => origin.Number == anchor.GetProperty("Number").GetInt32() &&
                            origin.Page == anchor.GetProperty("Page").GetInt32() && Math.Abs(origin.Y - anchor.GetProperty("Y").GetDouble()) < 1d),
                            "Every numbered note paragraph retains its independent one-point Word baseline.");
                if (row.TryGetProperty("MathAnchors", out var anchors))
                    foreach (var anchor in anchors.EnumerateArray())
                        TestAssert.True(mathOrigins.Any(origin => origin.Page == anchor.GetProperty("Page").GetInt32() &&
                            origin.Text == anchor.GetProperty("Text").GetString() && Math.Abs(origin.X - anchor.GetProperty("X").GetDouble()) < 1d &&
                            Math.Abs(origin.Y - anchor.GetProperty("Y").GetDouble()) < 1d),
                            "Every added mathematical glyph retains its independent one-point Word origin beside the picture.");
                TestAssert.Equal(row.GetProperty("Pages").GetInt32(), pages.Length);
                var expectedPictures = row.GetProperty("Images").EnumerateArray().ToArray();
                for (int i = 0; i < pictures.Count; i++)
                {
                    var picture = pictures[i]; var expected = expectedPictures[i];
                    TestAssert.True(picture.Page == expected.GetProperty("Page").GetInt32() && Math.Abs(picture.X - expected.GetProperty("X").GetDouble()) < 1d &&
                        Math.Abs(picture.Y - expected.GetProperty("Y").GetDouble()) < 1d && Math.Abs(picture.Width - expected.GetProperty("Width").GetDouble()) < 1d &&
                        Math.Abs(picture.Height - expected.GetProperty("Height").GetDouble()) < 1d, "Both picture rectangles retain independent one-point Word bounds.");
                }
            }
        }
        static double Number(Match match, string group) => double.Parse(match.Groups[group].Value, CultureInfo.InvariantCulture);
    }
}
