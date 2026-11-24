using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    private static readonly (string Id, string Control, string Neighbor)[] SimpleScriptCases =
    [
        ("minimum-simple-script-latin12", "ordinary-minimum-line-20.docx", "Alph\u00e1"),
        ("minimum-simple-script-greek12", "ordinary-minimum-line-20.docx", "\u0391\u03bb\u03c6\u03b1"),
        ("minimum-simple-script-cyrillic12", "ordinary-minimum-line-20.docx", "\u0410\u043b\u044c\u0444\u0430"),
        ("minimum-simple-script-latin11", "CompositionProbes/minimum-bookmark-ordinary11-control.docx", "Alph\u00e1"),
        ("minimum-simple-script-latin-noted12", "note-marked-minimum-line-20.docx", "Alph\u00e1")
    ];

    public static void DocxSimpleScriptNeighborsMatchIndependentWordMinimumFrames()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent simple-script controls use installed Windows fonts.");
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("minimum-simple-script-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string? id = row.GetProperty("Id").GetString();
            var fixture = SimpleScriptCases.Single(fixture => fixture.Id == id);
            string path = StaticNoteAdmissionCasePath(row.GetProperty("Input").GetString()!);
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var source = DocxTests.ReadDocx(path, mode);
            var pages = new DocxRenderer(null, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(row.GetProperty("Pages").GetInt32(), pages.Length);
            var shows = pages.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)).ToArray();
            foreach (var anchor in row.GetProperty("Anchors").EnumerateArray())
            {
                string text = anchor.GetProperty("Text").GetString()!;
                var show = shows.Single(show => show.Text.Contains(text, StringComparison.Ordinal));
                TestAssert.True(Math.Abs(show.Y - anchor.GetProperty("Y").GetDouble()) < .5d,
                    $"A supported simple-script neighbor preserves independent minimum frames: {id}/{mode}/{text}/{show.Y}.");
            }

            var ink = PortableRenderingTests.ReadPaintedSourceLabel(pages[0], fixture.Neighbor + " public body before note");
            double top = shows.Single(show => show.Text.StartsWith("Public minimum Top.", StringComparison.Ordinal)).Y;
            double bottom = shows.Single(show => show.Text.StartsWith("Public minimum AfterAlpha.", StringComparison.Ordinal)).Y;
            TestAssert.True(ink.All(glyph => glyph.Origin.Y < top && glyph.Origin.Y > bottom),
                $"The authored neighbor must paint between the two minimum paragraphs: {id}/{mode}.");
        }
    }

    public static void DocxSimpleScriptNeighborsKeepUnchangedParagraphGeometry()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (var fixtures in SimpleScriptCases.GroupBy(fixture => fixture.Control))
        {
            var source = DocxTests.ReadDocx(StaticNoteAdmissionCasePath(fixtures.Key), mode);
            string controlText = string.Concat(source.Paragraphs[1].Runs.Select(run => run.Text));
            var renderer = new DocxRenderer(new TestFaceFontResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var before = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            var expected = before.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true))
                .Where(show => show.Text.StartsWith("Public minimum ", StringComparison.Ordinal)).Select(show => (show.Text, show.X, show.Y)).ToArray();
            TestAssert.Equal(4, expected.Length);

            foreach (var fixture in fixtures)
            {
                var changed = DocxTests.ReadDocx(NestedMathCasePath(fixture.Id + ".docx"), mode);
                string changedText = string.Concat(changed.Paragraphs[1].Runs.Select(run => run.Text));
                TestAssert.True(changedText != controlText && changedText.StartsWith(fixture.Neighbor + " public body before note", StringComparison.Ordinal),
                    $"The reader must retain the authored neighbor: {fixture.Id}/{mode}.");
                var after = renderer.RenderPagesAsync(changed, null, default).ToBlockingEnumerable().ToArray();
                TestAssert.Equal(before.Length, after.Length);
                var actual = after.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true))
                    .Where(show => show.Text.StartsWith("Public minimum ", StringComparison.Ordinal)).Select(show => (show.Text, show.X, show.Y)).ToArray();
                TestAssert.True(expected.SequenceEqual(actual), $"Neighboring simple-script text cannot displace unchanged minimum paragraphs: {fixture.Id}/{mode}.");
            }
        }
    }

    public static void DocxSimpleScriptPlainFramesRetainShapingAndUnknownTextFallbacks()
    {
        var paragraph = DocxTests.ReadDocx(StaticNoteAdmissionCasePath("ordinary-minimum-line-20.docx"), OoxPdfDocxMarkupMode.Final).Paragraphs[1];
        bool Eligible(string text) => DocxPlainBodyFrame.HasPlainTextContent(paragraph with
            { Runs = [paragraph.Runs[0] with { Text = text }] });
        foreach (string text in new[] { "", "AV 123!\r\n", "\t", "AV\t123", "caf\u00e9 \u010c\u01ce", "\u1ea1 \u1f04", "\u0391\u03bb\u03c6\u03b1", "\u0410\u043b\u044c\u0444\u0430" })
            TestAssert.True(Eligible(text), "Qualified simple letters, printable ASCII and supported horizontal tabs retain plain frames.");
        foreach (string text in new[] { "e\u0301", "\u0483", "\u037e", "\u00d7", "\u00a0", "\u0000", "\u200f", "\u202e",
            "\u05d0", "\u0627", "\u0915", "\u0e01", "\u4e2d", "\U0001e900", "\U0001f600", "\ud800" })
            TestAssert.True(!Eligible(text), "Combining, bidi, complex/unknown scripts, controls and invalid scalars retain explicit fallback.");
        TestAssert.True(!DocxPlainBodyFrame.HasPlainTextContent(paragraph with
            { Runs = [paragraph.Runs[0] with { Text = "caf\u00e9", Revisions = [new(DocxRevisionKind.Insertion, "1", null, null, "ins", null, [])] }] }),
            "Simple text does not waive source revision ownership.");
    }
}
