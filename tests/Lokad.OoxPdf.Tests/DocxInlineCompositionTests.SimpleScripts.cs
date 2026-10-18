using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxSimpleScriptNeighborsMatchIndependentWordMinimumFrames()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent simple-script controls use installed Windows fonts.");
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("minimum-simple-script-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string path = StaticNoteAdmissionCasePath(row.GetProperty("Input").GetString()!);
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var source = DocxTests.ReadDocx(path, mode);
            var pages = new DocxRenderer(null, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(source, null, default).ToArray();
            TestAssert.Equal(row.GetProperty("Pages").GetInt32(), pages.Length);
            var shows = pages.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)).ToArray();
            foreach (var anchor in row.GetProperty("Anchors").EnumerateArray())
            {
                string text = anchor.GetProperty("Text").GetString()!;
                var show = shows.Single(show => show.Text.Contains(text, StringComparison.Ordinal));
                TestAssert.True(Math.Abs(show.Y - anchor.GetProperty("Y").GetDouble()) < .5d,
                    $"A supported simple-script neighbor preserves independent minimum frames: {row.GetProperty("Id").GetString()}/{mode}/{text}/{show.Y}.");
            }
        }
    }

    public static void DocxSimpleScriptNeighborsKeepUnchangedParagraphGeometry()
    {
        foreach (string variant in new[] { "latin12", "greek12", "cyrillic12", "latin11", "latin-noted12" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string control = variant == "latin11" ? "CompositionProbes/minimum-bookmark-ordinary11-control.docx"
                : variant == "latin-noted12" ? "note-marked-minimum-line-20.docx" : "ordinary-minimum-line-20.docx";
            var source = DocxTests.ReadDocx(StaticNoteAdmissionCasePath(control), mode);
            var changed = DocxTests.ReadDocx(NestedMathCasePath("minimum-simple-script-" + variant + ".docx"), mode);
            TestAssert.True(string.Concat(source.Paragraphs[1].Runs.Select(run => run.Text)) !=
                string.Concat(changed.Paragraphs[1].Runs.Select(run => run.Text)), "The neighboring text must actually change.");
            var renderer = new DocxRenderer(new TestFaceFontResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var before = renderer.RenderBlankPages(source, null, default).ToArray();
            var after = renderer.RenderBlankPages(changed, null, default).ToArray();
            TestAssert.Equal(before.Length, after.Length);
            var expected = before.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true))
                .Where(show => show.Text.StartsWith("Public minimum ", StringComparison.Ordinal)).Select(show => (show.Text, show.X, show.Y)).ToArray();
            var actual = after.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true))
                .Where(show => show.Text.StartsWith("Public minimum ", StringComparison.Ordinal)).Select(show => (show.Text, show.X, show.Y)).ToArray();
            TestAssert.Equal(4, expected.Length);
            TestAssert.True(expected.SequenceEqual(actual), "Neighboring simple-script text cannot displace unchanged minimum paragraphs.");
        }
    }

    public static void DocxSimpleScriptPlainFramesRetainShapingAndUnknownTextFallbacks()
    {
        var paragraph = DocxTests.ReadDocx(StaticNoteAdmissionCasePath("ordinary-minimum-line-20.docx"), OoxPdfDocxMarkupMode.Final).Paragraphs[1];
        bool Eligible(string text) => DocxPlainBodyFrame.HasPlainTextContent(paragraph with
            { Runs = [paragraph.Runs[0] with { Text = text }] });
        foreach (string text in new[] { "", "AV 123!\r\n", "caf\u00e9 \u010c\u01ce", "\u1ea1 \u1f04", "\u0391\u03bb\u03c6\u03b1", "\u0410\u043b\u044c\u0444\u0430" })
            TestAssert.True(Eligible(text), "Qualified simple letters and existing printable ASCII retain plain frames.");
        foreach (string text in new[] { "e\u0301", "\u0483", "\u037e", "\u00d7", "\u00a0", "\t", "\u0000", "\u200f", "\u202e",
            "\u05d0", "\u0627", "\u0915", "\u0e01", "\u4e2d", "\U0001e900", "\U0001f600", "\ud800" })
            TestAssert.True(!Eligible(text), "Combining, bidi, complex/unknown scripts, controls and invalid scalars retain explicit fallback.");
        TestAssert.True(!DocxPlainBodyFrame.HasPlainTextContent(paragraph with
            { Runs = [paragraph.Runs[0] with { Text = "caf\u00e9", Revisions = [new(DocxRevisionKind.Insertion, "1", null, null, "ins", null, [])] }] }),
            "Simple text does not waive source revision ownership.");
    }
}
