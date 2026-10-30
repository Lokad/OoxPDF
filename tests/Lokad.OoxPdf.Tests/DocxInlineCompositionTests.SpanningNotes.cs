using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxSpanningParallelCellsKeepNotesWithIndependentWordReferences()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent spanning-note page ownership uses installed Windows fonts.");
        CheckSpanningCellNotes(null, qualifySourcePages: true);
    }

    public static void DocxSingleSpanningNoteKeepsIndependentWordSourcePages()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent single-note source pages use installed Windows fonts.");
        CheckSpanningCellNotes(null, "parallel-cell-single-footnote-office.json", qualifySourcePages: true);
    }

    public static void DocxFirstSpanningNoteKeepsIndependentWordSourcePages()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent first-fragment note pages use installed Windows fonts.");
        CheckSpanningCellNotes(null, "parallel-cell-first-footnote-office.json", qualifySourcePages: true);
    }

    public static void DocxCrossRowSpanningNotesKeepIndependentWordOwners()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent cross-row note pages use installed Windows fonts.");
        CheckSpanningCellNotes(null, "parallel-cell-cross-row-footnote-office.json", qualifySourcePages: true);
    }

    public static void DocxSpanningNoteAtChangingCutKeepsMarkerAndTargetTogether()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("The cut-position control uses installed Windows fonts.");
        var source = ReadMutation("CompositionProbes/parallel-cell-boxed-inset-notes", OoxPdfDocxMarkupMode.Final, false, false,
            (xml, w, _) =>
            {
                var run = xml.Descendants(w + "footnoteReference").Single(n => (string?)n.Attribute(w + "id") == "2").Parent!;
                var paragraph = xml.Descendants(w + "p").Single(p => string.Concat(p.Descendants(w + "t").Select(t => t.Value)) == "Right source row 107.");
                run.Remove(); paragraph.Add(run);
            });
        var renderer = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var text = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable()
            .Select(page => string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Select(show => show.Text))).ToArray();
        int marker = Array.FindIndex(text, value => value.Contains("Right source row 107.", StringComparison.Ordinal));
        int target = Array.FindIndex(text, value => value.Contains("Right final-page note:", StringComparison.Ordinal));
        TestAssert.True(marker >= 0 && marker == target, "A capacity refinement must preserve complete source marker/target page ownership.");
    }

    internal static void CheckSpanningCellNotes(IFontResolver? resolver,
        string goldenFile = "parallel-cell-footnote-office.json", bool qualifySourcePages = false)
    {
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var gold = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, goldenFile)));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        foreach (bool transformed in new[] { false, true })
        {
            string id = row.GetProperty("Id").GetString()!;
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(
                SHA256.HashData(File.ReadAllBytes(Path.Combine(folder, id + ".docx")))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var source = ReadMutation("CompositionProbes/" + id, mode, transformed, transformed);
            var renderer = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            var text = pages.Select(page => string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Select(show => show.Text))).ToArray();
            string all = string.Concat(text);
            if (resolver is null) TestAssert.Equal(row.GetProperty("WordPages").GetInt32(), pages.Length);
            foreach (var note in row.GetProperty("WordNotes").EnumerateArray())
            {
                string label = note.GetProperty("SourceLabel").GetString()!;
                string content = note.GetProperty("Text").GetString()!;
                var markers = text.Select((value, index) => (value, index)).Where(item => item.value.Contains(label, StringComparison.Ordinal)).ToArray();
                var targets = text.Select((value, index) => (value, index)).Where(item => item.value.Contains(content, StringComparison.Ordinal)).ToArray();
                TestAssert.Equal(1, markers.Length);
                TestAssert.Equal(1, targets.Length);
                TestAssert.Equal(markers[0].index, targets[0].index);
                if (resolver is null)
                {
                    TestAssert.Equal(note.GetProperty("ReferencePage").GetInt32(), markers[0].index + 1);
                    TestAssert.Equal(note.GetProperty("Page").GetInt32(), targets[0].index + 1);
                }
            }
            if (qualifySourcePages)
                foreach (var anchor in row.GetProperty("WordProseAnchors").EnumerateArray())
                {
                    string prefix = anchor.GetProperty("Text").GetString()!;
                    int[] owners = text.Select((value, index) => (value, index))
                        .Where(item => item.value.Contains(prefix, StringComparison.Ordinal)).Select(item => item.index + 1).ToArray();
                    TestAssert.Equal(1, owners.Length);
                    TestAssert.Equal(anchor.GetProperty("Page").GetInt32(), owners[0]);
                }
            foreach (string side in new[] { "Left", "Right" })
            {
                var ordinals = Regex.Matches(all, side + @" source row (\d+)\.").Select(match => int.Parse(match.Groups[1].Value)).Order().ToArray();
                TestAssert.True(ordinals.SequenceEqual(Enumerable.Range(1, 120)), "All source paragraphs survive exactly once beside notes and independent cell windows.");
                TestAssert.Equal(1, text.Count(value => value.Contains(side + " terminal control.", StringComparison.Ordinal)));
            }
            TestAssert.Equal(4, pages.Sum(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Count(show =>
                page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(font.Font, show.Font)))));
            TestAssert.Equal(1, pages.Sum(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Count(show =>
                show.Text.Trim() == pages.Length.ToString(System.Globalization.CultureInfo.InvariantCulture))));
            TestAssert.Equal(row.GetProperty("WordNotes").GetArrayLength(), pages.Sum(page => page.Annotations.Count));
            using var cancelled = new CancellationTokenSource();
            cancelled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(source, null, cancelled.Token).ToBlockingEnumerable().ToArray());
        }
    }
}
