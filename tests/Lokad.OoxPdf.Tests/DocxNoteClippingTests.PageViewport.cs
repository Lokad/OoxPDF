using System.Globalization;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxNoteClippingTests
{
    public static void DocxNoteTableCellsRetainPageContainedInkBeyondFlowAdmission()
    {
        // Independent Word source: the note row reaches the physical paper
        // edge. The former body-bottom clip hid30additional complete words.
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "NoteTables", "unfit-endnote-row.docx"));
        TestAssert.Equal("274A9989DAB970D7D39E0DC1230EC77BA014B13E4C7C1CE4A87DDCFDCE355ED0",
            Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = DocxTests.ReadDocx(path, mode);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                    : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            int marginGlyphs = 0;
            foreach (PdfPage page in pages)
            foreach (var (block, clip) in ReadPhysicalTextClips(page))
            foreach (var show in DocxCommentsTests.ReadEmbeddedGlyphTextShows(new PdfPage(page.Width, page.Height, block, page.Fonts)))
            {
                if (!show.Text.Contains("FlowToken", StringComparison.Ordinal)) continue;
                var font = show.Font.Font;
                double scale = show.Size / font.UnitsPerEm;
                foreach (char character in show.Text)
                {
                    if (char.IsWhiteSpace(character) || !font.TryGetGlyphBounds(font.MapCodePoint(character), out var ink)) continue;
                    double bottom = show.Y + ink.YMin * scale, top = show.Y + ink.YMax * scale;
                    // The unfit tail remains physically clipped by the page;
                    // only glyphs wholly on paper are required to survive.
                    if (bottom < 0d || top > page.Height || top >= 72d) continue;
                    marginGlyphs++;
                    TestAssert.True(bottom >= clip.Y - .01d && top <= clip.Y + clip.Height + .01d,
                        "A page-contained note-cell glyph cannot be lost to the story's flow-admission clip.");
                }
            }
            TestAssert.True(marginGlyphs > 0, "The witness exercises visible table ink below its72ptbody margin: " + mode);
            TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_CLIPPED_NOTE_TABLE_ROW"),
                "The remaining row tail beyond the paper still reports unsupported fragmentation.");
        }
    }

    // Read only generated q/Q, physical rectangle clips and BT blocks for this
    // fixture. CID decoding stays in the existing glyph-show reader. Reject a
    // transformed text/clip instead of silently pretending it is physical.
    private static IEnumerable<(string Block, PdfRectangle Clip)> ReadPhysicalTextClips(PdfPage page)
    {
        const string number = @"[-+\d.eE]+";
        string rectangle = number + @"(?: " + number + @"){3}";
        string pattern = @"(?m)^(?:(?<save>q)|(?<restore>Q)|(?<clip>" + rectangle + @") re W\*? n|(?<matrix>" + number + @"(?: " + number + @"){5}) cm|(?<text>BT))\r?$";
        var saved = new Stack<(PdfRectangle Clip, bool Transformed)>();
        var clip = new PdfRectangle(0d, 0d, page.Width, page.Height);
        bool transformed = false;
        foreach (Match operation in Regex.Matches(page.Content, pattern))
        {
            if (operation.Groups["save"].Success) saved.Push((clip, transformed));
            else if (operation.Groups["restore"].Success) (clip, transformed) = saved.Pop();
            else if (operation.Groups["matrix"].Success) transformed = true;
            else if (operation.Groups["clip"].Success)
            {
                TestAssert.True(!transformed, "This witness supplies physical rectangle clips.");
                double[] v = operation.Groups["clip"].Value.Split(' ').Select(x => double.Parse(x, CultureInfo.InvariantCulture)).ToArray();
                double x = Math.Max(clip.X, v[0]), y = Math.Max(clip.Y, v[1]);
                clip = new PdfRectangle(x, y, Math.Min(clip.X + clip.Width, v[0] + v[2]) - x,
                    Math.Min(clip.Y + clip.Height, v[1] + v[3]) - y);
            }
            else
            {
                TestAssert.True(!transformed, "This witness emits its text in physical coordinates.");
                int end = page.Content.IndexOf("\nET", operation.Index, StringComparison.Ordinal);
                TestAssert.True(end >= 0, "Every generated text object terminates.");
                yield return (page.Content[operation.Index..(end + 3)], clip);
            }
        }
        TestAssert.Equal(0, saved.Count);
    }
}
