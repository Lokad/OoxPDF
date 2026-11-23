using System.Globalization;
using System.Numerics;
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
            System.Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
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
            foreach (var (block, clip, transform) in ReadPhysicalPaintFrames(page))
            foreach (var show in DocxCommentsTests.ReadEmbeddedGlyphTextShows(new PdfPage(page.Width, page.Height, block, page.Fonts)))
            {
                if (!show.Text.Contains("FlowToken", StringComparison.Ordinal)) continue;
                TestAssert.True(Math.Abs(transform.M12) < .000001d, "The source's note text has an axis-aligned vertical frame.");
                var font = show.Font.Font;
                double scale = show.Size / font.UnitsPerEm;
                foreach (char character in show.Text)
                {
                    if (char.IsWhiteSpace(character) || !font.TryGetGlyphBounds(font.MapCodePoint(character), out var ink)) continue;
                    double a = (show.Y + ink.YMin * scale) * transform.M22 + transform.M32;
                    double b = (show.Y + ink.YMax * scale) * transform.M22 + transform.M32;
                    double bottom = Math.Min(a, b), top = Math.Max(a, b);
                    // The unfit tail remains physically clipped by the page;
                    // only glyphs wholly on paper are required to survive.
                    if (bottom < 0d || top > page.Height || top >= 72d) continue;
                    marginGlyphs++;
                    TestAssert.True(bottom >= clip.Y - .01d && top <= clip.Y + clip.Height + .01d,
                        "A page-contained note-cell glyph cannot be lost to the story's flow-admission clip.");
                }
            }
            TestAssert.True(marginGlyphs > 0, "The witness exercises visible table ink below its72pt body margin: " + mode);
            TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_CLIPPED_NOTE_TABLE_ROW"),
                "The remaining row tail beyond the paper still reports unsupported fragmentation.");
        }
    }

    public static void DocxPhysicalPaintFramesRetainQuarterTurnClipBounds()
    {
        const string content = "q\n0 1 -1 0 80 10 cm\n5 7 20 30 re W n\n/Image Do\nQ\n";
        var frame = ReadPhysicalPaintFrames(new PdfPage(100d, 100d, content, [])).Single();
        TestAssert.Equal(new PdfRectangle(43d, 15d, 30d, 20d), frame.Clip);
        TestAssert.Equal(new Matrix3x2(0f, 1f, -1f, 0f, 80f, 10f), frame.Transform);
        string oblique = content.Replace("0 1 -1 0", "0.707107 0.707107 -0.707107 0.707107", StringComparison.Ordinal);
        TestAssert.Throws<InvalidOperationException>(() =>
            ReadPhysicalPaintFrames(new PdfPage(100d, 100d, oblique, [])).ToArray());
    }

    internal readonly record struct PhysicalPaintFrame(string Block, PdfRectangle Clip, Matrix3x2 Transform)
    {
        // Only normal fill text without a soft mask is qualified here. Geometry
        // callers can still inspect other paint; visibility callers must retain unknowns.
        public bool? VisibleTextPaint { get; init; }
    }

    // Read only generated q/Q, rectangle clips/strokes, matrices and text/image paints for these
    // fixtures. CID decoding stays in the existing glyph-show reader; standard
    // affine matrices map clips and glyph ink into physical page coordinates.
    internal static IEnumerable<PhysicalPaintFrame> ReadPhysicalPaintFrames(PdfPage page, bool includeVisibleFills = false)
    {
        const string number = @"[-+\d.eE]+";
        string rectangle = number + @"(?: " + number + @"){3}";
        const string paintPattern = @"(?m)^(?:(?<state>/\S+) gs|(?<mode>[0-7]) Tr)\r?$";
        string fills = includeVisibleFills ? "|(?<fill>" + rectangle + @" re f)" : "";
        string pattern = @"(?m)^(?:(?<save>q)|(?<restore>Q)|(?<clip>" + rectangle + @") re W\*? n|(?<matrix>" + number + @"(?: " + number + @"){5}) cm|(?<state>/\S+) gs|(?<mode>[0-7]) Tr|(?<text>BT)|(?<image>/[A-Za-z0-9_]+ Do)|(?<stroke>" + rectangle + @" re S)" + fills + @")\r?$";
        var saved = new Stack<(PdfRectangle Clip, Matrix3x2 Transform, (bool? FillVisible, bool Mask, int Mode) Paint)>();
        var clip = new PdfRectangle(0d, 0d, page.Width, page.Height);
        var transform = Matrix3x2.Identity;
        (bool? FillVisible, bool Mask, int Mode) paint = (true, false, 0);
        foreach (Match operation in Regex.Matches(page.Content, pattern))
        {
            if (operation.Groups["save"].Success) saved.Push((clip, transform, paint));
            else if (operation.Groups["restore"].Success) (clip, transform, paint) = saved.Pop();
            else if (operation.Groups["state"].Success || operation.Groups["mode"].Success) ApplyPaint(operation, ref paint);
            else if (operation.Groups["matrix"].Success)
            {
                float[] v = operation.Groups["matrix"].Value.Split(' ').Select(x => float.Parse(x, CultureInfo.InvariantCulture)).ToArray();
                transform = new Matrix3x2(v[0], v[1], v[2], v[3], v[4], v[5]) * transform;
            }
            else if (operation.Groups["clip"].Success)
            {
                TestAssert.True(Math.Abs(transform.M12) < .000001d && Math.Abs(transform.M21) < .000001d ||
                    Math.Abs(transform.M11) < .000001d && Math.Abs(transform.M22) < .000001d,
                    "The source uses axis-aligned rectangle clips.");
                float[] v = operation.Groups["clip"].Value.Split(' ').Select(x => float.Parse(x, CultureInfo.InvariantCulture)).ToArray();
                Vector2 first = Vector2.Transform(new(v[0], v[1]), transform);
                Vector2 last = Vector2.Transform(new(v[0] + v[2], v[1] + v[3]), transform);
                double x = Math.Max(clip.X, Math.Min(first.X, last.X)), y = Math.Max(clip.Y, Math.Min(first.Y, last.Y));
                clip = new PdfRectangle(x, y, Math.Min(clip.X + clip.Width, Math.Max(first.X, last.X)) - x,
                    Math.Min(clip.Y + clip.Height, Math.Max(first.Y, last.Y)) - y);
            }
            else if (operation.Groups["text"].Success)
            {
                int end = page.Content.IndexOf("\nET", operation.Index, StringComparison.Ordinal);
                TestAssert.True(end >= 0, "Every generated text object terminates.");
                string block = page.Content[operation.Index..(end + 3)];
                var shows = Regex.Matches(block, @"(?:>\s*Tj|\]\s*TJ)");
                var atShow = paint;
                bool changingPaint = false;
                foreach (Match change in Regex.Matches(block, paintPattern))
                {
                    if (shows.Count != 0 && change.Index < shows[0].Index) ApplyPaint(change, ref atShow);
                    else if (shows.Count != 0 && change.Index < shows[^1].Index) changingPaint = true;
                }
                bool? visible = shows.Count == 0 || changingPaint ? null : atShow.Mode switch
                {
                    0 when !atShow.Mask => atShow.FillVisible,
                    // Synthetic bold paints the glyph fill and adds a stroke.
                    2 when !atShow.Mask && atShow.FillVisible is true => true,
                    3 => false,
                    _ => null
                };
                yield return new(block, clip, transform) { VisibleTextPaint = visible };
            }
            else if (!operation.Groups["fill"].Success || (!paint.Mask && paint.FillVisible is true))
                yield return new(operation.Value.TrimEnd('\r'), clip, transform);
        }
        TestAssert.Equal(0, saved.Count);

        void ApplyPaint(Match operation, ref (bool? FillVisible, bool Mask, int Mode) state)
        {
            if (operation.Groups["mode"].Success) state.Mode = int.Parse(operation.Groups["mode"].Value, CultureInfo.InvariantCulture);
            else
            {
                string name = operation.Groups["state"].Value[1..];
                var resource = page.ExtGStates.SingleOrDefault(item => PdfEmbeddedFont.SanitizeName(item.ResourceName) == name);
                state.FillVisible = resource.ResourceName is null ? null :
                    double.Parse(PdfDocumentWriter.FormatNumber(resource.FillAlpha), CultureInfo.InvariantCulture) > 0d;
                // A graphics state without /SMask does not reset an existing mask.
                state.Mask |= resource.ResourceName is null || resource.LuminosityGroupName is not null;
            }
        }
    }
}
