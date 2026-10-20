using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxFooterTableContinuationFramesMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent continuation geometry uses installed Windows fonts.");
        CheckFooterTableContinuationFrames(null);
    }

    internal static void CheckFooterTableContinuationFrames(IFontResolver? resolver)
    {
        foreach (string file in new[] { "static-inline-footer-tables-office.json", "footer-cell-sibling-probes-office.json",
            "footer-table-wrapped-viewport-office.json" })
        {
            using var gold = ReadInlineBoxMathGold(file);
            foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
            foreach (bool split in new[] { false, true })
            {
                string id = row.GetProperty("Id").GetString()!;
                string input = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "InlineBoxMath", id + ".docx"));
                TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
                bool final = row.GetProperty("Mode").GetString() == "final-preserve";
                var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
                string part = row.TryGetProperty("MathPart", out var sourcePart) ? sourcePart.GetString()! : "word/" + id + ".xml";
                var document = ReadMutation("InlineBoxMath/" + id, mode, false, split, partOverride: part);
                var renderer = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                var page = renderer.RenderBlankPages(document, null, default).Single();
                string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Select(show => show.Text))
                    .Normalize(NormalizationForm.FormKC);
                foreach (string expected in new[] { "Before public math table.", "After public math table.", "Before inline",
                    "after equation.", "Inline box content." })
                    TestAssert.True(text.Contains(expected, StringComparison.Ordinal), id + ": admission preserves parent and child source text.");
                TestAssert.Equal(row.GetProperty("WordImages").GetArrayLength(), page.Images.Count);
                if (resolver is null)
                {
                    var target = row.GetProperty("WordFrames")[0].GetProperty("Rect");
                    double x = target[0].GetDouble(), top = target[1].GetDouble();
                    double width = target[2].GetDouble() - x, height = target[3].GetDouble() - top;
                    bool found = Regex.Matches(page.Content, @"([-\d.]+) ([-\d.]+) ([-\d.]+) ([-\d.]+) re\s+S").Any(match =>
                        Math.Abs(double.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture) - x) < 1d &&
                        Math.Abs(page.Height - double.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture) -
                            double.Parse(match.Groups[4].Value, CultureInfo.InvariantCulture) - top) < 1d &&
                        Math.Abs(double.Parse(match.Groups[3].Value, CultureInfo.InvariantCulture) - width) < .1d &&
                        Math.Abs(double.Parse(match.Groups[4].Value, CultureInfo.InvariantCulture) - height) < .1d);
                    TestAssert.True(found, id + ": frame placement retains independent Word origins and physical size.");
                }
                using var cancelled = new CancellationTokenSource();
                cancelled.Cancel();
                TestAssert.Throws<OperationCanceledException>(() => renderer.RenderBlankPages(document, null, cancelled.Token).ToArray());
            }
        }
    }

    public static void DocxFooterTableFrameAdmissionRetainsAuthoredBreakPlacement()
    {
        var document = ReadMutation("InlineBoxMath/footer-table-viewport-wrapped-margin", OoxPdfDocxMarkupMode.Final, false, false,
            (xml, w, _) =>
            {
                var box = xml.Descendants(w + "drawing").Single(drawing => drawing.Descendants(w + "txbxContent").Any());
                box.Parent!.AddBeforeSelf(new System.Xml.Linq.XElement(w + "r", new System.Xml.Linq.XElement(w + "br")));
            }, partOverride: "word/static-inline-footer-table-picture.xml");
        var page = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), default).Pages.Single();
        var cell = page.StaticTableRows.Single().Cells.Single();
        var frame = cell.InlineTextBoxes.Single();
        var line = cell.TextLines.Single(item => item.SourceLineIndex == frame.SourceLineIndex);
        TestAssert.True(frame.BoxTop - frame.BoxHeight < page.MarginBottom, "The explicit-break control exercises an overflowing frame.");
        TestAssert.True(Math.Abs(frame.BoxTop - line.BaselineY - line.InlineFrame!.Value.BaselineInset) < 1e-6,
            "An authored break retains its source line placement instead of moving into an earlier line.");
    }
}
