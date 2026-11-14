using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxBuiltInFallbackKeepsPageLocalNotePaintsAndWarnings()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadMutation("CompositionProbes/math-page-notes-arial-11-zero", mode, false, false,
                (xml, _, m) => xml.Descendants(m + "oMath").Remove());
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new DocxRenderer(new MissingTestFontResolver(), mode, CachedMathGeometry(mode))
                .RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(2, pages.Length);
            foreach (var page in pages)
            {
                TestAssert.Equal(0, page.Fonts.Count);
                TestAssert.True(page.FallbackFonts.Count != 0, "The test must exercise actual built-in font paint, not an embedded substitute.");
                var markers = new List<string>();
                foreach (Match region in Regex.Matches(page.Content, @"BT(.*?)ET", RegexOptions.Singleline))
                {
                    Match size = Regex.Match(region.Value, @"/(\S+)\s+([\d.]+)\s+Tf");
                    if (!size.Success || double.Parse(size.Groups[2].Value, CultureInfo.InvariantCulture) >= 9d) continue;
                    string text = OoxMissingFontTests.ExtractWinAnsiText(region.Value);
                    if (text.Length != 0 && text.All(char.IsAsciiDigit)) markers.Add(text);
                }
                TestAssert.Equal("09|10|09|10", string.Join('|', markers));
            }
            string all = string.Concat(pages.Select(page => OoxMissingFontTests.ExtractWinAnsiText(page.Content)));
            foreach (string word in new[] { "Alpha", "Bravo", "Charlie", "Delta" })
                TestAssert.Contains("Public destination " + word + ".", all);
            TestAssert.Contains("retain authored digits 108", all);
            var missing = diagnostics.Where(d => d.Id == "FONT_NO_USABLE_FACE").ToArray();
            TestAssert.True(missing.Length != 0, "Built-in fallback remains diagnosed after a numbering retry.");
            TestAssert.Equal(missing.Distinct().Count(), missing.Length);
        }
    }
}
