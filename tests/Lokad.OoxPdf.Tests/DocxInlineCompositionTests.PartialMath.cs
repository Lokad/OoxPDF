using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionPartialMathKeepsIndependentKnownOriginsAndWarnings()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent partial-math origins use installed Windows fonts.");
        CheckPartialMath(null);
    }

    internal static void CheckPartialMath(IFontResolver? resolver)
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("partial-math-office.json")));
        var controlOrigins = new Dictionary<string, string>();
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            string container = row.GetProperty("Container").GetString()!;
            string preset = row.GetProperty("Mode").GetString()!;
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(NestedMathCasePath(id + ".docx")))));
            bool final = preset == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            foreach (bool transformed in new[] { false, true })
            {
                var source = ReadMutation("CompositionProbes/" + id, mode, transformed, transformed,
                    partOverride: container == "header" ? "word/headerPartial.xml" : null);
                var diagnostics = new List<OoxPdfDiagnostic>();
                var page = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(source, diagnostics.Add, default).Single();
                var unsupported = diagnostics.Where(d => d.Id == "DOCX_UNSUPPORTED_EQUATION").ToArray();
                TestAssert.Equal(row.GetProperty("UnsupportedDiagnostics").GetInt32(), unsupported.Length);
                TestAssert.True(unsupported.All(d => d.PartName == (container == "header" ? "/word/headerPartial.xml" : "/word/document.xml")),
                    "Unsupported items retain their original part diagnostic.");
                var math = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                    .Where(show => page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) &&
                        ReferenceEquals(font.Font, show.Font))).ToArray();
                TestAssert.Equal(row.GetProperty("MathGlyphCount").GetInt32(), math.Length);
                TestAssert.Equal(row.GetProperty("MathText").GetString(),
                    string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
                string text = DocxPageFootnoteNumberingTests.DecodeText(page);
                foreach (string prose in new[] { "Before formula:", "Between formulae:", "after formula.", "Following prose." })
                    TestAssert.Contains(prose, text);
                TestAssert.True(math.All(show => show.X > 0d && show.X < page.Width && show.Y > 0d && show.Y < page.Height),
                    "Prepared siblings remain on their physical page.");
                string origins = string.Join("|", math.Select(show => show.X.ToString("R", CultureInfo.InvariantCulture) + "," +
                    show.Y.ToString("R", CultureInfo.InvariantCulture)));
                string key = container + "/" + preset;
                if (controlOrigins.TryGetValue(key, out string? control)) TestAssert.Equal(control, origins);
                else controlOrigins.Add(key, origins);
                if (resolver is not null) continue;
                foreach (var anchor in row.GetProperty("Anchors").EnumerateArray())
                {
                    var show = math[anchor.GetProperty("Offset").GetInt32()];
                    TestAssert.True(Math.Abs(show.X - anchor.GetProperty("X").GetDouble()) < 1d &&
                        Math.Abs(show.Y - anchor.GetProperty("Y").GetDouble()) < 1d,
                        "Known supported glyphs meet independent Word's existing one-point gate: " + id + "/" + preset);
                }
            }
        }
    }

    public static void DocxInlineCompositionPartialMathRequiresPreparedSourceOwnerAndKeepsFallbacks()
    {
        var document = ReadMutation("CompositionProbes/partial-math-body-unsupported", OoxPdfDocxMarkupMode.Final, false, false);
        var paragraph = document.Paragraphs.Single(p => p.Equations.Count == 2);
        foreach (var kept in paragraph.Equations)
        {
            var prepared = new DocxPreparedMathParagraph(kept, new(18d, 10d, 4d, []), 11d, 16d);
            var equations = new Dictionary<DocxEquation, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance) { [kept] = prepared };
            TestAssert.True(Measure(paragraph, equations, out var atoms),
                "Prepared content survives a missing item before or after it.");
            var fragment = TestAssert.NotNull(atoms).Fragments.Single(f => f.Math is not null);
            TestAssert.True(ReferenceEquals(prepared, fragment.Math) &&
                fragment.Source is DocxInlineEquation item && ReferenceEquals(item.Equation, kept),
                "A retained frame belongs to its actual ordered source equation.");
            TestAssert.True(atoms!.Fragments.All(f => f.Source is not DocxInlineEquation item || ReferenceEquals(item.Equation, kept)),
                "Missing boxes create no fabricated frame or source claim.");
            TestAssert.True(!Measure(paragraph with { InlineContent = null }, equations, out _) &&
                !Measure(paragraph with { LineSpacingPoints = 1d }, equations, out _),
                "Partial preparation cannot bypass source trust or paragraph eligibility.");
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxLayoutEngine.TryMeasureOrderedInlineAtoms(paragraph,
                equations, new TrackingMeasurer(), 12d, 9d, 16d, 1d, cancellation.Token, out _));
        }
        var absent = new Dictionary<DocxEquation, DocxPreparedMathParagraph>();
        TestAssert.True(!Measure(paragraph, absent, out _), "Entirely unavailable body math keeps its existing fallback.");
        var foreign = new DocxEquation(new([]), 999, 999, false, null, null, null);
        absent.Add(foreign, new(foreign, new(18d, 10d, 4d, []), 11d, 16d));
        TestAssert.True(!Measure(paragraph, absent, out _), "Another paragraph's prepared equation is not a ready sibling.");
        TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineAtoms(paragraph, absent, new TrackingMeasurer(),
            12d, 9d, 16d, 1d, default, out var staticAtoms, allowStaticFallback: true) &&
            staticAtoms!.Fragments.All(f => f.Math is null), "Static text retains its established all-missing-math fallback.");

        static bool Measure(DocxParagraph paragraph, IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph> equations,
            out DocxMeasuredInlineAtoms? atoms) => DocxLayoutEngine.TryMeasureOrderedInlineAtoms(paragraph,
                equations, new TrackingMeasurer(), 12d, 9d, 16d, 1d, default, out atoms);
    }
}
