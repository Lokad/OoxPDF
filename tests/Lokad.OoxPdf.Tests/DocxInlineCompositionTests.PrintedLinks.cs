using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionPrintedHyperlinkFramesMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent printed link frames require installed Windows fonts.");
        CheckPrintedHyperlinkFrames(null);
    }

    internal static void CheckPrintedHyperlinkFrames(IFontResolver? resolver)
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("justified-linked-images-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(NestedMathCasePath(id + ".docx")))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            foreach (bool transformed in new[] { false, true })
            {
                var document = ReadMutation("CompositionProbes/" + id, mode, transformed, transformed);
                var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
                var page = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(document, diagnostics.Add, default).Single();
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Linked mixed content retains its source formula.");
                var links = page.Annotations.Where(link => link.IsDestination).OrderByDescending(link => link.Y + link.Height).ToArray();
                TestAssert.True(links.Length > 0, "Supported source links produce clickable rectangles with either font resolver.");
                var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray();
                var math = shows.Where(show => page.Fonts.Any(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal) &&
                    ReferenceEquals(resource.Font, show.Font))).ToArray();
                TestAssert.Equal(16, math.Length);
                double minX = links.Min(link => link.X), maxX = links.Max(link => link.X + link.Width),
                    bottom = links.Min(link => link.Y), top = links.Max(link => link.Y + link.Height);
                var linked = shows.Where(show => show.Text.Any(c => !char.IsWhiteSpace(c)) &&
                    show.X >= minX && show.X <= maxX && show.Y >= bottom && show.Y <= top).ToArray();
                TestAssert.Equal(456, linked.Sum(show => show.Text.Count(c => !char.IsWhiteSpace(c))));
                TestAssert.True(linked.All(show => links.Any(link => show.X >= link.X - .001d && show.X <= link.X + link.Width + .001d &&
                    show.Y >= link.Y - .001d && show.Y <= link.Y + link.Height + .001d)),
                    "Every linked source text/math baseline origin remains inside a clickable line rectangle after print mapping.");
                if (resolver is not null) continue;
                TestAssert.Equal(row.GetProperty("LinkRectangles").GetArrayLength(), links.Length);
                foreach (var pair in links.Zip(row.GetProperty("LinkRectangles").EnumerateArray()))
                    TestAssert.True(Math.Abs(pair.First.Y - pair.Second.GetProperty("Bottom").GetDouble()) < 1d &&
                        Math.Abs(pair.First.Y + pair.First.Height - pair.Second.GetProperty("Top").GetDouble()) < 1d,
                        "Body/cell link tops and bottoms retain independent one-point Word gates in both presets.");
                double printScale = final ? 1d : DocxRenderer.ResolveWordCompatiblePrintScale(document,
                    DocxMarkupContext.FromMode(mode, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup));
                foreach (var pair in math.Zip(row.GetProperty("MathAnchors").EnumerateArray()))
                    TestAssert.True(pair.First.Text == pair.Second.GetProperty("Text").GetString() &&
                        Math.Abs(pair.First.X - pair.Second.GetProperty("X").GetDouble()) < 1d &&
                        Math.Abs(pair.First.Y - pair.Second.GetProperty("Y").GetDouble()) < 1d &&
                        Math.Abs(pair.First.Size * printScale - pair.Second.GetProperty("Size").GetDouble()) < .1d,
                        "Printed annotations do not change independent mathematical ink geometry.");
            }
        }
    }
}
