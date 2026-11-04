using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    internal static void CheckAdjacentMathGroups()
    {
        foreach ((string container, string? first, string? last, string kind, bool inline) in
            new (string, string?, string?, string, bool)[]
        {
            ("body", "left", "right", "adjacent", false), ("body", "right", "left", "adjacent", false),
            ("body", "left", null, "adjacent", false), ("body", "right", "empty", "adjacent", false),
            ("body", "left", "right", "three", false), ("body", "left", "right", "bookmark", false),
            ("body", "left", "right", "empty-run", false), ("body", "left", "left", "bare-tail", false),
            ("body", "left", "right", "hidden-break", false), ("body", "left", "right", "prose", true),
            ("body", "left", "right", "space", true), ("body", "left", "right", "hidden-space", true),
            ("body", "left", "right", "leading-space", true), ("body", "left", "right", "trailing-space", true),
            ("body", "left", "right", "link", true), ("body", "left", "right", "picture", true),
            ("cell", "left", "right", "adjacent", false), ("footer", "right", "left", "adjacent", false)
        })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string fixture = kind == "picture" ? "inline-math-navigation" : container == "footer"
                ? "math-explicit-display-right-footer" : "math-explicit-display-left-" + container;
            string part = container == "footer" ? "word/math-tabs-footer.xml" : "word/document.xml";
            int count = kind == "three" ? 3 : 2;
            DocxDocument Read(bool canonical) => ReadMutation("CompositionProbes/" + fixture, mode, false, false,
                (xml, w, m) =>
                {
                    var paragraph = xml.Descendants(w + "p").Single(p => p.Descendants(m + "oMath").Count() == 2 &&
                        (kind != "picture" || p.Descendants(w + "drawing").Any()));
                    XElement firstGroup, next;
                    if (kind == "picture")
                    {
                        // This fixture also has an unrelated two-formula paragraph.
                        // Keep the actual picture owner as an independent control.
                        paragraph.Parent!.Elements().Where(element => element != paragraph && element.Name != w + "sectPr").Remove();
                        paragraph.Elements(w + "r").Where(run => !run.Descendants(w + "drawing").Any()).Remove();
                        var equations = paragraph.Elements(m + "oMath").ToArray();
                        foreach (var equation in equations)
                            equation.ReplaceWith(new XElement(m + "oMathPara", new XElement(equation)));
                        firstGroup = paragraph.Elements(m + "oMathPara").First();
                        next = paragraph.Elements(m + "oMathPara").Last();
                    }
                    else
                    {
                        firstGroup = paragraph.Elements(m + "oMathPara").Single();
                        var second = firstGroup.Elements(m + "oMath").Last();
                        second.Remove(); next = new XElement(m + "oMathPara", second);
                        firstGroup.AddAfterSelf(next);
                    }
                    void Justify(XElement group, string? value)
                    {
                        group.Element(m + "oMathParaPr")?.Remove();
                        if (value is not null)
                            group.AddFirst(new XElement(m + "oMathParaPr", value == "empty" ? new XElement(m + "jc") :
                                new XElement(m + "jc", new XAttribute(m + "val", value))));
                    }
                    XElement Text(string value, bool hidden = false) => new(w + "r",
                        hidden ? new XElement(w + "rPr", new XElement(w + "vanish")) : null,
                        new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), value));
                    Justify(firstGroup, first); Justify(next, last);
                    if (kind == "three")
                    {
                        var third = new XElement(next); Justify(next, "center"); next.AddAfterSelf(third);
                    }
                    if (kind == "bare-tail" && !canonical) next.ReplaceWith(new XElement(next.Element(m + "oMath")!));
                    if (kind == "bookmark" || kind == "link")
                    {
                        firstGroup.AddBeforeSelf(new XElement(w + "bookmarkStart", new XAttribute(w + "id", "990"), new XAttribute(w + "name", "NormalizationTarget")));
                        firstGroup.AddAfterSelf(new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "990")));
                    }
                    if (kind == "link") next.ReplaceWith(new XElement(w + "hyperlink",
                        new XAttribute(w + "anchor", "NormalizationTarget"), new XElement(next)));
                    if (kind is "space" or "hidden-space" or "prose")
                        next.AddBeforeSelf(Text(kind == "prose" ? " boundary prose " : " ", kind == "hidden-space"));
                    if (kind == "leading-space") firstGroup.AddBeforeSelf(Text(" "));
                    if (kind == "trailing-space") next.AddAfterSelf(Text(" "));
                    if (kind == "empty-run") next.AddBeforeSelf(new XElement(w + "r"));
                    if (kind == "hidden-break") next.AddBeforeSelf(new XElement(w + "r",
                        new XElement(w + "rPr", new XElement(w + "vanish")), new XElement(w + "br")));
                    if (canonical)
                        foreach (var group in paragraph.Descendants(m + "oMathPara"))
                            Justify(group, kind == "hidden-break" ? first : last);
                }, partOverride: part);

            var baseline = Read(true); var source = Read(false);
            var paragraph = DocxBlockTraversal.EnumerateBodyParagraphs(source.EquationParts
                .Single(value => value.PartName == "/" + part).BodyElements).Single(p => p.Equations.Count == count);
            TestAssert.Equal(count, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
            TestAssert.Equal(inline ? 0 : count, DocxMathParagraphPolicy.GetDisplaySourceEquations(paragraph, default).Count);
            var metadata = paragraph.Equations.Select(eq => (eq.Display, eq.AlignmentValue, eq.PropertiesXml)).ToArray();
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var before = renderer.RenderPagesAsync(baseline, null, default).ToBlockingEnumerable().ToArray();
            var diagnostics = new List<OoxPdfDiagnostic>();
            var after = renderer.RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(before.Length, after.Length);
            for (int index = 0; index < before.Length; index++)
            {
                var expected = DocxCommentsTests.ReadEmbeddedGlyphTextShows(before[index], usePageCoordinates: true)
                    .Select(show => (show.Text, show.X, show.Y, show.Size)).ToArray();
                var actual = DocxCommentsTests.ReadEmbeddedGlyphTextShows(after[index], usePageCoordinates: true)
                    .Select(show => (show.Text, show.X, show.Y, show.Size)).ToArray();
                TestAssert.True(expected.Length != 0 && expected.SequenceEqual(actual),
                    "Native-equivalent group formatting preserves all glyphs and placement: " + container + "/" + kind + "/" + mode);
                TestAssert.Equal(before[index].Annotations.Count, after[index].Annotations.Count);
                if (kind == "picture") TestAssert.Equal(1, after[index].Images.Sum(image =>
                    Regex.Matches(after[index].Content, @"/" + Regex.Escape(image.ResourceName) + @"\s+Do\b").Count));
            }
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Known group formatting cannot discard supported formulas.");
            TestAssert.True(metadata.SequenceEqual(paragraph.Equations.Select(eq => (eq.Display, eq.AlignmentValue, eq.PropertiesXml))),
                "Effective grouping does not overwrite source properties.");
        }
    }

    private static DocxDocument ReadMathWithSourceCharacters(OoxPdfDocxMarkupMode mode, string? text, bool hidden, bool display = false) =>
        ReadMutation("docx-omml-isolated-inline-body-fraction-left", mode, false, false, (xml, w, m) =>
        {
            var equation = xml.Descendants(m + "oMath").Single();
            if (text is not null)
            {
                var run = new XElement(w + "r", hidden ? new XElement(w + "rPr", new XElement(w + "vanish")) : null,
                    new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), text));
                equation.AddBeforeSelf(run);
                if (!hidden) equation.AddAfterSelf(new XElement(run));
            }
            if (display) equation.ReplaceWith(new XElement(m + "oMathPara", new XElement(equation)));
        });

    internal static void CheckMathSourceCharacters()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var baseline = renderer.RenderPagesAsync(ReadMathWithSourceCharacters(mode, null, false), null, default).ToBlockingEnumerable().Single();
            var baselineMath = MathShows(baseline);
            foreach ((string text, bool hidden, bool display) in new[] { ("Invisible prose", true, false), (" ", false, false), (" ", false, true) })
            {
                var document = ReadMathWithSourceCharacters(mode, text, hidden, display);
                var paragraph = document.Paragraphs.Single(p => p.Equations.Count != 0);
                TestAssert.True(DocxMathParagraphPolicy.GetStandaloneRow(paragraph) is null, "Ordinary source characters retain inline mode before visibility filtering.");
                TestAssert.Equal(0, DocxMathParagraphPolicy.GetDisplaySourceEquations(paragraph, default).Count);
                var diagnostics = new List<OoxPdfDiagnostic>();
                var page = renderer.RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().Single();
                var math = MathShows(page);
                TestAssert.True(math.Length != 0 && math.Max(show => show.Size) < baselineMath.Max(show => show.Size) && math[0].X < baselineMath[0].X,
                    "Inline math follows the independently observed smaller left-aligned frame.");
                string painted = DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC);
                TestAssert.True(!painted.Contains("Invisible", StringComparison.Ordinal), "Hidden source text cannot paint.");
                TestAssert.Equal("x+2y", string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Ordinary source characters cannot discard the formula.");
            }
        }
    }

    private static (string Text, double X, double Y, double Size)[] MathShows(PdfPage page)
    {
        var fonts = page.Fonts.Where(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).Select(font => font.Font).ToHashSet();
        return DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).Where(show => fonts.Contains(show.Font))
            .Select(show => (show.Text, show.X, show.Y, show.Size)).ToArray();
    }

    public static void DocxInlineCompositionHiddenSourceMathMatchesIndependentWord()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for independent hidden-source math geometry.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases"));
        using var reference = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "CompositionProbes/math-hidden-source-office.json")));
        TestAssert.Equal(reference.RootElement.GetProperty("SourceSha256").GetString(), Convert.ToHexString(SHA256.HashData(
            File.ReadAllBytes(Path.Combine(cases, reference.RootElement.GetProperty("SourceFixture").GetString()! + ".docx")))));
        foreach (var row in reference.RootElement.GetProperty("Rows").EnumerateArray())
        {
            bool hidden = row.GetProperty("HiddenProse").GetBoolean();
            var mode = row.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var page = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderPagesAsync(ReadMathWithSourceCharacters(mode, hidden ? "Invisible prose" : null, hidden), null, default)
                .ToBlockingEnumerable().Single();
            var first = MathShows(page)[0]; var anchor = row.GetProperty("Anchor");
            TestAssert.True(Math.Abs(first.X - anchor.GetProperty("X").GetDouble()) < 1d &&
                Math.Abs(first.Y - anchor.GetProperty("Y").GetDouble()) < 1d && Math.Abs(first.Size - anchor.GetProperty("FontSize").GetDouble()) < .1d,
                "Math mode follows independent Word first-glyph geometry: " + mode + "/hidden=" + hidden);
        }
    }
}
