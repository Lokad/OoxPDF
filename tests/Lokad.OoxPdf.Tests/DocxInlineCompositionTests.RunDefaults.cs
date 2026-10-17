using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionRunDefaultsKeepWordSizeOwnershipAcrossContainers()
    {
        using var golden = ReadRunDefaultGolden();
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var document = ReadRunDefaultSource(row);
            var runs = DocxBlockTraversal.EnumerateSourceParagraphs(document).SelectMany(paragraph => paragraph.Runs)
                .Where(run => run.Text.Contains("Public ", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(6, runs.Select(run => run.Text.Trim()).Distinct().Count());
            double expected = row.GetProperty("NominalDefaultSize").GetDouble();
            bool owner = row.GetProperty("HasRunDefaultOwner").GetBoolean();
            foreach (var run in runs)
            {
                TestAssert.Equal(run.Text.Contains("explicit", StringComparison.Ordinal) ? 15d : expected, run.FontSize);
                TestAssert.Equal(owner, run.StyleResolution.HasDocumentDefaultRunProperties);
            }
            TestAssert.Equal(owner, document.StyleCatalog.HasRunDefaults);
            foreach (var paragraph in DocxBlockTraversal.EnumerateSourceParagraphs(document))
                TestAssert.Equal(expected, paragraph.ParagraphMarkFontSize);
        }
    }

    public static void DocxInlineCompositionRunDefaultsStayBelowParagraphAndDirectSizes()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadMutation("CompositionProbes/run-default-owner-calibri-font-run-defaults", mode, false, false,
                (xml, w, _) => xml.Root!.Add(new XElement(w + "style", new XAttribute(w + "type", "paragraph"),
                    new XAttribute(w + "default", "1"), new XAttribute(w + "styleId", "OwnedSizedDefault"),
                    new XElement(w + "name", new XAttribute(w + "val", "Owned sized default")),
                    new XElement(w + "rPr", new XElement(w + "sz", new XAttribute(w + "val", "26"))))),
                partOverride: "word/styles.xml");
            var runs = DocxBlockTraversal.EnumerateSourceParagraphs(document).SelectMany(paragraph => paragraph.Runs)
                .Where(run => run.Text.Contains("Public ", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(6, runs.Select(run => run.Text.Trim()).Distinct().Count());
            foreach (var run in runs)
                TestAssert.Equal(run.Text.Contains("explicit", StringComparison.Ordinal) ? 15d : 13d, run.FontSize);
        }
    }

    public static void DocxInlineCompositionRunDefaultsRenderIndependentSizesAndStoryContent()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native default-owner controls use installed Calibri/Arial.");
        CheckRunDefaultOwnerRendering(null);
    }

    internal static void CheckRunDefaultOwnerRendering(IFontResolver? resolver)
    {
        using var golden = ReadRunDefaultGolden();
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var mode = RunDefaultMode(row);
            var document = ReadRunDefaultSource(row);
            var pages = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, null, default).ToArray();
            TestAssert.Equal(row.GetProperty("ExpectedPages").GetInt32(), pages.Length);
            var shows = pages.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)).ToArray();
            foreach (var expected in row.GetProperty("Samples").EnumerateArray())
            {
                var actual = shows.Single(show => show.Text.Trim() == expected.GetProperty("Text").GetString());
                // Word emits 10/11pt as9.96/11.04 in these independent PDFs.
                TestAssert.True(Math.Abs(actual.Size - expected.GetProperty("FontSize").GetDouble()) <= .06d,
                    "Body/cell/static/note visible font sizes follow the declared default owner.");
            }
            var links = pages.SelectMany(page => page.Annotations).ToArray();
            TestAssert.Equal(1, links.Length);
            TestAssert.True(links[0].Destination is not null && links[0].Width > 0d && links[0].Height > 0d,
                "Default-size resolution preserves the complete custom note marker and target.");
        }
    }

    private static string RunDefaultCaseFolder() => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "CompositionProbes"));

    private static JsonDocument ReadRunDefaultGolden() => JsonDocument.Parse(File.ReadAllText(
        Path.Combine(RunDefaultCaseFolder(), "run-default-owners-office.json")));

    private static OoxPdfDocxMarkupMode RunDefaultMode(JsonElement row) =>
        row.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;

    private static DocxDocument ReadRunDefaultSource(JsonElement row)
    {
        string input = Path.Combine(RunDefaultCaseFolder(), row.GetProperty("Id").GetString()! + ".docx");
        TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
        return DocxTests.ReadDocx(input, RunDefaultMode(row));
    }
}
