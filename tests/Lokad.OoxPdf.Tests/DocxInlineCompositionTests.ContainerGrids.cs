using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxContainerTableGridsMatchIndependentWordHorizontalOriginsAndNestedOffsets()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent container-grid origins use installed Windows fonts.");
        using var cases = ReadContainerGridJson("container-grid.json");
        using var gold = ReadContainerGridJson("container-grid-office.json");
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string role = row.GetProperty("Role").GetString()!;
            var source = cases.RootElement.EnumerateArray().Single(item => item.GetProperty("Id").GetString() == row.GetProperty("Id").GetString());
            TestAssert.Equal(source.GetProperty("Sha256").GetString(), row.GetProperty("InputSha256").GetString());
            var mode = row.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var page = RenderContainerGrid(source, mode, null);
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).ToArray();
            double actualOrigin = role == "NestedRelative" ? shows.Single(show => show.Text.Trim() == "Outer origin control.").X : 0d;
            double wordOrigin = role == "NestedRelative" ? row.GetProperty("ParentOrigins")[0].GetProperty("X").GetDouble() : 0d;
            foreach (var expected in row.GetProperty("TextOrigins").EnumerateArray())
            {
                var actual = shows.Single(show => show.Text.Trim() == expected.GetProperty("Text").GetString());
                TestAssert.True(Math.Abs(actual.X - expected.GetProperty("X").GetDouble()) < 1d,
                    "The table retains its independently measured absolute origin: " + source.GetProperty("Id").GetString() + "/" + mode);
                if (role == "NestedRelative")
                    TestAssert.True(Math.Abs(actual.X - actualOrigin - expected.GetProperty("X").GetDouble() + wordOrigin) < 1d,
                        "The table owns its independently measured relative frame: " + source.GetProperty("Id").GetString() + "/" + mode);
            }
            var math = shows.Where(show => show.Text.Trim().Length != 0 && page.Fonts.Any(font =>
                font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(font.Font, show.Font))).ToArray();
            var native = row.GetProperty("Math").EnumerateArray().ToArray();
            TestAssert.Equal(native.Length, math.Length);
            for (int i = 0; i < math.Length; i++)
            {
                TestAssert.Equal(native[i].GetProperty("Text").GetString(), math[i].Text.Normalize(NormalizationForm.FormKC));
                TestAssert.True(Math.Abs(math[i].X - native[i].GetProperty("X").GetDouble()) < 1d,
                    "Mathematics retains the table's independently measured absolute origin.");
                if (role == "NestedRelative")
                    TestAssert.True(Math.Abs(math[i].X - actualOrigin - native[i].GetProperty("X").GetDouble() + wordOrigin) < 1d,
                        "Mathematics follows the same relative container origin as prose.");
            }
        }
    }

    public static void DocxContainerTableGridsPreservePortableContentAndPaintedLinks()
    {
        using var cases = ReadContainerGridJson("container-grid.json");
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (var row in cases.RootElement.EnumerateArray())
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var page = RenderContainerGrid(row, mode, resolver);
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).ToArray();
            var north = shows.Single(show => show.Text.Trim() == "BOX TABLE NORTH");
            var south = shows.Single(show => show.Text.Trim() == "BOX TABLE SOUTH");
            TestAssert.True(Math.Abs(north.X - south.X) < .002d, "Both rows share one grid across ordinary and mathematical content.");
            var math = shows.Where(show => show.Text.Trim().Length != 0 && page.Fonts.Any(font =>
                font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(font.Font, show.Font))).ToArray();
            TestAssert.Equal(row.GetProperty("ExpectedEquationCount").GetInt32() == 0 ? "" : "x+2y",
                string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
            var link = page.Annotations.Single(annotation => annotation.Uri == "https://example.invalid/static-box-table");
            TestAssert.True(north.X >= link.X && north.X <= link.X + link.Width && north.Y >= link.Y && north.Y <= link.Y + link.Height,
                $"The clickable source follows its painted table row: {row.GetProperty("Id").GetString()}/{mode}, glyph=({north.X},{north.Y}), link=({link.X},{link.Y},{link.Width},{link.Height}).");
            var destination = page.Annotations.Single(annotation => annotation.IsDestination).Destination!.Value;
            TestAssert.True(destination.PageIndex == 0 && destination.Top is double top && top >= north.Y && top <= page.Height,
                "The bookmark viewport reveals its painted target: " + row.GetProperty("Id").GetString() + "/" + mode);
        }
    }

    private static PdfPage RenderContainerGrid(JsonElement row, OoxPdfDocxMarkupMode mode, IFontResolver? resolver,
        CancellationToken cancellationToken = default)
    {
        const string prefix = "tests/Lokad.OoxPdf.Tests/Cases/";
        string input = row.GetProperty("Input").GetString()!;
        TestAssert.True(input.StartsWith(prefix, StringComparison.Ordinal), "Container-grid inputs stay in the public fixture tree.");
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", input[prefix.Length..]));
        using var stream = File.OpenRead(path);
        TestAssert.Equal(row.GetProperty("Sha256").GetString(), Convert.ToHexString(SHA256.HashData(stream)));
        stream.Position = 0;
        using var package = OoxPackage.Open(stream, default);
        var document = new DocxReader().Read(package, null, default, markupMode: mode);
        var geometry = mode == OoxPdfDocxMarkupMode.AllMarkup ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout;
        return new DocxRenderer(resolver, mode, geometry).RenderPagesAsync(document, null, cancellationToken).ToBlockingEnumerable().Single();
    }

    private static JsonDocument ReadContainerGridJson(string name) => JsonDocument.Parse(File.ReadAllText(Path.GetFullPath(
        Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "TableGridOrigins", name))));
}
