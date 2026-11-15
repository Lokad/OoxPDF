using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxNumberingTests
{
    public static void DocxListFormatsPaintIndependentComposedDocuments()
    {
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "HierarchicalLists"));
        using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "expected.json")));
        foreach (var row in manifest.RootElement.EnumerateArray())
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string path = Path.Combine(folder, row.GetProperty("Id").GetString() + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            var document = DocxTests.ReadDocx(path, mode);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var pages = renderer.RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(1, pages.Length);
            var shows = pages.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true, honorActualText: false)).ToArray();
            var expected = row.GetProperty("Labels").EnumerateArray().Select(value => value.GetString()!).ToArray();
            // A hyphenated label may paint through several PDF shows. Read
            // each physical row's visible prefix rather than assuming one show
            // per source label; longest matches distinguish IV. from IV.aa).
            var labels = shows.GroupBy(show => Math.Round(show.Y, 2))
                .Select(line => string.Concat(line.OrderBy(show => show.X).Select(show => show.Text)).TrimStart())
                .Select(line => expected.OrderByDescending(label => label.Length)
                    .FirstOrDefault(label => line.StartsWith(label, StringComparison.Ordinal)))
                .Where(label => label is not null).ToArray();
            TestAssert.True(labels.SequenceEqual(expected),
                "Actual embedded glyph paint retains every source-defined list label in document order: " +
                row.GetProperty("Id").GetString() + "; actual " + string.Join(" | ", labels));
            string text = string.Concat(shows.Select(show => show.Text));
            foreach (var prose in row.GetProperty("Text").EnumerateArray()) TestAssert.Contains(prose.GetString()!, text);
            TestAssert.Equal(row.GetProperty("Picture").GetBoolean() ? 1 : 0, pages.Sum(page => page.Images.Count));
            TestAssert.Equal(row.GetProperty("Picture").GetBoolean() ? 1 : 0, pages.Sum(page => page.Annotations.Count(link => link.Uri is not null)));
            TestAssert.Equal(row.GetProperty("Math").GetString(), string.Concat(pages.SelectMany(page =>
                DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true, honorActualText: false,
                    fontResourceFilter: name => name.StartsWith("Eq", StringComparison.Ordinal))).Select(show => show.Text)));
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Formatting a list cannot discard its supported equation.");
            foreach (var page in pages) PdfContentValidator.ValidatePage(page, 0, default);
            using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(document, null, cancelled.Token).ToBlockingEnumerable().ToArray());
        }
    }
}
