using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static class DocxPassiveStaticBodyTests
{
    public static void DocxPassiveStaticTextPreservesIndependentWordBodyBaselines()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent minimum body origins require installed Windows fonts.");
        CheckBodyBaselines(null, compareWord: true);
    }

    public static void DocxPassivePageFieldsPreserveIndependentWordBodyBaselines()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent minimum body/page-field origins require installed Windows fonts.");
        CheckBodyBaselines(null, compareWord: true, "minimum-passive-page-fields-office.json");
    }

    public static void DocxInactiveStaticVariantsPreserveIndependentWordBodyOrigins()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent inactive-template origins require installed Windows fonts.");
        CheckBodyBaselines(null, compareWord: true, "minimum-inactive-static-office.json");
    }

    public static void DocxInactiveNoteTemplatesPreserveSourceNumberingAndWordOrigins()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent note origins require installed Windows fonts.");
        CheckInactiveNoteTemplates(null, compareWord: true);
    }

    internal static void CheckInactiveNoteTemplates(IFontResolver? resolver, bool compareWord)
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(Case("CompositionProbes/minimum-inactive-notes-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string input = Case("CompositionProbes/" + row.GetProperty("Id").GetString() + ".docx");
            string original = Case(row.GetProperty("SourceInput").GetString()!.Replace("tests/Lokad.OoxPdf.Tests/Cases/", "", StringComparison.Ordinal));
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            TestAssert.Equal(row.GetProperty("SourceSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(original))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var geometry = final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
            var before = Shows(original);
            var after = Shows(input);
            TestAssert.True(before.Length > 8, "The control includes body, marker and note glyphs.");
            TestAssert.Equal(before.Length, after.Length);
            for (int index = 0; index < before.Length; index++)
            {
                TestAssert.Equal(before[index].Text, after[index].Text);
                TestAssert.Equal(before[index].Page, after[index].Page);
                TestAssert.True(Math.Abs(before[index].X - after[index].X) < .001d &&
                    Math.Abs(before[index].Y - after[index].Y) < .001d && Math.Abs(before[index].Size - after[index].Size) < .001d,
                    "Disabled templates preserve every source-owned body, marker and note glyph origin and size.");
            }
            if (compareWord)
                foreach (var anchor in row.GetProperty("WordProseAnchors").EnumerateArray().Where(anchor =>
                    anchor.GetProperty("Text").GetString()!.StartsWith("Public minimum ", StringComparison.Ordinal)))
                {
                    var show = after.Single(show => show.Text.Trim() == anchor.GetProperty("Text").GetString());
                    TestAssert.Equal(anchor.GetProperty("Page").GetInt32(), show.Page);
                    TestAssert.True(Math.Abs(show.X - anchor.GetProperty("X").GetDouble()) < 1d &&
                        Math.Abs(show.Y - anchor.GetProperty("Y").GetDouble()) < 1d,
                        "Minimum note-body lines retain independent one-point Word origins.");
                }

            (int Page, string Text, double X, double Y, double Size)[] Shows(string path) =>
                new DocxRenderer(resolver, mode, geometry).RenderBlankPages(DocxTests.ReadDocx(path, mode), null, default)
                    .SelectMany((page, index) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                        .Where(show => show.Text.Length != 0).Select(show =>
                            (index + 1, show.Text.Normalize(NormalizationForm.FormKC), show.X, page.Height - show.Y, show.Size)))
                    .ToArray();
        }
    }

    public static void DocxInactiveNoteTemplatesKeepEnabledStoryAndSectionFallbacks()
    {
        var document = DocxTests.ReadDocx(Case("CompositionProbes/minimum-inactive-first-header-noted12.docx"), OoxPdfDocxMarkupMode.Final);
        TestAssert.True(DocxPageFootnoteNumbering.IsEligibleDocument(document), "A disabled first template preserves existing note source admission.");
        TestAssert.True(!DocxPageFootnoteNumbering.IsEligibleDocument(document with
            { PageSettings = document.PageSettings with { TitlePage = true } }), "An enabled first story retains the original fallback.");
        var enabledEven = document with { PageSettings = document.PageSettings with { EvenAndOddHeaders = true },
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
                { ["EVEN"] = document.HeaderBodyElementsByType["first"] } };
        TestAssert.True(!DocxPageFootnoteNumbering.IsEligibleDocument(enabledEven), "Any selectable variant retains fallback, including case-insensitive source keys.");
        TestAssert.True(!DocxPageFootnoteNumbering.IsEligibleDocument(document with
            { BodyElements = document.BodyElements.Append(document.FinalSectionBreak!).ToArray() }), "Multiple source sections retain their existing admission.");
    }

    internal static void CheckBodyBaselines(IFontResolver? resolver, bool compareWord,
        string goldFile = "minimum-passive-static-office.json")
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(Case("CompositionProbes/" + goldFile)));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            string input = Case("CompositionProbes/" + id + ".docx");
            string original = Case(row.GetProperty("SourceInput").GetString()!.Replace("tests/Lokad.OoxPdf.Tests/Cases/", "", StringComparison.Ordinal));
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            TestAssert.Equal(row.GetProperty("SourceSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(original))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var geometry = final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
            var source = Lines(input);
            var baseline = Lines(original).Where(line => line.Text.Length != 0 &&
                !line.Text.StartsWith("Public passive ", StringComparison.Ordinal)).ToArray();
            var bodyTexts = baseline.Select(line => line.Text).ToHashSet(StringComparer.Ordinal);
            var body = source.Where(line => bodyTexts.Contains(line.Text)).ToArray();
            TestAssert.Equal(8, baseline.Length);
            TestAssert.Equal(8, body.Length);
            foreach (var before in baseline)
            {
                var after = body.Single(line => line.Text == before.Text);
                TestAssert.Equal(before.Page, after.Page);
                TestAssert.True(Math.Abs(before.X - after.X) < .001d && Math.Abs(before.Y - after.Y) < .001d,
                    "Adding a measured margin story must preserve the ordinary body's glyph origins with native or portable fonts.");
            }
            foreach (var expected in row.GetProperty("WordProseAnchors").EnumerateArray())
            {
                var actual = source.Single(line => line.Text == expected.GetProperty("Text").GetString());
                if (!compareWord) continue;
                TestAssert.Equal(expected.GetProperty("Page").GetInt32(), actual.Page);
                TestAssert.True(Math.Abs(actual.X - expected.GetProperty("X").GetDouble()) < 1d &&
                    Math.Abs(actual.Y - expected.GetProperty("Y").GetDouble()) < 1d,
                    "Body and passive story retain their independent one-point Word origins.");
            }

            (int Page, double X, double Y, string Text)[] Lines(string path) =>
                new DocxRenderer(resolver, mode, geometry).RenderBlankPages(DocxTests.ReadDocx(path, mode), null, default)
                    .SelectMany((page, index) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                        .GroupBy(show => show.Y).Select(line => (index + 1, line.Min(show => show.X), page.Height - line.Key,
                            Regex.Replace(string.Concat(line.OrderBy(show => show.X).Select(show => show.Text))
                                .Normalize(NormalizationForm.FormKC), @"\s+", " ").Trim())))
                    .ToArray();
        }
    }

    public static void DocxPassiveStaticTextKeepsIntrusiveUnsupportedAndUnmeasuredFallbacks()
    {
        foreach (string kind in new[] { "header", "footer" })
        {
            string input = Case("CompositionProbes/minimum-passive-" + kind + "-ordinary12.docx");
            var document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);
            var layout = Layout(document);
            TestAssert.True(DocxMinimumBodyLineLeading.TryCreateLayout(document, layout, out _), "A complete plain margin frame admits ordinary minimum leading.");
            var page = layout.Pages[0];
            var intrusive = kind == "header"
                ? layout with { HeaderContentBottomByPage = new Dictionary<int, double> { [0] = page.Height - page.MarginTop - 1d } }
                : layout with { FooterContentTopByPage = new Dictionary<int, double> { [0] = page.MarginBottom + 1d } };
            var unmeasured = layout with { HeaderContentBottomByPage = new Dictionary<int, double>() };
            var annotationLane = layout with { Pages = [page with { MarkupMarginReservePoints = 20d }] };
            foreach (var unsupported in new[] { intrusive, unmeasured, annotationLane })
                TestAssert.True(!DocxMinimumBodyLineLeading.TryCreateLayout(document, unsupported, out var unchanged) && ReferenceEquals(unsupported, unchanged),
                    "Intrusive, unmeasured and annotation-lane frames retain the existing fallback.");
            var dynamic = ReadUnsupportedFieldStory(input, kind);
            var dynamicLayout = Layout(dynamic);
            TestAssert.True(!DocxMinimumBodyLineLeading.TryCreateLayout(dynamic, dynamicLayout, out var fallback) && ReferenceEquals(dynamicLayout, fallback),
                "An unsupported cached field remains outside source-owned margin-text admission.");
        }
    }

    public static void DocxInactiveStaticVariantsKeepSelectedRichStoryFallbacks()
    {
        foreach (var (kind, variant) in new[] { ("header", "first"), ("footer", "even") })
        {
            string input = Case("CompositionProbes/minimum-inactive-" + variant + "-" + kind + ".docx");
            var inactive = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);
            TestAssert.True(DocxMinimumBodyLineLeading.TryCreateLayout(inactive, Layout(inactive), out _),
                "An unselected rich template does not own body metrics.");
            var active = ReadMutation(input, archive =>
            {
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                Update("word/document.xml", xml =>
                {
                    if (variant == "first") xml.Descendants(w + "sectPr").Single().Add(new XElement(w + "titlePg"));
                    else xml.Root!.Element(w + "body")!.Elements(w + "p").ElementAt(4)
                        .Element(w + "pPr")!.Add(new XElement(w + "pageBreakBefore"));
                });
                if (variant == "even") Update("word/settings.xml", xml => xml.Root!.Add(new XElement(w + "evenAndOddHeaders")));
                void Update(string part, Action<XDocument> change)
                {
                    var entry = archive.GetEntry(part)!;
                    XDocument xml;
                    using (var stream = entry.Open()) xml = XDocument.Load(stream);
                    change(xml);
                    entry.Delete();
                    using var output = archive.CreateEntry(part).Open();
                    xml.Save(output);
                }
            });
            var layout = Layout(active);
            TestAssert.Equal(variant == "first" ? 1 : 2, layout.Pages.Count);
            TestAssert.True(layout.Pages.Any(page => page.StaticTableRows.Count != 0), "The activated rich table actually owns a selected page story.");
            TestAssert.True(!DocxMinimumBodyLineLeading.TryCreateLayout(active, layout, out var fallback) && ReferenceEquals(layout, fallback),
                "Selected rich stories retain their independently staged body-metric fallback.");
        }
    }

    private static DocxLayout Layout(DocxDocument document)
    {
        var resolver = new TestFaceFontResolver();
        var measurer = new DocxFontPlanTextMeasurer(DocxFontPlan.Create(document, resolver, default),
            resolver.Resolve(new FontRequest("Calibri", false, false)), default, resolver);
        return new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, measurer, default);
    }

    private static DocxDocument ReadUnsupportedFieldStory(string input, string kind) => ReadMutation(input, archive =>
    {
        string part = "word/passive-body-" + kind + ".xml";
        var entry = archive.GetEntry(part)!;
        XDocument xml;
        using (var stream = entry.Open()) xml = XDocument.Load(stream);
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var paragraph = xml.Root!.Element(w + "p")!;
        var run = paragraph.Element(w + "r")!;
        run.Remove();
        paragraph.Add(new XElement(w + "fldSimple", new XAttribute(w + "instr", "DATE"), run));
        entry.Delete();
        using var output = archive.CreateEntry(part).Open();
        xml.Save(output);
    });

    public static void DocxDynamicFieldFormatKeepsUnqualifiedInstructionFallbacks()
    {
        foreach (string instruction in new[] { "PAGE \\* CHARFORMAT", "PAGE \\# \"0000\"", "PAGE \\* FutureFormat" })
        {
            var document = ReadMutation(Case("CompositionProbes/minimum-passive-header-page.docx"), archive =>
            {
                var entry = archive.GetEntry("word/passive-body-header.xml")!;
                XDocument xml;
                using (var stream = entry.Open()) xml = XDocument.Load(stream);
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                xml.Descendants(w + "fldSimple").Single().SetAttributeValue(w + "instr", instruction);
                entry.Delete();
                using var output = archive.CreateEntry("word/passive-body-header.xml").Open();
                xml.Save(output);
            });
            var paragraph = document.HeaderParagraphs.Single();
            TestAssert.Equal(11d, paragraph.Runs.Single(run => run.FieldKind == DocxFieldKind.Page).FontSize);
            TestAssert.Equal(instruction, paragraph.FieldReferences.Single().Instruction);
        }
    }

    private static DocxDocument ReadMutation(string input, Action<ZipArchive> mutate)
    {
        string temporary = Path.Combine(Path.GetTempPath(), "ooxpdf-passive-" + Guid.NewGuid().ToString("N") + ".docx");
        File.Copy(input, temporary);
        try
        {
            // Frozen inputs are read-only; the independent temporary copy is
            // intentionally mutable, without changing the frozen source.
            File.SetAttributes(temporary, FileAttributes.Normal);
            using (var archive = ZipFile.Open(temporary, ZipArchiveMode.Update))
                mutate(archive);
            return DocxTests.ReadDocx(temporary, OoxPdfDocxMarkupMode.AllMarkup);
        }
        finally { File.Delete(temporary); }
    }

    private static string Case(string relative) => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", relative));
}
