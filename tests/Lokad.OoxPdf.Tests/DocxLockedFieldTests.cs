using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static class DocxLockedFieldTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public static void DocxLockedFieldsPreserveIndependentWordValuesAndUnlockedCounters() =>
        CheckWord("locked-page-fields.docx");

    public static void DocxLockedFieldsPreserveEmptyResultsAndIndependentNestedCounters() =>
        CheckWord("locked-page-field-edges.docx");

    public static void DocxLockedFieldsUseBooleanLocksAndStoredFormatting()
    {
        foreach (string? value in new string?[] { null, "0", "false", "off", "1", "true", "on" })
        foreach (bool simple in new[] { true, false })
        foreach (string kind in new[] { "PAGE", "NUMPAGES" })
        {
            bool locked = value is "1" or "true" or "on";
            var paragraph = new XElement(W + "p", Field(kind, "99", value, simple, styled: true));
            var document = DocxTests.ReadDocx(Write([paragraph]), OoxPdfDocxMarkupMode.Final);
            var parsed = document.Paragraphs.Single();
            var run = parsed.Runs.Single();
            TestAssert.Equal(locked ? "99" : "{" + kind + "}", run.Text);
            TestAssert.Equal(locked ? 20d : 12d, run.FontSize);
            TestAssert.Equal(locked, run.Bold);
            TestAssert.Equal(!locked, parsed.FieldReferences.Single().UsesPlaceholder);
            TestAssert.Equal(locked, parsed.FieldReferences.Single().RendersCachedResult);
            TestAssert.True(parsed.FieldReferences.Single().HasCachedResult, "Stored values retain source metadata.");
            TestAssert.Equal(locked ? DocxCounterDependencies.None : kind == "PAGE"
                ? DocxCounterDependencies.PageNumber : DocxCounterDependencies.PageCount,
                DocxFieldText.GetDependencies(document.BodyElements));
        }
    }

    public static void DocxLockedFieldsSurviveCellsStaticStoriesAndCancellation()
    {
        var body = new XElement(W + "p", Run("Body: "), Field("PAGE", "99", "1", true));
        var cell = new XElement(W + "p", Run("Cell: "), Field("NUMPAGES", "88", "true", false));
        var table = new XElement(W + "tbl",
            new XElement(W + "tblPr", new XElement(W + "tblW", new XAttribute(W + "w", 3000), new XAttribute(W + "type", "dxa"))),
            new XElement(W + "tblGrid", new XElement(W + "gridCol", new XAttribute(W + "w", 3000))),
            new XElement(W + "tr", new XElement(W + "tc",
                new XElement(W + "tcPr", new XElement(W + "tcW", new XAttribute(W + "w", 3000), new XAttribute(W + "type", "dxa"))), cell)));
        var header = new XElement(W + "hdr", new XElement(W + "p", Run("Header: "), Field("PAGE", "66", "on", false)));
        var footer = new XElement(W + "ftr", new XElement(W + "p", Run("Footer: "), Field("NUMPAGES", "55", "1", true)));
        string input = Write([body, table], header, footer);
        var document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
        string text = RenderText(document, OoxPdfDocxMarkupMode.Final);
        foreach (string expected in new[] { "Body: 99", "Cell: 88", "Header: 66", "Footer: 55" })
            TestAssert.Contains(Compact(expected), Compact(text));
        TestAssert.Equal(DocxCounterDependencies.None, DocxFieldText.GetDependencies(document.BodyElements));
        TestAssert.True(document.HeaderParagraphs.Concat(document.FooterParagraphs)
            .SelectMany(paragraph => paragraph.FieldReferences).All(field => !field.UsesPlaceholder && field.RendersCachedResult),
            "Static stories retain their stored field values.");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        using var source = File.OpenRead(input);
        using var output = new MemoryStream();
        TestAssert.Throws<OperationCanceledException>(() => OoxPdfConverter.Convert(source, output,
            new OoxPdfOptions { FontResolver = new PortableRenderingTests.PinnedResolver() }, cancellation.Token));
        TestAssert.Equal(0L, output.Length);
    }

    private static void CheckWord(string file)
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(Case("locked-page-fields-office.json")));
        var row = gold.RootElement.GetProperty("Controls").EnumerateArray().Single(row => row.GetProperty("Input").GetString() == file);
        TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(Case(file)))));
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxTests.ReadDocx(Case(file), mode);
            TestAssert.Equal(Compact(row.GetProperty("WordText").GetString()!), Compact(RenderText(document, mode)));
        }
    }

    private static string RenderText(DocxDocument document, OoxPdfDocxMarkupMode mode)
    {
        var geometry = mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
            : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
        var pages = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode, geometry).RenderPagesAsync(document, null, default).ToBlockingEnumerable();
        return string.Join(" ", pages.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)).Select(show => show.Text));
    }

    private static string Compact(string text) => string.Concat(text.Where(character => !char.IsWhiteSpace(character)));
    private static string Case(string file) => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes", file));
    private static XElement Run(string text) => new(W + "r", new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), text));

    private static IEnumerable<XElement> Field(string kind, string cached, string? locked, bool simple, bool styled = false)
    {
        var result = Run(cached);
        if (styled) result.AddFirst(new XElement(W + "rPr", new XElement(W + "b"), new XElement(W + "sz", new XAttribute(W + "val", 40))));
        if (simple)
        {
            yield return new XElement(W + "fldSimple", new XAttribute(W + "instr", " " + kind + " "),
                locked is null ? null : new XAttribute(W + "fldLock", locked), result);
            yield break;
        }
        yield return new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "begin"),
            locked is null ? null : new XAttribute(W + "fldLock", locked)));
        foreach (string part in new[] { " " + kind[..2], kind[2..] + " " })
            yield return new XElement(W + "r", new XElement(W + "instrText", new XAttribute(XNamespace.Xml + "space", "preserve"), part));
        yield return new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "separate")));
        yield return result;
        yield return new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "end")));
    }

    private static string Write(XElement[] elements, XElement? header = null, XElement? footer = null)
    {
        using var archive = ZipFile.OpenRead(Case("locked-page-fields.docx"));
        var parts = archive.Entries.ToDictionary(entry => entry.FullName, entry =>
        {
            using var reader = new StreamReader(entry.Open());
            return reader.ReadToEnd();
        });
        var doc = XDocument.Parse(parts["word/document.xml"]);
        var body = doc.Root!.Element(W + "body")!;
        var section = new XElement(body.Element(W + "sectPr")!);
        body.ReplaceNodes(elements, section);
        var relationships = XDocument.Parse(parts["word/_rels/document.xml.rels"]);
        var types = XDocument.Parse(parts["[Content_Types].xml"]);
        foreach (var (name, story) in new[] { ("header", header), ("footer", footer) })
        {
            if (story is null) continue;
            parts["word/" + name + ".xml"] = story.ToString();
            section.AddFirst(new XElement(W + (name + "Reference"), new XAttribute(W + "type", "default"), new XAttribute(R + "id", name)));
            relationships.Root!.Add(new XElement(relationships.Root.Name.Namespace + "Relationship",
                new XAttribute("Id", name), new XAttribute("Type", R.NamespaceName + "/" + name), new XAttribute("Target", name + ".xml")));
            types.Root!.Add(new XElement(types.Root.Name.Namespace + "Override", new XAttribute("PartName", "/word/" + name + ".xml"),
                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml." + name + "+xml")));
        }
        parts["word/document.xml"] = doc.ToString();
        parts["word/_rels/document.xml.rels"] = relationships.ToString();
        parts["[Content_Types].xml"] = types.ToString();
        return TestFixtures.WriteTempPackage(".docx", parts);
    }
}
