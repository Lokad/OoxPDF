using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static class DocxStyleToggleTests
{
    private static readonly XNamespace W = OoxNamespaces.WordprocessingNamespace;
    private static readonly string[] Toggles = ["b", "i", "bCs", "iCs", "caps", "smallCaps", "strike", "vanish"];

    public static void DocxStyleTogglesComposeIndependentLevels()
    {
        // Independently specified truth table, including Word's enabled-default
        // behavior. Missing declarations inherit the document default at a level.
        (bool? Default, bool? Table, bool? Paragraph, bool? Character, bool Expected)[] cases = [
            (null, null, true, true, false), (false, null, true, false, true),
            (false, null, false, true, true), (false, null, false, false, false),
            (false, true, true, null, false), (false, true, true, true, true),
            (true, null, null, null, true), (true, null, false, null, false),
            (true, null, false, false, true), (true, null, false, true, false),
            (true, false, false, false, false), (true, false, false, true, true)
        ];
        foreach (string property in Toggles)
        foreach (var sample in cases)
        {
            var run = ReadRun(Parts(property, sample.Default, sample.Table, sample.Paragraph, sample.Character));
            TestAssert.True(Flag(run, property) == sample.Expected,
                $"Toggle {property} disagrees for defaults/table/paragraph/character {sample}.");
        }
    }

    public static void DocxStyleTogglesKeepBasedOnAssignmentAndDirectOverrides()
    {
        foreach (string property in Toggles)
        {
            TestAssert.True(Flag(ReadRun(Parts(property, paragraph: true, paragraphBase: true)), property),
                "Repeated declarations within one paragraph-style level do not cancel.");
            TestAssert.True(!Flag(ReadRun(Parts(property, paragraph: false, paragraphBase: true)), property),
                "The derived paragraph style overrides its parent at the same level.");
            TestAssert.True(!Flag(ReadRun(Parts(property, paragraph: true, character: true, characterBase: true)), property),
                "A character basedOn chain supplies one effective toggle, not two.");
            TestAssert.True(!Flag(ReadRun(Parts(property, table: true, tableBase: true, paragraph: true)), property),
                "A table basedOn chain supplies one effective toggle, not two.");
            foreach (bool direct in new[] { false, true })
                TestAssert.True(Flag(ReadRun(Parts(property, defaultValue: true, paragraph: false, character: false, direct: direct)), property) == direct,
                    "Direct formatting assigns the property absolutely after style composition.");
        }
    }

    public static void DocxStyleTogglesRecoverVisibleTextAcrossContainers()
    {
        foreach (string container in new[] { "body", "cell", "header" })
        {
            var actual = Paint(Parts("vanish", paragraph: true, character: true, container: container));
            var flattened = Paint(Parts("vanish", paragraph: true, character: true, direct: false, container: container));
            TestAssert.Contains("TOGGLE", actual.Text);
            TestAssert.Equal(flattened.Text, actual.Text);
            TestAssert.Equal(flattened.Content, actual.Content);
            TestAssert.Equal(0, actual.Diagnostics.Count);
        }
        var hidden = Paint(Parts("vanish", paragraph: true));
        TestAssert.True(!hidden.Text.Contains("TOGGLE", StringComparison.Ordinal), "One enabled style still hides text.");
    }

    public static void DocxStyleTogglesPreserveOrdinaryFormattingAndSourceOffsets()
    {
        var parts = Parts("b", paragraph: true, character: true);
        var styles = XDocument.Parse(parts["word/styles.xml"]);
        var paragraph = styles.Root!.Elements(W + "style").Single(style => (string?)style.Attribute(W + "styleId") == "Para").Element(W + "rPr")!;
        var character = styles.Root.Elements(W + "style").Single(style => (string?)style.Attribute(W + "styleId") == "Char").Element(W + "rPr")!;
        foreach (var owner in new[] { paragraph, character })
            owner.Add(new XElement(W + "u", new XAttribute(W + "val", "single")), new XElement(W + "dstrike"));
        paragraph.Add(new XElement(W + "sz", new XAttribute(W + "val", "28")), new XElement(W + "color", new XAttribute(W + "val", "112233")));
        character.Add(new XElement(W + "sz", new XAttribute(W + "val", "32")), new XElement(W + "color", new XAttribute(W + "val", "445566")));
        parts["word/styles.xml"] = styles.ToString();
        var run = ReadRun(parts);
        TestAssert.True(!run.Bold && run.Underline && run.DoubleStrike, "Only documented toggle properties compose between levels.");
        TestAssert.Equal(16d, run.FontSize);
        TestAssert.Equal("445566", run.ColorHex);
        TestAssert.Equal(0, run.SourceRunIndex);
        TestAssert.Equal(0, run.SourceTextOffsetInRun);
    }

    private static bool Flag(DocxTextRun run, string property) => property switch {
        "b" or "bCs" => run.Bold, "i" or "iCs" => run.Italic,
        "caps" => run.AllCaps, "smallCaps" => run.SmallCaps,
        "strike" => run.Strike, "vanish" => run.Hidden,
        _ => throw new ArgumentOutOfRangeException(nameof(property))
    };

    private static DocxTextRun ReadRun(Dictionary<string, string> parts)
    {
        using var input = TestFixtures.CreateZipPackage(parts);
        using var package = OoxPackage.Open(input, default);
        var document = new DocxReader().Read(package, null, default, OoxPdfDocxMarkupMode.Final);
        var paragraph = document.Tables.Count == 0 ? document.Paragraphs[0] : document.Tables[0].Rows[0].Cells[0].Paragraphs[0];
        return paragraph.Runs.Single();
    }

    private static (string Content, string Text, List<OoxPdfDiagnostic> Diagnostics) Paint(Dictionary<string, string> parts)
    {
        using var input = TestFixtures.CreateZipPackage(parts);
        using var package = OoxPackage.Open(input, default);
        var diagnostics = new List<OoxPdfDiagnostic>();
        var document = new DocxReader().Read(package, diagnostics.Add, default, OoxPdfDocxMarkupMode.Final);
        var pages = new DocxRenderer(new PortableRenderingTests.PinnedResolver())
            .RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
        TestAssert.Equal(1, pages.Length);
        return (pages[0].Content, DocxPageFootnoteNumberingTests.DecodeText(pages[0]), diagnostics);
    }

    private static Dictionary<string, string> Parts(string property, bool? defaultValue = null,
        bool? table = null, bool? paragraph = null, bool? character = null, bool? direct = null,
        bool? paragraphBase = null, bool? characterBase = null, bool? tableBase = null, string container = "body")
    {
        string Pr(bool? value) => value is { } flag ? $"<w:{property} w:val=\"{(flag ? "1" : "0")}\"/>" : "";
        string Style(string type, string id, bool? value, string? parent = null) =>
            $"<w:style w:type=\"{type}\" w:styleId=\"{id}\"><w:name w:val=\"{id}\"/>" +
            (parent is null ? "" : $"<w:basedOn w:val=\"{parent}\"/>") + $"<w:rPr>{Pr(value)}</w:rPr></w:style>";
        bool cell = table is not null || tableBase is not null || container == "cell";
        string text = property is "bCs" or "iCs" ? "\u0628" : "TOGGLE";
        string payload = $"<w:p><w:pPr><w:pStyle w:val=\"Para\"/></w:pPr><w:r><w:rPr><w:rStyle w:val=\"Char\"/>{Pr(direct)}</w:rPr><w:t>{text}</w:t></w:r></w:p>";
        string body = cell ? "<w:tbl><w:tblPr><w:tblStyle w:val=\"Table\"/><w:tblW w:w=\"4000\" w:type=\"dxa\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"4000\"/></w:tblGrid><w:tr><w:tc>" + payload + "</w:tc></w:tr></w:tbl>" : payload;
        if (container == "header") body = "<w:p><w:r><w:t>BODY</w:t></w:r></w:p>";
        var parts = new Dictionary<string, string> {
            ["[Content_Types].xml"] = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/><Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>" + (container == "header" ? "<Override PartName=\"/word/header.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/>" : "") + "</Types>",
            ["_rels/.rels"] = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"main\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>",
            ["word/document.xml"] = $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{OoxNamespaces.RelationshipsNamespace}\"><w:body>{body}<w:sectPr>" + (container == "header" ? "<w:headerReference w:type=\"default\" r:id=\"header\"/>" : "") + "</w:sectPr></w:body></w:document>",
            ["word/_rels/document.xml.rels"] = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"styles\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>" + (container == "header" ? "<Relationship Id=\"header\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header.xml\"/>" : "") + "</Relationships>",
            ["word/styles.xml"] = $"<w:styles xmlns:w=\"{W}\"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val=\"24\"/>{Pr(defaultValue)}</w:rPr></w:rPrDefault></w:docDefaults>" +
                (paragraphBase is null ? "" : Style("paragraph", "ParaBase", paragraphBase)) + Style("paragraph", "Para", paragraph, paragraphBase is null ? null : "ParaBase") +
                (characterBase is null ? "" : Style("character", "CharBase", characterBase)) + Style("character", "Char", character, characterBase is null ? null : "CharBase") +
                (tableBase is null ? "" : Style("table", "TableBase", tableBase)) + Style("table", "Table", table, tableBase is null ? null : "TableBase") + "</w:styles>"
        };
        if (container == "header") parts["word/header.xml"] = $"<w:hdr xmlns:w=\"{W}\">{payload}</w:hdr>";
        return parts;
    }
}
