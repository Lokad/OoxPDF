using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxNumberingTests
{
    public static void DocxListFormatsResolveReferencedLevelsAndConcreteOverrides()
    {
        string levels = FormatLevel(0, "upperRoman", 4, "%1.") +
            FormatLevel(1, "lowerLetter", 27, "%1.%2)") + FormatLevel(2, "decimalZero", 3, "%1.%2.%3");
        string body = string.Concat(new[] { 0, 1, 1, 2, 0, 1, 2 }.Select(level => FormatParagraph(level)));
        Check([], ["IV.", "IV.aa)", "IV.bb)", "IV.bb.03", "V.", "V.aa)", "V.aa.03"]);
        Check([Override(0, FormatLevel(0, "upperLetter", 26, "%1."))],
            ["Z.", "Z.aa)", "Z.bb)", "Z.bb.03", "AA.", "AA.aa)", "AA.aa.03"]);
        Check([Override(0, FormatLevel(0, "upperLetter", 26, "%1.")), Override(2, FormatLevel(2, "lowerRoman", 9, "%1/%2/%3"))],
            ["Z.", "Z.aa)", "Z.bb)", "Z/bb/ix", "AA.", "AA.aa)", "AA/aa/ix"]);

        void Check(string[] overrides, string[] expected)
        {
            var document = ReadFormatNumbering(levels, body, string.Concat(overrides));
            AssertFormatLabels(document, expected);
        }
        static string Override(int level, string definition) => $"<w:lvlOverride w:ilvl=\"{level}\">{definition}</w:lvlOverride>";
    }

    public static void DocxListFormatTransitionsKeepStartsAndCounterRestarts()
    {
        foreach (var row in new (string Format, int Start, string[] Expected)[]
        {
            ("decimal", 9, ["9)", "10)", "11)"]), ("decimalZero", 9, ["09)", "10)", "11)"]),
            ("lowerRoman", 3, ["iii)", "iv)", "v)"]), ("upperRoman", 8, ["VIII)", "IX)", "X)"]),
            ("upperRoman", 39, ["XXXIX)", "XL)", "XLI)"]),
            ("lowerLetter", 25, ["y)", "z)", "aa)"]), ("upperLetter", 52, ["ZZ)", "AAA)", "BBB)"]),
            ("futureFormat", 17, ["17)", "18)", "19)"]), ("upperRoman", 4000, ["4000)", "4001)", "4002)"])
        })
        {
            AssertFormatLabels(ReadFormatNumbering(FormatLevel(0, row.Format, row.Start, "%1)"),
                string.Concat(Enumerable.Repeat(FormatParagraph(0), 3))), row.Expected);
        }
        string levels = FormatLevel(0, "upperRoman", 3, "%1.") + FormatLevel(1, "lowerLetter", 2, "%1/%2");
        string start = "<w:lvlOverride w:ilvl=\"1\"><w:startOverride w:val=\"26\"/></w:lvlOverride>";
        AssertFormatLabels(ReadFormatNumbering(levels, string.Concat(new[] { 0, 1, 1, 0, 1 }.Select(level => FormatParagraph(level))), start),
            ["III.", "III/z", "III/aa", "IV.", "IV/z"]);
    }

    public static void DocxListLegalNumberingOwnsAllReferencedFormats()
    {
        foreach (string legal in new[] { "", "<w:isLgl/>", "<w:isLgl w:val=\"true\"/>", "<w:isLgl w:val=\"0\"/>" })
        {
            string levels = FormatLevel(0, "upperRoman", 4, "%1.") + FormatLevel(1, "lowerLetter", 27, "%1/%2") +
                FormatLevel(2, "decimalZero", 3, "%1/%2/%3", legal);
            bool enabled = legal.Contains("<w:isLgl/>", StringComparison.Ordinal) || legal.Contains("true", StringComparison.Ordinal);
            AssertFormatLabels(ReadFormatNumbering(levels, FormatParagraph(0) + FormatParagraph(1) + FormatParagraph(2)),
                ["IV.", "IV/aa", enabled ? "4/27/3" : "IV/aa/03"]);
        }
        string definitions = FormatLevel(0, "lowerRoman", 9, "%1/%2/%9 literal %x") +
            FormatLevel(1, "upperLetter", 2, "%1-%2", "<w:isLgl/>");
        string replacement = "<w:lvlOverride w:ilvl=\"1\">" + FormatLevel(1, "upperLetter", 2, "%1-%2", "<w:isLgl w:val=\"false\"/>") + "</w:lvlOverride>";
        AssertFormatLabels(ReadFormatNumbering(definitions, FormatParagraph(0) + FormatParagraph(1), replacement),
            ["ix// literal %x", "ix-B"]);
    }

    public static void DocxListFormatsKeepOriginalAndCurrentRevisionCounters()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxInlineCompositionTests.ReadMutation("CompositionProbes/numbering-mark-math-body", mode,
                false, false, (_, _, _) => { }, packageChange: archive =>
                {
                    XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                    var entry = archive.GetEntry("word/numbering.xml")!;
                    XDocument xml; using (var input = entry.Open()) xml = XDocument.Load(input);
                    foreach (var format in xml.Descendants(w + "numFmt")) format.SetAttributeValue(w + "val", "upperRoman");
                    entry.Delete(); using var output = archive.CreateEntry("word/numbering.xml").Open(); xml.Save(output);
                });
            var labels = DocxBlockTraversal.EnumerateSourceParagraphs(document).Where(p => p.ListLabel is not null)
                .Select(p => p.ListLabel!).ToArray();
            TestAssert.True(labels.Select(label => label.Text).SequenceEqual(new[] { "I.", "II.", "III." }), "Current counters resolve Roman labels before layout.");
            TestAssert.True(labels.Select(label => label.NumberingChange?.PreviousText).SequenceEqual(
                mode == OoxPdfDocxMarkupMode.AllMarkup ? new string?[] { null, "I.", "II." } : [null, null, null]),
                "The original timeline uses the same format and its own counter values.");
        }
    }

    private static void AssertFormatLabels(DocxDocument document, string[] expected) =>
        TestAssert.True(document.Paragraphs.Select(p => p.ListLabel?.Text).SequenceEqual(expected),
            "Expected " + string.Join(" | ", expected) + "; actual " + string.Join(" | ", document.Paragraphs.Select(p => p.ListLabel?.Text)));

    private static string FormatLevel(int level, string format, int start, string text, string extra = "") =>
        $"<w:lvl w:ilvl=\"{level}\"><w:start w:val=\"{start}\"/><w:numFmt w:val=\"{format}\"/><w:lvlText w:val=\"{text}\"/>{extra}</w:lvl>";

    private static string FormatParagraph(int level) => $"<w:p><w:pPr><w:numPr><w:ilvl w:val=\"{level}\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr><w:r><w:t>Independent list item</w:t></w:r></w:p>";

    private static DocxDocument ReadFormatNumbering(string levels, string body, string overrides = "")
    {
        using var input = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>""",
            ["_rels/.rels"] = """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="numbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>""",
            ["word/numbering.xml"] = "<w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:abstractNum w:abstractNumId=\"7\">" + levels + "</w:abstractNum><w:num w:numId=\"1\"><w:abstractNumId w:val=\"7\"/>" + overrides + "</w:num></w:numbering>",
            ["word/document.xml"] = "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" + body + "</w:body></w:document>"
        });
        return new DocxReader().Read(OoxPackage.Open(input, default), null, default, OoxPdfDocxMarkupMode.Final);
    }
}
