using System.IO.Compression;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Tests;

internal static class DocxMathTests
{
    public static void DocxMathNormalTextPropertiesRetainTypedDefaultsAndMalformedFallback()
    {
        foreach ((string property, bool? expected, bool unsupported) in new (string, bool?, bool)[]
        {
            ("", null, false), ("<m:nor/>", true, false), ("<m:nor m:val='1'/>", true, false),
            ("<m:nor m:val='true'/>", true, false), ("<m:nor m:val='0'/>", false, false),
            ("<m:nor m:val='false'/>", false, false), ("<m:nor m:val='unknown'/>", false, true),
            ("<m:nor/><m:nor/>", true, true), ("<m:nor extra='1'/>", true, true)
        })
        {
            DocxMathRow row = Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}'><m:r><m:rPr>" +
                property + "</m:rPr><m:t>sin</m:t></m:r></m:oMath>"));
            DocxMathRun run = (DocxMathRun)row.Children[0];
            TestAssert.Equal(expected, run.NormalTextValue);
            TestAssert.Equal(unsupported, run.HasUnsupportedProperties);
            TestAssert.Equal("sin", run.Text.Text);
            TestAssert.Equal(!unsupported, DocxMathRowLayout.TryGetRenderCodePoints(row, CancellationToken.None, out var points, out _));
            if (run.NormalText && !unsupported) { TestAssert.True(points.SequenceEqual("sin".Select(c => (int)c)),
                "Normal text owns its original alphabet instead of mathematical italic substitutions."); }
        }
    }

    public static void DocxMathArgumentPropertiesPreserveTypedSizeAndConservativeEligibility()
    {
        foreach ((string properties, int? size, bool unsupported, bool eligible) in new (string, int?, bool, bool)[]
        {
            ("", null, false, true), ("<m:argPr/>", null, false, true),
            ("<m:argPr><m:argSz m:val='0'/></m:argPr>", 0, false, true),
            ("<m:argPr><m:argSz m:val='-1'/></m:argPr>", -1, false, false),
            ("<m:argPr><m:argSz m:val='1'/></m:argPr>", 1, false, false),
            ("<m:argPr><m:argSz/></m:argPr>", 0, false, true),
            ("<m:argPr><m:argSz m:val='3'/></m:argPr>", null, true, false),
            ("<m:argPr><m:argSz m:val='no'/></m:argPr>", null, true, false),
            ("<m:argPr><m:argSz m:val='2147483648'/></m:argPr>", null, true, false),
            ("<m:argPr><m:argSz m:val='0' extra='1'/></m:argPr>", null, true, false),
            ("<m:argPr><m:argSz m:val='0'><m:unknown/></m:argSz></m:argPr>", null, true, false),
            ("<m:argPr><m:argSz m:val='0'/>text</m:argPr>", null, true, false),
            ("<m:argPr><m:argSz m:val='0'/><m:argSz m:val='0'/></m:argPr>", null, true, false),
            ("<m:argPr><m:unknown/></m:argPr>", null, true, false),
            ("<m:argPr/><m:argPr/>", null, true, false),
            ("<m:argPr/><m:ctrlPr/>", null, true, false),
            ("<m:ctrlPr/>", null, true, false)
        })
        {
            DocxMathRow row = Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}'>" + properties +
                "<m:r><m:rPr><m:sty m:val='p'/></m:rPr><m:t>x</m:t></m:r></m:oMath>"));
            TestAssert.Equal(size, row.ArgumentProperties?.Size);
            TestAssert.Equal(unsupported, row.ArgumentProperties?.HasUnsupportedProperties ?? false);
            TestAssert.Equal(eligible, row.HasSupportedProperties);
            TestAssert.Equal(eligible, DocxMathRowLayout.TryGetRenderCodePoints(row, CancellationToken.None, out _, out _));
            TestAssert.Equal("x", RowText(row));
            if (properties.Length != 0) { TestAssert.True(row.PropertiesXml.Count != 0, "Raw authored controls remain available for fallback."); }
        }
        TestAssert.True(!new DocxMathRow([]) { PropertiesXml = ["<unknown/>"] }.HasSupportedProperties,
            "Unparsed synthetic properties cannot silently acquire neutral semantics.");
    }

    public static void DocxMathNeutralArgumentPropertiesShareLiteralAndStructuredDiscovery()
    {
        foreach (string body in new[]
        {
            "<m:r><m:t>x</m:t></m:r>",
            "<m:sSubSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub>{0}<m:r><m:t>i</m:t></m:r></m:sub><m:sup>{0}<m:r><m:t>2</m:t></m:r></m:sup></m:sSubSup>",
            "<m:f><m:num>{0}<m:r><m:t>a</m:t></m:r></m:num><m:den>{0}<m:r><m:t>b</m:t></m:r></m:den></m:f>"
        })
        foreach (string properties in new[] { "<m:argPr/>", "<m:argPr><m:argSz m:val='0'/></m:argPr>" })
        {
            DocxMathRow row = Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}'>" + properties +
                body.Replace("{0}", properties, StringComparison.Ordinal) + "</m:oMath>"));
            foreach (bool display in new[] { false, true })
            {
                var equation = new DocxEquation(row, 0, 0, display, null, null, null);
                TestAssert.True(DocxMathLayout.TryGetRuns(equation, CancellationToken.None, out var runs),
                    "Neutral argument properties share the same literal/structured source discovery.");
                TestAssert.True(runs.Count != 0, "Source meaning must remain discoverable.");
                DocxMathRow unknown = row with { PropertiesXml = ["<unknown/>"], ArgumentProperties = null };
                TestAssert.True(!DocxMathLayout.TryGetRuns(equation with { Content = unknown }, CancellationToken.None, out _),
                    "Unknown root controls preserve the whole-equation fallback.");
            }
        }
    }

    public static void DocxMathPinnedControlsRetainStructureAndProseBoundaries()
    {
        foreach (string kind in new[] { "literal", "scripts", "fraction" })
        foreach (string placement in new[] { "inline", "after" })
        {
            DocxDocument document = ReadFixture(kind, placement);
            DocxParagraph paragraph = document.Paragraphs.Single(p => p.Equations.Count != 0);
            DocxEquation equation = paragraph.Equations.Single();
            TestAssert.Equal(placement == "after", equation.Display);
            TestAssert.Equal(placement == "inline" ? 1 : 0, equation.TextRunIndex);
            TestAssert.Equal(placement == "inline" ? 1 : 0, equation.SourceRunIndex);
            TestAssert.Equal("Before formula: Following prose.", string.Concat(document.Paragraphs.SelectMany(p => p.Runs).Select(run => run.Text)));
            if (kind == "literal")
            {
                TestAssert.Equal("α+β=γ", string.Concat(equation.Content.Children.Cast<DocxMathRun>().Select(run => run.Text.Text)));
                TestAssert.True(equation.Content.Children.Cast<DocxMathRun>().All(run => run.Text.FontFamily == "Cambria Math" && run.Text.FontSize == 12d), "Math run formatting must use the existing style cascade.");
            }
            else if (kind == "scripts")
            {
                var scripts = (DocxMathScripts)equation.Content.Children.Single();
                TestAssert.Equal("x", RowText(scripts.Base));
                TestAssert.Equal("i", RowText(TestAssert.NotNull(scripts.Subscript)));
                TestAssert.Equal("2", RowText(TestAssert.NotNull(scripts.Superscript)));
            }
            else
            {
                var fraction = (DocxMathFraction)equation.Content.Children.Single();
                TestAssert.Equal("a+b", RowText(fraction.Numerator));
                TestAssert.Equal("c", RowText(fraction.Denominator));
            }
        }
    }

    public static void DocxMathReaderPreservesMultiplePositionsFormattingAndRevisionVisibility()
    {
        XElement document = XElement.Parse("""
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
              <w:body><w:p>
                <w:r><w:t>A</w:t></w:r>
                <m:oMath><m:r><w:rPr><w:b/><w:sz w:val="36"/><w:color w:val="123456"/></w:rPr><m:rPr><m:sty m:val="bi"/><m:scr m:val="double-struck"/></m:rPr><m:t>x</m:t></m:r></m:oMath>
                <m:oMath><m:r><m:t>y</m:t></m:r></m:oMath>
                <w:r><w:t>B</w:t></w:r>
                <w:ins w:id="1" w:author="Author"><m:oMath><m:r><m:t>i</m:t></m:r></m:oMath></w:ins>
                <w:del w:id="2" w:author="Author"><m:oMath><m:r><m:t>d</m:t></m:r></m:oMath></w:del>
                <w:r><w:t>C</w:t></w:r>
              </w:p></w:body>
            </w:document>
            """);
        string input = WriteDocument(document);
        try
        {
            foreach (OoxPdfDocxMarkupMode mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.Original, OoxPdfDocxMarkupMode.AllMarkup })
            {
                DocxParagraph paragraph = DocxTests.ReadDocx(input, mode).Paragraphs.Single();
                TestAssert.Equal("ABC", string.Concat(paragraph.Runs.Select(run => run.Text)));
                TestAssert.Equal("0,1,2", string.Join(',', paragraph.Runs.Select(run => run.SourceRunIndex)));
                TestAssert.Equal(mode == OoxPdfDocxMarkupMode.AllMarkup ? 4 : 3, paragraph.Equations.Count);
                TestAssert.Equal("1,1,2", string.Join(',', paragraph.Equations.Take(3).Select(equation => equation.TextRunIndex)));
                var run = (DocxMathRun)paragraph.Equations[0].Content.Children.Single();
                TestAssert.Equal(18d, run.Text.FontSize);
                TestAssert.Equal("123456", run.Text.ColorHex);
                TestAssert.True(run.Text.Bold, "Direct math-run formatting must be retained.");
                TestAssert.Equal("bi", run.MathStyleValue);
                TestAssert.Equal("double-struck", run.ScriptValue);
                TestAssert.Contains("double-struck", run.PropertiesXml.Single());
                TestAssert.Equal(mode == OoxPdfDocxMarkupMode.Original ? "d" : "i", RowText(paragraph.Equations[2].Content));
                TestAssert.NotNull(paragraph.Equations[2].Revision);
            }
        }
        finally { File.Delete(input); }
    }

    public static void DocxMathParserRetainsUnsupportedAndMalformedMeaningWithoutFlattening()
    {
        XElement source = XElement.Parse("""
            <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
              <m:rad><m:e><m:r><m:t>root</m:t></m:r></m:e></m:rad>
              <m:f><m:num><m:r><m:t>one</m:t></m:r></m:num></m:f>
              <m:f><m:num/><m:num/><m:den/></m:f>
            </m:oMath>
            """);
        DocxMathRow row = Parse(source);
        TestAssert.Equal(3, row.Children.Count);
        TestAssert.True(row.Children.All(node => node is DocxMathUnsupported), "Incomplete constructs must remain explicit fallbacks.");
        var root = (DocxMathUnsupported)row.Children[0];
        TestAssert.Equal("rad", root.ElementName);
        TestAssert.Contains("root", TestAssert.NotNull(root.SourceXml));
        var missingDenominator = (DocxMathUnsupported)row.Children[1];
        TestAssert.Equal("num", missingDenominator.Children.Single().ElementName);
        TestAssert.Equal("one", RowText((DocxMathRow)missingDenominator.Children.Single().Content));
    }

    public static void DocxMathParserUsesIterativeTraversalAndObservesCancellation()
    {
        XElement current = new(MathNamespace + "r", new XElement(MathNamespace + "t", "x"));
        for (int i = 0; i < 2048; i++)
            current = new XElement(MathNamespace + "f", new XElement(MathNamespace + "num", current), new XElement(MathNamespace + "den"));
        XElement equation = new(MathNamespace + "oMath", current);
        DocxMathNode node = Parse(equation).Children.Single();
        int count = 0;
        while (node is DocxMathFraction fraction) { count++; node = fraction.Numerator.Children.Single(); }
        TestAssert.Equal(2048, count);
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => Parse(equation, cancelled.Token));
        using var duringRead = new CancellationTokenSource();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathParser.Read(equation, (_, text) =>
        {
            duringRead.Cancel();
            return PlainRun(text);
        }, duringRead.Token));
    }

    public static void DocxMathReaderUsesTheSameStructureForCellsAndStaticStories()
    {
        XElement document = XElement.Parse("""
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <w:body><w:p><m:oMath><m:r><m:t>body</m:t></m:r></m:oMath></w:p>
                <w:tbl><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:p><m:oMathPara><m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr><m:oMath><m:r><m:t>cell</m:t></m:r></m:oMath></m:oMathPara></w:p></w:tc></w:tr></w:tbl>
                <w:sectPr><w:headerReference w:type="default" r:id="mathHeader"/><w:footerReference w:type="default" r:id="mathFooter"/></w:sectPr>
              </w:body>
            </w:document>
            """);
        var parts = FixtureParts();
        parts["word/document.xml"] = Utf8(document.ToString());
        XNamespace relations = "http://schemas.openxmlformats.org/package/2006/relationships";
        XElement relationships = XElement.Parse(System.Text.Encoding.UTF8.GetString(parts["word/_rels/document.xml.rels"]));
        XNamespace contentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
        XElement types = XElement.Parse(System.Text.Encoding.UTF8.GetString(parts["[Content_Types].xml"]));
        foreach (string story in new[] { "header", "footer" })
        {
            string file = story + "-math.xml";
            relationships.Add(new XElement(relations + "Relationship", new XAttribute("Id", story == "header" ? "mathHeader" : "mathFooter"),
                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" + story), new XAttribute("Target", file)));
            types.Add(new XElement(contentTypes + "Override", new XAttribute("PartName", "/word/" + file),
                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml." + story + "+xml")));
            parts["word/" + file] = Utf8($"<w:{(story == "header" ? "hdr" : "ftr")} xmlns:w='{WordprocessingNamespace}' xmlns:m='{MathNamespace}'><w:p><m:oMath><m:r><m:t>{story}</m:t></m:r></m:oMath></w:p></w:{(story == "header" ? "hdr" : "ftr")}>");
        }
        parts["word/_rels/document.xml.rels"] = Utf8(relationships.ToString());
        parts["[Content_Types].xml"] = Utf8(types.ToString());
        string input = TestFixtures.WriteTempPackage(".docx", parts);
        try
        {
            DocxDocument result = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
            TestAssert.Equal("body", RowText(result.Paragraphs.Single().Equations.Single().Content));
            DocxEquation cell = result.Tables.Single().Rows.Single().Cells.Single().Paragraphs.Single().Equations.Single();
            TestAssert.Equal("cell", RowText(cell.Content));
            TestAssert.True(cell.Display, "Cell display structure must be retained.");
            TestAssert.Equal("left", cell.AlignmentValue);
            TestAssert.Contains("left", TestAssert.NotNull(cell.PropertiesXml));
            TestAssert.Equal("header", RowText(result.HeaderParagraphs.Single().Equations.Single().Content));
            TestAssert.Equal("footer", RowText(result.FooterParagraphs.Single().Equations.Single().Content));
        }
        finally { File.Delete(input); }
    }

    private static string RowText(DocxMathRow row) => string.Concat(row.Children.Cast<DocxMathRun>().Select(run => run.Text.Text));
    private static DocxTextRun PlainRun(string text) => new(text, 12d, null, false, false, false, null, "Cambria Math");
    private static DocxMathRow Parse(XElement element, CancellationToken token = default) => DocxMathParser.Read(element, (_, text) => PlainRun(text), token);
    private static DocxDocument ReadFixture(string kind, string placement) => DocxTests.ReadDocx(
        Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", $"docx-omml-{kind}-{placement}.docx")), OoxPdfDocxMarkupMode.Final);

    private static string WriteDocument(XElement document)
    {
        var parts = FixtureParts();
        parts["word/document.xml"] = Utf8(document.ToString(SaveOptions.DisableFormatting));
        return TestFixtures.WriteTempPackage(".docx", parts);
    }

    private static byte[] Utf8(string value) => System.Text.Encoding.UTF8.GetBytes(value);
    private static Dictionary<string, byte[]> FixtureParts()
    {
        string fixture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "docx-omml-literal-inline.docx"));
        using ZipArchive archive = ZipFile.OpenRead(fixture);
        var parts = new Dictionary<string, byte[]>();
        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            using Stream stream = entry.Open();
            using var bytes = new MemoryStream();
            stream.CopyTo(bytes);
            parts.Add(entry.FullName, bytes.ToArray());
        }
        return parts;
    }
}
