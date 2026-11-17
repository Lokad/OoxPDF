using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxMathCenterGroupTests
{
    public static void DocxMathCenterGroupRetainsLocalAndGlobalAlignmentSources()
    {
        foreach (var (local, global, expected) in new (string?, string?, DocxMathDisplayAlignment?)[]
        {
            (null, null, DocxMathDisplayAlignment.CenterGroup),
            (null, "center", DocxMathDisplayAlignment.Center),
            (null, "left", DocxMathDisplayAlignment.Left),
            ("right", "left", DocxMathDisplayAlignment.Right),
            ("centerGroup", "center", DocxMathDisplayAlignment.CenterGroup),
            ("center", "centerGroup", DocxMathDisplayAlignment.Center),
            (null, "future", DocxMathDisplayAlignment.Center),
            ("future", "center", null)
        })
        {
            var document = Read(local, global);
            var paragraph = Paragraph(document);
            TestAssert.Equal(global, document.Settings.MathParagraphAlignmentValue);
            TestAssert.True(paragraph.Equations.All(e => e.AlignmentValue == local && e.DefaultAlignmentValue == global),
                "Local absence and the raw document default remain independently inspectable.");
            TestAssert.Equal(expected, DocxMathParagraphPolicy.ResolveParagraphDisplayAlignment(paragraph, default));
        }
    }

    public static void DocxMathCenterGroupMeasuresMaximumRowsAndEndsAtProse()
    {
        foreach (bool prose in new[] { false, true })
        {
            var paragraph = Paragraph(Read("centerGroup", prose: prose));
            TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineContent(paragraph, Prepare(paragraph), new Measurer(),
                12d, 200d, 200d, 9d, 16d, 1d, default, out var measured), "The known display rows must measure.");
            var rows = TestAssert.NotNull(measured);
            TestAssert.Equal(prose ? 3 : 2, rows.Count);
            TestAssert.Equal(100d, rows[0].DisplayGroupAdvance);
            TestAssert.Equal(prose ? 40d : 100d, rows[^1].DisplayGroupAdvance);
            if (prose) TestAssert.True(rows[1].DisplayAlignment is null && rows[1].DisplayGroupAdvance is null,
                "Prose owns its ordinary row and separates the two display-group frames.");
        }
        var centered = Paragraph(Read("center"));
        TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineContent(centered, Prepare(centered), new Measurer(),
            12d, 200d, 200d, 9d, 16d, 1d, default, out var individual), "Explicit center remains supported.");
        TestAssert.True(TestAssert.NotNull(individual).All(row => row.DisplayGroupAdvance is null),
            "Individual row centering does not acquire a group-width constraint.");
    }

    public static void DocxMathCenterGroupPlacesBodyCellsAndFootersWithoutMovingBaselines()
    {
        foreach (string container in new[] { "body", "cell", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
        {
            var grouped = Layout(Read("centerGroup", container: container), mode);
            var centered = Layout(Read("center", container: container), mode);
            TestAssert.True(grouped.Length == 2 && centered.Length == 2,
                $"Expected two measured math rows in {container}/{mode}, got {grouped.Length}/{centered.Length}.");
            TestAssert.True(Math.Abs(grouped[0].X - grouped[1].X) < .000001,
                "Containers place both rows inside the same centered maximum-width group.");
            TestAssert.True(Math.Abs(centered[1].X - centered[0].X - (centered[0].Width - centered[1].Width) / 2d) < .000001,
                "Explicit center preserves independent row centers in the same container.");
            for (int i = 0; i < grouped.Length; i++)
                TestAssert.True(grouped[i].BaselineY == centered[i].BaselineY && grouped[i].LineHeight == centered[i].LineHeight &&
                    grouped[i].InlineMath.Count == centered[i].InlineMath.Count,
                    "Changing horizontal group justification preserves occupied vertical geometry and equation ownership.");
        }
    }

    public static void DocxMathCenterGroupCounterResumptionReusesMeasuredFrames()
    {
        var plain = Layout(Read("centerGroup"));
        var counter = Layout(Read("centerGroup", counter: true));
        TestAssert.Equal(2, counter.Length);
        TestAssert.True(Math.Abs(counter[0].X - counter[1].X) < .000001,
            "Counter-driven placement must preserve the common display-group origin, not two equally wrong row centers.");
        for (int i = 0; i < counter.Length; i++)
            TestAssert.True(Math.Abs(plain[i].X - counter[i].X) < .000001 && plain[i].Width == counter[i].Width,
                "PAGE forces source-cursor resumption without losing the already measured group frame.");
    }

    private static DocxTextLineLayout[] Layout(DocxDocument document,
        OoxPdfDocxMarkupGeometryMode mode = OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
    {
        var layout = new DocxLayoutEngine(mode).Create(document, new Measurer(), default, mathEquations: Prepare(Paragraph(document)));
        return layout.Pages.SelectMany(page => page.Items.Concat(page.StaticTextLines)).SelectMany(item => item switch
        {
            DocxTextLineLayout line => new[] { line },
            DocxTableRowLayout row => row.Cells.SelectMany(cell => cell.TextLines),
            _ => Enumerable.Empty<DocxTextLineLayout>()
        }).Where(line => line.InlineMath.Count != 0).ToArray();
    }

    private static DocxParagraph Paragraph(DocxDocument document) =>
        DocxBlockTraversal.EnumerateSourceParagraphs(document)
            .Distinct<DocxParagraph>(ReferenceEqualityComparer.Instance)
            .Single(paragraph => paragraph.Equations.Count == 3);

    private static Dictionary<DocxEquation, DocxPreparedMathParagraph> Prepare(DocxParagraph paragraph) =>
        paragraph.Equations.Select((equation, index) => (equation, width: new[] { 30d, 70d, 40d }[index]))
            .ToDictionary(pair => pair.equation, pair => new DocxPreparedMathParagraph(pair.equation,
                new(pair.width, 6d, 1d, []), 9d, 16d));

    private sealed class Measurer : IDocxTextMeasurer, IDocxStaticTextMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => text.Length * 5d;
        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => fontSize * .8d;
        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => fontSize * .2d;
    }

    private static DocxDocument Read(string? local, string? global = null, bool prose = false,
        string container = "body", bool counter = false)
    {
        string Equation(string text) => "<m:oMathPara>" + (local is null ? "" : $"<m:oMathParaPr><m:jc m:val=\"{local}\"/></m:oMathParaPr>") +
            $"<m:oMath><m:r><m:t>{text}</m:t></m:r></m:oMath></m:oMathPara>";
        string paragraph = "<w:p>" + Equation("a") + Equation("b") + "<w:r><w:br/></w:r>" +
            (prose ? "<w:r><w:t>Between groups.</w:t><w:br/></w:r>" : "") + Equation("c") +
            (counter ? "<w:r><w:br/></w:r><w:fldSimple w:instr=\"PAGE\"><w:r><w:t>1</w:t></w:r></w:fldSimple>" : "") + "</w:p>";
        const string namespaces = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"";
        bool staticStory = container is "header" or "footer" or "header-cell";
        string story = container == "header-cell" ? "header" : container;
        bool note = container is "footnote" or "endnote";
        string Cell(string content) => "<w:tbl><w:tblPr><w:tblW w:w=\"9360\" w:type=\"dxa\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"9360\"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w=\"9360\" w:type=\"dxa\"/></w:tcPr>" + content + "</w:tc></w:tr></w:tbl>";
        string Box(bool floating)
        {
            string properties = floating ? "<wp:positionH relativeFrom=\"margin\"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom=\"margin\"><wp:posOffset>0</wp:posOffset></wp:positionV>" : "";
            string tag = floating ? "anchor" : "inline";
            return "<w:p><w:r><w:drawing><wp:" + tag + " distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">" + properties +
                "<wp:extent cx=\"5080000\" cy=\"1270000\"/>" + (floating ? "<wp:wrapNone/>" : "") +
                "<wp:docPr id=\"1\" name=\"Math box\"/><a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent>" + paragraph +
                "</w:txbxContent></wps:txbx><wps:bodyPr lIns=\"0\" tIns=\"0\" rIns=\"0\" bIns=\"0\"/></wps:wsp></a:graphicData></a:graphic></wp:" + tag + "></w:drawing></w:r></w:p>";
        }
        var parts = new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>",
            ["_rels/.rels"] = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"main\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>",
            ["word/_rels/document.xml.rels"] = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"settings\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings\" Target=\"settings.xml\"/>" +
                (staticStory ? $"<Relationship Id=\"{story}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/{story}\" Target=\"{story}1.xml\"/>" : "") +
                (note ? $"<Relationship Id=\"{story}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/{story}s\" Target=\"{story}s.xml\"/>" : "") + "</Relationships>",
            ["word/settings.xml"] = $"<w:settings {namespaces}><m:mathPr><m:intLim m:val=\"subSup\"/><m:naryLim m:val=\"undOvr\"/>" + (global is null ? "" : $"<m:defJc m:val=\"{global}\"/>") + "</m:mathPr></w:settings>",
            ["word/document.xml"] = $"<w:document {namespaces}><w:body>" + (container switch
            {
                "cell" => Cell(paragraph),
                "footer" or "header" or "header-cell" => "<w:p><w:r><w:t>Body</w:t></w:r></w:p>",
                "footnote" or "endnote" => $"<w:p><w:r><w:t>Body</w:t><w:{story}Reference w:id=\"1\"/></w:r></w:p>",
                "inlinebox" => Box(false),
                "floatingbox" => Box(true),
                _ => paragraph
            }) + "<w:sectPr>" + (staticStory ? $"<w:{story}Reference w:type=\"default\" r:id=\"{story}\"/>" : "") +
                "<w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:right=\"1440\"/></w:sectPr></w:body></w:document>"
        };
        if (staticStory)
        {
            string tag = story == "header" ? "hdr" : "ftr";
            parts[$"word/{story}1.xml"] = $"<w:{tag} {namespaces}>" + (container == "header-cell" ? Cell(paragraph) : paragraph) + $"</w:{tag}>";
        }
        if (note) parts[$"word/{story}s.xml"] = $"<w:{story}s {namespaces}><w:{story} w:id=\"1\">{paragraph}</w:{story}></w:{story}s>";
        using var stream = TestFixtures.CreateZipPackage(parts);
        using var package = OoxPackage.Open(stream, default);
        return new DocxReader().Read(package, null, default, OoxPdfDocxMarkupMode.Final);
    }
}
