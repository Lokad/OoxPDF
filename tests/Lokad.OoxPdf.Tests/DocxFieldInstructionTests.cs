using System.Xml.Linq;
using System.Text;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxFieldInstructionTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    public static void DocxComputedFieldCachesUseResultOwnershipAcrossRunBoundaries()
    {
        foreach (var mode in Modes)
        foreach (string opcode in new[] { "PAGE", "NUMPAGES" })
        foreach (string cache in new[] { "stale", "7stale", "7", "Ⅶ", "1,234", "" })
        foreach (string shape in new[] { "separate", "split-cache", "packed" })
        {
            var nodes = Field([opcode], shape == "split-cache" ? null : cache).ToList();
            if (shape == "split-cache") nodes.InsertRange(nodes.Count - 1, cache.Select(c => Run(c.ToString())));
            nodes.Add(Run(" 42"));
            XElement[] source = shape == "packed"
                ? [new XElement(W + "r", nodes.SelectMany(run => run.Elements()).Select(child => new XElement(child)))]
                : nodes.ToArray();
            var paragraph = Read([.. source, new XElement(M + "oMath", new XElement(M + "r", new XElement(M + "t", "x")))], mode)
                .Paragraphs.Single();
            TestAssert.Equal("{" + opcode + "} 42", string.Concat(paragraph.Runs.Select(run => run.Text)));
            var field = paragraph.FieldReferences.Single();
            TestAssert.True(field.UsesPlaceholder && !field.RendersCachedResult && field.HasSeparate &&
                field.HasCachedResult == (cache.Length != 0), "Computed values replace their own cache, not following authored numbers.");
            TestAssert.True(paragraph.InlineContent is not null, "Stale spelling cannot invalidate an ordered paragraph's source.");
            TestAssert.Equal(1, DocxMathParagraphPolicy.GetPreparationEquations(paragraph, hasPageInputs: true).Count);
        }
    }

    public static void DocxComputedFieldCacheOwnersKeepStoredLocksAndNestedResults()
    {
        foreach (var mode in Modes)
        {
            foreach (string opcode in new[] { "PAGE", "NUMPAGES", "REF Target" })
            foreach (bool packed in new[] { false, true })
            {
                var nodes = Field([opcode], "stored7").ToList();
                nodes[0].Element(W + "fldChar")!.SetAttributeValue(W + "fldLock", "true");
                nodes.Add(Run(" 42"));
                XElement[] source = packed
                    ? [new XElement(W + "r", nodes.SelectMany(run => run.Elements()).Select(child => new XElement(child)))]
                    : nodes.ToArray();
                var paragraph = Read(source, mode).Paragraphs.Single();
                TestAssert.Equal("stored7 42", string.Concat(paragraph.Runs.Select(run => run.Text)));
                TestAssert.True(paragraph.FieldReferences.Single() is { HasCachedResult: true, RendersCachedResult: true, UsesPlaceholder: false },
                    "A lock retains its stored result and does not inherit another counter's replacement policy.");
            }
            var outer = Field(["REF Outer"], null, end: false).ToList();
            outer.Add(Run("Outer "));
            outer.AddRange(Field(["PAGE"], "stale7"));
            outer.Add(Run(" end"));
            outer.Add(new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "end"))));
            outer.Add(Run(" 42"));
            var nested = Read(outer.ToArray(), mode).Paragraphs.Single();
            TestAssert.Equal("Outer {PAGE} end 42", string.Concat(nested.Runs.Select(run => run.Text)));
            TestAssert.True(nested.FieldReferences.Single(f => f.Kind == DocxFieldKind.Page) is { UsesPlaceholder: true, RendersCachedResult: false },
                "A nested computed result discards only its own cached text.");
            TestAssert.True(nested.FieldReferences.Single(f => f.Kind == DocxFieldKind.Other) is { HasCachedResult: true, RendersCachedResult: true },
                "The enclosing stored field keeps its visible source result.");
        }
    }

    public static void DocxFieldInstructionSplittingPreservesCompleteOpcodesAndCachedValues()
    {
        foreach (var mode in Modes)
        foreach (string opcode in new[] { "PAGEREF Target", "NUMPAGESX" })
        foreach (int split in Enumerable.Range(1, opcode.Length - 1))
        {
            var paragraph = Read([.. Field([" " + opcode[..split], opcode[split..] + " "], "2")], mode).Paragraphs.Single();
            TestAssert.Equal("2", string.Concat(paragraph.Runs.Select(run => run.Text)));
            TestAssert.True(paragraph.Runs.All(run => run.FieldKind is null), "A partial opcode must not publish a dynamic counter.");
            var field = paragraph.FieldReferences.Single();
            TestAssert.Equal(DocxFieldKind.Other, field.Kind);
            TestAssert.Equal(" " + opcode + " ", field.Instruction);
            TestAssert.Equal(2, field.InstructionRunCount);
            TestAssert.True(field.HasCachedResult && field.RendersCachedResult && !field.UsesPlaceholder,
                "Cached results belong to the complete opcode, independently of instruction-run boundaries.");
        }
    }

    public static void DocxFieldInstructionBoundariesConfirmRealCountersAndKeepFollowingNumbers()
    {
        foreach (var mode in Modes)
        foreach (string opcode in new[] { "PAGE", "NUMPAGES" })
        foreach (bool separate in new[] { true, false })
        {
            var source = Field(opcode.Select(character => character.ToString()).ToArray(), separate ? "9" : null, separate);
            var paragraph = Read([.. source, Run(" 42")], mode).Paragraphs.Single();
            TestAssert.Equal("{" + opcode + "} 42", string.Concat(paragraph.Runs.Select(run => run.Text)));
            TestAssert.Equal(opcode == "PAGE" ? DocxFieldKind.Page : DocxFieldKind.NumPages, paragraph.Runs[0].FieldKind!.Value);
            var field = paragraph.FieldReferences.Single();
            TestAssert.True(field.UsesPlaceholder && !field.RendersCachedResult, "A completed real counter retains dynamic substitution.");
            TestAssert.Equal(opcode.Length, field.InstructionRunCount);
            TestAssert.Equal(separate, field.HasSeparate);
            TestAssert.Equal(separate, field.HasCachedResult);
        }
    }

    public static void DocxFieldInstructionSplittingKeepsOrderedMathAndLinks()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool split in new[] { false, true })
        {
            var code = split ? new[] { " PAGE", "REF Target " } : new[] { " PAGEREF Target " };
            var link = new XElement(W + "hyperlink", new XAttribute(W + "anchor", "Target"), Field(code, "2"));
            var math = new XElement(M + "oMath", new XElement(M + "r", new XElement(M + "t", "x")));
            var document = Read([Run("Reference "), link, Run(" end "), math], mode, target: true);
            var paragraph = document.Paragraphs[0];
            TestAssert.Equal("Reference 2 end ", string.Concat(paragraph.Runs.Select(run => run.Text)));
            TestAssert.Equal(1, paragraph.Equations.Count);
            TestAssert.True(paragraph.InlineContent is not null && paragraph.Hyperlinks.Count == 1,
                "Cached field results retain ordered inline source and hyperlink ownership.");
            var geometry = mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
            var pages = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode, geometry)
                .RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(2, pages.Length);
            string text = string.Concat(pages.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true))
                .Select(show => show.Text));
            TestAssert.Contains("Reference2endx", string.Concat(text.Normalize(NormalizationForm.FormKC)
                .Where(character => !char.IsWhiteSpace(character))));
            TestAssert.True(pages[0].Annotations.Any(annotation => annotation.Destination?.PageIndex == 1),
                "The cached cross-reference retains its internal link.");
        }
    }

    public static void DocxFieldInstructionUnclosedCountersKeepOrderedSourceFallback()
    {
        var math = new XElement(M + "oMath", new XElement(M + "r", new XElement(M + "t", "x")));
        var paragraph = Read([math, .. Field(["PAGE"], null, separate: false, end: false)],
            OoxPdfDocxMarkupMode.Final).Paragraphs.Single();
        TestAssert.Equal("{PAGE}", paragraph.Runs.Single().Text);
        TestAssert.True(paragraph.InlineContent?.OfType<DocxInlineText>().Any(item => item.Run.FieldKind == DocxFieldKind.Page) == true,
            "The existing unclosed-counter fallback must publish its late run into ordered source.");
        TestAssert.True(paragraph.FieldReferences.Single().UsesPlaceholder, "The source field retains its dynamic identity.");
    }

    private static IEnumerable<XElement> Field(string[] instruction, string? cached, bool separate = true, bool end = true)
    {
        yield return new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "begin")));
        foreach (string part in instruction)
            yield return new XElement(W + "r", new XElement(W + "instrText", new XAttribute(XNamespace.Xml + "space", "preserve"), part));
        if (separate)
            yield return new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "separate")));
        if (cached is not null) yield return Run(cached);
        if (end) yield return new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "end")));
    }

    private static DocxDocument Read(XElement[] content, OoxPdfDocxMarkupMode mode, bool target = false)
    {
        var body = new XElement(W + "body", new XElement(W + "p", content));
        if (target) body.Add(new XElement(W + "p", new XElement(W + "pPr", new XElement(W + "pageBreakBefore")),
            new XElement(W + "bookmarkStart", new XAttribute(W + "id", "1"), new XAttribute(W + "name", "Target")),
            Run("Target"), new XElement(W + "bookmarkEnd", new XAttribute(W + "id", "1"))));
        body.Add(new XElement(W + "sectPr", new XElement(W + "pgSz", new XAttribute(W + "w", 12240), new XAttribute(W + "h", 15840))));
        string path = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>",
            ["_rels/.rels"] = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"main\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>",
            ["word/document.xml"] = new XElement(W + "document", body).ToString()
        });
        return DocxTests.ReadDocx(path, mode);
    }

    private static XElement Run(string text) => new(W + "r", new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), text));
    private static readonly OoxPdfDocxMarkupMode[] Modes = Enum.GetValues<OoxPdfDocxMarkupMode>();
}
