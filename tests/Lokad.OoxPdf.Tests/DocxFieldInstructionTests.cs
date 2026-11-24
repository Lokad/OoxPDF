using System.Xml.Linq;
using System.Text;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

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

    public static void DocxComputedOuterFieldsSuppressNestedCountersInTheirReplacedResult()
    {
        var resolver = new PortableRenderingTests.PinnedResolver();
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool simple in new[] { false, true })
        foreach (bool acrossBreak in new[] { false, true })
        {
            var source = new List<XElement> { Run("Before ") };
            source.AddRange(Field([" PAGE "], null, end: false));
            source.Add(Run("91"));
            if (acrossBreak)
                source.Add(new XElement(W + "r", new XElement(W + "br", new XAttribute(W + "type", "page"))));
            if (simple)
                source.Add(new XElement(W + "fldSimple", new XAttribute(W + "instr", " NUMPAGES "), Run("73")));
            else
                source.AddRange(Field([" NUMPAGES "], "73"));
            source.Add(new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "end"))));
            source.Add(Run(" after 42"));

            DocxDocument document = Read(source.ToArray(), mode);
            TestAssert.Equal("Before {PAGE} after 42", string.Concat(document.Paragraphs.SelectMany(p => p.Runs).Select(run => run.Text)));
            var fields = document.Paragraphs.SelectMany(p => p.FieldReferences).ToArray();
            TestAssert.True(fields.Single(field => field.Kind == DocxFieldKind.Page) is { UsesPlaceholder: true, HasCachedResult: true },
                "The computed outer field owns one replacement.");
            TestAssert.True(fields.Single(field => field.Kind == DocxFieldKind.NumPages) is { UsesPlaceholder: false, HasCachedResult: true, RendersCachedResult: false },
                "The nested field retains its source and neither a generated nor a stored value paints.");

            var geometry = mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
            var pages = new DocxRenderer(resolver, mode, geometry).RenderPagesAsync(document, null, default)
                .ToBlockingEnumerable()
                .Select(page => string.Concat(DocxPageFootnoteNumberingTests.DecodeText(page).Where(character => !char.IsWhiteSpace(character))))
                .ToArray();
            TestAssert.Equal(acrossBreak ? 2 : 1, pages.Length);
            TestAssert.Equal(acrossBreak ? "Before1" : "Before1after42", pages[0]);
            if (acrossBreak)
                TestAssert.Equal("after42", pages[1]);
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

    public static void DocxUnownedInstructionTextRetainsAuthoredTextAndLinks()
    {
        var resolver = new PortableRenderingTests.PinnedResolver();
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var geometry = mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;

            var literal = Render(W + "t");
            var fieldText = Render(W + "instrText");
            TestAssert.Equal("BeforePAGE4729afterREFAnchor!", Compact(fieldText.Text));
            TestAssert.Equal(literal.Text, fieldText.Text);
            TestAssert.Equal(literal.Link, fieldText.Link);

            (string Text, PdfLinkAnnotation Link) Render(XName textElement)
            {
                var diagnostics = new List<OoxPdfDiagnostic>();
                var link = new XElement(W + "hyperlink", new XAttribute(W + "anchor", "Target"),
                    new XElement(W + "r", Text(W + "t", "Before"), Text(textElement, " PAGE ")));
                var sameRun = new XElement(W + "r", Text(W + "t", " after "),
                    Text(textElement, "REF Anchor"), Text(W + "t", "!"));
                var document = Read([link, Run("4729"), sameRun],
                    mode, target: true, diagnosticSink: diagnostics.Add);
                TestAssert.Equal(0, diagnostics.Count);
                TestAssert.Equal(0, document.Paragraphs[0].FieldReferences.Count);
                var pages = new DocxRenderer(resolver, mode, geometry).RenderPagesAsync(document, null, default)
                    .ToBlockingEnumerable().ToArray();
                TestAssert.Equal(2, pages.Length);
                var annotation = pages[0].Annotations.Single();
                TestAssert.Equal(1, annotation.Destination!.Value.PageIndex);
                return (DocxPageFootnoteNumberingTests.DecodeText(pages[0]), annotation);
            }
        }

        static XElement Text(XName name, string value) => new(name, new XAttribute(XNamespace.Xml + "space", "preserve"), value);
        static string Compact(string text) => string.Concat(text.Where(character => !char.IsWhiteSpace(character)));
    }

    public static void DocxLiteralDeletedInstructionTextFollowsDeletedTextVisibility()
    {
        foreach (var mode in Modes)
        {
            var literal = Render("delText");
            var instruction = Render("delInstrText");
            TestAssert.Equal(mode is OoxPdfDocxMarkupMode.Original or OoxPdfDocxMarkupMode.AllMarkup
                ? "BeforeDeletedwords4729after" : "Before4729after",
                string.Concat(instruction.Text.Where(character => !char.IsWhiteSpace(character))));
            TestAssert.Equal(literal.Text, instruction.Text);
            TestAssert.Equal(literal.Content, instruction.Content);

            (string Text, string Content) Render(string name)
            {
                var diagnostics = new List<OoxPdfDiagnostic>();
                var source = Read([Run("Before "), Revision("del", 1, RunText(name, "Deleted words ")), Run("4729 after")],
                    mode, diagnosticSink: diagnostics.Add);
                TestAssert.Equal(0, source.Paragraphs[0].FieldReferences.Count);
                TestAssert.Equal(0, diagnostics.Count);
                var page = RenderRevisionPages(source, mode)[0];
                return (DocxPageFootnoteNumberingTests.DecodeText(page), page.Content);
            }
        }
    }

    public static void DocxFieldRevisionCodesSelectViewAndKeepLockedResults()
    {
        foreach (var mode in Modes)
        {
            bool showsDeleted = mode is OoxPdfDocxMarkupMode.Original or OoxPdfDocxMarkupMode.AllMarkup;
            foreach (bool locked in new[] { false, true })
            {
                var deleted = Revision("del", 1, Marker("begin", locked), RunText("delInstrText", " PAGE "),
                    Marker("separate"), RunText("delText", "7811 "), Marker("end"));
                AssertField([deleted], showsDeleted ? DocxFieldKind.Page : null,
                    showsDeleted ? locked ? "7811" : "1" : "", computes: showsDeleted && !locked);
            }

            bool original = mode == OoxPdfDocxMarkupMode.Original;
            var former = Revision("del", 1, RunText("delInstrText", " PA"), RunText("delInstrText", "GE "));
            var current = Revision("ins", 2, RunText("instrText", " NUM"), RunText("instrText", "PAGES "));
            AssertField([Marker("begin"), former, current, Marker("separate"), Run("7811 "), Marker("end")],
                original ? DocxFieldKind.Page : DocxFieldKind.NumPages, original ? "1" : "2", computes: true);
            AssertField([Marker("begin"), former, current, Marker("end")],
                original ? DocxFieldKind.Page : DocxFieldKind.NumPages, original ? "1" : "2", computes: true);

            void AssertField(XElement[] contents, DocxFieldKind? expected, string visibleResult, bool computes)
            {
                var diagnostics = new List<OoxPdfDiagnostic>();
                var source = Read([Run("Before "), .. contents, Run(" 4729 after")], mode,
                    target: true, diagnosticSink: diagnostics.Add);
                TestAssert.Equal(0, diagnostics.Count);
                var fields = source.Paragraphs[0].FieldReferences;
                if (expected.HasValue)
                {
                    var field = fields.Single();
                    TestAssert.Equal(expected.Value, field.Kind);
                    TestAssert.Equal(computes, field.UsesPlaceholder);
                }
                else
                {
                    TestAssert.Equal(0, fields.Count);
                }
                PdfPage[] pages = RenderRevisionPages(source, mode);
                TestAssert.Equal(2, pages.Length);
                string actual = string.Concat(DocxPageFootnoteNumberingTests.DecodeText(pages[0]).Where(c => !char.IsWhiteSpace(c)));
                TestAssert.Equal("Before" + visibleResult + "4729after", actual);
            }
        }

        static XElement Marker(string kind, bool locked = false) => new(W + "r",
            new XElement(W + "fldChar", new XAttribute(W + "fldCharType", kind),
                locked ? new XAttribute(W + "fldLock", "1") : null));
    }

    private static PdfPage[] RenderRevisionPages(DocxDocument document, OoxPdfDocxMarkupMode mode)
    {
        var geometry = mode == OoxPdfDocxMarkupMode.AllMarkup
            ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
            : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout;
        return new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode, geometry)
            .RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
    }

    private static XElement Revision(string kind, int id, params XElement[] runs) =>
        new(W + kind, new XAttribute(W + "id", id), new XAttribute(W + "author", "Reviewer"), runs);

    private static XElement RunText(string kind, string value) =>
        new(W + "r", new XElement(W + kind, new XAttribute(XNamespace.Xml + "space", "preserve"), value));

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

    public static void DocxFieldInstructionRequiresCompleteNameAndKeepsSwitches()
    {
        var resolver = new PortableRenderingTests.PinnedResolver();
        foreach (var mode in Modes)
        {
            DocxDocument document = Read([
                new XElement(W + "fldSimple", new XAttribute(W + "instr", " PAGE2 "), Run("Simple cache")),
                Run(" 42 "),
                .. Field([" NUMPAGES", "_ARCHIVE "], "Complex cache"),
                Run(" 73 "),
                new XElement(W + "fldSimple", new XAttribute(W + "instr", @" PAGE\* MERGEFORMAT "), Run("29")),
                Run(" / "),
                .. Field([" NUMPAGES", @" \* MERGEFORMAT "], "87")
            ], mode);
            var fields = document.Paragraphs.Single().FieldReferences;
            var stored = fields.Where(field => field.Kind == DocxFieldKind.Other).ToArray();
            TestAssert.Equal(2, stored.Length);
            TestAssert.True(stored.All(field => field.HasCachedResult && field.RendersCachedResult && !field.UsesPlaceholder),
                "Unknown complete names retain their stored result.");
            TestAssert.True(fields.Where(field => field.UsesPlaceholder).Select(field => field.Kind)
                .SequenceEqual([DocxFieldKind.Page, DocxFieldKind.NumPages]),
                "A formatting switch still delimits a recognized counter.");

            var geometry = mode == OoxPdfDocxMarkupMode.AllMarkup
                ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout;
            var page = new DocxRenderer(resolver, mode, geometry).RenderPagesAsync(document, null, default)
                .ToBlockingEnumerable().Single();
            string actual = string.Concat(DocxPageFootnoteNumberingTests.DecodeText(page)
                .Where(character => !char.IsWhiteSpace(character)));
            TestAssert.Equal("Simplecache42Complexcache731/1", actual);
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
        var unclosed = Read([math, .. Field(["PAGE"], null, separate: false, end: false)],
            OoxPdfDocxMarkupMode.Final).Paragraphs.Single();
        TestAssert.True(unclosed.Runs.All(run => run.Text.Length == 0), "Unclosed code must not synthesize a counter.");
        TestAssert.Equal(1, unclosed.Equations.Count);
        TestAssert.Equal(0, unclosed.FieldReferences.Count);

        var closed = Read([new XElement(math), .. Field(["PAGE"], null, separate: false)],
            OoxPdfDocxMarkupMode.Final).Paragraphs.Single();
        TestAssert.Equal("{PAGE}", closed.Runs.Single().Text);
        TestAssert.True(closed.InlineContent?.OfType<DocxInlineText>().Any(item => item.Run.FieldKind == DocxFieldKind.Page) == true,
            "A matching end must publish the late counter run into ordered source.");
        TestAssert.True(closed.FieldReferences.Single().UsesPlaceholder, "The closed field retains its dynamic identity.");
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

    private static DocxDocument Read(XElement[] content, OoxPdfDocxMarkupMode mode, bool target = false,
        Action<OoxPdfDiagnostic>? diagnosticSink = null)
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
        using FileStream input = File.OpenRead(path);
        return new DocxReader().Read(OoxPackage.Open(input, CancellationToken.None), diagnosticSink, CancellationToken.None, mode);
    }

    private static XElement Run(string text) => new(W + "r", new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), text));
    private static readonly OoxPdfDocxMarkupMode[] Modes = Enum.GetValues<OoxPdfDocxMarkupMode>();
}
