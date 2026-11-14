using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxNumberingTests
{
    public static void DocxStyleNumberingRetainsSourceMarkersCountersAndMath()
    {
        foreach (string story in new[] { "body", "cell" })
        foreach (string kind in new[] { "none", "page", "column" })
        foreach (OoxPdfDocxMarkupMode mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            var document = DocxInlineCompositionTests.ReadMutation("StyleNumberingReferences/" + story, mode, false, false,
                (xml, w, _) =>
                {
                    if (kind == "none") return;
                    var paragraph = xml.Descendants(w + "p").First(p => p.Element(w + "pPr")?.Element(w + "pStyle")?.Attribute(w + "val")?.Value == "PublicNumberedSource");
                    paragraph.Element(w + "pPr")!.AddAfterSelf(new XElement(w + "r", new XElement(w + "br", new XAttribute(w + "type", kind))));
                });
            var paragraphs = DocxBlockTraversal.EnumerateSourceParagraphs(document).Distinct<DocxParagraph>(ReferenceEqualityComparer.Instance).ToArray();
            var numbered = paragraphs.Where(p => p.ListLabel is not null).ToArray();
            TestAssert.True(numbered.Select(p => p.ListLabel!.Text).SequenceEqual(new[] { "1.", "2.", "3." }),
                "Inherited numbering retains one marker/counter event per source item: " + story + "/" + kind + "/" + mode);
            TestAssert.Equal(DocxMarkupContext.FromMode(mode).AppliesInlineRevisionStyle ? "1." : null,
                numbered[1].ListLabel!.NumberingChange?.PreviousText);
            TestAssert.Equal(4, paragraphs.Sum(p => DocxMathParagraphPolicy.GetPreparationEquations(p).Count));
            foreach (var paragraph in paragraphs.Where(p => p.Equations.Count != 0))
                TestAssert.True(DocxInlineContent.HasSupportedSourceOwner(paragraph),
                    "A style-resolved list reference owns its actual revision metadata beside supported equations.");
        }
    }

    public static void DocxStyleNumberingPaintMatchesDirectSourceFormatting()
    {
        foreach (string story in new[] { "body", "cell" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            PdfPage[] Paint(string fixture, out List<OoxPdfDiagnostic> diagnostics)
            {
                diagnostics = [];
                var document = DocxInlineCompositionTests.ReadMutation(fixture, mode, false, false);
                return renderer.RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            }
            var direct = Paint("CompositionProbes/numbering-mark-math-" + story, out _);
            var styled = Paint("StyleNumberingReferences/" + story, out var diagnostics);
            TestAssert.Equal(direct.Length, styled.Length);
            for (int page = 0; page < styled.Length; page++)
            {
                TestAssert.Equal(direct[page].Content, styled[page].Content);
                TestAssert.Equal(direct[page].Images.Count, styled[page].Images.Count);
                PdfContentValidator.ValidatePage(styled[page], page, default);
            }
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Style inheritance cannot discard a fraction while conversion succeeds.");
            var text = string.Concat(styled.Select(DocxPageFootnoteNumberingTests.DecodeText));
            foreach (string expected in new[] { "1.", "2.", "3.", "Summary:", "planning ratio.", "The trend marker" }) TestAssert.Contains(expected, text);
            using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(
                DocxInlineCompositionTests.ReadMutation("StyleNumberingReferences/" + story, mode, false, false), null, cancelled.Token)
                .ToBlockingEnumerable().ToArray());
        }
    }

    public static void DocxStyleNumberingKeepsOverridesChainsAndUnresolvedFallbacks()
    {
        foreach (string variant in new[] { "chain", "cycle", "alias", "disabled", "unknown" })
        {
            var document = DocxInlineCompositionTests.ReadMutation("StyleNumberingReferences/body", OoxPdfDocxMarkupMode.Final, false, false,
                (xml, w, _) =>
                {
                    var style = xml.Root!.Elements(w + "style").Single(s => s.Attribute(w + "styleId")?.Value == "PublicNumberedSource");
                    if (variant is "chain" or "cycle")
                    {
                        style.SetAttributeValue(w + "styleId", "PublicNumberingBase");
                        xml.Root.Add(new XElement(w + "style", new XAttribute(w + "type", "paragraph"), new XAttribute(w + "styleId", "PublicNumberedSource"),
                            new XElement(w + "basedOn", new XAttribute(w + "val", "PublicNumberingBase"))));
                        if (variant == "cycle") style.Add(new XElement(w + "basedOn", new XAttribute(w + "val", "PublicNumberedSource")));
                    }
                    if (variant == "alias") style.Descendants(w + "numId").Single().SetAttributeValue(w + "val", "+0001");
                    if (variant is "disabled" or "unknown") style.Descendants(w + "numId").Single().SetAttributeValue(w + "val", variant == "disabled" ? "0" : "999");
                }, partOverride: "word/styles.xml");
            var first = document.Paragraphs.First(p => p.Equations.Count != 0);
            TestAssert.Equal(variant is "disabled" or "unknown" ? null : "1.", first.ListLabel?.Text);
            TestAssert.Equal(variant is "disabled" or "unknown" ? "1." : "2.", document.Paragraphs.First(p =>
                (p.ListLabel?.Text is "1." or "2.") && !ReferenceEquals(p, first)).ListLabel?.Text);
        }
        foreach (string id in new[] { "0", "999" })
        {
            var document = DocxInlineCompositionTests.ReadMutation("StyleNumberingReferences/body", OoxPdfDocxMarkupMode.Final, false, false,
                (xml, w, _) => xml.Descendants(w + "p").First(p => p.Element(w + "pPr")?.Element(w + "numPr") is not null)
                    .Element(w + "pPr")!.Element(w + "numPr")!.Add(new XElement(w + "numId", new XAttribute(w + "val", id))));
            var first = document.Paragraphs.First(p => p.Equations.Count != 0);
            TestAssert.True(first.ListLabel is null && DocxInlineContent.HasSupportedSourceOwner(first),
                "A direct disabled/unknown reference suppresses its marker without erasing the actual mark owner.");
        }
    }

    public static void DocxStyleNumberingResolvesAssociatedLevelsAndDefaultOwners()
    {
        foreach (bool defaultStyle in new[] { false, true })
        foreach (bool directLevel in new[] { false, true })
        foreach (bool ambiguous in new[] { false, true })
        foreach (bool concrete in new[] { false, true })
        {
            var document = ReadAssociatedStyleNumbering(defaultStyle, directLevel, ambiguous, concrete);
            var numbered = document.Paragraphs.Where(p => p.ListLabel is not null).ToArray();
            if (ambiguous && !directLevel)
            {
                TestAssert.Equal(1, numbered.Length);
                TestAssert.Equal("[7]", numbered[0].ListLabel!.Text);
            }
            else
            {
                TestAssert.True(numbered.Select(p => p.ListLabel!.Text).SequenceEqual(directLevel ? new[] { "1.", "[7]" } : new[] { "[7]", "[8]" }),
                    "An inherited reference uses the nearest associated style level; a paragraph's explicit level overrides it.");
                TestAssert.Equal(1, document.BodyElements.OfType<DocxPageBreakElement>().Sum(p => p.BreakCount));
            }
        }
    }

    private static DocxDocument ReadAssociatedStyleNumbering(bool defaultStyle, bool directLevel, bool ambiguous, bool concrete)
    {
        string appliedStyle = defaultStyle ? "" : "<w:pStyle w:val=\"Derived\"/>";
        string styles = "<w:style w:type=\"paragraph\" w:styleId=\"Numbered\"><w:pPr><w:numPr><w:numId w:val=\"1\"/></w:numPr></w:pPr></w:style>" +
            "<w:style w:type=\"paragraph\" w:styleId=\"Derived\"" + (defaultStyle ? " w:default=\"1\"" : "") + "><w:basedOn w:val=\"Numbered\"/></w:style>";
        string body = "<w:p><w:pPr>" + appliedStyle + (directLevel ? "<w:numPr><w:ilvl w:val=\"0\"/></w:numPr>" : "") +
            "</w:pPr><w:r><w:br w:type=\"page\"/></w:r></w:p>" +
            "<w:p><w:pPr><w:numPr><w:numId w:val=\"1\"/><w:ilvl w:val=\"1\"/></w:numPr></w:pPr><w:r><w:t>NEXT</w:t></w:r></w:p>";
        string secondLevel = "<w:lvl w:ilvl=\"1\"><w:start w:val=\"7\"/><w:pStyle w:val=\"Numbered\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"[%2]\"/></w:lvl>";
        using var input = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>""",
            ["_rels/.rels"] = """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="numbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>""",
            ["word/document.xml"] = "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" + body + "</w:body></w:document>",
            ["word/styles.xml"] = "<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" + styles + "</w:styles>",
            ["word/numbering.xml"] = "<w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:abstractNum w:abstractNumId=\"7\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.\"/>" +
                (ambiguous ? "<w:pStyle w:val=\"Numbered\"/>" : "") + "</w:lvl>" + (concrete ? "" : secondLevel) +
                "</w:abstractNum><w:num w:numId=\"1\"><w:abstractNumId w:val=\"7\"/>" +
                (concrete ? "<w:lvlOverride w:ilvl=\"1\">" + secondLevel + "</w:lvlOverride>" : "") + "</w:num></w:numbering>"
        });
        return new DocxReader().Read(OoxPackage.Open(input, default), null, default, OoxPdfDocxMarkupMode.Final);
    }
}
