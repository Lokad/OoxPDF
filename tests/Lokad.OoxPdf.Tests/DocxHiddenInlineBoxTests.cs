using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxHiddenInlineBoxTests
{
    private static readonly (string Fixture, string Part)[] Cases =
    [
        ("inline-box-owned-text-first", "word/document.xml"),
        ("inline-box-body-equation-inline-box-valid", "word/document.xml"),
        ("inline-box-math-cell", "word/document.xml"),
        ("inline-box-math-header", "word/inline-box-header.xml"),
        ("inline-box-math-footer", "word/inline-box-footer.xml")
    ];

    public static void DocxHiddenInlineBoxesPreserveSiblingPaintAcrossContainers()
    {
        foreach (var source in Cases)
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = Renderer(mode);
            var removed = Read(source, mode, (xml, _, w) => BoxRun(xml, w).Remove());
            var hidden = Read(source, mode, (xml, _, w) => SetHidden(BoxRun(xml, w), w, true));
            TestAssert.True(Paragraphs(hidden).All(p => p.InlineTextBoxes.Count == 0),
                "A hidden drawing run must not publish a text box: " + source.Fixture);
            var diagnostics = new List<OoxPdfDiagnostic>();
            AssertPaint(renderer.RenderPagesAsync(removed, null, default).ToBlockingEnumerable().ToArray(),
                renderer.RenderPagesAsync(hidden, diagnostics.Add, default).ToBlockingEnumerable().ToArray(), source.Fixture + "/" + mode);
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Intentionally hidden box mathematics is not a visible equation omission.");
        }
    }

    public static void DocxInlineBoxesUseInheritedVisibilityAndVisibleOverrides()
    {
        var source = Cases[2];
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        foreach (string owner in new[] { "direct", "character", "paragraph", "default-paragraph", "defaults", "table" })
        foreach (bool visible in new[] { false, true })
        {
            var document = Read(source, mode, (xml, styles, w) =>
            {
                var run = BoxRun(xml, w);
                if (owner == "direct") SetHidden(run, w, true);
                else if (owner == "defaults")
                {
                    var defaults = styles.Root!.Element(w + "docDefaults");
                    if (defaults is null) { defaults = new(w + "docDefaults"); styles.Root.AddFirst(defaults); }
                    defaults.Element(w + "rPrDefault")?.Remove();
                    defaults.AddFirst(new XElement(w + "rPrDefault", new XElement(w + "rPr", new XElement(w + "vanish"))));
                }
                else
                {
                    string type = owner == "character" ? "character" : owner == "table" ? "table" : "paragraph";
                    styles.Root!.Add(new XElement(w + "style", new XAttribute(w + "type", type), new XAttribute(w + "styleId", "HiddenBoxBase"),
                        new XElement(w + "rPr", new XElement(w + "vanish"))));
                    var derived = new XElement(w + "style", new XAttribute(w + "type", type), new XAttribute(w + "styleId", "HiddenBox"),
                        new XElement(w + "basedOn", new XAttribute(w + "val", "HiddenBoxBase")));
                    styles.Root.Add(derived);
                    if (owner == "character") Properties(run, w).Add(new XElement(w + "rStyle", new XAttribute(w + "val", "HiddenBox")));
                    else if (owner == "default-paragraph")
                    {
                        foreach (var style in styles.Root.Elements(w + "style").Where(s => (string?)s.Attribute(w + "type") == "paragraph"))
                            style.Attribute(w + "default")?.Remove();
                        derived.SetAttributeValue(w + "default", "1");
                    }
                    else
                    {
                        var parent = run.Ancestors(w + (owner == "table" ? "tbl" : "p")).First();
                        string properties = owner == "table" ? "tblPr" : "pPr", key = owner == "table" ? "tblStyle" : "pStyle";
                        var props = parent.Element(w + properties);
                        if (props is null) { props = new(w + properties); parent.AddFirst(props); }
                        props.Element(w + key)?.Remove(); props.AddFirst(new XElement(w + key, new XAttribute(w + "val", "HiddenBox")));
                    }
                }
                if (visible) SetHidden(run, w, false);
            });
            TestAssert.Equal(visible ? 1 : 0, Paragraphs(document).Sum(p => p.InlineTextBoxes.Count));
        }
    }

    public static void DocxHiddenInlineBoxesRetainRevisionFieldAndOmissionOwnership()
    {
        foreach (string variant in new[] { "insertion", "cached-field", "unsupported-sibling" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = Cases[1];
            DocxDocument Make(bool removed) => Read(source, mode, (xml, _, w) =>
            {
                var run = BoxRun(xml, w);
                if (variant == "insertion")
                {
                    var insertion = new XElement(w + "ins", new XAttribute(w + "id", "901"), new XAttribute(w + "author", "Public reviewer"));
                    run.ReplaceWith(insertion); insertion.Add(run);
                }
                if (variant == "cached-field")
                {
                    var field = new XElement(w + "fldSimple", new XAttribute(w + "instr", "QUOTE Cached"));
                    run.ReplaceWith(field); field.Add(run);
                }
                if (variant == "unsupported-sibling")
                {
                    XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
                    run.Ancestors(w + "p").First().Add(new XElement(m + "oMath", new XElement(m + "future")));
                }
                if (removed)
                {
                    // Removing the insertion scope avoids substituting the
                    // separate empty-revision ownership fallback for this control.
                    if (variant == "insertion") run.Parent!.Remove(); else run.Remove();
                }
                else SetHidden(run, w, true);
            });
            var renderer = Renderer(mode); var diagnostics = new List<OoxPdfDiagnostic>();
            var hidden = Make(false);
            AssertPaint(renderer.RenderPagesAsync(Make(true), null, default).ToBlockingEnumerable().ToArray(),
                renderer.RenderPagesAsync(hidden, diagnostics.Add, default).ToBlockingEnumerable().ToArray(), variant + "/" + mode);
            TestAssert.Equal(variant == "unsupported-sibling", diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_EQUATION"));
        }
    }

    public static void DocxHiddenBoxEquationSourcesSurviveConsumedParagraphMarks()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = Read(Cases[2], mode, (xml, _, w) =>
            {
                var run = BoxRun(xml, w); SetHidden(run, w, true);
                var paragraph = run.Ancestors(w + "p").First();
                var second = new XElement(paragraph); paragraph.AddAfterSelf(second);
                var props = paragraph.Element(w + "pPr");
                if (props is null) { props = new(w + "pPr"); paragraph.AddFirst(props); }
                var mark = props.Element(w + "rPr");
                if (mark is null) { mark = new(w + "rPr"); props.Add(mark); }
                mark.Add(new XElement(w + "del", new XAttribute(w + "id", "902"), new XAttribute(w + "author", "Public reviewer")));
                foreach (var id in second.Descendants().Where(e => e.Name.LocalName == "docPr")) id.SetAttributeValue("id", "902");
            });
            TestAssert.Equal(2, document.EquationParts.Single(p => p.PartName == document.MainPartName).SourceCount);
            TestAssert.True(Paragraphs(document).All(p => p.InlineTextBoxes.Count == 0), "Both source drawings are intentionally hidden.");
            var diagnostics = new List<OoxPdfDiagnostic>();
            _ = Renderer(mode).RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Consuming a paragraph mark must retain both intentionally hidden equation sources.");
        }
    }

    private static DocxRenderer Renderer(OoxPdfDocxMarkupMode mode) =>
        new(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
            mode == OoxPdfDocxMarkupMode.AllMarkup ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup :
                OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);

    private static IEnumerable<DocxParagraph> Paragraphs(DocxDocument document) =>
        DocxBlockTraversal.EnumerateSourceParagraphs(document).Distinct<DocxParagraph>(ReferenceEqualityComparer.Instance);

    private static XElement BoxRun(XDocument xml, XNamespace w) => xml.Descendants(w + "txbxContent").Single().Ancestors(w + "r").First();

    private static XElement Properties(XElement run, XNamespace w)
    {
        var props = run.Element(w + "rPr");
        if (props is null) { props = new(w + "rPr"); run.AddFirst(props); }
        return props;
    }

    private static void SetHidden(XElement run, XNamespace w, bool hidden)
    {
        var props = Properties(run, w); props.Elements(w + "vanish").Remove();
        props.Add(new XElement(w + "vanish", new XAttribute(w + "val", hidden ? "on" : "off")));
    }

    private static DocxDocument Read((string Fixture, string Part) source, OoxPdfDocxMarkupMode mode,
        Action<XDocument, XDocument, XNamespace> change)
    {
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "InlineBoxMath", source.Fixture + ".docx"));
        using var zip = ZipFile.OpenRead(path);
        var parts = zip.Entries.ToDictionary(entry => entry.FullName, entry =>
        { using var data = new MemoryStream(); using var stream = entry.Open(); stream.CopyTo(data); return data.ToArray(); });
        XDocument Xml(string name) { using var stream = new MemoryStream(parts[name]); return XDocument.Load(stream); }
        var xml = Xml(source.Part); var styles = Xml("word/styles.xml");
        change(xml, styles, "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        parts[source.Part] = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
        parts["word/styles.xml"] = Encoding.UTF8.GetBytes(styles.ToString(SaveOptions.DisableFormatting));
        using var input = TestFixtures.CreateZipPackage(parts);
        return new DocxReader().Read(OoxPackage.Open(input, default), null, default, markupMode: mode);
    }

    private static void AssertPaint(IReadOnlyList<PdfPage> expected, IReadOnlyList<PdfPage> actual, string context)
    {
        TestAssert.Equal(expected.Count, actual.Count);
        for (int page = 0; page < expected.Count; page++)
        {
            TestAssert.True(expected[page].Content == actual[page].Content, "Visible paint differs: " + context + "/page" + page);
            TestAssert.Equal(expected[page].Fonts.Count, actual[page].Fonts.Count);
            TestAssert.Equal(expected[page].Images.Count, actual[page].Images.Count);
            TestAssert.True(expected[page].Annotations.SequenceEqual(actual[page].Annotations), "Hidden boxes retain sibling navigation.");
        }
    }
}
