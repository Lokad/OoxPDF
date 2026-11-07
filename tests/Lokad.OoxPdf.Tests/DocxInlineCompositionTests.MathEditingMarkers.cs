using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxMathEditingMarkersKeepPortablePaintAcrossContainers()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (var source in new[]
        {
            (Fixture: "docx-omml-fraction-inline", Part: "word/document.xml"),
            (Fixture: "CompositionProbes/math-picture-first-table-single-cell", Part: "word/document.xml"),
            (Fixture: "CompositionProbes/equation-revision-header-ins", Part: "word/header1.xml"),
            (Fixture: "CompositionProbes/mixed-note-endnote-table", Part: "word/endnotes.xml")
        })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                    : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var direct = ReadMutation(source.Fixture, mode, false, false, partOverride: source.Part);
            var marked = ReadMutation(source.Fixture, mode, false, false, (xml, w, m) =>
            {
                var arguments = xml.Descendants(m + "num").ToArray();
                TestAssert.True(arguments.Length > 0, "The fixture must exercise mathematical content.");
                int markerId = 12;
                foreach (var argument in arguments)
                {
                    argument.AddFirst(new XElement(w + "proofErr", new XAttribute(w + "type", "spellStart")),
                        new XElement(w + "permStart", new XAttribute(w + "id", markerId),
                            new XAttribute(w + "edGrp", "everyone")));
                    argument.Add(new XElement(w + "permEnd", new XAttribute(w + "id", markerId)),
                        new XElement(w + "proofErr", new XAttribute(w + "type", "spellEnd")));
                    markerId++;
                }
            }, partOverride: source.Part);
            var before = renderer.RenderPagesAsync(direct, null, default).ToBlockingEnumerable().ToArray();
            var after = renderer.RenderPagesAsync(marked, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(before.Length, after.Length);
            for (int page = 0; page < before.Length; page++)
            {
                TestAssert.Equal(before[page].Content, after[page].Content);
                TestAssert.Equal(before[page].Fonts.Count, after[page].Fonts.Count);
                TestAssert.Equal(before[page].Images.Count, after[page].Images.Count);
                TestAssert.Equal(before[page].Annotations.Count, after[page].Annotations.Count);
            }
        }
    }
}
