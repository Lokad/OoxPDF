using System.IO.Compression;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Tests;

internal static partial class PortableRenderingTests
{
    public static void PptxInvalidParagraphLevelsKeepShapeAndCellContent()
    {
        foreach (bool table in new[] { false, true })
        {
            string witness = table ? "$120K" : "Hello";
            var firstLevel = RenderOptionalTextAttributes(witness, table);
            TestAssert.Equal(firstLevel.Content, RenderOptionalTextAttributes(witness, table, level: "unknown").Content);
            TestAssert.Equal(firstLevel.Content, RenderOptionalTextAttributes(witness, table, level: "2147483648").Content);

            var secondLevel = RenderOptionalTextAttributes(witness, table, level: "1");
            TestAssert.True(secondLevel.StartX > firstLevel.StartX + 7f,
                "The valid second-level paragraph style still indents the painted text.");
        }
    }

    public static void PptxInvalidCharacterSpacingKeepsInheritedSpacingAndContent()
    {
        foreach (bool table in new[] { false, true })
        {
            string witness = table ? "$120K" : "Hello";
            var inherited = RenderOptionalTextAttributes(witness, table, inheritedSpacing: "200");
            TestAssert.Equal(inherited.Content,
                RenderOptionalTextAttributes(witness, table, spacing: "", inheritedSpacing: "200").Content);
            var defaultSpacing = RenderOptionalTextAttributes(witness, table);
            TestAssert.True(inherited.GlyphOriginSpan > defaultSpacing.GlyphOriginSpan + 2f,
                "Valid inherited spacing still expands the authored glyph advances.");
            TestAssert.Equal(defaultSpacing.Content, RenderOptionalTextAttributes(witness, table, inheritedSpacing: "invalid").Content);
            TestAssert.Equal(defaultSpacing.Content,
                RenderOptionalTextAttributes(witness, table, spacing: "0", inheritedSpacing: "200").Content);

            var tightened = RenderOptionalTextAttributes(witness, table, spacing: "-125", inheritedSpacing: "200");
            TestAssert.True(tightened.GlyphOriginSpan < defaultSpacing.GlyphOriginSpan - 2f,
                "Valid signed direct spacing still tightens the authored glyph advances.");
        }
    }

    private static (string Content, float StartX, float GlyphOriginSpan) RenderOptionalTextAttributes(string witness, bool table,
        string? level = null, string? spacing = null, string? inheritedSpacing = null)
    {
        string fixture = table ? "pptx-ladder-10-composite-table-port.pptx" : "pptx-ladder-02-plain-text.pptx";
        using var bytes = new MemoryStream();
        using (FileStream source = File.OpenRead(Case(fixture)))
            source.CopyTo(bytes);

        using (var archive = new ZipArchive(bytes, ZipArchiveMode.Update, leaveOpen: true))
        {
            const string slidePath = "ppt/slides/slide1.xml";
            ZipArchiveEntry entry = archive.GetEntry(slidePath)!;
            XDocument slide;
            using (Stream part = entry.Open())
                slide = XDocument.Load(part);
            XElement run = slide.Descendants(DrawingNamespace + "r").Single(item =>
                (string?)item.Element(DrawingNamespace + "t") == witness);
            XElement paragraph = run.Parent!;
            XElement? properties = paragraph.Element(DrawingNamespace + "pPr");
            if (properties is null)
            {
                properties = new XElement(DrawingNamespace + "pPr");
                paragraph.AddFirst(properties);
            }
            properties.SetAttributeValue("lvl", level);
            if (inheritedSpacing is not null)
                properties.Add(new XElement(DrawingNamespace + "defRPr", new XAttribute("spc", inheritedSpacing)));
            run.Element(DrawingNamespace + "rPr")!.SetAttributeValue("spc", spacing);
            paragraph.Parent!.Element(DrawingNamespace + "lstStyle")!
                .Add(new XElement(DrawingNamespace + "lvl2pPr", new XAttribute("marL", "101600")));

            entry.Delete();
            using Stream output = archive.CreateEntry(slidePath).Open();
            slide.Save(output);
        }

        bytes.Position = 0;
        using var package = OoxPackage.Open(bytes, default);
        var diagnostics = new List<OoxPdfDiagnostic>();
        PdfPage page = new PptxRenderer(new PinnedResolver())
            .RenderPagesAsync(new PptxReader().Read(package, default), package, diagnostics.Add, default)
            .ToBlockingEnumerable().Single();
        TestAssert.Equal(0, diagnostics.Count);
        var glyphs = ReadPaintedSourceLabel(page, witness);
        TestAssert.Equal(witness.Length, glyphs.Count);
        return (page.Content, glyphs[0].Origin.X, glyphs[^1].Origin.X - glyphs[0].Origin.X);
    }
}
