using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PortableRenderingTests
{
    public static void PortableRenderingFloatingPictureLayersPreserveBooleanLexicalEquivalence()
    {
        var resolver = new PinnedResolver();
        foreach (var (token, behind) in new (string?, bool)[] {
            ("true", true), ("1", true), (" true ", true), ("\t1\r\n", true),
            (null, false), ("", true), ("on", true), (" on ", false),
            ("false", false), (" false ", false), ("0", false), ("off", false),
            (" ", false), ("unknown", false), ("\u00A0true\u00A0", false) })
        {
            AssertPictureLayer(RenderPictureLayerSource("body", token, resolver), behind);
        }
    }

    public static void PortableRenderingRelatedPictureLayersShareSourceBooleanSemantics()
    {
        var resolver = new PinnedResolver();
        foreach (string owner in new[] { "header", "footer", "footnote" })
        {
            PdfPage canonical = RenderPictureLayerSource(owner, "true", resolver);
            PdfPage padded = RenderPictureLayerSource(owner, " true ", resolver);
            PdfPage ahead = RenderPictureLayerSource(owner, "false", resolver);
            AssertPictureLayer(canonical, behind: true);
            AssertPictureLayer(padded, behind: true);
            AssertPictureLayer(ahead, behind: false);
            TestAssert.Equal(canonical.Content, padded.Content);
        }
    }

    private static void AssertPictureLayer(PdfPage page, bool behind)
    {
        const string label = "VISIBLE WORDS";
        TestAssert.Equal(1, page.Images.Count);
        var frames = DocxNoteClippingTests.ReadPhysicalPaintFrames(page).ToArray();
        int picture = Array.FindIndex(frames, frame => frame.Block == "/Im1 Do");
        int text = Array.FindIndex(frames, frame => string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(
            new PdfPage(page.Width, page.Height, frame.Block, page.Fonts)).Select(show => show.Text)).Contains(label, StringComparison.Ordinal));
        TestAssert.True(picture >= 0 && text >= 0, "Both authored picture and words must paint.");
        TestAssert.True(behind ? picture < text : text < picture,
            $"The source paint layer must put the opaque picture {(behind ? "behind" : "ahead of")} its words.");
        var image = frames[picture];
        TestAssert.True(image.Transform.M11 > 0 && image.Transform.M22 > 0 &&
            image.Transform.M12 == 0 && image.Transform.M21 == 0, "This source has an unrotated opaque picture.");
        foreach (var glyph in ReadPaintedSourceLabel(page, label))
        {
            // Ink/clip visibility alone cannot detect an opaque later picture.
            // Ensure the explicit paint-order assertion above protects overlapping content.
            TestAssert.True(glyph.Ink.X >= image.Transform.M31 - .01d &&
                glyph.Ink.Y >= image.Transform.M32 - .01d &&
                glyph.Ink.X + glyph.Ink.Width <= image.Transform.M31 + image.Transform.M11 + .01d &&
                glyph.Ink.Y + glyph.Ink.Height <= image.Transform.M32 + image.Transform.M22 + .01d,
                "The source picture must overlap every word, making paint order consequential.");
        }
    }

    private static PdfPage RenderPictureLayerSource(string owner, string? token, PinnedResolver resolver)
    {
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        XNamespace rel = "http://schemas.openxmlformats.org/package/2006/relationships";
        XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
        XElement paragraph = XElement.Parse("""
            <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                 xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <w:r><w:rPr><w:rFonts w:ascii="DejaVu Sans" w:hAnsi="DejaVu Sans"/><w:sz w:val="24"/></w:rPr><w:t>VISIBLE WORDS</w:t></w:r>
              <w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" locked="0" layoutInCell="1" allowOverlap="1">
                <wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
                <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="2540000" cy="762000"/>
                <wp:wrapNone/><wp:docPr id="1" name="Opaque red picture"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="Red"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                    <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2540000" cy="762000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>
                </a:graphicData></a:graphic>
              </wp:anchor></w:drawing></w:r>
            </w:p>
            """);
        paragraph.Descendants(wp + "anchor").Single().SetAttributeValue("behindDoc", token);
        XElement section = new(w + "sectPr", new XElement(w + "pgSz", new XAttribute(w + "w", 12240), new XAttribute(w + "h", 15840)),
            new XElement(w + "pgMar", new XAttribute(w + "top", 1440), new XAttribute(w + "bottom", 1440),
                new XAttribute(w + "left", 1440), new XAttribute(w + "right", 1440),
                new XAttribute(w + "header", 720), new XAttribute(w + "footer", 720)));
        XElement types = new(ct + "Types",
            new XElement(ct + "Default", new XAttribute("Extension", "rels"), new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")),
            new XElement(ct + "Default", new XAttribute("Extension", "xml"), new XAttribute("ContentType", "application/xml")),
            new XElement(ct + "Default", new XAttribute("Extension", "png"), new XAttribute("ContentType", "image/png")),
            Override("document", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"));
        XElement imageRelationships = new(rel + "Relationships", Relationship("rImage", "image", "media/red.png"));
        XElement body = new(w + "body");
        var parts = new Dictionary<string, byte[]> { ["word/media/red.png"] = TestFixtures.CreateRgbPng(1, 1, [255, 0, 0]) };
        parts["_rels/.rels"] = Bytes(new XElement(rel + "Relationships", Relationship("rDoc", "officeDocument", "word/document.xml")));
        if (owner == "body")
        {
            body.Add(paragraph);
            parts["word/_rels/document.xml.rels"] = Bytes(imageRelationships);
        }
        else
        {
            XElement control = new(w + "p", new XElement(w + "r", new XElement(w + "t", "Body control")));
            body.Add(control);
            if (owner == "footnote")
            {
                control.Add(new XElement(w + "r", new XElement(w + "footnoteReference", new XAttribute(w + "id", 1))));
                parts["word/footnotes.xml"] = Bytes(new XElement(w + "footnotes", new XElement(w + "footnote", new XAttribute(w + "id", 1), paragraph)));
            }
            else
            {
                section.AddFirst(new XElement(w + (owner + "Reference"), new XAttribute(w + "type", "default"), new XAttribute(r + "id", "rStory")));
                parts["word/" + owner + ".xml"] = Bytes(new XElement(w + (owner == "header" ? "hdr" : "ftr"), paragraph));
            }
            string part = owner == "footnote" ? "footnotes" : owner;
            types.Add(Override(part, "application/vnd.openxmlformats-officedocument.wordprocessingml." + part + "+xml"));
            parts["word/_rels/document.xml.rels"] = Bytes(new XElement(rel + "Relationships", Relationship("rStory", part, part + ".xml")));
            parts["word/_rels/" + part + ".xml.rels"] = Bytes(imageRelationships);
        }
        body.Add(section);
        parts["word/document.xml"] = Bytes(new XElement(w + "document", body));
        parts["[Content_Types].xml"] = Bytes(types);
        using var stream = TestFixtures.CreateZipPackage(parts);
        using var package = OoxPackage.Open(stream, default);
        var document = new DocxReader().Read(package, null, default, OoxPdfDocxMarkupMode.Final);
        return new DocxRenderer(resolver, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .RenderPagesAsync(document, null, default).ToBlockingEnumerable().Single();

        XElement Relationship(string id, string kind, string target) => new(rel + "Relationship", new XAttribute("Id", id),
            new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" + kind), new XAttribute("Target", target));
        XElement Override(string part, string type) => new(ct + "Override", new XAttribute("PartName", "/word/" + part + ".xml"), new XAttribute("ContentType", type));
        static byte[] Bytes(XElement xml) => Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
    }
}
