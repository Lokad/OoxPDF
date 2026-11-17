using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    private const string RepresentationTarget = "https://example.invalid/picture-representation/";

    public static void DocxPictureLinksKeepPaintedRepresentationsAcrossComposedStories()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode, CachedMathGeometry(mode));
            var control = Render("Composition/composed-picture-navigation-report", "raster");
            TestAssert.True(control.Sum(page => Links(page).Count()) >= 4,
                "The composed raster control has navigation across body/cell, footer and normal-note placements.");
            foreach (string variant in new[] { "vector-only", "vector-preview", "broken-vector-preview" })
            {
                var actual = Render("Composition/composed-picture-navigation-report", variant);
                TestAssert.Equal(control.Length, actual.Length);
                for (int index = 0; index < actual.Length; index++)
                {
                    TestAssert.Equal(DocxPageFootnoteNumberingTests.DecodeText(control[index]), DocxPageFootnoteNumberingTests.DecodeText(actual[index]));
                    TestAssert.Equal(string.Join("\n", Links(control[index])), string.Join("\n", Links(actual[index])));
                    CheckRepresentationPaint(actual[index], variant);
                }
            }

            PdfPage[] Render(string fixture, string variant) => renderer.RenderPagesAsync(
                ReadPictureRepresentation(fixture, mode, variant), null, default).ToBlockingEnumerable().ToArray();
        }

        static IEnumerable<string> Links(PdfPage page) => page.Annotations
            .Where(link => link.Uri?.StartsWith(RepresentationTarget, StringComparison.Ordinal) == true)
            .Select(link => FormattableString.Invariant($"{link.Uri}|{link.X:R}|{link.Y:R}|{link.Width:R}|{link.Height:R}"))
            .Order(StringComparer.Ordinal);
    }

    public static void DocxPictureLinksKeepHeaderVectorsAndSelectedRasterFallback()
    {
        var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), OoxPdfDocxMarkupMode.Final,
            OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        foreach (string variant in new[] { "raster", "vector-only", "broken-vector-preview" })
        {
            var page = renderer.RenderPagesAsync(ReadPictureRepresentation("docx-static-image-header",
                OoxPdfDocxMarkupMode.Final, variant), null, default).ToBlockingEnumerable().Single();
            TestAssert.Equal(1, page.Annotations.Count(link => link.Uri?.StartsWith(RepresentationTarget, StringComparison.Ordinal) == true));
            CheckRepresentationPaint(page, variant);
        }
    }

    public static void DocxPictureLinksKeepFailedZeroRotatedAndClippedVectorsUnlinked()
    {
        var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), OoxPdfDocxMarkupMode.Final,
            OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        foreach (string variant in new[] { "broken-raster", "broken-vector-only", "broken-vector-preview-broken", "zero-vector", "rotated-vector", "clipped-vector" })
        {
            string fixture = variant == "clipped-vector" ? "InlineBoxMath/inline-picture-link-properties-cell" : "InlineBoxMath/inline-picture-link-properties-body";
            var document = ReadPictureRepresentation(fixture, OoxPdfDocxMarkupMode.Final, variant);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = renderer.RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().Single();
            TestAssert.Equal(0, page.Annotations.Count(link => link.Uri?.StartsWith(RepresentationTarget, StringComparison.Ordinal) == true));
            if (variant.StartsWith("broken", StringComparison.Ordinal))
            {
                TestAssert.Equal(0, page.Images.Count);
                TestAssert.True(diagnostics.Any(d => d.Id is "IMAGE_UNSUPPORTED_FORMAT" or "SVG_UNSUPPORTED_CONTENT"),
                    "Failed representations retain their existing diagnostic.");
            }
            else if (variant != "zero-vector") CheckRepresentationPaint(page, variant);
            using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(document, null, cancellation.Token).ToBlockingEnumerable().ToArray());
        }
    }

    private static void CheckRepresentationPaint(PdfPage page, string variant)
    {
        PdfContentValidator.ValidatePage(page, 0, default);
        if (variant == "raster" || variant == "broken-vector-preview")
            TestAssert.True(page.Images.Count != 0, "The selected raster representation must actually be emitted.");
        else
        {
            TestAssert.Equal(0, page.Images.Count);
            TestAssert.Contains("0 0.667 0.4 rg", page.Content);
        }
    }

    private static DocxDocument ReadPictureRepresentation(string fixture, OoxPdfDocxMarkupMode mode, string variant) =>
        ReadMutation(fixture, mode, false, false, packageChange: zip =>
        {
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XNamespace pr = "http://schemas.openxmlformats.org/package/2006/relationships";
            XNamespace svg = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
            foreach (var entry in zip.Entries.Where(entry => entry.FullName.StartsWith("word/", StringComparison.Ordinal) &&
                entry.FullName.EndsWith(".xml", StringComparison.Ordinal) && !entry.FullName.Contains('/' + "_rels/", StringComparison.Ordinal)).ToArray())
            {
                XDocument source; using (var stream = entry.Open()) source = XDocument.Load(stream);
                var pictures = source.Descendants(wp + "inline").Where(inline => inline.Descendants(a + "blip").Any()).ToArray();
                if (pictures.Length == 0) continue;
                string part = entry.FullName;
                string relPath = "word/_rels/" + Path.GetFileName(part) + ".rels";
                var relEntry = zip.GetEntry(relPath);
                XDocument rels;
                if (relEntry is null) rels = new XDocument(new XElement(pr + "Relationships"));
                else { using var stream = relEntry.Open(); rels = XDocument.Load(stream); }
                for (int index = 0; index < pictures.Length; index++)
                {
                    var inline = pictures[index];
                    string linkId = "representationLink" + index;
                    var properties = TestAssert.NotNull(inline.Element(wp + "docPr"));
                    properties.Elements(a + "hlinkClick").Remove();
                    properties.Add(new XElement(a + "hlinkClick", new XAttribute(r + "id", linkId)));
                    rels.Root!.Add(new XElement(pr + "Relationship", new XAttribute("Id", linkId),
                        new XAttribute("Type", r.NamespaceName + "/hyperlink"), new XAttribute("TargetMode", "External"),
                        new XAttribute("Target", RepresentationTarget + Path.GetFileNameWithoutExtension(part) + "/" + index)));
                    var blip = inline.Descendants(a + "blip").First();
                    blip.RemoveAttributes(); blip.RemoveNodes();
                    bool vector = variant != "raster" && variant != "broken-raster";
                    bool preview = vector && variant.Contains("preview", StringComparison.Ordinal);
                    blip.SetAttributeValue(r + "embed", vector && !preview ? "representationVector" : "representationRaster");
                    if (preview) blip.Add(new XElement(a + "extLst", new XElement(a + "ext", new XAttribute("uri", "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"),
                        new XElement(svg + "svgBlip", new XAttribute(r + "embed", "representationVector")))));
                    if (variant == "zero-vector") inline.Element(wp + "extent")!.SetAttributeValue("cx", "0");
                    if (variant == "rotated-vector") inline.Descendants(a + "xfrm").First().SetAttributeValue("rot", "5400000");
                }
                foreach (var image in new[] { ("representationRaster", "representation.png"), ("representationVector", "representation.svg") })
                    rels.Root!.Add(new XElement(pr + "Relationship", new XAttribute("Id", image.Item1), new XAttribute("Type", r.NamespaceName + "/image"), new XAttribute("Target", "media/" + image.Item2)));
                if (variant == "clipped-vector")
                {
                    var row = source.Descendants(w + "tr").Single();
                    var props = row.Element(w + "trPr");
                    if (props is null) { props = new XElement(w + "trPr"); row.AddFirst(props); }
                    props.Elements(w + "trHeight").Remove();
                    props.Add(new XElement(w + "trHeight", new XAttribute(w + "val", "20"), new XAttribute(w + "hRule", "exact")));
                }
                entry.Delete(); using (var stream = zip.CreateEntry(part).Open()) source.Save(stream);
                relEntry?.Delete(); using (var stream = zip.CreateEntry(relPath).Open()) rels.Save(stream);
            }
            XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
            var typesEntry = zip.GetEntry("[Content_Types].xml")!;
            XDocument types; using (var stream = typesEntry.Open()) types = XDocument.Load(stream);
            foreach (var type in new[] { ("png", "image/png"), ("svg", "image/svg+xml") })
                if (!types.Root!.Elements(ct + "Default").Any(el => (string?)el.Attribute("Extension") == type.Item1))
                    types.Root.Add(new XElement(ct + "Default", new XAttribute("Extension", type.Item1), new XAttribute("ContentType", type.Item2)));
            typesEntry.Delete(); using (var stream = zip.CreateEntry("[Content_Types].xml").Open()) types.Save(stream);
            using (var stream = zip.CreateEntry("word/media/representation.png").Open()) stream.Write(
                variant is "broken-raster" or "broken-vector-preview-broken" ? "broken"u8.ToArray() : TestFixtures.CreateRgbPng(1, 1, [0, 170, 102]));
            using (var stream = zip.CreateEntry("word/media/representation.svg").Open()) stream.Write(Encoding.UTF8.GetBytes(
                variant.StartsWith("broken", StringComparison.Ordinal) ? "<svg" :
                "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 120 60\"><rect width=\"120\" height=\"60\" fill=\"#00AA66\"/></svg>"));
        });
}
