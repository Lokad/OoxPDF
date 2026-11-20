using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxPictureRecoveryTests
{
    public static void PptxPictureRecoveryRollsBackImageRegistrations()
    {
        foreach (bool grouped in new[] { false, true })
        {
            const string slidePart = "ppt/slides/slide1.xml";
            var parts = FallbackParts("raster-only");
            var slide = XDocument.Parse(Encoding.UTF8.GetString(parts[slidePart]));
            XElement tree = slide.Descendants(P + "spTree").Single();
            XElement first = tree.Element(P + "pic")!;
            XElement failed = PictureAt(3, 3657600);
            failed.Descendants(P + "cNvPr").Single().Add(new XElement(A + "hlinkClick", new XAttribute(R + "id", "unresolved")));
            XElement last = PictureAt(4, 5486400);
            last.Element(P + "blipFill")!.Element(A + "blip")!
                .AddAfterSelf(new XElement(A + "srcRect", new XAttribute("l", "10000")));
            tree.Add(grouped ? new XElement(P + "grpSp", new XElement(P + "grpSpPr"), PictureAt(5, 2286000), failed) : failed, last);
            parts[slidePart] = TestFixtures.Utf8(slide.ToString());

            using var input = TestFixtures.CreateZipPackage(parts);
            using var package = OoxPackage.Open(input, default);
            var diagnostics = new List<OoxPdfDiagnostic>();
            PdfPage page = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
                .RenderPagesAsync(new PptxReader().Read(package, default), package, Observe, default)
                .ToBlockingEnumerable().Single();

            // Group rollback restores the caller's image index. Its next sibling may
            // reuse the name already painted by the group's first child.
            string expected = grouped ? "Im1 Im2" : "Im1 Im3";
            PdfContentValidator.ValidatePage(page, 0, default);
            TestAssert.Equal(expected, string.Join(' ', page.Images.Select(image => image.ResourceName)));
            TestAssert.Equal(expected, string.Join(' ', Regex.Matches(page.Content, @"(?m)^/(Im\d+) Do\r?$")
                .Select(match => match.Groups[1].Value)));
            var diagnosed = diagnostics.Single();
            TestAssert.Equal("PPTX_NODE_RENDER_FAILED", diagnosed.Id);
            TestAssert.Equal(grouped ? "Group" : "Picture", diagnosed.Feature);

            XElement PictureAt(int id, int x)
            {
                var picture = new XElement(first);
                picture.Descendants(P + "cNvPr").Single().SetAttributeValue("id", id);
                picture.Descendants(A + "off").Single().SetAttributeValue("x", x);
                return picture;
            }

            void Observe(OoxPdfDiagnostic diagnostic)
            {
                // Inject a recoverable render failure at the internal renderer seam
                // after the picture has registered its image. In the grouped case,
                // fail the child recovery too, so the parent rewinds its whole group.
                if (diagnostic.Id == "PPTX_UNSUPPORTED_HYPERLINK" ||
                    grouped && diagnostic.Id == "PPTX_NODE_RENDER_FAILED" && diagnostic.Feature == "Picture")
                    throw new InvalidDataException("Injected failure after image paint");

                diagnostics.Add(diagnostic);
            }
        }
    }
}
