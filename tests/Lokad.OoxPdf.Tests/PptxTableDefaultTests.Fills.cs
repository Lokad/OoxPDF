using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxTableDefaultTests
{
    public static void PptxTableCellFillsKeepExplicitRemovalAndDirectAlpha()
    {
        foreach (bool styled in new[] { true, false })
        foreach (string treatment in new[] { "absent", "none", "transparent", "red", "unsupported" })
        {
            WithSlide("inherited", xml =>
            {
                if (!styled) xml.Descendants(A + "tableStyleId").Remove();
                // The fill contract does not depend on installed fonts or white header text.
                foreach (XElement run in xml.Descendants(A + "rPr"))
                {
                    run.SetAttributeValue("b", "0");
                    run.Element(A + "latin")?.Remove();
                    run.Add(new XElement(A + "latin", new XAttribute("typeface", "DejaVu Sans")));
                    run.AddFirst(new XElement(A + "solidFill", new XElement(A + "srgbClr", new XAttribute("val", "000000"))));
                }
                XElement cell = xml.Descendants(A + "tcPr").First();
                if (treatment == "none") cell.Add(new XElement(A + "noFill"));
                if (treatment == "unsupported") cell.Add(new XElement(A + "gradFill"));
                if (treatment is "red" or "transparent")
                    cell.Add(new XElement(A + "solidFill", new XElement(A + "srgbClr", new XAttribute("val", "FF0000"),
                        treatment == "transparent" ? new XElement(A + "alpha", new XAttribute("val", "0")) : null)));
            }, path =>
            {
                using var input = File.OpenRead(path);
                using var package = OoxPackage.Open(input, default);
                PdfPage page = new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(
                    new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
                bool directFill = treatment is "red" or "transparent";
                bool inheritedFill = styled && treatment is ("absent" or "unsupported");
                TestAssert.Equal(directFill || inheritedFill, page.Content.Contains("72 396 216 72 re f*", StringComparison.Ordinal));
                if (directFill) TestAssert.Contains("1 0 0 rg", page.Content);
                if (inheritedFill) TestAssert.Contains("0.306 0.655 0.18 rg", page.Content);
                TestAssert.Equal(treatment == "transparent", page.ExtGStates.Any(state => state.FillAlpha == 0d));
                foreach (string word in new[] { "Head", "Body" })
                    TestAssert.Equal(word.Length, PortableRenderingTests.ReadPaintedSourceLabel(page, word).Count);
                PdfContentValidator.ValidatePage(page, 0, default);
                if (styled && treatment == "none")
                {
                    using var output = new MemoryStream();
                    long written = PdfDocumentWriter.WriteStagedAsync(output, AsyncEnumerable.ToAsyncEnumerable(new[] { page }),
                        new OoxConversionLimits(), null, default).GetAwaiter().GetResult();
                    TestAssert.Equal(output.Length, written);
                    TestAssert.True(Encoding.ASCII.GetString(output.ToArray()).StartsWith("%PDF-", StringComparison.Ordinal),
                        "The no-fill page also passes ordinary staged PDF serialization.");
                }
            });
        }
    }
}
