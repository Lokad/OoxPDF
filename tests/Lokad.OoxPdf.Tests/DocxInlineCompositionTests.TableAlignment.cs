using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxContainerTableGridsAlignWholeTablesInBodyAndHeader()
    {
        // A 595.3pt sheet with 72pt margins and a 180pt table places the
        // centered grid at 207.65pt and the right-aligned grid at 343.3pt.
        foreach (var (fixture, part, alignment, expectedLeft) in new[]
        {
            ("body-zero", "word/document.xml", "center", 207.65d),
            ("header-zero", "word/static-inline-ordinary-header-box.xml", "right", 343.3d)
        })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            DocxDocument Source(string? value, int leftIndent) => ReadMutation("TableGridOrigins/Containers/" + fixture,
                mode, false, false, (xml, w, _) =>
                {
                    XElement properties = xml.Descendants(w + "tblPr").Single();
                    properties.Element(w + "tblInd")!.SetAttributeValue(w + "w", leftIndent);
                    if (value is not null) properties.Element(w + "tblW")!.AddAfterSelf(new XElement(w + "jc", new XAttribute(w + "val", value)));
                }, part);

            var renderer = MathContentControlRenderer(mode);
            var control = renderer.RenderPagesAsync(Source(null, 0), null, default).ToBlockingEnumerable().Single();
            var actual = renderer.RenderPagesAsync(Source(alignment, 480), null, default).ToBlockingEnumerable().Single();
            var controlFills = DocxParagraphShadingTests.ReadRectangles(control.Content).Where(fill => fill.Color == "DBEAFE").ToArray();
            var actualFills = DocxParagraphShadingTests.ReadRectangles(actual.Content).Where(fill => fill.Color == "DBEAFE").ToArray();
            TestAssert.Equal(2, controlFills.Length);
            TestAssert.Equal(2, actualFills.Length);
            double scale = controlFills[0].Width / 180d;
            double shift = expectedLeft * scale - controlFills[0].X;
            foreach (var (before, after) in controlFills.Zip(actualFills))
            {
                Close(expectedLeft * scale, after.X);
                Close(before.Y, after.Y);
                Close(before.Width, after.Width);
                Close(before.Height, after.Height);
            }
            var beforeText = DocxCommentsTests.ReadEmbeddedGlyphTextShows(control, true).ToArray();
            var afterText = DocxCommentsTests.ReadEmbeddedGlyphTextShows(actual, true).ToArray();
            TestAssert.True(beforeText.Select(item => item.Text).SequenceEqual(afterText.Select(item => item.Text)),
                "Alignment preserves all painted prose and equations in order.");
            foreach (string label in new[] { "BOX TABLE NORTH", "BOX TABLE SOUTH" })
            {
                var before = beforeText.Single(item => item.Text.Trim() == label);
                var after = afterText.Single(item => item.Text.Trim() == label);
                Close(before.X + shift, after.X);
                Close(before.Y, after.Y);
            }
            var beforeLink = control.Annotations.Single(link => link.Uri == "https://example.invalid/static-box-table");
            var afterLink = actual.Annotations.Single(link => link.Uri == beforeLink.Uri);
            Close(beforeLink.X + shift, afterLink.X);
            Close(beforeLink.Y, afterLink.Y);
            Close(beforeLink.Width, afterLink.Width);
            Close(beforeLink.Height, afterLink.Height);
            var beforeTarget = control.Annotations.Single(link => link.Destination.HasValue).Destination!.Value;
            var afterTarget = actual.Annotations.Single(link => link.Destination.HasValue).Destination!.Value;
            TestAssert.Equal(beforeTarget.PageIndex, afterTarget.PageIndex);
            Close(beforeTarget.Left!.Value + shift, afterTarget.Left!.Value);
            Close(beforeTarget.Top!.Value, afterTarget.Top!.Value);

            void Close(double expected, double observed) => TestAssert.True(Math.Abs(expected - observed) < .004d,
                $"{fixture}/{mode} expected {expected}pt, observed {observed}pt.");
        }
    }

    public static void DocxContainerTableGridsResolveAlignmentFromTableStyles()
    {
        const string fixture = "TableGridOrigins/Containers/body-zero";
        DocxTable Read(string? direct, bool mixedRows = false, bool rtl = false) => ReadMutation(fixture,
            OoxPdfDocxMarkupMode.Final, false, false, (xml, w, _) =>
            {
                XElement properties = xml.Descendants(w + "tblPr").Single();
                properties.AddFirst(new XElement(w + "tblStyle", new XAttribute(w + "val", "DerivedAlignment")));
                if (direct is not null) properties.Element(w + "tblW")!.AddAfterSelf(new XElement(w + "jc", new XAttribute(w + "val", direct)));
                if (mixedRows)
                {
                    XElement row = xml.Descendants(w + "tr").First();
                    row.Element(w + "tc")!.AddBeforeSelf(new XElement(w + "trPr", new XElement(w + "jc", new XAttribute(w + "val", "left"))));
                }
            }, packageChange: zip =>
            {
                var entry = zip.GetEntry("word/styles.xml")!;
                XDocument xml;
                using (var input = entry.Open()) xml = XDocument.Load(input);
                XElement root = xml.Root!;
                var w = root.Name.Namespace;
                root.Add(new XElement(w + "style", new XAttribute(w + "type", "table"), new XAttribute(w + "styleId", "BaseAlignment"),
                    new XElement(w + "tblPr", rtl ? new XElement(w + "bidiVisual") : null,
                        new XElement(w + "jc", new XAttribute(w + "val", "center")))),
                    new XElement(w + "style", new XAttribute(w + "type", "table"), new XAttribute(w + "styleId", "DerivedAlignment"),
                        new XElement(w + "basedOn", new XAttribute(w + "val", "BaseAlignment"))));
                entry.Delete();
                using var output = zip.CreateEntry("word/styles.xml").Open();
                xml.Save(output);
            }).Tables.Single();

        TestAssert.Equal(DocxTextAlignment.Center, Read(null).Alignment);
        TestAssert.Equal(DocxTextAlignment.Right, Read("right").Alignment);
        TestAssert.Equal(DocxTextAlignment.Left, Read("left").Alignment);
        TestAssert.Equal(DocxTextAlignment.Left, Read(null, mixedRows: true).Alignment);
        TestAssert.Equal(DocxTextAlignment.Left, Read(null, rtl: true).Alignment);
    }
}
