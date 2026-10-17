using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionPlainCellPageNotesUseMeasuredSlotsAcrossRunSizes()
    {
        foreach (string typography in new[] { "calibri-12-zero", "arial-11-zero" })
        foreach (bool largerFirstRun in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadMutation("CompositionProbes/cell-page-notes-plain-" + typography, mode, false, false,
                (xml, w, _) =>
                {
                    if (!largerFirstRun) return;
                    var run = xml.Descendants(w + "r").First(run => run.Element(w + "t") is not null);
                    var properties = run.Element(w + "rPr");
                    if (properties is null) { properties = new XElement(w + "rPr"); run.AddFirst(properties); }
                    var size = properties.Element(w + "sz");
                    if (size is null) { size = new XElement(w + "sz"); properties.Add(size); }
                    size.SetAttributeValue(w + "val", "30");
                });
            var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var layout = engine.Create(source, new DocxTests.FamilyWidthTextMeasurer(), default);
            int owners = 0;
            foreach (var page in layout.Pages)
            foreach (var row in page.Items.OfType<DocxTableRowLayout>())
            foreach (var cell in row.Cells)
            foreach (var line in cell.TextLines)
            {
                if (line.SourceParagraph is null || line.SourceParagraph.InlineReferences.Count == 0) continue;
                TestAssert.True(line.InlineFrame is null, "Ordinary text retains its existing placement path.");
                TestAssert.True(DocxInlineLineGeometry.TryGetOccupiedBounds(line, out double bottom, out double top),
                    "Placement publishes a finite occupied slot without an object or font-specific admission rule.");
                var clip = DocxTableCellClipGeometry.Resolve(row, cell);
                TestAssert.True(bottom >= clip.Y - .000001d && top <= clip.Y + clip.Height + .000001d,
                    "The actual ordinary line slot is contained in the cell paint clip.");
                owners++;
            }
            TestAssert.True(owners >= 4, "The regression observes every ordinary source-note paragraph.");
            TestAssert.True(DocxPageFootnoteNumbering.TryCreateDocument(source, layout, out var numbered, out var pages,
                allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences),
                "Complete ordinary cell slots reuse the shared numbering owner contract.");
            TestAssert.Equal("09,10,09,10", string.Join(',', DocxBlockTraversal.EnumerateBodyParagraphs(numbered)
                .SelectMany(paragraph => paragraph.InlineReferences).Select(reference => reference.DisplayText)));
            var final = engine.Create(numbered, new DocxTests.FamilyWidthTextMeasurer(), default);
            TestAssert.True(DocxPageFootnoteNumbering.HasStableNumberingPages(numbered, final, pages,
                allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences),
                "Changing generated marker advances keeps the complete source/note page map stable.");
        }
    }

    public static void DocxInlineCompositionPlainCellPageNotesKeepIndependentLabelsAndDestinations()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent ordinary cell controls use installed Windows fonts.");
        CheckPageNoteMath(null, "cell-page-notes-plain-office.json");
    }
}
