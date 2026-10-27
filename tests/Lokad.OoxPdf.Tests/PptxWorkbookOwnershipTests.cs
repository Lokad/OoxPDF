using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxWorkbookOwnershipTests
{
    public static void PptxWorkbookOptionalPartsIgnoreUnrelatedStylesAndStrings()
    {
        foreach (string kind in new[] { "styles", "sharedStrings" })
        {
            var parts = Parts(PptxTests.EmbeddedChartWorkbook());
            RemoveRelationship(parts, kind);
            parts["xl/" + kind + ".xml"] = TestFixtures.Utf8("<unreferenced-broken-xml");
            var workbook = Read(parts);
            TestAssert.True(workbook.Date1904, "The related workbook's date system remains available.");
            var value = workbook.ReadRangeCells("Sheet1!$B$2", default).Single();
            TestAssert.Equal("8.2", value.RawValue);
            TestAssert.True(value.ColumnHidden, "Related worksheet visibility must survive orphan metadata.");
            TestAssert.Equal(kind == "styles" ? (int?)null : 165, value.StyleNumberFormatId);
            var label = workbook.ReadRangeCells("Sheet1!$A$2", default).Single();
            TestAssert.Equal(kind == "sharedStrings" ? 0 : 2, label.SharedStringRunCount);
            if (kind == "styles") TestAssert.Equal("North", label.Text);
        }
    }

    public static void PptxWorkbookOptionalPartsKeepRelatedMetadataAndAlternateTargets()
    {
        foreach (bool alternate in new[] { false, true })
        {
            var parts = Parts(PptxTests.EmbeddedChartWorkbook());
            if (alternate)
            {
                foreach (string kind in new[] { "styles", "sharedStrings" })
                {
                    string name = "xl/" + kind + ".xml", target = "metadata/" + kind + ".xml";
                    parts["xl/" + target] = parts[name];
                    parts[name] = TestFixtures.Utf8("<orphan-at-the-usual-name");
                    ChangeXml(parts, "xl/_rels/workbook.xml.rels", xml =>
                        Relationship(xml, kind).SetAttributeValue("Target", target));
                    ChangeXml(parts, "[Content_Types].xml", xml =>
                    {
                        XElement entry = xml.Root!.Elements().Single(element => (string?)element.Attribute("PartName") == "/" + name);
                        var related = new XElement(entry);
                        related.SetAttributeValue("PartName", "/xl/" + target);
                        xml.Root.Add(related);
                    });
                }
            }
            var workbook = Read(parts);
            var value = workbook.ReadRangeCells("Sheet1!$B$2", default).Single();
            TestAssert.Equal(165, value.StyleNumberFormatId);
            TestAssert.Equal("m/d/yy", value.StyleNumberFormatCode);
            TestAssert.True(value.StyleNumberFormatIsDateLike && value.StyleAppliesNumberFormat == true,
                "Owned custom date formatting must remain intact.");
            TestAssert.True(value.ColumnHidden && workbook.ReadRangeCells("Sheet1!$B$3", default).Single().RowHidden,
                "Related row/column visibility remains available.");
            TestAssert.Equal("North", workbook.ReadRangeCells("Sheet1!$A$2", default).Single().Text);
        }
    }

    public static void PptxWorkbookOptionalPartsRetainReferencedFailuresAndCancellation()
    {
        foreach (string kind in new[] { "styles", "sharedStrings" })
        {
            var parts = Parts(PptxTests.EmbeddedChartWorkbook());
            parts["xl/" + kind + ".xml"] = TestFixtures.Utf8("<referenced-broken-xml");
            TestAssert.Throws<InvalidDataException>(() => Read(parts));
        }
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => Read(Parts(PptxTests.EmbeddedChartWorkbook()), cancellation.Token));
    }

    public static void PptxWorkbookOptionalPartsKeepCachedChartPaintWhenMetadataIsUnrelated()
    {
        string source = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
            "pptx-ladder-11-chart-bar-clustered-port.pptx"));
        var original = Parts(File.ReadAllBytes(source));
        var baseline = Convert(original);
        foreach (string kind in new[] { "styles", "sharedStrings" })
            foreach (bool orphan in new[] { false, true })
            {
                var package = new Dictionary<string, byte[]>(original);
                const string embedded = "ppt/embeddings/Microsoft_Excel_Sheet1.xlsx";
                var workbook = Parts(package[embedded]);
                RemoveRelationship(workbook, kind);
                string name = "xl/" + kind + ".xml";
                if (orphan) workbook[name] = TestFixtures.Utf8("<unrelated-broken-xml");
                else workbook.Remove(name);
                using (var bytes = TestFixtures.CreateZipPackage(workbook)) package[embedded] = bytes.ToArray();
                var current = Convert(package);
                TestAssert.True(baseline.Bytes.AsSpan().SequenceEqual(current.Bytes),
                    "An unrelated optional workbook part must not erase or change cached chart paint.");
                TestAssert.True(baseline.Diagnostics.SequenceEqual(current.Diagnostics),
                    "Unrelated metadata must not introduce chart-node recovery diagnostics.");
            }
    }

    private static (byte[] Bytes, OoxPdfDiagnostic[] Diagnostics) Convert(Dictionary<string, byte[]> parts)
    {
        using var input = TestFixtures.CreateZipPackage(parts);
        using var output = new MemoryStream();
        var diagnostics = new List<OoxPdfDiagnostic>();
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Pptx,
            FontResolver = new PortableRenderingTests.PinnedResolver(),
            DiagnosticSink = diagnostics.Add
        });
        return (output.ToArray(), diagnostics.ToArray());
    }

    private static Dictionary<string, byte[]> Parts(byte[] bytes)
    {
        using var input = new MemoryStream(bytes, writable: false);
        using var archive = new ZipArchive(input, ZipArchiveMode.Read);
        return archive.Entries.ToDictionary(entry => entry.FullName, entry =>
        {
            using var stream = entry.Open(); using var result = new MemoryStream(); stream.CopyTo(result); return result.ToArray();
        });
    }

    private static XElement Relationship(XDocument xml, string kind) => xml.Root!.Elements()
        .Single(element => ((string?)element.Attribute("Type"))?.EndsWith("/" + kind, StringComparison.Ordinal) == true);

    private static void RemoveRelationship(Dictionary<string, byte[]> parts, string kind) =>
        ChangeXml(parts, "xl/_rels/workbook.xml.rels", xml => Relationship(xml, kind).Remove());

    private static void ChangeXml(Dictionary<string, byte[]> parts, string name, Action<XDocument> change)
    {
        XDocument xml = XDocument.Parse(Encoding.UTF8.GetString(parts[name]));
        change(xml); parts[name] = TestFixtures.Utf8(xml.ToString());
    }

    private static PptxRenderer.ChartWorkbookData Read(Dictionary<string, byte[]> parts, CancellationToken token = default)
    {
        using var input = TestFixtures.CreateZipPackage(parts);
        OoxPackage package = OoxPackage.Open(input, default);
        MethodInfo read = typeof(PptxRenderer).GetMethod("ReadWorkbookDataCore", BindingFlags.NonPublic | BindingFlags.Static)!;
        try { return (PptxRenderer.ChartWorkbookData)read.Invoke(null, [package, token])!; }
        catch (TargetInvocationException ex) { throw ex.InnerException ?? ex; }
    }
}
