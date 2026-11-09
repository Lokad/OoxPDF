using System.IO.Compression;
using System.Reflection;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxNoteClippingTests
{
    public static void DocxPartiallyClippedNoteRowsReportTextAndMathContent()
    {
        foreach (bool math in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            using var input = BoundaryInput(math, overflow: true, exact: false);
            byte[] source = input.ToArray();
            var diagnostics = new List<OoxPdfDiagnostic>();
            using var output = new MemoryStream();
            OoxPdfConverter.Convert(input, output, Options(mode, diagnostics.Add));
            var warning = diagnostics.Single(d => d.Id == "DOCX_CLIPPED_NOTE_TABLE_ROW");
            TestAssert.Equal("/word/endnotes.xml", warning.PartName!);
            TestAssert.True(warning.PageIndex is > 0, "Identify the page of the partially clipped source line.");
            TestAssert.Equal(OoxPdfSeverity.Warning, warning.Severity);
            TestAssert.Equal("Clipped", warning.Fallback!);
            TestAssert.True(!diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_EQUATION") &&
                (!math || diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_EQUATION")),
                "The matrix is supported; its final frame clipping needs a distinct warning.");
            using var unobserved = new MemoryStream(source);
            using var control = new MemoryStream();
            OoxPdfConverter.Convert(unobserved, control, Options(mode, null));
            TestAssert.True(output.ToArray().AsSpan().SequenceEqual(control.ToArray()),
                "Observing partial clipping changes no rendered PDF bytes.");
        }
    }

    public static void DocxPartialNoteRowsKeepFittingAndIntentionalClipExceptions()
    {
        foreach (bool math in new[] { false, true })
        foreach (bool exact in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            using var input = BoundaryInput(math, overflow: exact, exact);
            using var output = new MemoryStream();
            var diagnostics = new List<OoxPdfDiagnostic>();
            OoxPdfConverter.Convert(input, output, Options(mode, diagnostics.Add));
            TestAssert.True(output.Length > 0 && !diagnostics.Any(d => d.Id == "DOCX_CLIPPED_NOTE_TABLE_ROW"),
                "Fitting text/math frames and authored exact heights retain their own clipping policy.");
            TestAssert.True(!diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_EQUATION"), "The control must retain supported math.");
        }
    }

    public static void DocxUnfitNoteRowsReportClippingWithoutChangingPdf()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            byte[] actual = Convert(400, false, mode, diagnostics.Add);
            byte[] withoutObserver = Convert(400, false, mode, null);
            var warning = diagnostics.Single(d => d.Id == "DOCX_CLIPPED_NOTE_TABLE_ROW");
            TestAssert.Equal("/word/endnotes.xml", warning.PartName!);
            TestAssert.True(warning.PageIndex is > 0, "Identify the actual page using the public one-based convention.");
            TestAssert.Equal(OoxPdfSeverity.Warning, warning.Severity);
            TestAssert.Equal("Clipped", warning.Fallback!);
            TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_EQUATION") &&
                !diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_EQUATION"), "Supported mathematical neighbors remain painted.");
            TestAssert.True(actual.AsSpan().SequenceEqual(withoutObserver), "Observing clipping must preserve exact PDF bytes.");
        }
    }

    public static void DocxFittingAndAuthoredExactNoteRowsKeepTheirClippingPolicy()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool exact in new[] { false, true })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            byte[] pdf = Convert(exact ? 400 : 24, exact, mode, diagnostics.Add);
            TestAssert.True(pdf.Length > 0, "The fitting or intentionally clipped document still converts.");
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_CLIPPED_NOTE_TABLE_ROW"),
                "Do not label fitting rows or authored exact-height clipping as unsupported row fragmentation.");
        }
    }

    public static void DocxNoteClippingUsesContainedFramesAndEmissionCoordinates()
    {
        using var input = Input(400, false);
        using var package = OoxPackage.Open(input, default);
        var source = new DocxReader().Read(package, null, default, OoxPdfDocxMarkupMode.Final);
        var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default);
        var row = layout.RelatedStories.Single(story => story.Story.Id == "37").TableRows.Single();
        var inspect = typeof(DocxRenderer).GetMethod("HasClippedNoteTableLines", BindingFlags.Static | BindingFlags.NonPublic)!
            .CreateDelegate<Func<DocxTableRowLayout, PdfRectangle, double, CancellationToken, bool>>();
        var outside = new PdfRectangle(0d, row.Y + row.Height + 100d, 1000d, 100d);
        TestAssert.True(inspect(row, outside, 0d, default), "A complete source-contained line is wholly outside the final clip.");
        var contained = outside with { Y = row.Y - 1d, Height = row.Height + 2d };
        TestAssert.True(!inspect(row, contained, 0d, default), "Retain fitting geometry.");
        var line = DocxTableCellClipGeometry.EnumerateCellLines(row, null, default).First(owner =>
            !string.IsNullOrWhiteSpace(owner.Line.Text)).Line;
        TestAssert.True(DocxInlineLineGeometry.TryGetOccupiedBounds(line, out double bottom, out double top),
            "Use an actual measured source line for partial clipping.");
        var single = row with { Cells = row.Cells.Select(cell => cell with {
            TextLines = cell.TextLines.Where(candidate => ReferenceEquals(candidate, line)).ToArray() }).ToArray() };
        var partial = contained with { Y = (bottom + top) / 2d };
        TestAssert.True(inspect(single, partial, 0d, default), "A line crossing the final frame is partly clipped.");
        TestAssert.True(!inspect(row, outside, row.Height + 1000d, default), "Use the same baseline translation as painting.");
        TestAssert.True(!inspect(row with { FragmentCount = 2 }, outside, 0d, default), "Real continuations may paint on another page.");
        TestAssert.True(!inspect(row with { HeightRuleValue = "exact", DeclaredHeightPoints = row.Height }, outside, 0d, default),
            "Exact row heights have an intentional clipping policy.");
        TestAssert.True(!inspect(row with { HeightRuleValue = "unknown" }, outside, 0d, default), "Unknown height policy is not automatic.");
        var empty = row with { Cells = row.Cells.Select(cell => cell with {
            TextLines = cell.TextLines.Select(line => line with { Text = " " }).ToArray() }).ToArray() };
        TestAssert.True(!inspect(empty, outside, 0d, default), "Blank line frames do not demonstrate lost content.");
        var unknown = row with { Cells = row.Cells.Select(cell => cell with {
            TextLines = cell.TextLines.Select(line => line with { InlineFrame = null, OccupiedLineBox = null,
                BodyLineBoxHeightPoints = null, BodyLineBoxBaselineInsetPoints = null }).ToArray() }).ToArray() };
        TestAssert.True(!inspect(unknown, outside, 0d, default), "Missing measured bounds are not proof of clipping.");
        var clippedByCell = row with { Cells = row.Cells.Select(cell => cell with { Height = .001d }).ToArray() };
        TestAssert.True(!inspect(clippedByCell, outside, 0d, default), "Exclude text already outside its source cell clip.");
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(() =>
            inspect(row, outside, 0d, cancellation.Token)).CancellationToken);
    }

    public static void DocxNoteClippingObserversPreserveAtomicPublication()
    {
        string directory = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ooxpdf-note-clipping-" + Guid.NewGuid().ToString("N"))).FullName;
        try
        {
            using var source = Input(400, false);
            string inputPath = Path.Combine(directory, "source.docx"), targetPath = Path.Combine(directory, "target.pdf");
            File.WriteAllBytes(inputPath, source.ToArray());
            foreach (bool file in new[] { false, true })
            foreach (bool cancel in new[] { false, true })
            {
                File.WriteAllBytes(targetPath, [1, 2, 3]);
                using var input = new MemoryStream(source.ToArray()); using var output = new MemoryStream();
                using var cancellation = new CancellationTokenSource();
                var expected = new InvalidOperationException("Caller clipping observer failed.");
                bool observed = false;
                var options = Options(OoxPdfDocxMarkupMode.Final, diagnostic => {
                    if (diagnostic.Id != "DOCX_CLIPPED_NOTE_TABLE_ROW") return;
                    observed = true;
                    if (cancel) cancellation.Cancel(); else throw expected;
                });
                Task conversion = file ? OoxPdfConverter.ConvertAsync(inputPath, targetPath, options, cancellation.Token)
                    : OoxPdfConverter.ConvertAsync(input, output, options, cancellation.Token);
                if (cancel) TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(() =>
                    conversion.GetAwaiter().GetResult()).CancellationToken);
                else TestAssert.True(ReferenceEquals(expected, TestAssert.Throws<InvalidOperationException>(() =>
                    conversion.GetAwaiter().GetResult())), "Preserve caller exception identity.");
                TestAssert.True(observed, "Reach the actual clipping diagnostic before failure.");
                TestAssert.Equal(0L, output.Length);
                TestAssert.True(File.ReadAllBytes(targetPath).AsSpan().SequenceEqual(new byte[] { 1, 2, 3 }), "Retain prior file publication.");
                TestAssert.Equal(2, Directory.GetFiles(directory).Length);
            }
        }
        finally { Directory.Delete(directory, true); }
    }

    private static byte[] Convert(int tokens, bool exact, OoxPdfDocxMarkupMode mode, Action<OoxPdfDiagnostic>? sink)
    {
        using var input = Input(tokens, exact); using var output = new MemoryStream();
        OoxPdfConverter.Convert(input, output, Options(mode, sink));
        return output.ToArray();
    }

    private static OoxPdfOptions Options(OoxPdfDocxMarkupMode mode, Action<OoxPdfDiagnostic>? sink) => new() {
        InputKind = OoxPdfInputKind.Docx, FontResolver = new PortableRenderingTests.PinnedResolver(mathFont: true),
        DocxMarkupMode = mode, DocxMarkupGeometryMode = mode == OoxPdfDocxMarkupMode.Final
            ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup,
        FixedCreationDate = DateTimeOffset.UnixEpoch, DiagnosticSink = sink };

    private static MemoryStream BoundaryInput(bool math, bool overflow, bool exact)
    {
        using var seed = Input(0, false);
        using var archive = new ZipArchive(seed, ZipArchiveMode.Read);
        var parts = archive.Entries.ToDictionary(entry => entry.FullName, entry => {
            using var source = entry.Open(); using var copy = new MemoryStream(); source.CopyTo(copy); return copy.ToArray(); });
        using var xmlStream = new MemoryStream(parts["word/endnotes.xml"]);
        var xml = XDocument.Load(xmlStream);
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        var note = xml.Root!.Elements(w + "endnote").Single(note => (string?)note.Attribute(w + "id") == "37");
        var row = note.Element(w + "tbl")!.Element(w + "tr")!;
        var cell = row.Element(w + "tc")!;
        cell.Elements().Where(element => element.Name != w + "tcPr").Remove();
        var paragraph = new XElement(w + "p", new XElement(w + "pPr", new XElement(w + "spacing",
            new XAttribute(w + "before", "0"), new XAttribute(w + "after", "0"),
            new XAttribute(w + "line", "240"), new XAttribute(w + "lineRule", "auto"))));
        cell.Add(paragraph);
        if (math)
        {
            var matrix = new XElement(m + "m");
            for (int index = 0; index < (overflow ? 80 : 4); index++)
            {
                var run = new XElement(m + "r",
                    new XElement(m + "rPr", new XElement(m + "sty", new XAttribute(m + "val", "p"))),
                    new XElement(w + "rPr", new XElement(w + "rFonts", new XAttribute(w + "ascii", "Cambria Math"),
                        new XAttribute(w + "hAnsi", "Cambria Math"))), new XElement(m + "t", "A"));
                matrix.Add(new XElement(m + "mr", new XElement(m + "e", run)));
            }
            paragraph.Add(new XElement(m + "oMath", matrix));
        }
        else paragraph.Add(new XElement(w + "r", new XElement(w + "rPr",
            new XElement(w + "rFonts", new XAttribute(w + "ascii", "Arial"), new XAttribute(w + "hAnsi", "Arial")),
            new XElement(w + "sz", new XAttribute(w + "val", overflow ? "1200" : "24"))), new XElement(w + "t", "g")));
        if (exact)
        {
            var properties = row.Element(w + "trPr");
            if (properties is null) { properties = new XElement(w + "trPr"); row.AddFirst(properties); }
            properties.Elements(w + "trHeight").Remove();
            properties.Add(new XElement(w + "trHeight", new XAttribute(w + "val", "16000"), new XAttribute(w + "hRule", "exact")));
        }
        using var updated = new MemoryStream(); xml.Save(updated, SaveOptions.DisableFormatting);
        parts["word/endnotes.xml"] = updated.ToArray();
        return TestFixtures.CreateZipPackage(parts);
    }

    private static MemoryStream Input(int tokens, bool exact)
    {
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
            "CompositionProbes/mixed-note-endnote-table.docx"));
        using var archive = ZipFile.OpenRead(path);
        var parts = archive.Entries.ToDictionary(entry => entry.FullName, entry => {
            using var source = entry.Open(); using var copy = new MemoryStream(); source.CopyTo(copy); return copy.ToArray(); });
        using var xmlStream = new MemoryStream(parts["word/endnotes.xml"]); var xml = XDocument.Load(xmlStream);
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var note = xml.Root!.Elements(w + "endnote").Single(note => (string?)note.Attribute(w + "id") == "37");
        var row = note.Element(w + "tbl")!.Element(w + "tr")!;
        row.Descendants(w + "p").Single().Add(new XElement(w + "r", new XElement(w + "t",
            " " + string.Join(" ", Enumerable.Range(0, tokens).Select(i => $"PublicRow{i:000}")))));
        if (exact)
        {
            var properties = row.Element(w + "trPr");
            if (properties is null) { properties = new XElement(w + "trPr"); row.AddFirst(properties); }
            properties.Elements(w + "trHeight").Remove();
            properties.Add(new XElement(w + "trHeight", new XAttribute(w + "val", "16000"), new XAttribute(w + "hRule", "exact")));
        }
        using var updated = new MemoryStream(); xml.Save(updated, SaveOptions.DisableFormatting);
        parts["word/endnotes.xml"] = updated.ToArray();
        return TestFixtures.CreateZipPackage(parts);
    }
}
