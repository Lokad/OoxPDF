using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxFootnotesTests
{
    public static void DocxAutomaticNoteReferencesWithEqualLabelsKeepDistinctLinks()
    {
        foreach (bool superscript in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
            var renderer = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var sources = new[] { false, true }.Select(joined => ReadEqualLabelNotes(joined, superscript, mode)).ToArray();
            var references = sources[1].Paragraphs.Single(p => p.Images.Count != 0).InlineReferences;
            TestAssert.Equal(2, references.Count);
            TestAssert.True(references.All(reference => reference.DisplayText == "1") &&
                !ReferenceEquals(references[0].DisplayRun, references[1].DisplayRun),
                "Equal visible labels retain distinct generated source owners.");
            var pages = sources.Select(source => renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray()).ToArray();
            TestAssert.Equal(pages[0].Length, pages[1].Length);
            for (int page = 0; page < pages[0].Length; page++)
            {
                var a = DocxCommentsTests.ReadEmbeddedGlyphTextShows(pages[0][page], true).Select(g => (g.Text, g.X, g.Y, g.Size));
                var b = DocxCommentsTests.ReadEmbeddedGlyphTextShows(pages[1][page], true).Select(g => (g.Text, g.X, g.Y, g.Size));
                TestAssert.True(a.SequenceEqual(b), "Joining equal note labels retains visible content and placement.");
                TestAssert.True(pages[0][page].Annotations.SequenceEqual(pages[1][page].Annotations),
                    "Distinct automatic notes cannot share each other's link when their labels and source offsets coincide.");
            }
        }
    }

    private static DocxDocument ReadEqualLabelNotes(bool joined, bool superscript, OoxPdfDocxMarkupMode mode)
    {
        string input = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "Composition", "report.docx"));
        var parts = new Dictionary<string, byte[]>();
        using (var zip = ZipFile.OpenRead(input))
            foreach (var entry in zip.Entries)
            { using var stream = entry.Open(); using var bytes = new MemoryStream(); stream.CopyTo(bytes); parts.Add(entry.FullName, bytes.ToArray()); }
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XDocument Read(string name)
        { using var bytes = new MemoryStream(parts[name]); return XDocument.Load(bytes); }
        void Write(string name, XDocument xml) => parts[name] = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
        var document = Read("word/document.xml");
        var marker = document.Descendants(w + "footnoteReference").Single();
        var first = marker.Parent!;
        first.Remove();
        var picture = document.Descendants(w + "drawing").Single().Parent!;
        picture.AddBeforeSelf(first);
        var second = new XElement(w + "r", new XElement(w + "endnoteReference", new XAttribute(w + "id", "1")));
        if (superscript)
            foreach (var run in new[] { first, second }) run.AddFirst(new XElement(w + "rPr",
                new XElement(w + "vertAlign", new XAttribute(w + "val", "superscript"))));
        if (joined) first.Add(second.Elements().Where(child => child.Name != w + "rPr").Select(child => new XElement(child)));
        else first.AddAfterSelf(second);
        document.Descendants(w + "sectPr").Single().Add(new XElement(w + "endnotePr",
            new XElement(w + "numFmt", new XAttribute(w + "val", "decimal"))));
        Write("word/document.xml", document);
        var notes = Read("word/footnotes.xml");
        foreach (var node in notes.Descendants())
            if (node.Name.Namespace == w && node.Name.LocalName.StartsWith("footnote", StringComparison.Ordinal))
                node.Name = w + node.Name.LocalName.Replace("footnote", "endnote", StringComparison.Ordinal);
        Write("word/endnotes.xml", notes);
        var relationships = Read("word/_rels/document.xml.rels");
        relationships.Root!.Add(new XElement(relationships.Root.Name.Namespace + "Relationship", new XAttribute("Id", "rIdEqualEndnote"),
            new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"), new XAttribute("Target", "endnotes.xml")));
        Write("word/_rels/document.xml.rels", relationships);
        var types = Read("[Content_Types].xml");
        types.Root!.Add(new XElement(types.Root.Name.Namespace + "Override", new XAttribute("PartName", "/word/endnotes.xml"),
            new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml")));
        Write("[Content_Types].xml", types);
        string temporary = TestFixtures.WriteTempPackage(".docx", parts);
        try { return DocxTests.ReadDocx(temporary, mode); }
        finally { File.Delete(temporary); }
    }
}
