using System.Collections;
using System.Reflection;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxStructuralMetadataWrappersKeepResolvedParentFrames()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Parent structural frame comparison uses installed Windows fonts.");
        CheckMetadataParentFrames(null, "structural-metadata-wrappers.json");
    }

    public static void DocxStructuralMetadataWrappersMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent structural metadata geometry uses installed Windows fonts.");
        const string golden = "structural-metadata-wrappers-office.json";
        CheckInlineBoxReference(golden, InlineBoxChecks.All);
    }

    public static void DocxStructuralMetadataWrappersKeepNestedRevisionAndSourceOwners()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupMode.Original })
        {
            var document = ReadMutation("Composition/report", mode, false, false, (xml, w, _) =>
            {
                XElement Paragraph(string text) => new(w + "p", new XElement(w + "r", new XElement(w + "t", text)));
                XElement Revised(string kind, string id, string text) => new(w + kind,
                    new XAttribute(w + "id", id), new XAttribute(w + "author", "Public"),
                    new XAttribute(w + "date", "2026-01-01T00:00:00Z"), MetadataWrapper(w, "customXml", Paragraph(text)));
                var body = xml.Root!.Element(w + "body")!;
                var section = new XElement(body.Element(w + "sectPr")!);
                body.ReplaceNodes(MetadataWrapper(w, "customXml",
                    new XElement(w + "customXmlPr", new XElement(w + "attr", new XAttribute(w + "name", "scope"),
                        new XAttribute(w + "val", "public"))),
                    Paragraph("Visible"), Revised("ins", "511", "Inserted"), Revised("del", "512", "Deleted"),
                    new XElement(XName.Get("customXml", "urn:future"), Paragraph("Unqualified"))), section);
            });
            string expected = mode switch
            {
                OoxPdfDocxMarkupMode.Final => "VisibleInserted",
                OoxPdfDocxMarkupMode.Original => "VisibleDeleted",
                _ => "VisibleInsertedDeleted"
            };
            TestAssert.Equal(expected, string.Concat(document.Paragraphs.SelectMany(p => p.Runs).Select(run => run.Text)));
            TestAssert.True(document.Paragraphs.First().Revisions.Count == 0,
                "Metadata does not create a revision owner for ordinary source text.");
            foreach (var paragraph in document.Paragraphs.Skip(1))
                TestAssert.True(paragraph.Revisions.Count == 1 && paragraph.Revisions[0].Author == "Public" &&
                    paragraph.Revisions[0].Kind is DocxRevisionKind.Insertion or DocxRevisionKind.Deletion,
                    "Nested metadata preserves the actual inherited revision instead of inventing another owner.");
        }
    }

    public static void DocxStructuralMetadataWrappersObserveCancellationWithoutYieldingOwners()
    {
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        // Exercise the live deferred visitor without publishing a test-only API.
        var visit = typeof(DocxReader).GetMethod("EnumerateRevisionScopedChildren", BindingFlags.Static | BindingFlags.NonPublic)!;
        foreach ((bool inlineContent, bool nestedFieldResults) in new[] { (false, false), (true, false), (true, true) })
        {
            XName owner = w + (inlineContent ? "r" : "p");
            var metadata = MetadataWrapper(w, "customXml",
                MetadataWrapper(w, "customXml", new XElement(w + "customXmlPr")));
            if (nestedFieldResults)
                metadata = new XElement(w + "fldSimple", new XAttribute(w + "instr", "REF Cache"), metadata);
            var source = new[] { new XElement(owner), metadata, new XElement(owner) };
            using var cancelled = new CancellationTokenSource();
            var sequence = (IEnumerable)visit.Invoke(null,
                [source, OoxPdfDocxMarkupMode.Final, cancelled.Token, new XName[] { owner }, null, inlineContent, nestedFieldResults])!;
            var iterator = sequence.GetEnumerator();
            try
            {
                TestAssert.True(iterator.MoveNext(), "The first actual source owner is yielded normally.");
                cancelled.Cancel();
                TestAssert.Throws<OperationCanceledException>(() => iterator.MoveNext());
            }
            finally { (iterator as IDisposable)?.Dispose(); }
        }
    }
}
