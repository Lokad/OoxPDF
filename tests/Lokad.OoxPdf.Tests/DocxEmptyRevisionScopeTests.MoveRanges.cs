using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxEmptyRevisionScopeTests
{
    public static void DocxEmptyMoveRangesPreserveMixedContentAcrossStories()
    {
        var sources = Cases.Concat(new (string Fixture, string Part)[]
        {
            ("CompositionProbes/math-note-content-footnote-inline", "word/footnotes.xml"),
            ("CompositionProbes/math-note-content-endnote-inline", "word/endnotes.xml")
        });
        foreach (var source in sources)
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.AllMarkup ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup :
                    OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var expected = renderer.RenderPagesAsync(Read(source, mode, null), null, default).ToBlockingEnumerable().ToArray();
            foreach (string kind in new[] { "moveFrom", "moveTo" })
            foreach (string payload in new[] { "empty", "bookmark" })
            foreach (string carrier in new[] { "direct", "hyperlink", "sdt", "customXml", "field", "split-revision" })
            {
                var document = ReadMoveRange(source, mode, kind, payload, carrier);
                var diagnostics = new List<OoxPdfDiagnostic>();
                var actual = renderer.RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
                AssertPaint(expected, actual, source.Fixture + "/" + kind + "/" + payload + "/" + carrier + "/" + mode);
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                    "An independently closed payload-free range cannot suppress supported neighboring mathematics.");
            }
        }
    }

    public static void DocxMoveRangesRetainUnknownContentOpenAndLegacyFallbacks()
    {
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        foreach (string kind in new[] { "moveFrom", "moveTo" })
        foreach (string payload in new[] { "empty", "bookmark", "literal", "empty-run", "unknown", "field-unknown", "equation", "moved",
            "computed", "locked-empty", "malformed-start", "malformed-end", "marker-text", "marker-attribute",
            "missing-id", "mismatched-id", "unclosed", "orphan" })
        {
            var document = ReadMoveRange(Cases[0], mode, kind, payload);
            var paragraph = DocxBlockTraversal.EnumerateBodyParagraphs(document).First(p => p.Equations.Count > 0);
            bool empty = payload is "empty" or "bookmark" or "locked-empty";
            TestAssert.Equal(empty ? paragraph.Equations.Count : 0,
                DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
            TestAssert.Equal(empty ? paragraph.Equations.Count : 0,
                DocxMathParagraphPolicy.GetPreparationEquations(paragraph, hasPageInputs: true).Count);
            TestAssert.True(paragraph.RevisionRanges.Count > 0,
                "Range provenance remains available independently of paint eligibility.");
            if (empty)
            {
                var range = paragraph.RevisionRanges.Single();
                TestAssert.True(range.Id == "961" && range.Name == "EmptyMoveRange" &&
                    range.Author == "Public reviewer" && range.Date == "2026-11-13T00:00:00Z" &&
                    range.StartSourceRunIndex == range.EndSourceRunIndex && range.StartTextOffset == range.EndTextOffset,
                    "Known empty source ranges retain their original named endpoints and coordinates.");
                TestAssert.Equal(0, paragraph.Revisions.Count);
            }
        }
        var baseline = Read(Cases[0], OoxPdfDocxMarkupMode.Final, null);
        var original = DocxBlockTraversal.EnumerateBodyParagraphs(baseline).First(p => p.Equations.Count > 0);
        var unproven = original with
        {
            RevisionRanges = [new(DocxRevisionKind.MoveFrom, "961", "EmptyMoveRange", "Public reviewer", null, 0, 0, 0, 0)]
        };
        TestAssert.Equal(0, DocxMathParagraphPolicy.GetPreparationEquations(unproven).Count);
        var recorded = ReadMoveRange(Cases[0], OoxPdfDocxMarkupMode.Final, "moveFrom", "empty");
        var structure = new DocxRenderer().InspectStructure(recorded);
        TestAssert.Equal(1, structure.RevisionRanges!.Count);
        TestAssert.True(structure.RevisionRanges[0].IsClosed && structure.RevisionRanges[0].HasName &&
            structure.RevisionRanges[0].HasAuthor && structure.RevisionRanges[0].HasDate,
            "Public-safe structure inspection retains closed range metadata even when it owns no ink.");
    }

    private static DocxDocument ReadMoveRange((string Fixture, string Part) source, OoxPdfDocxMarkupMode mode,
        string kind, string payload, string carrier = "direct") =>
        DocxInlineCompositionTests.ReadMutation(source.Fixture, mode, false, false, (xml, w, math) =>
        {
            var equation = xml.Descendants(math + "oMath").First();
            var start = new XElement(w + (kind + "RangeStart"), new XAttribute(w + "id", "961"),
                new XAttribute(w + "name", "EmptyMoveRange"), new XAttribute(w + "author", "Public reviewer"),
                new XAttribute(w + "date", "2026-11-13T00:00:00Z"));
            var end = new XElement(w + (kind + "RangeEnd"), new XAttribute(w + "id", "961"));
            XElement[] content = payload switch
            {
                "bookmark" => [new XElement(w + "bookmarkStart", new XAttribute(w + "id", "962"),
                    new XAttribute(w + "name", "EmptyMoveAnchor")), new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "962"))],
                "literal" => [new XElement(w + "r", new XElement(w + "t", "Retained source"))],
                "empty-run" => [new XElement(w + "r")],
                "unknown" => [new XElement(w + "future")],
                "equation" => [new XElement(equation)],
                "moved" => [new XElement(w + "moveFrom", new XAttribute(w + "id", "964"),
                    new XElement(w + "r", new XElement(w + "delText", "Retained moving source")))],
                "field-unknown" => [new XElement(w + "fldSimple", new XAttribute(w + "instr", "REF UnknownPayload"), new XElement(w + "future"))],
                "computed" => [new XElement(w + "fldSimple", new XAttribute(w + "instr", "PAGE"))],
                "locked-empty" => [new XElement(w + "fldSimple", new XAttribute(w + "instr", "PAGE"), new XAttribute(w + "fldLock", "true"))],
                _ => []
            };
            if (payload == "malformed-start") start.Add(new XElement(w + "future"));
            if (payload == "malformed-end") end.Add(new XElement(w + "future"));
            if (payload == "marker-text") start.Value = "Unknown source";
            if (payload == "marker-attribute") start.SetAttributeValue(w + "future", "unknown");
            if (payload == "missing-id") { start.Attribute(w + "id")!.Remove(); end.Attribute(w + "id")!.Remove(); }
            if (payload == "mismatched-id") end.SetAttributeValue(w + "id", "999");
            var children = new List<XElement>();
            if (payload != "orphan") children.Add(start);
            children.AddRange(content);
            if (payload != "unclosed") children.Add(end);
            object result = carrier switch
            {
                "hyperlink" => new XElement(w + "hyperlink", new XAttribute(w + "anchor", "EmptyMoveAnchor"), children),
                "sdt" => new XElement(w + "sdt", new XElement(w + "sdtContent", children)),
                "customXml" => new XElement(w + "customXml", children),
                "field" => new XElement(w + "fldSimple", new XAttribute(w + "instr", "REF EmptyMoveField"), children),
                "split-revision" => new object[] { start, new XElement(w + "ins", new XAttribute(w + "id", "963"), content, end) },
                _ => children
            };
            equation.AddAfterSelf(result);
        }, partOverride: source.Part);
}
