using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxMathRevisionBoundariesPreservePaintAndFiltering()
    {
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        foreach (string kind in new[] { "ins", "del" })
        foreach (bool control in new[] { false, true })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.AllMarkup ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                    : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var direct = renderer.RenderPagesAsync(ReadMathRevisionBoundary(mode, kind, control, "outer"), null, default)
                .ToBlockingEnumerable().ToArray();
            foreach (string marker in new[] { "bookmark", "proof", "permission" })
            {
                var document = ReadMathRevisionBoundary(mode, kind, control, marker);
                var marked = renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
                TestAssert.Equal(direct.Length, marked.Length);
                for (int page = 0; page < direct.Length; page++)
                {
                    TestAssert.Equal(direct[page].Content, marked[page].Content);
                    TestAssert.Equal(direct[page].Fonts.Count, marked[page].Fonts.Count);
                    TestAssert.Equal(direct[page].Images.Count, marked[page].Images.Count);
                }
                bool included = mode == OoxPdfDocxMarkupMode.AllMarkup ||
                    kind == "ins" && mode is OoxPdfDocxMarkupMode.Final or OoxPdfDocxMarkupMode.SimpleMarkup ||
                    kind == "del" && mode == OoxPdfDocxMarkupMode.Original;
                if (marker == "bookmark")
                    TestAssert.Equal(included ? 1 : 0,
                        document.Paragraphs.First().BookmarkAnchors.Count(anchor => anchor.Name == "RevisionTarget"));
                if (included)
                {
                    TestAssert.True(DocxMathLayout.TryGetRuns(document.Paragraphs.First().Equations.Single(), default,
                        out var runs) && runs.Any(run => run.Run.Text.Text == "x+2"),
                        "An included revision keeps its actual numerator run through source boundaries.");
                    TestAssert.True(marked.SelectMany(page => page.Annotations).Any(link => link.Destination?.PageIndex == 0),
                        "The included source bookmark retains its incoming internal link.");
                }
            }
        }
    }

    public static void DocxMathRevisionBoundariesKeepUnknownAndContentBearingFallbacks()
    {
        foreach (string marker in new[] { "foreign", "content", "move-content" })
        {
            var document = ReadMathRevisionBoundary(OoxPdfDocxMarkupMode.Final, "ins", false, marker);
            TestAssert.True(!DocxMathLayout.TryGetRuns(document.Paragraphs.First().Equations.Single(), default, out _),
                "Only exact empty source-boundary markers are transparent to the revision grammar.");
        }
        using var canceled = new CancellationTokenSource();
        canceled.Cancel();
        var equation = ReadMathRevisionBoundary(OoxPdfDocxMarkupMode.Final, "ins", true, "bookmark")
            .Paragraphs.First().Equations.Single();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryGetRuns(equation, canceled.Token, out _));
    }

    public static void DocxMathRevisionBoundariesPreserveMovedFractionContent()
    {
        foreach (var (mode, numeratorText, owners) in new[]
        {
            (OoxPdfDocxMarkupMode.Final, "2rate", new[] { DocxRevisionKind.MoveTo }),
            (OoxPdfDocxMarkupMode.SimpleMarkup, "2rate", new[] { DocxRevisionKind.MoveTo }),
            (OoxPdfDocxMarkupMode.Original, "rate2", new[] { DocxRevisionKind.MoveFrom }),
            (OoxPdfDocxMarkupMode.AllMarkup, "rate2rate", new[] { DocxRevisionKind.MoveFrom, DocxRevisionKind.MoveTo })
        })
        {
            var document = ReadMutation("docx-omml-normal-fraction-inline-arial-normal", mode, false, false, (xml, w, m) =>
            {
                XElement numerator = xml.Descendants(m + "num").Single();
                XElement moved = numerator.Element(m + "r")!;
                var unchanged = new XElement(moved);
                unchanged.Element(m + "t")!.Value = "2";
                numerator.ReplaceNodes(MoveRange("moveFrom", "41", "42"), unchanged, MoveRange("moveTo", "43", "44"));

                XElement[] MoveRange(string kind, string rangeId, string revisionId) =>
                [
                    new(w + (kind + "RangeStart"), new XAttribute(w + "id", rangeId), new XAttribute(w + "name", "FormulaMove"),
                        new XAttribute(w + "author", "Editor"), new XAttribute(w + "date", "2024-01-01T00:00:00Z")),
                    new(w + kind, new XAttribute(w + "id", revisionId), new XAttribute(w + "author", "Editor"),
                        new XAttribute(w + "date", "2024-01-01T00:00:00Z"), new XElement(moved)),
                    new(w + (kind + "RangeEnd"), new XAttribute(w + "id", rangeId))
                ];
            });
            var equation = document.Paragraphs.First().Equations.Single();
            TestAssert.True(DocxMathLayout.TryGetRuns(equation, default, out var runs), "A move inside a numerator preserves the supported fraction.");
            TestAssert.True(owners.Select(kind => (DocxRevisionKind?)kind).SequenceEqual(
                runs.Where(run => run.Run.Text.Text == "rate").Select(run => run.Run.Text.Revision?.Kind)),
                "The selected move operands retain their source and destination ownership.");

            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.AllMarkup ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                    : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var page = renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().Single();
            var equationPaint = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true,
                fontResourceFilter: name => name.StartsWith("Eq", StringComparison.Ordinal));
            TestAssert.Equal(numeratorText + "time", string.Concat(equationPaint.Select(show => show.Text)));
            string painted = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Select(show => show.Text));
            foreach (string label in new[] { "Before formula:", "after formula.", "Following prose." })
                TestAssert.Contains(label, painted);
        }
    }

    private static DocxDocument ReadMathRevisionBoundary(OoxPdfDocxMarkupMode mode, string kind,
        bool control, string marker, string? contentControl = null, string carrier = "sdt") => ReadMutation("CompositionProbes/math-run-revision-body-ins", mode, false, false,
        (xml, w, m) =>
        {
            var revision = xml.Descendants(w + "ins").Single();
            revision.Name = w + kind;
            if (control)
            {
                var fraction = revision.Ancestors(m + "f").Single();
                revision.ReplaceWith(revision.Elements().Select(element => new XElement(element)).ToArray());
                var whole = new XElement(w + kind, revision.Attributes(), new XElement(fraction));
                fraction.ReplaceWith(whole);
                revision = whole;
            }
            if (contentControl == "inside")
            {
                var children = revision.Elements().ToArray();
                revision.RemoveNodes();
                revision.Add(MathSourceContainer(children, w, carrier));
            }
            else if (contentControl == "outside")
            {
                var metadata = MathSourceContainer([new XElement(revision)], w, carrier);
                revision.ReplaceWith(metadata);
                revision = metadata.Descendants(w + kind).Single();
            }
            var owner = marker == "bookmark" ? revision : revision.Parent!;
            owner.AddFirst(new XElement(w + "bookmarkStart", new XAttribute(w + "id", "19114"),
                new XAttribute(w + "name", "RevisionTarget")));
            owner.Add(new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "19114")));
            if (marker == "proof")
            {
                revision.AddFirst(new XElement(w + "proofErr", new XAttribute(w + "type", "spellStart")));
                revision.Add(new XElement(w + "proofErr", new XAttribute(w + "type", "spellEnd")));
            }
            else if (marker == "permission")
            {
                revision.AddFirst(new XElement(w + "permStart", new XAttribute(w + "id", "19115"),
                    new XAttribute(w + "edGrp", "everyone")));
                revision.Add(new XElement(w + "permEnd", new XAttribute(w + "id", "19115")));
            }
            else if (marker == "foreign") revision.AddFirst(new XElement(XName.Get("proofErr", "urn:future")));
            else if (marker == "content") revision.AddFirst(new XElement(w + "proofErr", "Visible"));
            else if (marker == "move-content") revision.AddFirst(new XElement(w + "moveToRangeStart", "Visible"));
            XNamespace relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            var link = xml.Descendants(w + "hyperlink").Single(item => !item.Ancestors(m + "oMath").Any());
            link.Attribute(relationships + "id")!.Remove();
            link.SetAttributeValue(w + "anchor", "RevisionTarget");
        });
}
