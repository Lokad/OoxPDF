using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

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
            var document = ReadMutation("docx-omml-normal-fraction-inline-arial-normal", mode, false, false, (xml, _, m) =>
            {
                XElement numerator = xml.Descendants(m + "num").Single();
                XElement moved = numerator.Element(m + "r")!;
                var unchanged = new XElement(moved);
                unchanged.Element(m + "t")!.Value = "2";
                numerator.ReplaceNodes(MathOperandMove(moved, "moveFrom", "41", "42"), unchanged,
                    MathOperandMove(moved, "moveTo", "43", "44"));
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

    public static void DocxMathRevisionBoundariesPairOperandsAcrossParagraphsAndPages()
    {
        foreach (bool pageBreak in new[] { false, true })
        foreach (var (mode, sourceRow, destinationRow) in new[]
        {
            (OoxPdfDocxMarkupMode.Final, "2time", "3ratecost"),
            (OoxPdfDocxMarkupMode.SimpleMarkup, "2time", "3ratecost"),
            (OoxPdfDocxMarkupMode.Original, "rate2time", "3cost"),
            (OoxPdfDocxMarkupMode.AllMarkup, "rate2time", "3ratecost")
        })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            var document = ReadMovedOperands(mode, pageBreak, diagnostics.Add, equivalentDecimalIds: pageBreak);
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.AllMarkup ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                    : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var pages = renderer.RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            string[] painted = pages.Select(page => string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page,
                usePageCoordinates: true, fontResourceFilter: name => name.StartsWith("Eq", StringComparison.Ordinal))
                .Select(show => show.Text))).ToArray();
            TestAssert.True(painted.SequenceEqual(pageBreak ? [sourceRow, destinationRow] : [sourceRow + destinationRow]),
                $"Both complete fractions follow the authored move and page ownership in {mode} (page break {pageBreak}).");
            TestAssert.Equal(0, diagnostics.Count(d => d.Id == "DOCX_UNSUPPORTED_TRACKED_CHANGES"));
            TestAssert.Equal(mode == OoxPdfDocxMarkupMode.Final ? 0 : 1,
                diagnostics.Count(d => d.Id == "DOCX_APPROXIMATED_TRACKED_CHANGES"));
            TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_EQUATION"),
                "The supported fractions still report their equation approximation.");
        }
    }

    public static void DocxMathRevisionBoundariesKeepMoveFallbackDiagnostics()
    {
        var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true),
            OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);

        Check(move => move[0].SetAttributeValue(WordprocessingNamespace + "name", "DifferentMove"),
            broadWarning: true, "DOCX_UNSUPPORTED_EQUATION");
        Check(move => move[2].AddAfterSelf(new XElement(move[2])),
            broadWarning: true, "DOCX_APPROXIMATED_EQUATION");
        Check(move => move[1].Add(new XElement(MathNamespace + "future")),
            broadWarning: false, "DOCX_UNSUPPORTED_EQUATION");

        void Check(Action<XElement[]> change, bool broadWarning, string equationDiagnostic)
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            var document = ReadMovedOperands(OoxPdfDocxMarkupMode.Final, pageBreak: false, diagnostics.Add, change);
            _ = renderer.RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().Single();
            TestAssert.Equal(broadWarning ? 1 : 0, diagnostics.Count(d => d.Id == "DOCX_UNSUPPORTED_TRACKED_CHANGES"));
            TestAssert.True(diagnostics.Any(d => d.Id == equationDiagnostic),
                $"Expected {equationDiagnostic} for the affected math move.");
        }
    }

    public static void DocxMathRevisionBoundariesKeepStoryAndUniqueMoveIds()
    {
        var run = new XElement(MathNamespace + "r", new XElement(MathNamespace + "t", "rate"));
        var from = new XElement(MathNamespace + "oMath", MathOperandMove(run, "moveFrom", "41", "42"));
        var to = new XElement(MathNamespace + "oMath", MathOperandMove(run, "moveTo", "43", "44"));
        var scope = new XElement(WordprocessingNamespace + "footnotes",
            new XElement(WordprocessingNamespace + "footnote", new XAttribute(WordprocessingNamespace + "id", "1"),
                new XElement(WordprocessingNamespace + "p", from)),
            new XElement(WordprocessingNamespace + "footnote", new XAttribute(WordprocessingNamespace + "id", "2"),
                new XElement(WordprocessingNamespace + "p", to)));
        TestAssert.Equal(2, scope.Elements().Count());
        TestAssert.Equal(0, DocxSourceXml.ReadPairedMoveOwnerIds(from, null, default).Count);
        TestAssert.Equal(0, DocxSourceXml.ReadPairedMoveOwnerIds(to, null, default).Count);

        var paired = new XElement(WordprocessingNamespace + "body",
            new XElement(WordprocessingNamespace + "p", new XElement(from), new XElement(to)));
        var unrelated = new XElement(WordprocessingNamespace + "moveTo", new XAttribute(WordprocessingNamespace + "id", "80"),
            new XElement(WordprocessingNamespace + "r", new XElement(WordprocessingNamespace + "t", "Independent")));
        paired.Add(new XElement(WordprocessingNamespace + "p", unrelated));
        TestAssert.Equal(2, DocxSourceXml.ReadPairedMoveOwnerIds(paired.Descendants(MathNamespace + "oMath").First(), null, default).Count);
        foreach (string duplicateId in new[] { "44", "+0044" })
        {
            var ambiguous = new XElement(paired);
            ambiguous.Elements().Last().Element(WordprocessingNamespace + "moveTo")!
                .SetAttributeValue(WordprocessingNamespace + "id", duplicateId);
            TestAssert.Equal(0, DocxSourceXml.ReadPairedMoveOwnerIds(ambiguous.Descendants(MathNamespace + "oMath").First(), null, default).Count);
        }
    }

    private static DocxDocument ReadMovedOperands(OoxPdfDocxMarkupMode mode, bool pageBreak,
        Action<OoxPdfDiagnostic> diagnostics, Action<XElement[]>? changeDestinationMove = null,
        bool equivalentDecimalIds = false) =>
        ReadMutation("docx-omml-normal-fraction-inline-arial-normal", mode, false, false, (xml, w, m) =>
        {
            XElement source = xml.Descendants(m + "oMath").Single().Parent!;
            var destination = new XElement(source);
            XElement firstNumerator = source.Descendants(m + "num").Single();
            XElement lastNumerator = destination.Descendants(m + "num").Single();
            var rate = new XElement(firstNumerator.Element(m + "r")!);
            firstNumerator.Element(m + "r")!.Element(m + "t")!.Value = "2";
            lastNumerator.Element(m + "r")!.Element(m + "t")!.Value = "3";
            destination.Descendants(m + "den").Single().Element(m + "r")!.Element(m + "t")!.Value = "cost";
            var sourceMove = MathOperandMove(rate, "moveFrom", "41", "42");
            firstNumerator.AddFirst(sourceMove);
            var destinationMove = MathOperandMove(rate, "moveTo", "43", "44");
            lastNumerator.Add(destinationMove);
            if (equivalentDecimalIds)
            {
                sourceMove[2].SetAttributeValue(w + "id", "+0041");
                destinationMove[1].SetAttributeValue(w + "id", "+0044");
                destinationMove[2].SetAttributeValue(w + "id", "00043");
            }
            changeDestinationMove?.Invoke(destinationMove);
            if (pageBreak)
            {
                source.Add(new XElement(w + "r", new XElement(w + "br", new XAttribute(w + "type", "page"))));
                source.Add(destination.Elements().Where(child => child.Name != w + "pPr"));
            }
            else source.AddAfterSelf(destination);
        }, diagnosticSink: diagnostics);

    private static XElement[] MathOperandMove(XElement operand, string kind, string rangeId, string revisionId)
    {
        XNamespace w = WordprocessingNamespace;
        return
        [
            new(w + (kind + "RangeStart"), new XAttribute(w + "id", rangeId), new XAttribute(w + "name", "FormulaMove"),
                new XAttribute(w + "author", "Editor"), new XAttribute(w + "date", "2024-01-01T00:00:00Z")),
            new(w + kind, new XAttribute(w + "id", revisionId), new XAttribute(w + "author", "Editor"),
                new XAttribute(w + "date", "2024-01-01T00:00:00Z"), new XElement(operand)),
            new(w + (kind + "RangeEnd"), new XAttribute(w + "id", rangeId))
        ];
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
