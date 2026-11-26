using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PortableRenderingTests
{
    public static void PortableRenderingPackedMathRunRevisionsKeepEquivalentPaint()
    {
        foreach (var (story, part) in new[] { ("body", "word/document.xml"), ("header", "word/header1.xml") })
        foreach (string kind in new[] { "ins", "del" })
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            var beforeDiagnostics = new List<OoxPdfDiagnostic>();
            var before = Render(packed: false, beforeDiagnostics);
            string expected = RevisionBodyEquationText(kind, mode);
            TestAssert.Equal(expected, ReadPortableEquationGlyphs(before));
            CheckPackedPaint();
            if ((story, kind) is ("body", "ins") or ("header", "del"))
            {
                CheckPackedPaint(outerMathProperties: true);
                CheckPackedPaint(outerWordProperties: true);
                CheckPackedPaint(outerMathProperties: true, outerWordProperties: true);
                CheckPackedPaint(outerMathProperties: true, outerWordProperties: true, addMarkers: (revision, w) =>
                {
                    revision.AddBeforeSelf(new XElement(w + "tab"));
                    revision.AddAfterSelf(new XElement(w + "lastRenderedPageBreak"));
                });
                CheckPackedPaint(outerMathProperties: true, outerWordProperties: true, addMarkers: (revision, w) =>
                {
                    revision.AddFirst(new XElement(w + "tab"));
                    revision.Add(new XElement(w + "lastRenderedPageBreak"));
                });
            }

            void CheckPackedPaint(bool outerMathProperties = false, bool outerWordProperties = false,
                Action<XElement, XNamespace>? addMarkers = null)
            {
                var diagnostics = new List<OoxPdfDiagnostic>();
                var page = Render(packed: true, diagnostics, outerMathProperties, outerWordProperties, addMarkers);
                TestAssert.Equal(expected, ReadPortableEquationGlyphs(page));
                TestAssert.Equal(before.Content, page.Content);
                TestAssert.True(before.Annotations.SequenceEqual(page.Annotations),
                    "Moving the revision boundary keeps the same navigation.");
                TestAssert.True(beforeDiagnostics.SequenceEqual(diagnostics),
                    "Moving the revision boundary keeps the same equation and review diagnostics.");
                PdfContentValidator.ValidatePage(page, 0, default);
            }

            PdfPage Render(bool packed, List<OoxPdfDiagnostic> diagnostics,
                bool outerMathProperties = false, bool outerWordProperties = false, Action<XElement, XNamespace>? addMarkers = null)
            {
                var document = DocxInlineCompositionTests.ReadMutation($"CompositionProbes/math-run-revision-{story}-{kind}",
                    mode, false, false, (xml, w, m) =>
                    {
                        XElement revision = xml.Descendants(w + kind).Single();
                        XElement run = revision.Element(m + "r")!;
                        var mathProperties = new XElement(m + "rPr", new XElement(m + "sty", new XAttribute(m + "val", "p")));
                        var wordProperties = new XElement(w + "rPr", new XElement(w + "color", new XAttribute(w + "val", "1840A0")));
                        run.AddFirst(mathProperties, wordProperties);
                        if (!packed) return;

                        if (kind == "del") run.Element(m + "t")!.Name = w + "delText";
                        var content = run.Nodes().ToArray();
                        run.RemoveNodes();
                        run.Remove();
                        revision.ReplaceWith(run);
                        revision.Add(content);
                        run.Add(revision);
                        if (outerMathProperties)
                        {
                            mathProperties.Remove();
                            revision.AddBeforeSelf(mathProperties);
                        }
                        if (outerWordProperties)
                        {
                            wordProperties.Remove();
                            revision.AddBeforeSelf(wordProperties);
                        }
                        addMarkers?.Invoke(revision, w);
                    }, partOverride: part, diagnosticSink: diagnostics.Add);
                var renderer = new DocxRenderer(new PinnedResolver(mathFont: true), mode,
                    mode == OoxPdfDocxMarkupMode.AllMarkup ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup :
                        OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
                return renderer.RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().Single();
            }
        }
    }

    public static void PortableRenderingDeletedMathTextRetainsRevisionPaint()
    {
        foreach (string kind in new[] { "math-run-revision", "equation-revision" })
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            string fixture = $"CompositionProbes/{kind}-body-del";
            var beforeDiagnostics = new List<OoxPdfDiagnostic>();
            var afterDiagnostics = new List<OoxPdfDiagnostic>();
            var baseline = DocxInlineCompositionTests.ReadMutation(fixture, mode, false, false,
                diagnosticSink: beforeDiagnostics.Add);
            var changed = DocxInlineCompositionTests.ReadMutation(fixture, mode, false, false, (xml, w, m) =>
            {
                var deletedText = xml.Descendants(w + "del").Single().Descendants(m + "t").ToArray();
                TestAssert.Equal(kind == "math-run-revision" ? 1 : 2, deletedText.Length);
                foreach (var literal in deletedText) literal.Name = w + "delText";
            }, diagnosticSink: afterDiagnostics.Add);
            var renderer = new DocxRenderer(new PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.AllMarkup ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup :
                    OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var before = renderer.RenderPagesAsync(baseline, beforeDiagnostics.Add, default).ToBlockingEnumerable().ToArray();
            var after = renderer.RenderPagesAsync(changed, afterDiagnostics.Add, default).ToBlockingEnumerable().ToArray();
            string expected = RevisionBodyEquationText("del", mode);
            if (kind == "equation-revision" && mode is OoxPdfDocxMarkupMode.Final or OoxPdfDocxMarkupMode.SimpleMarkup)
                expected = expected[1..]; // The entire first fraction, including its denominator, is deleted.
            TestAssert.Equal(expected, string.Concat(after.Select(ReadPortableEquationGlyphs)));
            TestAssert.Equal(before.Length, after.Length);
            for (int page = 0; page < before.Length; page++)
            {
                TestAssert.Equal(before[page].Content, after[page].Content);
                TestAssert.True(before[page].Annotations.SequenceEqual(after[page].Annotations),
                    "Deleted literal text preserves the selected view's navigation.");
            }
            TestAssert.True(beforeDiagnostics.SequenceEqual(afterDiagnostics),
                "Deleted literal text preserves the selected view's equation diagnostics.");
            CheckWrittenPages(after);
        }
    }

    public static void PortableRenderingEmptyRevisedMathRunsKeepFractions()
    {
        foreach (string kind in new[] { "ins", "del" })
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            string fixture = "CompositionProbes/math-run-revision-body-" + kind;
            var beforeSource = DocxInlineCompositionTests.ReadMutation(fixture, mode, false, false);
            var renderer = new DocxRenderer(new PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.AllMarkup ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup :
                    OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var beforeDiagnostics = new List<OoxPdfDiagnostic>();
            var before = renderer.RenderPagesAsync(beforeSource, beforeDiagnostics.Add, default).ToBlockingEnumerable().ToArray();
            string[] beforeGlyphs = before.Select(ReadPortableEquationGlyphs).ToArray();
            TestAssert.Equal(RevisionBodyEquationText(kind, mode), string.Concat(beforeGlyphs));

            foreach (string contents in new[] { "", "<w:lastRenderedPageBreak/>", "<m:rPr><m:nor m:val='0'/></m:rPr>" })
            {
                var afterSource = DocxInlineCompositionTests.ReadMutation(fixture, mode, false, false, (xml, w, m) =>
                {
                    var revision = xml.Descendants(w + kind).Single();
                    revision.Add(XElement.Parse($"<m:r xmlns:m='{m}' xmlns:w='{w}'>{contents}</m:r>"));
                });
                var afterDiagnostics = new List<OoxPdfDiagnostic>();
                var after = renderer.RenderPagesAsync(afterSource, afterDiagnostics.Add, default).ToBlockingEnumerable().ToArray();
                TestAssert.Equal(before.Length, after.Length);
                for (int page = 0; page < before.Length; page++)
                {
                    TestAssert.Equal(beforeGlyphs[page], ReadPortableEquationGlyphs(after[page]));
                    TestAssert.Equal(before[page].Content, after[page].Content);
                    TestAssert.Equal(before[page].Fonts.Count, after[page].Fonts.Count);
                    TestAssert.Equal(before[page].Images.Count, after[page].Images.Count);
                    TestAssert.Equal(before[page].Annotations.Count, after[page].Annotations.Count);
                    PdfContentValidator.ValidatePage(after[page], page, default);
                }
                TestAssert.True(beforeDiagnostics.SequenceEqual(afterDiagnostics),
                    "An empty readable run cannot change the revision's supported visible mathematics.");
            }
        }
    }

    public static void PortableRenderingMathSmartTagsPreserveRevisedFractions()
    {
        foreach (string kind in new[] { "ins", "del" })
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            string fixture = "CompositionProbes/math-run-revision-body-" + kind;
            var beforeNotes = new List<OoxPdfDiagnostic>();
            var afterNotes = new List<OoxPdfDiagnostic>();
            var direct = DocxInlineCompositionTests.ReadMutation(fixture, mode, false, false, diagnosticSink: beforeNotes.Add);
            var wrapped = DocxInlineCompositionTests.ReadMutation(fixture, mode, false, false, (xml, w, _) =>
            {
                XElement revision = xml.Descendants(w + kind).Single();
                revision.ReplaceWith(new XElement(w + "smartTag", new XAttribute(w + "element", "operand"),
                    new XElement(w + "smartTagPr", new XElement(w + "attr",
                        new XAttribute(w + "name", "role"), new XAttribute(w + "val", "numerator"))), revision));
            }, diagnosticSink: afterNotes.Add);
            var renderer = new DocxRenderer(new PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.AllMarkup ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup :
                    OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var expected = renderer.RenderPagesAsync(direct, beforeNotes.Add, default).ToBlockingEnumerable().Single();
            var measured = renderer.RenderPagesAsync(wrapped, afterNotes.Add, default).ToBlockingEnumerable().Single();
            TestAssert.Equal(RevisionBodyEquationText(kind, mode), ReadPortableEquationGlyphs(measured));
            TestAssert.Equal(expected.Content, measured.Content);
            TestAssert.True(expected.Annotations.Count > 0 && expected.Annotations.SequenceEqual(measured.Annotations),
                "A math metadata wrapper must preserve the surrounding link.");
            TestAssert.True(beforeNotes.SequenceEqual(afterNotes), "The revision keeps its selected view's diagnostics.");
            CheckWrittenPages([measured]);
        }
    }

    private static string RevisionBodyEquationText(string kind, OoxPdfDocxMarkupMode mode)
    {
        // The first fraction has a revised numerator. The other equations are unmarked.
        bool numeratorVisible = kind == "ins" ? mode != OoxPdfDocxMarkupMode.Original :
            mode is OoxPdfDocxMarkupMode.Original or OoxPdfDocxMarkupMode.AllMarkup;
        return (numeratorVisible ? "x+2" : "") + "ya+b=cx+2ya+b=c";
    }

    private static string ReadPortableEquationGlyphs(PdfPage page) => string.Concat(
        DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true, honorActualText: false,
            fontResourceFilter: name => name.StartsWith("Eq", StringComparison.Ordinal)).Select(show => show.Text))
        .Normalize(NormalizationForm.FormKC);

    public static void PortableRenderingMathRunWordContentKeepsComposedFractions() =>
        CheckComposedMathRunContent("x\u20112", (numerators, w, m) =>
        {
            numerators[0].ReplaceWith(new XElement(w + "t", "x\u20112"));
            numerators[1].ReplaceWith(new XElement(w + "t", "x"), new XElement(w + "noBreakHyphen"), new XElement(w + "t", "2"));
            numerators[2].ReplaceWith(new XElement(w + "t", "x"), new XElement(w + "noBreakHyphen"), new XElement(m + "t", "2"));
        });

    public static void PortableRenderingMathDiscretionaryHyphensKeepUnbrokenFractions() =>
        CheckComposedMathRunContent("x2", (numerators, w, m) =>
        {
            numerators[0].Value = "x\u00AD2";
            numerators[1].ReplaceWith(new XElement(w + "t", "x"), new XElement(w + "softHyphen"), new XElement(m + "t", "2"));
            numerators[2].ReplaceWith(new XElement(w + "t", "x\u00AD2"));
        });

    public static void PortableRenderingMathZeroWidthBreaksKeepUnbrokenFractions() =>
        CheckComposedMathRunContent("x2", (numerators, w, m) =>
        {
            numerators[0].Value = "x\u200B2";
            numerators[1].ReplaceWith(new XElement(w + "t", "x\u200B2"));
            numerators[2].Value = "x";
            numerators[2].Parent!.AddAfterSelf(new XElement(m + "r", new XElement(m + "t", "\u200B")),
                new XElement(m + "r", new XElement(m + "t", "2")));
        });

    public static void PortableRenderingMathWordJoinersKeepUnbrokenFractions() =>
        CheckComposedMathRunContent("x+2", (numerators, w, m) =>
        {
            numerators[0].Value = "x+\u20602";
            numerators[1].ReplaceWith(new XElement(w + "t", "x+\uFEFF2"));
            numerators[2].Value = "x+";
            numerators[2].Parent!.AddAfterSelf(new XElement(m + "r", new XElement(m + "t", "\u2060")),
                new XElement(m + "r", new XElement(m + "t", "2")));
        });

    public static void PortableRenderingMathEmptyRunFormattingKeepsComposedFractions() =>
        CheckComposedMathRunContent("x2", (numerators, w, m) =>
        {
            foreach (var numerator in numerators) numerator.Value = "x2";
            numerators[0].Parent!.AddAfterSelf(new XElement(m + "r",
                new XElement(w + "rPr", new XElement(w + "vanish")), new XElement(m + "t")));
            numerators[1].Parent!.AddAfterSelf(new XElement(m + "r",
                new XElement(w + "rPr", new XElement(w + "highlight", new XAttribute(w + "val", "yellow"))),
                new XElement(m + "t", "\u200B")));
            numerators[2].Parent!.AddAfterSelf(new XElement(m + "r", new XElement(m + "t")));
        });

    public static void PortableRenderingDisabledMathAlignmentKeepsComposedFractions() =>
        CheckComposedMathRunContent("x+2", (numerators, _, m) =>
        {
            for (int index = 0; index < numerators.Length; index++)
            {
                XElement run = numerators[index].Parent!;
                XElement properties = run.Element(m + "rPr") ?? new XElement(m + "rPr");
                if (properties.Parent is null) run.AddFirst(properties);
                properties.Add(new XElement(m + "aln", new XAttribute(m + "val", index == 1 ? "false" : "0")));
            }
        });

    public static void PortableRenderingLiteralMathRunBreakKeepsComposedFractions() =>
        CheckComposedMathRunContent("x+2", (numerators, _, m) =>
        {
            for (int index = 0; index < numerators.Length; index++)
            {
                XElement run = numerators[index].Parent!;
                XElement properties = run.Element(m + "rPr") ?? new XElement(m + "rPr");
                if (properties.Parent is null) run.AddFirst(properties);
                properties.Add(new XElement(m + "brk", index == 1 ? new XAttribute(m + "alnAt", "1") : null));
            }
        });

    public static void PortableRenderingNeutralMathArgumentControlsKeepComposedFractions() =>
        CheckComposedMathRunContent("x+2", (numerators, w, m) =>
        {
            XElement[] arguments = numerators.Select(numerator => numerator.Parent!.Parent!).ToArray();
            arguments[0].Add(new XElement(m + "ctrlPr"));
            arguments[1].Add(new XElement(m + "ctrlPr",
                new XElement(w + "rPr", new XElement(w + "noProof"))));
            arguments[2].AddFirst(new XElement(m + "argPr",
                new XElement(m + "argSz", new XAttribute(m + "val", "0"))));
            arguments[2].Add(new XElement(m + "ctrlPr",
                new XElement(w + "rPr", new XElement(w + "webHidden", new XAttribute(w + "val", "1")))));
        });

    public static void PortableRenderingVisibleTransparentPhantomsKeepComposedFractions() =>
        CheckComposedMathRunContent("x+2", (numerators, w, m) =>
        {
            for (int index = 0; index < numerators.Length; index++)
            {
                XElement run = numerators[index].Parent!;
                var transparency = new XElement(m + "transp");
                if (index == 1) transparency.SetAttributeValue(m + "val", "true");
                var properties = new XElement(m + "phantPr", transparency);
                if (index == 1)
                    properties.AddFirst(new XElement(m + "show", new XAttribute(m + "val", "1")));
                if (index == 2)
                {
                    transparency.AddBeforeSelf(new XElement(m + "zeroWid", new XAttribute(m + "val", "0")),
                        new XElement(m + "zeroAsc", new XAttribute(m + "val", "false")),
                        new XElement(m + "zeroDesc", new XAttribute(m + "val", "0")));
                    properties.Add(new XElement(m + "ctrlPr", new XElement(w + "rPr", new XElement(w + "noProof"))));
                }
                run.ReplaceWith(new XElement(m + "phant", properties, new XElement(m + "e", new XElement(run))));
            }
        });

    private static void CheckComposedMathRunContent(string numerator, Action<XElement[], XNamespace, XNamespace> rewrite)
    {
        const string fixture = "CompositionProbes/composed-static-svg-report";
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var beforeDiagnostics = new List<OoxPdfDiagnostic>();
            var afterDiagnostics = new List<OoxPdfDiagnostic>();
            var beforeSource = DocxInlineCompositionTests.ReadMutation(fixture, mode, false, false, (xml, _, m) =>
            {
                foreach (var literal in ReadNumerators(xml, m)) literal.Value = numerator;
            },
                diagnosticSink: beforeDiagnostics.Add);
            var afterSource = DocxInlineCompositionTests.ReadMutation(fixture, mode, false, false,
                (xml, w, m) => rewrite(ReadNumerators(xml, m), w, m), diagnosticSink: afterDiagnostics.Add);
            var renderer = new DocxRenderer(new PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var before = renderer.RenderPagesAsync(beforeSource, beforeDiagnostics.Add, default).ToBlockingEnumerable().ToArray();
            var after = renderer.RenderPagesAsync(afterSource, afterDiagnostics.Add, default).ToBlockingEnumerable().ToArray();
            var shows = after.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page,
                usePageCoordinates: true, honorActualText: false,
                fontResourceFilter: name => name.StartsWith("Eq", StringComparison.Ordinal)));
            string text = string.Concat(shows.Select(show => show.Text));
            TestAssert.Equal(3 * numerator.Count(character => character == '\u2011'), text.Count(character => character == '\u2011'));
            TestAssert.True(!text.Any(c => OoxTextBreakRules.IsNonPrintingBreakControl(c)),
                "Nonprinting break controls must stay out of the painted glyph stream.");
            string expected = numerator.Normalize(NormalizationForm.FormKC) + "y";
            TestAssert.Equal(3, text.Normalize(NormalizationForm.FormKC).Split(expected, StringSplitOptions.None).Length - 1);
            TestAssert.Equal(before.Length, after.Length);
            for (int page = 0; page < before.Length; page++)
            {
                TestAssert.Equal(before[page].Content, after[page].Content);
                TestAssert.True(before[page].Annotations.SequenceEqual(after[page].Annotations),
                    "WordprocessingML run content in fractions preserves navigation.");
            }
            TestAssert.True(beforeDiagnostics.SequenceEqual(afterDiagnostics), "WordprocessingML run content preserves equation diagnostics.");
            CheckWrittenPages(after);
        }

        static XElement[] ReadNumerators(XDocument xml, XNamespace math)
        {
            var numerators = xml.Descendants(math + "f").Select(fraction =>
                fraction.Element(math + "num")!.Descendants(math + "t").Single()).ToArray();
            TestAssert.Equal(3, numerators.Length);
            return numerators;
        }
    }

    public static void PortableRenderingMathControlMetadataKeepsComposedFractions()
    {
        const string fixture = "CompositionProbes/composed-static-svg-report";
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var beforeDiagnostics = new List<OoxPdfDiagnostic>();
            var beforeSource = DocxInlineCompositionTests.ReadMutation(fixture, mode, false, false,
                diagnosticSink: beforeDiagnostics.Add);
            var renderer = new DocxRenderer(new PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var before = renderer.RenderPagesAsync(beforeSource, beforeDiagnostics.Add, default).ToBlockingEnumerable().ToArray();

            foreach (string propertyName in new[] { "noProof", "webHidden" })
            {
                var afterDiagnostics = new List<OoxPdfDiagnostic>();
                var afterSource = DocxInlineCompositionTests.ReadMutation(fixture, mode, false, false, (xml, w, m) =>
                {
                    var fractions = xml.Descendants(m + "f").ToArray();
                    TestAssert.Equal(3, fractions.Length);
                    string?[] values = [null, "1", "0"];
                    for (int i = 0; i < values.Length; i++)
                    {
                        var property = new XElement(w + propertyName);
                        property.SetAttributeValue(w + "val", values[i]);
                        fractions[i].AddFirst(new XElement(m + "fPr",
                            new XElement(m + "ctrlPr", new XElement(w + "rPr", property))));
                    }
                }, diagnosticSink: afterDiagnostics.Add);
                var after = renderer.RenderPagesAsync(afterSource, afterDiagnostics.Add, default).ToBlockingEnumerable().ToArray();
                string text = string.Concat(after.Select(ReadPortableEquationGlyphs));
                TestAssert.Equal(3, text.Split("x+2y", StringSplitOptions.None).Length - 1);
                TestAssert.Equal(before.Length, after.Length);
                for (int page = 0; page < before.Length; page++)
                {
                    TestAssert.Equal(before[page].Content, after[page].Content);
                    TestAssert.True(before[page].Annotations.SequenceEqual(after[page].Annotations),
                        $"{propertyName} preserves navigation.");
                }
                TestAssert.True(beforeDiagnostics.SequenceEqual(afterDiagnostics), $"{propertyName} preserves equation diagnostics.");
                CheckWrittenPages(after);
            }
        }
    }

    public static void PortableRenderingMathControlFormattingHistoryKeepsCurrentFractions()
    {
        const string fixture = "CompositionProbes/composed-static-svg-report";
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var baseline = Render(withHistory: false, hasPreviousProperties: false);
            var changed = Render(withHistory: true, hasPreviousProperties: true);
            string text = string.Concat(changed.Pages.Select(ReadPortableEquationGlyphs));
            TestAssert.Equal(3, text.Split("x+2y", StringSplitOptions.None).Length - 1);
            TestAssert.Equal(baseline.Pages.Length, changed.Pages.Length);
            for (int i = 0; i < changed.Pages.Length; i++)
            {
                TestAssert.Equal(baseline.Pages[i].Content, changed.Pages[i].Content);
                TestAssert.True(baseline.Pages[i].Annotations.SequenceEqual(changed.Pages[i].Annotations),
                    "Control formatting history preserves the current fractions' navigation.");
            }

            string revisionId = mode == OoxPdfDocxMarkupMode.Final ? "DOCX_UNSUPPORTED_FORMATTING_REVISIONS" :
                "DOCX_APPROXIMATED_FORMATTING_REVISIONS";
            TestAssert.Equal(1, changed.Diagnostics.Count(diagnostic => diagnostic.Id == revisionId));
            TestAssert.True(baseline.Diagnostics.SequenceEqual(changed.Diagnostics.Where(diagnostic => diagnostic.Id != revisionId)),
                "Only the formatting-revision diagnostic distinguishes the current paint controls.");
            CheckWrittenPages(changed.Pages);

            var malformed = Render(withHistory: true, hasPreviousProperties: false);
            TestAssert.True(malformed.Diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_UNSUPPORTED_EQUATION"),
                "A change without its former property record stays unsupported.");

            (PdfPage[] Pages, List<OoxPdfDiagnostic> Diagnostics) Render(bool withHistory, bool hasPreviousProperties)
            {
                var diagnostics = new List<OoxPdfDiagnostic>();
                var document = DocxInlineCompositionTests.ReadMutation(fixture, mode, false, false, (xml, w, m) =>
                {
                    XElement[] fractions = xml.Descendants(m + "f").ToArray();
                    TestAssert.Equal(3, fractions.Length);
                    for (int i = 0; i < fractions.Length; ++i)
                    {
                        var properties = new XElement(w + "rPr", new XElement(w + "color", new XAttribute(w + "val", "1964AE")));
                        if (withHistory)
                        {
                            var history = new XElement(w + "rPrChange", new XAttribute(w + "id", 700 + i),
                                new XAttribute(w + "author", "Public Reviewer"));
                            if (hasPreviousProperties)
                                history.Add(new XElement(w + "rPr", new XElement(w + "color", new XAttribute(w + "val", "CC0000"))));
                            properties.Add(history);
                        }
                        fractions[i].AddFirst(new XElement(m + "fPr", new XElement(m + "ctrlPr", properties)));
                    }
                }, diagnosticSink: diagnostics.Add);
                return (renderer.RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().ToArray(), diagnostics);
            }
        }
    }

    public static void PortableRenderingMathRunPaginationHintsKeepComposedContent()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (var source in new[]
        {
            (Fixture: "InlineBoxMath/inline-sibling-mixed", Part: "word/document.xml"),
            (Fixture: "InlineBoxMath/inline-box-math-cell", Part: "word/document.xml"),
            (Fixture: "InlineBoxMath/inline-box-math-header", Part: "word/inline-box-header.xml"),
            (Fixture: "CompositionProbes/math-note-content-footnote-inline", Part: "word/footnotes.xml")
        })
        {
            var beforeSource = DocxInlineCompositionTests.ReadMutation(source.Fixture, mode, false, false, partOverride: source.Part);
            var afterSource = DocxInlineCompositionTests.ReadMutation(source.Fixture, mode, false, false, (xml, w, m) =>
            {
                var runs = xml.Descendants(m + "r").ToArray();
                TestAssert.True(runs.Length != 0, "The source contains mathematical runs.");
                foreach (var run in runs) run.Add(new XElement(w + "lastRenderedPageBreak"));
            }, partOverride: source.Part);
            var renderer = new DocxRenderer(new PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var beforeDiagnostics = new List<OoxPdfDiagnostic>();
            var afterDiagnostics = new List<OoxPdfDiagnostic>();
            var before = renderer.RenderPagesAsync(beforeSource, beforeDiagnostics.Add, default).ToBlockingEnumerable().ToArray();
            var after = renderer.RenderPagesAsync(afterSource, afterDiagnostics.Add, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(before.Length, after.Length);
            TestAssert.True(beforeDiagnostics.SequenceEqual(afterDiagnostics), "Stored pagination markers do not change equation support diagnostics.");
            for (int page = 0; page < before.Length; page++)
            {
                TestAssert.Equal(before[page].Content, after[page].Content);
                TestAssert.Equal(before[page].Fonts.Count, after[page].Fonts.Count);
                TestAssert.Equal(before[page].Images.Count, after[page].Images.Count);
                TestAssert.Equal(before[page].Annotations.Count, after[page].Annotations.Count);
                PdfContentValidator.ValidatePage(after[page], page, default);
            }
        }
    }
}
