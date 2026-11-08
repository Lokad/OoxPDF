using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxFootnotesTests
{
    public static void DocxCompleteEndnoteKerningProtectsOwnedCustomBaselinePrefixes()
    {
        foreach (string mark in new[] { "*", "**", "*tail" })
        {
            DocxDocument source = ReadCustomDocumentEndnoteFlowFixture(mark);
            DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default);
            TestAssert.True(DocxNoteKerning.TryCreateDocument(source, layout, out var adjusted, out var pages, default),
                "A unique typed reference protects the custom prefix through the shared complete-note correction.");
            TestAssert.True(source.BodyElements.Zip(adjusted.BodyElements).All(pair => ReferenceEquals(pair.First, pair.Second)),
                "Custom note typography preserves body run/reference owners.");
            foreach (var pair in source.RelatedStories.Zip(adjusted.RelatedStories).Where(pair => pair.First.IsNormalStoryType))
            {
                DocxTextRun original = pair.First.Paragraphs[0].Runs[0];
                DocxTextRun[] runs = pair.Second.Paragraphs[0].Runs.ToArray();
                TestAssert.Equal(original with { Text = "*" }, runs[0]);
                TestAssert.True(!runs[0].IsAutomaticNoteStoryMarker, "A custom prefix does not acquire an automatic XML role.");
                if (mark.Length == 1) TestAssert.True(ReferenceEquals(original, runs[0]), "A complete one-character owner stays exact.");
                else
                {
                    TestAssert.Equal(mark[1..], runs[1].Text);
                    TestAssert.Equal(original.SourceRunIndex, runs[1].SourceRunIndex);
                    TestAssert.Equal(original.SourceTextOffsetInRun + 1, runs[1].SourceTextOffsetInRun);
                    TestAssert.True(!runs[1].PairKerningEnabled && runs[1].UseLogicalSizeForPdfAdvances,
                        "Adjacent source text resolves its own threshold without consuming the protected marker.");
                }
                TestAssert.True(!runs[^1].PairKerningEnabled && runs[^1].UseLogicalSizeForPdfAdvances,
                    "Regular note text uses the shared disabled-threshold rule.");
                TestAssert.Equal(mark, original.Text);
            }
            DocxLayout final = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(adjusted, new DocxTests.FamilyWidthTextMeasurer(), default);
            TestAssert.True(DocxNoteKerning.HasStableOwners(final, pages, default), "Owned custom markers keep the complete note-page map.");
        }
    }

    public static void DocxCompleteEndnoteKerningRejectsAmbiguousCustomSourceOwners()
    {
        DocxDocument plain = ReadCustomDocumentEndnoteFlowFixture("*");
        DocxParagraph body = plain.Paragraphs[0];
        DocxInlineReference reference = body.InlineReferences.Single();
        DocxRelatedStory first = plain.RelatedStories.First(s => s.IsNormalStoryType);
        DocxDocument ReplaceBody(DocxParagraph p) => plain with { BodyElements = plain.BodyElements.Select(e =>
            e is DocxParagraphElement pe && ReferenceEquals(pe.Paragraph, body) ? pe with { Paragraph = p } : e).ToArray() };
        foreach (DocxDocument source in new[]
        {
            ReplaceBody(body with { InlineReferences = [reference with { CustomMarkText = null }] }),
            ReplaceBody(body with { InlineReferences = [reference, reference] }),
            ReplaceBody(body with { InlineReferences = [reference with { SourceRunIndex = reference.SourceRunIndex + 100 }] }),
            ReplaceBody(body with { InlineReferences = [reference with { TextOffsetInRun = reference.TextOffsetInRun + 1 }] }),
            ReplaceBody(body with { Runs = body.Runs.Concat(body.Runs.Where(r => r.SourceRunIndex == reference.SourceRunIndex)).ToArray() }),
            plain with { RelatedStories = plain.RelatedStories.Select(s => ReferenceEquals(s, first) ? s with
            {
                BodyElements = [new DocxParagraphElement(first.Paragraphs[0] with
                { Runs = [first.Paragraphs[0].Runs[0] with { Text = "?" }, first.Paragraphs[0].Runs[1]] })],
            } : s).ToArray() },
        })
        {
            DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default);
            TestAssert.True(!DocxNoteKerning.TryCreateDocument(source, layout, out var adjusted, out _, default) && ReferenceEquals(source, adjusted),
                "Missing, repeated, overlapping and mismatched source owners keep the exact fallback.");
        }
    }

    public static void DocxCompleteEndnoteKerningPreservesBaselineAutomaticMarkerOwners()
    {
        DocxDocument source = ReadDocumentEndnoteFlowFixture(0, 0);
        source = source with { RelatedStories = source.RelatedStories.Select(s => s.IsNormalStoryType ? s with
        {
            BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
            {
                Runs = [p.Runs[0] with { VerticalAlignmentValue = "baseline", KerningMinimumHalfPointsValue = "broken" },
                    p.Runs[1] with { KerningMinimumHalfPointsValue = "0", PairKerningEnabled = true }],
            })).ToArray(),
        } : s).ToArray() };
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default);
        TestAssert.True(DocxNoteKerning.TryCreateDocument(source, layout, out var adjusted, out var pages, default),
            "Explicit automatic marker ownership admits the complete baseline frame through the shared rule.");
        foreach (var pair in source.RelatedStories.Zip(adjusted.RelatedStories).Where(pair => pair.First.IsNormalStoryType))
        {
            TestAssert.True(ReferenceEquals(pair.First.Paragraphs[0].Runs[0], pair.Second.Paragraphs[0].Runs[0]),
                "Baseline automatic markers preserve exact run ownership, raw threshold and advance policy.");
            TestAssert.True(!pair.Second.Paragraphs[0].Runs[1].PairKerningEnabled && pair.Second.Paragraphs[0].Runs[1].UseLogicalSizeForPdfAdvances,
                "Only the regular note text receives the shared resolved threshold and nominal advances.");
        }
        DocxLayout final = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(adjusted, new DocxTests.FamilyWidthTextMeasurer(), default);
        TestAssert.True(DocxNoteKerning.HasStableOwners(final, pages, default), "The owned retry preserves complete note pages.");
    }

    public static void DocxCompleteEndnoteKerningResolvesThresholdsWithoutRewritingOwners()
    {
        foreach ((double size, string? minimum, bool enabled, bool nominalAdvances) in new (double, string?, bool, bool)[]
        {
            (10d, null, false, true), (12d, "0", false, true), (14d, "0", false, true),
            (10d, "20", true, true), (12d, "24", true, false), (14d, "28", true, true),
            (10d, "21", false, true), (12d, "25", false, true), (14d, "29", false, true),
        })
        {
            DocxDocument source = ReadWrappedDocumentEndnoteFlowFixture(8, 4, 8);
            source = source with { RelatedStories = source.RelatedStories.Select(s => s.Kind == DocxRelatedStoryKind.Endnote && s.IsNormalStoryType ? s with
            {
                BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                {
                    Runs = p.Runs.Select(r => r.VerticalAlignment == DocxRunVerticalAlignment.Baseline ? r with
                    {
                        FontSize = size, KerningMinimumHalfPointsValue = minimum, PairKerningEnabled = !enabled,
                    } : r).ToArray(),
                })).ToArray(),
            } : s).ToArray() };
            DocxLayout first = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(source, new DocxTests.FamilyWidthTextMeasurer(), default);
            TestAssert.True(DocxNoteKerning.TryCreateDocument(source, first, out DocxDocument adjusted, out var pages, default),
                "An admitted complete paragraph resolves authored kerning against its nominal size.");
            TestAssert.Equal(2, pages.Count);
            TestAssert.True(source.BodyElements.Zip(adjusted.BodyElements).All(pair => ReferenceEquals(pair.First, pair.Second)),
                "Endnote correction preserves exact body owners and marker positions.");
            foreach (var pair in source.RelatedStories.Zip(adjusted.RelatedStories).Where(pair => pair.First.IsNormalStoryType))
            {
                TestAssert.True(ReferenceEquals(pair.First.Paragraphs[0].Runs[0], pair.Second.Paragraphs[0].Runs[0]),
                    "Automatic superscript marker runs preserve their source ownership and policy.");
                DocxTextRun oldRun = pair.First.Paragraphs[0].Runs[1], newRun = pair.Second.Paragraphs[0].Runs[1];
                TestAssert.Equal(minimum!, newRun.KerningMinimumHalfPointsValue!);
                TestAssert.Equal(enabled, newRun.PairKerningEnabled);
                TestAssert.Equal(oldRun with { PairKerningEnabled = enabled, UseLogicalSizeForPdfAdvances = nominalAdvances }, newRun);
                TestAssert.Equal(!enabled, oldRun.PairKerningEnabled);
            }
            DocxLayout final = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(adjusted, new DocxTests.FamilyWidthTextMeasurer(), default);
            TestAssert.True(DocxNoteKerning.HasStableOwners(final, pages, default),
                "The bounded retry retains the exact complete note-page mapping.");
            TestAssert.True(!DocxNoteKerning.HasStableOwners(final, new Dictionary<string, DocxNoteKerning.PageRange>(), default),
                "A changed ownership map cannot qualify the retry.");
        }
    }

    public static void DocxCompleteEndnoteKerningUsesNominalAdvancesAcrossExportGrid()
    {
        var font = OpenTypeFont.Load(TestFontBuilder.CreateTestFont());
        var embedded = PdfEmbeddedFont.Create(font, "AB".EnumerateRunes().Select(r => r.Value), default);
        DocxTextRun source = ReadDocumentEndnoteFlowFixture(0, 0).RelatedStories.First(s => s.IsNormalStoryType).Paragraphs[0].Runs[1];
        foreach ((double nominal, double physical) in new[] { (10d, 9.96d), (12d, 12d), (14d, 14.04d) })
        foreach (bool pairs in new[] { false, true })
        {
            var style = source with { Text = "AB", FontSize = nominal, PairKerningEnabled = pairs,
                UseLogicalSizeForPdfAdvances = true, CharacterSpacingPoints = 0d };
            var plan = DocxTextEmissionPlanner.Create(style, nominal, 0d, false, DocxTextStateCharacterSpacingSource.None);
            TestAssert.Equal(physical, plan.PdfFontSize);
            TestAssert.Equal(nominal, plan.NominalAdvanceFontSize!.Value);
            string payload = embedded.EncodeGlyphPositioningArray("AB", 0d, plan.PdfFontSize, true, pairs, plan.NominalAdvanceFontSize)!;
            double adjustment = System.Text.RegularExpressions.Regex.Matches(payload, @"(?<= )-?(?:\d+(?:\.\d*)?|\.\d+)(?= )")
                .Sum(m => double.Parse(m.Value, CultureInfo.InvariantCulture));
            double expected = .600d * nominal - (pairs ? .050d * nominal : 0d) + .620d * physical;
            double actual = 1.220d * physical - adjustment * physical / 1000d;
            TestAssert.True(Math.Abs(actual - expected) < .00001d,
                "Actual TJ keeps nominal preceding/pair advances and the final glyph's physical width on both sides of the export grid.");
            var profile = DocxTextEmissionPlanner.MeasureAdvanceProfile("AB", embedded, 1.220d * nominal - (pairs ? .050d * nominal : 0d), plan);
            TestAssert.True(Math.Abs(profile.PlannedEmittedAdvance - actual) < .00001d,
                "Emission inspection predicts the real stream rather than only the layout estimate.");
        }
    }

    public static void DocxCompleteEndnoteKerningRetainsMalformedAndUnqualifiedFallbacks()
    {
        DocxDocument plain = ReadWrappedDocumentEndnoteFlowFixture(8, 4, 8);
        DocxRelatedStory first = plain.RelatedStories.First(s => s.IsNormalStoryType);
        DocxParagraph paragraph = first.Paragraphs.Single();
        foreach (DocxDocument source in new[]
        {
            plain with { HeaderParagraphs = [plain.Paragraphs[0]] },
            plain with { RelatedStories = plain.RelatedStories.Select(s => ReferenceEquals(s, first) ? s with
            {
                BodyElements = [new DocxParagraphElement(paragraph with { LineSpacingPoints = null,
                    Spacing = paragraph.Spacing with { LineRuleValue = "exact" } })],
            } : s).ToArray() },
            plain with { RelatedStories = plain.RelatedStories.Select(s => ReferenceEquals(s, first) ? s with
            {
                BodyElements = [new DocxParagraphElement(paragraph with
                {
                    Runs = [paragraph.Runs[0] with { VerticalAlignmentValue = "baseline", IsAutomaticNoteStoryMarker = false }, paragraph.Runs[1]],
                })],
            } : s).ToArray() },
            plain with { RelatedStories = plain.RelatedStories.Select(s => ReferenceEquals(s, first) ? s with
            {
                BodyElements = [new DocxParagraphElement(paragraph with
                {
                    Runs = [paragraph.Runs[0], paragraph.Runs[1] with { KerningMinimumHalfPointsValue = "broken" }],
                })],
            } : s).ToArray() },
        })
        {
            DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(source, new DocxTests.FamilyWidthTextMeasurer(), default);
            TestAssert.True(!DocxNoteKerning.TryCreateDocument(source, layout, out var adjusted, out _, default) && ReferenceEquals(source, adjusted),
                "Unsupported ownership or thresholds retain the original document and layout policy.");
        }
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxNoteKerning.TryCreateDocument(plain,
            new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(plain, new DocxTests.FamilyWidthTextMeasurer(), default),
            out _, out _, cancelled.Token));
    }

    public static void DocxContinuedEndnoteKerningPreservesSourceMarkersAndPageRanges()
    {
        DocxDocument source = ReadWrappedDocumentEndnoteFlowFixture(0, 0, 200);
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new DocxTests.FamilyWidthTextMeasurer();
        DocxLayout first = engine.Create(source, measurer, default);
        TestAssert.True(DocxNoteKerning.TryCreateDocument(source, first, out var adjusted, out var pages, default),
            "Plain continued paragraphs share the authored threshold policy with fitting endnotes.");
        TestAssert.True(pages.Values.Any(range => range.Slices > 1 && range.Last > range.First),
            "The regression exercises an actual multi-page note rather than only a fitting block.");
        TestAssert.True(source.BodyElements.Zip(adjusted.BodyElements).All(pair => ReferenceEquals(pair.First, pair.Second)),
            "Body reference ownership is unchanged by continued-note typography.");
        foreach (var pair in source.RelatedStories.Zip(adjusted.RelatedStories).Where(pair => pair.First.IsNormalStoryType))
        {
            TestAssert.True(ReferenceEquals(pair.First.Paragraphs[0].Runs[0], pair.Second.Paragraphs[0].Runs[0]),
                "Each source marker retains its exact owner and emission policy.");
            TestAssert.True(pair.First.Paragraphs[0].Runs[1].PairKerningEnabled,
                "Source run policy remains immutable.");
            TestAssert.True(!pair.Second.Paragraphs[0].Runs[1].PairKerningEnabled &&
                pair.Second.Paragraphs[0].Runs[1].UseLogicalSizeForPdfAdvances,
                "Absent thresholds use nominal regular-text advances on every slice.");
        }
        DocxLayout final = engine.Create(adjusted, measurer, default);
        TestAssert.True(DocxNoteKerning.HasStableOwners(final, pages, default),
            "The typography retry preserves every source owner and first destination page.");
        var changed = pages.ToDictionary(p => p.Key, p => p.Value);
        string id = changed.First(p => p.Value.Slices > 1).Key;
        changed[id] = changed[id] with { Last = changed[id].Last + 1 };
        TestAssert.True(DocxNoteKerning.HasStableOwners(final, changed, default),
            "A legitimate continuation-range change does not invalidate stable first placements.");
        changed[id] = pages[id] with { Slices = pages[id].Slices + 1 };
        TestAssert.True(!DocxNoteKerning.HasStableOwners(final, changed, default),
            "A slice count exceeding its page range remains invalid source ownership metadata.");
    }

    public static void DocxContinuedEndnoteKerningRejectsUnqualifiedSlicesAndKeptParagraphs()
    {
        DocxDocument source = ReadWrappedDocumentEndnoteFlowFixture(0, 0, 200);
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new DocxTests.FamilyWidthTextMeasurer();
        DocxLayout layout = engine.Create(source, measurer, default);
        TestAssert.True(DocxNoteKerning.TryCreateDocument(source, layout, out _, out var pages, default),
            "The control has qualified multi-page typography ownership.");
        DocxLayout unqualified = layout with { Pages = layout.Pages.Select(page => page with
        {
            PlacedRelatedStories = page.PlacedRelatedStories.Select(p => p.ContentTopOffset > 0d
                ? p with { UsesPlainDocumentEndnoteTypography = false } : p).ToArray(),
        }).ToArray() };
        TestAssert.True(!DocxNoteKerning.TryCreateDocument(source, unqualified, out var unchanged, out _, default) &&
            ReferenceEquals(source, unchanged), "Every continued slice must retain layout-owned admission.");
        TestAssert.True(!DocxNoteKerning.HasStableOwners(unqualified, pages, default),
            "Unqualified continuation pages cannot pass the post-retry check.");
        DocxDocument kept = source with { RelatedStories = source.RelatedStories.Select(story => story.IsNormalStoryType
            ? story with { BodyElements = story.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                { KeepRules = p.KeepRules with { KeepLines = true } })).ToArray() } : story).ToArray() };
        TestAssert.True(!DocxNoteKerning.TryCreateDocument(kept, engine.Create(kept, measurer, default),
            out unchanged, out _, default) && ReferenceEquals(kept, unchanged),
            "Continuation does not bypass the established semantic keep-rule fallback.");
    }

    public static void DocxContinuedEndnoteKerningAdmitsRealContinuationReflowWithStableFirstOwners()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new LogicalBoundaryTextMeasurer();
        DocxDocument? control = null;
        DocxDocument? adjustedControl = null;
        DocxLayout? finalControl = null;
        IReadOnlyDictionary<string, DocxNoteKerning.PageRange>? expected = null;
        for (int repeats = 20; repeats <= 340 && control is null; repeats++)
        {
            DocxDocument source = ReadWrappedDocumentEndnoteFlowFixture(0, 0, repeats);
            string? shortId = source.RelatedStories.Where(s => s.IsNormalStoryType).Skip(1).First().Id;
            source = WithEndnoteText(source, shortId, "Short note.");
            DocxLayout first = engine.Create(source, measurer, default);
            if (!DocxNoteKerning.TryCreateDocument(source, first, out var adjusted, out var pages, default)) continue;
            DocxLayout final = engine.Create(adjusted, measurer, default);
            var actual = final.Pages.SelectMany((page, index) => page.PlacedRelatedStories
                .Where(s => s.StoryLayout.Story.IsNormalStoryType).Select(s => (Id: s.StoryLayout.Story.Id!, Page: index)))
                .GroupBy(s => s.Id).ToDictionary(g => g.Key, g => (First: g.Min(s => s.Page), Last: g.Max(s => s.Page), Slices: g.Count()));
            if (final.Pages.Count != first.Pages.Count || actual.Count != pages.Count ||
                actual.Any(p => !pages.TryGetValue(p.Key, out var previous) || p.Value.First != previous.First) ||
                !actual.Any(p => p.Value.Last != pages[p.Key].Last || p.Value.Slices != pages[p.Key].Slices)) continue;
            control = source; adjustedControl = adjusted; finalControl = final; expected = pages;
        }
        TestAssert.True(control is not null && finalControl is not null && expected is not null,
            "A bounded width/capacity scan creates real continuation reflow while retaining first owners and total page count.");
        TestAssert.True(DocxNoteKerning.HasStableOwners(finalControl!, expected!, default),
            "Corrected typography can change continuation slices without falling back to the uncorrected advances.");
        TestAssert.True(control!.BodyElements.Zip(adjustedControl!.BodyElements).All(p => ReferenceEquals(p.First, p.Second)),
            "Reflow retains exact source body and inline-reference owners.");
        TestAssert.True(control.RelatedStories.Zip(adjustedControl.RelatedStories).Where(p => p.First.IsNormalStoryType)
            .All(p => ReferenceEquals(p.First.Paragraphs[0].Runs[0], p.Second.Paragraphs[0].Runs[0])),
            "Every source marker preserves its original emission policy through the retry.");
        var missingOwner = expected!.ToDictionary(p => p.Key, p => p.Value);
        missingOwner.Remove(missingOwner.Keys.First());
        TestAssert.True(!DocxNoteKerning.HasStableOwners(finalControl!, missingOwner, default),
            "A missing source owner retains the explicit fallback.");
    }

    public static void DocxContinuedEndnoteKerningAllowsNoteOnlyPageGrowthWithStableBodyAndFirstOwners()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new LogicalBoundaryTextMeasurer();
        DocxDocument? control = null;
        DocxDocument? adjustedControl = null;
        DocxLayout? before = null;
        DocxLayout? after = null;
        IReadOnlyDictionary<string, DocxNoteKerning.PageRange>? expected = null;
        for (int repeats = 20; repeats <= 340 && control is null; repeats++)
        {
            DocxDocument source = ReadWrappedDocumentEndnoteFlowFixture(0, 0, repeats);
            DocxLayout first = engine.Create(source, measurer, default);
            if (!DocxNoteKerning.TryCreateDocument(source, first, out var adjusted, out var pages, default)) continue;
            DocxLayout final = engine.Create(adjusted, measurer, default);
            var actual = final.Pages.SelectMany((page, index) => page.PlacedRelatedStories
                .Where(s => s.StoryLayout.Story.IsNormalStoryType).Select(s => (Id: s.StoryLayout.Story.Id!, Page: index)))
                .GroupBy(s => s.Id).ToDictionary(g => g.Key, g => g.Min(s => s.Page));
            if (final.Pages.Count <= first.Pages.Count || actual.Count != pages.Count ||
                actual.Any(p => !pages.TryGetValue(p.Key, out var previous) || p.Value != previous.First)) continue;
            control = source; adjustedControl = adjusted; before = first; after = final; expected = pages;
        }
        TestAssert.True(control is not null && before is not null && after is not null,
            "A bounded advance/capacity scan adds real note-only pages while retaining first destinations.");
        TestAssert.True(DocxPlainBodyFrame.IsEligible(control!), "The admitted body has no dynamic fields, headers, footers or composed content.");
        TestAssert.True(DocxNoteKerning.HasStableOwners(after!, expected!, default),
            "Every note retains its qualified source owner and first placement despite the added tail page.");
        var oldBody = before!.Pages.SelectMany((page, index) => page.Items.OfType<DocxTextLineLayout>()
            .Select(l => (index, l.Text, l.X, l.BaselineY, l.Width, l.FontSize, l.SourceBlockIndex, l.SourceParagraphIndex, l.SourceLineIndex)));
        var newBody = after!.Pages.SelectMany((page, index) => page.Items.OfType<DocxTextLineLayout>()
            .Select(l => (index, l.Text, l.X, l.BaselineY, l.Width, l.FontSize, l.SourceBlockIndex, l.SourceParagraphIndex, l.SourceLineIndex)));
        TestAssert.True(oldBody.SequenceEqual(newBody), "Added note pages preserve every body line, geometry and source coordinate.");
        TestAssert.True(after.Pages.Skip(before.Pages.Count).All(p => p.Items.Count == 0), "Every newly added page contains only related-story content.");
        TestAssert.True(control!.BodyElements.Zip(adjustedControl!.BodyElements).All(p => ReferenceEquals(p.First, p.Second)),
            "Source body and inline references remain immutable through the bounded retry.");
        foreach (DocxRelatedStoryLayout story in after.RelatedStories.Where(s => s.Story.IsNormalStoryType))
        {
            var slices = after.Pages.SelectMany(p => p.PlacedRelatedStories).Where(s => ReferenceEquals(s.StoryLayout.Story, story.Story)).ToArray();
            TestAssert.Equal(story.TextLines.Count, slices.Sum(s => s.TextLines.Count));
            TestAssert.Equal(1, slices.Count(s => s.ContentTopOffset == 0d));
        }
    }

    public static void DocxContinuedEndnoteKerningReflowsFirstPlacementsWithStableSourceAndBody()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new LogicalBoundaryTextMeasurer();
        DocxDocument? control = null;
        DocxDocument? adjustedControl = null;
        DocxLayout? before = null;
        DocxLayout? after = null;
        IReadOnlyDictionary<string, DocxNoteKerning.PageRange>? expected = null;
        for (int repeats = 20; repeats <= 340 && control is null; repeats++)
        {
            DocxDocument source = ReadWrappedDocumentEndnoteFlowFixture(0, 0, repeats);
            DocxLayout first = engine.Create(source, measurer, default);
            if (!DocxNoteKerning.TryCreateDocument(source, first, out var adjusted, out var pages, default)) continue;
            DocxLayout final = engine.Create(adjusted, measurer, default);
            var actual = final.Pages.SelectMany((page, index) => page.PlacedRelatedStories
                .Where(s => s.StoryLayout.Story.IsNormalStoryType).Select(s => (Id: s.StoryLayout.Story.Id!, Page: index)))
                .GroupBy(s => s.Id).ToDictionary(g => g.Key, g => g.Min(s => s.Page));
            if (actual.Count != pages.Count || !actual.Any(p => pages.TryGetValue(p.Key, out var previous) && p.Value != previous.First)) continue;
            control = source; adjustedControl = adjusted; before = first; after = final; expected = pages;
        }
        TestAssert.True(control is not null && before is not null && after is not null,
            "A bounded advance scan exercises a real first-placement change with complete note ownership.");
        TestAssert.True(DocxPlainBodyFrame.IsEligible(control!), "The admitted body has no fields or composed static stories.");
        TestAssert.True(DocxNoteKerning.HasStableOwners(after!, expected!, default),
            "Corrected wrapping can move first destinations without rejecting the existing source owners.");
        var oldBody = before!.Pages.SelectMany((page, index) => page.Items.OfType<DocxTextLineLayout>()
            .Select(l => (index, l.Text, l.X, l.BaselineY, l.Width, l.FontSize, l.SourceBlockIndex, l.SourceParagraphIndex, l.SourceLineIndex)));
        var newBody = after!.Pages.SelectMany((page, index) => page.Items.OfType<DocxTextLineLayout>()
            .Select(l => (index, l.Text, l.X, l.BaselineY, l.Width, l.FontSize, l.SourceBlockIndex, l.SourceParagraphIndex, l.SourceLineIndex)));
        TestAssert.True(oldBody.SequenceEqual(newBody), "First-note page reflow preserves every body operation and source coordinate.");
        TestAssert.True(control!.BodyElements.Zip(adjustedControl!.BodyElements).All(p => ReferenceEquals(p.First, p.Second)),
            "Authored body references remain immutable through the bounded retry.");
        TestAssert.True(control.RelatedStories.Zip(adjustedControl.RelatedStories).Where(p => p.First.IsNormalStoryType)
            .All(p => ReferenceEquals(p.First.Paragraphs[0].Runs[0], p.Second.Paragraphs[0].Runs[0])),
            "Source markers retain their exact owners and numbering policy.");
        foreach (DocxRelatedStoryLayout story in after.RelatedStories.Where(s => s.Story.IsNormalStoryType))
        {
            var slices = after.Pages.SelectMany(p => p.PlacedRelatedStories).Where(s => ReferenceEquals(s.StoryLayout.Story, story.Story)).ToArray();
            var placedLines = slices.SelectMany(s => s.TextLines)
                .Select(l => (l.Text, l.SourceParagraph, l.SourceBlockIndex, l.SourceParagraphIndex, l.SourceLineIndex));
            var sourceLines = story.TextLines
                .Select(l => (l.Text, l.SourceParagraph, l.SourceBlockIndex, l.SourceParagraphIndex, l.SourceLineIndex));
            TestAssert.True(placedLines.SequenceEqual(sourceLines), "Every source line is emitted once in order despite its page translation.");
            TestAssert.Equal(1, slices.Count(s => s.ContentTopOffset == 0d));
        }
        var substituted = expected!.ToDictionary(p => p.Key, p => p.Value);
        string id = substituted.Keys.First();
        substituted.Add("foreign-owner", substituted[id]); substituted.Remove(id);
        TestAssert.True(!DocxNoteKerning.HasStableOwners(after, substituted, default), "A substituted owner is rejected despite equal counts.");
        var invalid = expected!.ToDictionary(p => p.Key, p => p.Value);
        invalid[id] = invalid[id] with { First = -1 };
        TestAssert.True(!DocxNoteKerning.HasStableOwners(after, invalid, default), "Malformed original ownership ranges remain invalid.");
    }

    public static void DocxContinuedEndnoteKerningAdmitsFittingStaticFramesAndRetainsFrameFallbacks()
    {
        DocxDocument plain = ReadWrappedDocumentEndnoteFlowFixture(0, 0, 80);
        DocxParagraph staticParagraph = DocxTests.CreateDocxLayoutParagraph("public frame", 12d, 14d) with
        {
            LineSpacingPoints = null,
            Spacing = DocxParagraphSpacing.Empty with { BeforeValue = "0", AfterValue = "0", LineValue = "240", LineRuleValue = "auto" },
        };
        DocxDocument Compose(bool header, bool footer, DocxParagraph paragraph, int count = 1, DocxDocument? body = null)
        {
            DocxDocument bodySource = body ?? plain;
            var paragraphs = Enumerable.Repeat(paragraph, count).ToArray();
            var headers = new Dictionary<string, IReadOnlyList<DocxParagraph>>();
            var footers = new Dictionary<string, IReadOnlyList<DocxParagraph>>();
            if (header) headers["default"] = paragraphs;
            if (footer) footers["default"] = paragraphs;
            var headerElements = headers.ToDictionary(p => p.Key, p => (IReadOnlyList<DocxBodyElement>)p.Value.Select(v => (DocxBodyElement)new DocxParagraphElement(v)).ToArray());
            var footerElements = footers.ToDictionary(p => p.Key, p => (IReadOnlyList<DocxBodyElement>)p.Value.Select(v => (DocxBodyElement)new DocxParagraphElement(v)).ToArray());
            var settings = bodySource.PageSettings with { HeaderDistancePoints = 36d, FooterDistancePoints = 36d,
                HeaderParagraphsByType = headers, FooterParagraphsByType = footers,
                HeaderBodyElementsByType = headerElements, FooterBodyElementsByType = footerElements };
            return bodySource with { PageSettings = settings, FinalSectionBreak = bodySource.FinalSectionBreak! with { PageSettings = settings },
                HeaderParagraphs = header ? paragraphs : [], FooterParagraphs = footer ? paragraphs : [],
                HeaderParagraphsByType = headers, FooterParagraphsByType = footers,
                HeaderBodyElementsByType = headerElements, FooterBodyElementsByType = footerElements };
        }
        var measurer = new LogicalBoundaryTextMeasurer();
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        foreach ((bool header, bool footer) in new[] { (true, false), (false, true), (true, true) })
        {
            DocxDocument source = Compose(header, footer, staticParagraph);
            TestAssert.True(DocxPlainBodyFrame.HasPlainBody(source) && !DocxPlainBodyFrame.IsEligible(source),
                "Fitting static admission leaves the original body typography and page-start qualifier unchanged.");
            DocxLayout before = engine.Create(source, measurer, default);
            TestAssert.True(before.Pages.All(p => p.StaticTextLines.Count != 0), "The control renders its real static stories on every page.");
            TestAssert.True(DocxNoteKerning.TryCreateDocument(source, before, out var adjusted, out var pages, default),
                "Automatic plain static stories fit outside the measured body frame and share note typography.");
            DocxLayout after = engine.Create(adjusted, measurer, default);
            TestAssert.True(DocxNoteKerning.HasStableOwners(after, pages, default), "All note slices retain complete source ownership.");
            TestAssert.True(ReferenceEquals(source.PageSettings, adjusted.PageSettings) && ReferenceEquals(source.HeaderParagraphs, adjusted.HeaderParagraphs) &&
                ReferenceEquals(source.FooterParagraphs, adjusted.FooterParagraphs), "The retry preserves authored static story ownership and policy.");
            TestAssert.True(before.Pages.Take(Math.Min(before.Pages.Count, after.Pages.Count)).Zip(after.Pages)
                .All(pair => pair.First.StaticTextLines.Select(l => (l.Text, l.X, l.BaselineY, l.Width, l.FontSize,
                    l.SourceBlockIndex, l.SourceParagraphIndex, l.SourceLineIndex)).SequenceEqual(
                    pair.Second.StaticTextLines.Select(l => (l.Text, l.X, l.BaselineY, l.Width, l.FontSize,
                        l.SourceBlockIndex, l.SourceParagraphIndex, l.SourceLineIndex)))), "Static geometry stays exact on retained pages.");
            DocxDocument overflow = Compose(header, footer, staticParagraph, 10);
            TestAssert.True(DocxNoteKerning.TryCreateDocument(overflow, engine.Create(overflow, measurer, default), out _, out _, default),
                "Plain static overflow leaves a usable frame bounded by the shared story edges.");
            DocxDocument consumed = Compose(header, footer, staticParagraph, 100);
            TestAssert.True(!DocxNoteKerning.TryCreateDocument(consumed, engine.Create(consumed, measurer, default), out _, out _, default),
                "Static content consuming the entire frame retains its fallback and mandatory line progress.");
        }
        foreach (DocxParagraph paragraph in new[] { staticParagraph with { LineSpacingPoints = 14d },
            staticParagraph with { Runs = [staticParagraph.Runs[0] with { Text = "{NUMPAGES}", FieldKind = DocxFieldKind.NumPages }] } })
        {
            DocxDocument source = Compose(true, true, paragraph);
            TestAssert.True(!DocxNoteKerning.TryCreateDocument(source, engine.Create(source, measurer, default), out _, out _, default),
                "Fixed-height or dynamic static stories retain their explicitly unqualified frame policy.");
        }
        DocxDocument variant = Compose(true, false, staticParagraph);
        var variantParagraphs = variant.PageSettings.HeaderParagraphsByType.ToDictionary(p => p.Key, p => p.Value);
        variantParagraphs["even"] = Enumerable.Repeat(staticParagraph, 100).ToArray();
        var variantElements = variantParagraphs.ToDictionary(p => p.Key,
            p => (IReadOnlyList<DocxBodyElement>)p.Value.Select(v => (DocxBodyElement)new DocxParagraphElement(v)).ToArray());
        var variantSettings = variant.PageSettings with { EvenAndOddHeaders = true,
            HeaderParagraphsByType = variantParagraphs, HeaderBodyElementsByType = variantElements };
        variant = variant with { PageSettings = variantSettings, FinalSectionBreak = variant.FinalSectionBreak! with { PageSettings = variantSettings },
            HeaderParagraphsByType = variantParagraphs, HeaderBodyElementsByType = variantElements };
        TestAssert.True(!DocxNoteKerning.TryCreateDocument(variant, engine.Create(variant, measurer, default), out _, out _, default),
            "A fitting default story does not admit a consumed frame on a reachable even continuation page.");
        DocxDocument mixedBody = ReadDocumentEndnoteFlowFixture(0, 0);
        DocxDocument footnotes = ReadNotePrecedenceFixture("footnote", "decimal", null, null, true);
        mixedBody = mixedBody with { RelatedStories = mixedBody.RelatedStories.Concat(footnotes.RelatedStories
            .Where(s => s.IsNormalStoryType && s.Kind == DocxRelatedStoryKind.Footnote)).ToArray() };
        DocxDocument mixed = Compose(true, true, staticParagraph, body: mixedBody);
        DocxLayout mixedLayout = engine.Create(mixed, measurer, default);
        var mixedNotes = mixedLayout.Pages.SelectMany(p => p.PlacedRelatedStories)
            .Where(s => s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote).ToArray();
        TestAssert.Equal(2, mixedNotes.Length);
        var withoutStatic = engine.Create(mixedBody, measurer, default).Pages.SelectMany(p => p.PlacedRelatedStories)
            .Where(s => s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote).ToArray();
        TestAssert.True(mixedNotes.Select(note => (note.TopY, note.Height)).SequenceEqual(
            withoutStatic.Select(note => (note.TopY, note.Height))),
            "Independent static stories and unreferenced footnote sources preserve complete endnote paragraph boundaries.");
        TestAssert.True(!DocxNoteKerning.TryCreateDocument(mixed, mixedLayout, out _, out _, default),
            "The static-frame extension preserves the existing mixed-note typography fallback.");
    }

    public static void DocxMinimumEndnoteTypographySharesPlainParagraphFlow()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new NoteContentMetricMeasurer();
        foreach (double minimum in new[] { 8d, 18d })
        foreach (string threshold in new[] { "0", "20", "30" })
        {
            DocxDocument source = ReadWrappedDocumentEndnoteFlowFixture(0, 0, 200);
            source = source with { RelatedStories = source.RelatedStories.Select(s => s.IsNormalStoryType ? s with
            {
                BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                {
                    LineSpacingPoints = minimum, Spacing = p.Spacing with { LineRuleValue = "atLeast",
                        LineValue = (minimum * 20d).ToString(CultureInfo.InvariantCulture) },
                    Runs = [p.Runs[0], p.Runs[1] with { FontSize = 10d, KerningMinimumHalfPointsValue = threshold,
                        PairKerningEnabled = threshold != "20" }],
                })).ToArray(),
            } : s).ToArray() };
            DocxLayout layout = engine.Create(source, measurer, default);
            var slices = layout.Pages.SelectMany(p => p.PlacedRelatedStories)
                .Where(s => s.StoryLayout.Story.IsNormalStoryType).ToArray();
            TestAssert.True(slices.Any(s => s.ContentTopOffset > 0d) && slices.All(s => s.UsesPlainDocumentEndnoteTypography),
                "Minimum endnote continuations admit source-owned typography independently of automatic paragraph flow.");
            TestAssert.True(DocxNoteKerning.TryCreateDocument(source, layout, out var adjusted, out var owners, default),
                "Minimum content shares the existing absent, met and unmet threshold policy.");
            foreach (var pair in source.RelatedStories.Zip(adjusted.RelatedStories).Where(p => p.First.IsNormalStoryType))
            {
                var before = pair.First.Paragraphs[0]; var after = pair.Second.Paragraphs[0];
                TestAssert.True(ReferenceEquals(before.Runs[0], after.Runs[0]), "The source marker remains immutable.");
                TestAssert.Equal(before.Spacing, after.Spacing);
                TestAssert.Equal(before.LineSpacingPoints!.Value, after.LineSpacingPoints!.Value);
                TestAssert.Equal(threshold == "20", after.Runs[1].PairKerningEnabled);
            }
            TestAssert.True(ReferenceEquals(source.BodyElements, adjusted.BodyElements), "Body owners remain unchanged.");
            TestAssert.True(DocxNoteKerning.HasStableOwners(engine.Create(adjusted, measurer, default), owners, default),
                "All canonical minimum lines retain complete continuation ownership after the content retry.");
        }
        DocxDocument fitting = ReadDocumentEndnoteFlowFixture(0, 0);
        fitting = fitting with { RelatedStories = fitting.RelatedStories.Select(s => s.IsNormalStoryType ? s with
        {
            BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
            {
                LineSpacingPoints = 18d, Spacing = p.Spacing with { LineRuleValue = "atLeast", LineValue = "360" },
            })).ToArray(),
        } : s).ToArray() };
        var notes = engine.Create(fitting, measurer, default).Pages.Single().PlacedRelatedStories
            .Where(s => s.StoryLayout.Story.IsNormalStoryType).ToArray();
        TestAssert.True(notes.All(s => s.UsesPlainDocumentEndnoteTypography), "Fitting minima share content admission.");
        TestAssert.True(Math.Abs(notes[0].TopY - notes[1].TopY - notes[0].Height) < .000001d,
            "Plain minimum paragraphs share the Office-qualified adjacent paragraph flow without an extra note gap.");
    }

    public static void DocxEndnoteKerningRejectsIncompleteCanonicalLineOwners()
    {
        DocxDocument source = ReadWrappedDocumentEndnoteFlowFixture(0, 0, 200);
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        DocxLayout layout = engine.Create(source, new NoteContentMetricMeasurer(), default);
        TestAssert.True(DocxNoteKerning.TryCreateDocument(source, layout, out _, out var owners, default),
            "The automatic endnote control has complete canonical source owners.");
        foreach (Func<IReadOnlyList<DocxTextLineLayout>, IReadOnlyList<DocxTextLineLayout>> modify in new
            Func<IReadOnlyList<DocxTextLineLayout>, IReadOnlyList<DocxTextLineLayout>>[]
        {
            lines => lines.Skip(1).ToArray(),
            lines => lines.Reverse().ToArray(),
            lines => lines.Select((line, i) => i == 0 ? line with { SourceLineIndex = -1 } : line).ToArray(),
            lines => lines.Select((line, i) => i == 0 ? line with { SourceParagraph = line.SourceParagraph! with { } } : line).ToArray(),
        })
        {
            bool changed = false;
            DocxLayout invalid = layout with { Pages = layout.Pages.Select(page => page with
            {
                PlacedRelatedStories = page.PlacedRelatedStories.Select(placed =>
                {
                    if (changed || !placed.StoryLayout.Story.IsNormalStoryType) return placed;
                    changed = true;
                    var lines = modify(placed.StoryLayout.TextLines);
                    return placed with { StoryLayout = placed.StoryLayout with { TextLines = lines }, TextLines = lines };
                }).ToArray(),
            }).ToArray() };
            TestAssert.True(!DocxNoteKerning.TryCreateDocument(source, invalid, out var unchanged, out _, default) &&
                ReferenceEquals(source, unchanged), "Missing, reordered or substituted source lines retain the original document.");
            TestAssert.True(!DocxNoteKerning.HasStableOwners(invalid, owners, default),
                "Equal note IDs and page ranges cannot qualify incomplete canonical line ownership.");
        }
    }
}
