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
    public static void DocxMixedEndnoteCapacityExcludesLatentSeparatorSpaceWithoutMovingInk()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new NoteContentMetricMeasurer();
        foreach (bool custom in new[] { false, true })
        {
            DocxDocument source = WithMixedNoteStaticFrame(ReadMixedNoteRetryFixture(custom, 140), true, true, 1);
            DocxRelatedStory end = source.RelatedStories.Single(s => s.IsNormalStoryType && s.Kind == DocxRelatedStoryKind.Endnote);
            DocxParagraph paragraph = end.Paragraphs[0];
            DocxParagraph fitting = paragraph with { Runs = paragraph.Runs.Select((r, index) =>
                index == paragraph.Runs.Count - 1 ? r with { Text = " Public fitting complement." } : r).ToArray() };
            DocxDocument fittingSource = source with { RelatedStories = source.RelatedStories.Select(s =>
                ReferenceEquals(s, end) ? s with { BodyElements = [new DocxParagraphElement(fitting)] } : s).ToArray() };
            DocxLayout layout = engine.Create(source, measurer, default);
            DocxLayout fittingLayout = engine.Create(fittingSource, measurer, default);
            DocxPlacedRelatedStoryLayout separator = layout.Pages[0].PlacedRelatedStories.Single(s =>
                s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && s.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
            DocxPlacedRelatedStoryLayout fittingSeparator = fittingLayout.Pages[0].PlacedRelatedStories.Single(s =>
                s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && s.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
            var ends = layout.Pages.SelectMany(p => p.PlacedRelatedStories).Where(s =>
                s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote).ToArray();
            double headBottom = ends[0].TopY - ends[0].Height;
            TestAssert.True(ends.Length > 1 && separator.SeparatorY is { } ruleY &&
                headBottom < separator.TopY && headBottom > ruleY + separator.SeparatorThickness,
                "The continued head may use empty separator padding while clearing its visible rule.");
            TestAssert.True(separator.TopY == fittingSeparator.TopY && separator.Height == fittingSeparator.Height &&
                separator.SeparatorY == fittingSeparator.SeparatorY && separator.SeparatorThickness == fittingSeparator.SeparatorThickness &&
                separator.SeparatorWidth == fittingSeparator.SeparatorWidth,
                "Head capacity leaves the complementary separator clip and visible rule exactly unchanged.");
            var foot = layout.Pages[0].PlacedRelatedStories.Single(s => s.StoryLayout.Story.IsNormalStoryType &&
                s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote);
            var fittingFoot = fittingLayout.Pages[0].PlacedRelatedStories.Single(s => s.StoryLayout.Story.IsNormalStoryType &&
                s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote);
            TestAssert.True(foot.TopY == fittingFoot.TopY && foot.Height == fittingFoot.Height && headBottom >= foot.TopY &&
                ReferenceEquals(foot.StoryLayout.Story.Paragraphs[0], fittingFoot.StoryLayout.Story.Paragraphs[0]),
                "The reused capacity clears the complete original complementary footnote box and source.");
            DocxRelatedStoryLayout canonical = layout.RelatedStories.Single(s => s.Story.IsNormalStoryType &&
                s.Story.Kind == DocxRelatedStoryKind.Endnote);
            TestAssert.Equal(string.Concat(canonical.TextLines.Select(l => l.Text)),
                string.Concat(ends.SelectMany(s => s.StoryLayout.TextLines).Select(l => l.Text)));
            TestAssert.True(ends.All(s => !s.UsesCompleteDocumentEndnoteFlow),
                "Separator capacity preserves the existing mixed paragraph-boundary policy.");
        }
    }

    public static void DocxDeferredMixedEndnoteUsesResolvedSeparatorFirstLineOrigin()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var witnessed = new HashSet<(double? Ascent, bool Custom, int Size, bool Registered)>();
        foreach (double? ascent in new double?[] { null, 12d })
        foreach (bool custom in new[] { false, true })
        foreach (int size in new[] { 8, 12, 16 })
        foreach (bool registered in new[] { false, true })
        foreach (int footerLines in new[] { 1, 6 })
        foreach (int footRepeats in new[] { 96, 100, 104, 120 })
        {
            var measurer = new NoteContentMetricMeasurer(ascent);
            DocxDocument source = ReadMixedNoteRetryFixture(custom, 32, footRepeats, (_, settings, parts) =>
            {
                XDocument notes = XDocument.Parse(parts["word/endnotes.xml"]);
                XElement separator = notes.Root!.Elements(NoteWord + "endnote").Single(n =>
                    (string?)n.Attribute(NoteWord + "type") == "separator");
                XElement properties = new(NoteWord + "rPr", new XElement(NoteWord + "sz",
                    new XAttribute(NoteWord + "val", size * 2)));
                separator.Element(NoteWord + "p")!.Element(NoteWord + "pPr")!.Add(new XElement(properties));
                separator.Descendants(NoteWord + "r").Single().AddFirst(new XElement(properties));
                parts["word/endnotes.xml"] = notes.ToString();
                if (registered) settings.Root!.Add(new XElement(NoteWord + "endnotePr",
                    new XElement(NoteWord + "endnote", new XAttribute(NoteWord + "id", "-1"))));
            });
            source = WithMixedNoteStaticFrame(source, false, true, footerLines);
            DocxLayout layout = engine.Create(source, measurer, default);
            DocxLayoutPage page = layout.Pages.First(p => p.PlacedRelatedStories.Any(s =>
                s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote));
            if (page.Items.Count != 0) continue;
            witnessed.Add((ascent, custom, size, registered));
            DocxPlacedRelatedStoryLayout separator = page.PlacedRelatedStories.Single(s =>
                s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote &&
                s.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
            DocxTextLineLayout first = separator.StoryLayout.TextLines.Single();
            TestAssert.Equal(registered ? (double)size : 12d, first.Segments[0].StyleRun.FontSize);
            DocxTextRun mark = first.Segments[0].StyleRun;
            TestAssert.True(Math.Abs(separator.TopY - separator.Height -
                (page.Height - page.MarginTop - Math.Max(-first.BaselineY,
                    measurer.MeasureWindowsAscender(mark, mark.FontSize)))) < .000001d,
                "The head baseline preserves the source floor and the visible mark's measured first ascent.");
            var ends = layout.Pages.SelectMany(p => p.PlacedRelatedStories).Where(s =>
                s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote).ToArray();
            TestAssert.True(ends.All(s => s.UsesPlainDocumentEndnoteTypography && !s.UsesCompleteDocumentEndnoteFlow),
                "Source origin placement retains the existing mixed typography and paragraph-boundary policy.");
            TestAssert.True(ends.SelectMany(s => s.TextLines).Select(l => l.SourceLineIndex).SequenceEqual(
                layout.RelatedStories.Single(s => s.Story.IsNormalStoryType &&
                    s.Story.Kind == DocxRelatedStoryKind.Endnote).TextLines.Select(l => l.SourceLineIndex)),
                "Every original canonical line retains its source owner and order.");
        }
        TestAssert.Equal(24, witnessed.Count);
        foreach (double unavailable in new[] { 0d, -1d })
        {
            int fallbacks = 0;
            foreach (int footerLines in new[] { 1, 6 })
            foreach (int footRepeats in new[] { 96, 100, 104, 120 })
            {
                DocxDocument source = WithMixedNoteStaticFrame(ReadMixedNoteRetryFixture(true, 32, footRepeats),
                    false, true, footerLines);
                DocxLayout layout = engine.Create(source, new NoteContentMetricMeasurer(unavailable), default);
                DocxLayoutPage page = layout.Pages.First(p => p.PlacedRelatedStories.Any(s =>
                    s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote));
                if (page.Items.Count != 0) continue;
                fallbacks++;
                DocxPlacedRelatedStoryLayout separator = page.PlacedRelatedStories.Single(s =>
                    s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote &&
                    s.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
                TestAssert.Equal(page.Height - page.MarginTop + 3d, separator.TopY);
            }
            TestAssert.True(fallbacks > 0, "Unavailable separator metrics retain witnessed deferred-head placement.");
        }
    }

    public static void DocxDeferredMixedEndnoteConsumesSeparatorDeficitOnce()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        foreach (bool custom in new[] { false, true })
        foreach (bool matching in new[] { false, true })
        {
            int deferred = 0;
            foreach (int footerLines in new[] { 1, 6 })
            foreach (int footRepeats in new[] { 96, 100, 104, 120 })
            {
                var measurer = new NoteContentMetricMeasurer(10.8d, 1.15d,
                    matching ? 1.15d : 1.220703125d);
                DocxDocument source = WithMixedNoteStaticFrame(ReadMixedNoteRetryFixture(custom, 32, footRepeats),
                    false, true, footerLines);
                DocxLayout layout = engine.Create(source, measurer, default);
                DocxLayoutPage page = layout.Pages.First(p => p.PlacedRelatedStories.Any(s =>
                    s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote));
                if (page.Items.Count != 0) continue;
                deferred++;
                DocxPlacedRelatedStoryLayout end = page.PlacedRelatedStories.Single(s =>
                    s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote);
                double firstBaseline = end.TextLines[0].BaselineY;
                // Independent Word controls put the Calibri-content baseline
                // 25.22pt below the fresh frame's top with a 12pt Arial separator.
                // Matching metrics retain the separator's own minimum gap.
                double inset = page.Height - page.MarginTop - firstBaseline;
                TestAssert.True(Math.Abs(inset - (matching ? 25.08d : 25.22d)) < .02d,
                    $"The source-first-inset gap agrees with Word and preserves the mark-gap floor: {inset}, custom={custom}, matching={matching}.");
                DocxRelatedStoryLayout canonical = layout.RelatedStories.Single(s => s.Story.IsNormalStoryType &&
                    s.Story.Kind == DocxRelatedStoryKind.Endnote);
                TestAssert.True(layout.Pages.SelectMany(p => p.PlacedRelatedStories).Where(s =>
                    s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote)
                    .SelectMany(s => s.TextLines).Select(l => l.SourceLineIndex).SequenceEqual(
                        canonical.TextLines.Select(l => l.SourceLineIndex)),
                    "Gap correction preserves every canonical source owner and order.");
            }
            TestAssert.True(deferred > 0, "Both marker and matching/mixed metric profiles witness a deferred head.");
        }
    }

    public static void DocxFittingMixedEndnoteSharesOwnedSeparatorDeficit()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        foreach (bool custom in new[] { false, true })
        foreach (bool matching in new[] { false, true })
        foreach (int footerLines in new[] { 1, 6 })
        foreach (int footRepeats in new[] { 0, 16 })
        {
            var measurer = new NoteContentMetricMeasurer(10.8d, 1.15d,
                matching ? 1.15d : 1.220703125d, useStrikeoutMetrics: true);
            DocxDocument source = WithMixedNoteStaticFrame(ReadMixedNoteRetryFixture(custom, 32, footRepeats),
                false, true, footerLines);
            DocxLayout layout = engine.Create(source, measurer, default);
            DocxLayoutPage page = layout.Pages.Single(p => p.PlacedRelatedStories.Any(s =>
                s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote));
            TestAssert.True(page.Items.Count > 0 && page.PlacedRelatedStories.Any(s =>
                s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote),
                "The source witnesses a complete fitting mixed head on its occupied body/footnote page.");
            DocxPlacedRelatedStoryLayout end = page.PlacedRelatedStories.Single(s =>
                s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote);
            DocxPlacedRelatedStoryLayout separator = page.PlacedRelatedStories.Single(s =>
                s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote &&
                s.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
            // Independent Word controls put the 12pt Arial rule 16.44pt above
            // Calibri content. Matching metrics preserve the 16.31pt mark-gap floor.
            double expected = matching ? 16.31d : 16.44d;
            double actual = separator.SeparatorY!.Value - end.TextLines[0].BaselineY;
            TestAssert.True(Math.Abs(actual - expected) < .02d,
                $"Fitting first-bearing ownership consumes the deficit once: {actual}, custom={custom}, matching={matching}.");
            DocxRelatedStoryLayout canonical = layout.RelatedStories.Single(s =>
                s.Story.IsNormalStoryType && s.Story.Kind == DocxRelatedStoryKind.Endnote);
            TestAssert.True(end.TextLines.Select(l => l.SourceLineIndex).SequenceEqual(
                canonical.TextLines.Select(l => l.SourceLineIndex)),
                "The fitting gap retains the complete canonical source owner and order.");
        }
    }

    public static void DocxEndnoteSeparatorRuleFollowsStrikeoutMetrics()
    {
        // RV06 endnote probes match the footnote strikeout rule, so section-end endnote
        // separators share the same font-derived geometry instead of a fixed 3.74pt offset.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:endnotePr><w:pos w:val="sectEnd"/></w:endnotePr></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body</w:t></w:r></w:p></w:endnote></w:endnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        var resolver = new StrikeoutFontResolver();
        DocxFontPlan fontPlan = DocxFontPlan.Create(document, resolver, CancellationToken.None);
        var measurer = DocxFontPlanTextMeasurer.CreateAsync(fontPlan, resolver.Resolve(new FontRequest("StrikeFace")), CancellationToken.None, resolver).GetAwaiter().GetResult();
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, measurer, CancellationToken.None);
        DocxPlacedRelatedStoryLayout separator = layout.Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .Single(story => story.SeparatorY is not null);
        DocxTextLineLayout separatorLine = separator.TextLines.Single();
        double markSize = separatorLine.Segments[0].StyleRun.EffectiveProperties.FontSize;
        double ruleY = separator.SeparatorY ?? double.NaN;
        TestAssert.True(Math.Abs(ruleY - separatorLine.BaselineY - 0.25d * markSize) < 0.000001d, "Endnote rule bottom must sit 0.25em above the mark baseline with no ride.");
        TestAssert.True(Math.Abs(separator.SeparatorThickness - 0.05d * markSize) < 0.000001d, "Endnote rule thickness must follow the strikeout size.");
    }

    public static void DocxEndnoteSeparatorIsPlacedWithRuleMark()
    {
        // RV06 endnote probes: Office draws the endnote separator rule with a mark
        // space at the rule end, like footnotes.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:endnotePr><w:pos w:val="sectEnd"/></w:endnotePr></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body</w:t></w:r></w:p></w:endnote></w:endnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxPlacedRelatedStoryLayout separator = layout.Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .Single(story => story.SeparatorY is not null);
        DocxTextLineLayout separatorLine = separator.TextLines.Single();
        TestAssert.Equal(1, separatorLine.Segments.Count);
        TestAssert.Equal(" ", separatorLine.Segments[0].Text);
        TestAssert.Equal(144d, separatorLine.Segments[0].X - separatorLine.X);
    }

    public static void DocxDocumentEndnoteSeparatorIsPlacedWithRuleMark()
    {
        // RV06 endnote probes: document-end endnotes draw the separator rule with a
        // mark space at the rule end, like section-end endnotes.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body</w:t></w:r></w:p></w:endnote></w:endnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxPlacedRelatedStoryLayout separator = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .Single(story => story.SeparatorY is not null);
        DocxTextLineLayout separatorLine = separator.TextLines.Single();
        TestAssert.Equal(1, separatorLine.Segments.Count);
        TestAssert.Equal(" ", separatorLine.Segments[0].Text);
        TestAssert.Equal(144d, separatorLine.Segments[0].X - separatorLine.X);
    }

    public static void DocxEndnoteSeparatorIgnoresDirectParagraphSpacing()
    {
        // RV06 endnote-spacing probes (Word 16.0, edge-endsepsp/endsepspb): Office
        // holds document-end endnotes byte-identical across separator before/after
        // 0 vs 24pt, so separator stories lay out with latent-default spacing while
        // the renderer flows direct values into the separator height (24pt error).
        double topAfter0 = LayoutEndnoteBodyTopWithSeparatorSpacing("<w:pPr><w:spacing w:after=\"0\"/></w:pPr>");
        double topAfter24 = LayoutEndnoteBodyTopWithSeparatorSpacing("<w:pPr><w:spacing w:after=\"480\"/></w:pPr>");
        double topDefault = LayoutEndnoteBodyTopWithSeparatorSpacing(string.Empty);
        TestAssert.True(Math.Abs(topAfter0 - topAfter24) < 0.000001d, "Endnote separator after-spacing must not move the note block; observed shift=" + Math.Abs(topAfter0 - topAfter24).ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(topAfter0 - topDefault) < 0.000001d, "Endnote separator after-spacing must match latent defaults; observed shift=" + Math.Abs(topAfter0 - topDefault).ToString(CultureInfo.InvariantCulture) + ".");
    }

    public static void DocxEndnoteSeparatorIgnoresExactLineSpacing()
    {
        // RV06 exact-line probe (Word 16.0, edge-endsepspex): Office shifts the
        // endnote block identically with exact-24 and auto separator lines
        // (minus 8.76 both), so separator stories also ignore exact line rules
        // while the renderer clamps the mark line box to 24pt (shift minus 3.76).
        double topExact = LayoutEndnoteBodyTopWithSeparatorSpacing("<w:pPr><w:spacing w:line=\"480\" w:lineRule=\"exact\"/></w:pPr>");
        double topAuto = LayoutEndnoteBodyTopWithSeparatorSpacing(string.Empty);
        TestAssert.True(Math.Abs(topExact - topAuto) < 0.000001d, "Endnote separator exact line spacing must match auto lines; observed shift=" + Math.Abs(topExact - topAuto).ToString(CultureInfo.InvariantCulture) + ".");
    }

    public static void DocxEndnoteSeparatorGapUsesMarkMetrics()
    {
        // RV06 height-model probes (Word COM references edge-endsepheight-* plus the
        // 10/12/14pt mark grids): document-end content placement below the separator
        // grows with mark size (Office rule-to-first slope carries a size-driven
        // below-gap), while the renderer drops a legacy constant 3pt. A synthetic
        // singleLineEm of 1.5em makes the mark-metrics gap (1.5 minus 0.94) times the
        // 10pt mark size, pinning 5.6pt against the legacy constant. Pre-fix the
        // separator-to-content distance renders at 3pt.
        double gap = LayoutEndnoteSeparatorContentGapWithMarkSize(20, 1.5d);
        TestAssert.True(Math.Abs(gap - 5.6d) < 0.000001d, "Document-end below-separator gap must follow mark metrics.");
    }

    public static void DocxEndnoteSeparatorHeightGrowsWithFirstInset()
    {
        // RV06 height-model probes (empty/exact-24/text separator variants place
        // identically in Word): Office separator placement ignores laid-out line
        // boxes, growing only by the first-line inset with mark size. The laid-out
        // box slope would separate 10pt and 20pt marks by a full line box per 10pt;
        // the size-driven placement height separates them by one first inset.
        // Pre-fix the placed heights differ by the laid-out box slope.
        double shortHeight = LayoutEndnoteSeparatorHeightWithMarkSize(20);
        double tallHeight = LayoutEndnoteSeparatorHeightWithMarkSize(40);
        TestAssert.True(Math.Abs((tallHeight - shortHeight) - 9.4d) < 0.000001d, "Document-end separator placement height must grow by one first inset per 10pt.");
    }

    public static void DocxEndnoteSeparatorHeightCorrectsSingleLineDeficit()
    {
        // RV06 four-family separator grids (Word COM references edge-endsepgrid-cal/tmr/tah/vdn):
        // Office separator height carries a singleLine-deficit term against the validated
        // Calibri anchor (measured singleLineEm 1.2207): Times, Tahoma and Verdana share 1.15
        // yet place about 0.7pt lower at 12pt, which no laid-out box spread explains. A synthetic
        // 1.15em mark must stand 0.84 times the 0.0707 deficit times 12pt above the anchored
        // height. Pre-fix both heights are equal.
        double calHeight = LayoutEndnoteSeparatorHeightWithSingleLineEm(24, 1.2207d);
        double lowHeight = LayoutEndnoteSeparatorHeightWithSingleLineEm(24, 1.15d);
        TestAssert.True(Math.Abs((lowHeight - calHeight) - 0.71d) < 0.01d, "Document-end separator height must correct the singleLine deficit; observed shift=" + (lowHeight - calHeight).ToString(CultureInfo.InvariantCulture) + ".");
    }

    public static void DocxEndnoteSeparatorHeightAddsHheaAscenderExcess()
    {
        // RV06 endnote block-shift probes (Word COM references: five-family mark axis on pinned Tahoma content plus Tahoma-mark size sweep, genuine embeds both sides): Office document-end separator totals run about 0.8 taller on Tahoma/Verdana marks, flat across 8/10/14pt marks, while Calibri/Times/Georgia/Arial and the other probed families sit inside 0.25. The excess keys on hhea ascender over the auto-inset floor. A synthetic 1.005em mark on 1.005em content must stand (1.005 minus 0.94) times 1.27 times 10pt above the anchored height (mark and content terms meet in the minimum), while a 0.89em mark matches the 0.90em control exactly (gated at the floor). Pre-fix all three heights are equal.
        double controlHeight = LayoutEndnoteSeparatorHeightWithHheaAscender(20, 0.90d);
        double gatedHeight = LayoutEndnoteSeparatorHeightWithHheaAscender(20, 0.89d);
        double bigHeight = LayoutEndnoteSeparatorHeightWithHheaAscender(20, 1.005d);
        TestAssert.True(Math.Abs((bigHeight - controlHeight) - 0.83d) < 0.01d, "Document-end separator height must add the hhea ascender excess; observed shift=" + (bigHeight - controlHeight).ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(gatedHeight - controlHeight) < 0.000001d, "Document-end separator height must gate the hhea excess at the auto-inset floor.");
    }

    public static void DocxEndnoteAbsentSeparatorNormalizesToDefault()
    {
        // RV06 height-model probes (Word COM reference edge-endsepheight-absent):
        // Word normalizes separator-less endnote parts with a default separator
        // (absent renders identically to mark: same rule, same placement), while
        // the renderer skips placement entirely (content a full separator higher
        // with no rule). Pre-fix no separator story is placed. Footnote parts
        // keep legacy absent behavior.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with endnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body one</w:t></w:r></w:p><w:p><w:r><w:t>Note body two</w:t></w:r></w:p></w:endnote></w:endnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxPlacedRelatedStoryLayout separator = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
        TestAssert.True(separator.Height > 0d, "Synthesized default separator must take vertical space.");
        DocxPlacedRelatedStoryLayout content = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal));
        TestAssert.True(Math.Abs(((separator.TopY - separator.Height) - content.TopY) - 3d) < 0.000001d, "Default separator content must keep the legacy gap with the FamilyWidth measurer.");
    }

    public static void DocxEndnoteSeparatorHeightAddsSmallAscSupplement()
    {
        double control = LayoutEndnoteSeparatorHeightWithWindowsAscender(16, 0.9521d);
        double small = LayoutEndnoteSeparatorHeightWithWindowsAscender(16, 0.89d);
        TestAssert.True(Math.Abs((small - control) - 0.50d) < 0.01d, "Document-end separator height must add the small ascender supplement; observed shift=" + (small - control).ToString(CultureInfo.InvariantCulture) + ".");
    }

    public static void DocxEndnoteContentGapFollowsContentFont()
    {
        double control = LayoutEndnoteContentGapWithFamilies("Calibri", "Calibri");
        double mixed = LayoutEndnoteContentGapWithFamilies("Calibri", "Tahoma");
        TestAssert.True(Math.Abs(control - 3.37d) < 0.01d, "Uniform endnote gap must keep the shared value; control=" + control.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(mixed - 6.72d) < 0.01d, "Endnote gap must follow content font; mixed=" + mixed.ToString(CultureInfo.InvariantCulture) + ".");
    }

    public static void DocxEndnoteContentGapAddsWindowsAscenderExcess()
    {
        // RV06 fnwrap content-gap probes (Word 16.0): the shared helper serves
        // document-end endnotes too, so the same synthetic big-ascender excess
        // applies at 12pt while uniform content keeps the shared value.
        double control = LayoutEndnoteContentGapWithWindowsAscender("Calibri", "Calibri");
        double mixed = LayoutEndnoteContentGapWithWindowsAscender("Calibri", "BigAsc");
        TestAssert.True(Math.Abs(control - 3.37d) < 0.01d, "Uniform endnote gap must keep the shared value; control=" + control.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(mixed - 4.54d) < 0.01d, "Big-ascender endnote gap must add the Windows excess; mixed=" + mixed.ToString(CultureInfo.InvariantCulture) + ".");
    }

    public static void DocxEndnoteTakeExcludesTrailingAfterSpacing()
    {
        // RV06 after-spacing sweep: endnote takes exclude the take last-line
        // trailing after-spacing like footnote takes, so the head page takes one
        // more mixed line than the full-box fit allows at 16pt after-spacing.
        int headTake = LayoutEndnoteTakeWithAfterSpacingAndAfter(0, 320);
        TestAssert.Equal(22, headTake);
    }

    public static void DocxEndnoteContinuationTakeExcludesTrailingAfterSpacing()
    {
        // Same exclusion on endnote continuation takes with identical head takes
        // on both sides at 12pt after-spacing, isolating the continuation leg.
        int continuationTake = LayoutEndnoteTakeWithAfterSpacingAndAfter(1, 240);
        TestAssert.Equal(26, continuationTake);
    }

    public static void DocxEndnoteSeparatorTextMatchesMarkEmission()
    {
        // RV06 height-model probes (Word COM reference edge-endsepheight-text):
        // Office drops non-mark separator text while drawing the rule (text variant
        // renders the same 12 ops as the mark variant with no dashes); the renderer
        // emits the dashes as extra text shows. Pre-fix the text variant shows more.
        (FontFaceResolution Resolution, OpenTypeFont Font)? font = DocxTests.FindUsableInstalledFont();
        if (font is null)
        {
            TestAssert.Skip("Environmental precondition not met: (font is null)");
        }

        var resolver = new DocxTests.SingleResolutionFontResolver(font.Value.Resolution);
        int markShows = CountSeparatorVariantTextShows("<w:p><w:r><w:separator/></w:r></w:p>", resolver);
        int textShows = CountSeparatorVariantTextShows("<w:p><w:r><w:t xml:space=\"preserve\">---</w:t></w:r></w:p>", resolver);
        TestAssert.Equal(markShows, textShows);
    }
}
