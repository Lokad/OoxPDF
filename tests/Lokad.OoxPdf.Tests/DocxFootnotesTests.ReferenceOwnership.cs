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
    public static void DocxNumericNoteIdsKeepStoriesAndNavigationAcrossAliasSpellings()
    {
        XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        foreach (string fixture in new[] { "note-nav-two-footnotes", "note-nav-mixed-notes" })
        foreach (OoxPdfDocxMarkupMode mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string seed = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", fixture + ".docx"));
            var original = new Dictionary<string, byte[]>();
            using (ZipArchive zip = ZipFile.OpenRead(seed))
            foreach (ZipArchiveEntry entry in zip.Entries)
            {
                using Stream source = entry.Open();
                using var bytes = new MemoryStream();
                source.CopyTo(bytes); original[entry.FullName] = bytes.ToArray();
            }
            // Include a repeated reference; aliases must preserve its existing
            // label/navigation behavior as well as the distinct notes.
            XDocument body = XDocument.Parse(Encoding.UTF8.GetString(original["word/document.xml"]));
            XElement first = body.Descendants(word + "footnoteReference").First();
            first.Parent!.AddAfterSelf(new XElement(first.Parent!));
            original["word/document.xml"] = Encoding.UTF8.GetBytes(body.ToString());
            using var baselineInput = TestFixtures.CreateZipPackage(original);
            DocxDocument baseline = new DocxReader().Read(OoxPackage.Open(baselineInput, default), null, default, mode);
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            PdfPage[] baselinePages = renderer.RenderPagesAsync(baseline, null, default).ToBlockingEnumerable().ToArray();
            foreach (var spelling in new[] { (Declarations: false, References: true), (Declarations: true, References: false), (Declarations: true, References: true) })
            {
                var parts = new Dictionary<string, byte[]>(original);
                foreach (string name in new[] { "word/document.xml", "word/settings.xml", "word/footnotes.xml", "word/endnotes.xml" })
                {
                    if (!parts.TryGetValue(name, out byte[]? content)) continue;
                    XDocument xml = XDocument.Parse(Encoding.UTF8.GetString(content));
                    foreach (XElement element in xml.Descendants())
                    {
                        bool reference = element.Name == word + "footnoteReference" || element.Name == word + "endnoteReference" ||
                            (name == "word/settings.xml" && (element.Name == word + "footnote" || element.Name == word + "endnote"));
                        bool declaration = name is "word/footnotes.xml" or "word/endnotes.xml" &&
                            (element.Name == word + "footnote" || element.Name == word + "endnote");
                        if (((spelling.References && reference) || (spelling.Declarations && declaration)) && element.Attribute(word + "id") is { } id)
                            id.Value = id.Value.StartsWith('-') ? "-0" + id.Value[1..] : "0" + id.Value;
                    }
                    parts[name] = Encoding.UTF8.GetBytes(xml.ToString());
                }
                using var input = TestFixtures.CreateZipPackage(parts);
                DocxDocument candidate = new DocxReader().Read(OoxPackage.Open(input, default), null, default, mode);
                TestAssert.True(candidate.RelatedStories.Select(s => (s.Kind, s.Id, s.Type)).SequenceEqual(baseline.RelatedStories.Select(s => (s.Kind, s.Id, s.Type))),
                    "Numeric aliases must retain distinct notes and reserved-story identities.");
                PdfPage[] pages = renderer.RenderPagesAsync(candidate, null, default).ToBlockingEnumerable().ToArray();
                TestAssert.True(pages.Select(p => p.Content).SequenceEqual(baselinePages.Select(p => p.Content)), "Numeric aliases must retain every painted note.");
                TestAssert.True(pages.SelectMany(p => p.Annotations).SequenceEqual(baselinePages.SelectMany(p => p.Annotations)), "Numeric aliases must preserve navigation destinations.");
            }
        }
    }

    public static void DocxCustomNoteMarksHonorExplicitFalseFlags()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (string flag in new[] { "0", "false", "off", "FALSE", "OFF" })
        {
            DocxDocument document = ReadSectionNumberingFixture(kind, (xml, _, _) =>
                xml.Descendants(NoteWord + kind + "Reference").First().SetAttributeValue(NoteWord + "customMarkFollows", flag));
            CheckNoteLabels(document, ["iv", "v", "09", "10"]);
            TestAssert.Equal(flag, document.Paragraphs[0].InlineReferences[0].CustomMarkFollowsValue!);
            TestAssert.Equal(4, RenderCustomNoteLinks(document).Length);
        }
    }

    public static void DocxCustomNoteMarksLinkOnlyTheirFirstVisibleCharacter()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (OoxPdfDocxMarkupGeometryMode geometry in new[]
            { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
        {
            double? starWidth = null;
            foreach (string mark in new[] { "*", "**", "*tail", "AB", "\u2020" })
            {
                DocxDocument document = ReadCustomMarkFixture(kind, mark);
                PdfLinkAnnotation[] links = RenderCustomNoteLinks(document, geometry);
                TestAssert.Equal(4, links.Length);
                TestAssert.Equal(4, links.Select(l => l.Destination).Distinct().Count());
                TestAssert.True(links.All(l => l.Width > 0 && l.Height > 0 && l.IsDestination), "Custom marks require actual measurable destinations.");
                TestAssert.True(document.Paragraphs[0].InlineReferences[0].DisplayText is null, "Custom marks keep authored text without an automatic label.");
                TestAssert.Equal("iv", document.Paragraphs[1].InlineReferences[0].DisplayText!);
                TestAssert.True(document.Paragraphs[0].Runs.Any(r => r.Text == mark), "The complete custom-mark text remains visible.");
                if (mark == "*") starWidth = links[0].Width;
                if (mark is "**" or "*tail") TestAssert.True(Math.Abs(starWidth!.Value - links[0].Width) < .000001d, "The link covers only the first character.");
            }
        }
    }

    public static void DocxCustomNoteMarksRequireOwnedVisibleSourceAndTarget()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        {
            foreach (string mark in new[] { "", " ", "\ud83d\ude00" })
                TestAssert.Equal(3, RenderCustomNoteLinks(ReadCustomMarkFixture(kind, mark)).Length);
            DocxDocument source = ReadCustomMarkFixture(kind, "*");
            TestAssert.Equal(3, RenderCustomNoteLinks(source with
                { RelatedStories = source.RelatedStories.Where(s => s.Id != "37").ToArray() }).Length);
            DocxParagraph paragraph = source.Paragraphs[0];
            int sourceRun = paragraph.InlineReferences[0].SourceRunIndex;
            foreach (DocxParagraph hidden in new[]
                { paragraph with { Runs = paragraph.Runs.Where(r => r.SourceRunIndex != sourceRun).ToArray() },
                  paragraph with { Runs = paragraph.Runs.Select(r => r.SourceRunIndex == sourceRun ? r with { SourceRunIndex = 1000 } : r).ToArray() } })
            {
                DocxBodyElement[] elements = source.BodyElements.Select(e => e is DocxParagraphElement p && ReferenceEquals(p.Paragraph, paragraph)
                    ? (DocxBodyElement)new DocxParagraphElement(hidden) : e).ToArray();
                TestAssert.Equal(3, RenderCustomNoteLinks(source with { BodyElements = elements }).Length);
            }
            foreach (string flag in new[] { "unknown", "2" })
            {
                DocxDocument malformed = ReadSectionNumberingFixture(kind, (xml, _, _) =>
                    xml.Descendants(NoteWord + kind + "Reference").First().SetAttributeValue(NoteWord + "customMarkFollows", flag));
                TestAssert.True(malformed.Paragraphs[0].InlineReferences[0].DisplayText is null, "Malformed flags retain the prior fallback.");
                TestAssert.Equal(3, RenderCustomNoteLinks(malformed).Length);
            }
        }
    }

    public static void DocxCustomNoteMarksKeepRepeatedMarksAndDestinationsDistinct()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        {
            DocxDocument document = ReadSectionNumberingFixture(kind, (xml, _, parts) =>
            {
                foreach (XElement reference in xml.Descendants(NoteWord + kind + "Reference"))
                {
                    reference.SetAttributeValue(NoteWord + "customMarkFollows", "on");
                    reference.AddAfterSelf(new XElement(NoteWord + "t", "*"));
                }
                XDocument notes = XDocument.Parse(parts["word/" + kind + "s.xml"]);
                foreach (XElement marker in notes.Descendants(NoteWord + kind + "Ref")) { marker.Name = NoteWord + "t"; marker.Value = "*"; }
                parts["word/" + kind + "s.xml"] = notes.ToString();
            });
            PdfLinkAnnotation[] links = RenderCustomNoteLinks(document);
            TestAssert.Equal(4, links.Length);
            TestAssert.Equal(4, links.Select(l => l.Destination).Distinct().Count());
            for (int i = 0; i < links.Length; i++) TestAssert.Equal(kind == "endnote" ? 1 : i / 2, links[i].Destination!.Value.PageIndex);
        }
    }

    public static void DocxCustomNoteMarksOwnOffsetsWithinSharedSourceRuns()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        {
            DocxDocument document = ReadSectionNumberingFixture(kind, (xml, _, parts) =>
            {
                XElement[] references = xml.Descendants(NoteWord + kind + "Reference").Take(2).ToArray();
                XElement run = references[0].Parent!;
                references[0].SetAttributeValue(NoteWord + "customMarkFollows", "1");
                references[0].AddBeforeSelf(new XElement(NoteWord + "t", "pre "));
                references[0].AddAfterSelf(new XElement(NoteWord + "t", "*"));
                XElement second = new(references[1]);
                references[1].Parent!.Remove();
                second.SetAttributeValue(NoteWord + "customMarkFollows", "1");
                run.Add(new XElement(NoteWord + "t", " "), second, new XElement(NoteWord + "t", "#"));
                XDocument notes = XDocument.Parse(parts["word/" + kind + "s.xml"]);
                foreach (string id in new[] { "37", "4" })
                {
                    XElement marker = notes.Root!.Elements(NoteWord + kind).Single(n => (string?)n.Attribute(NoteWord + "id") == id)
                        .Descendants(NoteWord + kind + "Ref").Single();
                    marker.Name = NoteWord + "t"; marker.Value = id == "37" ? "*" : "#";
                }
                parts["word/" + kind + "s.xml"] = notes.ToString();
            });
            foreach (OoxPdfDocxMarkupGeometryMode geometry in new[]
                { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
            {
                PdfLinkAnnotation[] links = RenderCustomNoteLinks(document, geometry);
                TestAssert.Equal(4, links.Length);
                TestAssert.Equal(4, links.Select(l => l.Destination).Distinct().Count());
            }
            DocxDocument prefixed = ReadSectionNumberingFixture(kind, (xml, _, parts) =>
            {
                XElement reference = xml.Descendants(NoteWord + kind + "Reference").First();
                reference.SetAttributeValue(NoteWord + "customMarkFollows", "1");
                reference.AddBeforeSelf(new XElement(NoteWord + "t", "pre "));
                reference.AddAfterSelf(new XElement(NoteWord + "t", "*"));
                XDocument notes = XDocument.Parse(parts["word/" + kind + "s.xml"]);
                XElement marker = notes.Root!.Elements(NoteWord + kind).Single(n => (string?)n.Attribute(NoteWord + "id") == "37")
                    .Descendants(NoteWord + kind + "Ref").Single();
                marker.Name = NoteWord + "t"; marker.Value = "*";
                parts["word/" + kind + "s.xml"] = notes.ToString();
            });
            TestAssert.Equal(4, RenderCustomNoteLinks(prefixed).Length);
        }
    }

    private static PdfLinkAnnotation[] RenderCustomNoteLinks(DocxDocument document,
        OoxPdfDocxMarkupGeometryMode geometry = OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout) =>
        new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, geometry)
            .RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().SelectMany(p => p.Annotations).ToArray();

    private static DocxDocument ReadCustomMarkFixture(string kind, string mark) =>
        ReadSectionNumberingFixture(kind, (xml, _, parts) =>
        {
            XElement reference = xml.Descendants(NoteWord + kind + "Reference").First();
            reference.SetAttributeValue(NoteWord + "customMarkFollows", "1");
            reference.AddAfterSelf(new XElement(NoteWord + "t", mark));
            XDocument notes = XDocument.Parse(parts["word/" + kind + "s.xml"]);
            XElement marker = notes.Root!.Elements(NoteWord + kind).Single(n => (string?)n.Attribute(NoteWord + "id") == "37")
                .Descendants(NoteWord + kind + "Ref").Single();
            marker.Name = NoteWord + "t"; marker.Value = mark;
            parts["word/" + kind + "s.xml"] = notes.ToString();
        });

    public static void DocxCompleteFootnoteKerningProtectsOwnedCustomBaselinePrefixes()
    {
        foreach (string mark in new[] { "*", "**", "*tail" })
        {
            DocxDocument source = ReadCustomDocumentEndnoteFlowFixture(mark, footnote: true);
            DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default);
            TestAssert.True(DocxNoteKerning.TryCreateDocument(source, layout, out var adjusted, out var pages, default, DocxRelatedStoryKind.Footnote),
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
            TestAssert.True(DocxNoteKerning.HasStableOwners(final, pages, default, DocxRelatedStoryKind.Footnote), "Owned custom markers keep the complete note-page map.");
        }
    }

    public static void DocxCompleteFootnoteKerningRejectsAmbiguousCustomSourceOwners()
    {
        DocxDocument plain = ReadCustomDocumentEndnoteFlowFixture("*", footnote: true);
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
                { Runs = [first.Paragraphs[0].Runs[0] with { VerticalAlignmentValue = "superscript" }, first.Paragraphs[0].Runs[1]] })],
            } : s).ToArray() },
            plain with { RelatedStories = plain.RelatedStories.Select(s => ReferenceEquals(s, first) ? s with
            {
                BodyElements = [new DocxParagraphElement(first.Paragraphs[0] with
                { Runs = [first.Paragraphs[0].Runs[0] with { Text = "?" }, first.Paragraphs[0].Runs[1]] })],
            } : s).ToArray() },
        })
        {
            DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default);
            TestAssert.True(!DocxNoteKerning.TryCreateDocument(source, layout, out var adjusted, out _, default, DocxRelatedStoryKind.Footnote) && ReferenceEquals(source, adjusted),
                "Missing, repeated, overlapping and mismatched source owners keep the exact fallback.");
        }
    }

    private static DocxDocument ReadCustomDocumentEndnoteFlowFixture(string mark, bool footnote = false)
    {
        DocxDocument source = footnote ? WithPlainFootnoteTestHeader(ReadWrappedFootnoteFlowFixture(200)) :
            ReadDocumentEndnoteFlowFixture(0, 0);
        return source with
        {
            BodyElements = source.BodyElements.Select(e => e is DocxParagraphElement pe && pe.Paragraph.InlineReferences.Count > 0
                ? pe with { Paragraph = pe.Paragraph with
                {
                    Runs = pe.Paragraph.Runs.Select(r => pe.Paragraph.InlineReferences.Any(reference =>
                        r.SourceRunIndex == reference.SourceRunIndex && r.SourceTextOffsetInRun == reference.TextOffsetInRun)
                        ? r with { Text = mark } : r).ToArray(),
                    InlineReferences = pe.Paragraph.InlineReferences.Select(r => r with
                    { DisplayText = null, CustomMarkFollowsValue = "1", CustomMarkText = "*" }).ToArray(),
                } } : e).ToArray(),
            RelatedStories = source.RelatedStories.Select(s => s.IsNormalStoryType ? s with
            {
                BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                {
                    Runs = [p.Runs[0] with { Text = mark, VerticalAlignmentValue = "baseline", IsAutomaticNoteStoryMarker = false,
                        KerningMinimumHalfPointsValue = "0", PairKerningEnabled = true },
                        p.Runs[1] with { KerningMinimumHalfPointsValue = "0", PairKerningEnabled = true }],
                })).ToArray(),
            } : s).ToArray(),
        };
    }

    public static void DocxAutomaticNoteStoryMarkerRetainsExactXmlOwnership()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (bool field in new[] { false, true })
        {
            DocxDocument source = ReadSectionNumberingFixture(kind, (_, _, parts) =>
            {
                XDocument notes = XDocument.Parse(parts["word/" + kind + "s.xml"]);
                XElement mark = notes.Descendants(NoteWord + kind + "Ref").First();
                XElement run = mark.Parent!;
                XElement props = run.Element(NoteWord + "rPr")!;
                props.Element(NoteWord + "vertAlign")?.Remove();
                props.Add(new XElement(NoteWord + "vertAlign", new XAttribute(NoteWord + "val", "baseline")));
                mark.AddBeforeSelf(new XElement(NoteWord + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), "pre "));
                mark.AddAfterSelf(new XElement(NoteWord + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), " tail 105"));
                if (field)
                {
                    props.AddAfterSelf(new XElement(NoteWord + "fldChar", new XAttribute(NoteWord + "fldCharType", "begin")),
                        new XElement(NoteWord + "instrText", " REF markerRole "),
                        new XElement(NoteWord + "fldChar", new XAttribute(NoteWord + "fldCharType", "separate")));
                    run.Add(new XElement(NoteWord + "fldChar", new XAttribute(NoteWord + "fldCharType", "end")));
                }
                parts["word/" + kind + "s.xml"] = notes.ToString();
            });
            DocxRelatedStory story = source.RelatedStories.First(s => s.IsNormalStoryType);
            DocxTextRun marker = story.Paragraphs[0].Runs.Single(r => r.IsAutomaticNoteStoryMarker);
            string label = source.Paragraphs.SelectMany(p => p.InlineReferences).Single(r => r.Id == story.Id).DisplayText!;
            TestAssert.Equal(label, marker.Text);
            TestAssert.Equal(4, marker.SourceTextOffsetInRun);
            TestAssert.Equal(DocxRunVerticalAlignment.Baseline, marker.VerticalAlignment);
            DocxTextRun[] neighbors = story.Paragraphs[0].Runs.Where(r => r.SourceRunIndex == marker.SourceRunIndex && !r.IsAutomaticNoteStoryMarker).ToArray();
            TestAssert.Equal("pre  tail 105", string.Concat(neighbors.Select(r => r.Text)));
            TestAssert.True(neighbors.All(r => r.SourceTextOffsetInRun == 0 || r.SourceTextOffsetInRun == 4 + label.Length),
                "Only the explicit automatic element owns a marker role; adjacent text retains exact source offsets.");
            TestAssert.True(source.Paragraphs.SelectMany(p => p.Runs).All(r => !r.IsAutomaticNoteStoryMarker),
                "Related-story marker metadata does not change body-reference ownership.");
        }
        foreach (string kind in new[] { "footnote", "endnote" })
        {
            DocxDocument custom = ReadCustomMarkFixture(kind, "*");
            DocxRelatedStory story = custom.RelatedStories.Single(s => s.IsNormalStoryType && s.Id == "37");
            TestAssert.True(story.Paragraphs.SelectMany(p => p.Runs).All(r => !r.IsAutomaticNoteStoryMarker),
                "Authored custom text never inherits an automatic marker role from its style or value.");
        }
    }
}
