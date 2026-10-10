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

internal static class DocxFootnotesTests
{
    public static void DocxSimpleSectionPaginationUsesFollowingStartType()
    {
        foreach ((string? preceding, string? following, int expected) in new[]
        {
            ((string?)"nextPage", (string?)"continuous", 1),
            ((string?)"continuous", (string?)"nextPage", 2),
            ((string?)null, (string?)"continuous", 1),
            ((string?)"continuous", (string?)null, 2)
        })
        {
            DocxDocument document = ReadSectionNumberingFixture("footnote", (xml, _, parts) =>
            {
                SetSectionTypes(xml, preceding, following);
                SetNotePresence(xml, parts, "footnote", false, false);
            });
            TestAssert.Equal(expected, RenderSectionPages(document).Length);
            TestAssert.Equal(preceding, document.BodyElements.OfType<DocxSectionBreakElement>().Single().TypeValue?.ToValueString());
            TestAssert.Equal(following, document.FinalSectionBreak!.TypeValue?.ToValueString());
        }
    }

    public static void DocxSimpleSectionPaginationSeparatesPageBottomFootnotes()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach ((bool first, bool second) in new[] { (true, true), (true, false), (false, true), (false, false) })
        foreach (string? position in kind == "footnote" ? new string?[] { "pageBottom", "beneathText", null } : ["docEnd"])
        {
            DocxDocument document = ReadSectionNumberingFixture(kind, (xml, _, parts) =>
            {
                SetSectionTypes(xml, "continuous", "continuous");
                SetNotePresence(xml, parts, kind, first, second);
                foreach (XElement section in xml.Descendants(NoteWord + "sectPr"))
                {
                    XElement props = section.Element(NoteWord + kind + "Pr")!;
                    props.Elements(NoteWord + "pos").Remove();
                    if (position is not null) props.AddFirst(new XElement(NoteWord + "pos", new XAttribute(NoteWord + "val", position)));
                }
            });
            PdfPage[] pages = RenderSectionPages(document);
            bool splits = kind == "footnote" && first && position != "beneathText";
            TestAssert.Equal(splits ? 2 : 1, pages.Length);
            PdfLinkAnnotation[] links = pages.SelectMany(p => p.Annotations).ToArray();
            TestAssert.Equal((first ? 2 : 0) + (second ? 2 : 0), links.Length);
            TestAssert.True(links.All(l => l.IsDestination && l.Width > 0 && l.Height > 0), "Visible note markers remain linked.");
            if (splits && second)
            {
                TestAssert.Equal(2, pages[0].Annotations.Count);
                TestAssert.Equal(2, pages[1].Annotations.Count);
                TestAssert.True(pages[0].Annotations.All(l => l.Destination!.Value.PageIndex == 0) &&
                    pages[1].Annotations.All(l => l.Destination!.Value.PageIndex == 1), "Notes follow the section's source page.");
            }
        }
    }

    public static void DocxSimpleSectionPaginationIgnoresNotesAlreadyPlacedOnEarlierPages()
    {
        DocxDocument document = ReadSectionNumberingFixture("footnote", (xml, _, parts) =>
        {
            SetSectionTypes(xml, "continuous", "continuous");
            XElement[] paragraphs = xml.Root!.Element(NoteWord + "body")!.Elements(NoteWord + "p").ToArray();
            paragraphs[0].Add(new XElement(NoteWord + "r", new XElement(NoteWord + "br", new XAttribute(NoteWord + "type", "page"))));
            paragraphs[1].Descendants(NoteWord + "footnoteReference").Single().Parent!.Remove();
            XDocument notes = XDocument.Parse(parts["word/footnotes.xml"]);
            notes.Root!.Elements(NoteWord + "footnote").Single(n => (string?)n.Attribute(NoteWord + "id") == "4").Remove();
            parts["word/footnotes.xml"] = notes.ToString();
        });
        PdfPage[] pages = RenderSectionPages(document);
        TestAssert.Equal(2, pages.Length);
        TestAssert.Equal(1, pages[0].Annotations.Count);
        TestAssert.Equal(2, pages[1].Annotations.Count);
        TestAssert.True(pages[1].Annotations.All(l => l.Destination!.Value.PageIndex == 1), "An earlier note must not insert a third page.");
    }

    public static void DocxSimpleSectionPaginationRetainsExcludedFlow()
    {
        int variant = 0;
        foreach (Action<XDocument, XDocument, Dictionary<string, string>> modify in new Action<XDocument, XDocument, Dictionary<string, string>>[]
        {
            (xml, _, _) =>
            {
                XElement body = xml.Root!.Element(NoteWord + "body")!;
                XElement paragraph = body.Elements(NoteWord + "p").First();
                paragraph.Remove();
                body.AddFirst(new XElement(NoteWord + "tbl", new XElement(NoteWord + "tr", new XElement(NoteWord + "tc", paragraph))));
            },
            (xml, _, _) =>
            {
                foreach (XElement section in xml.Descendants(NoteWord + "sectPr"))
                    section.Add(new XElement(NoteWord + "cols", new XAttribute(NoteWord + "num", "2"), new XAttribute(NoteWord + "space", "720")));
            },
            (xml, _, _) => xml.Descendants(NoteWord + "sectPr").Last().Element(NoteWord + "pgMar")!.SetAttributeValue(NoteWord + "left", "1800"),
            (xml, _, _) => xml.Root!.Element(NoteWord + "body")!.Elements(NoteWord + "p").Skip(2).Remove(),
            (xml, _, _) => xml.Root!.Element(NoteWord + "body")!.Elements(NoteWord + "p").First().Add(
                new XElement(NoteWord + "ins", new XElement(NoteWord + "r", new XElement(NoteWord + "t", "tracked"))))
        })
        {
            DocxDocument document = ReadSectionNumberingFixture("footnote", (xml, settings, parts) =>
            {
                SetSectionTypes(xml, "continuous", "continuous");
                modify(xml, settings, parts);
            });
            int pages = RenderSectionPages(document).Length;
            // The synthetic face already needs two pages in the two-column case.
            int expected = variant == 1 ? 2 : 1;
            TestAssert.True(pages == expected, $"Excluded variant {variant} must retain {expected} pages; got {pages}.");
            variant++;
        }
        DocxDocument ordinary = ReadSectionNumberingFixture("footnote", (xml, _, _) => SetSectionTypes(xml, "continuous", "continuous"));
        var scaled = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, wordCompatiblePrintScale: 0.5d);
        TestAssert.Equal(1, scaled.Create(ordinary, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None).Pages.Count);
    }

    public static void DocxSimpleSectionPaginationPreservesLabelsAndRepeatedDestinations()
    {
        foreach (OoxPdfDocxMarkupGeometryMode geometry in new[]
            { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
        {
            DocxDocument document = ReadSectionNumberingFixture("footnote", (xml, _, _) =>
            {
                SetSectionTypes(xml, "continuous", "continuous");
                SetNoteSection(xml.Descendants(NoteWord + "sectPr").Last(), "footnote", "lowerRoman", "4", "eachSect");
            });
            CheckNoteLabels(document, ["iv", "v", "iv", "v"]);
            PdfPage[] pages = RenderSectionPages(document, geometry);
            PdfLinkAnnotation[] links = pages.SelectMany(p => p.Annotations).ToArray();
            TestAssert.Equal(2, pages.Length);
            TestAssert.Equal(4, links.Length);
            TestAssert.Equal(4, links.Select(l => l.Destination).Distinct().Count());
            TestAssert.Equal(105, document.Settings.FootnoteReferenceSettings.NumberStart!.Value);
            TestAssert.Equal(DocxSectionBreakType.Continuous, document.BodyElements.OfType<DocxSectionBreakElement>().Single().TypeValue!.Value);
        }
    }

    private static PdfPage[] RenderSectionPages(DocxDocument document,
        OoxPdfDocxMarkupGeometryMode geometry = OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup) =>
        new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, geometry)
            .RenderBlankPages(document, null, CancellationToken.None).ToArray();

    private static void SetSectionTypes(XDocument xml, string? preceding, string? following)
    {
        XElement[] sections = xml.Descendants(NoteWord + "sectPr").ToArray();
        for (int i = 0; i < sections.Length; i++)
        {
            sections[i].Elements(NoteWord + "type").Remove();
            string? value = i == 0 ? preceding : following;
            if (value is not null) sections[i].Add(new XElement(NoteWord + "type", new XAttribute(NoteWord + "val", value)));
        }
    }

    private static void SetNotePresence(XDocument xml, Dictionary<string, string> parts, string kind, bool first, bool second)
    {
        XElement[] references = xml.Descendants(NoteWord + kind + "Reference").ToArray();
        XDocument notes = XDocument.Parse(parts["word/" + kind + "s.xml"]);
        for (int i = 0; i < references.Length; i++)
        {
            if (i < 2 ? first : second) continue;
            string id = (string)references[i].Attribute(NoteWord + "id")!;
            references[i].Parent!.Remove();
            notes.Root!.Elements(NoteWord + kind).Single(n => (string?)n.Attribute(NoteWord + "id") == id).Remove();
        }
        parts["word/" + kind + "s.xml"] = notes.ToString();
    }

    public static void DocxSectionNoteNumbersUseClosingSectionPropertiesAndRestart()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        {
            DocxDocument document = ReadSectionNumberingFixture(kind);
            CheckNoteLabels(document, ["iv", "v", "09", "10"]);
            DocxNoteReferenceSettings authored = kind == "footnote"
                ? document.Settings.FootnoteReferenceSettings : document.Settings.EndnoteReferenceSettings;
            TestAssert.Equal(105, authored.NumberStart!.Value);
            TestAssert.Equal("decimal", authored.NumberFormatValue!);
        }
    }

    public static void DocxSectionNoteNumbersUseDocumentOccurrencesForContinuousStarts()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        {
            CheckNoteLabels(ReadSectionNumberingFixture(kind, (xml, _, _) =>
                SetNoteSection(xml.Descendants(NoteWord + "sectPr").Last(), kind, "decimalZero", "9", "continuous")),
                ["iv", "v", "11", "12"]);
            CheckNoteLabels(ReadSectionNumberingFixture(kind, (xml, _, parts) =>
            {
                SetNoteSection(xml.Descendants(NoteWord + "sectPr").Last(), kind, "lowerRoman", "4", "eachSect");
                AddThirdNoteSection(xml, parts, kind, "decimalZero", "9", "continuous");
            }), ["iv", "v", "iv", "v", "13", "14"]);
            CheckNoteLabels(ReadSectionNumberingFixture(kind, (xml, _, parts) =>
            {
                SetNoteSection(xml.Descendants(NoteWord + "sectPr").First(), kind, "lowerRoman", "4", "continuous");
                SetNoteSection(xml.Descendants(NoteWord + "sectPr").Last(), kind, "decimal", "9", "continuous");
                AddThirdNoteSection(xml, parts, kind, "upperRoman", "1", "continuous");
            }), ["iv", "v", "11", "12", "V", "VI"]);
        }
    }

    public static void DocxSectionNoteNumbersUseDefaultsDespiteDocumentRestartSettings()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        {
            DocxDocument document = ReadSectionNumberingFixture(kind, (xml, settings, _) =>
            {
                foreach (XElement section in xml.Descendants(NoteWord + "sectPr"))
                    section.Element(NoteWord + kind + "Pr")!.Remove();
                settings.Descendants(NoteWord + kind + "Pr").Single().Add(
                    new XElement(NoteWord + "numRestart", new XAttribute(NoteWord + "val", "eachSect")));
            });
            CheckNoteLabels(document, kind == "footnote" ? ["1", "2", "3", "4"] : ["i", "ii", "iii", "iv"]);
            DocxNoteReferenceSettings authored = kind == "footnote"
                ? document.Settings.FootnoteReferenceSettings : document.Settings.EndnoteReferenceSettings;
            TestAssert.Equal("eachSect", authored.NumberRestartValue!);
            TestAssert.Equal(105, authored.NumberStart!.Value);
        }
    }

    public static void DocxSectionNoteNumbersKeepIndependentFootnoteAndEndnoteAdmission()
    {
        DocxDocument document = ReadSectionNumberingFixture("footnote", (xml, settings, parts) =>
        {
            foreach (XElement reference in xml.Descendants(NoteWord + "footnoteReference").ToArray())
                reference.Parent!.AddAfterSelf(new XElement(NoteWord + "r", new XElement(NoteWord + "endnoteReference",
                    new XAttribute(NoteWord + "id", (string)reference.Attribute(NoteWord + "id")!))));
            foreach (XElement section in xml.Descendants(NoteWord + "sectPr"))
                SetNoteSection(section, "endnote", "upperRoman", "4", "eachPage");
            settings.Root!.Add(new XElement(NoteWord + "endnotePr",
                new XElement(NoteWord + "numFmt", new XAttribute(NoteWord + "val", "decimal")),
                new XElement(NoteWord + "numStart", new XAttribute(NoteWord + "val", "105"))));
            parts["word/endnotes.xml"] = parts["word/footnotes.xml"].Replace("footnote", "endnote", StringComparison.Ordinal)
                .Replace("Footnote", "Endnote", StringComparison.Ordinal);
            XDocument relationships = XDocument.Parse(parts["word/_rels/document.xml.rels"]);
            XElement relation = new(relationships.Root!.Elements().Single(e => ((string?)e.Attribute("Type"))?.EndsWith("/footnotes", StringComparison.Ordinal) == true));
            relation.SetAttributeValue("Id", "rIdEndnotes");
            relation.SetAttributeValue("Type", ((string)relation.Attribute("Type")!).Replace("footnotes", "endnotes", StringComparison.Ordinal));
            relation.SetAttributeValue("Target", "endnotes.xml");
            relationships.Root.Add(relation);
            parts["word/_rels/document.xml.rels"] = relationships.ToString();
            XDocument types = XDocument.Parse(parts["[Content_Types].xml"]);
            XElement type = new(types.Root!.Elements().Single(e => (string?)e.Attribute("PartName") == "/word/footnotes.xml"));
            type.SetAttributeValue("PartName", "/word/endnotes.xml");
            type.SetAttributeValue("ContentType", ((string)type.Attribute("ContentType")!).Replace("footnotes", "endnotes", StringComparison.Ordinal));
            types.Root.Add(type);
            parts["[Content_Types].xml"] = types.ToString();
        });
        CheckNoteLabels(document, ["iv", "105", "v", "106", "09", "107", "10", "108"]);
        TestAssert.Equal(8, document.RelatedStories.Count(s => s.Kind is DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote && s.Id is not "-1" and not "0"));
    }

    public static void DocxSectionNoteNumbersSurviveTablesRunBreaksAndEmptySections()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        {
            CheckNoteLabels(ReadSectionNumberingFixture(kind, (xml, _, _) =>
            {
                XElement body = xml.Root!.Element(NoteWord + "body")!;
                XElement paragraph = body.Elements(NoteWord + "p").First();
                paragraph.Remove();
                body.AddFirst(new XElement(NoteWord + "tbl", new XElement(NoteWord + "tr", new XElement(NoteWord + "tc", paragraph))));
                body.Elements(NoteWord + "p").First().Add(new XElement(NoteWord + "r", new XElement(NoteWord + "br", new XAttribute(NoteWord + "type", "page"))));
                body.AddFirst(new XElement(NoteWord + "p", new XElement(NoteWord + "pPr", new XElement(xml.Descendants(NoteWord + "sectPr").Last()))));
            }), ["iv", "v", "09", "10"]);
        }
    }

    public static void DocxSectionNoteNumbersKeepUnqualifiedFallbacks()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (Action<XDocument, XDocument, Dictionary<string, string>> modify in new Action<XDocument, XDocument, Dictionary<string, string>>[]
        {
            (xml, _, _) => SetNoteSection(xml.Descendants(NoteWord + "sectPr").Last(), kind, "decimal", "4", "eachPage"),
            (xml, _, _) => SetNoteSection(xml.Descendants(NoteWord + "sectPr").Last(), kind, "cardinalText", "4", "eachSect"),
            (xml, _, _) => SetNoteSection(xml.Descendants(NoteWord + "sectPr").Last(), kind, "decimal", "0", "eachSect"),
            (xml, _, _) => SetNoteSection(xml.Descendants(NoteWord + "sectPr").Last(), kind, "decimal", "32768", "eachSect"),
            (xml, _, _) => SetNoteSection(xml.Descendants(NoteWord + "sectPr").Last(), kind, "decimal", "broken", "eachSect"),
            (xml, _, _) => xml.Descendants(NoteWord + "sectPr").First().Add(new XElement(NoteWord + "sectPrChange", new XElement(NoteWord + "sectPr"))),
            (xml, _, _) => xml.Descendants(NoteWord + "body").Single().Elements(NoteWord + "p").First().Add(new XElement(NoteWord + "ins", new XElement(NoteWord + "r", new XElement(NoteWord + "t", "tracked")))),
            (_, settings, _) => settings.Descendants(NoteWord + kind + "Pr").Single().Add(new XElement(NoteWord + "numRestart", new XAttribute(NoteWord + "val", "eachPage")))
        })
            CheckNoteLabels(ReadSectionNumberingFixture(kind, modify), ["105", "106", "107", "108"]);
    }

    public static void DocxNoteHitAreaSpacingExtendsSingleLineParagraphs()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (bool custom in new[] { false, true })
        foreach (OoxPdfDocxMarkupGeometryMode geometry in new[]
            { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
        {
            double baseline = RenderCustomNoteLinks(ReadNoteHitAreaFixture(kind, custom, 0), geometry)[0].Height;
            foreach (int after in new[] { 6, 12, 24, 36 })
            {
                PdfLinkAnnotation[] links = RenderCustomNoteLinks(ReadNoteHitAreaFixture(kind, custom, after), geometry);
                TestAssert.Equal(4, links.Length);
                TestAssert.True(Math.Abs(links[0].Height - baseline - after) < .000001d, "One-line note hit areas include paragraph after-spacing.");
            }
        }
    }

    public static void DocxNoteHitAreaSpacingKeepsMultilineSlots()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (bool custom in new[] { false, true })
        foreach (bool before in new[] { false, true })
        {
            Action<XDocument> wrap = xml =>
            {
                XElement paragraph = xml.Root!.Element(NoteWord + "body")!.Elements(NoteWord + "p").First();
                XElement text = before ? paragraph.Descendants(NoteWord + "t").First() : paragraph.Descendants(NoteWord + "t").Last();
                text.Value = string.Join(' ', Enumerable.Repeat("Public wrapped paragraph text", 35));
            };
            double baseline = RenderCustomNoteLinks(ReadNoteHitAreaFixture(kind, custom, 0, wrap))[0].Height;
            double spaced = RenderCustomNoteLinks(ReadNoteHitAreaFixture(kind, custom, 24, wrap))[0].Height;
            TestAssert.True(Math.Abs(spaced - baseline) < .000001d, "Marks on any multiline paragraph retain a line-only slot.");
        }
    }

    public static void DocxNoteHitAreaSpacingIncludesExactLineSlots()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (bool custom in new[] { false, true })
        {
            Action<XDocument> exact = xml =>
            {
                XElement spacing = xml.Root!.Element(NoteWord + "body")!.Elements(NoteWord + "p").First().Element(NoteWord + "pPr")!.Element(NoteWord + "spacing")!;
                spacing.SetAttributeValue(NoteWord + "line", "480"); spacing.SetAttributeValue(NoteWord + "lineRule", "exact");
            };
            foreach (OoxPdfDocxMarkupGeometryMode geometry in new[]
                { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
            {
                double baseline = RenderCustomNoteLinks(ReadNoteHitAreaFixture(kind, custom, 0, exact), geometry)[0].Height;
                double spaced = RenderCustomNoteLinks(ReadNoteHitAreaFixture(kind, custom, 24, exact), geometry)[0].Height;
                TestAssert.True(Math.Abs(spaced - baseline - 24d) < .000001d, "Exact single-line slots also include after-spacing.");
            }
        }
    }

    public static void DocxNoteHitAreaSpacingRetainsExcludedParagraphs()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (int variant in Enumerable.Range(0, 6))
        {
            Action<XDocument> exclude = xml =>
            {
                XElement body = xml.Root!.Element(NoteWord + "body")!;
                XElement paragraph = body.Elements(NoteWord + "p").First();
                XElement props = paragraph.Element(NoteWord + "pPr")!;
                XElement spacing = props.Element(NoteWord + "spacing")!;
                if (variant == 0) spacing.SetAttributeValue(NoteWord + "beforeLines", "100");
                if (variant == 1) props.Add(new XElement(NoteWord + "contextualSpacing"));
                if (variant == 2) spacing.SetAttributeValue(NoteWord + "afterAutospacing", "1");
                if (variant == 3) spacing.SetAttributeValue(NoteWord + "afterLines", "100");
                if (variant == 4)
                    foreach (XElement section in xml.Descendants(NoteWord + "sectPr"))
                    { section.Element(NoteWord + "cols")?.Remove(); section.Add(new XElement(NoteWord + "cols", new XAttribute(NoteWord + "num", "2"))); }
                if (variant == 5)
                    paragraph.Add(new XElement(NoteWord + "r", new XElement(NoteWord + "br", new XAttribute(NoteWord + "type", "page"))));
            };
            double baseline = RenderCustomNoteLinks(ReadNoteHitAreaFixture(kind, false, 0, exclude))[0].Height;
            double spaced = RenderCustomNoteLinks(ReadNoteHitAreaFixture(kind, false, 24, exclude))[0].Height;
            TestAssert.True(Math.Abs(spaced - baseline) < .000001d, "Excluded spacing/flow settings retain the prior hit-area height; variant " + variant);
        }
    }

    private static DocxDocument ReadNoteHitAreaFixture(string kind, bool custom, int after, Action<XDocument>? modify = null) =>
        ReadSectionNumberingFixture(kind, (xml, _, parts) =>
        {
            XElement paragraph = xml.Root!.Element(NoteWord + "body")!.Elements(NoteWord + "p").First();
            XElement props = paragraph.Element(NoteWord + "pPr")!;
            XElement? spacing = props.Element(NoteWord + "spacing");
            if (spacing is null) { spacing = new XElement(NoteWord + "spacing"); props.Add(spacing); }
            spacing.SetAttributeValue(NoteWord + "after", (after * 20).ToString(CultureInfo.InvariantCulture));
            if (custom)
            {
                XElement reference = paragraph.Descendants(NoteWord + kind + "Reference").Single();
                reference.SetAttributeValue(NoteWord + "customMarkFollows", "1"); reference.AddAfterSelf(new XElement(NoteWord + "t", "*"));
                XDocument notes = XDocument.Parse(parts["word/" + kind + "s.xml"]);
                XElement marker = notes.Root!.Elements(NoteWord + kind).Single(n => (string?)n.Attribute(NoteWord + "id") == "37")
                    .Descendants(NoteWord + kind + "Ref").Single(); marker.Name = NoteWord + "t"; marker.Value = "*";
                parts["word/" + kind + "s.xml"] = notes.ToString();
            }
            modify?.Invoke(xml);
        });

    public static void DocxNoteBeforeSpacingOwnsOnlyTheUnconsumedGap()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (bool custom in new[] { false, true })
        foreach (OoxPdfDocxMarkupGeometryMode geometry in new[]
            { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
        foreach (int? precedingAfter in new int?[] { null, 6, 24, 36 })
        {
            PdfLinkAnnotation baseline = RenderCustomNoteLinks(ReadNoteBeforeSpacingFixture(kind, custom, 0, precedingAfter), geometry)[0];
            foreach (int before in new[] { 6, 12, 36 })
            {
                PdfLinkAnnotation[] links = RenderCustomNoteLinks(ReadNoteBeforeSpacingFixture(kind, custom, before, precedingAfter), geometry);
                TestAssert.Equal(4, links.Length);
                double ownedGap = Math.Max(0d, before - (precedingAfter ?? 0));
                TestAssert.True(Math.Abs(links[0].Height - baseline.Height - ownedGap) < .000001d,
                    "Word gives the note paragraph only the before-gap beyond its predecessor's after-spacing.");
                TestAssert.True(Math.Abs(links[0].Y + links[0].Height - baseline.Y - baseline.Height) < .000001d,
                    "The gap's top stays anchored while the line moves through before-spacing.");
                TestAssert.True(Math.Abs(links[0].Width - baseline.Width) < .000001d, "Spacing must not change horizontal marker ownership.");
            }
        }
    }

    public static void DocxNoteBeforeSpacingIncludesInheritedAndSectionSlots()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        {
            DocxDocument inherited = ReadSectionNumberingFixture(kind, (xml, _, parts) =>
            {
                XElement props = xml.Root!.Element(NoteWord + "body")!.Elements(NoteWord + "p").First().Element(NoteWord + "pPr")!;
                props.Element(NoteWord + "spacing")!.SetAttributeValue(NoteWord + "before", null);
                props.Element(NoteWord + "spacing")!.SetAttributeValue(NoteWord + "after", "480");
                props.AddFirst(new XElement(NoteWord + "pStyle", new XAttribute(NoteWord + "val", "PublicNoteGap")));
                XDocument styles = XDocument.Parse(parts["word/styles.xml"]);
                styles.Root!.Add(new XElement(NoteWord + "style", new XAttribute(NoteWord + "type", "paragraph"),
                    new XAttribute(NoteWord + "styleId", "PublicNoteGap"), new XElement(NoteWord + "basedOn", new XAttribute(NoteWord + "val", "Normal")),
                    new XElement(NoteWord + "pPr", new XElement(NoteWord + "spacing", new XAttribute(NoteWord + "before", "240")))));
                parts["word/styles.xml"] = styles.ToString();
            });
            PdfLinkAnnotation direct = RenderCustomNoteLinks(ReadNoteBeforeSpacingFixture(kind, false, 12, null))[0];
            PdfLinkAnnotation styled = RenderCustomNoteLinks(inherited)[0];
            TestAssert.True(Math.Abs(direct.Height - styled.Height) < .000001d && Math.Abs(direct.Y - styled.Y) < .000001d,
                "Inherited and direct before-spacing must resolve the same note slot.");
            DocxDocument section = ReadNoteHitAreaFixture(kind, false, 0, xml =>
            {
                XElement props = xml.Root!.Element(NoteWord + "body")!.Elements(NoteWord + "p").ElementAt(2).Element(NoteWord + "pPr")!;
                props.Element(NoteWord + "spacing")!.SetAttributeValue(NoteWord + "before", "240");
                props.Element(NoteWord + "spacing")!.SetAttributeValue(NoteWord + "after", "480");
            });
            PdfLinkAnnotation baseline = RenderCustomNoteLinks(ReadNoteHitAreaFixture(kind, false, 0))[2];
            PdfLinkAnnotation changed = RenderCustomNoteLinks(section)[2];
            TestAssert.True(Math.Abs(changed.Height - baseline.Height - 36d) < .000001d,
                "A new section owns its before-spacing without the preceding section's pending gap.");
        }
    }

    public static void DocxNoteBeforeSpacingIncludesExactLineSlots()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (bool custom in new[] { false, true })
        {
            Action<XDocument> exact = xml =>
            {
                XElement spacing = xml.Root!.Element(NoteWord + "body")!.Elements(NoteWord + "p").First().Element(NoteWord + "pPr")!.Element(NoteWord + "spacing")!;
                spacing.SetAttributeValue(NoteWord + "line", "480"); spacing.SetAttributeValue(NoteWord + "lineRule", "exact");
            };
            PdfLinkAnnotation baseline = RenderCustomNoteLinks(ReadNoteHitAreaFixture(kind, custom, 24, exact))[0];
            PdfLinkAnnotation changed = RenderCustomNoteLinks(ReadNoteHitAreaFixture(kind, custom, 24, xml =>
            {
                exact(xml); xml.Root!.Element(NoteWord + "body")!.Elements(NoteWord + "p").First().Element(NoteWord + "pPr")!
                    .Element(NoteWord + "spacing")!.SetAttributeValue(NoteWord + "before", "240");
            }))[0];
            TestAssert.True(Math.Abs(changed.Height - baseline.Height - 12d) < .000001d, "Exact line slots own the independent before-gap.");
        }
    }

    public static void DocxNoteBeforeSpacingRetainsComplexGapFallbacks()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (int variant in Enumerable.Range(0, 7))
        {
            Action<XDocument> guard = xml =>
            {
                XElement body = xml.Root!.Element(NoteWord + "body")!, paragraph = body.Elements(NoteWord + "p").First();
                XElement props = paragraph.Element(NoteWord + "pPr")!, spacing = props.Element(NoteWord + "spacing")!;
                spacing.SetAttributeValue(NoteWord + "before", "240");
                if (variant == 0) spacing.SetAttributeValue(NoteWord + "beforeAutospacing", "1");
                if (variant == 1) spacing.SetAttributeValue(NoteWord + "beforeLines", "100");
                if (variant == 2) props.Add(new XElement(NoteWord + "contextualSpacing"));
                if (variant == 3) paragraph.Descendants(NoteWord + "t").Last().Value = string.Join(' ', Enumerable.Repeat("Public wrapped paragraph text", 35));
                if (variant == 4) paragraph.Add(new XElement(NoteWord + "r", new XElement(NoteWord + "br", new XAttribute(NoteWord + "type", "page"))));
                if (variant is 5 or 6)
                {
                    var previous = new XElement(NoteWord + "p", new XElement(NoteWord + "pPr",
                        new XElement(NoteWord + "spacing", new XAttribute(NoteWord + "before", "0"), new XAttribute(NoteWord + "after", "120"))),
                        new XElement(NoteWord + "r", new XElement(NoteWord + "t", variant is 5 or 6 ?
                            string.Join(' ', Enumerable.Repeat("Public wrapped predecessor", 35)) : "Public preceding paragraph")));
                    if (variant == 6) previous.Element(NoteWord + "pPr")!.Add(new XElement(NoteWord + "contextualSpacing"));
                    paragraph.AddBeforeSelf(previous);
                }
            };
            PdfLinkAnnotation baseline = RenderCustomNoteLinks(ReadNoteHitAreaFixture(kind, false, 0, guard))[0];
            PdfLinkAnnotation spaced = RenderCustomNoteLinks(ReadNoteHitAreaFixture(kind, false, 24, guard))[0];
            TestAssert.True(Math.Abs(baseline.Height - spaced.Height) < .000001d,
                "Complex gap ownership retains the prior rectangle; variant " + variant);
        }
    }

    private static DocxDocument ReadNoteBeforeSpacingFixture(string kind, bool custom, int before, int? precedingAfter) =>
        ReadNoteHitAreaFixture(kind, custom, 24, xml =>
        {
            XElement paragraph = xml.Root!.Element(NoteWord + "body")!.Elements(NoteWord + "p").First();
            paragraph.Element(NoteWord + "pPr")!.Element(NoteWord + "spacing")!
                .SetAttributeValue(NoteWord + "before", (before * 20).ToString(CultureInfo.InvariantCulture));
            if (precedingAfter is { } after)
                paragraph.AddBeforeSelf(new XElement(NoteWord + "p", new XElement(NoteWord + "pPr", new XElement(NoteWord + "spacing",
                    new XAttribute(NoteWord + "before", "0"), new XAttribute(NoteWord + "after", (after * 20).ToString(CultureInfo.InvariantCulture)),
                    new XAttribute(NoteWord + "line", "240"), new XAttribute(NoteWord + "lineRule", "auto"))),
                    new XElement(NoteWord + "r", new XElement(NoteWord + "t", "Public preceding paragraph."))));
        });

    public static void DocxMultilineNoteBeforeSpacingExtendsOnlyFirstLineMarks()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (bool custom in new[] { false, true })
        foreach (OoxPdfDocxMarkupGeometryMode geometry in new[]
            { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
        {
            PdfLinkAnnotation baseline = RenderCustomNoteLinks(ReadMultilineNoteSpacingFixture(kind, custom, 0, 0, true), geometry)[0];
            foreach (int before in new[] { 6, 12, 24 })
            foreach (int after in new[] { 0, 24 })
            {
                PdfLinkAnnotation[] links = RenderCustomNoteLinks(ReadMultilineNoteSpacingFixture(kind, custom, before, after, true), geometry);
                TestAssert.Equal(4, links.Length);
                TestAssert.True(Math.Abs(links[0].Height - baseline.Height - before) < .000001d,
                    "A first-line note mark owns before-spacing and keeps line-only after-spacing when the paragraph wraps.");
                TestAssert.True(Math.Abs(links[0].Y + links[0].Height - baseline.Y - baseline.Height) < .000001d,
                    "The first-line rectangle stays anchored to the before-gap's top.");
                TestAssert.True(Math.Abs(links[0].Width - baseline.Width) < .000001d, "Wrapping does not change marker width ownership.");
            }
        }
    }

    public static void DocxMultilineNoteBeforeSpacingKeepsLaterLineSlots()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (bool custom in new[] { false, true })
        {
            double baseline = RenderCustomNoteLinks(ReadMultilineNoteSpacingFixture(kind, custom, 0, 0, false))[0].Height;
            foreach (int before in new[] { 6, 12, 24 })
            foreach (int after in new[] { 0, 24 })
                TestAssert.True(Math.Abs(RenderCustomNoteLinks(ReadMultilineNoteSpacingFixture(kind, custom, before, after, false))[0].Height - baseline) < .000001d,
                    "A note mark on a later rendered line owns neither the paragraph's before-gap nor after-spacing.");
        }
    }

    public static void DocxMultilineNoteBeforeSpacingIncludesExactAndCollapsedGaps()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (bool custom in new[] { false, true })
        foreach (bool exact in new[] { false, true })
        {
            Action<XDocument> modify = xml =>
            {
                XElement paragraph = xml.Root!.Element(NoteWord + "body")!.Elements(NoteWord + "p").First();
                if (exact)
                {
                    XElement spacing = paragraph.Element(NoteWord + "pPr")!.Element(NoteWord + "spacing")!;
                    spacing.SetAttributeValue(NoteWord + "line", "480"); spacing.SetAttributeValue(NoteWord + "lineRule", "exact");
                }
                else paragraph.AddBeforeSelf(new XElement(NoteWord + "p", new XElement(NoteWord + "pPr", new XElement(NoteWord + "spacing",
                    new XAttribute(NoteWord + "before", "0"), new XAttribute(NoteWord + "after", "120"),
                    new XAttribute(NoteWord + "line", "240"), new XAttribute(NoteWord + "lineRule", "auto"))),
                    new XElement(NoteWord + "r", new XElement(NoteWord + "t", "Public preceding paragraph."))));
            };
            PdfLinkAnnotation baseline = RenderCustomNoteLinks(ReadMultilineNoteSpacingFixture(kind, custom, 0, 24, true, modify))[0];
            PdfLinkAnnotation changed = RenderCustomNoteLinks(ReadMultilineNoteSpacingFixture(kind, custom, 12, 24, true, modify))[0];
            TestAssert.True(Math.Abs(changed.Height - baseline.Height - (exact ? 12d : 6d)) < .000001d,
                "Exact first-line slots and unconsumed adjacent gaps follow the same first-line ownership rule.");
        }
    }

    public static void DocxMultilineNoteBeforeSpacingRetainsExcludedLayouts()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (int variant in Enumerable.Range(0, 6))
        {
            Action<XDocument> guard = xml =>
            {
                XElement paragraph = xml.Root!.Element(NoteWord + "body")!.Elements(NoteWord + "p").First();
                XElement props = paragraph.Element(NoteWord + "pPr")!, spacing = props.Element(NoteWord + "spacing")!;
                if (variant == 0)
                {
                    props.Add(new XElement(NoteWord + "contextualSpacing"));
                    paragraph.AddBeforeSelf(new XElement(NoteWord + "p", new XElement(NoteWord + "pPr",
                        new XElement(NoteWord + "spacing", new XAttribute(NoteWord + "after", "120"))),
                        new XElement(NoteWord + "r", new XElement(NoteWord + "t", "Public preceding paragraph."))));
                }
                if (variant == 1) spacing.SetAttributeValue(NoteWord + "beforeAutospacing", "1");
                if (variant == 2) spacing.SetAttributeValue(NoteWord + "beforeLines", "100");
                if (variant == 3) paragraph.Add(new XElement(NoteWord + "r", new XElement(NoteWord + "br", new XAttribute(NoteWord + "type", "page"))));
                if (variant == 4)
                    foreach (XElement section in xml.Descendants(NoteWord + "sectPr"))
                    { section.Element(NoteWord + "cols")?.Remove(); section.Add(new XElement(NoteWord + "cols", new XAttribute(NoteWord + "num", "2"))); }
                if (variant == 5) paragraph.AddBeforeSelf(new XElement(NoteWord + "p", new XElement(NoteWord + "pPr",
                    new XElement(NoteWord + "contextualSpacing"), new XElement(NoteWord + "spacing", new XAttribute(NoteWord + "after", "120"))),
                    new XElement(NoteWord + "r", new XElement(NoteWord + "t",
                        "Public preceding paragraph. " + string.Concat(Enumerable.Repeat("wrapped words ", 35))))));
            };
            double baseline = RenderCustomNoteLinks(ReadMultilineNoteSpacingFixture(kind, false, 0, 24, true, guard))[0].Height;
            double changed = RenderCustomNoteLinks(ReadMultilineNoteSpacingFixture(kind, false, 12, 24, true, guard))[0].Height;
            TestAssert.True(Math.Abs(changed - baseline) < .000001d, "Unqualified multiline gap settings retain the prior slot; variant " + variant);
        }
    }

    private static DocxDocument ReadMultilineNoteSpacingFixture(string kind, bool custom, int before, int after, bool first,
        Action<XDocument>? modify = null) => ReadNoteHitAreaFixture(kind, custom, after, xml =>
        {
            XElement paragraph = xml.Root!.Element(NoteWord + "body")!.Elements(NoteWord + "p").First();
            paragraph.Element(NoteWord + "pPr")!.Element(NoteWord + "spacing")!
                .SetAttributeValue(NoteWord + "before", (before * 20).ToString(CultureInfo.InvariantCulture));
            XElement text = first ? paragraph.Descendants(NoteWord + "t").Last() : paragraph.Descendants(NoteWord + "t").First();
            text.Value = string.Join(' ', Enumerable.Repeat("Public wrapped paragraph text", 35));
            modify?.Invoke(xml);
        });

    public static void DocxMultilineContextualNoteBeforeSpacingOwnsResolvedGaps()
    {
        CheckMultilineContextualNoteGaps(sameStyle: true);
    }

    public static void DocxMultilineContextualNoteBeforeSpacingKeepsDifferentStyleOwnership()
    {
        CheckMultilineContextualNoteGaps(sameStyle: false);
    }

    private static void CheckMultilineContextualNoteGaps(bool sameStyle)
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (bool custom in new[] { false, true })
        foreach (bool exact in new[] { false, true })
        foreach (OoxPdfDocxMarkupGeometryMode geometry in new[]
            { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
        {
            double lineHeight = RenderCustomNoteLinks(ReadMultilineContextualNoteFixture(kind, custom, 0, 0, 0, true, sameStyle, exact), geometry)[0].Height;
            foreach ((int before, int previousAfter) in new[] { (12, 6), (6, 12), (24, 6) })
            foreach (int after in new[] { 0, 24 })
            {
                DocxDocument document = ReadMultilineContextualNoteFixture(kind, custom, before, after, previousAfter, true, sameStyle, exact);
                PdfLinkAnnotation[] links = RenderCustomNoteLinks(document, geometry);
                TestAssert.Equal(4, links.Length);
                TestAssert.True(Math.Abs(links[0].Height - lineHeight - Math.Max(0, before - previousAfter)) < .000001d,
                    "A first-line multiline mark owns only its unconsumed before-gap, including contextual predecessor suppression.");
                TestAssert.Equal((double)before, document.Paragraphs[1].SpacingBeforePoints);
                TestAssert.Equal((double)after, document.Paragraphs[1].SpacingAfterPoints);
                TestAssert.Equal((double)previousAfter, document.Paragraphs[0].SpacingAfterPoints);
            }
        }
    }

    public static void DocxMultilineContextualNoteBeforeSpacingKeepsLaterLineSlots()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (bool custom in new[] { false, true })
        foreach (bool exact in new[] { false, true })
        foreach (bool sameStyle in new[] { false, true })
        {
            double baseline = RenderCustomNoteLinks(ReadMultilineContextualNoteFixture(kind, custom, 0, 0, 0, false, sameStyle, exact))[0].Height;
            foreach (int before in new[] { 6, 12, 24 })
            {
                PdfLinkAnnotation mark = RenderCustomNoteLinks(ReadMultilineContextualNoteFixture(kind, custom, before, 24, 6, false, sameStyle, exact))[0];
                TestAssert.True(Math.Abs(mark.Height - baseline) < .000001d,
                    "Later-line marks retain their line slot after contextual predecessors.");
            }
        }
    }

    public static void DocxMultilineContextualNoteBeforeSpacingRetainsExcludedPredecessors()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (int variant in new[] { 0, 1, 2, 3 })
        {
            Action<XDocument> guard = xml =>
            {
                XElement previous = xml.Root!.Element(NoteWord + "body")!.Elements(NoteWord + "p").First();
                if (variant == 0) previous.Descendants(NoteWord + "t").First().Value = "Public previous text " + string.Concat(Enumerable.Repeat("wrapped words ", 35));
                if (variant == 1) previous.Element(NoteWord + "pPr")!.Element(NoteWord + "spacing")!.SetAttributeValue(NoteWord + "afterAutospacing", "1");
                if (variant == 2) previous.Elements(NoteWord + "r").Remove();
                if (variant == 3) previous.ElementsAfterSelf(NoteWord + "p").First().Element(NoteWord + "pPr")!.Add(new XElement(NoteWord + "contextualSpacing"));
            };
            double baseline = RenderCustomNoteLinks(ReadMultilineContextualNoteFixture(kind, false, 0, 24, 6, true, true, false, guard))[0].Height;
            double spaced = RenderCustomNoteLinks(ReadMultilineContextualNoteFixture(kind, false, 12, 24, 6, true, true, false, guard))[0].Height;
            TestAssert.True(Math.Abs(baseline - spaced) < .000001d,
                $"Excluded predecessor variant {variant} ({kind}) retains its line slot: baseline={baseline}, spaced={spaced}.");
        }
    }

    private static DocxDocument ReadMultilineContextualNoteFixture(string kind, bool custom, int before, int after,
        int previousAfter, bool first, bool sameStyle, bool exact, Action<XDocument>? modify = null)
    {
        DocxDocument document = ReadMultilineNoteSpacingFixture(kind, custom, before, after, first, xml =>
        {
            XElement paragraph = xml.Root!.Element(NoteWord + "body")!.Elements(NoteWord + "p").First();
            if (exact)
            {
                XElement spacing = paragraph.Element(NoteWord + "pPr")!.Element(NoteWord + "spacing")!;
                spacing.SetAttributeValue(NoteWord + "line", "480"); spacing.SetAttributeValue(NoteWord + "lineRule", "exact");
            }
            paragraph.AddBeforeSelf(new XElement(NoteWord + "p", new XElement(NoteWord + "pPr",
                new XElement(NoteWord + "spacing", new XAttribute(NoteWord + "before", "0"), new XAttribute(NoteWord + "after", previousAfter * 20),
                    new XAttribute(NoteWord + "line", "240"), new XAttribute(NoteWord + "lineRule", "auto")),
                new XElement(NoteWord + "contextualSpacing")),
                new XElement(NoteWord + "r", new XElement(NoteWord + "t", "Public preceding paragraph."))));
            modify?.Invoke(xml);
        });
        if (sameStyle) return document;
        DocxParagraph previous = document.Paragraphs[0];
        DocxParagraph different = previous with { StyleId = "PublicPrevious" };
        return document with
        {
            BodyElements = document.BodyElements.Select(e => e is DocxParagraphElement p && ReferenceEquals(p.Paragraph, previous)
                ? (DocxBodyElement)new DocxParagraphElement(different) : e).ToArray()
        };
    }

    public static void DocxImplicitContextualStyleMatchesIndependentWordFlags()
    {
        foreach (string mode in new[] { "no-paragraph-styles", "off" })
        foreach (bool explicitPrevious in new[] { true, false })
        foreach ((bool previousFlag, bool currentFlag, double gap) in new[]
            { (false, false, 12d), (false, true, 6d), (true, false, 6d), (true, true, 0d) })
        {
            DocxDocument source = ReadDefaultContextualStyleFixture("footnote", "Normal", explicitPrevious, previousFlag, currentFlag, mode);
            DocxParagraph previous = source.Paragraphs[0] with
            { Runs = [new DocxTextRun("Previous", 10d, null, false, false, false, null, null)], LineSpacingPoints = 10d, InlineReferences = [] };
            DocxParagraph current = source.Paragraphs[1] with
            { Runs = [new DocxTextRun("Current", 10d, null, false, false, false, null, null)],
              SpacingAfterPoints = 0d, LineSpacingPoints = 10d, InlineReferences = [] };
            DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(previous), new DocxParagraphElement(current)], []);
            DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
                .Pages[0].Items.OfType<DocxTextLineLayout>().ToArray();
            TestAssert.True(Math.Abs(lines[0].BaselineY - lines[1].BaselineY - 10d - gap) < .000001d,
                "Office's implicit Normal identity follows the independent flag matrix in either direction.");
        }
    }

    public static void DocxImplicitContextualStyleAlignsAutomaticAndExactNoteBounds()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (string mode in new[] { "no-paragraph-styles", "off" })
        foreach (bool explicitPrevious in new[] { true, false })
        foreach (bool exact in new[] { true, false })
        {
            DocxDocument source = ReadDefaultContextualStyleFixture(kind, "Normal", explicitPrevious, true, true, mode);
            DocxDocument baseline = ReadDefaultContextualStyleFixture(kind, "Normal", explicitPrevious, true, true, "allomitted");
            if (exact)
            {
                DocxDocument SetExact(DocxDocument document) => document with
                { BodyElements = document.BodyElements.Select(e => e is DocxParagraphElement p && p.Paragraph.InlineReferences.Count > 0
                    ? new DocxParagraphElement(p.Paragraph with { LineSpacingPoints = 24d,
                        Spacing = p.Paragraph.Spacing with { LineRuleValue = "exact", LineValue = "480" } }) : e).ToArray() };
                source = SetExact(source); baseline = SetExact(baseline);
            }
            PdfLinkAnnotation[] actual = RenderCustomNoteLinks(source), expected = RenderCustomNoteLinks(baseline);
            TestAssert.Equal(expected.Length, actual.Length);
            for (int i = 0; i < actual.Length; i++)
            {
                TestAssert.True(Math.Abs(actual[i].Height - expected[i].Height) < .000001d,
                    "Implicit Normal owns the same first-line note gap as the same-style control.");
                TestAssert.True(Math.Abs(actual[i].Y - expected[i].Y) < .000001d,
                    "Implicit Normal keeps each note link on its painted automatic/exact slot.");
            }
        }
    }

    public static void DocxImplicitContextualStylePreservesAuthoredProperties()
    {
        foreach (string mode in new[] { "no-paragraph-styles", "off" })
        {
            DocxDocument source = ReadDefaultContextualStyleFixture("footnote", "Normal", true, true, true, mode);
            TestAssert.Equal("Normal", source.Paragraphs[0].StyleId!);
            TestAssert.True(source.Paragraphs[1].StyleId is null && source.Paragraphs[1].StyleResolution.StyleId is null,
                "Application-default identity does not invent an authored style.");
            TestAssert.Equal(12d, source.Paragraphs[1].SpacingBeforePoints);
            TestAssert.Equal(24d, source.Paragraphs[1].SpacingAfterPoints);
            TestAssert.Equal(mode == "off", source.Paragraphs[0].StyleResolution.StyleFound);
            TestAssert.Equal(mode == "off" ? 1 : 0, source.StyleCatalog.ParagraphStyles.Count);
        }
    }

    public static void DocxImplicitContextualStyleRetainsDeclaredAndAmbiguousPrecedence()
    {
        DocxDocument source = ReadDefaultContextualStyleFixture("footnote", "PublicBase", true, true, true);
        DocxParagraph previous = source.Paragraphs[0] with { StyleId = "Normal" };
        DocxParagraph current = source.Paragraphs[1];
        TestAssert.True(!DocxLayoutEngine.HasSameContextualSpacingStyle(previous, current),
            "An explicit named default takes precedence over the implicit Normal identity.");
        DocxDocument ambiguous = ReadDefaultContextualStyleFixture("footnote", "Normal", true, true, true, "multiple");
        TestAssert.True(!DocxLayoutEngine.HasSameContextualSpacingStyle(ambiguous.Paragraphs[0], ambiguous.Paragraphs[1]),
            "Multiple declarations keep the prior raw-identifier fallback.");
        DocxDocument nondefault = ReadDefaultContextualStyleFixture("footnote", "Normal", true, true, true, "nondefault");
        TestAssert.True(!DocxLayoutEngine.HasSameContextualSpacingStyle(nondefault.Paragraphs[0], nondefault.Paragraphs[1]),
            "The implicit fallback does not equate another explicit style with Normal.");
    }


    public static void DocxDefaultContextualStyleMatchesWordGapMatrix()
    {
        foreach (string defaultId in new[] { "Normal", "PublicBase" })
        foreach (bool explicitPrevious in new[] { true, false })
        foreach ((bool previousFlag, bool currentFlag, double largerBefore, double largerAfter) in new[]
        { (false, false, 12d, 12d), (false, true, 6d, 12d), (true, false, 6d, 0d), (true, true, 0d, 0d) })
        foreach (bool beforeLarger in new[] { true, false })
        {
            DocxDocument source = ReadDefaultContextualStyleFixture("footnote", defaultId, explicitPrevious, previousFlag, currentFlag);
            DocxParagraph previous = source.Paragraphs[0] with
            { Runs = [new DocxTextRun("Previous", 10d, null, false, false, false, null, null)],
              SpacingAfterPoints = beforeLarger ? 6d : 12d, LineSpacingPoints = 10d, InlineReferences = [] };
            DocxParagraph current = source.Paragraphs[1] with
            { Runs = [new DocxTextRun("Current", 10d, null, false, false, false, null, null)],
              SpacingBeforePoints = beforeLarger ? 12d : 6d, SpacingAfterPoints = 0d, LineSpacingPoints = 10d, InlineReferences = [] };
            DocxDocument document = DocxTests.CreateLayoutTestDocument(
                [new DocxParagraphElement(previous), new DocxParagraphElement(current)], []);
            DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
                .Pages[0].Items.OfType<DocxTextLineLayout>().ToArray();
            double expected = beforeLarger ? largerBefore : largerAfter;
            TestAssert.True(Math.Abs(lines[0].BaselineY - lines[1].BaselineY - 10d - expected) < .000001d,
                "Declared default styles follow Word's ownership matrix in either direction.");
            TestAssert.Equal(beforeLarger ? 6d : 12d, lines[1].PendingAfterSpacing ?? -1d);
            TestAssert.Equal(beforeLarger ? 12d : 6d, lines[1].ParagraphBeforeSpacing ?? -1d);
        }
    }

    public static void DocxDefaultContextualStylePreservesAuthoredIdentity()
    {
        foreach (string defaultId in new[] { "Normal", "PublicBase" })
        foreach (bool explicitPrevious in new[] { true, false })
        {
            DocxDocument source = ReadDefaultContextualStyleFixture("footnote", defaultId, explicitPrevious, true, true);
            DocxParagraph omitted = source.Paragraphs[explicitPrevious ? 1 : 0];
            DocxParagraph explicitStyle = source.Paragraphs[explicitPrevious ? 0 : 1];
            TestAssert.True(omitted.StyleId is null && omitted.StyleResolution.StyleId is null && !omitted.StyleResolution.StyleFound,
                "The authored missing style and its existing cascade provenance remain intact.");
            TestAssert.Equal(defaultId, explicitStyle.StyleId!);
            TestAssert.True(explicitStyle.StyleResolution.StyleFound, "Explicit style provenance remains resolved.");
            TestAssert.Equal(12d, source.Paragraphs[1].SpacingBeforePoints);
            TestAssert.Equal(24d, source.Paragraphs[1].SpacingAfterPoints);
            TestAssert.True(source.Paragraphs[1].Runs.All(r => r.FontSize <= 12d),
                "Identity resolution does not introduce the default style's paragraph or run overrides.");
        }
    }

    public static void DocxDefaultContextualStyleAlignsNoteBoundsWithOmittedControls()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (string defaultId in new[] { "Normal", "PublicBase" })
        foreach (bool explicitPrevious in new[] { true, false })
        foreach (OoxPdfDocxMarkupGeometryMode geometry in new[]
            { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
        {
            DocxDocument source = ReadDefaultContextualStyleFixture(kind, defaultId, explicitPrevious, true, true);
            DocxDocument omitted = ReadDefaultContextualStyleFixture(kind, defaultId, explicitPrevious, true, true, "allomitted");
            PdfLinkAnnotation[] expected = RenderCustomNoteLinks(omitted, geometry);
            PdfLinkAnnotation[] actual = RenderCustomNoteLinks(source, geometry);
            TestAssert.Equal(expected.Length, actual.Length);
            for (int i = 0; i < actual.Length; i++)
            {
                TestAssert.True(Math.Abs(expected[i].Height - actual[i].Height) < .000001d,
                    "Default-style aliases own the same note hit-area gap as omitted-style controls.");
                TestAssert.True(Math.Abs(expected[i].Y - actual[i].Y) < .000001d,
                    "Default-style aliases keep note hit areas on the same painted line.");
            }
        }
    }

    public static void DocxDefaultContextualStyleRetainsAmbiguousAndRichFallbacks()
    {
        foreach (string variant in new[] { "multiple", "malformed", "nondefault", "rich" })
        {
            DocxDocument source = ReadDefaultContextualStyleFixture("footnote", "Normal", true, true, true, variant);
            DocxBodyElement[] elements = source.BodyElements.Select(e => e is DocxParagraphElement p
                ? new DocxParagraphElement(p.Paragraph with { StyleId = p.Paragraph.StyleId ?? "PublicFallback" }) : e).ToArray();
            PdfLinkAnnotation[] expected = RenderCustomNoteLinks(source with { BodyElements = elements });
            PdfLinkAnnotation[] actual = RenderCustomNoteLinks(source);
            TestAssert.Equal(expected.Length, actual.Length);
            for (int i = 0; i < actual.Length; i++)
            {
                TestAssert.Equal(expected[i].Height, actual[i].Height);
                TestAssert.Equal(expected[i].Y, actual[i].Y);
            }
        }
    }

    private static DocxDocument ReadDefaultContextualStyleFixture(string kind, string defaultId,
        bool explicitPrevious, bool previousFlag, bool currentFlag, string variant = "ordinary") =>
        ReadSectionNumberingFixture(kind, (xml, _, parts) =>
        {
            XDocument styles = XDocument.Parse(parts["word/styles.xml"]);
            XElement? normal = styles.Root!.Elements(NoteWord + "style").SingleOrDefault(s => (string?)s.Attribute(NoteWord + "styleId") == "Normal");
            if (normal is null)
            {
                normal = new XElement(NoteWord + "style", new XAttribute(NoteWord + "type", "paragraph"),
                    new XAttribute(NoteWord + "styleId", "Normal"));
                styles.Root.Add(normal);
            }
            foreach (XElement style in styles.Root.Elements(NoteWord + "style")) style.SetAttributeValue(NoteWord + "default", "0");
            XElement declared = new(normal); declared.SetAttributeValue(NoteWord + "styleId", defaultId);
            declared.SetAttributeValue(NoteWord + "default", variant == "off" ? "off" : "true");
            if (defaultId == "Normal") normal.ReplaceWith(declared); else styles.Root.Add(declared);
            if (variant == "undeclared") declared.Attribute(NoteWord + "default")!.Remove();
            if (variant is "multiple" or "nondefault")
            {
                XElement other = new(declared); other.SetAttributeValue(NoteWord + "styleId", "PublicOther");
                other.SetAttributeValue(NoteWord + "default", variant == "multiple" ? "1" : "0"); styles.Root.Add(other);
            }
            XElement body = xml.Root!.Element(NoteWord + "body")!;
            XElement current = body.Elements(NoteWord + "p").First();
            XElement props = current.Element(NoteWord + "pPr")!;
            props.Element(NoteWord + "spacing")?.Remove();
            props.Add(new XElement(NoteWord + "spacing", new XAttribute(NoteWord + "before", "240"), new XAttribute(NoteWord + "after", "480")));
            props.Element(NoteWord + "contextualSpacing")?.Remove();
            if (currentFlag) props.Add(new XElement(NoteWord + "contextualSpacing"));
            current.Descendants(NoteWord + "t").Last().Value += " and following text " + string.Concat(Enumerable.Repeat("after note ", 35));
            XElement previousProps = new(NoteWord + "pPr", new XElement(NoteWord + "spacing", new XAttribute(NoteWord + "before", "0"), new XAttribute(NoteWord + "after", "120")));
            if (previousFlag) previousProps.Add(new XElement(NoteWord + "contextualSpacing"));
            XElement previous = new(NoteWord + "p", previousProps, new XElement(NoteWord + "r", new XElement(NoteWord + "t", "Public preceding paragraph.")));
            XElement explicitProps = explicitPrevious ? previousProps : props;
            explicitProps.AddFirst(new XElement(NoteWord + "pStyle", new XAttribute(NoteWord + "val", variant == "nondefault" ? "PublicOther" : defaultId)));
            if (variant == "allomitted") explicitProps.Element(NoteWord + "pStyle")!.Remove();
            if (variant == "malformed") props.AddFirst(new XElement(NoteWord + "pStyle"));
            if (variant == "rich") previous.Element(NoteWord + "r")!.Add(new XElement(NoteWord + "br"));
            current.AddBeforeSelf(previous);
            if (variant == "no-paragraph-styles")
                styles.Root.Elements(NoteWord + "style").Where(s => (string?)s.Attribute(NoteWord + "type") == "paragraph").Remove();
            parts["word/styles.xml"] = styles.ToString();
        });

    public static void DocxCurrentContextualMultilineNoteBeforeSpacingOwnsSectionFirstGaps()
    {
        CheckCurrentContextualMultilineNoteGaps("none");
    }

    public static void DocxCurrentContextualMultilineNoteBeforeSpacingOwnsDifferentStyleGaps()
    {
        CheckCurrentContextualMultilineNoteGaps("different");
    }

    private static void CheckCurrentContextualMultilineNoteGaps(string predecessor)
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (bool custom in new[] { false, true })
        foreach (bool exact in new[] { false, true })
        foreach (OoxPdfDocxMarkupGeometryMode geometry in new[]
            { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
        {
            double baseline = RenderCustomNoteLinks(ReadCurrentContextualMultilineNoteFixture(kind, custom, 0, true, exact, predecessor), geometry)[0].Height;
            foreach (int before in new[] { 6, 12, 24 })
            {
                DocxDocument document = ReadCurrentContextualMultilineNoteFixture(kind, custom, before, true, exact, predecessor);
                PdfLinkAnnotation[] links = RenderCustomNoteLinks(document, geometry);
                TestAssert.Equal(4, links.Length);
                double owned = predecessor == "none" ? before : Math.Max(0, before - 6);
                TestAssert.True(Math.Abs(links[0].Height - baseline - owned) < .000001d,
                    "Current contextual spacing retains positive first-line before-gaps at section starts or after different styles.");
                DocxParagraph paragraph = document.Paragraphs[predecessor == "none" ? 0 : 1];
                TestAssert.Equal((double)before, paragraph.SpacingBeforePoints);
                TestAssert.Equal(24d, paragraph.SpacingAfterPoints);
                TestAssert.True(paragraph.Spacing.ContextualSpacing == true, "The authored contextual setting remains active.");
            }
        }
    }

    public static void DocxCurrentContextualMultilineNoteBeforeSpacingRetainsSuppressedAndLaterSlots()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (bool custom in new[] { false, true })
        foreach (bool exact in new[] { false, true })
        foreach (string predecessor in new[] { "none", "same", "different" })
        {
            foreach (bool first in new[] { false, true })
            {
                if (first && predecessor != "same") continue;
                double baseline = RenderCustomNoteLinks(ReadCurrentContextualMultilineNoteFixture(kind, custom, 0, first, exact, predecessor))[0].Height;
                foreach (int before in new[] { 6, 12, 24 })
                    TestAssert.True(Math.Abs(RenderCustomNoteLinks(ReadCurrentContextualMultilineNoteFixture(kind, custom, before, first, exact, predecessor))[0].Height - baseline) < .000001d,
                        "Matching styles suppress ownership; later-line marks retain their line-only slot.");
            }
        }
    }

    public static void DocxCurrentContextualMultilineNoteBeforeSpacingRetainsComplexFallbacks()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (int variant in new[] { 0, 1, 2, 3 })
        {
            Action<XDocument> guard = xml =>
            {
                XElement paragraph = xml.Root!.Element(NoteWord + "body")!.Elements(NoteWord + "p").First();
                if (variant == 0) paragraph.Element(NoteWord + "pPr")!.Element(NoteWord + "spacing")!.SetAttributeValue(NoteWord + "beforeAutospacing", "1");
                if (variant == 1) paragraph.ElementsAfterSelf(NoteWord + "p").First().Elements(NoteWord + "r").Last().Add(new XElement(NoteWord + "br"));
                if (variant == 2) paragraph.AddBeforeSelf(new XElement(NoteWord + "p", new XElement(NoteWord + "pPr",
                    new XElement(NoteWord + "contextualSpacing"), new XElement(NoteWord + "spacing", new XAttribute(NoteWord + "after", "120")))));
                if (variant == 3)
                    foreach (XElement section in xml.Descendants(NoteWord + "sectPr"))
                    { section.Element(NoteWord + "cols")?.Remove(); section.Add(new XElement(NoteWord + "cols", new XAttribute(NoteWord + "num", "2"))); }
            };
            double baseline = RenderCustomNoteLinks(ReadCurrentContextualMultilineNoteFixture(kind, false, 0, true, false, "none", guard))[0].Height;
            double spaced = RenderCustomNoteLinks(ReadCurrentContextualMultilineNoteFixture(kind, false, 12, true, false, "none", guard))[0].Height;
            TestAssert.True(Math.Abs(spaced - baseline) < .000001d, "Complex current-contextual gaps retain their line-slot fallback.");
        }
    }

    private static DocxDocument ReadCurrentContextualMultilineNoteFixture(string kind, bool custom, int before,
        bool first, bool exact, string predecessor, Action<XDocument>? modify = null)
    {
        DocxDocument document = ReadMultilineNoteSpacingFixture(kind, custom, before, 24, first, xml =>
        {
            XElement paragraph = xml.Root!.Element(NoteWord + "body")!.Elements(NoteWord + "p").First();
            XElement props = paragraph.Element(NoteWord + "pPr")!;
            props.Add(new XElement(NoteWord + "contextualSpacing"));
            if (exact)
            {
                XElement spacing = props.Element(NoteWord + "spacing")!;
                spacing.SetAttributeValue(NoteWord + "line", "480"); spacing.SetAttributeValue(NoteWord + "lineRule", "exact");
            }
            if (predecessor != "none") paragraph.AddBeforeSelf(new XElement(NoteWord + "p", new XElement(NoteWord + "pPr",
                new XElement(NoteWord + "spacing", new XAttribute(NoteWord + "after", "120"))),
                new XElement(NoteWord + "r", new XElement(NoteWord + "t", "Public preceding paragraph."))));
            modify?.Invoke(xml);
        });
        if (predecessor != "different") return document;
        DocxParagraph previous = document.Paragraphs[0];
        DocxParagraph different = previous with { StyleId = "PublicPrevious" };
        return document with { BodyElements = document.BodyElements.Select(e => e is DocxParagraphElement p && ReferenceEquals(p.Paragraph, previous)
            ? (DocxBodyElement)new DocxParagraphElement(different) : e).ToArray() };
    }

    public static void DocxContextualNoteContributionLinksFollowSuppressedGaps()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (bool custom in new[] { false, true })
        foreach (OoxPdfDocxMarkupGeometryMode geometry in new[]
            { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
        {
            double lineHeight = RenderCustomNoteLinks(ReadNoteHitAreaFixture(kind, custom, 0), geometry)[0].Height;
            foreach (bool previousFlag in new[] { false, true })
            foreach (bool currentFlag in new[] { false, true })
            {
                DocxDocument document = ReadNoteHitAreaFixture(kind, custom, 24, xml =>
                {
                    XElement paragraph = xml.Root!.Element(NoteWord + "body")!.Elements(NoteWord + "p").First();
                    XElement props = paragraph.Element(NoteWord + "pPr")!;
                    props.Element(NoteWord + "spacing")!.SetAttributeValue(NoteWord + "before", "240");
                    props.Add(new XElement(NoteWord + "contextualSpacing", new XAttribute(NoteWord + "val", currentFlag ? "1" : "0")));
                    paragraph.AddBeforeSelf(new XElement(NoteWord + "p", new XElement(NoteWord + "pPr",
                        new XElement(NoteWord + "spacing", new XAttribute(NoteWord + "before", "0"), new XAttribute(NoteWord + "after", "120"),
                            new XAttribute(NoteWord + "line", "240"), new XAttribute(NoteWord + "lineRule", "auto")),
                        new XElement(NoteWord + "contextualSpacing", new XAttribute(NoteWord + "val", previousFlag ? "1" : "0"))),
                        new XElement(NoteWord + "r", new XElement(NoteWord + "t", "Public preceding paragraph."))));
                });
                PdfLinkAnnotation[] links = RenderCustomNoteLinks(document, geometry);
                TestAssert.Equal(4, links.Length);
                TestAssert.True(Math.Abs(links[0].Height - lineHeight - (currentFlag ? 0d : 30d)) < .000001d,
                    "A contextual mark omits its suppressed before/after contributions; predecessor suppression preserves the unconsumed before-gap.");
                TestAssert.Equal(12d, document.Paragraphs[1].SpacingBeforePoints);
                TestAssert.Equal(24d, document.Paragraphs[1].SpacingAfterPoints);
                TestAssert.Equal(6d, document.Paragraphs[0].SpacingAfterPoints);
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
                TestAssert.Equal(2, links.Length);
                TestAssert.Equal(2, links.Select(l => l.Destination).Distinct().Count());
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
            .RenderBlankPages(document, null, CancellationToken.None).SelectMany(p => p.Annotations).ToArray();

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

    public static void DocxSectionNoteNumbersDoNotCountCustomMarks()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        {
            DocxDocument document = ReadSectionNumberingFixture(kind, (xml, _, parts) =>
            {
                XElement reference = xml.Descendants(NoteWord + kind + "Reference").First();
                reference.SetAttributeValue(NoteWord + "customMarkFollows", "1");
                reference.AddAfterSelf(new XElement(NoteWord + "t", "*"));
                XDocument notes = XDocument.Parse(parts["word/" + kind + "s.xml"]);
                XElement marker = notes.Root!.Elements(NoteWord + kind).Single(n => (string?)n.Attribute(NoteWord + "id") == "37")
                    .Descendants(NoteWord + kind + "Ref").Single();
                marker.Name = NoteWord + "t";
                marker.Value = "*";
                parts["word/" + kind + "s.xml"] = notes.ToString();
            });
            DocxInlineReference[] references = document.Paragraphs.SelectMany(p => p.InlineReferences).ToArray();
            TestAssert.True(references[0].DisplayText is null, "A custom mark keeps its authored text.");
            TestAssert.Equal("iv", references[1].DisplayText!);
            TestAssert.Equal("09", references[2].DisplayText!);
            TestAssert.Equal("10", references[3].DisplayText!);
            DocxRelatedStory custom = document.RelatedStories.Single(s => s.Kind.ToString().Equals(kind, StringComparison.OrdinalIgnoreCase) && s.Id == "37");
            TestAssert.Equal("*", custom.Paragraphs[0].Runs[0].Text);
        }
    }

    public static void DocxSectionNoteNumberLinksKeepRepeatedLabelsDistinct()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (OoxPdfDocxMarkupGeometryMode geometry in new[]
            { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
        {
            DocxDocument document = ReadSectionNumberingFixture(kind, (xml, _, _) =>
                SetNoteSection(xml.Descendants(NoteWord + "sectPr").Last(), kind, "lowerRoman", "4", "eachSect"));
            CheckNoteLabels(document, ["iv", "v", "iv", "v"]);
            var renderer = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, geometry);
            PdfPage[] pages = renderer.RenderBlankPages(document, null, CancellationToken.None).ToArray();
            PdfLinkAnnotation[] links = pages.SelectMany(p => p.Annotations).ToArray();
            TestAssert.Equal(2, pages.Length);
            TestAssert.Equal(4, links.Length);
            TestAssert.Equal(4, links.Select(l => l.Destination).Distinct().Count());
            TestAssert.True(links.All(l => l.IsDestination && l.Width > 0 && l.Height > 0), "Every note retains a measurable link.");
            for (int i = 0; i < links.Length; i++)
                TestAssert.Equal(kind == "endnote" ? 1 : i / 2, links[i].Destination!.Value.PageIndex);
        }
    }

    private static readonly XNamespace NoteWord = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static DocxDocument ReadSectionNumberingFixture(string kind,
        Action<XDocument, XDocument, Dictionary<string, string>>? modify = null)
    {
        string seed = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "note-sections-" + kind + ".docx"));
        var parts = new Dictionary<string, string>();
        using (ZipArchive archive = ZipFile.OpenRead(seed))
        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            using var reader = new StreamReader(entry.Open());
            parts[entry.FullName] = reader.ReadToEnd();
        }
        XDocument document = XDocument.Parse(parts["word/document.xml"]);
        XDocument settings = XDocument.Parse(parts["word/settings.xml"]);
        modify?.Invoke(document, settings, parts);
        parts["word/document.xml"] = document.ToString();
        parts["word/settings.xml"] = settings.ToString();
        string input = TestFixtures.WriteTempPackage(".docx", parts);
        try { return DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup); }
        finally { File.Delete(input); }
    }

    private static void SetNoteSection(XElement section, string kind, string? format, string? start, string? restart)
    {
        section.Element(NoteWord + kind + "Pr")?.Remove();
        var properties = new XElement(NoteWord + kind + "Pr");
        if (format is not null) properties.Add(new XElement(NoteWord + "numFmt", new XAttribute(NoteWord + "val", format)));
        if (start is not null) properties.Add(new XElement(NoteWord + "numStart", new XAttribute(NoteWord + "val", start)));
        if (restart is not null) properties.Add(new XElement(NoteWord + "numRestart", new XAttribute(NoteWord + "val", restart)));
        section.AddFirst(properties);
    }

    private static void AddThirdNoteSection(XDocument xml, Dictionary<string, string> parts, string kind,
        string format, string start, string restart)
    {
        XElement body = xml.Root!.Element(NoteWord + "body")!;
        XElement final = body.Element(NoteWord + "sectPr")!;
        final.Remove();
        XElement[] tail = body.Elements(NoteWord + "p").TakeLast(2).Select(p => new XElement(p)).ToArray();
        body.Elements(NoteWord + "p").Last().Element(NoteWord + "pPr")!.Add(final);
        XDocument notes = XDocument.Parse(parts["word/" + kind + "s.xml"]);
        XElement template = notes.Root!.Elements(NoteWord + kind).First(n => n.Attribute(NoteWord + "type") is null);
        for (int i = 0; i < tail.Length; i++)
        {
            string id = i == 0 ? "101" : "19";
            tail[i].Descendants(NoteWord + kind + "Reference").Single().SetAttributeValue(NoteWord + "id", id);
            body.Add(tail[i]);
            XElement note = new(template);
            note.SetAttributeValue(NoteWord + "id", id);
            notes.Root.Add(note);
        }
        XElement section = new(final);
        SetNoteSection(section, kind, format, start, restart);
        body.Add(section);
        parts["word/" + kind + "s.xml"] = notes.ToString();
    }

    public static void DocxSingleSectionNoteNumbersUseExplicitSectionProperties()
    {
        foreach ((string fixture, string[] expected) in new[]
        {
            ("note-number-section-only-105", new[] { "105", "106" }),
            ("note-section-footnote-roman-4", new[] { "iv", "v" }),
            ("note-section-endnote-start-5", new[] { "v" }),
            ("note-number-both-105", new[] { "105", "106" })
        })
        {
            DocxDocument document = ReadNoteNavigationFixture(fixture);
            DocxInlineReference[] references = document.Paragraphs.SelectMany(p => p.InlineReferences).ToArray();
            TestAssert.Equal(expected.Length, references.Length);
            for (int i = 0; i < expected.Length; i++)
            {
                TestAssert.Equal(expected[i], references[i].DisplayText!);
                DocxRelatedStory story = document.RelatedStories.Single(s => s.Kind == references[i].Kind && s.Id == references[i].Id);
                TestAssert.Equal(expected[i], story.Paragraphs[0].Runs[0].Text);
            }
        }
    }

    public static void DocxSingleSectionNoteNumbersKeepUnsupportedFallbacks()
    {
        DocxDocument document = ReadNoteNavigationFixture("note-nav-numbered-ids");
        TestAssert.Equal(105, document.Settings.FootnoteReferenceSettings.NumberStart!.Value);
        TestAssert.Equal("1", document.Paragraphs[0].InlineReferences[0].DisplayText!);
        // Parsing settings remains independent of applying a section's overrides.
        DocxDocument section = ReadNoteNavigationFixture("note-number-section-only-105");
        TestAssert.True(section.Settings.FootnoteReferenceSettings.NumberStart is null,
            "Section numbering must not rewrite the authored document settings snapshot.");
        XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        DocxDocument malformed = ReadModifiedNoteNumberingFixture("note-number-section-only-105", xml =>
            xml.Descendants(word + "numStart").Single().SetAttributeValue(word + "val", "0"));
        TestAssert.Equal("1", malformed.Paragraphs[0].InlineReferences[0].DisplayText!);
        DocxDocument multiple = ReadModifiedNoteNumberingFixture("note-number-section-only-105", xml =>
            xml.Descendants(word + "body").Single().AddFirst(new XElement(word + "p", new XElement(word + "pPr",
                new XElement(xml.Descendants(word + "sectPr").Single())))));
        TestAssert.Equal("105", multiple.Paragraphs.SelectMany(p => p.InlineReferences).First().DisplayText!);
        DocxDocument restarting = ReadModifiedNoteNumberingFixture("note-number-section-only-105", xml =>
            xml.Descendants(word + "footnotePr").Single().Add(new XElement(word + "numRestart",
                new XAttribute(word + "val", "eachPage"))));
        TestAssert.Equal("1", restarting.Paragraphs[0].InlineReferences[0].DisplayText!);
    }

    public static void DocxSingleSectionNoteNumbersUseSectionDefaultsInsteadOfDocumentSettings()
    {
        foreach ((string kind, string documentFormat, string? sectionFormat, int? sectionStart, bool properties, string[] expected) in new[]
        {
            ("footnote", "decimal", (string?)null, (int?)null, true, new[] { "1", "2" }),
            ("footnote", "decimal", (string?)null, (int?)null, false, new[] { "1", "2" }),
            ("footnote", "upperRoman", (string?)null, (int?)null, true, new[] { "1", "2" }),
            ("footnote", "upperRoman", (string?)null, (int?)4, true, new[] { "4", "5" }),
            ("footnote", "upperRoman", (string?)"decimalZero", (int?)null, true, new[] { "01", "02" }),
            ("footnote", "upperRoman", (string?)"lowerLetter", (int?)4, true, new[] { "d", "e" }),
            ("endnote", "decimal", (string?)null, (int?)null, true, new[] { "i", "ii" }),
            ("endnote", "decimal", (string?)null, (int?)null, false, new[] { "i", "ii" }),
            ("endnote", "decimal", (string?)null, (int?)4, true, new[] { "iv", "v" }),
            ("endnote", "upperRoman", (string?)"decimalZero", (int?)null, true, new[] { "01", "02" }),
            ("endnote", "upperRoman", (string?)"lowerLetter", (int?)4, true, new[] { "d", "e" })
        })
        {
            DocxDocument document = ReadNotePrecedenceFixture(kind, documentFormat, sectionFormat, sectionStart, properties);
            DocxNoteReferenceSettings authored = kind == "endnote"
                ? document.Settings.EndnoteReferenceSettings : document.Settings.FootnoteReferenceSettings;
            TestAssert.Equal(105, authored.NumberStart!.Value);
            TestAssert.Equal(documentFormat, authored.NumberFormatValue!);
            CheckNoteLabels(document, expected);
        }
    }

    public static void DocxNoteDecimalZeroLabelsCrossTenWithoutExtraPadding()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
            CheckNoteLabels(ReadNotePrecedenceFixture(kind, "decimal", "decimalZero", 9, true), ["09", "10"]);
    }

    public static void DocxSingleSectionNoteNumbersPreserveInvalidAndRestartFallbacks()
    {
        XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        CheckNoteLabels(ReadNotePrecedenceFixture("footnote", "decimal", null, 0, true), ["105", "106"]);
        CheckNoteLabels(ReadNotePrecedenceFixture("footnote", "decimal", "cardinalText", null, true), ["105", "106"]);
        // An unsupported format retains the prior partial override of a valid start.
        CheckNoteLabels(ReadNotePrecedenceFixture("footnote", "decimal", "cardinalText", 4, true), ["4", "5"]);
        CheckNoteLabels(ReadNotePrecedenceFixture("footnote", "decimal", "upperRoman", 4, true, (xml, _) =>
            xml.Descendants(word + "footnotePr").Single().Add(new XElement(word + "numRestart",
                new XAttribute(word + "val", "eachPage")))), ["105", "106"]);
        CheckNoteLabels(ReadNotePrecedenceFixture("footnote", "decimal", "decimalZero", null, true, (_, settings) =>
            settings.Descendants(word + "footnotePr").Single().Add(new XElement(word + "numRestart",
                new XAttribute(word + "val", "eachPage")))), ["105", "106"]);
        CheckNoteLabels(ReadNotePrecedenceFixture("footnote", "decimal", null, null, true, (xml, _) =>
            xml.Descendants(word + "body").Single().AddFirst(new XElement(word + "p", new XElement(word + "pPr",
                new XElement(xml.Descendants(word + "sectPr").Single()))))), ["1", "2"]);
        CheckNoteLabels(ReadNotePrecedenceFixture("footnote", "decimal", null, null, true, (xml, _) =>
            xml.Descendants(word + "sectPr").Single().Remove()), ["105", "106"]);
    }

    private static void CheckNoteLabels(DocxDocument document, string[] expected)
    {
        DocxInlineReference[] references = DocxBlockTraversal.EnumerateBodyParagraphs(document.BodyElements).SelectMany(p => p.InlineReferences).ToArray();
        TestAssert.Equal(expected.Length, references.Length);
        for (int i = 0; i < expected.Length; i++)
        {
            TestAssert.Equal(expected[i], references[i].DisplayText!);
            DocxRelatedStory story = document.RelatedStories.Single(s => s.Kind == references[i].Kind && s.Id == references[i].Id);
            TestAssert.Equal(expected[i], story.Paragraphs[0].Runs[0].Text);
        }
    }

    private static DocxDocument ReadNotePrecedenceFixture(string kind, string documentFormat, string? sectionFormat,
        int? sectionStart, bool sectionProperties, Action<XDocument, XDocument>? modify = null)
    {
        string seed = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "note-nav-numbered-ids.docx"));
        var parts = new Dictionary<string, string>();
        using (ZipArchive archive = ZipFile.OpenRead(seed))
        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            using var reader = new StreamReader(entry.Open());
            parts[entry.FullName] = reader.ReadToEnd();
        }
        XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XDocument document = XDocument.Parse(parts["word/document.xml"]);
        XElement section = document.Descendants(word + "sectPr").Single();
        section.Elements().Where(e => e.Name == word + "footnotePr" || e.Name == word + "endnotePr").Remove();
        if (kind == "endnote")
        {
            foreach (XElement reference in document.Descendants(word + "footnoteReference")) reference.Name = word + "endnoteReference";
            foreach (XElement style in document.Descendants(word + "rStyle"))
                if ((string?)style.Attribute(word + "val") == "FootnoteReference") style.SetAttributeValue(word + "val", "EndnoteReference");
            parts["word/endnotes.xml"] = parts["word/footnotes.xml"].Replace("footnote", "endnote", StringComparison.Ordinal)
                .Replace("Footnote", "Endnote", StringComparison.Ordinal);
            parts.Remove("word/footnotes.xml");
            foreach (string part in new[] { "[Content_Types].xml", "word/_rels/document.xml.rels" })
                parts[part] = parts[part].Replace("footnotes", "endnotes", StringComparison.Ordinal);
        }
        if (sectionProperties)
        {
            var properties = new XElement(word + kind + "Pr",
                new XElement(word + "pos", new XAttribute(word + "val", kind == "footnote" ? "pageBottom" : "docEnd")));
            if (sectionFormat is not null) properties.Add(new XElement(word + "numFmt", new XAttribute(word + "val", sectionFormat)));
            if (sectionStart is not null) properties.Add(new XElement(word + "numStart", new XAttribute(word + "val", sectionStart.Value)));
            section.Add(properties);
        }
        var settings = new XDocument(new XElement(word + "settings", new XElement(word + kind + "Pr",
            new XElement(word + "numFmt", new XAttribute(word + "val", documentFormat)),
            new XElement(word + "numStart", new XAttribute(word + "val", 105)))));
        modify?.Invoke(document, settings);
        parts["word/document.xml"] = document.ToString();
        parts["word/settings.xml"] = settings.ToString();
        string input = TestFixtures.WriteTempPackage(".docx", parts);
        try { return DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup); }
        finally { File.Delete(input); }
    }

    public static void DocxDocumentEndnotesFollowTrailingBodySpacingAtNominalScale()
    {
        var renderer = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup,
            OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        double Top(string fixture) => renderer.InspectLayout(ReadNoteNavigationFixture(fixture)).Pages[0]
            .PlacedRelatedStories.Single(s => s.Kind == "Endnote" && s.Id == "1").TopY;
        double zero = Top("note-nav-endnote-body");
        TestAssert.True(Math.Abs(zero - Top("note-body-after-8") - 8d) < 0.001d,
            "Document endnotes must follow eight points of trailing body spacing.");
        TestAssert.True(Math.Abs(zero - Top("note-body-after-24") - 24d) < 0.001d,
            "Document endnotes must follow twenty-four points of trailing body spacing.");
        TestAssert.True(Math.Abs(zero - Top("note-body-default-after-8") - 8d) < 0.001d,
            "Inherited body spacing participates in document-endnote flow.");
        foreach (string fixture in new[] { "note-separator-after-8", "note-separator-after-24", "note-separator-before-24", "note-separator-default-after-8" })
        {
            TestAssert.True(Math.Abs(zero - Top(fixture)) < 0.001d,
                "The separator's own paragraph spacing must not replace body flow spacing.");
        }
    }

    private static DocxDocument ReadModifiedNoteNumberingFixture(string fixture, Action<XDocument> modify)
    {
        string seed = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", fixture + ".docx"));
        var parts = new Dictionary<string, string>();
        using (ZipArchive archive = ZipFile.OpenRead(seed))
        {
            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                using var reader = new StreamReader(entry.Open());
                parts[entry.FullName] = reader.ReadToEnd();
            }
        }
        XDocument xml = XDocument.Parse(parts["word/document.xml"]);
        modify(xml);
        parts["word/document.xml"] = xml.ToString();
        string input = TestFixtures.WriteTempPackage(".docx", parts);
        try { return DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup); }
        finally { File.Delete(input); }
    }

    public static void DocxReaderRetainsAutomaticNoteStoryMarks()
    {
        foreach (string fixture in NoteNavigationFixtures)
        {
            DocxDocument document = ReadNoteNavigationFixture(fixture);
            foreach (DocxParagraph paragraph in DocxBlockTraversal.EnumerateBodyParagraphs(document.BodyElements))
            foreach (DocxInlineReference reference in paragraph.InlineReferences)
            {
                if (reference.Kind is not (DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote)) continue;
                DocxRelatedStory story = document.RelatedStories.Single(story => story.Kind == reference.Kind && story.Id == reference.Id);
                TestAssert.Equal(reference.DisplayText!, story.Paragraphs[0].Runs[0].Text);
                TestAssert.Equal(DocxRunVerticalAlignment.Superscript, story.Paragraphs[0].Runs[0].VerticalAlignment);
            }
        }
    }

    public static void DocxNoteReferenceAnnotationsNavigateAcrossPages()
    {
        foreach (string fixture in NoteNavigationFixtures)
        foreach (OoxPdfDocxMarkupGeometryMode geometry in new[]
            { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
        {
            DocxDocument document = ReadNoteNavigationFixture(fixture);
            var renderer = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, geometry);
            PdfPage[] pages = renderer.RenderBlankPages(document, null, CancellationToken.None).ToArray();
            PdfLinkAnnotation[] links = pages.SelectMany(page => page.Annotations).ToArray();
            int expected = fixture is "note-nav-two-footnotes" or "note-nav-mixed-notes" or "note-nav-numbered-ids" ? 2 : 1;
            TestAssert.Equal(expected, links.Length);
            TestAssert.True(links.All(link => link.IsDestination && link.Width > 0d && link.Height > 0d),
                "Every visible note marker must have a measurable internal link.");
            TestAssert.Equal(expected, links.Select(link => link.Destination).Distinct().Count());
            foreach (PdfLinkAnnotation link in links)
            {
                PdfLinkDestination target = link.Destination!.Value;
                TestAssert.True(target.PageIndex >= 0 && target.PageIndex < pages.Length &&
                    target.Top > 0d && target.Top <= pages[target.PageIndex].Height,
                    "Note destinations must resolve to a rendered page location.");
            }
            if (fixture == "note-nav-endnote-page-two")
            {
                TestAssert.Equal(2, pages.Length);
                TestAssert.Equal(1, links[0].Destination!.Value.PageIndex);
            }
        }
    }

    public static void DocxNoteReferenceAnnotationsRequireVisibleSourceAndTarget()
    {
        DocxDocument source = ReadNoteNavigationFixture("note-nav-footnote-body");
        var renderer = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup,
            OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        DocxDocument orphan = source with { RelatedStories = [] };
        TestAssert.Equal(0, renderer.RenderBlankPages(orphan, null, CancellationToken.None).Sum(page => page.Annotations.Count));
        DocxParagraph paragraph = source.Paragraphs.Single();
        int markerRun = paragraph.InlineReferences.Single().SourceRunIndex;
        DocxDocument hidden = source with { BodyElements = [new DocxParagraphElement(paragraph with
            { Runs = paragraph.Runs.Where(run => run.SourceRunIndex != markerRun).ToArray() })] };
        TestAssert.Equal(0, renderer.RenderBlankPages(hidden, null, CancellationToken.None).Sum(page => page.Annotations.Count));
    }

    public static void DocxReaderNoteStoryMarksFollowFilteredBodyLabels()
    {
        string seed = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "note-nav-two-footnotes.docx"));
        var parts = new Dictionary<string, string>();
        using (ZipArchive archive = ZipFile.OpenRead(seed))
        {
            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                using var reader = new StreamReader(entry.Open());
                parts[entry.FullName] = reader.ReadToEnd();
            }
        }
        XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XDocument xml = XDocument.Parse(parts["word/document.xml"]);
        XElement marker = xml.Descendants(word + "footnoteReference").First().Parent!;
        marker.ReplaceWith(new XElement(word + "del", new XAttribute(word + "id", "9"),
            new XAttribute(word + "author", "Public reviewer"), new XElement(marker)));
        parts["word/document.xml"] = xml.ToString();
        string input = TestFixtures.WriteTempPackage(".docx", parts);
        try
        {
            foreach (OoxPdfDocxMarkupMode mode in new[]
                { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.Original, OoxPdfDocxMarkupMode.AllMarkup })
            {
                DocxDocument document = DocxTests.ReadDocx(input, mode);
                DocxInlineReference[] references = document.Paragraphs.SelectMany(paragraph => paragraph.InlineReferences).ToArray();
                TestAssert.Equal(mode == OoxPdfDocxMarkupMode.Final ? 1 : 2, references.Length);
                foreach (DocxInlineReference reference in references)
                {
                    DocxRelatedStory story = document.RelatedStories.Single(story => story.Kind == reference.Kind && story.Id == reference.Id);
                    TestAssert.Equal(reference.DisplayText!, story.Paragraphs[0].Runs[0].Text);
                }
            }
        }
        finally { File.Delete(input); }
    }

    public static void DocxFittingDocumentEndnotesStayAboveFootnoteArea()
    {
        DocxDocument document = ReadNoteNavigationFixture("note-nav-mixed-notes");
        var renderer = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup,
            OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        DocxLayoutSnapshot fitting = renderer.InspectLayout(document);
        TestAssert.Equal(1, fitting.Pages.Count);
        var footnotes = fitting.Pages[0].PlacedRelatedStories.Where(story => story.Kind == "Footnote").ToArray();
        var endnotes = fitting.Pages[0].PlacedRelatedStories.Where(story => story.Kind == "Endnote").ToArray();
        TestAssert.True(endnotes.Length > 0 && endnotes.Min(story => story.TopY - story.Height) >= footnotes.Max(story => story.TopY),
            "Fitting document endnotes belong below the body and above the complete footnote area.");
        DocxRelatedStory endnote = document.RelatedStories.Single(story => story.Kind == DocxRelatedStoryKind.Endnote && story.IsNormalStoryType);
        DocxParagraph paragraph = endnote.Paragraphs[0];
        DocxRelatedStory longEndnote = endnote with { BodyElements = [new DocxParagraphElement(paragraph with
            { Runs = [paragraph.Runs[0], paragraph.Runs[1] with { Text = string.Join(' ', Enumerable.Repeat("Long endnote content", 500)) }] })] };
        DocxDocument overflowing = document with { RelatedStories = document.RelatedStories.Select(story => ReferenceEquals(story, endnote) ? longEndnote : story).ToArray() };
        DocxLayoutSnapshot fallback = renderer.InspectLayout(overflowing);
        TestAssert.True(fallback.Pages.Count > 1 && !fallback.Pages[0].PlacedRelatedStories.Any(story => story.Kind == "Endnote"),
            "Endnotes that exceed the free gap retain continuation placement and cannot overlap footnotes.");
    }

    private static readonly string[] NoteNavigationFixtures =
    ["note-nav-footnote-body", "note-nav-footnote-table", "note-nav-endnote-body", "note-nav-endnote-page-two",
     "note-nav-two-footnotes", "note-nav-mixed-notes", "note-nav-numbered-ids", "note-nav-shared-run"];

    private static DocxDocument ReadNoteNavigationFixture(string fixture) => DocxTests.ReadDocx(
        Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", fixture + ".docx")),
        OoxPdfDocxMarkupMode.AllMarkup);

    public static void DocxWordCompatibleAllMarkupClampsCommentConnectorAnchorsAtPageEdge()
    {
        DocxTextRun run = new("Edge anchor", 10d, null, false, false, false, null, null)
        {
            SourceRunIndex = 0,
            SourceTextOffsetInRun = 0
        };
        DocxParagraph paragraph = new(
            [run],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1.2d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null)
        {
            InlineReferences =
            [
                new DocxInlineReference(DocxRelatedStoryKind.Comment, "1", null, SourceRunIndex: 1, RunChildIndex: 0, TextOffsetInRun: 0, DisplayText: null)
            ],
            CommentRanges =
            [
                new DocxCommentRange("1", 0, 0, 0, 0, 1, 0)
            ]
        };
        DocxRelatedStory commentStory = new(
            DocxRelatedStoryKind.Comment,
            "/word/comments.xml",
            "1",
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Public edge comment body", 10d, 12d))],
            [],
            [], null);
        // Near-edge clamp calibration: anchor = L - L * (1 - s) - 3.18 with s = 200 / 346.5,
        // so L = 4 lands at -0.87 and clamps to 0.5 (L = 10 no longer reaches the edge
        // under the lane-fit print scale).
        DocxDocument document = new(
            200d,
            240d,
            4d,
            120d,
            20d,
            20d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(paragraph)],
            [],
            [])
        {
            RelatedStories = [commentStory],
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };
        var renderer = new DocxRenderer(
            fontResolver: null,
            markupMode: OoxPdfDocxMarkupMode.AllMarkup,
            markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);

        DocxMarkupBalloonPlacementSnapshot placement = renderer.InspectMarkupBalloons(document)
            .Single(item => item.Kind == "Comment");

        TestAssert.True(
            placement.AnchorConnectorX >= 0.5d &&
            placement.AnchorConnectorX < 1d,
            string.Create(
                CultureInfo.InvariantCulture,
                $"Word-compatible all-markup should clamp near-edge comment connector anchors inside the page media box. AnchorX={placement.AnchorConnectorX}."));
        TestAssert.True(placement.AnchorConnectorClamped, "Near-edge connector placement snapshots should expose that the anchor was clamped.");
    }

    public static void DocxReaderSplitsNestedVisibleInlineContainerPageBreaks()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/break" TargetMode="External"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p>
                      <w:hyperlink r:id="rIdLink">
                        <w:ins>
                          <w:r><w:t>Before</w:t><w:br w:type="page"/><w:t>After</w:t></w:r>
                        </w:ins>
                      </w:hyperlink>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        TestAssert.Equal(3, document.BodyElements.Count);
        TestAssert.True(document.BodyElements[0] is DocxParagraphElement, "Text before the nested break should remain a paragraph fragment.");
        TestAssert.True(document.BodyElements[1] is DocxPageBreakElement, "Nested page breaks should become explicit body break elements.");
        TestAssert.True(document.BodyElements[2] is DocxParagraphElement, "Text after the nested break should remain a paragraph fragment.");
        DocxParagraph before = ((DocxParagraphElement)document.BodyElements[0]).Paragraph;
        DocxParagraph after = ((DocxParagraphElement)document.BodyElements[2]).Paragraph;
        TestAssert.Equal("Before", string.Concat(before.Runs.Select(run => run.Text)));
        TestAssert.Equal("After", string.Concat(after.Runs.Select(run => run.Text)));
        TestAssert.Equal(1, before.Hyperlinks.Count);
        TestAssert.Equal(1, after.Hyperlinks.Count);
    }

    public static void DocxReaderSplitsMoveToFinalViewPageBreaks()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:moveTo w:id="9" w:author="Author" w:date="2026-06-02T00:00:00Z">
                        <w:r><w:t>Before</w:t><w:br w:type="page"/><w:t>After</w:t></w:r>
                      </w:moveTo>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        TestAssert.Equal(3, document.BodyElements.Count);
        TestAssert.True(document.BodyElements[0] is DocxParagraphElement, "Moved-to text before the break should remain a paragraph fragment.");
        TestAssert.True(document.BodyElements[1] is DocxPageBreakElement, "Moved-to page breaks should become explicit body break elements.");
        TestAssert.True(document.BodyElements[2] is DocxParagraphElement, "Moved-to text after the break should remain a paragraph fragment.");
        TestAssert.Equal("Before", string.Concat(((DocxParagraphElement)document.BodyElements[0]).Paragraph.Runs.Select(run => run.Text)));
        TestAssert.Equal("After", string.Concat(((DocxParagraphElement)document.BodyElements[2]).Paragraph.Runs.Select(run => run.Text)));
    }

    public static void DocxLayoutReservesPageFootnoteStoryArea()
    {
        DocxParagraph anchor = DocxTests.CreateDocxLayoutParagraph("Anchor paragraph", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "7",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 6)
            ]
        };
        DocxParagraph filler = DocxTests.CreateDocxLayoutParagraph("Filler paragraph keeps body close to the bottom", 10d, 12d);
        DocxParagraph footnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Footnote body line one wraps line two", 10d, 12d);
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "7",
            [new DocxParagraphElement(footnoteParagraph)],
            [],
            [], null);
        var document = new DocxDocument(
            220d,
            112d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [
                new DocxParagraphElement(anchor),
                new DocxParagraphElement(filler),
                new DocxParagraphElement(filler),
                new DocxParagraphElement(filler),
                new DocxParagraphElement(filler)
            ],
            [],
            [])
        {
            RelatedStories = [footnoteStory]
        };

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None));
        DocxLayoutPageSnapshot footnotePage = snapshot.Pages.Single(page => page.PlacedFootnoteStoryCount == 1);
        double footnoteTop = footnotePage.PlacedRelatedItems.Max(item => item.Y + item.Height);
        double bodyBottom = footnotePage.Items.Min(item => item.Y);

        TestAssert.True(footnotePage.PlacedRelatedStoryTextLineCount >= 1, "Resolved footnote story text should be placed on the body-reference page.");
        TestAssert.True(bodyBottom >= footnoteTop, $"Body layout should reserve page space above placed footnotes; body bottom {bodyBottom} footnote top {footnoteTop}.");
    }

    public static void DocxFootnoteStoryLaysOutUnscaledDesignSpace()
    {
        // RV06 anchor probe (edge-footanchor-5, Word 16.0): footnote stories lay
        // out unscaled in design space while emission maps uniformly (WC first
        // baseline 236.62 = affine-mapped design). Pre-fix the word-compatible
        // story layout scaled metrics in place.
        DocxParagraph SizedLine(string text, double fontSize)
        {
            return DocxTests.CreateDocxLayoutParagraph(text, fontSize, 12d) with
            {
                LineSpacingPoints = null
            };
        }

        double[] Pitches(OoxPdfDocxMarkupGeometryMode mode, double scale)
        {
            DocxParagraph anchor = DocxTests.CreateDocxLayoutParagraph("Anchor", 10d, 12d) with
            {
                InlineReferences =
                [
                    new DocxInlineReference(
                        DocxRelatedStoryKind.Footnote,
                        "7",
                        CustomMarkFollowsValue: null,
                        DisplayText: "1",
                        SourceRunIndex: 0,
                        RunChildIndex: 1,
                        TextOffsetInRun: 6)
                ]
            };
            var footnoteStory = new DocxRelatedStory(
                DocxRelatedStoryKind.Footnote,
                "/word/footnotes.xml",
                "7",
                [
                    new DocxParagraphElement(SizedLine("Alpha", 12d)),
                    new DocxParagraphElement(SizedLine("Beta", 15d)),
                    new DocxParagraphElement(SizedLine("Gamma", 12d))
                ],
                [],
                [],
                null);
            var document = new DocxDocument(
                220d,
                200d,
                10d,
                10d,
                10d,
                10d,
                DocxPageSettings.Empty,
                [],
                [],
                [],
                [new DocxParagraphElement(anchor)],
                [],
                [])
            {
                RelatedStories = [footnoteStory]
            };
            DocxTextLineLayout[] lines = new DocxLayoutEngine(mode, scale)
                .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
                .Pages.SelectMany(page => page.PlacedRelatedStories).SelectMany(story => story.TextLines).ToArray();
            TestAssert.Equal(3, lines.Length);
            return [lines[0].BaselineY - lines[1].BaselineY, lines[1].BaselineY - lines[2].BaselineY];
        }

        // Footnote stories lay out unscaled in design space (emission maps
        // uniformly); the scaled layout keeps full unscaled transitions.
        double[] preserve = Pitches(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, 1d);
        double[] scaled = Pitches(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, 0.75d);
        TestAssert.True(
            Math.Abs(scaled[0] - preserve[0]) < 0.001d,
            $"Footnote story layout should stay design space. preserve={preserve[0]} scaled={scaled[0]}.");
        TestAssert.True(
            Math.Abs(scaled[1] - preserve[1]) < 0.001d,
            $"Footnote story layout should stay design space. preserve={preserve[1]} scaled={scaled[1]}.");
    }

    public static void DocxEndnoteStartTopRecoversDesignBodyBottom()
    {
        // RV06 anchor probe (edge-endanchor-5, Word 16.0): scaled body bottom
        // 689.77 with frame top 720, first inset 11.28 and print scale 0.75874
        // recovers the design body end 683.75; scale 1.0 and degenerate frames
        // keep legacy behavior.
        TestAssert.True(
            Math.Abs(DocxLayoutEngine.ResolveDesignBodyBottomForEndnoteStart(689.77d, 720d, 11.28d, 0.75874d) - 683.75d) < 0.02d,
            "Design body end should recover 683.75.");
        TestAssert.Equal(100d, DocxLayoutEngine.ResolveDesignBodyBottomForEndnoteStart(100d, 720d, 11.28d, 1d));
        TestAssert.Equal(700d, DocxLayoutEngine.ResolveDesignBodyBottomForEndnoteStart(700d, 720d, null, 0.75d));
    }

    public static void DocxLayoutStacksMultipleFootnotesOnOnePage()
    {
        DocxParagraph anchor = DocxTests.CreateDocxLayoutParagraph("Anchor paragraph with two footnote markers", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "31",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 0,
                    TextOffsetInRun: 8),
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "32",
                    CustomMarkFollowsValue: null,
                    DisplayText: "2",
                    SourceRunIndex: 0,
                    RunChildIndex: 0,
                    TextOffsetInRun: 22)
            ]
        };
        DocxParagraph firstFootnote = DocxTests.CreateDocxLayoutParagraph("First footnote body", 10d, 12d);
        DocxParagraph secondFootnote = DocxTests.CreateDocxLayoutParagraph("Second footnote body", 10d, 12d);
        var firstStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "31",
            [new DocxParagraphElement(firstFootnote)],
            [],
            [], null);
        var secondStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "32",
            [new DocxParagraphElement(secondFootnote)],
            [],
            [], null);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(anchor)], []) with
        {
            RelatedStories = [firstStory, secondStory]
        };

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None));
        DocxLayoutPageSnapshot footnotePage = snapshot.Pages.Single(page => page.PlacedFootnoteStoryCount == 2);
        DocxPlacedRelatedStoryLayoutSnapshot[] placedStories = footnotePage.PlacedRelatedStories.ToArray();
        DocxPlacedRelatedStoryLayoutSnapshot firstPlaced = placedStories.Single(story => story.Id == "31");
        DocxPlacedRelatedStoryLayoutSnapshot secondPlaced = placedStories.Single(story => story.Id == "32");

        TestAssert.True(firstPlaced.TopY > secondPlaced.TopY, "The first footnote body should be placed above the later footnote body instead of both stories sharing the page bottom.");
        TestAssert.True(secondPlaced.TopY <= firstPlaced.TopY - firstPlaced.Height + 0.001d, "Stacked footnote bodies should not overlap vertically.");
        TestAssert.True(firstPlaced.SeparatorY is not null && secondPlaced.SeparatorY is null, "Generic separator geometry should be emitted once above the stacked footnote group.");
    }

    public static void DocxRendererClipsOverlongPlacedFootnoteStoryToPageRegion()
    {
        DocxParagraph bodyParagraph = DocxTests.CreateDocxLayoutParagraph("body with long footnote", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "16",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        DocxParagraph footnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Footnote overflow body", 10d, 12d);
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "16",
            Enumerable.Range(0, 8).Select(_ => new DocxParagraphElement(footnoteParagraph)).Cast<DocxBodyElement>().ToArray(),
            [],
            [], null);
        var document = new DocxDocument(
            220d,
            82d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(bodyParagraph)],
            [bodyParagraph],
            [])
        {
            RelatedStories = [footnoteStory]
        };

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None));
        DocxPlacedRelatedStoryLayoutSnapshot placedStory = snapshot.Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .First(story => story.Kind == "Footnote" && story.Id == "16");
        // Overlong notes slice across pages now; the head slice keeps full-story height with a page-slice height.
        PdfPage placedPage = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .RenderBlankPages(document, null, CancellationToken.None)
            .First(page => page.Content.Contains(" re W n", StringComparison.Ordinal));

        TestAssert.True(placedStory.ContentHeight > placedStory.Height, "The overlong footnote should retain full-story height while exposing the clipped page slice height.");
        TestAssert.Equal(0d, placedStory.ContentTopOffset);
        TestAssert.Contains(" re W n", placedPage.Content);
    }

    public static void DocxLayoutPlacesFootnoteOnRenderedMarkerRunPage()
    {
        var leadingRun = new DocxTextRun(
            string.Concat(Enumerable.Repeat("alpha beta gamma delta ", 12)),
            10d,
            null,
            false,
            false,
            false,
            null,
            null);
        var markerRun = new DocxTextRun(
            "marker",
            10d,
            null,
            false,
            false,
            false,
            null,
            null);
        DocxParagraph anchor = new(
            [leadingRun, markerRun],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            12d,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null)
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "11",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 1,
                    RunChildIndex: 1,
                    TextOffsetInRun: 0)
            ]
        };
        DocxParagraph footnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Footnote body", 10d, 12d);
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "11",
            [new DocxParagraphElement(footnoteParagraph)],
            [],
            [], null);
        var document = new DocxDocument(
            220d,
            82d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(anchor)],
            [anchor],
            [])
        {
            RelatedStories = [footnoteStory]
        };

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout);
        int footnotePageIndex = Array.FindIndex(snapshot.Pages.ToArray(), page => page.PlacedFootnoteStoryCount == 1);

        TestAssert.True(layout.Pages.Count > 1, "The marker paragraph should span pages so note ownership is not reducible to source-block ownership.");
        TestAssert.True(layout.Pages[0].Items.OfType<DocxTextLineLayout>().Any(line => line.SourceBlockIndex == 0), "The first page should contain the source paragraph before the marker run is rendered.");
        TestAssert.Equal(0, snapshot.Pages[0].PlacedFootnoteStoryCount);
        TestAssert.True(footnotePageIndex > 0, "The footnote story should be placed on the page where the marker run is rendered, not the first page containing the source paragraph.");
        TestAssert.True(layout.Pages[footnotePageIndex].Items.OfType<DocxTextLineLayout>().Any(line => line.Segments.Any(segment => segment.SourceTextRunIndex == 1)), "The selected footnote page should contain the marker source run.");
    }

    public static void DocxLayoutPlacesFootnoteOnRenderedMarkerOffsetPage()
    {
        const int markerOffset = 180;
        string text = string.Concat(Enumerable.Repeat("alpha beta gamma delta ", 12));
        var run = new DocxTextRun(
            text,
            10d,
            null,
            false,
            false,
            false,
            null,
            null);
        DocxParagraph anchor = new(
            [run],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            12d,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null)
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "12",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: markerOffset)
            ]
        };
        DocxParagraph footnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Footnote body", 10d, 12d);
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "12",
            [new DocxParagraphElement(footnoteParagraph)],
            [],
            [], null);
        var document = new DocxDocument(
            220d,
            82d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(anchor)],
            [anchor],
            [])
        {
            RelatedStories = [footnoteStory]
        };

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout);
        int footnotePageIndex = Array.FindIndex(snapshot.Pages.ToArray(), page => page.PlacedFootnoteStoryCount == 1);

        TestAssert.True(layout.Pages.Count > 1, "The single marker run should span pages so note ownership requires source offsets.");
        TestAssert.Equal(0, snapshot.Pages[0].PlacedFootnoteStoryCount);
        TestAssert.True(footnotePageIndex > 0, "The footnote story should be placed on the page containing the marker offset, not the first page containing the source run.");
        TestAssert.True(
            layout.Pages[footnotePageIndex].Items.OfType<DocxTextLineLayout>().Any(line =>
                line.Segments.Any(segment =>
                    segment.SourceTextRunIndex == 0 &&
                    markerOffset >= segment.SourceTextOffsetInRun &&
                    markerOffset <= segment.SourceTextOffsetInRun + segment.Text.Length)),
            "The selected footnote page should contain the marker source offset inside the rendered run segment.");
    }

    public static void DocxLayoutPlacesFootnoteOnReaderProvenanceMarkerOffsetPage()
    {
        string leadingText = string.Concat(Enumerable.Repeat("alpha beta gamma delta ", 12));
        const string markerText = "1";
        int markerOffset = leadingText.Length;
        var leadingRun = new DocxTextRun(
            leadingText,
            10d,
            null,
            false,
            false,
            false,
            null,
            null)
        {
            SourceRunIndex = 0,
            SourceTextOffsetInRun = 0
        };
        var markerRun = new DocxTextRun(
            markerText,
            10d,
            null,
            false,
            false,
            false,
            null,
            null,
            0d,
            false,
            "superscript",
            false,
            null,
            false,
            null,
            null,
            null,
            null,
            null,
            false,
            null,
            false,
            null,
            null)
        {
            SourceRunIndex = 0,
            SourceTextOffsetInRun = markerOffset
        };
        var trailingRun = new DocxTextRun(
            " trailing text after marker",
            10d,
            null,
            false,
            false,
            false,
            null,
            null)
        {
            SourceRunIndex = 0,
            SourceTextOffsetInRun = markerOffset
        };
        DocxParagraph anchor = new(
            [leadingRun, markerRun, trailingRun],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            12d,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null)
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "13",
                    CustomMarkFollowsValue: null,
                    DisplayText: markerText,
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: markerOffset)
            ]
        };
        DocxParagraph footnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Footnote body", 10d, 12d);
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "13",
            [new DocxParagraphElement(footnoteParagraph)],
            [],
            [], null);
        var document = new DocxDocument(
            220d,
            82d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(anchor)],
            [anchor],
            [])
        {
            RelatedStories = [footnoteStory]
        };

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout);
        int footnotePageIndex = Array.FindIndex(snapshot.Pages.ToArray(), page => page.PlacedFootnoteStoryCount == 1);

        TestAssert.True(layout.Pages.Count > 1, "The normalized run sequence should span pages so source provenance is required.");
        TestAssert.Equal(0, snapshot.Pages[0].PlacedFootnoteStoryCount);
        TestAssert.True(footnotePageIndex > 0, "The footnote story should be placed on the page containing the normalized marker run, not the first page containing earlier text from the same source run.");
        TestAssert.True(
            layout.Pages[footnotePageIndex].Items.OfType<DocxTextLineLayout>().Any(line =>
                line.Segments.Any(segment =>
                    segment.Text == markerText &&
                    segment.SourceTextRunIndex == 0 &&
                    segment.SourceTextOffsetInRun == markerOffset)),
            "The selected footnote page should contain the normalized marker display run with original OOXML source provenance.");
    }

    public static void DocxLayoutPlacesSectEndEndnoteOnOwningSectionEndPage()
    {
        DocxParagraph firstSectionParagraph = DocxTests.CreateDocxLayoutParagraph("first section endnote marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Endnote,
                    "21",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 6)
            ]
        };
        DocxParagraph secondSectionParagraph = DocxTests.CreateDocxLayoutParagraph("second section body", 10d, 12d);
        DocxParagraph endnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Endnote body", 10d, 12d);
        var endnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Endnote,
            "/word/endnotes.xml",
            "21",
            [new DocxParagraphElement(endnoteParagraph)],
            [],
            [], null);
        DocxPageSettings firstSectionSettings = DocxPageSettings.Empty with
        {
            EndnoteReferenceSettings = DocxNoteReferenceSettings.Empty with { PositionValue = "sectEnd" }
        };
        var sectionBreak = new DocxSectionBreakElement(firstSectionSettings, DocxSectionBreakType.NextPage, null, null, null, []);
        var document = new DocxDocument(
            220d,
            112d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [
                new DocxParagraphElement(firstSectionParagraph),
                sectionBreak,
                new DocxParagraphElement(secondSectionParagraph)
            ],
            [firstSectionParagraph, secondSectionParagraph],
            [])
        {
            RelatedStories = [endnoteStory]
        };

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None));
        DocxLayoutPageSnapshot endnotePage = snapshot.Pages.Single(page => page.PlacedEndnoteStoryCount == 1);

        TestAssert.True(snapshot.Pages.Count >= 2, "The next-page section break should separate the two sections.");
        TestAssert.Equal("sectEnd", snapshot.Pages[0].SectionEndnotePositionValue ?? string.Empty);
        TestAssert.Equal(1, snapshot.Pages[0].PlacedEndnoteStoryCount);
        TestAssert.Equal(0, snapshot.Pages[^1].PlacedEndnoteStoryCount);
        TestAssert.True(endnotePage.PlacedRelatedStories.Any(story => story.Kind == "Endnote" && story.SourceBlockIndex == 0), "A sectEnd endnote should be owned by the source section end page, not rewritten as a document-end story.");
    }

    public static void DocxLayoutWrapsDocumentEndEndnoteUsingFinalSectionPageWidth()
    {
        DocxPageSettings wideFirstSection = new(
            "12240",
            "12000",
            null,
            "1440",
            "1440",
            "1440",
            "1440",
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null);
        DocxPageSettings narrowFinalSection = new(
            "4400",
            "6000",
            null,
            "200",
            "200",
            "200",
            "200",
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null);
        DocxParagraph anchor = DocxTests.CreateDocxLayoutParagraph("Narrow document-end endnote marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Endnote,
                    "25",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 7)
            ]
        };
        DocxParagraph endnoteParagraph = DocxTests.CreateDocxLayoutParagraph("alpha beta gamma delta epsilon zeta eta theta", 10d, 12d);
        var endnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Endnote,
            "/word/endnotes.xml",
            "25",
            [new DocxParagraphElement(endnoteParagraph)],
            [],
            [], null);
        var document = new DocxDocument(
            612d,
            792d,
            72d,
            72d,
            72d,
            72d,
            narrowFinalSection,
            [],
            [],
            [],
            [
                new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Wide first section", 10d, 12d)),
                new DocxSectionBreakElement(wideFirstSection, DocxSectionBreakType.NextPage, null, null, null, []),
                new DocxParagraphElement(anchor)
            ],
            [],
            [])
        {
            RelatedStories = [endnoteStory]
        };

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxPlacedRelatedStoryLayout placedStory = layout.Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .Single(story => story.StoryLayout.Story.Id == "25");
        DocxTextLineLayout[] lines = placedStory.TextLines.ToArray();

        TestAssert.Equal(-1, placedStory.SourceBlockIndex);
        TestAssert.True(Math.Abs(placedStory.Width - 200d) < 0.001d, "The document-end endnote should inherit the final section body width.");
        TestAssert.True(lines.Length >= 2, "Document-end endnotes should wrap using the target page width, not the first section width.");
        TestAssert.True(lines.All(line => line.Width <= placedStory.Width + 0.001d), "Document-end endnote lines should not exceed the final section body width.");
    }

    public static void DocxLayoutWrapsSectEndEndnoteUsingOwningSectionPageWidth()
    {
        DocxPageSettings wideFirstSection = new(
            "12240",
            "12000",
            null,
            "1440",
            "1440",
            "1440",
            "1440",
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null);
        DocxPageSettings narrowFinalSection = new(
            "4400",
            "6000",
            null,
            "200",
            "200",
            "200",
            "200",
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null)
        {
            EndnoteReferenceSettings = DocxNoteReferenceSettings.Empty with { PositionValue = "sectEnd" }
        };
        DocxParagraph anchor = DocxTests.CreateDocxLayoutParagraph("Narrow section-end endnote marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Endnote,
                    "26",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 7)
            ]
        };
        DocxParagraph endnoteParagraph = DocxTests.CreateDocxLayoutParagraph("alpha beta gamma delta epsilon zeta eta theta", 10d, 12d);
        var endnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Endnote,
            "/word/endnotes.xml",
            "26",
            [new DocxParagraphElement(endnoteParagraph)],
            [],
            [], null);
        var document = new DocxDocument(
            612d,
            792d,
            72d,
            72d,
            72d,
            72d,
            narrowFinalSection,
            [],
            [],
            [],
            [
                new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Wide first section", 10d, 12d)),
                new DocxSectionBreakElement(wideFirstSection, DocxSectionBreakType.NextPage, null, null, null, []),
                new DocxParagraphElement(anchor)
            ],
            [],
            [])
        {
            RelatedStories = [endnoteStory]
        };

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxPlacedRelatedStoryLayout placedStory = layout.Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .Single(story => story.StoryLayout.Story.Id == "26");
        DocxTextLineLayout[] lines = placedStory.TextLines.ToArray();

        TestAssert.Equal(2, placedStory.SourceBlockIndex);
        TestAssert.True(Math.Abs(placedStory.Width - 200d) < 0.001d, "The section-end endnote should inherit the owning section body width.");
        TestAssert.True(lines.Length >= 2, "Section-end endnotes should wrap using the target section width, not the first section width.");
        TestAssert.True(lines.All(line => line.Width <= placedStory.Width + 0.001d), "Section-end endnote lines should not exceed the owning section body width.");
    }

    public static void DocxLayoutKeepsSectEndEndnoteOverflowPageBeforeFollowingSection()
    {
        DocxParagraph firstSectionParagraph = DocxTests.CreateDocxLayoutParagraph("first section endnote marker", 10d, 18d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Endnote,
                    "22",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 6),
                new DocxInlineReference(
                    DocxRelatedStoryKind.Endnote,
                    "23",
                    CustomMarkFollowsValue: null,
                    DisplayText: "2",
                    SourceRunIndex: 0,
                    RunChildIndex: 2,
                    TextOffsetInRun: 12)
            ]
        };
        DocxParagraph secondSectionParagraph = DocxTests.CreateDocxLayoutParagraph("second section body", 10d, 12d);
        DocxParagraph endnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Endnote overflow body", 10d, 18d);
        var firstEndnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Endnote,
            "/word/endnotes.xml",
            "22",
            Enumerable.Range(0, 4).Select(_ => new DocxParagraphElement(endnoteParagraph)).Cast<DocxBodyElement>().ToArray(),
            [],
            [], null);
        var secondEndnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Endnote,
            "/word/endnotes.xml",
            "23",
            Enumerable.Range(0, 4).Select(_ => new DocxParagraphElement(endnoteParagraph)).Cast<DocxBodyElement>().ToArray(),
            [],
            [], null);
        DocxPageSettings firstSectionSettings = DocxPageSettings.Empty with
        {
            EndnoteReferenceSettings = DocxNoteReferenceSettings.Empty with { PositionValue = "sectEnd" }
        };
        var document = new DocxDocument(
            220d,
            112d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [
                new DocxParagraphElement(firstSectionParagraph),
                new DocxSectionBreakElement(firstSectionSettings, DocxSectionBreakType.NextPage, null, null, null, []),
                new DocxParagraphElement(secondSectionParagraph)
            ],
            [firstSectionParagraph, secondSectionParagraph],
            [])
        {
            RelatedStories = [firstEndnoteStory, secondEndnoteStory]
        };

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None));
        int secondSectionPageIndex = snapshot.Pages
            .Select((page, pageIndex) => (page, pageIndex))
            .First(item => item.page.Items.Any(pageItem => pageItem.SourceBlockIndex == 2))
            .pageIndex;
        int overflowEndnotePageIndex = snapshot.Pages
            .Select((page, pageIndex) => (page, pageIndex))
            .Where(item => item.page.PlacedEndnoteStoryCount == 1)
            .Select(item => item.pageIndex)
            .Last();

        TestAssert.True(snapshot.Pages.Count >= 3, "Multiple sectEnd endnotes that cannot all fit on the section end page should create a section-owned continuation page.");
        TestAssert.True(overflowEndnotePageIndex > 0 && overflowEndnotePageIndex < secondSectionPageIndex, "Section-end endnote overflow must be inserted before the following section instead of falling back to document-end placement.");
        TestAssert.Equal("sectEnd", snapshot.Pages[overflowEndnotePageIndex].SectionEndnotePositionValue ?? string.Empty);
        TestAssert.True(snapshot.Pages[overflowEndnotePageIndex].PlacedRelatedStories.Any(story => story.Kind == "Endnote" && story.SourceBlockIndex == 0), "The inserted section-end continuation page should retain marker-owned endnote provenance.");
        TestAssert.Equal(0, snapshot.Pages[secondSectionPageIndex].PlacedEndnoteStoryCount);
    }

    public static void RelatedStoryPageIndexMatchesLegacyMatching()
    {
        // R12: the once-per-pass page index answers reference/block queries exactly
        // like the legacy per-check page walks: null-tolerance and identity paragraph
        // matching, offset and run-fallback segment checks, vacuous truth for
        // never-rendered runs, and last-page section ranges.
        DocxParagraph source = MakeNoteParagraph();
        DocxParagraph other = MakeNoteParagraph();
        DocxTextLineLayout lineA = MakeOwnerLine(null, 0, 7, "hello");
        DocxTextLineLayout lineB = MakeOwnerLine(other, 0, 7, "world");
        DocxTextLineLayout lineC = MakeOwnerLine(null, 1, 7, "!");
        DocxLayoutPage page0 = MakeOwnerPage([lineA]);
        DocxLayoutPage page1 = MakeOwnerPage([lineB, lineC]);
        DocxLayoutEngine.RelatedStoryPageIndex index = DocxLayoutEngine.RelatedStoryPageIndex.Build([page0, page1], CancellationToken.None);

        DocxInlineReference reference = new(DocxRelatedStoryKind.Footnote, "1", null, null, SourceRunIndex: 7, RunChildIndex: 0, TextOffsetInRun: 2);
        var location = new DocxLayoutEngine.DocxInlineReferenceLocation(0, source, reference);
        TestAssert.True(index.IsReferenceRenderedOnPage(0, location), "Null-tolerance line must match any paragraph of its block.");
        TestAssert.True(!index.IsReferenceRenderedOnPage(1, location), "Identity mismatch and block mismatch must reject page 1.");
        TestAssert.Equal(0, index.FindFirstPageWithReference(location));

        var block1 = new DocxLayoutEngine.DocxInlineReferenceLocation(1, source, reference with { TextOffsetInRun = 0 });
        TestAssert.True(!index.IsReferenceRenderedOnPage(0, block1), "Block mismatch must reject page 0.");
        TestAssert.True(index.IsReferenceRenderedOnPage(1, block1), "Null-tolerance line must match block 1 on page 1.");

        var missingBlock = new DocxLayoutEngine.DocxInlineReferenceLocation(9, source, reference);
        TestAssert.True(index.IsReferenceRenderedOnPage(0, missingBlock), "Blocks rendered nowhere stay vacuously true (legacy).");
        TestAssert.True(index.IsReferenceRenderedOnPage(1, missingBlock), "Blocks rendered nowhere stay vacuously true (legacy).");
        TestAssert.Equal(0, index.FindFirstPageWithReference(missingBlock));

        var unrenderedRun = new DocxLayoutEngine.DocxInlineReferenceLocation(0, source, reference with { SourceRunIndex = 99 });
        TestAssert.True(index.IsReferenceRenderedOnPage(1, unrenderedRun), "Never-rendered runs stay vacuously true (legacy).");
        var negativeRun = new DocxLayoutEngine.DocxInlineReferenceLocation(0, source, reference with { SourceRunIndex = -1 });
        TestAssert.True(index.IsReferenceRenderedOnPage(1, negativeRun), "Negative runs stay vacuously true (legacy).");

        int[] page0Blocks = index.SortedBlocks(0);
        TestAssert.Equal(1, page0Blocks.Length);
        TestAssert.Equal(0, page0Blocks[0]);
        int[] page1Blocks = index.SortedBlocks(1);
        TestAssert.Equal(2, page1Blocks.Length);
        TestAssert.Equal(0, page1Blocks[0]);
        TestAssert.Equal(1, page1Blocks[1]);
        TestAssert.Equal(1, index.FindLastPageWithBlockInRange(0, 0));
        TestAssert.Equal(1, index.FindLastPageWithBlockInRange(0, 1));
        TestAssert.Equal(-1, index.FindLastPageWithBlockInRange(5, 9));
    }

    private static DocxParagraph MakeNoteParagraph()
    {
        return new DocxParagraph(
            [],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1.2d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
    }

    private static DocxTextLineLayout MakeOwnerLine(DocxParagraph? paragraph, int block, int run, string text)
    {
        var styleRun = new DocxTextRun(text, 10d, null, false, false, false, null, null);
        var segment = new DocxTextSegmentLayout(
            text,
            styleRun,
            0d,
            text.Length * 5d,
            10d,
            0d,
            0d,
            default,
            false,
            run,
            0,
            DocxTextSegmentRole.Text);
        return new DocxTextLineLayout(
            Text: text,
            StyleRun: styleRun,
            FontSize: 10d,
            X: 0d,
            BaselineY: 10d,
            Width: text.Length * 5d,
            Segments: [segment],
            SourceBlockIndex: block,
            SourceParagraphIndex: null,
            SourceLineIndex: null,
            Story: null,
            LineHeight: null,
            AppliedBeforeSpacing: null,
            IsFirstParagraphLine: null,
            EndsWithIntraTokenBreak: false,
            SingleLineHeight: null,
            ListLabelSingleLineHeight: null,
            BodyWindowsLineHeight: null,
            ListLabelWindowsLineHeight: null,
            EffectiveLineSpacingFactor: null,
            LineSpacingFactorFloorApplied: null,
            PendingAfterSpacing: null,
            ParagraphBeforeSpacing: null,
            ParagraphAfterSpacing: null,
            ContextualSpacingSuppressed: null,
            SourceParagraph: paragraph,
            LineHeightSource: null,
            EmitsTerminalParagraphMark: false);
    }

    private static DocxLayoutPage MakeOwnerPage(IReadOnlyList<DocxTextLineLayout> lines)
    {
        return new DocxLayoutPage(
            612d,
            792d,
            72d,
            72d,
            0d,
            72d,
            72d,
            DocxPageSettings.Empty,
            new DocxSectionLayoutProperties(null, null, null, null, null, null, []),
            [],
            [],
            [],
            [],
            [],
            lines);
    }

    public static void DocxFootnoteTrailingLinesShareDrawableLinePositions()
    {
        // RV06 footnote-align probe: Office centers/rights drawable footnote text and
        // keeps one row-end space beyond authored trailing in footnotes too.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">F trail   </w:t></w:r></w:p><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">F trail</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxTextLineLayout[] footnoteLines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .PlacedRelatedStories
            .SelectMany(story => story.TextLines)
            .ToArray();
        DocxTextLineLayout trailLine = footnoteLines.Single(line => line.Text.EndsWith("   ", StringComparison.Ordinal));
        DocxTextLineLayout cleanLine = footnoteLines.Single(line => line.Text == "F trail");
        TestAssert.Equal("F trail    ", trailLine.Text);
        TestAssert.Equal(cleanLine.X, trailLine.X);
    }

    public static void DocxFootnoteSeparatorMarkSitsAtRuleEnd()
    {
        // RV06 footnote probe: Office renders the separator mark as a space at the
        // rule end (body-left plus the 144pt rule width).
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxPlacedRelatedStoryLayout separator = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .PlacedRelatedStories
            .Single(story => story.SeparatorY is not null);
        DocxTextLineLayout separatorLine = separator.TextLines.Single();
        TestAssert.Equal(1, separatorLine.Segments.Count);
        TestAssert.Equal(" ", separatorLine.Segments[0].Text);
        TestAssert.Equal(144d, separatorLine.Segments[0].X - separatorLine.X);
    }

    private sealed class StrikeoutFontResolver(int strikeoutSize = 50, int strikeoutPosition = 300) : IFontResolver
    {
        public FontFaceResolution Resolve(FontRequest request)
        {
            byte[] faceBytes = TestFontBuilder.CreateTestFont();
            PatchStrikeoutMetrics(faceBytes, (ushort)strikeoutSize, (short)strikeoutPosition);
            return new FontFaceResolution(
                request.FamilyName,
                "StrikeFace",
                new FontStyleKey(request.Bold, request.Italic),
                new MemoryFontProgramSource("test:strikeface", faceBytes),
                IsFallback: false);
        }

        private static void PatchStrikeoutMetrics(byte[] faceBytes, ushort size, short position)
        {
            // The shared synthetic face stores a placeholder strikeout size, so pin
            // real OS/2 strikeout geometry (0.05em size, 0.30em position) in place.
            int tableCount = (faceBytes[4] << 8) | faceBytes[5];
            for (int tableIndex = 0; tableIndex < tableCount; tableIndex++)
            {
                int record = 12 + 16 * tableIndex;
                if (faceBytes[record] == 0x4F && faceBytes[record + 1] == 0x53 && faceBytes[record + 2] == 0x2F && faceBytes[record + 3] == 0x32)
                {
                    int offset = (faceBytes[record + 8] << 24) | (faceBytes[record + 9] << 16) | (faceBytes[record + 10] << 8) | faceBytes[record + 11];
                    faceBytes[offset + 26] = (byte)(size >> 8);
                    faceBytes[offset + 27] = (byte)(size & 0xFF);
                    faceBytes[offset + 28] = (byte)((position >> 8) & 0xFF);
                    faceBytes[offset + 29] = (byte)(position & 0xFF);
                    return;
                }
            }

            throw new InvalidOperationException("Synthetic test font is missing the OS/2 table.");
        }
    }

    public static void DocxFootnoteSeparatorRuleFollowsStrikeoutMetrics()
    {
        // RV06 separator probes (Word 16.0, Times, Aptos and Calibri at 10, 12 and 14pt):
        // the footnote separator rule follows OS/2 strikeout geometry, with the rule top
        // at the strikeout position and the thickness at the strikeout size, instead of
                // The patched synthetic face pins strikeout 0.30em and 0.05em,
        // so the rule bottom must sit 0.25em above the separator bottom while the mark baseline
        // rides at the separator bottom with no extra ride.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        var resolver = new StrikeoutFontResolver();
        DocxFontPlan fontPlan = DocxFontPlan.Create(document, resolver, CancellationToken.None);
        var measurer = new DocxFontPlanTextMeasurer(fontPlan, resolver.Resolve(new FontRequest("StrikeFace")), CancellationToken.None, resolver);
        DocxPlacedRelatedStoryLayout separator = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, measurer, CancellationToken.None)
            .Pages[0]
            .PlacedRelatedStories
            .Single(story => story.SeparatorY is not null);
        DocxTextLineLayout separatorLine = separator.TextLines.Single();
        double markSize = separatorLine.Segments[0].StyleRun.EffectiveProperties.FontSize;
        double ruleY = separator.SeparatorY ?? double.NaN;
        TestAssert.True(Math.Abs(ruleY - separatorLine.BaselineY - 0.25d * markSize) < 0.000001d, "Footnote rule bottom must sit 0.25em above the mark baseline with no ride.");
        TestAssert.True(Math.Abs(separator.SeparatorThickness - 0.05d * markSize) < 0.000001d, "Footnote rule thickness must follow the strikeout size.");
    }

    public static void DocxFootnoteSeparatorRuleSnapsThicknessToPixelGrid()
    {
        // RV06 rule-thickness probes (Word 16.0, Times grids plus Tahoma slash Calibri 14pt):
        // Office rule thickness snaps strikeout size to whole 600dpi pixels (seven points exact),
        // so a 0.067em synthetic face at explicit 10pt marks renders 0.60 instead of 0.67,
                // The patched synthetic face pins strikeout 0.25em position and 0.067em size at
        // explicit 10pt marks, so the thickness must snap 0.67 to the 5-pixel 0.60.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/><Relationship Id="rIdS1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/styles.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        var resolver = new StrikeoutFontResolver(67, 250);
        DocxFontPlan fontPlan = DocxFontPlan.Create(document, resolver, CancellationToken.None);
        var measurer = new DocxFontPlanTextMeasurer(fontPlan, resolver.Resolve(new FontRequest("StrikeFace")), CancellationToken.None, resolver);
        DocxPlacedRelatedStoryLayout separator = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, measurer, CancellationToken.None)
            .Pages[0]
            .PlacedRelatedStories
            .Single(story => story.SeparatorY is not null);
        DocxTextLineLayout separatorLine = separator.TextLines.Single();
        double markSize = separatorLine.Segments[0].StyleRun.EffectiveProperties.FontSize;
        double ruleY = separator.SeparatorY ?? double.NaN;
        TestAssert.True(Math.Abs(markSize - 10d) < 0.000001d, "Separator marks must resolve at explicit 10pt; markSize=" + markSize.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(separator.SeparatorThickness - 0.72d) < 0.000001d, "Rule thickness must snap 0.67 to the 6-pixel 0.72; thickness=" + separator.SeparatorThickness.ToString(CultureInfo.InvariantCulture) + ".");
    }

    public static void DocxFootnoteSeparatorRuleKeepsLegacyConstantsWithoutStrikeoutMetrics()
    {
        // Measurers without strikeout metrics keep the legacy footnote constants.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxPlacedRelatedStoryLayout separator = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .PlacedRelatedStories
            .Single(story => story.SeparatorY is not null);
        DocxTextLineLayout separatorLine = separator.TextLines.Single();
        double ruleY = separator.SeparatorY ?? double.NaN;
        TestAssert.True(Math.Abs(ruleY - separatorLine.BaselineY - 2.1d) < 0.000001d, "Legacy footnote rule offset must stay 2.1pt above the separator bottom with no ride.");
        TestAssert.Equal(0.75d, separator.SeparatorThickness);
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
        var measurer = new DocxFontPlanTextMeasurer(fontPlan, resolver.Resolve(new FontRequest("StrikeFace")), CancellationToken.None, resolver);
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

    public static void DocxWordCompatibleAllMarkupMapsSeparatorRuleThickness()
    {
        // RV06 endnote probes: word-compatible separator rules map uniformly to
        // emission space like the surrounding story text, but the rule thickness
        // still emits at design size while the rule width is print-scaled.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r><w:r><w:commentRangeStart w:id="0"/></w:r><w:r><w:t xml:space="preserve">commented</w:t></w:r><w:r><w:commentRangeEnd w:id="0"/></w:r><w:r><w:commentReference w:id="0"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body</w:t></w:r></w:p></w:footnote></w:footnotes>""",
            ["word/comments.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="0" w:author="Reviewer" w:initials="R" w:date="2024-01-02T00:00:00Z"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:comment></w:comments>"""
        });;
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.AllMarkup);
        }
        var resolver = new StrikeoutFontResolver();
        DocxFontPlan fontPlan = DocxFontPlan.Create(document, resolver, CancellationToken.None);
        var measurer = new DocxFontPlanTextMeasurer(fontPlan, resolver.Resolve(new FontRequest("StrikeFace")), CancellationToken.None, resolver);
        DocxPlacedRelatedStoryLayout separator = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .Create(document, measurer, CancellationToken.None)
            .Pages[0]
            .PlacedRelatedStories
            .Single(story => story.SeparatorY is not null);
        double designThickness = separator.SeparatorThickness;
        double designWidth = Math.Min(144d, separator.Width);
        var renderer = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        PdfPage page = renderer.RenderBlankPages(document, null, CancellationToken.None).Single();
        Match[] ruleMatches = Regex.Matches(page.Content, @"(?<x>-?[0-9.]+) (?<y>-?[0-9.]+) (?<w>[0-9.]+) (?<h>[0-9.]+) re f[^*]")
            .Cast<Match>()
            .Where(match => double.Parse(match.Groups["x"].Value, CultureInfo.InvariantCulture) < 200d && double.Parse(match.Groups["w"].Value, CultureInfo.InvariantCulture) > 50d)
            .ToArray();
        TestAssert.Equal(1, ruleMatches.Length);
        double emitWidth = double.Parse(ruleMatches[0].Groups["w"].Value, CultureInfo.InvariantCulture);
        double emitHeight = double.Parse(ruleMatches[0].Groups["h"].Value, CultureInfo.InvariantCulture);
        double printScale = emitWidth / designWidth;
        TestAssert.True(printScale < 0.999d, "Separator rule width must carry the print scale, proving the emission map is active.");
        TestAssert.True(Math.Abs(emitHeight - designThickness * printScale) < 0.002d, "Word-compatible separator rule thickness must scale with the print map.");
    }



    public static void DocxWordCompatibleContinuationRuleSpansDesignBodyWidth()
    {
        // RV06 wclong probe (Word 16.0): word-compatible continuation rules span the
        // full design body mapped once (Office 354.98 emitted), while the renderer
        // emits the shrunk layout body mapped again (269.49 = shrunk times scale).
        // The continuation rule width must be the design body (layout body plus the
        // markup reserve) with the end mark at the rule end.
        var footnoteParas = new StringBuilder();
        for (int line = 0; line < 60; line++)
        {
            footnoteParas.Append("<w:p><w:r><w:t>Note body line " + line.ToString(CultureInfo.InvariantCulture) + " words here</w:t></w:r></w:p>");
        }
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r><w:r><w:commentRangeStart w:id="0"/></w:r><w:r><w:t xml:space="preserve">commented</w:t></w:r><w:r><w:commentRangeEnd w:id="0"/></w:r><w:r><w:commentReference w:id="0"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2">""" + footnoteParas.ToString() + """</w:footnote></w:footnotes>""",
            ["word/comments.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="0" w:author="Reviewer" w:initials="R" w:date="2024-01-02T00:00:00Z"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:comment></w:comments>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.AllMarkup);
        }
        var resolver = new StrikeoutFontResolver();
        DocxFontPlan fontPlan = DocxFontPlan.Create(document, resolver, CancellationToken.None);
        var measurer = new DocxFontPlanTextMeasurer(fontPlan, resolver.Resolve(new FontRequest("StrikeFace")), CancellationToken.None, resolver);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .Create(document, measurer, CancellationToken.None);
        var continuations = layout.Pages.SelectMany(page => page.PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                story.StoryLayout.Story.Type == DocxRelatedStoryType.ContinuationSeparator &&
                story.SeparatorY is not null)
            .Select(story => (Page: page, Story: story))).ToArray();
        TestAssert.True(continuations.Length != 0, "Overflowing WC footnotes must place continuation separators.");
        foreach ((DocxLayoutPage page, DocxPlacedRelatedStoryLayout continuation) in continuations)
        {
            double designBody = page.Width - page.MarginLeft - page.MarginRight + page.MarkupMarginReservePoints;
            TestAssert.True(Math.Abs(continuation.SeparatorWidth - designBody) < 0.01d, "WC continuation rule width must be the design body width.");
            TestAssert.True(Math.Abs(continuation.Width - designBody) < 0.01d, "WC continuation story width must cover the design body for the rule clip.");
            if (continuation.TextLines.Count == 1 && continuation.TextLines[0].Segments.Count == 1 &&
                string.IsNullOrWhiteSpace(continuation.TextLines[0].Segments[0].Text))
            {
                double markX = continuation.TextLines[0].Segments[0].X;
                TestAssert.True(Math.Abs(markX - (continuation.X + designBody)) < 0.01d, "WC continuation end mark must sit at the rule end.");
            }
        }
        // Reserve-margin pages keep identity emission over the shrunk body, so the
        // reserve must not join the continuation rule back there.
        DocxLayout reserveLayout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin)
            .Create(document, measurer, CancellationToken.None);
        var reserveContinuations = reserveLayout.Pages.SelectMany(page => page.PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                story.StoryLayout.Story.Type == DocxRelatedStoryType.ContinuationSeparator &&
                story.SeparatorY is not null)
            .Select(story => (Page: page, Story: story))).ToArray();
        TestAssert.True(reserveContinuations.Length != 0, "Overflowing reserve-margin footnotes must place continuation separators.");
        foreach ((DocxLayoutPage reservePage, DocxPlacedRelatedStoryLayout reserveContinuation) in reserveContinuations)
        {
            double shrunkBody = reservePage.Width - reservePage.MarginLeft - reservePage.MarginRight;
            TestAssert.True(Math.Abs(reserveContinuation.SeparatorWidth - shrunkBody) < 0.01d, "Reserve-margin continuation rule width must stay in the shrunk body.");
        }
    }

    public static void DocxLongFootnoteContinuationReservesHeaderZone()
    {
        // RV06 wclong header probes (Word 16.0, h1/h5/nohdr): continued notes leave
        // the static header zone clear (h5 takes 22 vs 25 blind with the rule below
        // all header baselines; nohdr agrees 25/25), while the renderer starts
        // continuation pages at the full page top so the rule lands inside the
        // header text. Every continuation rule must sit below the header baselines
        // of its own page.
        var headerParas = new StringBuilder();
        for (int line = 0; line < 5; line++)
        {
            headerParas.Append("<w:p><w:r><w:t>Header line " + line.ToString(CultureInfo.InvariantCulture) + " words here</w:t></w:r></w:p>");
        }
        var footnoteParas = new StringBuilder();
        for (int line = 0; line < 60; line++)
        {
            footnoteParas.Append("<w:p><w:r><w:t>Note body line " + line.ToString(CultureInfo.InvariantCulture) + " words here</w:t></w:r></w:p>");
        }
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/header1.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">""" + headerParas.ToString() + """</w:hdr>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2">""" + footnoteParas.ToString() + """</w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        var continuations = layout.Pages.SelectMany(page => page.PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                story.StoryLayout.Story.Type == DocxRelatedStoryType.ContinuationSeparator &&
                story.SeparatorY is not null)
            .Select(story => (Page: page, Story: story))).ToArray();
        TestAssert.True(continuations.Length != 0, "Overflowing footnotes with headers must place continuation separators.");
        foreach ((DocxLayoutPage page, DocxPlacedRelatedStoryLayout continuation) in continuations)
        {
            double[] headerBaselines = page.StaticTextLines.Select(line => line.BaselineY).ToArray();
            TestAssert.True(headerBaselines.Length != 0, "Continuation pages must carry the static header.");
            double lowestHeaderBaseline = headerBaselines.Min();
            TestAssert.True(continuation.SeparatorY <= lowestHeaderBaseline - 1d, "Continuation rules must sit below the header baselines of their page.");
        }
    }
    public static void DocxLongFootnoteHeadTakeMatchesOfficeWithAptosMetrics()
    {
        // RV06 p1-clamp probes (Word 16.0, h1/b-sweep): the head page takes 24 note
        // lines (2 mixed-size plus long 0..21) with full-width rules below the body,
        // while the renderer takes 23 (long 0..20) with the rule 3.73 low. Synthetic
        // metrics take frame-exact counts that cannot see the real-advance shortfall,
        // so this pins the Office take with installed Aptos advances and skips without
        // a usable Aptos face like other font-environmental tests.
        if (!HasUsableAptosFootnoteFace())
        {
            TestAssert.Skip("Environmental precondition not met: Aptos is not installed.");
        }

        var footnoteParas = new StringBuilder();
        footnoteParas.Append("<w:p><w:r><w:t xml:space=\"preserve\">Footnote twelve opening words</w:t></w:r></w:p>");
        footnoteParas.Append("<w:p><w:r><w:rPr><w:b/><w:sz w:val=\"30\"/><w:szCs w:val=\"30\"/></w:rPr><w:t xml:space=\"preserve\">Footnote fifteen second line</w:t></w:r></w:p>");
        for (int line = 0; line < 60; line++)
        {
            footnoteParas.Append("<w:p><w:r><w:t>Long note line " + line.ToString(CultureInfo.InvariantCulture) + " with words to fill pages.</w:t></w:r></w:p>");
        }
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body text with a footnote reference</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r><w:r><w:t xml:space="preserve"> and </w:t></w:r><w:commentRangeStart w:id="1"/><w:r><w:t xml:space="preserve">a review note</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r><w:r><w:t xml:space="preserve"> trailing words.</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/header1.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t xml:space="preserve">Header twelve opening words</w:t></w:r></w:p></w:hdr>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2">""" + footnoteParas.ToString() + """</w:footnote></w:footnotes>""",
            ["word/comments.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="1" w:author="Reviewer" w:date="2026-06-01T00:00:00Z"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:comment></w:comments>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        var resolver = new WindowsFontResolver();
        DocxFontPlan fontPlan = DocxFontPlan.Create(document, resolver, CancellationToken.None);
        var measurer = new DocxFontPlanTextMeasurer(fontPlan, resolver.Resolve(new FontRequest("Aptos")), CancellationToken.None, resolver);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, measurer, CancellationToken.None);
        var headLines = layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .ToArray();
        var longIndexes = headLines
            .Select(line => line.Text)
            .Select(text => text.StartsWith("Long note line ", StringComparison.Ordinal) && int.TryParse(text.Substring(15).Split(' ')[0], out int index) ? index : -1)
            .Where(index => index >= 0)
            .ToArray();
        DocxTextLineLayout firstLine = headLines[0];        string diagnosis = "headLines=" + headLines.Length + ",long=[" + string.Join(",", longIndexes.Take(3)) + ".." + (longIndexes.Length == 0 ? "none" : longIndexes[^1].ToString(CultureInfo.InvariantCulture)) + "],firstLine(h=" + (firstLine.LineHeight?.ToString("F3", CultureInfo.InvariantCulture) ?? "?") + ",single=" + (firstLine.SingleLineHeight?.ToString("F3", CultureInfo.InvariantCulture) ?? "?") + ",factor=" + (firstLine.EffectiveLineSpacingFactor?.ToString("F4", CultureInfo.InvariantCulture) ?? "?") + ",after=" + (firstLine.ParagraphAfterSpacing?.ToString("F3", CultureInfo.InvariantCulture) ?? "?") + ")";
        TestAssert.Equal(24, headLines.Length);
        TestAssert.True(longIndexes.Length == 22 && longIndexes[0] == 0 && longIndexes[^1] == 21, "Head page must take long lines 0..21 like Office. " + diagnosis);
    }

    public static void DocxLongFootnoteHeadBlockSitsAtOfficePosition()
    {
        // RV06 p1-position probes (Word 16.0, h1/b-sweep): with takes exact, the head
        // block must also sit at Office absolute positions (h1-Final rule 682.54 with
        // mixed1 at 664.18), while the renderer seats the whole block ~8 high on
        // capacity slack. Windows-only Aptos pins this like the take test.
        if (!HasUsableAptosFootnoteFace())
        {
            TestAssert.Skip("Environmental precondition not met: Aptos is not installed.");
        }

        var footnoteParas = new StringBuilder();
        footnoteParas.Append("<w:p><w:r><w:t xml:space=\"preserve\">Footnote twelve opening words</w:t></w:r></w:p>");
        footnoteParas.Append("<w:p><w:r><w:rPr><w:b/><w:sz w:val=\"30\"/><w:szCs w:val=\"30\"/></w:rPr><w:t xml:space=\"preserve\">Footnote fifteen second line</w:t></w:r></w:p>");
        for (int line = 0; line < 60; line++)
        {
            footnoteParas.Append("<w:p><w:r><w:t>Long note line " + line.ToString(CultureInfo.InvariantCulture) + " with words to fill pages.</w:t></w:r></w:p>");
        }
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body text with a footnote reference</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r><w:r><w:t xml:space="preserve"> and </w:t></w:r><w:commentRangeStart w:id="1"/><w:r><w:t xml:space="preserve">a review note</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r><w:r><w:t xml:space="preserve"> trailing words.</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/header1.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t xml:space="preserve">Header twelve opening words</w:t></w:r></w:p></w:hdr>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2">""" + footnoteParas.ToString() + """</w:footnote></w:footnotes>""",
            ["word/comments.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="1" w:author="Reviewer" w:date="2026-06-01T00:00:00Z"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:comment></w:comments>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        var resolver = new WindowsFontResolver();
        DocxFontPlan fontPlan = DocxFontPlan.Create(document, resolver, CancellationToken.None);
        var measurer = new DocxFontPlanTextMeasurer(fontPlan, resolver.Resolve(new FontRequest("Aptos")), CancellationToken.None, resolver);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, measurer, CancellationToken.None);
        DocxLayoutPage page = layout.Pages[0];
        DocxPlacedRelatedStoryLayout separator = page.PlacedRelatedStories.Single(story => story.SeparatorY is not null);
        DocxTextLineLayout mixedFirst = page.PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .First(line => line.Text.StartsWith("Footnote twelve opening", StringComparison.Ordinal));
        TestAssert.True(Math.Abs((separator.SeparatorY ?? double.NaN) - 682.54d) < 1.5d, "Head separator rule must sit at the Office position; observed rule bottom=" + (separator.SeparatorY?.ToString(CultureInfo.InvariantCulture) ?? "?") + ".");
        TestAssert.True(Math.Abs(mixedFirst.BaselineY - 664.18d) < 1.5d, "Head footnote content must start at the Office position; observed mixed1 baseline=" + mixedFirst.BaselineY.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private static bool HasUsableAptosFootnoteFace()
    {
        try
        {
            FontFaceResolution resolved = new WindowsFontResolver().Resolve(new FontRequest("Aptos"));
            if (resolved.IsFallback)
            {
                return false;
            }

            OpenTypeFont? font = FontProgramLoader.Load(resolved, CancellationToken.None);
            return font is not null && font.HasTrueTypeOutlines;
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or NotSupportedException or ArgumentOutOfRangeException or UnauthorizedAccessException)
        {
            return false;
        }
    }
    public static void DocxFootnoteStoriesTopHangsBelowBodyStart()
    {
        // RV06 p1-clamp probes (Word 16.0, b/ba/2para/bodyb sweeps): Office storiesTop
        // is independent of body size/after/length/before and follows body start with
        // footnote-side moves, while the renderer hung content one body em below the
        // lowest body baseline. Takes come from the same clamped capacity, but the
        // overflowing head block seats bottom-up (storiesTop = margin + takeHeight)
        // instead of leaving capacity slack above the margin.
        var footnoteParas = new StringBuilder();
        for (int line = 0; line < 60; line++)
        {
            footnoteParas.Append("<w:p><w:r><w:t>Note body line " + line.ToString(CultureInfo.InvariantCulture) + " words here</w:t></w:r></w:p>");
        }
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2">""" + footnoteParas.ToString() + """</w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxLayoutPage page = layout.Pages[0];
        DocxPlacedRelatedStoryLayout body = page.PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal));
        // RV06 line-box maximum: filler runs box the Windows extents (1.2em), so the clamped take counts fewer lines with bottom-up seating intact.
        TestAssert.Equal(25, body.TextLines.Count);
        TestAssert.True(Math.Abs(body.TopY - body.Height - page.MarginBottom) < 0.000001d, "The overflowing head block must sit bottom-up on the margin; observed bottom=" + (body.TopY - body.Height).ToString(CultureInfo.InvariantCulture) + ".");
    }

    public static void DocxFootnoteSeparatorGapFollowsSingleLineMetrics()
    {
        // RV06 separator-bottom probes (Word 16.0, Times/Aptos/Calibri): the Office
        // footnote gap above the body equals one single-spaced line box minus the
        // first-baseline inset, so it derives from the mark font instead of a fixed
        // 3pt constant. The patched synthetic face has a 1.15em single line, giving
        // a 0.21em gap.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        var resolver = new StrikeoutFontResolver();
        DocxFontPlan fontPlan = DocxFontPlan.Create(document, resolver, CancellationToken.None);
        var measurer = new DocxFontPlanTextMeasurer(fontPlan, resolver.Resolve(new FontRequest("StrikeFace")), CancellationToken.None, resolver);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, measurer, CancellationToken.None);
        DocxPlacedRelatedStoryLayout separator = layout.Pages[0].PlacedRelatedStories.Single(story => story.SeparatorY is not null);
        DocxPlacedRelatedStoryLayout body = layout.Pages[0].PlacedRelatedStories.Single(story => story.SeparatorY is null);
        double markSize = separator.TextLines.Single().Segments[0].StyleRun.EffectiveProperties.FontSize;
        double separatorBottom = separator.TopY - separator.Height;
        TestAssert.True(Math.Abs(separatorBottom - body.TopY - 0.21d * markSize) < 0.000001d, "Footnote separator gap must follow the single-line box minus the first-baseline inset.");
    }

    public static void DocxFootnoteSeparatorGapKeepsLegacyConstantWithoutLineMetrics()
    {
        // Measurers without single-line metrics keep the legacy 3pt footnote gap.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        var resolver = new StrikeoutFontResolver();
        DocxFontPlan fontPlan = DocxFontPlan.Create(document, resolver, CancellationToken.None);
        var measurer = new DocxFontPlanTextMeasurer(fontPlan, resolver.Resolve(new FontRequest("StrikeFace")), CancellationToken.None, resolver);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxPlacedRelatedStoryLayout separator = layout.Pages[0].PlacedRelatedStories.Single(story => story.SeparatorY is not null);
        DocxPlacedRelatedStoryLayout body = layout.Pages[0].PlacedRelatedStories.Single(story => story.SeparatorY is null);
        double separatorBottom = separator.TopY - separator.Height;
        TestAssert.True(Math.Abs(separatorBottom - body.TopY - 3d) < 0.000001d, "Legacy footnote separator gap must stay 3pt.");
    }

    public static void DocxLongFootnoteSplitsAcrossPagesWithContinuationSeparator()
    {
        // RV06 footlong probe (Word 16.0): a 60-paragraph footnote splits 24/25/11
        // across 3 pages with full-width continuation rules plus end marks, while the
        // renderer stacks every line on the reference page off-page with no
        // continuation story placed. Continuation separators parse but never place.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body line 0 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 1 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 2 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 3 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 4 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 5 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 6 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 7 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 8 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 9 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 10 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 11 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 12 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 13 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 14 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 15 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 16 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 17 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 18 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 19 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 20 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 21 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 22 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 23 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 24 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 25 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 26 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 27 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 28 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 29 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 30 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 31 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 32 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 33 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 34 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 35 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 36 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 37 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 38 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 39 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 40 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 41 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 42 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 43 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 44 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 45 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 46 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 47 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 48 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 49 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 50 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 51 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 52 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 53 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 54 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 55 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 56 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 57 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 58 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 59 words here</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        foreach (DocxLayoutPage page in layout.Pages)
        {
            foreach (DocxPlacedRelatedStoryLayout story in page.PlacedRelatedStories)
            {
                if (story.StoryLayout.Story.Kind != DocxRelatedStoryKind.Footnote || story.StoryLayout.Story.Type != null)
                {
                    continue;
                }

                foreach (DocxTextLineLayout line in story.TextLines)
                {
                    TestAssert.True(line.BaselineY >= page.MarginBottom - 1d && line.BaselineY <= page.Height - page.MarginTop + 1d, "Footnote lines must stay inside their page.");
                }
            }
        }
        IReadOnlyList<DocxPlacedRelatedStoryLayout> continuations = layout.Pages.Skip(1).SelectMany(page => page.PlacedRelatedStories).Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.ContinuationSeparator && story.SeparatorY is not null).ToArray();
        TestAssert.True(continuations.Count != 0, "Overflowing footnotes must place continuation separators on continuation pages.");
    }

    public static void DocxLongFootnoteWithoutSeparatorSplitsAcrossPages()
    {
        // RV06 footlong-absent probe (Word 16.0): separator-less footnotes (Word draws
        // default rules) still stack whole stories off-page on overflow. This pins
        // in-bounds lines plus continuation separators for the separator-less path.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body line 0 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 1 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 2 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 3 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 4 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 5 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 6 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 7 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 8 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 9 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 10 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 11 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 12 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 13 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 14 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 15 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 16 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 17 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 18 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 19 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 20 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 21 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 22 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 23 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 24 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 25 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 26 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 27 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 28 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 29 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 30 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 31 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 32 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 33 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 34 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 35 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 36 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 37 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 38 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 39 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 40 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 41 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 42 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 43 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 44 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 45 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 46 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 47 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 48 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 49 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 50 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 51 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 52 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 53 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 54 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 55 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 56 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 57 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 58 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 59 words here</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        var bodyLines = layout.Pages.SelectMany(page => page.PlacedRelatedStories).Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal)).SelectMany(story => story.TextLines).ToArray();
        TestAssert.True(bodyLines.Length == 60, "All separator-less footnote lines must be placed.");
        TestAssert.True(bodyLines.All(line => layout.Pages.Any(page => line.BaselineY >= page.MarginBottom - 1d && line.BaselineY <= page.Height - page.MarginTop + 1d)), "Every separator-less footnote line must sit inside some page.");
        int contRules = layout.Pages.SelectMany(page => page.PlacedRelatedStories).Count(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.ContinuationSeparator && story.SeparatorY is not null);
        TestAssert.True(contRules != 0, "Overflowing separator-less footnotes must place continuation separators.");
    }

    public static void DocxSplitEndnotesKeepEachLineOnExactlyOnePage()
    {
        // RV06 endlong probe (Word 16.0): a 60-paragraph section-end endnote splits
        // 23/25/12 with full-width continuation rules, one text op per line, while the
        // renderer triplicates every line across 3 slice pages (180 ops for 60 lines)
        // because slices shift without filtering and only clip at emission.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:endnotePr><w:pos w:val="sectEnd"/></w:endnotePr></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body line 0 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 1 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 2 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 3 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 4 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 5 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 6 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 7 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 8 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 9 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 10 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 11 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 12 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 13 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 14 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 15 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 16 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 17 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 18 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 19 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 20 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 21 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 22 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 23 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 24 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 25 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 26 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 27 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 28 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 29 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 30 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 31 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 32 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 33 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 34 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 35 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 36 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 37 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 38 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 39 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 40 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 41 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 42 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 43 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 44 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 45 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 46 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 47 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 48 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 49 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 50 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 51 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 52 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 53 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 54 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 55 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 56 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 57 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 58 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 59 words here</w:t></w:r></w:p></w:endnote></w:endnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        TestAssert.True(layout.Pages.Count > 1, "Overflowing endnotes must paginate.");
        var lineKeys = layout.Pages.SelectMany(page => page.PlacedRelatedStories).Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && story.StoryLayout.Story.Type != DocxRelatedStoryType.Separator && story.StoryLayout.Story.Type != DocxRelatedStoryType.ContinuationSeparator).SelectMany(story => story.TextLines).Select(line => (line.SourceBlockIndex, line.SourceParagraphIndex, line.SourceLineIndex)).ToArray();
        TestAssert.Equal(lineKeys.Distinct().Count(), lineKeys.Length);
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

    public static void DocxFootnoteRemainderAccountsSeparatorUntilPlaced()
    {
        DocxRelatedStory story = new(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "31",
            [],
            [],
            [], null);
        var storyLayout = new DocxRelatedStoryLayout(story, 0, [], [], [], [], 0d);
        var location = new DocxLayoutEngine.DocxInlineReferenceLocation(
            0,
            DocxTests.CreateDocxLayoutParagraph("Anchor", 10d, 12d),
            new DocxInlineReference(DocxRelatedStoryKind.Footnote, "31", null, "1", 0, 1, 4));
        var inFlight = new DocxLayoutEngine.InFlightRelatedStory(storyLayout, null, null, 0, location);
        double[] lineBoxes = [10d, 10d, 10d, 10d, 10d];

        double remainder = DocxLayoutEngine.FootnoteRemainderHeight(inFlight, lineBoxes, null);

        TestAssert.Equal(53d, remainder);
    }

    public static void DocxFootnoteRemainderShrinksWithPlacedTakes()
    {
        DocxRelatedStory story = new(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "32",
            [],
            [],
            [], null);
        DocxTextRun lineRun = new("Shared note", 10d, null, false, false, false, null, null);
        DocxTextLineLayout RemainderLine(int index)
        {
            return new DocxTextLineLayout("Shared note line " + index.ToString(CultureInfo.InvariantCulture), lineRun, 10d, 72d, 700d - (12d * index), 200d, [], null, null, null, null, 12d, null, null, false, null, null, null, null, null, null, null, null, null, null, null, null, false);
        }

        var storyLayout = new DocxRelatedStoryLayout(story, 0, [RemainderLine(0), RemainderLine(1), RemainderLine(2), RemainderLine(3), RemainderLine(4)], [], [], [], 60d);
        var location = new DocxLayoutEngine.DocxInlineReferenceLocation(
            0,
            DocxTests.CreateDocxLayoutParagraph("Anchor", 10d, 12d),
            new DocxInlineReference(DocxRelatedStoryKind.Footnote, "32", null, "1", 0, 1, 4));
        var inFlight = new DocxLayoutEngine.InFlightRelatedStory(storyLayout, null, null, 0, location)
        {
            PlacedLineCount = 2,
            SeparatorPlaced = true
        };
        double[] lineBoxes = [10d, 10d, 10d, 10d, 10d];

        double remainder = DocxLayoutEngine.FootnoteRemainderHeight(inFlight, lineBoxes, null);

        TestAssert.Equal(30d, remainder);
        TestAssert.Equal(3, inFlight.RemainingLineCount);
        TestAssert.Equal(30d, inFlight.RemainingContentHeight(lineBoxes));
    }

    public static void DocxNarrowedFootnoteRemainderRebasesAtStoryOrigin()
    {
        DocxTextRun run = new("Shared note", 10d, null, false, false, false, null, null);
        DocxTextLineLayout Line(string text, double baseline)
        {
            return new DocxTextLineLayout(text, run, 10d, 72d, baseline, 200d, [], null, null, null, null, 12d, null, null, false, null, null, null, null, null, null, null, null, null, null, null, null, false);
        }

        DocxRelatedStory story = new(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "33",
            [],
            [],
            [], null);
        var storyLayout = new DocxRelatedStoryLayout(
            story, 0, [Line("Shared note line zero", 700d), Line("Shared note line one", 688d), Line("Shared note line two", 676d)], [], [], [], 36d);

        DocxRelatedStoryLayout narrowed = DocxLayoutEngine.NarrowStoryTextLinesForOffset(storyLayout, 1);

        TestAssert.Equal(2, narrowed.TextLines.Count);
        TestAssert.Equal(700d, narrowed.TextLines[0].BaselineY);
        TestAssert.Equal(688d, narrowed.TextLines[1].BaselineY);
        TestAssert.Equal(24d, narrowed.ContentHeight);
    }

    public static void DocxSharedFootnotePagesKeepBodyAboveNotes()
    {
        DocxDocument document = CreateInterleavingDocument(noteParagraphCount: 40, fillerParagraphCount: 60, footnoteId: "41");
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout);

        string[] placedTexts = layout.Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Select(line => line.Text)
            .ToArray();
        TestAssert.Equal(40, placedTexts.Length);
        TestAssert.Equal(40, placedTexts.Distinct().Count());

        foreach (DocxLayoutPageSnapshot page in snapshot.Pages)
        {
            if (page.Items.Count == 0 || page.PlacedFootnoteStoryCount == 0)
            {
                continue;
            }

            double bodyBottom = page.Items.Min(item => item.Y);
            double footnoteTop = page.PlacedRelatedItems.Max(item => item.Y + item.Height);
            TestAssert.True(bodyBottom >= footnoteTop, "Shared pages must keep body above placed notes.");
        }
    }

    public static void DocxExhaustedFootnoteDrainUsesDedicatedPages()
    {
        DocxDocument document = CreateInterleavingDocument(noteParagraphCount: 200, fillerParagraphCount: 0, footnoteId: "42");
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        string[] placedTexts = layout.Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Select(line => line.Text)
            .ToArray();
        TestAssert.True(layout.Pages.Count > 2, "The long note must span several pages.");
        TestAssert.Equal(200, placedTexts.Length);
        TestAssert.Equal(200, placedTexts.Distinct().Count());
        TestAssert.Equal(0, layout.Pages[^1].Items.Count);
        TestAssert.True(layout.Pages[^1].PlacedRelatedStories.Count != 0, "The exhausted drain must append dedicated note pages.");
    }

    public static void DocxMidBodyFootnoteAnchorKeepsPrecedingFillersOnPage()
    {
        const int fillersBefore = 30;
        DocxDocument document = CreateMidBodyInterleavingDocument(fillersBefore, noteParagraphCount: 60, footnoteId: "43", fillersAfter: 30);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        int PageWithBlock(int blockIndex)
        {
            for (int pageIndex = 0; pageIndex < layout.Pages.Count; pageIndex++)
            {
                foreach (DocxLayoutItem item in layout.Pages[pageIndex].Items)
                {
                    if (item is DocxTextLineLayout line && line.SourceBlockIndex == blockIndex)
                    {
                        return pageIndex;
                    }
                }
            }

            return -1;
        }

        TestAssert.True(PageWithBlock(fillersBefore - 1) >= 0, "Preceding fillers must be placed.");
        TestAssert.Equal(PageWithBlock(fillersBefore - 1), PageWithBlock(fillersBefore));

        string[] placedTexts = layout.Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Select(line => line.Text)
            .ToArray();
        TestAssert.Equal(60, placedTexts.Length);
        TestAssert.Equal(60, placedTexts.Distinct().Count());
    }

    public static void DocxTableAnchoredFootnoteSharesBodyPages()
    {
        DocxParagraph cellAnchor = DocxTests.CreateDocxLayoutParagraph("Cell note marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "51",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        var cell = new DocxTableCell(
            string.Empty,
            [cellAnchor, DocxTests.CreateDocxLayoutParagraph("Cell trailing line", 10d, 12d)],
            null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [200d], [new DocxTableRow([cell], 30d)]);
        DocxParagraphElement[] noteElements = Enumerable.Range(0, 120)
            .Select(index => new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Shared note line " + index.ToString(CultureInfo.InvariantCulture) + " with trailing words", 10d, 12d)))
            .ToArray();
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "51",
            noteElements,
            [],
            [], null);
        var bodyElements = new List<DocxParagraphElement> { };
        var paragraphs = new List<DocxParagraph> { cellAnchor };
        for (int index = 0; index < 60; index++)
        {
            DocxParagraph filler = DocxTests.CreateDocxLayoutParagraph("Filler body line " + index.ToString(CultureInfo.InvariantCulture) + " keeps body flowing", 10d, 12d);
            bodyElements.Add(new DocxParagraphElement(filler));
            paragraphs.Add(filler);
        }

        var document = new DocxDocument(
            612d,
            792d,
            72d,
            72d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            new DocxBodyElement[] { new DocxTableElement(table) }.Concat(bodyElements).ToArray(),
            paragraphs.ToArray(),
            [])
        {
            RelatedStories = [footnoteStory]
        };
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        bool sharedPageExists = false;
        foreach (DocxLayoutPage page in layout.Pages)
        {
            bool hasFiller = page.Items.OfType<DocxTextLineLayout>().Any(line => line.Text.StartsWith("Filler body line", StringComparison.Ordinal));
            bool hasTable = page.Items.OfType<DocxTableRowLayout>().Any() || page.Items.OfType<DocxTextLineLayout>().Any(line => line.Text.StartsWith("Cell ", StringComparison.Ordinal));
            bool hasNote = page.PlacedRelatedStories
                .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                    (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
                .SelectMany(story => story.TextLines)
                .Any();
            if (hasFiller && hasNote && !hasTable)
            {
                sharedPageExists = true;
            }
        }

        TestAssert.True(sharedPageExists, "Table-anchored notes must share body pages instead of dedicating continuation pages.");

        string[] placedTexts = layout.Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Select(line => line.Text)
            .ToArray();
        TestAssert.Equal(120, placedTexts.Length);
        TestAssert.Equal(120, placedTexts.Distinct().Count());
    }

    public static void DocxKeepChainBlockIndexesFollowKeepNext()
    {
        DocxParagraph Keep(string text)
        {
            return DocxTests.CreateDocxLayoutParagraph(text, 10d, 12d, keepRules: new DocxParagraphKeepRules(true, null, null, null, null, null));
        }

        DocxParagraph plain = DocxTests.CreateDocxLayoutParagraph("Plain", 10d, 12d);
        var table = new DocxTable(null, [60d], [new DocxTableRow([new DocxTableCell("Cell", [plain], null, null, null, null, [], DocxTableCellMargins.Empty)], 30d)]);
        IReadOnlyList<DocxBodyElement> elements = [new DocxParagraphElement(Keep("Keep zero")), new DocxParagraphElement(Keep("Keep one")), new DocxParagraphElement(plain), new DocxTableElement(table), new DocxParagraphElement(Keep("Keep four"))];

        TestAssert.True(new List<int> { 0, 1, 2 }.SequenceEqual(DocxLayoutEngine.KeepChainBlockIndexes(elements, 0)), "Keep chain from 0 must span both keep paras plus the plain follower.");
        TestAssert.True(new List<int> { 1, 2 }.SequenceEqual(DocxLayoutEngine.KeepChainBlockIndexes(elements, 1)), "Keep chain from 1 must span its keep para plus the plain follower.");
        TestAssert.True(new List<int> { 2 }.SequenceEqual(DocxLayoutEngine.KeepChainBlockIndexes(elements, 2)), "Plain paragraphs carry no chain.");
        TestAssert.True(new List<int> { 3 }.SequenceEqual(DocxLayoutEngine.KeepChainBlockIndexes(elements, 3)), "Table elements carry no chain.");
        TestAssert.True(new List<int> { 4 }.SequenceEqual(DocxLayoutEngine.KeepChainBlockIndexes(elements, 4)), "Trailing keep with no follower carries no chain.");
    }

    public static void DocxSharedContinuedFootnotePagesCarryContinuationRules()
    {
        DocxDocument document = CreateInterleavingDocument(noteParagraphCount: 120, fillerParagraphCount: 60, footnoteId: "44", includeSeparators: true);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        bool headPageSeen = false;
        foreach (DocxLayoutPage page in layout.Pages)
        {
            bool hasNormalLines = page.PlacedRelatedStories
                .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                    (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
                .SelectMany(story => story.TextLines)
                .Any();
            if (!hasNormalLines)
            {
                continue;
            }

            if (!headPageSeen)
            {
                headPageSeen = true;
                continue;
            }
        TestAssert.True(
                page.PlacedRelatedStories.Any(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                    story.StoryLayout.Story.Type == DocxRelatedStoryType.ContinuationSeparator),
                "Continued footnote pages must open with a continuation rule.");

        }

    }

    private static DocxDocument CreateMidBodyInterleavingDocument(int fillersBefore, int noteParagraphCount, string footnoteId, int fillersAfter)
    {
        DocxParagraph anchor = DocxTests.CreateDocxLayoutParagraph("Body note marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    footnoteId,
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        DocxParagraphElement[] noteElements = Enumerable.Range(0, noteParagraphCount)
            .Select(index => new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Shared note line " + index.ToString(CultureInfo.InvariantCulture) + " with trailing words", 10d, 12d)))
            .ToArray();
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            footnoteId,
            noteElements,
            [],
            [], null);
        var bodyElements = new List<DocxParagraphElement>();
        var paragraphs = new List<DocxParagraph>();
        for (int index = 0; index < fillersBefore; index++)
        {
            DocxParagraph filler = DocxTests.CreateDocxLayoutParagraph("Filler body line " + index.ToString(CultureInfo.InvariantCulture) + " keeps body flowing", 10d, 12d);
            bodyElements.Add(new DocxParagraphElement(filler));
            paragraphs.Add(filler);
        }

        bodyElements.Add(new DocxParagraphElement(anchor));
        paragraphs.Add(anchor);
        for (int index = 0; index < fillersAfter; index++)
        {
            DocxParagraph filler = DocxTests.CreateDocxLayoutParagraph("Filler body line " + (fillersBefore + index).ToString(CultureInfo.InvariantCulture) + " keeps body flowing", 10d, 12d);
            bodyElements.Add(new DocxParagraphElement(filler));
            paragraphs.Add(filler);
        }

        return new DocxDocument(
            612d,
            792d,
            72d,
            72d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            bodyElements.ToArray(),
            paragraphs.ToArray(),
            [])
        {
            RelatedStories = [footnoteStory]
        };
    }

    private static DocxDocument CreateInterleavingDocument(int noteParagraphCount, int fillerParagraphCount, string footnoteId, bool includeSeparators = false)
    {
        DocxParagraph anchor = DocxTests.CreateDocxLayoutParagraph("Body note marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    footnoteId,
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        DocxParagraphElement[] noteElements = Enumerable.Range(0, noteParagraphCount)
            .Select(index => new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Shared note line " + index.ToString(CultureInfo.InvariantCulture) + " with trailing words", 10d, 12d)))
            .ToArray();
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            footnoteId,
            noteElements,
            [],
            [], null);
        var bodyElements = new List<DocxParagraphElement> { new DocxParagraphElement(anchor) };
        DocxRelatedStory? separatorStory = null;
        DocxRelatedStory? continuationStory = null;
        if (includeSeparators)
        {
            separatorStory = new DocxRelatedStory(DocxRelatedStoryKind.Footnote, "/word/footnotes.xml", "0", [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Note separator mark", 10d, 12d))], [], [], DocxRelatedStoryType.Separator);
            continuationStory = new DocxRelatedStory(DocxRelatedStoryKind.Footnote, "/word/footnotes.xml", "1", [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Note continuation mark", 10d, 12d))], [], [], DocxRelatedStoryType.ContinuationSeparator);
        }

        var paragraphs = new List<DocxParagraph> { anchor };
        for (int index = 0; index < fillerParagraphCount; index++)
        {
            DocxParagraph filler = DocxTests.CreateDocxLayoutParagraph("Filler body line " + index.ToString(CultureInfo.InvariantCulture) + " keeps body flowing", 10d, 12d);
            bodyElements.Add(new DocxParagraphElement(filler));
            paragraphs.Add(filler);
        }

        return new DocxDocument(
            612d,
            792d,
            72d,
            72d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            bodyElements.ToArray(),
            paragraphs.ToArray(),
            [])
        {
            RelatedStories = separatorStory is null || continuationStory is null ? [footnoteStory] : [footnoteStory, separatorStory, continuationStory]
        };
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

    public static void DocxFootnoteSeparatorIgnoresDirectParagraphSpacing()
    {
        // Shared story-layout path with endnote separators (Office evidence is
        // endnote-only; no diverging footnote evidence): direct separator spacing
        // must not move footnote content, and the separator rule must not ride the
        // direct spacing either (the footnote content hangs off the stories top
        // while the rule sits in the separator story, so the rule is the moving
        // part here).
        double topAfter0 = LayoutFootnoteBodyTopWithSeparatorSpacing("<w:pPr><w:spacing w:after=\"0\"/></w:pPr>");
        double topAfter24 = LayoutFootnoteBodyTopWithSeparatorSpacing("<w:pPr><w:spacing w:after=\"480\"/></w:pPr>");
        double topDefault = LayoutFootnoteBodyTopWithSeparatorSpacing(string.Empty);
        TestAssert.True(Math.Abs(topAfter0 - topAfter24) < 0.000001d, "Footnote separator after-spacing must not move the note block; observed shift=" + Math.Abs(topAfter0 - topAfter24).ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(topAfter0 - topDefault) < 0.000001d, "Footnote separator after-spacing must match latent defaults; observed shift=" + Math.Abs(topAfter0 - topDefault).ToString(CultureInfo.InvariantCulture) + ".");
        double ruleAfter0 = LayoutFootnoteSeparatorRuleWithSpacing("<w:pPr><w:spacing w:after=\"0\"/></w:pPr>");
        double ruleAfter24 = LayoutFootnoteSeparatorRuleWithSpacing("<w:pPr><w:spacing w:after=\"480\"/></w:pPr>");
        double ruleDefault = LayoutFootnoteSeparatorRuleWithSpacing(string.Empty);
        TestAssert.True(Math.Abs(ruleAfter0 - ruleAfter24) < 0.000001d, "Footnote separator rule must not ride direct after-spacing; observed shift=" + Math.Abs(ruleAfter0 - ruleAfter24).ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(ruleAfter0 - ruleDefault) < 0.000001d, "Footnote separator rule must match latent defaults; observed shift=" + Math.Abs(ruleAfter0 - ruleDefault).ToString(CultureInfo.InvariantCulture) + ".");
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
    private sealed class HheaAscenderTextMeasurer(IDocxTextMeasurer inner, double hheaAscenderEm) : IDocxTextMeasurer, IDocxLineMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => inner is IDocxLineMetricsProvider lineMetrics ? lineMetrics.MeasureSingleLineHeight(run, fontSize) : fontSize;

        public double MeasureHheaAscender(DocxTextRun? run, double fontSize) => hheaAscenderEm * fontSize;
    }
    private static double LayoutEndnoteSeparatorHeightWithHheaAscender(int docDefaultsHalfPoints, double hheaAscenderEm)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/><Relationship Id="rIdS1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>""",
            ["word/styles.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="SZ"/><w:szCs w:val="SZ"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>""".Replace("SZ", docDefaultsHalfPoints.ToString(CultureInfo.InvariantCulture)),
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with endnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body one</w:t></w:r></w:p><w:p><w:r><w:t>Note body two</w:t></w:r></w:p></w:endnote></w:endnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new HheaAscenderTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), hheaAscenderEm), CancellationToken.None);
        return layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator).Height;
    }
    private static double LayoutEndnoteSeparatorHeightWithSingleLineEm(int docDefaultsHalfPoints, double singleLineEm)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/><Relationship Id="rIdS1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>""",
            ["word/styles.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="SZ"/><w:szCs w:val="SZ"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>""".Replace("SZ", docDefaultsHalfPoints.ToString(CultureInfo.InvariantCulture)),
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with endnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body one</w:t></w:r></w:p><w:p><w:r><w:t>Note body two</w:t></w:r></w:p></w:endnote></w:endnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new SingleLineEmTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), singleLineEm), CancellationToken.None);
        return layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator).Height;
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

    private sealed class SingleLineEmTextMeasurer(IDocxTextMeasurer inner, double singleLineEm) : IDocxTextMeasurer
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public bool TryGetSingleLineEm(DocxTextRun? run, out double value)
        {
            value = singleLineEm;
            return true;
        }
    }

    private static double LayoutEndnoteSeparatorContentGapWithMarkSize(int docDefaultsHalfPoints, double singleLineEm)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/><Relationship Id="rIdS1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>""",
            ["word/styles.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="SZ"/><w:szCs w:val="SZ"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>""".Replace("SZ", docDefaultsHalfPoints.ToString(CultureInfo.InvariantCulture)),
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with endnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body one</w:t></w:r></w:p><w:p><w:r><w:t>Note body two</w:t></w:r></w:p></w:endnote></w:endnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new SingleLineEmTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), singleLineEm), CancellationToken.None);
        DocxPlacedRelatedStoryLayout separator = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
        DocxPlacedRelatedStoryLayout content = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal));
        return (separator.TopY - separator.Height) - content.TopY;
    }
    private static double LayoutEndnoteSeparatorHeightWithMarkSize(int docDefaultsHalfPoints)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/><Relationship Id="rIdS1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>""",
            ["word/styles.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="SZ"/><w:szCs w:val="SZ"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>""".Replace("SZ", docDefaultsHalfPoints.ToString(CultureInfo.InvariantCulture)),
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with endnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body one</w:t></w:r></w:p><w:p><w:r><w:t>Note body two</w:t></w:r></w:p></w:endnote></w:endnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        return layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator).Height;
    }

    public static void DocxEndnoteSeparatorHeightAddsSmallAscSupplement()
    {
        double control = LayoutEndnoteSeparatorHeightWithWindowsAscender(16, 0.9521d);
        double small = LayoutEndnoteSeparatorHeightWithWindowsAscender(16, 0.89d);
        TestAssert.True(Math.Abs((small - control) - 0.50d) < 0.01d, "Document-end separator height must add the small ascender supplement; observed shift=" + (small - control).ToString(CultureInfo.InvariantCulture) + ".");
    }

    private sealed class WindowsAscenderTextMeasurer(IDocxTextMeasurer inner, double windowsAscenderEm) : IDocxTextMeasurer, IDocxStaticTextMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => windowsAscenderEm * fontSize;

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => inner is IDocxStaticTextMetricsProvider staticMetrics ? staticMetrics.MeasureWindowsDescender(run, fontSize) : fontSize * 0.2d;
    }

    private static double LayoutEndnoteSeparatorHeightWithWindowsAscender(int docDefaultsHalfPoints, double windowsAscenderEm)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/><Relationship Id="rIdS1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>""",
            ["word/styles.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="SZ"/><w:szCs w:val="SZ"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>""".Replace("SZ", docDefaultsHalfPoints.ToString(CultureInfo.InvariantCulture)),
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with endnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body one</w:t></w:r></w:p><w:p><w:r><w:t>Note body two</w:t></w:r></w:p></w:endnote></w:endnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new WindowsAscenderTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), windowsAscenderEm), CancellationToken.None);
        return layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator).Height;
    }

    public static void DocxFootnoteLineHeightAddsDescDeficitExtra()
    {
        double tallFirst = LayoutFootnotePitchWithDescFaces("Tahoma", "Calibri");
        double tallSecond = LayoutFootnotePitchWithDescFaces("Calibri", "Tahoma");
        TestAssert.True(Math.Abs(tallFirst - 20.40d) < 0.01d, "Footnote line height must add the descender deficit extra; tallFirst=" + tallFirst.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(tallSecond - tallFirst) < 0.000001d, "Footnote descender extra must ignore run order; tallSecond=" + tallSecond.ToString(CultureInfo.InvariantCulture) + " tallFirst=" + tallFirst.ToString(CultureInfo.InvariantCulture) + ".");
    }

    public static void DocxFootnoteLineHeightIgnoresWhitespaceOnlyRuns()
    {
        double pitch = LayoutFootnotePitchWithSpaceOnlyRun();
        TestAssert.True(Math.Abs(pitch - 19.58d) < 0.01d, "Footnote line height must ignore whitespace only runs; pitch=" + pitch.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private static double LayoutFootnotePitchWithSpaceOnlyRun()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="Tahoma" w:hAnsi="Tahoma"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body one start tail one</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="Tahoma" w:hAnsi="Tahoma"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body two start tail two</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DescDeficitTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxTextLineLayout[] lines = layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Where(line => line.Text.StartsWith("Note body", StringComparison.Ordinal))
            .Take(2)
            .ToArray();
        TestAssert.Equal(2, lines.Length);
        return lines[0].BaselineY - lines[1].BaselineY;
    }

    public static void DocxLineHeightPrefersWindowsBoxOverHhea()
    {
        // RV06 line-box probes (Word 16.0): Office auto line boxes take max(Windows box, hhea box).
        (double footnotePitch, double bodyPitch) = LayoutWinBoxPitches();
        TestAssert.True(Math.Abs(footnotePitch - 23.98d) < 0.02d, "Footnote line height must prefer the Windows box.");
        TestAssert.True(Math.Abs(bodyPitch - 23.98d) < 0.02d, "Body line height must prefer the Windows box.");
    }

    private static (double FootnotePitch, double BodyPitch) LayoutWinBoxPitches()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="WinBox" w:hAnsi="WinBox"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Body winbox line one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="WinBox" w:hAnsi="WinBox"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Body winbox line two</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="WinBox" w:hAnsi="WinBox"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note winbox line one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="WinBox" w:hAnsi="WinBox"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note winbox line two</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new WinBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxTextLineLayout[] footnoteLines = layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Where(line => line.Text.StartsWith("Note winbox", StringComparison.Ordinal))
            .Take(2)
            .ToArray();
        TestAssert.Equal(2, footnoteLines.Length);
        DocxTextLineLayout[] bodyLines = layout.Pages[0].Items.OfType<DocxTextLineLayout>()
            .Where(line => line.Text.StartsWith("Body winbox", StringComparison.Ordinal))
            .Take(2)
            .ToArray();
        TestAssert.Equal(2, bodyLines.Length);
        return (footnoteLines[0].BaselineY - footnoteLines[1].BaselineY, bodyLines[0].BaselineY - bodyLines[1].BaselineY);
    }

    private sealed class WinBoxTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => 10d;

        public double MeasureHheaLineHeight(DocxTextRun? run, double fontSize) => 12d;

        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => fontSize * 0.9d;

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => fontSize * 0.25d;
    }

    public static void DocxLineHeightPrefersTypographicBoxOverHhea()
    {
        // RV06 line-box probes (Word 16.0): Abadi runs box the unfloored typo box alone (typoFull 1.0698 against hhea 1.3047 with single floored at 1.15).
        (double footnotePitch, double bodyPitch) = LayoutTypoBoxPitches();
        TestAssert.True(Math.Abs(footnotePitch - 22.87d) < 0.02d, "Footnote line height must prefer the typographic box.");
        TestAssert.True(Math.Abs(bodyPitch - 22.87d) < 0.02d, "Body line height must prefer the typographic box.");
    }

    private static (double FootnotePitch, double BodyPitch) LayoutTypoBoxPitches()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="TypoBox" w:hAnsi="TypoBox"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Body typobox line one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="TypoBox" w:hAnsi="TypoBox"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Body typobox line two</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="TypoBox" w:hAnsi="TypoBox"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note typobox line one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="TypoBox" w:hAnsi="TypoBox"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note typobox line two</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new TypoBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxTextLineLayout[] footnoteLines = layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Where(line => line.Text.StartsWith("Note typobox", StringComparison.Ordinal))
            .Take(2)
            .ToArray();
        TestAssert.Equal(2, footnoteLines.Length);
        DocxTextLineLayout[] bodyLines = layout.Pages[0].Items.OfType<DocxTextLineLayout>()
            .Where(line => line.Text.StartsWith("Body typobox", StringComparison.Ordinal))
            .Take(2)
            .ToArray();
        TestAssert.Equal(2, bodyLines.Length);
        return (footnoteLines[0].BaselineY - footnoteLines[1].BaselineY, bodyLines[0].BaselineY - bodyLines[1].BaselineY);
    }

    private sealed class TypoBoxTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider, IDocxTypographicMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => 13.8d;

        public double MeasureHheaLineHeight(DocxTextRun? run, double fontSize) => 15.66d;

        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => 10.96d;

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => 2.89d;

        public bool UseTypographicMetrics(DocxTextRun? run) => true;

        public double MeasureTypographicLineHeight(DocxTextRun? run, double fontSize) => 12.84d;
    }

    private sealed class DescDeficitTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => 10d;

        public double MeasureHheaLineHeight(DocxTextRun? run, double fontSize) => 10d;

        // Windows extents stay inside the 10pt hhea box so the deficit mechanism stays isolated under the line-box maximum (Tahoma ascender still leads for order-invariance).
        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "Tahoma", StringComparison.Ordinal) ? fontSize * 0.6d : fontSize * 0.5d;

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "Tahoma", StringComparison.Ordinal) ? fontSize * 0.2d : fontSize * 0.27d;
    }

    private static double LayoutFootnotePitchWithDescFaces(string firstFamily, string secondFamily)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body one start tail one</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> mixed tail one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body two start tail two</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> mixed tail two</w:t></w:r></w:p></w:footnote></w:footnotes>""".Replace("F1", firstFamily).Replace("F2", secondFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DescDeficitTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxTextLineLayout[] lines = layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Where(line => line.Text.StartsWith("Note body", StringComparison.Ordinal))
            .Take(2)
            .ToArray();
        TestAssert.Equal(2, lines.Length);
        return lines[0].BaselineY - lines[1].BaselineY;
    }

    public static void DocxFootnoteFirstInsetFollowsDeeperDescender()
    {
        // RV06 descender probes (Word 16.0, Magneto/Calibri/Informal resolved sheets):
        // Office bottom bearing follows the deeper descender, so the inset is the hhea
        // box minus the maximum descender while equal descenders resolve identically
        // to the ascender-plus-gap form on both sides.
        double same = LayoutFootnoteFirstBaselineWithDescFaces("DescSame", "DescSame");
        double most = LayoutFootnoteFirstBaselineWithDescFaces("DescMost", "DescMost");
        TestAssert.True(Math.Abs((most - same) - 1.20d) < 0.02d, "Deeper-descender first baseline must sit above the equal-descender one by the inset difference; most=" + most.ToString(CultureInfo.InvariantCulture) + " same=" + same.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private sealed class DescMostTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider, IDocxHheaLineGapProvider, IDocxHheaDescenderProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => 12d;

        public double MeasureHheaLineHeight(DocxTextRun? run, double fontSize) => 14d;

        public double MeasureHheaAscender(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "DescMost", StringComparison.Ordinal) ? 12d : 11.6d;

        public double MeasureHheaLineGap(DocxTextRun? run, double fontSize) => 0d;

        public double MeasureHheaDescender(DocxTextRun? run, double fontSize) => 2.4d;

        // Windows extents stay inside the hhea box so the inset mechanism stays isolated under the line-box maximum.
        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => fontSize * 0.8d;

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "DescMost", StringComparison.Ordinal) ? 3.6d : 2.4d;
    }

    private static double LayoutFootnoteFirstBaselineWithDescFaces(string firstFamily, string secondFamily)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body one start tail one</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> mixed tail one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body two start tail two</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> mixed tail two</w:t></w:r></w:p></w:footnote></w:footnotes>""".Replace("F1", firstFamily).Replace("F2", secondFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DescMostTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxTextLineLayout[] lines = layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Where(line => line.Text.StartsWith("Note body", StringComparison.Ordinal))
            .Take(2)
            .ToArray();
        TestAssert.Equal(2, lines.Length);
        return lines[0].BaselineY;
    }

    public static void DocxFootnoteFirstInsetFollowsWindowsDescender()
    {
        // RV06 windows-descender probes (Word 16.0, Informal Roman 10/12/14pt size
        // matrix plus Consolas): Office bottom bearing follows the Windows descender
        // even when the hhea descender runs deeper, so two synthetic faces sharing
        // the hhea box and Windows descender keep the same first baseline while
        // differing only in hhea descender depth.
        double deep = LayoutFootnoteFirstBaselineWithWinDescFaces("HheaDeep", "HheaDeep");
        double flat = LayoutFootnoteFirstBaselineWithWinDescFaces("HheaFlat", "HheaFlat");
        TestAssert.True(Math.Abs(deep - flat) < 0.02d, "Windows-descender first baselines must ignore hhea descender depth; deep=" + deep.ToString(CultureInfo.InvariantCulture) + " flat=" + flat.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private sealed class WinDescTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider, IDocxHheaLineGapProvider, IDocxHheaDescenderProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => 12d;

        public double MeasureHheaLineHeight(DocxTextRun? run, double fontSize) => 14d;

        public double MeasureHheaAscender(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "HheaDeep", StringComparison.Ordinal) ? 12d : 11.6d;

        public double MeasureHheaLineGap(DocxTextRun? run, double fontSize) => 0d;

        public double MeasureHheaDescender(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "HheaDeep", StringComparison.Ordinal) ? 3.0d : 2.4d;

        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => fontSize * 0.9d;

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => 2.4d;
    }

    private static double LayoutFootnoteFirstBaselineWithWinDescFaces(string firstFamily, string secondFamily)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body one start tail one</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> mixed tail one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body two start tail two</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> mixed tail two</w:t></w:r></w:p></w:footnote></w:footnotes>""".Replace("F1", firstFamily).Replace("F2", secondFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new WinDescTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxTextLineLayout[] lines = layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Where(line => line.Text.StartsWith("Note body", StringComparison.Ordinal))
            .Take(2)
            .ToArray();
        TestAssert.Equal(2, lines.Length);
        return lines[0].BaselineY;
    }

    public static void DocxTextBoxFirstInsetFollowsWindowsDescender()
    {
        // RV06 floatinset probes (Word 16.0, Tahoma slash Calibri slash Informal floating
        // textboxes): Office textbox first insets follow the hhea-box-minus-Windows-descender
        // tier, so a uniform synthetic face keeps its tier inset in textbox stories
        // instead of the legacy floor.
        double baseline = LayoutTextBoxFirstBaselineWithDescFaces("BoxTbx");
        TestAssert.True(Math.Abs(baseline - -11.6d) < 0.02d, "Textbox first baseline must sit at the tier inset; baseline=" + baseline.ToString(CultureInfo.InvariantCulture) + ".");
    }

    public static void DocxTextBoxFirstInsetFollowsMaxTier()
    {
        // RV06 floatmixsz probes (Word 16.0, Tahoma 10/14pt mixed with Calibri 12pt, both orders): Office
        // mixed-textbox insets take the max tier across runs (Tah10 reads Cal-tier, Tah14 reads Tah-tier),
        // unlike footnote mixed notes which keep the legacy floor so take capacities hold. The crossed
        // synthetic (high ascender on the low tier and vice versa) separates tier selection from
        // ascender selection. Pre-fix both orders read the ascender winner.
        double hiFirst = LayoutMixedTextBoxFirstBaseline("MixHi", "MixLo");
        double loFirst = LayoutMixedTextBoxFirstBaseline("MixLo", "MixHi");
        TestAssert.True(Math.Abs(hiFirst - -11.6d) < 0.02d, "Max-first mixed baselines must sit at the max tier inset; hiFirst=" + hiFirst.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(loFirst - -11.6d) < 0.02d, "Min-first mixed baselines must sit at the max tier inset; loFirst=" + loFirst.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private sealed class MixMaxTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider, IDocxHheaLineGapProvider, IDocxHheaDescenderProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => 12d;

        public double MeasureHheaLineHeight(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "MixHi", StringComparison.Ordinal) ? 11d : 14d;

        public double MeasureHheaAscender(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "MixHi", StringComparison.Ordinal) ? 12d : 9d;

        public double MeasureHheaLineGap(DocxTextRun? run, double fontSize) => 0d;

        public double MeasureHheaDescender(DocxTextRun? run, double fontSize) => 2.4d;

        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => fontSize * 0.9d;

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => 2.4d;
    }

    private static double LayoutMixedTextBoxFirstBaseline(string firstFamily, string secondFamily)
    {
        var runA = new DocxTextRun("Box note one ", 12d, null, false, false, false, null, firstFamily);
        var runB = new DocxTextRun("mixed tail", 12d, null, false, false, false, null, secondFamily);
        var paragraph = new DocxParagraph(
            [runA, runB],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var anchor = DocxTests.CreateDocxLayoutParagraph("Anchor", 10d, 12d);
        var drawing = DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(paragraph)]);
        var document = new DocxDocument(
            612d, 792d, 72d, 72d, 72d, 72d,
            DocxPageSettings.Empty,
            [drawing],
            [], [], [new DocxParagraphElement(anchor)], [], []);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new MixMaxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxFloatingDrawingLayout placed = layout.FloatingDrawings.Single();
        var textBoxLayout = TestAssert.NotNull(placed.TextBoxLayout, "Floating drawing must carry its textbox story layout.");
        DocxTextLineLayout first = textBoxLayout.TextLines.First(line => line.Text.StartsWith("Box note", StringComparison.Ordinal));
        return first.BaselineY;
    }

    private static double LayoutTextBoxFirstBaselineWithDescFaces(string family)
    {
        var run = new DocxTextRun("Box note one", 12d, null, false, false, false, null, family);
        var paragraph = new DocxParagraph(
            [run],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var anchor = DocxTests.CreateDocxLayoutParagraph("Anchor", 10d, 12d);
        var drawing = DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(paragraph)]);
        var document = new DocxDocument(
            612d, 792d, 72d, 72d, 72d, 72d,
            DocxPageSettings.Empty,
            [drawing],
            [], [], [new DocxParagraphElement(anchor)], [], []);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DescMostTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxFloatingDrawingLayout placed = layout.FloatingDrawings.Single();
        var textBoxLayout = TestAssert.NotNull(placed.TextBoxLayout, "Floating drawing must carry its textbox story layout.");
        DocxTextLineLayout first = textBoxLayout.TextLines.First(line => line.Text.StartsWith("Box note", StringComparison.Ordinal));
        return first.BaselineY;
    }

    public static void DocxFootnoteFirstInsetKeepsLegacyAboveMarkSingleEm()
    {
        // RV06 gap-slice guard (Palatino/Algerian first baselines): uniform notes whose
        // content single-em exceeds the mark reference stay on the legacy floor so
        // gap-overflowing notes keep validated placement while same-geometry Gate pairs
        // differ only by ascender.
        double big = LayoutFootnoteFirstBaselineWithFloorFaces("GateBig", "GateBig");
        double small = LayoutFootnoteFirstBaselineWithFloorFaces("GateSmall", "GateSmall");
        TestAssert.True(Math.Abs((small - big) - 0.72d) < 0.02d, "Gate pair must keep the legacy floor above mark single-em; small=" + small.ToString(CultureInfo.InvariantCulture) + " big=" + big.ToString(CultureInfo.InvariantCulture) + ".");
    }

    public static void DocxFootnoteFirstInsetFollowsHheaAscenderPlusGap()
    {
        // RV06 box-law instruments (eighteen-family shape probes, Word 16.0): Office
        // first insets follow min-over-runs hhea ascender plus hhea gap with no 0.94em
        // floor, so a synthetic small face keeps its summed inset while a big-ascender
        // face keeps its own sum on both sides.
        double big = LayoutFootnoteFirstBaselineWithFloorFaces("BigTop", "BigTop");
        double small = LayoutFootnoteFirstBaselineWithFloorFaces("LowFloor", "LowFloor");
        TestAssert.True(Math.Abs((small - big) - 2.04d) < 0.02d, "Small-first minus big-first baselines must span the summed-inset difference; small=" + small.ToString(CultureInfo.InvariantCulture) + " big=" + big.ToString(CultureInfo.InvariantCulture) + ".");

    }

    private sealed class FloorInsetTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider, IDocxHheaLineGapProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "GateBig", StringComparison.Ordinal) || string.Equals(run?.FontFamily, "GateSmall", StringComparison.Ordinal) ? 16d : 10d;

        public double MeasureHheaLineHeight(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "GateBig", StringComparison.Ordinal) || string.Equals(run?.FontFamily, "GateSmall", StringComparison.Ordinal) ? 16d : 10d;

        public double MeasureHheaAscender(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "BigTop", StringComparison.Ordinal) || string.Equals(run?.FontFamily, "GateBig", StringComparison.Ordinal) ? fontSize : fontSize * 0.8d;

        public double MeasureHheaLineGap(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "LowFloor", StringComparison.Ordinal) ? fontSize * 0.03d : 0d;

        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => fontSize * 0.9d;

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => fontSize * 0.2d;
    }

    private static double LayoutFootnoteFirstBaselineWithFloorFaces(string firstFamily, string secondFamily)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body one start tail one</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> mixed tail one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body two start tail two</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> mixed tail two</w:t></w:r></w:p></w:footnote></w:footnotes>""".Replace("F1", firstFamily).Replace("F2", secondFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new FloorInsetTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxTextLineLayout[] lines = layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Where(line => line.Text.StartsWith("Note body", StringComparison.Ordinal))
            .Take(2)
            .ToArray();
        TestAssert.Equal(2, lines.Length);
        return lines[0].BaselineY;
    }

    public static void DocxFootnoteLineHeightFollowsHheaWithoutSingleFloor()
    {
        // RV06 box-law instruments (Corsiva/YiBaiti/Haettenschweiler/Arial/Ebrima
        // same-batch probes, Word 16.0): Office uniform pitches follow the per-family
        // hhea box with no single-line floor (Corsiva 13.50 against 13.46, YiBaiti
        // 12.36 against 12.32, Haett 12.88 against 12.80, Arial 13.81 against 13.80,
        // Ebrima 16.30 against 16.29), so a synthetic small-hhea face keeps its hhea
        // pitch while a big-hhea face keeps the hhea maximum on both sides.
        double big = LayoutFootnotePitchWithFloorFaces("BigBox", "BigBox");
        double small = LayoutFootnotePitchWithFloorFaces("SmallBox", "SmallBox");
        TestAssert.True(Math.Abs(big - 24.22d) < 0.01d, "Big-hhea pitch must keep the hhea maximum; big=" + big.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(small - 19.58d) < 0.01d, "Small-hhea pitch must drop the single-line floor; small=" + small.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private sealed class FloorBoxTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "SmallBox", StringComparison.Ordinal) || string.Equals(run?.FontFamily, "BigBox", StringComparison.Ordinal) ? 14d : 10d;

        public double MeasureHheaLineHeight(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "BigBox", StringComparison.Ordinal) ? 14d : 10d;

        // Windows extents stay inside the hhea box so the floor mechanism stays isolated under the line-box maximum (Tahoma ascender still leads).
        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "Tahoma", StringComparison.Ordinal) ? fontSize * 0.6d : fontSize * 0.5d;

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "Tahoma", StringComparison.Ordinal) ? fontSize * 0.2d : fontSize * 0.27d;
    }

    private static double LayoutFootnotePitchWithFloorFaces(string firstFamily, string secondFamily)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body one start tail one</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> mixed tail one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body two start tail two</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> mixed tail two</w:t></w:r></w:p></w:footnote></w:footnotes>""".Replace("F1", firstFamily).Replace("F2", secondFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new FloorBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxTextLineLayout[] lines = layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Where(line => line.Text.StartsWith("Note body", StringComparison.Ordinal))
            .Take(2)
            .ToArray();
        TestAssert.Equal(2, lines.Length);
        return lines[0].BaselineY - lines[1].BaselineY;
    }

    public static void DocxFootnoteLineHeightSelfGatesWithoutSmallRun()
    {
        // RV06 box-law instruments: two big-ascender runs with matching descenders
        // carry no descender deficit, so the line keeps the uniform pitch while the
        // same helper with a small-ascender second run grows by the deficit extra.
        double control = LayoutFootnotePitchWithSelfGateFaces("Tahoma", "Calibri");
        double bigPair = LayoutFootnotePitchWithSelfGateFaces("Tahoma", "BigTop");
        TestAssert.True(Math.Abs(control - 20.40d) < 0.01d, "Mixed control pitch must keep the descender deficit extra; control=" + control.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(bigPair - 19.58d) < 0.01d, "Big-ascender pair pitch must self-gate to the uniform value; bigPair=" + bigPair.ToString(CultureInfo.InvariantCulture) + ".");
    }

    public static void DocxFootnoteLineHeightSelfGatesWithDeepOwnDescender()
    {
        // RV06 box-law instruments (Palatino-like): the max-ascender run carrying the
        // deepest descender leaves a non-positive deficit, so the line keeps the
        // uniform pitch while the same helper with a shallow second run grows.
        double control = LayoutFootnotePitchWithSelfGateFaces("Tahoma", "Calibri");
        double deepOwn = LayoutFootnotePitchWithSelfGateFaces("DeepDsc", "Calibri");
        TestAssert.True(Math.Abs(control - 20.40d) < 0.01d, "Mixed control pitch must keep the descender deficit extra; control=" + control.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(deepOwn - 19.58d) < 0.01d, "Deep-descender max-ascender pitch must self-gate to the uniform value; deepOwn=" + deepOwn.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private sealed class SelfGateTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => 10d;

        public double MeasureHheaLineHeight(DocxTextRun? run, double fontSize) => 10d;

        // Windows extents stay inside the hhea box so the self-gate mechanism stays isolated under the line-box maximum (ascender order preserved).
        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => SelfGateAscenderEm(run) * fontSize * 0.45d;

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => SelfGateDescenderEm(run) * fontSize;

        private static double SelfGateAscenderEm(DocxTextRun? run)
        {
            if (string.Equals(run?.FontFamily, "DeepDsc", StringComparison.Ordinal))
            {
                return 1.05d;
            }

            if (string.Equals(run?.FontFamily, "Calibri", StringComparison.Ordinal))
            {
                return 0.9d;
            }

            return 1.0d;
        }

        private static double SelfGateDescenderEm(DocxTextRun? run)
        {
            if (string.Equals(run?.FontFamily, "DeepDsc", StringComparison.Ordinal))
            {
                return 0.35d;
            }

            if (string.Equals(run?.FontFamily, "Calibri", StringComparison.Ordinal))
            {
                return 0.27d;
            }

            return 0.2d;
        }
    }

    private static double LayoutFootnotePitchWithSelfGateFaces(string firstFamily, string secondFamily)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body one start tail one</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> mixed tail one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body two start tail two</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> mixed tail two</w:t></w:r></w:p></w:footnote></w:footnotes>""".Replace("F1", firstFamily).Replace("F2", secondFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new SelfGateTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxTextLineLayout[] lines = layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Where(line => line.Text.StartsWith("Note body", StringComparison.Ordinal))
            .Take(2)
            .ToArray();
        TestAssert.Equal(2, lines.Length);
        return lines[0].BaselineY - lines[1].BaselineY;
    }

    public static void DocxFootnoteContentGapFollowsFirstParagraph()
    {
        // RV06 mixed-gap probes (Word 16.0, Pal/Cal orders with r2f diff minus 1.20):
        // Office content gap follows the first paragraph while story-max would give
        // plus 0.4, so a Calibri-first mixed note keeps the Calibri-uniform gap while
        // Tahoma-first keeps the Tahoma-uniform gap.
        double tahUniform = LayoutFootnoteContentGapWithFamilies("Calibri", "Tahoma");
        double calUniform = LayoutFootnoteContentGapWithFamilies("Calibri", "Calibri");
        double tahFirst = LayoutMixedFootnoteContentGap("Tahoma", "Calibri");
        double calFirst = LayoutMixedFootnoteContentGap("Calibri", "Tahoma");
        TestAssert.True(Math.Abs(tahFirst - tahUniform) < 0.02d, "Tahoma-first mixed gap must keep the Tahoma-uniform gap; tahFirst=" + tahFirst.ToString(CultureInfo.InvariantCulture) + " tahUniform=" + tahUniform.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(calFirst - calUniform) < 0.02d, "Calibri-first mixed gap must keep the Calibri-uniform gap; calFirst=" + calFirst.ToString(CultureInfo.InvariantCulture) + " calUniform=" + calUniform.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private static double LayoutMixedFootnoteContentGap(string firstFamily, string secondFamily)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body two</w:t></w:r></w:p></w:footnote></w:footnotes>""".Replace("F1", firstFamily).Replace("F2", secondFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new FamilyGapTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxPlacedRelatedStoryLayout separator = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
        DocxPlacedRelatedStoryLayout content = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal));
        return (separator.TopY - separator.Height) - content.TopY;
    }

    public static void DocxFootnoteContentGapFollowsContentFont()
    {
        double control = LayoutFootnoteContentGapWithFamilies("Calibri", "Calibri");
        double mixed = LayoutFootnoteContentGapWithFamilies("Calibri", "Tahoma");
        TestAssert.True(Math.Abs(control - 3.37d) < 0.01d, "Uniform footnote gap must keep the shared value; control=" + control.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(mixed - 6.72d) < 0.01d, "Footnote gap must follow content font; mixed=" + mixed.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private sealed class FamilyGapTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public bool TryGetSingleLineEm(DocxTextRun? run, out double value)
        {
            value = string.Equals(run?.FontFamily, "Tahoma", StringComparison.Ordinal) ? 1.5d : 1.2207d;
            return true;
        }
    }

    private static double LayoutFootnoteContentGapWithFamilies(string markFamily, string contentFamily)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:rPr><w:rFonts w:ascii="MFAM" w:hAnsi="MFAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="CFAM" w:hAnsi="CFAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="CFAM" w:hAnsi="CFAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body two</w:t></w:r></w:p></w:footnote></w:footnotes>""".Replace("MFAM", markFamily).Replace("CFAM", contentFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new FamilyGapTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxPlacedRelatedStoryLayout separator = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
        DocxPlacedRelatedStoryLayout content = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal));
        return (separator.TopY - separator.Height) - content.TopY;
    }

    public static void DocxEndnoteContentGapFollowsContentFont()
    {
        double control = LayoutEndnoteContentGapWithFamilies("Calibri", "Calibri");
        double mixed = LayoutEndnoteContentGapWithFamilies("Calibri", "Tahoma");
        TestAssert.True(Math.Abs(control - 3.37d) < 0.01d, "Uniform endnote gap must keep the shared value; control=" + control.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(mixed - 6.72d) < 0.01d, "Endnote gap must follow content font; mixed=" + mixed.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private static double LayoutEndnoteContentGapWithFamilies(string markFamily, string contentFamily)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with endnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:rPr><w:rFonts w:ascii="MFAM" w:hAnsi="MFAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="CFAM" w:hAnsi="CFAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="CFAM" w:hAnsi="CFAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body two</w:t></w:r></w:p></w:endnote></w:endnotes>""".Replace("MFAM", markFamily).Replace("CFAM", contentFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new FamilyGapTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxPlacedRelatedStoryLayout separator = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
        DocxPlacedRelatedStoryLayout content = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal));
        return (separator.TopY - separator.Height) - content.TopY;
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

    private static double LayoutEndnoteContentGapWithWindowsAscender(string markFamily, string contentFamily)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with endnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:rPr><w:rFonts w:ascii="MFAM" w:hAnsi="MFAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="CFAM" w:hAnsi="CFAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="CFAM" w:hAnsi="CFAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body two</w:t></w:r></w:p></w:endnote></w:endnotes>""".Replace("MFAM", markFamily).Replace("CFAM", contentFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new FamilyWascTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxPlacedRelatedStoryLayout separator = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
        DocxPlacedRelatedStoryLayout content = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal));
        return (separator.TopY - separator.Height) - content.TopY;
    }

    public static void DocxFootnoteTakeExcludesTrailingAfterSpacing()
    {
        // RV06 after-spacing sweep (Word 16.0, seven after values): take needs exclude
        // the take last-line trailing after-spacing, so the head page takes one more
        // mixed line than the full-box fit allows. Synthetic desc-deficit metrics stand
        // in for the Calibri plus Tahoma mix with explicit 16pt note after-spacing, where
        // the legacy full-box fit takes 21 against 22 with the trailing-after exclusion.
        int headTake = LayoutFootnoteTakeWithAfterSpacingAndAfter(0, 320);
        TestAssert.Equal(22, headTake);
    }

    public static void DocxFootnoteContinuationTakeChargesRideTolerance()
    {
        // RV06 take-battery probes (Word 16.0, cal9-a8 laps 0.72 with cal-pal 0.36 while
        // a4 refuses the next line at lap 4.54): the take-side charge is the emitted rule
        // ride net of a 2.1pt overflow tolerance with full-box needs, rather than the full
        // rule block with trailing-after exclusion. At 24pt after-spacing the full-box
        // needs dominate, so the first continuation page takes 17 lines against 18 under
        // the legacy exclusion-plus-full-block model; the Office-anchored take-more
        // direction is pinned by the cal9-a8 29/31 and cal-pal 23/24/13 probe takes.
        // Same synthetic note shape as the overhead test at 24pt after-spacing.
        int continuationTake = LayoutFootnoteTakeWithAfterSpacingAndAfter(1, 480);
        TestAssert.Equal(17, continuationTake);
    }

    public static void DocxFootnoteContinuationTakeChargesRuleOverhead()
    {
        // RV06 continuation probes: the footnote bottom-anchor query charges the rule
        // overhead net of already-reserved top space, so the first continuation page
        // takes fewer lines than the full-frame fit allows. Same synthetic note shape
        // as the trailing-after test at 10pt after-spacing, counting the first
        // continuation page, where a zeroed overhead takes 29 against 28 charged.
        int continuationTake = LayoutFootnoteTakeWithAfterSpacingAndAfter(1, 200);
        TestAssert.Equal(28, continuationTake);
    }

    private static int LayoutFootnoteTakeWithAfterSpacingAndAfter(int pageIndex, int afterTwips)
    {
        var footnoteParas = new System.Text.StringBuilder();
        for (int line = 0; line < 60; line++)
        {
            footnoteParas.Append("<w:p><w:pPr><w:spacing w:after=\"" + afterTwips.ToString(CultureInfo.InvariantCulture) + "\"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Calibri\" w:hAnsi=\"Calibri\"/><w:sz w:val=\"24\"/><w:szCs w:val=\"24\"/></w:rPr><w:t xml:space=\"preserve\">Note body line " + line.ToString(CultureInfo.InvariantCulture) + " start tail</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=\"Tahoma\" w:hAnsi=\"Tahoma\"/><w:sz w:val=\"24\"/><w:szCs w:val=\"24\"/></w:rPr><w:t xml:space=\"preserve\"> mixed tail</w:t></w:r></w:p>");
        }

        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2">""" + footnoteParas.ToString() + """</w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DescDeficitTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        return layout.Pages[pageIndex].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Count();
    }

    public static void DocxFootnoteContentGapAddsWindowsAscenderExcess()
    {
        // RV06 fnwrap content-gap probes (Word 16.0): Office hangs big-ascender content
        // lower by the Windows-ascender excess over the mark, so a synthetic big-ascender
        // content face at 1.05em against a 0.9521em mark gains 0.94pt of gap at 12pt
        // while uniform content keeps the shared value.
        double control = LayoutFootnoteContentGapWithWindowsAscender("Calibri", "Calibri");
        double mixed = LayoutFootnoteContentGapWithWindowsAscender("Calibri", "BigAsc");
        TestAssert.True(Math.Abs(control - 3.37d) < 0.01d, "Uniform footnote gap must keep the shared value; control=" + control.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(mixed - 4.54d) < 0.01d, "Big-ascender content gap must add the Windows excess; mixed=" + mixed.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private sealed class FamilyWascTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer, IDocxStaticTextMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public bool TryGetSingleLineEm(DocxTextRun? run, out double value)
        {
            value = 1.2207d;
            return true;
        }

        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => (string.Equals(run?.FontFamily, "BigAsc", StringComparison.Ordinal) ? 1.05d : 0.9521d) * fontSize;

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => inner is IDocxStaticTextMetricsProvider staticMetrics ? staticMetrics.MeasureWindowsDescender(run, fontSize) : fontSize * 0.2d;
    }

    private sealed class TypoWascTextMeasurer(IDocxTextMeasurer inner, bool useTypographicMetrics) : IDocxTextMeasurer, IDocxStaticTextMetricsProvider, IDocxTypographicMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public bool TryGetSingleLineEm(DocxTextRun? run, out double value)
        {
            value = 1.2207d;
            return true;
        }

        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => (string.Equals(run?.FontFamily, "BigAsc", StringComparison.Ordinal) ? 1.05d : 0.9521d) * fontSize;

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => inner is IDocxStaticTextMetricsProvider staticMetrics ? staticMetrics.MeasureWindowsDescender(run, fontSize) : fontSize * 0.2d;

        public bool UseTypographicMetrics(DocxTextRun? run) => useTypographicMetrics;
    }

    private static double LayoutFootnoteContentGapWithTypographicAscender(string markFamily, string contentFamily)
    {
        return LayoutFootnoteContentGapWithWindowsAscender(markFamily, contentFamily, true);
    }

    private static double LayoutFootnoteContentGapWithWindowsAscender(string markFamily, string contentFamily)
    {
        return LayoutFootnoteContentGapWithWindowsAscender(markFamily, contentFamily, false);
    }

    public static void DocxFootnoteContentGapScalesShelterWithRunSize()
    {
        // RV06 endnote mark-size probes (m28/m28c): the inset shelter must scale with the run size, so big marks on small content keep the run-sized shelter instead of releasing the supplement early. Same-size pairs behave identically either way.
        double same = LayoutFootnoteContentGapWithSizedWindowsAscender("Calibri", 28, "BigAsc", 28);
        double mixed = LayoutFootnoteContentGapWithSizedWindowsAscender("Calibri", 28, "BigAsc", 20);
        TestAssert.True(Math.Abs(same - 5.30d) < 0.02d, "Same-size content gap must keep the mark-sized shelter; same=" + same.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(mixed - 4.91d) < 0.02d, "Mixed-size content gap must shelter with the run size; mixed=" + mixed.ToString(CultureInfo.InvariantCulture) + ".");
    }
    private static double LayoutFootnoteContentGapWithSizedWindowsAscender(string markFamily, int markHalfPoints, string contentFamily, int contentHalfPoints)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/><Relationship Id="rIdS1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>""",
            ["word/styles.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="MFAM" w:hAnsi="MFAM"/><w:sz w:val="MSZ"/><w:szCs w:val="MSZ"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>""".Replace("MFAM", markFamily).Replace("MSZ", markHalfPoints.ToString(CultureInfo.InvariantCulture)),
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:rPr><w:rFonts w:ascii="MFAM" w:hAnsi="MFAM"/><w:sz w:val="MSZ"/><w:szCs w:val="MSZ"/></w:rPr><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="CFAM" w:hAnsi="CFAM"/><w:sz w:val="CSZ"/><w:szCs w:val="CSZ"/></w:rPr><w:t xml:space="preserve">Note body one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="CFAM" w:hAnsi="CFAM"/><w:sz w:val="CSZ"/><w:szCs w:val="CSZ"/></w:rPr><w:t xml:space="preserve">Note body two</w:t></w:r></w:p></w:footnote></w:footnotes>""".Replace("MFAM", markFamily).Replace("CFAM", contentFamily).Replace("MSZ", markHalfPoints.ToString(CultureInfo.InvariantCulture)).Replace("CSZ", contentHalfPoints.ToString(CultureInfo.InvariantCulture))
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new TypoWascTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), false), CancellationToken.None);
        DocxPlacedRelatedStoryLayout separator = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
        DocxPlacedRelatedStoryLayout content = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal));
        return (separator.TopY - separator.Height) - content.TopY;
    }
    private static double LayoutFootnoteContentGapWithWindowsAscender(string markFamily, string contentFamily, bool useTypographicMetrics)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:rPr><w:rFonts w:ascii="MFAM" w:hAnsi="MFAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="CFAM" w:hAnsi="CFAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="CFAM" w:hAnsi="CFAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body two</w:t></w:r></w:p></w:footnote></w:footnotes>""".Replace("MFAM", markFamily).Replace("CFAM", contentFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new TypoWascTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), useTypographicMetrics), CancellationToken.None);
        DocxPlacedRelatedStoryLayout separator = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
        DocxPlacedRelatedStoryLayout content = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal));
        return (separator.TopY - separator.Height) - content.TopY;
    }

    public static void DocxFootnoteContentGapSkipsTypographicSupplement()
    {
        // RV06 Aptos probes (Word 16.0, uniform Aptos lacks the excess while Segoe keeps
        // 1.524 of it): only Aptos sets OS/2 USE_TYPO_METRICS among probed families, so a
        // typographic content face with big-ascender metrics keeps the shared gap value
        // while the same face without the flag gains the Windows excess.
        double control = LayoutFootnoteContentGapWithWindowsAscender("Calibri", "Calibri");
        double flagged = LayoutFootnoteContentGapWithTypographicAscender("Calibri", "BigAsc");
        double unflagged = LayoutFootnoteContentGapWithWindowsAscender("Calibri", "BigAsc");
        TestAssert.True(Math.Abs(control - 3.37d) < 0.01d, "Uniform footnote gap must keep the shared value; control=" + control.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(flagged - 3.37d) < 0.01d, "Typographic content gap must skip the supplement; flagged=" + flagged.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(unflagged - 4.54d) < 0.01d, "Unflagged big-ascender content gap must keep the Windows excess; unflagged=" + unflagged.ToString(CultureInfo.InvariantCulture) + ".");
    }

    public static void DocxFootnoteContentGapYieldsToCoveredInset()
    {
        // RV06 Tah-Verdana probes: the Windows excess must yield where the first-line
        // inset already covers it, so Tahoma-first Verdana-second content keeps the
        // shared value while uniform big-ascender content keeps its excess.
        double control = LayoutFootnoteContentGapWithCoveredInset("Calibri", "Calibri");
        double mixed = LayoutFootnoteContentGapWithCoveredInset("Calibri", "Tahoma", "Verdana");
        TestAssert.True(Math.Abs(control - 3.37d) < 0.01d, "Uniform footnote gap must keep the shared value; control=" + control.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(mixed - 3.37d) < 0.01d, "Covered-inset content gap must yield to the shared value; mixed=" + mixed.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private sealed class FamilyCoveredInsetTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => inner is IDocxLineMetricsProvider lineMetrics ? lineMetrics.MeasureSingleLineHeight(run, fontSize) : fontSize;

        public double MeasureHheaAscender(DocxTextRun? run, double fontSize)
        {
            double em = string.Equals(run?.FontFamily, "Tahoma", StringComparison.Ordinal) ? 1.0005d : string.Equals(run?.FontFamily, "Verdana", StringComparison.Ordinal) ? 1.0054d : 0.75d;
            return em * fontSize;
        }

        public bool TryGetSingleLineEm(DocxTextRun? run, out double value)
        {
            value = 1.2207d;
            return true;
        }

        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize)
        {
            double em = string.Equals(run?.FontFamily, "Tahoma", StringComparison.Ordinal) ? 1.0005d : string.Equals(run?.FontFamily, "Verdana", StringComparison.Ordinal) ? 1.0054d : 0.9521d;
            return em * fontSize;
        }

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => inner is IDocxStaticTextMetricsProvider staticMetrics ? staticMetrics.MeasureWindowsDescender(run, fontSize) : fontSize * 0.2d;
    }

    private static double LayoutFootnoteContentGapWithCoveredInset(string markFamily, string firstFamily, string secondFamily = "")
    {
        string secondRun = string.IsNullOrEmpty(secondFamily) ? "" : "<w:r><w:rPr><w:rFonts w:ascii=\"F2\" w:hAnsi=\"F2\"/><w:sz w:val=\"24\"/><w:szCs w:val=\"24\"/></w:rPr><w:t xml:space=\"preserve\"> mixed tail</w:t></w:r>";
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:rPr><w:rFonts w:ascii="MFAM" w:hAnsi="MFAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body one start tail one</w:t></w:r>" + secondRun + "</w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body two start tail two</w:t></w:r>" + secondRun + "</w:p></w:footnote></w:footnotes>""".Replace("MFAM", markFamily).Replace("F1", firstFamily).Replace("F2", secondFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new FamilyCoveredInsetTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxPlacedRelatedStoryLayout separator = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
        DocxPlacedRelatedStoryLayout content = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal));
        return (separator.TopY - separator.Height) - content.TopY;
    }

    public static void DocxFootnoteOverflowHeadSliceHangsContentByGap()
    {
        // RV06 mixlong probes (Word 16.0): overflowing head slices hang content below
        // the take-driven rule like fitting notes, so a synthetic big-ascender overflow
        // places its first baseline lower by the Windows excess at 12pt while takes,
        // take heights and uniform content stay put.
        double control = LayoutFootnoteOverflowHeadFirstBaseline("Calibri", "Calibri");
        double mixed = LayoutFootnoteOverflowHeadFirstBaseline("Calibri", "BigAsc");
        TestAssert.True(Math.Abs((control - mixed) - 1.17d) < 0.02d, "Overflow head slice must hang content by the Windows excess; control=" + control.ToString(CultureInfo.InvariantCulture) + " mixed=" + mixed.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private static double LayoutFootnoteOverflowHeadFirstBaseline(string firstFamily, string secondFamily)
    {
        var footnoteParas = new System.Text.StringBuilder();
        for (int line = 0; line < 60; line++)
        {
            footnoteParas.Append("<w:p><w:r><w:rPr><w:rFonts w:ascii=\"F1\" w:hAnsi=\"F1\"/><w:sz w:val=\"24\"/><w:szCs w:val=\"24\"/></w:rPr><w:t xml:space=\"preserve\">Note body line " + line.ToString(CultureInfo.InvariantCulture) + " start tail</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=\"F2\" w:hAnsi=\"F2\"/><w:sz w:val=\"24\"/><w:szCs w:val=\"24\"/></w:rPr><w:t xml:space=\"preserve\"> mixed tail</w:t></w:r></w:p>".Replace("F1", firstFamily).Replace("F2", secondFamily));
        }

        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:rPr><w:rFonts w:ascii="MFAM" w:hAnsi="MFAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2">""" + footnoteParas.ToString() + """</w:footnote></w:footnotes>""".Replace("MFAM", firstFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new FamilyWascTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxTextLineLayout firstLine = layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .First();
        return firstLine.BaselineY;
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

    private static int LayoutEndnoteTakeWithAfterSpacingAndAfter(int pageIndex, int afterTwips)
    {
        var noteParas = new System.Text.StringBuilder();
        for (int line = 0; line < 60; line++)
        {
            noteParas.Append("<w:p><w:pPr><w:spacing w:after=\"" + afterTwips.ToString(CultureInfo.InvariantCulture) + "\"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Calibri\" w:hAnsi=\"Calibri\"/><w:sz w:val=\"24\"/><w:szCs w:val=\"24\"/></w:rPr><w:t xml:space=\"preserve\">Note body line " + line.ToString(CultureInfo.InvariantCulture) + " start tail</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=\"Tahoma\" w:hAnsi=\"Tahoma\"/><w:sz w:val=\"24\"/><w:szCs w:val=\"24\"/></w:rPr><w:t xml:space=\"preserve\"> mixed tail</w:t></w:r></w:p>");
        }

        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:endnotePr><w:pos w:val="sectEnd"/></w:endnotePr></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2">""" + noteParas.ToString() + """</w:endnote></w:endnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DescDeficitTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        return layout.Pages[pageIndex].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Count();
    }

    public static void DocxFootnoteNinePtHeadTakeReservesClampMargin()
    {
        // RV06 size sweep (Word 16.0, cal9-a0 takes 47 against 48 while 10pt and
        // first-10pt boundaries hold exact): uniform 9pt-after-0 head takes 53 under the
        // 4.0pt uniform reserve against 54 under the 2.5pt reserve; the Office-anchored
        // take-fewer direction is pinned by the cal9-a0 47/13 probe takes with size-10
        // 43/17 and first-10 36/24 holding exact.
        int ninePtHeadTake = LayoutFootnoteUniformHeadTakeWithAfterAndSize(0, 0);
        TestAssert.Equal(53, ninePtHeadTake);
    }

    public static void DocxFootnoteHeadTakeReservesClampMargin()
    {
        // RV06 take-battery probes (Word 16.0): Office head takes fit one more
        // line than clamp-minus-separator allows, so the clamp carries a 2.5pt
        // reserve shaved from the take side, pinned where a 10pt after-spacing
        // head take drops from 29 to 28 against the untrimmed fit.
        int headTake = LayoutFootnoteUniformHeadTakeWithAfter(200);
        TestAssert.Equal(28, headTake);
    }

    private static int LayoutFootnoteUniformHeadTakeWithAfter(int afterTwips)
    {
        return LayoutFootnoteUniformHeadTakeWithAfterAndSize(afterTwips, 24);
    }

    private static int LayoutFootnoteUniformHeadTakeWithAfterAndSize(int afterTwips, int sizeHalfPoints)
    {
        var footnoteParas = new System.Text.StringBuilder();
        for (int line = 0; line < 60; line++)
        {
            footnoteParas.Append("<w:p><w:pPr><w:spacing w:after=\"" + afterTwips.ToString(CultureInfo.InvariantCulture) + "\"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Calibri\" w:hAnsi=\"Calibri\"/><w:sz w:val=\"" + sizeHalfPoints.ToString(CultureInfo.InvariantCulture) + "\"/><w:szCs w:val=\"" + sizeHalfPoints.ToString(CultureInfo.InvariantCulture) + "\"/></w:rPr><w:t xml:space=\"preserve\">Note body line " + line.ToString(CultureInfo.InvariantCulture) + " uniform tail</w:t></w:r></w:p>");
        }

        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2">""" + footnoteParas.ToString() + """</w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DescDeficitTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        return layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Count();
    }

    public static void DocxFootnoteFirstBaselineIgnoresLaterRunFonts()
    {
        // RV05 fnmix probes (Word COM references edge-fnmix/edge-fnmix3: note first
        // baselines sit at 85.46 regardless of run order or family): related-story
        // content keeps legacy widest-run insets, so a trailing Tahoma run must not
        // move the first baseline. Pre-scoping (max-hhea) it drops by 0.73.
        double control = LayoutFootnoteFirstBaselineWithSecondRun("Calibri");
        double mixed = LayoutFootnoteFirstBaselineWithSecondRun("Tahoma");
        TestAssert.True(Math.Abs(mixed - control) < 0.000001d, "Footnote first baseline must ignore later-run fonts; mixed=" + mixed.ToString(CultureInfo.InvariantCulture) + " control=" + control.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private sealed class StoryHheaTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer, IDocxLineMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => inner is IDocxLineMetricsProvider lineMetrics ? lineMetrics.MeasureSingleLineHeight(run, fontSize) : fontSize;

        public double MeasureHheaAscender(DocxTextRun? run, double fontSize)
        {
            double em = run?.FontFamily == "Tahoma" ? 1.0005d : 0.75d;
            return em * fontSize;
        }
    }

    private static double LayoutFootnoteFirstBaselineWithSecondRun(string secondFamily)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note start words </w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="FAM" w:hAnsi="FAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">middle mixed words</w:t></w:r></w:p></w:footnote></w:footnotes>""".Replace("FAM", secondFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new StoryHheaTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        return layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .First(line => line.Text.StartsWith("Note start words", StringComparison.Ordinal)).BaselineY;
    }
    public static void DocxFootnoteLineHeightUsesMaxAcrossRuns()
    {
        double control = LayoutFootnotePitchWithRuns("Calibri", "Calibri");
        double tall = LayoutFootnotePitchWithRuns("Tahoma", "Tahoma");
        double mixed = LayoutFootnotePitchWithRuns("Calibri", "Tahoma");
        double swapped = LayoutFootnotePitchWithRuns("Tahoma", "Calibri");
        TestAssert.True(Math.Abs(mixed - tall) < 0.000001d, "Footnote line height must use max across runs; mixed=" + mixed.ToString(CultureInfo.InvariantCulture) + " tall=" + tall.ToString(CultureInfo.InvariantCulture) + " control=" + control.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(swapped - tall) < 0.000001d, "Footnote line height must ignore run order; swapped=" + swapped.ToString(CultureInfo.InvariantCulture) + " tall=" + tall.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private sealed class StoryHeightTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer, IDocxLineMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "Tahoma", StringComparison.Ordinal) ? 20d : 10d;
    }

    private static double LayoutFootnotePitchWithRuns(string firstFamily, string secondFamily)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body one start tail one</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> mixed tail one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body two start tail two</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> mixed tail two</w:t></w:r></w:p></w:footnote></w:footnotes>""".Replace("F1", firstFamily).Replace("F2", secondFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new StoryHeightTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxTextLineLayout[] lines = layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Where(line => line.Text.StartsWith("Note body", StringComparison.Ordinal))
            .Take(2)
            .ToArray();
        TestAssert.Equal(2, lines.Length);
        return lines[0].BaselineY - lines[1].BaselineY;
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

    private static int CountSeparatorVariantTextShows(string separatorParagraph, DocxTests.SingleResolutionFontResolver resolver)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with endnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0">SEPPARA</w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body one</w:t></w:r></w:p><w:p><w:r><w:t>Note body two</w:t></w:r></w:p></w:endnote></w:endnotes>""".Replace("SEPPARA", separatorParagraph)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        PdfPage page = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();
        return DocxTests.CountPdfTextShows(page.Content);
    }
    public static void DocxFootnoteSeparatorIgnoresExactLineSpacing()
    {
        // Shared story-layout path with endnote separators (Office evidence is
        // endnote-only; no diverging footnote evidence): the footnote rule rides
        // the separator bottom, so an exact line rule that Office ignores would
        // move our rule with the taller story.
        double ruleExact = LayoutFootnoteSeparatorRuleWithSpacing("<w:pPr><w:spacing w:line=\"480\" w:lineRule=\"exact\"/></w:pPr>");
        double ruleAuto = LayoutFootnoteSeparatorRuleWithSpacing(string.Empty);
        TestAssert.True(Math.Abs(ruleExact - ruleAuto) < 0.000001d, "Footnote separator exact line spacing must match auto lines; observed shift=" + Math.Abs(ruleExact - ruleAuto).ToString(CultureInfo.InvariantCulture) + ".");
    }

    private static double LayoutEndnoteBodyTopWithSeparatorSpacing(string separatorParagraphProperties)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with endnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p>""" + separatorParagraphProperties + """<w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body one</w:t></w:r></w:p><w:p><w:r><w:t>Note body two</w:t></w:r></w:p></w:endnote></w:endnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        return layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal)).TopY;
    }

    private static double LayoutFootnoteBodyTopWithSeparatorSpacing(string separatorParagraphProperties)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p>""" + separatorParagraphProperties + """<w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body one</w:t></w:r></w:p><w:p><w:r><w:t>Note body two</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        return layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal)).TopY;
    }

    private static double LayoutFootnoteSeparatorRuleWithSpacing(string separatorParagraphProperties)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p>""" + separatorParagraphProperties + """<w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body one</w:t></w:r></w:p><w:p><w:r><w:t>Note body two</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        return layout.Pages[0].PlacedRelatedStories.Single(story => story.SeparatorY is not null).SeparatorY ?? 0d;
    }
    }
