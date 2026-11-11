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

    public static void DocxTableMinimumNoteLinksKeepSlotDuringBaselineCorrection()
    {
        DocxDocument UniformMinimum(string kind, int before)
        {
            DocxDocument source = ReadTableNoteSpacingFixture(kind, before, 24, variant: "atleast");
            DocxTable table = source.BodyElements.OfType<DocxTableElement>().Single().Table;
            DocxTableCell original = table.Rows[0].Cells[0];
            DocxParagraph[] paragraphs = original.Paragraphs.Select(p => p with
                { Runs = p.Runs.Select(r => r with { FontFamily = "Calibri", FontSize = 12d }).ToArray() }).ToArray();
            DocxTableCell cell = original with
                { Paragraphs = paragraphs, BodyElements = paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p)).ToArray() };
            DocxTable updated = table with { Rows = [table.Rows[0] with { Cells = [cell] }] };
            return source with { BodyElements = source.BodyElements.Select(e => e is DocxTableElement
                ? (DocxBodyElement)new DocxTableElement(updated) : e).ToArray(), FallbackTables = [updated] };
        }
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (OoxPdfDocxMarkupGeometryMode geometry in new[]
            { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
        {
            PdfLinkAnnotation baseline = RenderCustomNoteLinks(UniformMinimum(kind, 0), geometry)[0];
            PdfLinkAnnotation actual = RenderCustomNoteLinks(UniformMinimum(kind, 12), geometry)[0];
            TestAssert.True(Math.Abs(actual.Height - baseline.Height - 6d) < .000001d,
                "The corrected minimum-height baseline keeps its line slot and gives the first note mark its remaining before contribution.");
            TestAssert.True(Math.Abs(actual.Y + actual.Height - baseline.Y - baseline.Height) < .000001d,
                "The slot top remains at the same predecessor boundary while its bottom follows the painted baseline.");
            TestAssert.Equal(baseline.Width, actual.Width);
            TestAssert.Equal(baseline.Destination!.Value.PageIndex, actual.Destination!.Value.PageIndex);
        }
    }

    public static void DocxTableNoteBeforeSpacingMatchesIndependentWordOwnership()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (bool previousFlag in new[] { false, true })
        foreach (bool currentFlag in new[] { false, true })
        foreach (bool different in new[] { false, true })
        foreach (int before in new[] { 6, 12, 24 })
        {
            double slot = RenderCustomNoteLinks(ReadTableNoteSpacingFixture(kind, 0, 24, previousFlag, currentFlag, different))[0].Height;
            DocxDocument document = ReadTableNoteSpacingFixture(kind, before, 24, previousFlag, currentFlag, different);
            double expected = currentFlag && !different ? 0d : Math.Max(0d, before - 6d);
            TestAssert.True(Math.Abs(RenderCustomNoteLinks(document)[0].Height - slot - expected) < .000001d,
                "Word assigns the remaining collapsed before-gap to the first-line table note mark.");
        }
    }

    public static void DocxTableNoteBeforeSpacingIncludesCustomAndExactSlots()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (bool exact in new[] { false, true })
        foreach (OoxPdfDocxMarkupGeometryMode geometry in new[]
            { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
        {
            string variant = exact ? "custom-exact" : "custom-auto";
            PdfLinkAnnotation[] baseline = RenderCustomNoteLinks(ReadTableNoteSpacingFixture(kind, 0, 24, variant: variant), geometry);
            PdfLinkAnnotation[] actual = RenderCustomNoteLinks(ReadTableNoteSpacingFixture(kind, 12, 24, variant: variant), geometry);
            TestAssert.Equal(4, actual.Length);
            TestAssert.True(Math.Abs(actual[0].Height - baseline[0].Height - 6d) < .000001d,
                "Custom first-character marks and exact line slots own the same 6pt before contribution.");
            TestAssert.Equal(baseline[0].Width, actual[0].Width);
            TestAssert.Equal(baseline[0].Destination!.Value.PageIndex, actual[0].Destination!.Value.PageIndex);
        }
    }

    public static void DocxTableNoteBeforeSpacingPreservesFinalAfterAndAuthoredProperties()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        {
            DocxDocument emptyAfter = ReadTableNoteSpacingFixture(kind, 12, 0);
            DocxDocument finalAfter = ReadTableNoteSpacingFixture(kind, 12, 24);
            PdfLinkAnnotation emptyLink = RenderCustomNoteLinks(emptyAfter)[0];
            PdfLinkAnnotation finalLink = RenderCustomNoteLinks(finalAfter)[0];
            TestAssert.Equal(emptyLink.Height, finalLink.Height);
            TestAssert.Equal(emptyLink.Y, finalLink.Y);
            DocxTable table = finalAfter.BodyElements.OfType<DocxTableElement>().Single().Table;
            TestAssert.Equal(12d, table.Rows[0].Cells[0].Paragraphs[1].SpacingBeforePoints);
            TestAssert.Equal(24d, table.Rows[0].Cells[0].Paragraphs[1].SpacingAfterPoints);
            var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            DocxLayout zero = engine.Create(emptyAfter, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
            DocxLayout spaced = engine.Create(finalAfter, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
            double zeroHeight = zero.Pages[0].Items.OfType<DocxTableRowLayout>().Single().Height;
            double spacedHeight = spaced.Pages[0].Items.OfType<DocxTableRowLayout>().Single().Height;
            TestAssert.True(Math.Abs(spacedHeight - zeroHeight - 24d) < .000001d,
                "Word retains the final paragraph after-space inside the row rather than the multiline note link.");
        }
    }

    public static void DocxTableNoteBeforeSpacingRetainsLaterLineAndComplexFallbacks()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (string variant in new[] { "later", "rich-previous", "before-auto", "before-lines", "center", "two-rows" })
        {
            double slot = RenderCustomNoteLinks(ReadTableNoteSpacingFixture(kind, 0, 24, variant: variant))[0].Height;
            double actual = RenderCustomNoteLinks(ReadTableNoteSpacingFixture(kind, 12, 24, variant: variant))[0].Height;
            TestAssert.True(Math.Abs(actual - slot) < .000001d,
                "Later-line marks and unsupported paragraph/cell/row models keep the existing line-only hit area.");
        }
    }

    private static DocxDocument ReadTableNoteSpacingFixture(string kind, int before, int after,
        bool previousFlag = true, bool currentFlag = false, bool different = false, string variant = "") =>
        ReadSectionNumberingFixture(kind, (xml, _, parts) =>
        {
            XElement body = xml.Root!.Element(NoteWord + "body")!;
            XElement current = body.Elements(NoteWord + "p").First();
            XElement props = current.Element(NoteWord + "pPr")!;
            props.Element(NoteWord + "spacing")?.Remove();
            XElement spacing = new(NoteWord + "spacing", new XAttribute(NoteWord + "before", before * 20),
                new XAttribute(NoteWord + "after", after * 20), new XAttribute(NoteWord + "line", variant.Contains("exact") || variant == "atleast" ? "480" : "240"),
                new XAttribute(NoteWord + "lineRule", variant == "atleast" ? "atLeast" : variant.Contains("exact") ? "exact" : "auto"));
            props.Add(spacing);
            if (currentFlag) props.Add(new XElement(NoteWord + "contextualSpacing"));
            if (different) props.AddFirst(new XElement(NoteWord + "pStyle", new XAttribute(NoteWord + "val", "PublicOther")));
            if (variant == "before-auto") spacing.SetAttributeValue(NoteWord + "beforeAutospacing", "1");
            if (variant == "before-lines") spacing.SetAttributeValue(NoteWord + "beforeLines", "100");
            current.Descendants(NoteWord + "t").Last().Value += " " + string.Concat(Enumerable.Repeat("after note ", 35));
            if (variant == "later") current.Descendants(NoteWord + "t").First().Value += string.Concat(Enumerable.Repeat("before note ", 45));
            XElement previousProps = new(NoteWord + "pPr", new XElement(NoteWord + "spacing",
                new XAttribute(NoteWord + "before", "0"), new XAttribute(NoteWord + "after", "120")));
            if (previousFlag) previousProps.Add(new XElement(NoteWord + "contextualSpacing"));
            XElement run = new(NoteWord + "r", new XElement(NoteWord + "t", "Public preceding paragraph."));
            if (variant == "rich-previous") run.Add(new XElement(NoteWord + "br"));
            XElement previous = new(NoteWord + "p", previousProps, run);
            XElement cell = new(NoteWord + "tc", new XElement(NoteWord + "tcPr",
                new XElement(NoteWord + "tcW", new XAttribute(NoteWord + "w", "9360"), new XAttribute(NoteWord + "type", "dxa"))), previous);
            if (variant == "center") cell.Element(NoteWord + "tcPr")!.Add(new XElement(NoteWord + "vAlign", new XAttribute(NoteWord + "val", "center")));
            XElement table = new(NoteWord + "tbl", new XElement(NoteWord + "tblPr",
                new XElement(NoteWord + "tblW", new XAttribute(NoteWord + "w", "9360"), new XAttribute(NoteWord + "type", "dxa"))),
                new XElement(NoteWord + "tblGrid", new XElement(NoteWord + "gridCol", new XAttribute(NoteWord + "w", "9360"))), new XElement(NoteWord + "tr", cell));
            current.AddBeforeSelf(table); current.Remove(); cell.Add(current);
            if (variant == "two-rows") table.Add(new XElement(NoteWord + "tr", new XElement(NoteWord + "tc", new XElement(NoteWord + "p", new XElement(NoteWord + "r", new XElement(NoteWord + "t", "Guard row."))))));
            if (variant.StartsWith("custom", StringComparison.Ordinal))
            {
                XElement reference = current.Descendants(NoteWord + kind + "Reference").Single();
                reference.SetAttributeValue(NoteWord + "customMarkFollows", "1"); reference.AddAfterSelf(new XElement(NoteWord + "t", "*"));
                XDocument notes = XDocument.Parse(parts["word/" + kind + "s.xml"]);
                XElement note = notes.Root!.Elements(NoteWord + kind).Single(n => (string?)n.Attribute(NoteWord + "id") == (string?)reference.Attribute(NoteWord + "id"));
                XElement marker = note.Descendants(NoteWord + kind + "Ref").Single(); marker.Name = NoteWord + "t"; marker.Value = "*";
                parts["word/" + kind + "s.xml"] = notes.ToString();
            }
        });

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
}
