using System.IO.Compression;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxFootnotesTests
{
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
        foreach (string restart in new[] { "eachSect", "eachPage" })
        {
            DocxDocument document = ReadSectionNumberingFixture(kind, (xml, settings, _) =>
            {
                foreach (XElement section in xml.Descendants(NoteWord + "sectPr"))
                    section.Element(NoteWord + kind + "Pr")!.Remove();
                settings.Descendants(NoteWord + kind + "Pr").Single().Add(
                    new XElement(NoteWord + "numRestart", new XAttribute(NoteWord + "val", restart)));
            });
            CheckNoteLabels(document, kind == "footnote" ? ["1", "2", "3", "4"] : ["i", "ii", "iii", "iv"]);
            DocxNoteReferenceSettings authored = kind == "footnote"
                ? document.Settings.FootnoteReferenceSettings : document.Settings.EndnoteReferenceSettings;
            TestAssert.Equal(restart, authored.NumberRestartValue!);
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

    public static void DocxSectionNoteNumbersUseExplicitSectionsDespiteDocumentEachPage()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        {
            DocxDocument document = ReadSectionNumberingFixture(kind, (xml, settings, _) =>
            {
                settings.Descendants(NoteWord + kind + "Pr").Single().Add(
                    new XElement(NoteWord + "numRestart", new XAttribute(NoteWord + "val", "eachPage")));
                SetNoteSection(xml.Descendants(NoteWord + "sectPr").Last(), kind, "decimalZero", "9", "continuous");
            });
            CheckNoteLabels(document, ["iv", "v", "11", "12"]);
            DocxNoteReferenceSettings authored = kind == "footnote"
                ? document.Settings.FootnoteReferenceSettings : document.Settings.EndnoteReferenceSettings;
            TestAssert.Equal("eachPage", authored.NumberRestartValue!);
            TestAssert.Equal(105, authored.NumberStart!.Value);
        }
    }

    public static void DocxSingleSectionNoteNumbersUseDefaultsDespiteDocumentEachPage()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (string format in new[] { "decimal", "decimalZero" })
        foreach ((string? sectionFormat, int? sectionStart, bool properties) in new[]
        {
            ((string?)null, (int?)null, false),
            ((string?)null, (int?)null, true),
            ((string?)"decimalZero", (int?)null, true),
            ((string?)null, (int?)9, true),
            ((string?)"upperRoman", (int?)4, true),
        })
        {
            DocxDocument document = ReadNotePrecedenceFixture(kind, format, sectionFormat, sectionStart, properties,
                (_, settings) => settings.Descendants(NoteWord + kind + "Pr").Single().Add(
                    new XElement(NoteWord + "numRestart", new XAttribute(NoteWord + "val", "eachPage"))));
            string[] expected = sectionFormat == "upperRoman" ? ["IV", "V"] :
                sectionFormat == "decimalZero" ? ["01", "02"] :
                sectionStart == 9 ? kind == "endnote" ? ["ix", "x"] : ["9", "10"] :
                kind == "endnote" ? ["i", "ii"] : ["1", "2"];
            CheckNoteLabels(document, expected);
            DocxNoteReferenceSettings authored = kind == "footnote"
                ? document.Settings.FootnoteReferenceSettings : document.Settings.EndnoteReferenceSettings;
            TestAssert.Equal(format, authored.NumberFormatValue!);
            TestAssert.Equal(105, authored.NumberStart!.Value);
            TestAssert.Equal("eachPage", authored.NumberRestartValue!);
        }
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
            (_, settings, _) => settings.Descendants(NoteWord + kind + "Pr").Single().Add(new XElement(NoteWord + "numRestart", new XAttribute(NoteWord + "val", "unrecognized")))
        })
            CheckNoteLabels(ReadSectionNumberingFixture(kind, modify), ["105", "106", "107", "108"]);
    }

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
            PdfPage[] pages = renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().ToArray();
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
                new XAttribute(word + "val", "unrecognized")))), ["105", "106"]);
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

}
