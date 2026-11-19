using System.Numerics;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PortableRenderingTests
{
    public static void PortableRenderingUnplacedCellBoxesPreserveSiblingAndFollowingContent()
    {
        string valid = Box(1, 1828800, "Placed box");
        string unplaced = Box(2, 0, "Zero width box");
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool fixedLine in new[] { false, true })
        {
            var baseline = Render(mode, fixedLine, valid);
            var expectedBox = ReadPaintedSourceLabel(baseline, "Placed box");
            var expectedBody = ReadPaintedSourceLabel(baseline, "Following body");
            foreach (string boxes in new[] { unplaced + valid, valid + unplaced })
            {
                var actual = Render(mode, fixedLine, boxes);
                AssertSameOrigins(expectedBox, ReadPaintedSourceLabel(actual, "Placed box"), "Cell box");
                AssertSameOrigins(expectedBody, ReadPaintedSourceLabel(actual, "Following body"), "Following body");
            }

            var twoBoxes = Render(mode, fixedLine, Box(2, 1828800, "Prior box") + valid);
            var prior = ReadPaintedSourceLabel(twoBoxes, "Prior box");
            var placed = ReadPaintedSourceLabel(twoBoxes, "Placed box");
            var following = ReadPaintedSourceLabel(twoBoxes, "Following body");
            TestAssert.True(prior[0].Origin.Y > placed[0].Origin.Y && placed[0].Origin.Y > following[0].Origin.Y,
                "Admitted boxes and following prose paint in source order.");
        }

        static void AssertSameOrigins(IReadOnlyList<(Vector2 Origin, double Size, PdfRectangle Ink)> expected,
            IReadOnlyList<(Vector2 Origin, double Size, PdfRectangle Ink)> actual, string label)
        {
            TestAssert.Equal(expected.Count, actual.Count);
            foreach (var pair in expected.Zip(actual))
                TestAssert.True(Vector2.Distance(pair.First.Origin, pair.Second.Origin) < .002f,
                    label + " keeps its painted position when an adjacent box is unplaced.");
        }

        static PdfPage Render(OoxPdfDocxMarkupMode mode, bool fixedLine, string boxes)
        {
            string spacing = fixedLine ? " w:line=\"480\" w:lineRule=\"exact\"" : "";
            using var stream = TestFixtures.CreateZipPackage(new Dictionary<string, string>
            {
                ["[Content_Types].xml"] = """
                    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                    </Types>
                    """,
                ["_rels/.rels"] = """
                    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                    </Relationships>
                    """,
                ["word/document.xml"] = $$"""
                    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                                xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                                xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                      <w:body>
                        <w:tbl><w:tblPr><w:tblW w:type="dxa" w:w="4000"/></w:tblPr>
                          <w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>
                            <w:tcPr><w:tcW w:type="dxa" w:w="4000"/></w:tcPr>
                            <w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="exact"/></w:pPr>
                              <w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>Cell text</w:t></w:r></w:p>
                            <w:p><w:pPr><w:spacing w:before="0" w:after="0"{{spacing}}/></w:pPr>{{boxes}}</w:p>
                          </w:tc></w:tr>
                        </w:tbl>
                        <w:p><w:r><w:t>Following body</w:t></w:r></w:p>
                        <w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
                          <w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/>
                        </w:sectPr>
                      </w:body>
                    </w:document>
                    """
            });
            using var package = OoxPackage.Open(stream, default);
            var document = new DocxReader().Read(package, null, default, markupMode: mode);
            return RenderAlignmentProbe(document, mode).Single();
        }

        static string Box(int id, int width, string text) => $$"""
            <w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">
              <wp:extent cx="{{width}}" cy="457200"/><wp:docPr id="{{id}}" name="{{text}}"/>
              <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{{width}}" cy="457200"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>
                  <wps:txbx><w:txbxContent><w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>{{text}}</w:t></w:r></w:p></w:txbxContent></wps:txbx>
                  <wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/>
                </wps:wsp>
              </a:graphicData></a:graphic>
            </wp:inline></w:drawing></w:r>
            """;
    }

    public static void PortableRenderingJustifiedInlineBoxChildKeepsPreparedParentContent()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var baseline = RenderAlignmentProbe(DocxInlineCompositionTests.ReadMutation("InlineBoxMath/inline-box-body-equation-inline-box-valid",
                mode, false, false), mode).Single();
            var expected = ReadPaintedSourceLabel(baseline, "x+2y", mathOnly: true);
            TestAssert.Equal(4, expected.Count);
            foreach (bool unsupportedChild in new[] { false, true })
            {
                var source = DocxInlineCompositionTests.ReadMutation("InlineBoxMath/inline-box-body-equation-inline-box-valid", mode, false, false,
                    (xml, w, m) =>
                    {
                        var paragraph = xml.Descendants(w + "txbxContent").Single().Elements(w + "p").Single();
                        if (unsupportedChild) paragraph.Add(new XElement(m + "oMath", new XElement(m + "future")));
                        else SetJustified(paragraph, w);
                    });
                var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
                var page = RenderAlignmentProbe(source, mode, diagnostics.Add).Single();
                var actual = ReadPaintedSourceLabel(page, "x+2y", mathOnly: true);
                TestAssert.Equal(expected.Count, actual.Count);
                foreach (var pair in expected.Zip(actual)) AssertSamePaintOrigin(pair.First, pair.Second);
                ReadPaintedSourceLabel(page, "Inline box content.");
                TestAssert.Equal(unsupportedChild ? 1 : 0, diagnostics.Count(d => d.Id == "DOCX_UNSUPPORTED_EQUATION"));
            }
        }
    }

    public static void PortableRenderingJustifiedStaticParagraphsKeepFooterInkAndPhysicalFrames()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var baseline = RenderAlignmentProbe(DocxInlineCompositionTests.ReadMutation("Composition/report", mode, false, false), mode);
            foreach (string part in new[] { "word/header1.xml", "word/footer1.xml" })
            {
                var source = DocxInlineCompositionTests.ReadMutation("Composition/report", mode, false, false, (xml, w, m) =>
                    SetJustified(xml.Descendants(w + "p").First(p => !p.Descendants(m + "oMath").Any() &&
                        p.Descendants(w + "t").Any(t => t.Value.Length != 0)), w), partOverride: part);
                var pages = RenderAlignmentProbe(source, mode);
                TestAssert.Equal(baseline.Length, pages.Length);
                for (int index = 0; index < pages.Length; index++)
                {
                    const string footer = "Public synthetic composition portfolio";
                    var before = ReadPaintedSourceLabel(baseline[index], footer);
                    var after = ReadPaintedSourceLabel(pages[index], footer);
                    TestAssert.Equal(before.Count, after.Count);
                    AssertSamePaintOrigin(before[0], after[0]);
                    AssertSameStaticMathPaint(baseline[index], pages[index]);
                }
            }
        }
    }

    public static void PortableRenderingPlainStaticCountersKeepFooterInkAndIndependentFrames()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var baseline = RenderAlignmentProbe(DocxInlineCompositionTests.ReadMutation("Composition/report", mode, false, false), mode);
            foreach (string part in new[] { "word/header1.xml", "word/footer1.xml" })
            {
                var source = DocxInlineCompositionTests.ReadMutation("Composition/report", mode, false, false, (xml, w, m) =>
                {
                    var paragraph = xml.Descendants(w + "p").First(p => !p.Descendants(m + "oMath").Any() &&
                        p.Descendants(w + "t").Any(t => t.Value.Length != 0));
                    void Text(string value) => paragraph.Add(new XElement(w + "r", new XElement(w + "t",
                        new XAttribute(XNamespace.Xml + "space", "preserve"), value)));
                    void Field(string instruction) => paragraph.Add(new XElement(w + "fldSimple",
                        new XAttribute(w + "instr", instruction), new XElement(w + "r", new XElement(w + "t", "99"))));
                    Text(" Page "); Field(" PAGE "); Text(" of "); Field(" NUMPAGES ");
                }, partOverride: part);
                var pages = RenderAlignmentProbe(source, mode);
                TestAssert.Equal(baseline.Length, pages.Length);
                for (int index = 0; index < pages.Length; index++)
                {
                    const string footer = "Public synthetic composition portfolio";
                    var before = ReadPaintedSourceLabel(baseline[index], footer);
                    var after = ReadPaintedSourceLabel(pages[index], footer);
                    TestAssert.Equal(before.Count, after.Count);
                    AssertSamePaintOrigin(before[0], after[0]);
                    ReadPaintedSourceLabel(pages[index], "Page " + (index + 1) + " of " + pages.Length);
                    AssertSameStaticMathPaint(baseline[index], pages[index]);
                }
            }
        }
    }

    public static void PortableRenderingStaticFallbackKeepsItsCounterpartVisible()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxInlineCompositionTests.ReadMutation("Composition/report", mode, false, false);
            var baseline = RenderAlignmentProbe(document, mode);
            var pictureSource = document.Paragraphs.SelectMany(paragraph => paragraph.Images).Single();
            foreach (double rotation in new[] { 0d, 90d })
            {
                var headers = document.PageSettings.HeaderBodyElementsByType.ToDictionary(pair => pair.Key, pair => pair.Value);
                var owner = headers["default"].OfType<DocxParagraphElement>()
                    .First(element => element.Paragraph.Equations.Count == 0 && element.Paragraph.Runs.Count != 0);
                var paragraph = owner.Paragraph;
                var picture = pictureSource with
                {
                    WidthPoints = 20d, HeightPoints = 20d, PictureFrame = new(20d, 20d), RotationDegrees = rotation,
                    SourceRunIndex = paragraph.Runs.Count - 1, SourceTextOffsetInRun = paragraph.Runs[^1].Text.Length
                };
                paragraph = paragraph with
                {
                    Images = [.. paragraph.Images, picture],
                    InlineContent = [.. paragraph.InlineContent ?? DocxInlineContent.FromPrefix(
                        paragraph.Runs, paragraph.Images, default, paragraph.InlineTextBoxes), new DocxInlinePicture(picture)]
                };
                headers["default"] = headers["default"].Select(element => ReferenceEquals(element, owner)
                    ? owner with { Paragraph = paragraph } : element).ToArray();
                var source = document with { PageSettings = document.PageSettings with { HeaderBodyElementsByType = headers } };
                var pages = RenderAlignmentProbe(source, mode);
                TestAssert.Equal(baseline.Length, pages.Length);
                for (int index = 0; index < pages.Length; index++)
                {
                    const string footer = "Public synthetic composition portfolio";
                    var expected = ReadPaintedSourceLabel(baseline[index], footer);
                    var actual = ReadPaintedSourceLabel(pages[index], footer);
                    TestAssert.Equal(expected.Count, actual.Count);
                    AssertSamePaintOrigin(expected[0], actual[0]);
                    string math = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(baseline[index], usePageCoordinates: true,
                        fontResourceFilter: name => name.StartsWith("Eq", StringComparison.Ordinal)).Select(show => show.Text));
                    ReadPaintedSourceLabel(pages[index], math, mathOnly: true);
                }
            }
        }
    }

    private static void AssertSameStaticMathPaint(PdfPage before, PdfPage after)
    {
        var expected = DocxCommentsTests.ReadEmbeddedGlyphTextShows(before, usePageCoordinates: true,
            fontResourceFilter: name => name.StartsWith("Eq", StringComparison.Ordinal));
        var actual = DocxCommentsTests.ReadEmbeddedGlyphTextShows(after, usePageCoordinates: true,
            fontResourceFilter: name => name.StartsWith("Eq", StringComparison.Ordinal));
        TestAssert.True(expected.Select(show => (show.Text, show.Size, show.X, show.Y))
            .SequenceEqual(actual.Select(show => (show.Text, show.Size, show.X, show.Y))),
            "Equivalent static paragraph edits retain neighboring prepared math and physical frames.");
    }

    private static PdfPage[] RenderAlignmentProbe(DocxDocument document, OoxPdfDocxMarkupMode mode,
        Action<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>? diagnosticSink = null) =>
        new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
            mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .RenderPagesAsync(document, diagnosticSink, default).ToBlockingEnumerable().ToArray();

    private static void SetJustified(XElement paragraph, XNamespace w)
    {
        var properties = paragraph.Element(w + "pPr");
        if (properties is null) { properties = new XElement(w + "pPr"); paragraph.AddFirst(properties); }
        properties.Element(w + "jc")?.Remove();
        properties.Add(new XElement(w + "jc", new XAttribute(w + "val", "both")));
    }

    private static void AssertSamePaintOrigin((Vector2 Origin, double Size, PdfRectangle Ink) before,
        (Vector2 Origin, double Size, PdfRectangle Ink) after) =>
        TestAssert.True(Vector2.Distance(before.Origin, after.Origin) < .002f && Math.Abs(before.Size - after.Size) < .001d,
            "The equivalent single-line alignment keeps its source origin and physical font size.");

    internal static IReadOnlyList<(Vector2 Origin, double Size, PdfRectangle Ink)> ReadPaintedSourceLabel(PdfPage page,
        string label, bool mathOnly = false, string? fontPrefix = null)
    {
        // Match source characters across text-show boundaries, then check their
        // font ink against each physical paint frame. Existing readers own PDF syntax.
        var characters = new List<(int Offset, int Length, ushort Glyph, OpenTypeFont Font, double Size,
            PdfRectangle Clip, Matrix3x2 Transform, bool? VisiblePaint)>();
        var bindings = new Dictionary<PdfEmbeddedFont, (OpenTypeFont Program, Dictionary<ushort, ushort> Glyphs)>();
        var text = new StringBuilder();
        foreach (var frame in DocxNoteClippingTests.ReadPhysicalPaintFrames(page))
        foreach (var show in DocxCommentsTests.ReadEmbeddedGlyphTextShows(new PdfPage(page.Width, page.Height, frame.Block, page.Fonts),
            usePageCoordinates: false, includeTextTransforms: true, honorActualText: false,
            fontResourceFilter: mathOnly ? name => name.StartsWith("Eq", StringComparison.Ordinal) :
                fontPrefix is not null ? name => name.StartsWith(fontPrefix, StringComparison.Ordinal) : null,
            glyphPaint: paint =>
        {
            if (!bindings.TryGetValue(paint.Font, out var binding))
            {
                var glyphs = new Dictionary<ushort, ushort>();
                foreach (ushort original in paint.Font.UnicodeByOriginalGlyph.Keys)
                {
                    TestAssert.True(paint.Font.TryGetEncodedCid(original, out ushort cid), "Every source outline has an embedded CID.");
                    glyphs[cid] = cid; // Canonical subset/whole-font CIDs address their program directly.
                }
                foreach (var alias in paint.Font.UnicodeByCid.Where(pair => !glyphs.ContainsKey(pair.Key)))
                {
                    ushort original = paint.Font.UnicodeByOriginalGlyph.Keys.Single(glyph =>
                        paint.Font.TryGetEncodedCid(glyph, alias.Value, out ushort cid) && cid == alias.Key);
                    TestAssert.True(paint.Font.TryGetEncodedCid(original, out ushort canonical), "Source aliases keep a canonical outline.");
                    glyphs.Add(alias.Key, canonical);
                }
                binding = (OpenTypeFont.Load(paint.Font.FontProgramBytes.ToArray()), glyphs);
                bindings.Add(paint.Font, binding);
            }
            ushort glyph = binding.Glyphs[paint.Cid];
            var rune = new Rune(paint.Unicode);
            if (!Rune.IsWhiteSpace(rune))
            {
                string normalized = rune.ToString().Normalize(NormalizationForm.FormKC);
                characters.Add((text.Length, normalized.Length, glyph, binding.Program, paint.Size, frame.Clip,
                    paint.Matrix * frame.Transform, frame.VisibleTextPaint));
                text.Append(normalized);
            }
            return Math.Round(binding.Program.GetAdvanceWidth(glyph) * 1000d / binding.Program.UnitsPerEm) * paint.Size / 1000d;
        })) { }
        string wanted = string.Concat(label.Where(c => !char.IsWhiteSpace(c))).Normalize(NormalizationForm.FormKC);
        string actual = text.ToString();
        int start = actual.IndexOf(wanted, StringComparison.Ordinal);
        TestAssert.True(start >= 0 && start == actual.LastIndexOf(wanted, StringComparison.Ordinal),
            "Each complete source label paints once: " + label);
        var result = new List<(Vector2 Origin, double Size, PdfRectangle Ink)>();
        int covered = start;
        foreach (var character in characters.Where(c => c.Offset >= start && c.Offset < start + wanted.Length))
        {
            TestAssert.True(character.VisiblePaint is true,
                "Every source glyph must have known visible fill paint: " + label);
            // A normalized UTF-16 span can belong to one supplementary scalar
            // or compatibility glyph. Require complete span ownership rather
            // than equating UTF-16 length with the number of painted glyphs.
            TestAssert.Equal(covered, character.Offset);
            covered += character.Length;
            TestAssert.True(covered <= start + wanted.Length, "A source label must own complete glyphs: " + label);
            var font = character.Font;
            TestAssert.True(font.TryGetGlyphBounds(character.Glyph, out var ink), "Pinned fonts supply the source glyph ink.");
            double scale = character.Size / font.UnitsPerEm;
            double left = double.PositiveInfinity, bottom = double.PositiveInfinity;
            double right = double.NegativeInfinity, top = double.NegativeInfinity;
            foreach (double x in new[] { ink.XMin, ink.XMax })
            foreach (double y in new[] { ink.YMin, ink.YMax })
            {
                var point = Vector2.Transform(new((float)(x * scale), (float)(y * scale)), character.Transform);
                left = Math.Min(left, point.X); right = Math.Max(right, point.X);
                bottom = Math.Min(bottom, point.Y); top = Math.Max(top, point.Y);
                TestAssert.True(point.X >= character.Clip.X - .01d && point.X <= character.Clip.X + character.Clip.Width + .01d &&
                    point.Y >= character.Clip.Y - .01d && point.Y <= character.Clip.Y + character.Clip.Height + .01d,
                    "Every source glyph must remain inside the paper and effective clips: " + label);
            }
            result.Add((Vector2.Transform(Vector2.Zero, character.Transform), character.Size,
                new PdfRectangle(left, bottom, right - left, top - bottom)));
        }
        TestAssert.Equal(start + wanted.Length, covered);
        return result;
    }
}
