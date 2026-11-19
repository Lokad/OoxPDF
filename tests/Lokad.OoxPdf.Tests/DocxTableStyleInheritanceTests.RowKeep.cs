using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxTableStyleInheritanceTests
{
    public static void DocxTableStyleRowKeepControlsPageAdmission()
    {
        const string keep = """<w:trPr><w:cantSplit w:val="1"/></w:trPr>""";
        const string split = """<w:trPr><w:cantSplit w:val="0"/></w:trPr>""";
        const string noFirstRow = """<w:trPr><w:cnfStyle w:firstRow="0"/></w:trPr>""";
        const string bandSize = """<w:tblPr><w:tblStyleRowBandSize w:val="1"/></w:tblPr>""";
        const string first = """<w:tblLook w:firstRow="1"/>""";
        const string noFirst = """<w:tblLook w:firstRow="0"/>""";
        const string edges = """<w:tblLook w:firstRow="1" w:lastRow="1"/>""";
        const string bands = """<w:tblLook w:noHBand="0"/>""";
        const string noBands = """<w:tblLook w:noHBand="1"/>""";
        string firstKeep = Region("firstRow", keep);
        string lastKeep = Region("lastRow", keep);
        string bandKeep = Region("band1Horz", keep);
        string intro = string.Concat(Enumerable.Range(1, 4).Select(n => Line($"Intro{n}")));
        string row = string.Concat(Enumerable.Range(1, 6).Select(n => Line($"Row{n}")));
        const string section = """
            <w:pgSz w:w="6000" w:h="4000"/>
            <w:pgMar w:top="400" w:bottom="400" w:left="400" w:right="400" w:header="0" w:footer="0"/>
            """;
        const string defaults = """
            <w:docDefaults><w:rPrDefault><w:rPr>
              <w:rFonts w:ascii="DejaVu Sans" w:hAnsi="DejaVu Sans"/><w:sz w:val="20"/>
            </w:rPr></w:rPrDefault></w:docDefaults>
            """;

        foreach (var source in new (string Name, string Styles, string? StyleId, string RowProperties, string Look, string? EffectiveToken)[]
        {
            ("direct keep", Style("Probe", ""), "Probe", keep, first, "1"),
            ("named keep", Style("Probe", keep), "Probe", "", first, "1"),
            ("inherited keep", Style("Base", keep) + Child(""), "Probe", "", first, "1"),
            ("default keep", Style("Probe", keep, isDefault: true), null, "", first, "1"),
            ("conditional keep", Style("Probe", split + firstKeep), "Probe", "", first, "1"),
            ("inherited conditional keep", Style("Base", firstKeep) +
                Child(Region("firstRow", """<w:rPr><w:b w:val="0"/></w:rPr>""")), "Probe", "", first, "1"),
            ("trailing keep", Style("Probe", lastKeep), "Probe", "", edges, "1"),
            ("leading edge wins", Style("Probe", lastKeep + Region("firstRow", split)), "Probe", "", edges, "0"),
            ("band keep", Style("Probe", bandSize + bandKeep), "Probe", "", bands, "1"),
            ("disabled band", Style("Probe", bandSize + bandKeep), "Probe", "", noBands, null),
            ("disabled condition", Style("Probe", firstKeep), "Probe", "", noFirst, null),
            ("explicit row membership", Style("Probe", firstKeep), "Probe", noFirstRow, first, null),
            ("cell condition", Style("Probe", Region("firstCol", keep)), "Probe", "", first, null),
            ("child overrides base keep", Style("Base", keep) + Child(split), "Probe", "", first, "0"),
            ("direct override", Style("Probe", firstKeep), "Probe", split, first, "0"),
            ("no keep", Style("Probe", ""), "Probe", "", first, null)
        })
        {
            string styleSelector = source.StyleId is null ? "" : $"""<w:tblStyle w:val="{source.StyleId}"/>""";
            string body = $$"""
                {{intro}}<w:tbl><w:tblPr>{{styleSelector}}<w:tblW w:w="4000" w:type="dxa"/>{{source.Look}}</w:tblPr>
                  <w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>
                  <w:tr>{{source.RowProperties}}<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>{{row}}</w:tc></w:tr>
                </w:tbl>
                """;
            foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
            {
                bool keepTogether = source.EffectiveToken == "1";
                var document = ReadStyledTableProbe(body, mode, section, defaults + source.Styles);
                var parsed = document.Tables.Single().Rows.Single();
                TestAssert.True((keepTogether, source.EffectiveToken) == (parsed.CantSplit, parsed.CantSplitValue),
                    $"{source.Name} in {mode} should supply the effective keep setting and its token.");

                var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode,
                    mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                        OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
                PdfPage[] pages = renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
                TestAssert.Equal(2, pages.Length);
                string pageOne = keepTogether ? "Intro1 Intro2 Intro3 Intro4" : "Intro1 Intro2 Intro3 Intro4 Row1 Row2 Row3 Row4";
                string pageTwo = keepTogether ? "Row1 Row2 Row3 Row4 Row5 Row6" : "Row5 Row6";
                string firstPageText = Text(pages[0]);
                string secondPageText = Text(pages[1]);
                TestAssert.True(firstPageText == pageOne && secondPageText == pageTwo,
                    $"{source.Name} in {mode} should paint [{pageOne}] and [{pageTwo}], found [{firstPageText}] and [{secondPageText}].");
            }
        }

        static string Text(PdfPage page)
        {
            string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                includeTextTransforms: true).Select(show => show.Text));
            return string.Join(' ', text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
        }
        static string Line(string text) =>
            Paragraph(text, """<w:spacing w:before="0" w:after="0" w:line="400" w:lineRule="exact"/>""");
        static string Region(string name, string content) =>
            $"""<w:tblStylePr w:type="{name}">{content}</w:tblStylePr>""";
        static string Child(string content) => Style("Probe", """<w:basedOn w:val="Base"/>""" + content);
        static string Style(string id, string content, bool isDefault = false) =>
            $"""<w:style w:type="table" w:styleId="{id}" w:default="{(isDefault ? "1" : "0")}">{content}</w:style>""";
    }
}
