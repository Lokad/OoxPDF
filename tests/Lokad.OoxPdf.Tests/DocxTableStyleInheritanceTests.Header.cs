namespace Lokad.OoxPdf.Tests;

internal static partial class DocxTableStyleInheritanceTests
{
    public static void DocxTableStyleHeaderRowsRepeatOnlyTheLeadingGroup()
    {
        const string header = """<w:tblHeader w:val="1"/>""";
        const string noHeader = """<w:tblHeader w:val="0"/>""";
        const string keep = """<w:cantSplit/>""";
        string conditional = Region("firstRow", "<w:trPr>" + header + keep + "</w:trPr>");

        foreach (var source in new (string Name, string Styles, string FirstRow, string OtherRows,
            bool FirstRowEnabled, bool UseDefault, int ExtraHeaderIndex, int HeaderCount, bool FirstRowKept)[]
        {
            ("direct", Style("Probe", ""), header, "", true, false, -1, 1, false),
            ("named", Style("Probe", "<w:trPr>" + header + "</w:trPr>"), "", noHeader, true, false, -1, 1, false),
            ("default", Style("Probe", "<w:trPr>" + header + "</w:trPr>", isDefault: true), "", noHeader, true, true, -1, 1, false),
            ("conditional", Style("Probe", conditional), "", "", true, false, -1, 1, true),
            ("inherited", Style("Base", conditional) + Child(""), "", "", true, false, -1, 1, true),
            ("disabled", Style("Probe", conditional), "", "", false, false, -1, 0, false),
            ("direct off", Style("Probe", conditional), noHeader, "", true, false, -1, 0, true),
            ("two leading headers", Style("Probe", conditional), "", "", true, false, 1, 2, true),
            ("late direct marker", Style("Probe", ""), header, "", true, false, 6, 1, false)
        })
        {
            string styleSelector = source.UseDefault ? "" : """<w:tblStyle w:val="Probe"/>""";
            string first = source.FirstRowEnabled ? "1" : "0";
            string rows = string.Concat(Enumerable.Range(0, 9).Select(index =>
                Row(index == 0 ? "ColumnKey" : "Row" + index,
                    index == 0 ? source.FirstRow : index == source.ExtraHeaderIndex ? header : source.OtherRows)));
            string body = $$"""
                <w:tbl><w:tblPr>{{styleSelector}}<w:tblW w:w="4000" w:type="dxa"/><w:tblLook w:firstRow="{{first}}"/></w:tblPr>
                  <w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>{{rows}}
                </w:tbl>
                """;

            foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
            {
                var document = ReadStyledTableProbe(body, mode, RowProbeSection, RowProbeDefaults + source.Styles);
                var parsed = document.Tables.Single().Rows;
                var pages = RenderStyledTableProbe(document, mode);
                TestAssert.Equal(2, pages.Length);
                string pageOne = ReadRowProbePageText(pages[0]);
                string pageTwo = ReadRowProbePageText(pages[1]);
                string expectedPageTwo = source.HeaderCount switch
                {
                    0 => "Row6 Row7 Row8",
                    1 => "ColumnKey Row6 Row7 Row8",
                    _ => "ColumnKey Row1 Row6 Row7 Row8"
                };
                TestAssert.True(pageOne == "ColumnKey Row1 Row2 Row3 Row4 Row5" && pageTwo == expectedPageTwo,
                    $"{source.Name} in {mode} should repeat only leading headers, found [{pageOne}] and [{pageTwo}].");
                TestAssert.True(parsed.Select(row => row.IsHeader).SequenceEqual(Enumerable.Range(0, 9).Select(i => i < source.HeaderCount)),
                    $"{source.Name} in {mode} should classify only the contiguous header group.");
                TestAssert.Equal(source.FirstRowKept, parsed[0].CantSplit);
                if (source.ExtraHeaderIndex >= 0) TestAssert.Equal("1", parsed[source.ExtraHeaderIndex].HeaderValue);
            }
        }

        static string Row(string label, string flags) => $$"""
            <w:tr><w:trPr>{{flags}}<w:trHeight w:val="480" w:hRule="exact"/></w:trPr>
              <w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>
                {{Paragraph(label, """<w:spacing w:before="0" w:after="0" w:line="320" w:lineRule="exact"/>""")}}
              </w:tc>
            </w:tr>
            """;
    }
}
