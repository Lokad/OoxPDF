using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxCommentsTests
{
    public static void DocxCommentNumericIdentityKeepsRangesThreadsAndHiddenOwners()
    {
        foreach (var (canonical, spellings) in new (string, string[])[]
        {
            ("42", ["0042", "+42", " \t+0042 \r\n"]),
            ("0", ["+000", "-000"]),
            ("-2", ["-0002"]),
            ("123456789012345678901234567890", ["+000123456789012345678901234567890"]),
            ("4x2", ["4x2"]), // Keep the existing malformed-ID fallback, without numeric coercion.
        })
        foreach (string spelling in spellings)
        for (int owner = 0; owner < 4; owner++)
        {
            string[] ids = [canonical, canonical, canonical, canonical];
            ids[owner] = spelling;
            var document = ReadCommentIdentitySource(ids[0], ids[1], ids[2], ids[3]);
            var paragraph = DocxBlockTraversal.EnumerateBodyParagraphs(document).Single();
            var reference = paragraph.InlineReferences.Single();
            TestAssert.Equal(1, paragraph.CommentRanges.Count);
            var range = paragraph.CommentRanges.Single();
            TestAssert.Equal(canonical, reference.Id!);
            TestAssert.Equal(canonical, range.Id!);
            TestAssert.True(range.StartSourceRunIndex is not null && range.EndSourceRunIndex is not null &&
                range.ReferenceSourceRunIndex is not null, "Equivalent numeric keys retain the complete source range and reference.");
            TestAssert.Equal(canonical, document.PackageCommentAnchorIds.Single());
            TestAssert.Equal(0, document.HiddenCommentAnchorIds.Count);
            var parent = document.RelatedStories.Single(story => story.CommentMetadata?.Author == "Reviewer");
            var reply = document.RelatedStories.Single(story => story.CommentMetadata?.Author == "Reply");
            TestAssert.Equal(canonical, parent.Id!);
            TestAssert.Equal(canonical, reply.CommentMetadata!.ParentCommentId!);
            TestAssert.Equal("00AB0001", parent.CommentMetadata!.ParagraphId!);
            TestAssert.Equal("00AB0001", reply.CommentMetadata.ParentParagraphId!);
            TestAssert.Equal("43", reply.Id!);
        }

        var distinct = ReadCommentIdentitySource("42", "99", "99", "99");
        var anchor = distinct.Paragraphs.Single().InlineReferences.Single();
        TestAssert.True(distinct.RelatedStories.All(story => story.Id != anchor.Id), "Different numeric IDs cannot acquire a comment owner.");
        var malformed = ReadCommentIdentitySource("4x2", "42", "42", "42");
        TestAssert.Equal("4x2", malformed.RelatedStories.Single(story => story.CommentMetadata?.Author == "Reviewer").Id!);
        var hidden = ReadCommentIdentitySource("42", "+0042", "0042", "42", mode: OoxPdfDocxMarkupMode.Final, hidden: true);
        TestAssert.Equal("42", hidden.PackageCommentAnchorIds.Single());
        TestAssert.Equal("42", hidden.HiddenCommentAnchorIds.Single());
    }

    internal static DocxDocument ReadCommentIdentitySource(string story, string reference, string start, string end,
        OoxPdfDocxMarkupMode mode = OoxPdfDocxMarkupMode.AllMarkup, bool cell = false, bool hidden = false)
    {
        const string word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string office = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const string relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
        string anchor = $"<w:commentRangeStart w:id=\"{start}\"/><w:r><w:t>ANCHOR</w:t></w:r><w:commentRangeEnd w:id=\"{end}\"/><w:r><w:commentReference w:id=\"{reference}\"/></w:r>";
        if (hidden) anchor = "<w:del w:id=\"7\" w:author=\"Reviewer\">" + anchor + "</w:del><w:r><w:t>RETAINED</w:t></w:r>";
        string body = "<w:p>" + anchor + "</w:p>";
        if (cell) body = "<w:tbl><w:tblPr><w:tblW w:w=\"6000\" w:type=\"dxa\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"6000\"/></w:tblGrid><w:tr><w:tc>" + body + "</w:tc></w:tr></w:tbl>";
        using var input = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/><Override PartName=\"/word/comments.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml\"/><Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/></Types>",
            ["_rels/.rels"] = $"<Relationships xmlns=\"{relationships}\"><Relationship Id=\"main\" Type=\"{office}/officeDocument\" Target=\"word/document.xml\"/></Relationships>",
            ["word/_rels/document.xml.rels"] = $"<Relationships xmlns=\"{relationships}\"><Relationship Id=\"comments\" Type=\"{office}/comments\" Target=\"comments.xml\"/><Relationship Id=\"styles\" Type=\"{office}/styles\" Target=\"styles.xml\"/><Relationship Id=\"threads\" Type=\"http://schemas.microsoft.com/office/2011/relationships/commentsExtended\" Target=\"commentsExtended.xml\"/></Relationships>",
            ["word/document.xml"] = $"<w:document xmlns:w=\"{word}\"><w:body>{body}<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:right=\"1440\"/></w:sectPr></w:body></w:document>",
            ["word/styles.xml"] = $"<w:styles xmlns:w=\"{word}\"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii=\"DejaVu Sans\" w:hAnsi=\"DejaVu Sans\"/><w:sz w:val=\"24\"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>",
            ["word/comments.xml"] = $$"""
                <w:comments xmlns:w="{{word}}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
                  <w:comment w:id="{{story}}" w:author="Reviewer"><w:p w14:paraId="00AB0001"><w:r><w:t>REVIEWNOTE</w:t></w:r></w:p></w:comment>
                  <w:comment w:id="+0043" w:author="Reply"><w:p w14:paraId="00AB0002"><w:r><w:t>REPLYBODY</w:t></w:r></w:p></w:comment>
                </w:comments>
                """,
            ["word/commentsExtended.xml"] = "<w15:commentsEx xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\"><w15:commentEx w15:paraId=\"00AB0001\"/><w15:commentEx w15:paraId=\"00AB0002\" w15:paraIdParent=\"00AB0001\"/></w15:commentsEx>",
        });
        using var package = OoxPackage.Open(input, default);
        return new DocxReader().Read(package, null, default, mode);
    }
}
