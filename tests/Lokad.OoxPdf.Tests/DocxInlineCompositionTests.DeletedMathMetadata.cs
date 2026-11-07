using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxDeletedMathMetadataKeepsModeFilteringAndSourceOwners()
    {
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            var direct = ReadDeletedMathMetadata(mode, null).Paragraphs.First();
            TestAssert.Equal(mode == OoxPdfDocxMarkupMode.SimpleMarkup ? 0 : 1, direct.Equations.Count);
            foreach (string wrapper in new[] { "sdt", "sdtContent", "smartTag", "customXml", "nested" })
            {
                var wrapped = ReadDeletedMathMetadata(mode, wrapper).Paragraphs.First();
                TestAssert.Equal(direct.Equations.Count, wrapped.Equations.Count);
                TestAssert.Equal(direct.Revisions.Count, wrapped.Revisions.Count);
                TestAssert.Equal(string.Concat(direct.Runs.Select(run => run.Text)),
                    string.Concat(wrapped.Runs.Select(run => run.Text)));
                if (mode == OoxPdfDocxMarkupMode.SimpleMarkup) continue;
                var equation = wrapped.Equations.Single();
                TestAssert.Equal(direct.Equations.Single().Display, equation.Display);
                TestAssert.Equal(direct.Equations.Single().AlignmentValue, equation.AlignmentValue);
                TestAssert.Equal(direct.Equations.Single().SourceRunIndex, equation.SourceRunIndex);
                TestAssert.Equal(mode == OoxPdfDocxMarkupMode.Final, equation.Revision is null);
                if (mode == OoxPdfDocxMarkupMode.Final)
                    TestAssert.True(equation.Content.Children.Single() is DocxMathFraction
                        { Numerator.Children.Count: 0, Denominator.Children.Count: 0 },
                        "Metadata preserves the excluded equation's structure without deleted glyphs.");
            }
        }
    }

    public static void DocxDeletedMathMetadataKeepsPortablePaintAndNeighborContent()
    {
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.AllMarkup
                    ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                    : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var direct = renderer.RenderPagesAsync(ReadDeletedMathMetadata(mode, null), null, default)
                .ToBlockingEnumerable().ToArray();
            foreach (string wrapper in new[] { "sdt", "smartTag", "nested" })
            {
                var wrapped = renderer.RenderPagesAsync(ReadDeletedMathMetadata(mode, wrapper), null, default)
                    .ToBlockingEnumerable().ToArray();
                TestAssert.Equal(direct.Length, wrapped.Length);
                for (int page = 0; page < direct.Length; page++)
                {
                    TestAssert.Equal(direct[page].Content, wrapped[page].Content);
                    TestAssert.Equal(direct[page].Fonts.Count, wrapped[page].Fonts.Count);
                    TestAssert.Equal(direct[page].Images.Count, wrapped[page].Images.Count);
                    TestAssert.Equal(direct[page].Annotations.Count, wrapped[page].Annotations.Count);
                }
            }
        }
    }

    public static void DocxDeletedMathMetadataDoesNotReviveOtherDeletedContent()
    {
        var source = ReadDeletedMathMetadata(OoxPdfDocxMarkupMode.Final, "nested", (xml, w, m) =>
        {
            var content = xml.Descendants(w + "sdtContent").Single();
            content.Add(new XElement(w + "r", new XElement(w + "delText", "DeletedPayload")),
                new XElement(w + "hyperlink", new XAttribute(w + "anchor", "DeletedDestination"),
                    new XElement(w + "r", new XElement(w + "delText", "DeletedLink"))),
                new XElement(w + "bookmarkStart", new XAttribute(w + "id", "881"),
                    new XAttribute(w + "name", "DeletedDestination")));
        }).Paragraphs.First();
        var direct = ReadDeletedMathMetadata(OoxPdfDocxMarkupMode.Final, null).Paragraphs.First();
        TestAssert.Equal(string.Concat(direct.Runs.Select(run => run.Text)),
            string.Concat(source.Runs.Select(run => run.Text)));
        TestAssert.Equal(direct.Hyperlinks.Count, source.Hyperlinks.Count);
        TestAssert.Equal(direct.BookmarkAnchors.Count, source.BookmarkAnchors.Count);
        TestAssert.Equal(1, source.Equations.Count);
        foreach (string owner in new[] { "unknown", "ins", "del", "bdo" })
        {
            var unsupported = ReadDeletedMathMetadata(OoxPdfDocxMarkupMode.Final, "sdt", (xml, w, m) =>
            {
                var content = xml.Descendants(w + "sdtContent").Single();
                var equation = content.Element(m + "oMathPara")!;
                equation.ReplaceWith(new XElement(w + owner, new XElement(equation)));
            });
            TestAssert.Equal(0, unsupported.Paragraphs.First().Equations.Count);
        }
    }

    private static DocxDocument ReadDeletedMathMetadata(OoxPdfDocxMarkupMode mode, string? wrapper,
        Action<XDocument, XNamespace, XNamespace>? mutate = null) =>
        ReadMutation("CompositionProbes/equation-revision-body-del", mode, false, false, (xml, w, m) =>
        {
            var deletion = xml.Descendants(w + "del").Single();
            var equation = deletion.Element(m + "oMathPara")!;
            if (wrapper is not null)
            {
                XElement wrapped = wrapper switch
                {
                    "sdt" => new XElement(w + "sdt", new XElement(w + "sdtContent", new XElement(equation))),
                    "nested" => new XElement(w + "sdt", new XElement(w + "sdtContent",
                        MetadataWrapper(w, "smartTag", new XElement(equation)))),
                    "sdtContent" => new XElement(w + "sdtContent", new XElement(equation)),
                    _ => MetadataWrapper(w, wrapper, new XElement(equation))
                };
                equation.ReplaceWith(wrapped);
            }
            mutate?.Invoke(xml, w, m);
        });
}
