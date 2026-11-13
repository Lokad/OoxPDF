using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxEmptyRevisionScopeTests
{
    private static readonly (string Fixture, string Part)[] Cases =
    [
        ("CompositionProbes/inline-math-navigation", "word/document.xml"),
        ("CompositionProbes/math-picture-first-table-single-cell", "word/document.xml"),
        ("InlineBoxMath/inline-box-math-header", "word/inline-box-header.xml"),
        ("InlineBoxMath/inline-box-math-footer", "word/inline-box-footer.xml")
    ];

    public static void DocxEmptyRevisionScopesKeepMixedContentAcrossStories()
    {
        foreach (var source in Cases)
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.AllMarkup ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup :
                    OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var baseline = Read(source, mode, null);
            var expected = renderer.RenderPagesAsync(baseline, null, default).ToBlockingEnumerable().ToArray();
            foreach (string kind in new[] { "ins", "del" })
            foreach (string payload in new[] { "empty", "bookmark" })
            {
                var document = Read(source, mode, kind, payload);
                var diagnostics = new List<OoxPdfDiagnostic>();
                var actual = renderer.RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
                AssertPaint(expected, actual, source.Fixture + "/" + kind + "/" + payload + "/" + mode);
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                    "A source-range-only revision cannot discard supported neighboring mathematics.");
            }
        }
    }

    public static void DocxEmptyRevisionScopesKeepSourceRangesThroughTransparentContainers()
    {
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        foreach (string kind in new[] { "ins", "del", "moveFrom", "moveTo" })
        foreach (string carrier in new[] { "direct", "customXml", "sdt" })
        {
            var document = Read(Cases[0], mode, kind, "bookmark", carrier);
            var paragraph = DocxBlockTraversal.EnumerateBodyParagraphs(document).First(p => p.Equations.Count > 0);
            TestAssert.Equal(0, paragraph.Revisions.Count);
            TestAssert.Equal(paragraph.Equations.Count, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
            bool selected = kind is "ins" or "moveTo" ? mode != OoxPdfDocxMarkupMode.Original :
                mode is OoxPdfDocxMarkupMode.Original or OoxPdfDocxMarkupMode.AllMarkup;
            TestAssert.Equal(selected, paragraph.BookmarkAnchors.Any(a => a.Name == "EmptyRevisionScope"));
            TestAssert.True(DocxInlineContent.HasValidHyperlinkSourceRanges(paragraph, default),
                "Annotations keep their captured source positions when the empty scope is removed from paint ownership.");
        }
    }

    public static void DocxRevisionScopesRetainActualUnknownAndMovingOwnerGuards()
    {
        foreach (string kind in new[] { "ins", "del", "moveFrom", "moveTo" })
        foreach (string payload in new[] { "literal", "empty-run", "unknown" })
        {
            var document = Read(Cases[0], OoxPdfDocxMarkupMode.AllMarkup, kind, payload);
            var paragraph = DocxBlockTraversal.EnumerateBodyParagraphs(document).First(p => p.Equations.Count > 0);
            TestAssert.Equal(1, paragraph.Revisions.Count);
            bool supported = (kind is "ins" or "del") && payload != "unknown";
            TestAssert.Equal(supported, DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, default));
            TestAssert.Equal(supported ? paragraph.Equations.Count : 0,
                DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
        }
    }

    public static void DocxHyperlinkScopesShareParsedSourceOwnership()
    {
        foreach (var source in Cases)
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.AllMarkup ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup :
                    OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var expected = renderer.RenderPagesAsync(Read(source, mode, null), null, default).ToBlockingEnumerable().ToArray();
            foreach (string kind in new[] { "ins", "del" })
            foreach (string payload in new[] { "empty", "bookmark" })
            {
                var document = Read(source, mode, kind, payload, "hyperlink");
                var diagnostics = new List<OoxPdfDiagnostic>();
                var actual = renderer.RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
                AssertPaint(expected, actual, source.Fixture + "/hyperlink/" + kind + "/" + payload + "/" + mode);
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                    "A hyperlink with only source annotations cannot suppress neighboring mathematics.");
            }
        }
    }

    public static void DocxHyperlinkScopesRetainActualAndUnknownOwnerGuards()
    {
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        foreach (string kind in new[] { "ins", "del", "moveFrom", "moveTo" })
        foreach (string payload in new[] { "empty", "bookmark", "literal", "empty-run", "unknown" })
        {
            var document = Read(Cases[0], mode, kind, payload, "hyperlink");
            var paragraph = DocxBlockTraversal.EnumerateBodyParagraphs(document).First(p => p.Equations.Count > 0);
            bool included = kind is "ins" or "moveTo" ? mode != OoxPdfDocxMarkupMode.Original :
                mode is OoxPdfDocxMarkupMode.Original or OoxPdfDocxMarkupMode.AllMarkup;
            bool hasOwner = included && payload is not ("empty" or "bookmark");
            TestAssert.Equal(hasOwner ? 1 : 0, paragraph.Revisions.Count);
            bool supported = !hasOwner || (kind is "ins" or "del") && payload != "unknown";
            TestAssert.Equal(supported, DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, default));
            TestAssert.Equal(supported ? paragraph.Equations.Count : 0,
                DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
            TestAssert.Equal(included && payload == "bookmark",
                paragraph.BookmarkAnchors.Any(a => a.Name == "EmptyRevisionScope"));
            TestAssert.True(DocxInlineContent.HasValidHyperlinkSourceRanges(paragraph, default),
                "Source ownership propagation retains captured hyperlink ranges.");
        }
    }

    public static void DocxMetadataPropertiesKeepRevisionScopedContentAcrossStories()
    {
        foreach (var source in Cases)
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.AllMarkup ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup :
                    OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var expected = renderer.RenderPagesAsync(Read(source, mode, null), null, default).ToBlockingEnumerable().ToArray();
            foreach (string kind in new[] { "ins", "del" })
            foreach (string carrier in new[] { "customXml-properties", "smartTag-properties" })
            foreach (string payload in new[] { "empty", "bookmark" })
            {
                var document = Read(source, mode, kind, payload, carrier);
                var diagnostics = new List<OoxPdfDiagnostic>();
                var actual = renderer.RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
                AssertPaint(expected, actual, source.Fixture + "/" + kind + "/" + carrier + "/" + payload + "/" + mode);
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                    "Owned metadata properties cannot turn an annotation-only scope into revised payload.");
            }
        }
    }

    public static void DocxMetadataPropertiesRetainActualUnknownAndMovingOwnerGuards()
    {
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        foreach (string kind in new[] { "ins", "del", "moveFrom", "moveTo" })
        foreach (string carrier in new[] { "customXml-properties", "smartTag-properties",
            "customXml-misplaced-properties", "customXml-foreign-properties" })
        foreach (string payload in new[] { "empty", "bookmark", "literal", "empty-run", "unknown" })
        {
            var document = Read(Cases[0], mode, kind, payload, carrier);
            var paragraph = DocxBlockTraversal.EnumerateBodyParagraphs(document).First(p => p.Equations.Count > 0);
            bool included = kind is "ins" or "moveTo" ? mode != OoxPdfDocxMarkupMode.Original :
                mode is OoxPdfDocxMarkupMode.Original or OoxPdfDocxMarkupMode.AllMarkup;
            bool unknownProperties = carrier is "customXml-misplaced-properties" or "customXml-foreign-properties";
            bool hasOwner = included && (unknownProperties || payload is not ("empty" or "bookmark"));
            TestAssert.Equal(hasOwner ? 1 : 0, paragraph.Revisions.Count);
            // Unknown properties retain an unclaimed owner, but cannot replace
            // or invalidate a known run's actual revision owner.
            bool supported = !hasOwner || (kind is "ins" or "del") && payload is "literal" or "empty-run";
            TestAssert.True(supported == DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, default),
                "Revision ownership changed: " + mode + "/" + kind + "/" + carrier + "/" + payload);
            TestAssert.Equal(supported ? paragraph.Equations.Count : 0,
                DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
            TestAssert.Equal(included && payload == "bookmark",
                paragraph.BookmarkAnchors.Any(a => a.Name == "EmptyRevisionScope"));
        }
    }

    private static DocxDocument Read((string Fixture, string Part) source, OoxPdfDocxMarkupMode mode,
        string? kind, string payload = "empty", string carrier = "direct") =>
        DocxInlineCompositionTests.ReadMutation(source.Fixture, mode, false, false, (xml, w, m) =>
        {
            if (kind is null) return;
            var equation = xml.Descendants(m + "oMath").First();
            var scope = new XElement(w + kind, new XAttribute(w + "id", "903"),
                new XAttribute(w + "author", "Public reviewer"), new XAttribute(w + "date", "2026-11-13T00:00:00Z"));
            XElement[] children = payload switch
            {
                "bookmark" => new[]
                {
                    new XElement(w + "bookmarkStart", new XAttribute(w + "id", "904"), new XAttribute(w + "name", "EmptyRevisionScope")),
                    new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "904"))
                },
                "literal" => [new XElement(w + "r", new XElement(w + "t", "Revised source"))],
                "empty-run" => [new XElement(w + "r")],
                "unknown" => [new XElement(w + "future")],
                _ => Array.Empty<XElement>()
            };
            if (carrier == "hyperlink") scope.Add(new XElement(w + "hyperlink", new XAttribute(w + "anchor", "EmptyRevisionScope"), children));
            else if (carrier == "sdt") scope.Add(new XElement(w + "sdt", new XElement(w + "sdtContent", children)));
            else if (carrier == "customXml") scope.Add(new XElement(w + "customXml", children));
            else if (carrier.EndsWith("-properties", StringComparison.Ordinal))
            {
                string owner = carrier.StartsWith("smartTag", StringComparison.Ordinal) ? "smartTag" : "customXml";
                XName property = carrier == "customXml-foreign-properties" ? XName.Get("customXmlPr", "urn:future") :
                    w + (carrier == "customXml-misplaced-properties" ? "smartTagPr" : owner + "Pr");
                scope.Add(new XElement(w + owner, new XAttribute(w + "uri", "urn:public:metadata"),
                    new XAttribute(w + "element", "calculation"),
                    new XElement(property, new XElement(w + "attr", new XAttribute(w + "name", "public-property"),
                        new XAttribute(w + "val", "metadata-value"))), children));
            }
            else scope.Add(children);
            equation.AddAfterSelf(scope);
        }, partOverride: source.Part);

    private static void AssertPaint(IReadOnlyList<PdfPage> expected, IReadOnlyList<PdfPage> actual, string context)
    {
        TestAssert.Equal(expected.Count, actual.Count);
        for (int index = 0; index < expected.Count; index++)
        {
            TestAssert.True(expected[index].Content == actual[index].Content, "Visible paint changed after an empty source revision: " + context);
            TestAssert.Equal(expected[index].Fonts.Count, actual[index].Fonts.Count);
            TestAssert.Equal(expected[index].Images.Count, actual[index].Images.Count);
            TestAssert.True(expected[index].Annotations.All(a => actual[index].Annotations.Contains(a)),
                "The new empty-scope anchor cannot replace existing links or destinations.");
        }
    }
}
