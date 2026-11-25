using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxMathTests
{
    public static void DocxMathDeletedTextRequiresDeletionOwner()
    {
        foreach (string? wrapper in new[] { null, "ins", "del" })
        {
            var xml = EditingMarkerFraction();
            var run = xml.Descendants(MathNamespace + "r").First();
            var deletedText = run.Element(MathNamespace + "t")!;
            deletedText.Name = WordprocessingNamespace + "delText";
            if (wrapper == "del")
                deletedText.Add(new XElement(WordprocessingNamespace + "br"));
            if (wrapper is not null)
                run.ReplaceWith(new XElement(WordprocessingNamespace + wrapper, new XElement(run)));
            var parsed = DocxMathParser.Read(xml, (_, text) => PlainRun(text), default, _ => true);
            var fraction = (DocxMathFraction)parsed.Children.Single();
            TestAssert.True(fraction.Numerator.Children[0] is DocxMathUnsupported,
                "Deleted literal text needs an actual deletion owner and cannot contain nested elements.");
        }
    }

    public static void DocxMathWordTextPreservesOrderAndRevisionOwnership()
    {
        XName wordText = WordprocessingNamespace + "t";
        var mixed = EditingMarkerFraction();
        mixed.Descendants(MathNamespace + "r").First().Add(
            new XElement(wordText, new XAttribute(XNamespace.Xml + "space", "preserve"), " q "),
            new XElement(MathNamespace + "t", "x"));
        var fraction = (DocxMathFraction)Parse(mixed).Children.Single();
        TestAssert.Equal("a q x+b", RowText(fraction.Numerator));

        foreach (bool included in new[] { false, true })
        {
            var revised = EditingMarkerFraction();
            var run = revised.Descendants(MathNamespace + "r").First();
            run.Element(MathNamespace + "t")!.Name = wordText;
            run.ReplaceWith(new XElement(WordprocessingNamespace + "ins", new XElement(run)));
            var parsed = DocxMathParser.Read(revised, (_, text) => PlainRun(text), default, _ => included);
            var numerator = ((DocxMathFraction)parsed.Children.Single()).Numerator;
            var revision = TestAssert.NotNull(numerator.Children[0] as DocxMathRow);
            TestAssert.Equal(included ? "a" : "", RowText(revision));
        }

        var malformed = EditingMarkerFraction();
        malformed.Descendants(MathNamespace + "r").First().Add(
            new XElement(wordText, new XElement(WordprocessingNamespace + "br")));
        var fallback = (DocxMathFraction)Parse(malformed).Children.Single();
        TestAssert.True(fallback.Numerator.Children[0] is DocxMathUnsupported,
            "Content nested inside a WordprocessingML text element retains the unsupported fallback.");
    }

    public static void DocxMathNonBreakingHyphenPreservesRunOrderAndRevisionOwnership()
    {
        XName hyphen = WordprocessingNamespace + "noBreakHyphen";
        var xml = EditingMarkerFraction();
        xml.Descendants(MathNamespace + "r").First().Add(new XElement(hyphen), new XElement(WordprocessingNamespace + "t", "z"));
        TestAssert.Equal("a\u2011z+b", RowText(((DocxMathFraction)Parse(xml).Children.Single()).Numerator));

        foreach (string owner in new[] { "ins", "del" })
        foreach (bool included in new[] { false, true })
        {
            var revised = EditingMarkerFraction();
            var run = revised.Descendants(MathNamespace + "r").First();
            run.ReplaceNodes(new XElement(hyphen));
            run.ReplaceWith(new XElement(WordprocessingNamespace + owner, new XElement(run)));
            var parsed = DocxMathParser.Read(revised, (_, text) => PlainRun(text), default, _ => included);
            var numerator = ((DocxMathFraction)parsed.Children.Single()).Numerator;
            var revision = TestAssert.NotNull(numerator.Children[0] as DocxMathRow);
            TestAssert.Equal(included ? "\u2011" : "", RowText(revision));
        }

        foreach (var malformed in new[]
        {
            new XElement(hyphen, "Unexpected text"),
            new XElement(hyphen, new XElement(MathNamespace + "t", "x")),
            new XElement(hyphen, new XAttribute("extra", "unknown"))
        })
        {
            var invalid = EditingMarkerFraction();
            invalid.Descendants(MathNamespace + "r").First().Add(malformed);
            var fraction = (DocxMathFraction)Parse(invalid).Children.Single();
            TestAssert.True(fraction.Numerator.Children[0] is DocxMathUnsupported,
                "A nonbreaking hyphen marker must be empty and carry no payload attributes.");
        }
    }

    public static void DocxMathOptionalBreaksKeepOrderedInkAndEmptyRevisionRules()
    {
        XName marker = WordprocessingNamespace + "softHyphen";
        foreach (var breakControl in new[] { new XElement(marker), new XElement(MathNamespace + "t", "\u00AD"),
            new XElement(MathNamespace + "t", "\u200B"), new XElement(WordprocessingNamespace + "t", "\u00AD\u200B"),
            new XElement(MathNamespace + "t", "\u2060"), new XElement(WordprocessingNamespace + "t", "\uFEFF") })
        {
            var equation = EditingMarkerFraction();
            equation.Descendants(MathNamespace + "r").First().Add(breakControl, new XElement(WordprocessingNamespace + "t", "z"));
            TestAssert.Equal("az+b", RowText(((DocxMathFraction)Parse(equation).Children.Single()).Numerator));

            var revised = EditingMarkerFraction();
            var first = revised.Descendants(MathNamespace + "r").First();
            first.ReplaceNodes(new XElement(breakControl));
            first.ReplaceWith(new XElement(WordprocessingNamespace + "ins", new XElement(first)));
            var math = DocxMathParser.Read(revised, (_, text) => PlainRun(text), default, _ => true);
            TestAssert.True(((DocxMathFraction)math.Children.Single()).Numerator.Children[0] is DocxMathUnsupported,
                "A nonprinting break control alone supplies no visible revision operand.");
        }

        var malformed = EditingMarkerFraction();
        malformed.Descendants(MathNamespace + "r").First().Add(new XElement(marker, "Unexpected text"));
        TestAssert.True(((DocxMathFraction)Parse(malformed).Children.Single()).Numerator.Children[0] is DocxMathUnsupported,
            "A content-bearing optional-hyphen marker retains the unsupported-run fallback.");
    }

    public static void DocxMathRevisionAdmissionSharesReadableEmptyRunRules()
    {
        foreach (string kind in new[] { "ins", "del" })
        foreach (bool included in new[] { false, true })
        foreach (string contents in new[] { "", "<m:t/>", "<w:lastRenderedPageBreak/>",
            "<m:rPr><m:nor m:val='0'/></m:rPr>" })
        {
            var xml = EditingMarkerFraction();
            var run = xml.Descendants(MathNamespace + "r").First();
            var empty = XElement.Parse($"<m:r xmlns:m='{MathNamespace}' xmlns:w='{WordprocessingNamespace}'>{contents}</m:r>");
            run.ReplaceWith(new XElement(WordprocessingNamespace + kind, new XElement(run), empty));
            var parsed = DocxMathParser.Read(xml, (_, text) => PlainRun(text), default, _ => included);
            var fraction = (DocxMathFraction)parsed.Children.Single();
            var revision = fraction.Numerator.Children[0] as DocxMathRow;
            TestAssert.NotNull(revision, "A readable empty run does not invalidate its actual revision container.");
            TestAssert.Equal(included ? 2 : 0, revision!.Children.Count);
            if (included) TestAssert.Equal("a", RowText(revision));
            TestAssert.Equal("c", RowText(fraction.Denominator));

            var emptyOnly = EditingMarkerFraction();
            emptyOnly.Descendants(MathNamespace + "r").First().ReplaceWith(
                new XElement(WordprocessingNamespace + kind, new XElement(empty)));
            var unsupported = DocxMathParser.Read(emptyOnly, (_, text) => PlainRun(text), default, _ => included);
            TestAssert.True(((DocxMathFraction)unsupported.Children.Single()).Numerator.Children[0] is DocxMathUnsupported,
                "Empty-only wrappers retain the existing requirement for an actual revision payload.");
        }

        foreach (string kind in new[] { "ins", "del" })
        foreach (bool included in new[] { false, true })
        foreach (string contents in new[] { "<m:future/>", "<w:lastRenderedPageBreak extra='future'/>",
            "<m:rPr><m:future/></m:rPr>", "<m:rPr><m:nor m:val='future'/></m:rPr>" })
        {
            var xml = EditingMarkerFraction();
            var run = xml.Descendants(MathNamespace + "r").First();
            var unknown = XElement.Parse($"<m:r xmlns:m='{MathNamespace}' xmlns:w='{WordprocessingNamespace}'>{contents}</m:r>");
            run.ReplaceWith(new XElement(WordprocessingNamespace + kind, new XElement(run), unknown));
            var parsed = DocxMathParser.Read(xml, (_, text) => PlainRun(text), default, _ => included);
            TestAssert.True(((DocxMathFraction)parsed.Children.Single()).Numerator.Children[0] is DocxMathUnsupported,
                "An unknown empty payload retains its previous revision fallback in either view.");
        }

        foreach (bool included in new[] { false, true })
        {
            var xml = EditingMarkerFraction();
            var run = xml.Descendants(MathNamespace + "r").First();
            run.ReplaceWith(new XElement(WordprocessingNamespace + "del", new XElement(run),
                new XElement(MathNamespace + "r", new XElement(MathNamespace + "t", "Visible"), new XElement(MathNamespace + "future"))));
            var parsed = DocxMathParser.Read(xml, (_, text) => PlainRun(text), default, _ => included);
            var row = (DocxMathRow)((DocxMathFraction)parsed.Children.Single()).Numerator.Children[0];
            TestAssert.True(included ? row.Children.Last() is DocxMathUnsupported : row.Children.Count == 0,
                "Visible unknown content retains its existing deferred fallback and excluded-literal filtering.");
        }
    }

    public static void DocxMathRunNonPrintingMarkersRetainTextAndFallbacks()
    {
        foreach (string marker in new[] { "lastRenderedPageBreak", "proofErr", "permStart", "permEnd" })
        {
            var xml = EditingMarkerFraction();
            foreach (var run in xml.Descendants(MathNamespace + "r"))
            {
                run.AddFirst(new XElement(WordprocessingNamespace + marker));
                run.Add(new XElement(WordprocessingNamespace + marker));
            }
            var fraction = (DocxMathFraction)Parse(xml).Children.Single();
            TestAssert.Equal("a+b", RowText(fraction.Numerator));
            TestAssert.Equal("c", RowText(fraction.Denominator));
        }

        foreach (string marker in new[]
        {
            "<w:lastRenderedPageBreak extra='future'/>",
            "<w:lastRenderedPageBreak>Visible</w:lastRenderedPageBreak>",
            "<w:lastRenderedPageBreak><m:t>Visible</m:t></w:lastRenderedPageBreak>",
            "<x:lastRenderedPageBreak xmlns:x='urn:future'/>"
        })
        {
            var xml = EditingMarkerFraction();
            var content = XElement.Parse($"<root xmlns:w='{WordprocessingNamespace}' xmlns:m='{MathNamespace}'>{marker}</root>");
            xml.Descendants(MathNamespace + "r").First().Add(content.Elements().Single());
            var fraction = (DocxMathFraction)Parse(xml).Children.Single();
            TestAssert.True(fraction.Numerator.Children[0] is DocxMathUnsupported,
                "Unknown or content-bearing run metadata retains the explicit fallback.");
        }

        var misplaced = EditingMarkerFraction();
        misplaced.Descendants(MathNamespace + "num").Single().AddFirst(new XElement(WordprocessingNamespace + "lastRenderedPageBreak"));
        TestAssert.True(((DocxMathFraction)Parse(misplaced).Children.Single()).Numerator.Children[0] is DocxMathUnsupported,
            "A cached run position does not become a valid mathematical argument.");
    }
}
