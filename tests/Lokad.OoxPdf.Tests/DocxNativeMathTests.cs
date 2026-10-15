using System.Text;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxNativeMathTests
{
    public static void DocxNativeFractionDelimitersRetainChildRulesAndSourceMeaning()
    {
        RequireMathFont();
        foreach (string mode in new[] { "display", "inline" })
        foreach (string suffix in new[] { "", "-brackets", "-stack", "-child20", "-parent20" })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = Renderer().RenderBlankPages(NamedFixture($"docx-omml-delimiter-{mode}-fraction{suffix}"), diagnostics.Add, CancellationToken.None).Single();
            TestAssert.Contains(suffix == "-brackets" ? "[n+1k]" : "(n+1k)", DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_EQUATION") && diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Bounded fractions/stacks retain readable grouped native source instead of disappearing with the enclosing delimiter.");
        }
        var composed = Renderer().RenderBlankPages(NamedFixture("docx-omml-function-inline-sin-delimiter-fraction"), null, CancellationToken.None).Single();
        TestAssert.Contains("sin(n+1k)", DocxPageFootnoteNumberingTests.DecodeText(composed).Normalize(NormalizationForm.FormKC));
    }

    public static void DocxNativeMixedFunctionsKeepContextAndAdjacentSourceMeaning()
    {
        RequireMathFont();
        foreach ((string suffix, string meaning) in new[] {
            ("display-sin-literal-binary", "a+sinx=b"), ("display-sin-literal-ordinary", "asinx=b"),
            ("inline-sin-delimiter-binary", "a+sin(n+1)=b"), ("inline-sin-delimiter-ordinary", "asin(n+1)=b"),
            ("display-sin-fraction-binary", "a+sinxk=b"), ("inline-log-normal-both-ordinary", "alogx=b"),
            ("inline-cos-control20-ordinary", "acosx=b"),
            ("inline-sin-delimiter-ordinary-then-cos-control20", "asin(n+1)cosx=b"),
            ("inline-cos-control20-ordinary-then-sin-delimiter", "acosxsin(n+1)=b"),
            ("display-sin-literal-ordinary-parent20", "asinx=b"),
            ("inline-sin-delimiter-ordinary-parent20", "asin(n+1)=b"),
            ("inline-cos-control20-ordinary-parent16", "acosx=b") })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = Renderer().RenderBlankPages(NamedFixture("docx-omml-function-row-" + suffix), diagnostics.Add, CancellationToken.None).Single();
            TestAssert.Contains(meaning, DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_EQUATION") && diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Mixed functions retain prefix, argument, adjacent controls and outgoing relation source in native paint.");
        }
    }

    public static void DocxNativeFunctionsRetainNamesAndBoundedArgumentMeaning()
    {
        RequireMathFont();
        foreach ((string suffix, string meaning) in new[] { ("display-sin-literal", "sinx"), ("inline-sin-delimiter", "sin(n+1)"),
            ("display-log-script", "logx2"), ("inline-cos-control20", "cosx"), ("display-sin-fraction", "sinxk"),
            ("inline-log-normal-name", "logx"), ("inline-log-normal-both", "logx") })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = Renderer().RenderBlankPages(NamedFixture("docx-omml-function-" + suffix), diagnostics.Add, CancellationToken.None).Single();
            TestAssert.Contains(meaning, DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_EQUATION") && diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Known function names and bounded arguments retain actual native glyph/rule painting.");
            if (suffix == "inline-log-normal-both")
            {
                var fonts = page.Fonts.Where(f => f.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).ToArray();
                TestAssert.True(fonts.Length == 1 && fonts[0].Font.Font.FamilyName.Contains("Arial", StringComparison.OrdinalIgnoreCase),
                    "A metric-only Cambria owner cannot publish an unused whole-font resource when every visible glyph uses Arial.");
            }
        }
    }

    public static void DocxNativeMixedNaryRowsKeepPrefixAndAdjacentControlMeaning()
    {
        RequireMathFont();
        foreach ((string suffix, string meaning) in new[] {
            ("display-sum-binary", "a+\u2211xni=1=b"), ("display-product-binary", "a+\u220Fxni=1=b"),
            ("inline-sum-binary", "a+\u2211xni=1=b"), ("inline-product-binary", "a+\u220Fxni=1=b"), ("inline-integral-binary", "a+\u222Bx10=b"),
            ("display-sum-ordinary", "a\u2211xni=1=b"), ("inline-sum-ordinary", "a\u2211xni=1=b"),
            ("display-adjacent-product", "a\u2211xni=1\u220Fxni=1=b"), ("inline-adjacent-integral", "a\u2211xni=1\u222Bx10=b") })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = Renderer().RenderBlankPages(NamedFixture("docx-omml-nary-row-" + suffix), diagnostics.Add, CancellationToken.None).Single();
            TestAssert.Contains(meaning, DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_EQUATION") && diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Mixed rows retain all literal, body/limit and per-node operator source meaning.");
        }
    }

    public static void DocxNativeNaryGrowthPreservesOffAndFitsExplicitTallBodies()
    {
        RequireMathFont();
        foreach ((string suffix, string character, string body, string limits) in new[] {
            ("display-sum-off", "\u2211", "x", "ni=1"), ("display-product-off", "\u220F", "x", "ni=1"),
            ("inline-sum-off", "\u2211", "x", "ni=1"), ("inline-product-off", "\u220F", "x", "ni=1"), ("inline-integral-off", "\u222B", "x", "10"),
            ("display-sum-tall-off", "\u2211", "g+y", "ni=1"), ("display-sum-tall-on", "\u2211", "g+y", "ni=1"),
            ("display-sum-body24-on", "\u2211", "g+y", "ni=1"), ("display-sum-body48-on", "\u2211", "g+y", "ni=1"),
            ("display-product-body24-on", "\u220F", "g+y", "ni=1"), ("display-product-body48-on", "\u220F", "g+y", "ni=1"),
            ("inline-sum-body36-on", "\u2211", "g+y", "ni=1"), ("inline-integral-body36-on", "\u222B", "g+y", "10") })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = Renderer().RenderBlankPages(NamedFixture("docx-omml-nary-growth-" + suffix), diagnostics.Add, CancellationToken.None).Single();
            TestAssert.Contains(character + body + limits, DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_EQUATION") && diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Known growth-off/default and explicit body-fit roles retain native operator, body and limits.");
        }
    }

    public static void DocxNativeMixedDelimiterRowsKeepLiteralMeaningAndMultipleControls()
    {
        RequireMathFont();
        foreach (string mode in new[] { "display", "inline" })
        foreach ((string kind, string meaning) in new[] { ("single", "a+(n+1)=b"), ("multiple", "a+(n+1|k+2)=b"), ("mixed-growth", "a+(n+1)+(k+2)=b") })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = Renderer().RenderBlankPages(NamedFixture($"docx-omml-delimiter-row-{mode}-{kind}"), diagnostics.Add, CancellationToken.None).Single();
            TestAssert.Contains(meaning, DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_EQUATION") && diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Supported mixed rows retain every literal, child argument and implicit delimiter control.");
        }
    }

    public static void DocxNativeInlineIntegralsRetainLiteralSizedAndStructuredBodyMeaning()
    {
        RequireMathFont();
        foreach ((string name, string meaning) in new[] { ("docx-omml-nary-inline-integral-subSup", "x"),
            ("docx-omml-integral-inline-size8", "x"), ("docx-omml-integral-inline-size16", "x"), ("docx-omml-integral-inline-size20", "x"),
            ("docx-omml-integral-inline-stacked-request", "x"), ("docx-omml-integral-inline-document-default", "x"),
            ("docx-omml-integral-inline-fraction-body", "xk"), ("docx-omml-integral-inline-script-body", "xi2") })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = Renderer().RenderBlankPages(NamedFixture(name), diagnostics.Add, CancellationToken.None).Single();
            TestAssert.Contains("\u222B" + meaning + "10", DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_EQUATION") && diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Qualified inline integrals retain their native operator, body and both limits across font sizes and location requests.");
        }
    }

    public static void DocxNativeNaryStructuredBodiesRetainFractionStackAndScriptMeaning()
    {
        RequireMathFont();
        foreach (string mode in new[] { "display", "inline" })
        foreach ((string kind, string body, string character, string meaning) in new[] {
            ("sum", "fraction", "\u2211", "xk"), ("product", "stack", "\u220F", "xk"), ("sum", "script", "\u2211", "xi2") })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = Renderer().RenderBlankPages(NamedFixture($"docx-omml-nary-body-{mode}-{kind}-{body}"), diagnostics.Add, CancellationToken.None).Single();
            TestAssert.Contains(character + meaning + "ni=1", DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_EQUATION") && diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Bounded body boxes retain native operator, inner mathematical meaning and both outer limits.");
        }
    }

    public static void DocxNativeNaryDocumentDefaultsKeepSharedOwnerAndExplicitPrecedence()
    {
        RequireMathFont();
        foreach ((string suffix, string character, string? location) in new[] {
            ("display-sum-omitted", "\u2211", (string?)null), ("display-product-setting-side", "\u220F", null),
            ("display-sum-setting-stacked", "\u2211", null), ("inline-product-setting-stacked", "\u220F", null),
            ("display-sum-explicit-side", "\u2211", "subSup"), ("display-product-present-default", "\u220F", "undOvr"),
            ("display-product-setting-default", "\u220F", null), ("inline-sum-setting-side", "\u2211", null) })
        {
            var source = NamedFixture("docx-omml-nary-default-" + suffix);
            var equation = source.Paragraphs.SelectMany(p => p.Equations).Single();
            var nary = (DocxMathNary)equation.Content.Children.Single();
            TestAssert.True(ReferenceEquals(source.Settings.MathLimitSettings, equation.LimitSettings), "Equations share the document's immutable settings owner.");
            TestAssert.Equal(location, nary.LimitLocationValue);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = Renderer().RenderBlankPages(source, diagnostics.Add, CancellationToken.None).Single();
            TestAssert.Contains(character + "xni=1", DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_EQUATION") && diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Document defaults and authored node precedence retain native body and limit painting.");
        }
    }

    public static void DocxNativeEmptyHiddenNaryLimitsRetainRemainingBodyAndArguments()
    {
        RequireMathFont();
        foreach (string mode in new[] { "display", "inline" })
        foreach ((string kind, string character, string hidden, string meaning) in new[] {
            ("sum", "\u2211", "lower", "xn"), ("product", "\u220F", "upper", "xi=1"), ("sum", "\u2211", "both", "x") })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = Renderer().RenderBlankPages(NamedFixture($"docx-omml-nary-hidden-empty-{mode}-{kind}-{hidden}"), diagnostics.Add, CancellationToken.None).Single();
            TestAssert.Contains(character + meaning, DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_EQUATION") && diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Empty hidden arguments retain native operator, body and visible limit painting.");
        }
    }

    public static void DocxNativeStoredHiddenNaryArgumentsRetainCompleteSourceFallback()
    {
        RequireMathFont();
        foreach (string mode in new[] { "display", "inline" })
        foreach (string kind in new[] { "sum-lower", "product-upper", "sum-both" })
        {
            var source = NamedFixture($"docx-omml-nary-hidden-{mode}-{kind}");
            var equation = source.Paragraphs.SelectMany(p => p.Equations).Single();
            var nary = (DocxMathNary)equation.Content.Children.Single();
            TestAssert.True(nary.Subscript.Children.Count > 0 && nary.Superscript.Children.Count > 0 &&
                !DocxMathLayout.TryGetRuns(equation, CancellationToken.None, out _), "Stored hidden arguments stay in typed source without claiming omission is correct.");
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = Renderer().RenderBlankPages(source, diagnostics.Add, CancellationToken.None).ToArray();
            TestAssert.True(pages.All(page => page.Fonts.All(font => !font.ResourceName.StartsWith("Eq", StringComparison.Ordinal))),
                "A stored hidden argument cannot produce partial equation font resources.");
            TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_EQUATION") && diagnostics.All(d => d.Id != "DOCX_APPROXIMATED_EQUATION"),
                "Word's reassigned stored text needs its own profile; no partial native painting is published.");
        }
    }

    public static void DocxNativeSideNaryOperatorsKeepExplicitAndInlineLimitPlacement()
    {
        RequireMathFont();
        foreach ((string suffix, string character) in new[] { ("display-sum-subSup", "\u2211"), ("inline-sum-subSup", "\u2211"),
            ("inline-sum-undOvr", "\u2211"), ("inline-product-undOvr", "\u220F"),
            ("display-sum-subSup-control20", "\u2211"), ("inline-product-undOvr-body16", "\u220F") })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = Renderer().RenderBlankPages(NamedFixture("docx-omml-nary-" + suffix), diagnostics.Add, CancellationToken.None).Single();
            TestAssert.Contains(character + "xni=1", DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_EQUATION") && diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Qualified side limits retain native operator, body and both script arguments.");
        }
    }

    public static void DocxNativeStackedNaryOperatorsKeepBodyAndLimitMeaning()
    {
        RequireMathFont();
        foreach ((string kind, string character, string suffix) in new[] { ("sum", "\u2211", ""), ("product", "\u220F", ""),
            ("sum", "\u2211", "-control20"), ("product", "\u220F", "-body16") })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = Renderer().RenderBlankPages(NamedFixture($"docx-omml-nary-display-{kind}-undOvr{suffix}"), diagnostics.Add, CancellationToken.None).Single();
            TestAssert.Contains(character + "xni=1", DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_EQUATION") && diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Stacked sum/product operators retain native body and both limits.");
        }
    }

    public static void DocxNativeDelimiterArgumentsKeepVisibleOrderAndEmptySeparators()
    {
        RequireMathFont();
        foreach(string placement in new[]{"display","inline"})
        foreach(string kind in new[]{"default","empty","large20"})
        {
            var diagnostics=new List<OoxPdfDiagnostic>();
            var page=Renderer().RenderBlankPages(NamedFixture($"docx-omml-delimiter-arguments-{placement}-{kind}"),diagnostics.Add,CancellationToken.None).Single();
            TestAssert.Contains(kind=="empty"?"(n+1k+2)":"(n+1|k+2)",DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.True(diagnostics.Any(d=>d.Id=="DOCX_APPROXIMATED_EQUATION") && diagnostics.All(d=>d.Id!="DOCX_UNSUPPORTED_EQUATION"),
                "Supported argument lists retain actual native control and child glyph painting.");
        }
    }

    public static void DocxNativeFixedDelimitersRetainAuthoredGrowthAndChildSizes()
    {
        RequireMathFont();
        foreach(string placement in new[]{"display","inline"})
        foreach((string kind,double childSize,double controlSize) in new[]{("literal12",12d,12d),("literal20",20.04d,12d),("control20",12d,20.04d)})
        {
            var diagnostics=new List<OoxPdfDiagnostic>();
            var page=Renderer().RenderBlankPages(NamedFixture($"docx-omml-delimiter-fixed-{placement}-{kind}"),diagnostics.Add,CancellationToken.None).Single();
            TestAssert.Contains("(n+1)",DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            foreach(double size in new[]{childSize,controlSize})TestAssert.Contains(" "+size.ToString("0.##",System.Globalization.CultureInfo.InvariantCulture)+" Tf",page.Content);
            TestAssert.True(diagnostics.Any(d=>d.Id=="DOCX_APPROXIMATED_EQUATION") && diagnostics.All(d=>d.Id!="DOCX_UNSUPPORTED_EQUATION"),
                "An authored growth-off flag must retain actual native delimiter painting.");
        }
    }

    public static void DocxNativeDelimiterHeightContrastsKeepParentFramesAndChildMeaning()
    {
        RequireMathFont();
        foreach(string placement in new[]{"display","inline"})
        foreach((string kind,string meaning,double size) in new[]{("literal16","(n+1)",15.96d),("literal20","(n+1)",20.04d),("descender12","(g+y)",12d)})
        {
            var diagnostics=new List<OoxPdfDiagnostic>();
            var page=Renderer().RenderBlankPages(NamedFixture($"docx-omml-delimiter-height-{placement}-{kind}"),diagnostics.Add,CancellationToken.None).Single();
            TestAssert.Contains(meaning,DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.Contains(" "+size.ToString("0.##",System.Globalization.CultureInfo.InvariantCulture)+" Tf",page.Content);
            TestAssert.True(diagnostics.Any(d=>d.Id=="DOCX_APPROXIMATED_EQUATION") && diagnostics.All(d=>d.Id!="DOCX_UNSUPPORTED_EQUATION"),
                "Measured height/descender contrasts retain native delimiter painting and authored child sizes.");
            if(placement=="display" && kind!="descender12")
            {
                var font=page.Fonts.Single(f=>f.ResourceName.StartsWith("Eq",StringComparison.Ordinal));
                TestAssert.True(font.Font.TryGetEncodedCid(font.Font.Font.MapCodePoint(0x1D45B),out ushort cid),"Expected the authored n glyph.");
                var operation=System.Text.RegularExpressions.Regex.Match(page.Content,
                    @"1 0 0 1 [-\d.]+ ([-\d.]+) Tm\s*<"+cid.ToString("X4",System.Globalization.CultureInfo.InvariantCulture)+@"> Tj");
                var frame=System.Text.RegularExpressions.Regex.Match(page.Content,@"q\s*1 0 0 1 [-\d.]+ ([-\d.]+) cm");
                TestAssert.True(operation.Success && frame.Success,"Expected the stored math glyph and frame baseline.");
                double baseline=double.Parse(operation.Groups[1].Value,System.Globalization.CultureInfo.InvariantCulture)+
                    double.Parse(frame.Groups[1].Value,System.Globalization.CultureInfo.InvariantCulture);
                TestAssert.True(Math.Abs(baseline-(kind=="literal16"?707.5d:704.74d))<0.1d,
                    "The independently measured Word baseline must not include a larger child's typographic line gap.");
            }
        }
    }

    public static void DocxNativeLiteralDelimitersKeepVisibleCharactersAndEmptySides()
    {
        RequireMathFont();
        foreach(string placement in new[]{"display","inline"})
        foreach((string kind,string meaning) in new[]{("parentheses","(n+1)"),("brackets","[n+1]"),("one-sided","{n+1")})
        {
            var diagnostics=new List<OoxPdfDiagnostic>();
            var page=Renderer().RenderBlankPages(NamedFixture($"docx-omml-delimiter-{placement}-{kind}"),diagnostics.Add,CancellationToken.None).Single();
            TestAssert.Contains(meaning,DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.True(diagnostics.Any(d=>d.Id=="DOCX_APPROXIMATED_EQUATION") && diagnostics.All(d=>d.Id!="DOCX_UNSUPPORTED_EQUATION"),
                "Qualified literal delimiters retain actual native glyph painting.");
            TestAssert.Equal(0,System.Text.RegularExpressions.Regex.Matches(page.Content," re f").Count);
        }
    }

    public static void DocxNativeRootScriptsRetainMeaningCrampedSizesAndOwnedBars()
    {
        RequireMathFont();
        foreach(string placement in new[]{"display","inline"})
        foreach((string tag,string meaning) in new[]{("sub","xi"),("sup","x2"),("paired","xi2"),
            ("paired-cube","n+1i23"),("descender","xg"),("large20","xi")})
        {
            var diagnostics=new List<OoxPdfDiagnostic>();
            var page=Renderer().RenderBlankPages(NamedFixture($"docx-omml-root-script-{placement}-{tag}"),diagnostics.Add,CancellationToken.None).Single();
            TestAssert.Contains("√"+meaning,DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.Equal(1,System.Text.RegularExpressions.Regex.Matches(page.Content," re f").Count);
            TestAssert.Contains(tag=="large20"?" 14.52 Tf":" 8.52 Tf",page.Content);
            TestAssert.True(diagnostics.Any(d=>d.Id=="DOCX_APPROXIMATED_EQUATION") && diagnostics.All(d=>d.Id!="DOCX_UNSUPPORTED_EQUATION"),
                "Bounded scripted radicands retain native glyph/script/rule painting.");
        }
    }

    public static void DocxNativeRootFractionsKeepMeaningAndIndependentChildRules()
    {
        RequireMathFont();
        foreach(string placement in new[]{"display","inline"})
        foreach(string kind in new[]{"bar","noBar"})
        foreach(string degree in new[]{"square","cube"})
        {
            var diagnostics=new List<OoxPdfDiagnostic>();
            var page=Renderer().RenderBlankPages(NamedFixture($"docx-omml-root-fraction-{placement}-{kind}-{degree}"),diagnostics.Add,CancellationToken.None).Single();
            TestAssert.Contains("√n+1k"+(degree=="cube"?"3":""),DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.Equal(kind=="bar"?2:1,System.Text.RegularExpressions.Regex.Matches(page.Content," re f").Count);
            TestAssert.True(diagnostics.Any(d=>d.Id=="DOCX_APPROXIMATED_EQUATION") && diagnostics.All(d=>d.Id!="DOCX_UNSUPPORTED_EQUATION"),
                "A qualified child fraction/stack retains native root painting and its own rule count.");
        }
    }

    public static void DocxNativeRootDegreesPaintLiteralMeaningAndScriptScriptSize()
    {
        RequireMathFont();
        foreach(string placement in new[]{"display","inline"})
        foreach((string tag,string degree) in new[]{("cube","3"),("letter","n"),("expression","n+1"),("large-control","3")})
        {
            var diagnostics=new List<OoxPdfDiagnostic>();
            PdfPage page=Renderer().RenderBlankPages(NamedFixture($"docx-omml-root-degree-{placement}-{tag}"),diagnostics.Add,CancellationToken.None).Single();
            TestAssert.Contains("√n+1"+degree,DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.Contains(" 6.96 Tf",page.Content);
            TestAssert.Equal(1,System.Text.RegularExpressions.Regex.Matches(page.Content," re f").Count);
            TestAssert.True(diagnostics.Any(d=>d.Id=="DOCX_APPROXIMATED_EQUATION") && diagnostics.All(d=>d.Id!="DOCX_UNSUPPORTED_EQUATION"),
                "A visible degree must be painted along with the owned root and bar.");
        }
    }

    public static void DocxNativeAuthoredRootControlsKeepMeaningSizeAndColor()
    {
        RequireMathFont();
        foreach(string placement in new[]{"display","inline"})
        foreach((string tag,double size) in new[]{("neutral",12d),("small8",8.04d),("large20",20.04d),("color-font",12d),("italic",12d)})
        {
            var diagnostics=new List<OoxPdfDiagnostic>();
            PdfPage page=Renderer().RenderBlankPages(NamedFixture($"docx-omml-root-control-{placement}-{tag}"),diagnostics.Add,CancellationToken.None).Single();
            TestAssert.Contains("√n+1",DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.Contains(" "+size.ToString("0.##",System.Globalization.CultureInfo.InvariantCulture)+" Tf",page.Content);
            int rule=page.Content.IndexOf(" re f",StringComparison.Ordinal);TestAssert.True(rule>=0,"Expected the authored control's root rule.");
            if(tag=="color-font")
            {
                var colors=System.Text.RegularExpressions.Regex.Matches(page.Content[..rule],@"([-\d.]+) ([-\d.]+) ([-\d.]+) rg");
                TestAssert.True(colors.Count>0 && Math.Abs(double.Parse(colors[^1].Groups[2].Value,System.Globalization.CultureInfo.InvariantCulture)-128d/255d)<0.001d,
                    "Authored control color applies to the rule independently of child text.");
            }
            TestAssert.True(diagnostics.Any(d=>d.Id=="DOCX_APPROXIMATED_EQUATION") && diagnostics.All(d=>d.Id!="DOCX_UNSUPPORTED_EQUATION"),
                "Supported authored controls retain actual native painting.");
        }
    }

    public static void DocxNativeRootControlsRetainInheritedColorAndIndependentChildFaces()
    {
        RequireMathFont();
        foreach(string placement in new[]{"display","inline"})
        foreach((string suffix,string meaning,string family) in new[]{("math-color","n+1","Cambria"),
            ("normal-default","rate","Cambria"),("normal-calibri-color","rate","Calibri"),("normal-arial-color","rate","Arial")})
        {
            var diagnostics=new List<OoxPdfDiagnostic>();
            PdfPage page=Renderer().RenderBlankPages(NamedFixture($"docx-omml-root-owner-{placement}-{suffix}"),diagnostics.Add,CancellationToken.None).Single();
            TestAssert.Contains("√"+meaning,DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.True(page.Fonts.Any(f=>f.ResourceName.StartsWith("Eq",StringComparison.Ordinal) &&
                f.Font.Font.FamilyName.Contains(family,StringComparison.OrdinalIgnoreCase)),"The child retains its authored face.");
            int rule=page.Content.IndexOf(" re f",StringComparison.Ordinal);TestAssert.True(rule>=0,"Expected the owned root rule.");
            var colors=System.Text.RegularExpressions.Regex.Matches(page.Content[..rule],@"([-\d.]+) ([-\d.]+) ([-\d.]+) rg");
            TestAssert.True(colors.Count==0 || Enumerable.Range(1,3).All(i=>double.Parse(colors[^1].Groups[i].Value,
                System.Globalization.CultureInfo.InvariantCulture)==0d),"Root control color must remain inherited black independently of colored children.");
            TestAssert.True(diagnostics.Any(d=>d.Id=="DOCX_APPROXIMATED_EQUATION") && diagnostics.All(d=>d.Id!="DOCX_UNSUPPORTED_EQUATION"),
                "Independently owned normal-text children must retain actual native root painting.");
        }
    }

    public static void DocxNativeRadicalsPaintReadableRootsOwnedBarsAndChildSizes()
    {
        RequireMathFont();
        foreach(string placement in new[]{"display","inline"})
        foreach((string shape,string meaning,double size) in new[]{("small8","x",8.04d),("operators12","n+1",12d),
            ("descenders12","g+y",12d),("large20","a+b",20.04d)})
        {
            var diagnostics=new List<OoxPdfDiagnostic>();
            PdfPage page=Renderer().RenderBlankPages(NamedFixture($"docx-omml-radical-{placement}-{shape}"),diagnostics.Add,CancellationToken.None).Single();
            TestAssert.Contains("√"+meaning,DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.Contains(" re f",page.Content);
            TestAssert.Contains(" "+size.ToString("0.##",System.Globalization.CultureInfo.InvariantCulture)+" Tf",page.Content);
            TestAssert.True(diagnostics.Any(d=>d.Id=="DOCX_APPROXIMATED_EQUATION") && diagnostics.All(d=>d.Id!="DOCX_UNSUPPORTED_EQUATION"),
                "Qualified root boxes must report actual native source painting.");
        }
    }

    public static void DocxNativeStructuredSequencesRetainSourceOrderAndIndependentRules()
    {
        RequireMathFont();
        foreach (string placement in new[] { "display", "inline" })
        foreach ((string pair, string meaning, int rules) in new[]
        {
            ("fraction-script", "y=n+1k+x2+z", 1), ("script-fraction", "y=x2+n+1k+z", 1),
            ("fraction-fraction", "y=n+1k+n+1k+z", 2), ("stack-script", "y=n+1k+x2+z", 0)
        })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            PdfPage page = Renderer().RenderBlankPages(NamedFixture($"docx-omml-sequence-{placement}-{pair}"),
                diagnostics.Add, CancellationToken.None).Single();
            TestAssert.Contains(meaning, DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.Equal(rules, System.Text.RegularExpressions.Regex.Matches(page.Content, " re f").Count);
            TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_EQUATION") &&
                diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Sequence diagnostics must describe the actual painted structures.");
        }
    }

    public static void DocxNativeMixedScriptRowsPaintLiteralAndScriptMeaning()
    {
        RequireMathFont();
        foreach (string placement in new[] { "display", "inline" })
        foreach (string suffix in new[] { "sub-prefix", "sub-both", "sup-prefix", "sup-both", "sup-size20" })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            PdfPage page = Renderer().RenderBlankPages(NamedFixture($"docx-omml-mixed-script-{placement}-{suffix}"),
                diagnostics.Add, CancellationToken.None).Single();
            TestAssert.Contains(suffix.EndsWith("prefix", StringComparison.Ordinal) ? "y=x2" : "y=x2+z",
                DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_EQUATION") &&
                diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Mixed script rows must report actual native source painting.");
        }
    }

    public static void DocxNativeMixedFractionRowsPaintLiteralAtomsAndRetainAuthoredRules()
    {
        RequireMathFont();
        foreach (string placement in new[] { "display", "inline" })
        foreach (string kind in new[] { "bar", "stack" })
        foreach (string tail in new[] { "prefix", "both" })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            PdfPage page = Renderer().RenderBlankPages(NamedFixture($"docx-omml-mixed-fraction-{placement}-{kind}-{tail}"),
                diagnostics.Add, CancellationToken.None).Single();
            TestAssert.Contains(tail == "both" ? "y=n+1k+z" : "y=n+1k",
                DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.Equal(kind == "bar", page.Content.Contains(" re f", StringComparison.Ordinal));
            TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_EQUATION") &&
                diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Compound source rows report their actual native painting.");
        }
    }

    public static void DocxNativeFractionScriptBasesRetainNestedMeaningAndAuthoredRules()
    {
        RequireMathFont();
        foreach (string placement in new[] { "display", "inline" })
        foreach (string kind in new[] { "bar", "stack" })
        foreach (string role in new[] { "sub", "sup" })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            PdfPage page = Renderer().RenderBlankPages(NamedFixture($"docx-omml-nested-base-{placement}-{kind}-{role}"),
                diagnostics.Add, CancellationToken.None).Single();
            TestAssert.Contains("n+1k2", DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.Equal(kind == "bar", page.Content.Contains(" re f", StringComparison.Ordinal));
            TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_EQUATION") &&
                diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Nested source diagnostics must describe the actual painted structure.");
        }
    }

    public static void DocxNativeNoBarStacksPaintTheirSourceMeaningAndOwnChildFaces()
    {
        RequireMathFont();
        foreach (string placement in new[] { "display", "inline" })
        foreach (string kind in new[] { "math12", "math20", "normal12" })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            PdfPage page = Renderer().RenderBlankPages(NamedFixture($"docx-omml-stack-{placement}-{kind}"),
                diagnostics.Add, CancellationToken.None).Single();
            TestAssert.Contains("n+1k", DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.True(!page.Content.Contains(" re f", StringComparison.Ordinal), "An authored noBar stack must not paint a fraction rule.");
            PdfFontResource child = page.Fonts.First(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal));
            TestAssert.True(child.Font.Font.FamilyName.Contains(kind == "normal12" ? "Calibri" : "Cambria", StringComparison.OrdinalIgnoreCase),
                "A stack retains its source glyph face independently of the parent metric owner.");
            TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_EQUATION") &&
                diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Diagnostics must describe the actual painted stack.");
        }
    }

    public static void DocxNativeNoBarDisplayStackRetainsWordBaselineSeparation()
    {
        RequireMathFont();
        PdfPage page = Renderer().RenderBlankPages(NamedFixture("docx-omml-stack-display-math12"), null, CancellationToken.None).Single();
        PdfFontResource font = page.Fonts.Single(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal));
        double Baseline(int codePoint)
        {
            TestAssert.True(font.Font.TryGetEncodedCid(font.Font.Font.MapCodePoint(codePoint), out ushort cid), "Expected an owned source glyph.");
            var operation = System.Text.RegularExpressions.Regex.Match(page.Content,
                @"1 0 0 1 [-\d.]+ ([-\d.]+) Tm\s*<" + cid.ToString("X4", System.Globalization.CultureInfo.InvariantCulture) + @"> Tj");
            TestAssert.True(operation.Success, "Expected the stored child baseline.");
            return double.Parse(operation.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
        }
        TestAssert.True(Math.Abs(Baseline(0x1D45B) - Baseline(0x1D458) - 17.16d) < 0.15d,
            "The independent Word stack keeps n/k baselines 17.16pt apart.");
    }

    public static void DocxNativeSlantedScriptCenteringUsesLogicalAdvanceAndRetainsInk()
    {
        RequireMathFont();
        foreach ((string suffix, double wordX) in new[] { ("sub-default-i", 295.01d), ("sup-default-i", 294.65d) })
        {
            var renderer = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            PdfPage page = renderer.RenderBlankPages(NamedFixture("docx-omml-styled-owner-" + suffix),
                null, CancellationToken.None).Single();
            PdfFontResource font = page.Fonts.Single(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal));
            TestAssert.True(font.Font.TryGetEncodedCid(font.Font.Font.MapCodePoint(0x1D465), out ushort cid),
                "Expected the mathematical base's owned CID.");
            var operation = System.Text.RegularExpressions.Regex.Match(page.Content,
                @"1 0 0 1 ([-\d.]+) [-\d.]+ Tm\s*<" + cid.ToString("X4", System.Globalization.CultureInfo.InvariantCulture) + @"> Tj");
            TestAssert.True(operation.Success, "Expected the untransformed mathematical base placement.");
            double actual = double.Parse(operation.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
            var frame = System.Text.RegularExpressions.Regex.Match(page.Content, @"q\s*1 0 0 1 ([-\d.]+) [-\d.]+ cm");
            TestAssert.True(frame.Success, "Expected the source equation's stored frame transform.");
            actual += double.Parse(frame.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
            TestAssert.True(Math.Abs(actual - wordX) < 0.20d,
                $"Word centers the logical script advance: actual{actual:F3}/expected{wordX:F3}; slanted ink width shifts the base by roughly0.75pt.");
            TestAssert.Contains("0.333", page.Content);
        }
    }

    public static void DocxNativeNormalTextStylesSelectAuthoredFacesAndMeasuredMathSlant()
    {
        RequireMathFont();
        foreach (string face in new[] { "calibri", "arial" })
        foreach (string style in new[] { "b", "i", "bi" })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            PdfPage page = Renderer().RenderBlankPages(NamedFixture($"docx-omml-normal-text-style-{face}-{style}"),
                diagnostics.Add, CancellationToken.None).Single();
            TestAssert.Contains("sin x", DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            PdfFontResource text = page.Fonts.First(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal));
            TestAssert.Equal(style.Contains('b'), text.Font.Font.Os2.WeightClass >= 600);
            TestAssert.Equal(style.Contains('i'), Math.Abs(text.Font.Font.Post.ItalicAngle) > 0.01d);
            TestAssert.True(diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "An available normal-text face owns its authored weight and slant.");
        }
        PdfPage synthetic = Renderer().RenderBlankPages(NamedFixture("docx-omml-normal-text-style-default-i"),
            null, CancellationToken.None).Single();
        TestAssert.Contains("sin x", DocxPageFootnoteNumberingTests.DecodeText(synthetic).Normalize(NormalizationForm.FormKC));
        TestAssert.Contains("0.333", synthetic.Content);
    }

    public static void DocxNativeNormalFractionsRetainParentMetricsAndChildFontMeaning()
    {
        RequireMathFont();
        foreach ((string suffix, string meaning, int resources) in new[]
        {
            ("display-default-normal", "ratetime", 1), ("display-calibri-normal", "ratetime", 1),
            ("inline-arial-normal", "ratetime", 1), ("display-calibri-mixed", "ratex", 2), ("inline-arial-mixed", "ratex", 2)
        })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            PdfPage page = Renderer().RenderBlankPages(NamedFixture("docx-omml-normal-fraction-" + suffix),
                diagnostics.Add, CancellationToken.None).Single();
            TestAssert.Contains(meaning, DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.Equal(resources, page.Fonts.Count(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)));
            TestAssert.Contains(" re f", page.Content);
            TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_APPROXIMATED_EQUATION") &&
                diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Normal/mixed fraction children retain actual native rules, parent metrics and source-face ownership.");
        }
    }

    public static void DocxNativeNormalScriptArgumentsUseTheParentScaleAndOwnTheirTextFace()
    {
        RequireMathFont();
        foreach ((string name, double size, int resources) in new[]
        {
            ("sub-default-default", 8.52d, 1), ("sub-calibri-default", 8.52d, 2),
            ("sup-arial-default", 8.52d, 2), ("sub-calibri-minus1", 6.96d, 2), ("sup-arial-plus1", 12d, 2)
        })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            PdfPage page = Renderer().RenderBlankPages(NamedFixture("docx-omml-normal-script-" + name),
                diagnostics.Add, CancellationToken.None).Single();
            TestAssert.Contains("xmax", DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.Contains(" " + size.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture) + " Tf", page.Content);
            TestAssert.Equal(resources, page.Fonts.Count(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)));
            TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_APPROXIMATED_EQUATION") &&
                diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Normal-text script arguments must retain actual parent/child font and source-role ownership.");
        }
    }

    public static void DocxNativeNormalTextRetainsTextAndMathematicalVariableOwnership()
    {
        RequireMathFont();
        foreach (string placement in new[] { "display", "inline" })
        foreach (string face in new[] { "default", "calibri" })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            PdfPage page = Renderer().RenderBlankPages(NamedFixture($"docx-omml-normal-text-{placement}-{face}-true"),
                diagnostics.Add, CancellationToken.None).Single();
            TestAssert.Contains("sin x", DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_APPROXIMATED_EQUATION") &&
                diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Normal text and mathematical variables must retain their shared native equation ownership.");
            var resources = page.Fonts.Where(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(face == "default" ? 1 : 2, resources.Length);
        }
    }

    public static void DocxNativeExplicitFalseNormalTextUsesMathGlyphsAndFace()
    {
        RequireMathFont();
        var diagnostics = new List<OoxPdfDiagnostic>();
        PdfPage page = Renderer().RenderBlankPages(NamedFixture("docx-omml-normal-text-display-calibri-false"),
            diagnostics.Add, CancellationToken.None).Single();
        PdfFontResource font = page.Fonts.Single(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal));
        TestAssert.True(font.Font.Font.MapCodePoint(0x1D460) != 0, "Explicit false retains the mathematical italic alphabet.");
        TestAssert.True(font.Font.TryGetEncodedCid(font.Font.Font.MapCodePoint(0x1D460), out _),
            "The math face owns the italic s glyph, rather than the authored normal-text face.");
        TestAssert.True(diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"),
            "An explicit false property is a supported math-mode control.");
    }

    public static void DocxNativeVariantSymbolsRetainTheirDistinctStyledSubsetMappings()
    {
        RequireMathFont();
        foreach ((string style, int[] points) in new[]
        {
            ("i", new[] { 0x1D717, 0x1D719, 0x1D716 }),
            ("b", new[] { 0x1D6DD, 0x1D6DF, 0x1D6DC }),
            ("bi", new[] { 0x1D751, 0x1D753, 0x1D750 })
        })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            PdfPage page = Renderer().RenderBlankPages(NamedFixture($"docx-omml-symbol-style-variants-{style}"),
                diagnostics.Add, CancellationToken.None).Single();
            PdfFontResource resource = page.Fonts.Single(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal));
            foreach (int point in points)
            {
                TestAssert.True(resource.Font.TryGetEncodedCid(resource.Font.Font.MapCodePoint(point), out ushort cid),
                    "A distinct variant glyph must own an encoded CID.");
                TestAssert.Equal(point, resource.Font.UnicodeByCid[cid]);
            }
            TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_APPROXIMATED_EQUATION") &&
                diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Supported variant symbols must report their native painting.");
        }
    }

    public static void DocxNativeGreekStylesRetainLowerUpperAndScriptMeaning()
    {
        RequireMathFont();
        foreach ((string kind, string meaning) in new[] { ("lower", "α+β=γ"), ("upper", "Γ+Θ=Σ"), ("scripts", "αβ2") })
        foreach (string style in new[] { "b", "bi" })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            PdfPage page = Renderer().RenderBlankPages(NamedFixture($"docx-omml-greek-style-{kind}-{style}"),
                diagnostics.Add, CancellationToken.None).Single();
            TestAssert.Contains(meaning, DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_APPROXIMATED_EQUATION") &&
                diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Styled Greek letters must retain their actual native glyph/subset ownership.");
        }
    }

    public static void DocxNativeExplicitMathStylesPaintTheirSourceMeaningInDisplayAndInlineRows()
    {
        RequireMathFont();
        foreach (string placement in new[] { "display", "inline" })
        foreach (string style in new[] { "b", "bi" })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            PdfPage page = Renderer().RenderBlankPages(NamedFixture($"docx-omml-style-literal-{placement}-12-{style}"),
                diagnostics.Add, CancellationToken.None).Single();
            TestAssert.Contains("Ah+h=x", DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_APPROXIMATED_EQUATION") &&
                diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Explicit math styles must report actual native painting through the shared box/font path.");
        }
    }

    public static void DocxNativeExplicitMathScriptStylesOwnTheirLetterAndDigitSubsetMappings()
    {
        RequireMathFont();
        foreach ((string style, int x, int i) in new[] { ("b", 0x1D431, 0x1D422), ("bi", 0x1D499, 0x1D48A) })
        {
            PdfPage page = Renderer().RenderBlankPages(NamedFixture($"docx-omml-style-scripts-display-12-{style}"),
                null, CancellationToken.None).Single();
            TestAssert.Contains("xi2", DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            PdfFontResource resource = page.Fonts.Single(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal));
            foreach ((int point, int level) in new[] { (x, 0), (i, 1), (0x1D7D0, 1) })
            {
                ushort glyph = resource.Font.Font.MapCodePoint(point);
                if (level != 0) { TestAssert.True(resource.Font.Font.TryGetMathScriptGlyph(glyph, level, CancellationToken.None, out glyph),
                    "Expected the style's script glyph ownership."); }
                TestAssert.True(resource.Font.TryGetEncodedCid(glyph, out ushort cid), "Styled source glyphs must own a published subset CID.");
                TestAssert.Equal(point, resource.Font.UnicodeByCid[cid]);
            }
        }
    }

    public static void DocxNativeDisplayFractionOperatorGapUsesTheParentMathSize()
    {
        RequireMathFont();
        PdfPage page = Renderer().RenderBlankPages(NamedFixture("docx-omml-display-spacing-fraction-child18-parent12-auto240"),
            null, CancellationToken.None).Single();
        PdfFontResource resource = page.Fonts.Single(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal));
        double X(int point)
        {
            TestAssert.True(resource.Font.TryGetEncodedCid(resource.Font.Font.MapCodePoint(point), out ushort cid),
                "Expected an owned ordinary display argument glyph.");
            var operation = System.Text.RegularExpressions.Regex.Match(page.Content,
                @"1 0 0 1 ([-\d.]+) [-\d.]+ Tm\s*<" + cid.ToString("X4", System.Globalization.CultureInfo.InvariantCulture) + @"> Tj");
            TestAssert.True(operation.Success, "Expected stored numerator glyph placement.");
            return double.Parse(operation.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
        }
        TestAssert.True(Math.Abs(X('+') - X(0x1D44E) - 12.60d) < 0.60d,
            "The Word12.60pt transition remains approximate; glyph-owned14.47pt operator spacing exceeds the qualified residual.");
    }

    public static void DocxNativeDisplayFractionParentMetricsRetainWordChildBaselineSeparation()
    {
        RequireMathFont();
        PdfPage page = Renderer().RenderBlankPages(NamedFixture("docx-omml-display-spacing-fraction-child18-parent12-auto240"),
            null, CancellationToken.None).Single();
        PdfFontResource resource = page.Fonts.Single(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal));
        double Y(int point)
        {
            TestAssert.True(resource.Font.TryGetEncodedCid(resource.Font.Font.MapCodePoint(point), out ushort cid),
                "Expected an owned display-fraction child glyph.");
            var operation = System.Text.RegularExpressions.Regex.Match(page.Content,
                @"1 0 0 1 [-\d.]+ ([-\d.]+) Tm\s*<" + cid.ToString("X4", System.Globalization.CultureInfo.InvariantCulture) + @"> Tj");
            TestAssert.True(operation.Success, "Expected stored numerator/denominator baselines.");
            return double.Parse(operation.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
        }
        TestAssert.True(Math.Abs(Y(0x1D44E) - Y(0x1D450) - 17.16d) < 0.15d,
            "Word separates enlarged children by17.16pt using parent12 metrics; child-owned25.66pt shifts fail.");
    }

    public static void DocxNativeDisplayAutomaticSpacingPaintsLiteralFractionAndScriptMeaning()
    {
        RequireMathFont();
        foreach ((string kind, string place, string meaning) in new[]
            { ("literal", "first", "α+β=γ"), ("fraction", "first", "a+bc"), ("scripts", "after", "xi2") })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            PdfPage page = Renderer().RenderBlankPages(NamedFixture($"docx-omml-display-spacing-{kind}-{place}-auto360"),
                diagnostics.Add, CancellationToken.None).Single();
            TestAssert.Contains(meaning, DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_APPROXIMATED_EQUATION") &&
                diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Expanded automatic spacing must retain actual native mathematical meaning.");
        }
    }

    public static void DocxNativeDisplaySpacingUsesTheMathFontLineRatherThanTheParentSize()
    {
        RequireMathFont();
        foreach ((string stem, double expected) in new[]
            { ("literal-first", 7.08d), ("literal-child18-parent12", 10.56d) })
        {
            double Tail(int line)
            {
                DocxDocument document = NamedFixture($"docx-omml-display-spacing-{stem}-auto{line}");
                // Inspect actual following-prose placement; source child18/parent12
                // independently distinguishes math-font spacing from parent spacing.
                return Renderer().InspectTextEmission(document).Lines.Last().Segments[0].BaselineY;
            }
            TestAssert.True(Math.Abs(Tail(240) - Tail(360) - expected) < 0.15d,
                "The extra paragraph advance must agree with the independent Word font-line contrast.");
        }
    }

    public static void DocxNativeDisplayMathCompressedAutomaticSlotsKeepTheirFallback()
    {
        RequireMathFont();
        var diagnostics = new List<OoxPdfDiagnostic>();
        PdfPage page = Renderer().RenderBlankPages(NamedFixture("docx-omml-display-spacing-literal-first-auto120"),
            diagnostics.Add, CancellationToken.None).Single();
        TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_UNSUPPORTED_EQUATION") &&
            diagnostics.All(diagnostic => diagnostic.Id != "DOCX_APPROXIMATED_EQUATION"),
            "Compressed slots need their own baseline contract and retain explicit fallback.");
        TestAssert.True(!page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)),
            "Unpainted math cannot publish unused resources.");
    }

    public static void DocxNativeSignedArgumentSizesPaintTheirSourceOwnedRoles()
    {
        RequireMathFont();
        foreach ((string name, double size) in new[]
            { ("docx-omml-argsize-both-12-minus1", 6.96d), ("docx-omml-argsize-both-12-plus1", 12d),
              ("docx-omml-argsize-both-12-absentval", 8.52d) })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            PdfPage page = Renderer().RenderBlankPages(NamedFixture(name), diagnostics.Add, CancellationToken.None).Single();
            TestAssert.Contains("xi2", DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.Contains(" " + size.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture) + " Tf", page.Content);
            TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_APPROXIMATED_EQUATION") &&
                diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Supported signed/default argument controls must report their actual native painting.");
        }
    }

    public static void DocxNativeNeutralArgumentPropertiesPaintDefaultScriptRoles()
    {
        RequireMathFont();
        var diagnostics = new List<OoxPdfDiagnostic>();
        PdfPage page = Renderer().RenderBlankPages(NamedFixture("docx-omml-script-scale-arg0"), diagnostics.Add, CancellationToken.None).Single();
        TestAssert.Contains("xi2", DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
        TestAssert.Contains(" 8.52 Tf", page.Content);
        TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_APPROXIMATED_EQUATION"),
            "Explicit neutral argument size must paint the default source script roles.");
        TestAssert.True(diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"),
            "Neutral properties cannot report ignored mathematics.");
    }

    public static void DocxNativeDisplayMathRetainsMeaningAndReportsPaintedContent()
    {
        RequireMathFont();
        DocxDocument document = Fixture("literal", "after");
        var diagnostics = new List<OoxPdfDiagnostic>();
        PdfPage page = Renderer().RenderBlankPages(document, diagnostics.Add, CancellationToken.None).Single();
        string text = DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC);
        TestAssert.Contains("α+β=γ", text);
        TestAssert.Contains("Before formula:", text);
        TestAssert.Contains("Following prose.", text);
        TestAssert.True(page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)), "Native math must own an embedded font resource.");
        TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_APPROXIMATED_EQUATION" && diagnostic.PartName == "/word/document.xml"), "Painted math must be reported as approximate with source provenance.");
        TestAssert.True(diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"), "A fully painted source part must not report ignored mathematics.");
    }

    public static void DocxNativeMathUnsupportedStructuredRowsKeepTheirFallback()
    {
        RequireMathFont();
        DocxDocument source = Fixture("fraction", "inline");
        DocxParagraph paragraph = source.Paragraphs[0];
        DocxEquation equation = paragraph.Equations[0];
        DocxMathFraction fraction = (DocxMathFraction)equation.Content.Children[0];
        DocxParagraph unsupported = paragraph with { Equations = [equation with
        { Content = new([fraction with { Numerator = new([fraction]) }]) }] };
        DocxDocument document = source with { BodyElements = source.BodyElements.Select(element =>
            element is DocxParagraphElement p && ReferenceEquals(p.Paragraph, paragraph) ? new DocxParagraphElement(unsupported) : element).ToArray() };
        var diagnostics = new List<OoxPdfDiagnostic>();
        PdfPage page = Renderer().RenderBlankPages(document, diagnostics.Add, CancellationToken.None).Single();
        TestAssert.True(!page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)), "Unsupported math must not publish unused equation fonts.");
        TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_UNSUPPORTED_EQUATION" && diagnostic.Fallback == "Ignored"), "Unsupported structures retain an explicit ignored fallback.");
        TestAssert.True(diagnostics.All(diagnostic => diagnostic.Id != "DOCX_APPROXIMATED_EQUATION"), "Unpainted mathematics cannot claim native rendering.");
    }

    public static void DocxNativeInlineFractionPreservesSourceRolesAlongsideProse()
    {
        RequireMathFont();
        var diagnostics = new List<OoxPdfDiagnostic>();
        PdfPage page = Renderer().RenderBlankPages(Fixture("fraction", "inline"), diagnostics.Add, CancellationToken.None).Single();
        string text = DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC);
        TestAssert.Contains("a+bc", text);
        TestAssert.Contains("Before formula:", text);
        TestAssert.Contains("Following prose.", text);
        TestAssert.Contains(" 8.52 Tf", page.Content);
        TestAssert.Contains(" re f", page.Content);
        TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_APPROXIMATED_EQUATION"),
            "A painted inline fraction must report actual native approximate rendering.");
        TestAssert.True(diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"),
            "Painted numerator/denominator/rule cannot report ignored mathematics.");
    }

    public static void DocxNativeInlineFractionChildSizesDoNotTakeOwnershipOfParentShifts()
    {
        RequireMathFont();
        DocxDocument document = NamedFixture("docx-omml-inline-fraction-tall18");
        DocxEquation equation = document.Paragraphs[0].Equations[0];
        TestAssert.True(equation.ParentFontSize is { } parent && Math.Abs(parent - 12d) < 1e-9,
            "The source paragraph owns the parent math metric size independently of18pt child formatting.");
        PdfPage page = Renderer().RenderBlankPages(document, null, CancellationToken.None).Single();
        PdfFontResource resource = page.Fonts.Single(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal));
        double Y(int point)
        {
            TestAssert.True(resource.Font.Font.TryGetMathScriptGlyph(resource.Font.Font.MapCodePoint(point), 1,
                CancellationToken.None, out ushort glyph), "Expected a script-style argument glyph.");
            TestAssert.True(resource.Font.TryGetEncodedCid(glyph, out ushort cid), "Expected an owned fraction subset CID.");
            var operation = System.Text.RegularExpressions.Regex.Match(page.Content,
                @"1 0 0 1 [-\d.]+ ([-\d.]+) Tm\s*<" + cid.ToString("X4", System.Globalization.CultureInfo.InvariantCulture) + @"> Tj");
            TestAssert.True(operation.Success, "Expected a stored child baseline.");
            return double.Parse(operation.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
        }
        TestAssert.True(Math.Abs(Y(0x1D44E) - Y(0x1D450) - 13.08d) < 0.15d,
            "The independently measured parent shifts separate tall child baselines by13.08pt; child-owned19.60pt spacing fails.");
    }

    public static void DocxNativeInlineScriptsPreserveMeaningAlongsideSourceOwnedProse()
    {
        RequireMathFont();
        var diagnostics = new List<OoxPdfDiagnostic>();
        PdfPage page = Renderer().RenderBlankPages(Fixture("scripts", "inline"), diagnostics.Add, CancellationToken.None).Single();
        string text = DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC);
        TestAssert.Contains("xi2", text);
        TestAssert.Contains("Before formula:", text);
        TestAssert.Contains("Following prose.", text);
        TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_APPROXIMATED_EQUATION"),
            "Supported inline script roles must report native approximate painting.");
        TestAssert.True(diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"),
            "Painted inline scripts cannot report ignored mathematics.");
    }

    public static void DocxNativeInlineScriptsWrapFollowingProseAndRetainOverwideFallback()
    {
        RequireMathFont();
        var diagnostics = new List<OoxPdfDiagnostic>();
        PdfPage fitting = Renderer().RenderBlankPages(NamedFixture("docx-omml-inline-script-wrap110"), diagnostics.Add, CancellationToken.None).Single();
        string text = DocxPageFootnoteNumberingTests.DecodeText(fitting).Normalize(NormalizationForm.FormKC);
        TestAssert.Contains("xi2", text);
        TestAssert.Contains("after formula.", text);
        TestAssert.True(diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"),
            "Following-prose wrapping must share the same script source owner.");
        diagnostics.Clear();
        PdfPage overwide = Renderer().RenderBlankPages(NamedFixture("docx-omml-inline-script-overwide30"), diagnostics.Add, CancellationToken.None).Single();
        TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_UNSUPPORTED_EQUATION"),
            "Unqualified formula relocation retains its whole-equation fallback.");
        TestAssert.True(!overwide.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)),
            "Fallback output cannot publish an unused mathematical subset.");
    }

    public static void DocxNativeInlineScriptFrameKeepsFollowingProseWithinItsAdmittedBody()
    {
        RequireMathFont();
        DocxDocument document = NamedFixture("docx-omml-inline-script-boundary140");
        DocxRenderer renderer = Renderer();
        PdfPage page = renderer.RenderBlankPages(document, null, CancellationToken.None).Single();
        string text = DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC);
        TestAssert.Contains("xi2", text);
        TestAssert.Contains("Following prose.", text);
        var tail = renderer.InspectTextEmission(document).Lines.Last(line => line.SourceBlockIndex == 8);
        TestAssert.True(tail.Segments.All(segment => segment.BaselineY >= document.MarginBottomPoints),
            "The shared inline script frame must keep its following prose within the admitted source body.");
    }

    public static void DocxNativeDisplayScriptsPaintDistinctBaseAndScriptRoles()
    {
        RequireMathFont();
        var diagnostics = new List<OoxPdfDiagnostic>();
        PdfPage page = Renderer().RenderBlankPages(Fixture("scripts", "after"), diagnostics.Add, CancellationToken.None).Single();
        string text = DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC);
        TestAssert.Contains("xi2", text);
        TestAssert.Contains("Following prose.", text);
        TestAssert.Contains(" 8.52 Tf", page.Content);
        TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_APPROXIMATED_EQUATION"),
            "Painted source roles and script-style glyphs must report native approximate math.");
        TestAssert.True(diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"),
            "A fully painted display script cannot report ignored source content.");
    }

    public static void DocxNativeDisplayScriptsOwnTheirSelectedVariantSubsetMappings()
    {
        RequireMathFont();
        PdfPage page = Renderer().RenderBlankPages(Fixture("scripts", "after"), null, CancellationToken.None).Single();
        PdfFontResource resource = page.Fonts.Single(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal));
        foreach (int point in new[] { 0x1D456, (int)'2' })
        {
            ushort normal = resource.Font.Font.MapCodePoint(point);
            TestAssert.True(resource.Font.Font.TryGetMathScriptGlyph(normal, 1, CancellationToken.None, out ushort chosen),
                "Native script roles must resolve their source font's script-style glyph.");
            TestAssert.True(resource.Font.TryGetEncodedCid(chosen, out ushort cid), "The selected variant must own a published subset CID.");
            TestAssert.Equal(point, resource.Font.UnicodeByCid[cid]);
        }
    }

    public static void DocxNativeDisplayIndividualScriptsRetainTheirDistinctMeaning()
    {
        RequireMathFont();
        foreach ((string name, string expected) in new[] { ("docx-omml-script-sub12", "xi"), ("docx-omml-script-sup12", "x2") })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            PdfPage page = Renderer().RenderBlankPages(NamedFixture(name), diagnostics.Add, CancellationToken.None).Single();
            TestAssert.Contains(expected, DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC));
            TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_APPROXIMATED_EQUATION"),
                "Single source-script roles must report their actual native painting.");
        }
    }

    public static void DocxNativeDisplayScriptFrameAdmitsItsFollowingProseAtTheBodyBoundary()
    {
        RequireMathFont();
        DocxDocument document = NamedFixture("docx-omml-script-boundary12");
        DocxRenderer renderer = Renderer();
        PdfPage page = renderer.RenderBlankPages(document, null, CancellationToken.None).Single();
        string text = DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC);
        TestAssert.Contains("xi2", text);
        TestAssert.Contains("Following prose.", text);
        var tail = renderer.InspectTextEmission(document).Lines.Last(line => line.SourceBlockIndex == 7);
        TestAssert.True(tail.Segments.All(segment => segment.BaselineY >= document.MarginBottomPoints),
            "The shared script frame must admit following prose without crossing the source body boundary.");
    }

    public static void DocxNativeDisplayDescenderScriptUsesGlyphRatherThanBoxDropMinimum()
    {
        RequireMathFont();
        PdfPage page = Renderer().RenderBlankPages(NamedFixture("docx-omml-script-descender12"), null, CancellationToken.None).Single();
        PdfFontResource resource = page.Fonts.Single(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal));
        ushort original = resource.Font.Font.MapCodePoint(0x1D456);
        TestAssert.True(resource.Font.Font.TryGetMathScriptGlyph(original, 1, CancellationToken.None, out ushort script), "Expected script-style i.");
        TestAssert.True(resource.Font.TryGetEncodedCid(script, out ushort cid), "Expected an owned script subset CID.");
        string code = cid.ToString("X4", System.Globalization.CultureInfo.InvariantCulture);
        var operation = System.Text.RegularExpressions.Regex.Match(page.Content,
            @"1 0 0 1 [-\d.]+ ([-\d.]+) Tm\s*<" + code + @"> Tj");
        TestAssert.True(operation.Success, "The native script must retain a stored relative baseline.");
        double baseline = double.Parse(operation.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
        TestAssert.True(baseline >= -3.3d,
            "The Word-backed descender control cannot receive the extra composed-box drop (old baseline below -4.4pt).");
    }

    public static void DocxNativeDisplayMathCornerKernMovesTheScriptToItsMeasuredOfficeGap()
    {
        RequireMathFont();
        PdfPage page = Renderer().RenderBlankPages(NamedFixture("docx-omml-kern-y-sub"), null, CancellationToken.None).Single();
        PdfFontResource resource = page.Fonts.Single(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal));
        ushort basis = resource.Font.Font.MapCodePoint(0x1D466);
        TestAssert.True(resource.Font.Font.TryGetMathScriptGlyph(resource.Font.Font.MapCodePoint(0x1D456), 1,
            CancellationToken.None, out ushort script), "Expected script-style i.");
        TestAssert.True(resource.Font.TryGetEncodedCid(basis, out ushort baseCid), "The base must own a published subset CID.");
        TestAssert.True(resource.Font.TryGetEncodedCid(script, out ushort scriptCid), "The script must own a published subset CID.");
        double X(ushort cid)
        {
            string code = cid.ToString("X4", System.Globalization.CultureInfo.InvariantCulture);
            var operation = System.Text.RegularExpressions.Regex.Match(page.Content,
                @"1 0 0 1 ([-\d.]+) [-\d.]+ Tm\s*<" + code + @"> Tj");
            TestAssert.True(operation.Success, "Expected a stored source glyph position.");
            return double.Parse(operation.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
        }
        TestAssert.True(Math.Abs(X(scriptCid) - X(baseCid) - 6.36d) < 0.15d,
            "The independently measured Word gap is6.36pt; the unkerned6.71pt and minimum-sum5.66pt placements fail this contrast.");
    }

    public static void DocxNativeDisplayFractionPaintsDistinctSourceRolesAndStoredRule()
    {
        RequireMathFont();
        var diagnostics = new List<OoxPdfDiagnostic>();
        PdfPage page = Renderer().RenderBlankPages(Fixture("fraction", "after"), diagnostics.Add, CancellationToken.None).Single();
        string text = DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC);
        TestAssert.Contains("a+bc", text);
        TestAssert.Contains("Following prose.", text);
        TestAssert.Contains(" re f", page.Content);
        TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_APPROXIMATED_EQUATION"),
            "Painted numerator, denominator and rule must report approximate native rendering.");
        TestAssert.True(diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"),
            "A fully painted display fraction cannot report ignored source content.");
    }

    public static void DocxNativeDisplayFractionUsesItsFullFrameBeforeFollowingProseTurnsPage()
    {
        RequireMathFont();
        DocxDocument document = NamedFixture("docx-omml-display-fraction-boundary12");
        PdfPage[] pages = Renderer().RenderBlankPages(document, null, CancellationToken.None).ToArray();
        TestAssert.Equal(2, pages.Length);
        string first = DocxPageFootnoteNumberingTests.DecodeText(pages[0]).Normalize(NormalizationForm.FormKC);
        string second = DocxPageFootnoteNumberingTests.DecodeText(pages[1]).Normalize(NormalizationForm.FormKC);
        TestAssert.Contains("a+bc", first);
        TestAssert.True(!first.Contains("Following prose.", StringComparison.Ordinal),
            "Fraction ink, font leading and denominator must consume one complete frame before following prose is admitted.");
        TestAssert.Contains("Following prose.", second);
    }

    public static void DocxNativeInlineMathRetainsLiteralMeaningAndSharedProseBaseline()
    {
        RequireMathFont();
        DocxDocument document = Fixture("literal", "inline");
        var diagnostics = new List<OoxPdfDiagnostic>();
        PdfPage page = Renderer().RenderBlankPages(document, diagnostics.Add, CancellationToken.None).Single();
        string text = DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC);
        TestAssert.Contains("α+β=γ", text);
        TestAssert.Contains("Before formula:", text);
        TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_APPROXIMATED_EQUATION"), "Fitting inline content must report native approximate painting.");
        TestAssert.True(diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"), "Painted inline content cannot report ignored math.");
    }

    public static void DocxNativeInlineMathMovesFollowingProsePastItsMeasuredBox()
    {
        RequireMathFont();
        DocxDocument document = NamedFixture("docx-omml-inline-middle12");
        DocxRenderer renderer = Renderer();
        PdfPage page = renderer.RenderBlankPages(document, null, CancellationToken.None).Single();
        string text = DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC);
        TestAssert.Contains("α+β=γ", text);
        TestAssert.Contains(" after formula.", text);
        var segments = renderer.InspectTextEmission(document).Lines[0].Segments;
        double begin = segments[0].X;
        TestAssert.True(segments.Where(segment => !segment.IsTerminalLineSpace).Max(segment => segment.X) > begin + 120d,
            "Prose after a source-owned formula must move beyond its mathematical box.");
        TestAssert.True(segments.Where(segment => segment.IsTerminalLineSpace).All(segment => segment.X > begin + 200d),
            "The paragraph terminal cannot overlap inline mathematical content or following prose.");
    }

    public static void DocxNativeInlineMathOverfullRowsRetainTheWholeIgnoredFallback()
    {
        RequireMathFont();
        var diagnostics = new List<OoxPdfDiagnostic>();
        PdfPage page = Renderer().RenderBlankPages(NamedFixture("docx-omml-inline-overfull12"), diagnostics.Add, CancellationToken.None).Single();
        TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_UNSUPPORTED_EQUATION"), "Formula-driven wrapping retains an explicit fallback until qualified.");
        TestAssert.True(diagnostics.All(diagnostic => diagnostic.Id != "DOCX_APPROXIMATED_EQUATION"), "An overfull row cannot claim painted math.");
        TestAssert.True(!page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)), "Fallback output cannot publish unused math fonts.");
    }

    public static void DocxNativeInlineMathWrapsFollowingProseWithoutLosingItsSourceBoundary()
    {
        RequireMathFont();
        DocxDocument document = NamedFixture("docx-omml-wrap-middle140");
        var diagnostics = new List<OoxPdfDiagnostic>();
        DocxRenderer renderer = Renderer();
        PdfPage page = renderer.RenderBlankPages(document, diagnostics.Add, CancellationToken.None).Single();
        string text = DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC);
        TestAssert.Contains("α+β=γ", text);
        TestAssert.Contains("after formula.", text);
        TestAssert.Contains("Following prose.", text);
        TestAssert.True(!text.Contains('\u200B'), "Layout markers cannot leak into emitted prose.");
        var lines = renderer.InspectTextEmission(document).Lines.Where(line => line.SourceBlockIndex == 0).ToArray();
        TestAssert.Equal(2, lines.Length);
        TestAssert.True(lines[0].Segments[0].BaselineY > lines[1].Segments[0].BaselineY,
            "Following prose must own a subsequent line, preserving the mathematical source position.");
        TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_APPROXIMATED_EQUATION"),
            "A fitting formula remains painted when only its following prose requires wrapping.");
        TestAssert.True(diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"),
            "Supported prose wrapping cannot claim ignored mathematical content.");
    }

    public static void DocxNativeInlineMathWrapsBeginningAndLongTailProseWithoutChangingSourceText()
    {
        RequireMathFont();
        foreach (string name in new[] { "docx-omml-prosewrap-begin90", "docx-omml-prosewrap-long140" })
        {
            DocxDocument document = NamedFixture(name);
            DocxRenderer renderer = Renderer();
            PdfPage page = renderer.RenderBlankPages(document, null, CancellationToken.None).Single();
            string text = DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC);
            TestAssert.Contains("α+β=γ", text);
            TestAssert.Contains("Following prose.", text);
            string prose = string.Concat(document.Paragraphs[0].Runs.Select(run => run.Text));
            TestAssert.True(prose.Split(' ', StringSplitOptions.RemoveEmptyEntries).All(word => text.Contains(word, StringComparison.Ordinal)),
                "Every following prose word must survive source-owned formula wrapping.");
            var lines = renderer.InspectTextEmission(document).Lines.Where(line => line.SourceBlockIndex == 0).ToArray();
            TestAssert.True(lines.Length >= 2, "The mathematical box must participate in following-prose line breaking.");
        }
    }

    public static void DocxNativeMathSettingsAndExplicitRunFacesRemainDistinct()
    {
        DocxDocument document = Fixture("literal", "after");
        TestAssert.Equal("Cambria Math", document.Settings.MathFontFamily);
        DocxEquation equation = document.Paragraphs.Single(paragraph => paragraph.Equations.Count != 0).Equations.Single();
        TestAssert.True(equation.Content.Children.Cast<DocxMathRun>().All(run => run.ExplicitFontFamily == "Cambria Math"), "Authored math faces must remain separate from inherited prose faces.");
    }

    public static void DocxNativeMathUsesItsOccupiedBodyFrameBeforeFollowingProseTurnsPage()
    {
        RequireMathFont();
        string input = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "docx-omml-display-boundary12.docx"));
        DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
        PdfPage[] pages = Renderer().RenderBlankPages(document, null, CancellationToken.None).ToArray();
        TestAssert.Equal(2, pages.Length);
        string first = DocxPageFootnoteNumberingTests.DecodeText(pages[0]).Normalize(NormalizationForm.FormKC);
        string second = DocxPageFootnoteNumberingTests.DecodeText(pages[1]).Normalize(NormalizationForm.FormKC);
        TestAssert.Contains("α+β=γ", first);
        TestAssert.True(!first.Contains("Following prose.", StringComparison.Ordinal), "The following paragraph must turn at the source body boundary.");
        TestAssert.Contains("Following prose.", second);
    }

    public static void DocxNativeMathMissingFontsAndUnqualifiedFormattingRetainIgnoredDiagnostics()
    {
        DocxDocument document = Fixture("literal", "after");
        var diagnostics = new List<OoxPdfDiagnostic>();
        PdfPage page = new DocxRenderer(new MapFontResolver([], "Unavailable"), OoxPdfDocxMarkupMode.Final,
            OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, diagnostics.Add, CancellationToken.None).Single();
        TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_UNSUPPORTED_EQUATION"), "Unavailable math fonts retain ignored-source diagnostics.");
        TestAssert.True(!page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)), "A missing math face cannot claim embedded equation output.");
        DocxParagraph paragraph = document.Paragraphs.Single(value => value.Equations.Count != 0);
        TestAssert.NotNull(DocxMathParagraphPolicy.GetDisplayRow(paragraph));
        TestAssert.True(DocxMathParagraphPolicy.GetDisplayRow(paragraph with
        { Spacing = paragraph.Spacing with { LineValue = "480", LineRuleValue = "exact" } }) is null,
            "Unqualified explicit math line rules must keep their existing fallback.");
        TestAssert.True(DocxMathParagraphPolicy.GetDisplayRow(paragraph with
        { InlineReferences = [new DocxInlineReference(DocxRelatedStoryKind.Footnote, "1", null, "1", 0, 0, 0)] }) is null,
            "References require their own semantic owner before math painting is admitted.");
    }

    private static DocxRenderer Renderer() => new(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
    private static DocxDocument Fixture(string kind, string placement) => DocxTests.ReadDocx(Path.GetFullPath(
        Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", $"docx-omml-{kind}-{placement}.docx")), OoxPdfDocxMarkupMode.Final);
    private static DocxDocument NamedFixture(string name) => DocxTests.ReadDocx(Path.GetFullPath(
        Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", name + ".docx")), OoxPdfDocxMarkupMode.Final);
    private static void RequireMathFont()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for this native rendering control.");
    }
}
