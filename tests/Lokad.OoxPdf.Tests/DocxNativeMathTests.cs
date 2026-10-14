using System.Text;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxNativeMathTests
{
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
