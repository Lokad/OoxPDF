using System.Buffers.Binary;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxMathNaryTests
{
    public static void MathFontNaryDocumentLimitSettingsRetainValuesMalformedSourceAndCancellation()
    {
        foreach (string? value in new string?[] { null, "", "subSup", "undOvr" })
        {
            string property = value is null ? "" : value.Length == 0 ? "<m:naryLim/><m:intLim/>" :
                $"<m:naryLim m:val='{value}'/><m:intLim m:val='{value}'/>";
            var settings = Settings(property);
            TestAssert.Equal(string.IsNullOrEmpty(value) ? null : value, settings.NaryValue);
            TestAssert.Equal(string.IsNullOrEmpty(value) ? null : value, settings.IntegralValue);
            TestAssert.True(!settings.UnsupportedNary && !settings.UnsupportedIntegral && settings.PropertiesXml.Count == 1,
                "Omitted versus present-default source remains retained without inventing authored values.");
        }
        foreach (string malformed in new[] { "<m:naryLim m:val='bad'/>", "<m:naryLim/><m:naryLim/>",
            "<m:naryLim extra='1'/>", "<m:naryLim>text</m:naryLim>", "<m:naryLim><m:future/></m:naryLim>", "<w:naryLim/>" })
        {
            var settings = Settings(malformed + "<m:intLim m:val='subSup'/>");
            TestAssert.True(settings.UnsupportedNary && !settings.UnsupportedIntegral && settings.PropertiesXml[0].Contains("naryLim", StringComparison.Ordinal),
                "Malformed or duplicate n-ary defaults remain explicit without invalidating an independent integral setting.");
        }
        var duplicate = DocxMathLimitSettings.Read(XElement.Parse("<w:settings xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'><m:mathPr/><m:mathPr/></w:settings>"), CancellationToken.None);
        TestAssert.True(duplicate.UnsupportedNary && duplicate.UnsupportedIntegral && duplicate.PropertiesXml.Count == 2,
            "Duplicate property owners retain their source and cannot silently choose one document default.");
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLimitSettings.Read(null, cancelled.Token));
    }

    public static void MathFontNaryDocumentDefaultsSharePlacementAndKeepNodePrecedence()
    {
        var font = OpenTypeFont.Load(FontBytes(0x2211));
        var resource = new PdfFontResource("Defaults", PdfEmbeddedFont.CreateForGlyphMappings(font,
            new[] { ((ushort)128, 0x2211), ((ushort)130, 0x2211), (font.MapCodePoint('B'), (int)'B'), (font.MapCodePoint('C'), (int)'C') }, CancellationToken.None));
        foreach (bool display in new[] { true, false })
        foreach (string location in new[] { "undOvr", "subSup" })
        {
            var settings = Settings($"<m:naryLim m:val='{location}'/><m:intLim m:val='bad'/>");
            var nary = Nary("\u2211") with { LimitLocationValue = null };
            var equation = Equation(nary) with { Display = display, LimitSettings = settings };
            TestAssert.True(DocxMathLayout.TryMeasure(equation, _ => resource, CancellationToken.None, out var inherited, out var reason), reason ?? "Expected document-owned limits.");
            TestAssert.True(DocxMathLayout.TryMeasure(equation with { Content = new([nary with { LimitLocationValue = location }]) },
                _ => resource, CancellationToken.None, out var explicitBox, out _), "Expected equivalent explicit placement.");
            TestAssert.True(TestAssert.NotNull(inherited).Glyphs.SequenceEqual(TestAssert.NotNull(explicitBox).Glyphs) && nary.LimitLocationValue is null,
                "Resolution reuses placement without overwriting the authored node or consuming unrelated invalid integral defaults.");
            var invalid = equation with { LimitSettings = Settings("<m:naryLim m:val='bad'/>") };
            TestAssert.True(!DocxMathLayout.TryMeasure(invalid, _ => resource, CancellationToken.None, out var absent, out _) && absent is null,
                "An unqualified document default cannot claim native placement.");
            TestAssert.True(DocxMathLayout.TryMeasure(invalid with { Content = new([nary with { LimitLocationValue = location }]) },
                _ => resource, CancellationToken.None, out _, out _), "Authored node placement takes precedence over unused malformed defaults.");
        }
    }

    private static DocxMathLimitSettings Settings(string properties) => DocxMathLimitSettings.Read(XElement.Parse(
        "<w:settings xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'><m:mathPr>" + properties + "</m:mathPr></w:settings>"), CancellationToken.None);

    public static void MathFontNaryStackedLimitsKeepIndependentResourcesAndCompleteFallback()
    {
        foreach (int point in new[] { 0x2211, 0x220F })
        {
            var font = OpenTypeFont.Load(FontBytes(point));
            var parent = new PdfFontResource("Operator", PdfEmbeddedFont.CreateForGlyphMappings(font,
                new[] { ((ushort)128, point) }, CancellationToken.None));
            var text = new PdfFontResource("Children", PdfEmbeddedFont.Create(font, new[] { (int)'B', (int)'C' }, CancellationToken.None));
            var nary = Nary(char.ConvertFromUtf32(point));
            var equation = Equation(nary);
            TestAssert.True(DocxMathLayout.TryGetRuns(equation, CancellationToken.None, out var sources) && sources.Count == 3 &&
                sources[0].ScriptLevel == 0 && sources.Skip(1).All(s => s.ScriptLevel == 1) &&
                ReferenceEquals(sources[0].ControlOwner, nary), "The operator owns one control while literal body and limits retain their source levels.");
            TestAssert.True(DocxMathLayout.TryMeasure(equation, _ => text, CancellationToken.None, out var box, out var reason,
                resolveControlFont: _ => parent), reason ?? "Expected native stacked limits.");
            var measured = TestAssert.NotNull(box);
            TestAssert.Equal(char.ConvertFromUtf32(point) + "BCB", string.Concat(measured.Glyphs.Select(g => char.ConvertFromUtf32(g.SourceCodePoint))));
            TestAssert.True(ReferenceEquals(measured.Glyphs[0].Font, parent) && measured.Glyphs[0].OriginalGlyph == 128 &&
                measured.Glyphs.Skip(1).All(g => ReferenceEquals(g.Font, text)) && measured.Glyphs[2].BaselineY > 0 &&
                measured.Glyphs[3].BaselineY < 0 && measured.ParentMetrics?.Size == 12d,
                "Selected operator variants and independently encoded children share one bounded occupied frame.");
            TestAssert.True(measured.Glyphs[0].Color.Red == 255 && measured.Glyphs.Skip(1).All(g => g.Color.Blue == 255),
                "Control color and child colors retain independent source ownership.");
            var missing = new PdfFontResource("Missing", PdfEmbeddedFont.CreateForGlyphMappings(font,
                new[] { ((ushort)130, point) }, CancellationToken.None));
            TestAssert.True(!DocxMathLayout.TryMeasure(equation, _ => text, CancellationToken.None, out var absent, out _,
                resolveControlFont: _ => missing) && absent is null, "A missing selected variant CID cannot publish a partial body or operator.");
            var graphics = new PdfGraphicsBuilder(); DocxMathBoxPainter.Paint(measured, graphics, 20d, 50d, CancellationToken.None);
            using var output = new MemoryStream();
            PdfDocumentWriter.WriteBlank(output, [new PdfPage(200d, 150d, graphics.ToString(), [parent, text])], CancellationToken.None);
            TestAssert.True(output.Length > 0, "Native operator and child resources remain valid PDF content.");
        }
    }

    public static void MathFontNaryEligibilityBoundsAndCancellationRemainExplicit()
    {
        var nary = Nary("\u2211");
        foreach (var invalid in new[] { nary with { CharacterValue = "\u222B" }, nary with { LimitLocationValue = "unknown" },
            nary with { GrowValue = false },
            nary with { SubscriptHiddenValue = true }, nary with { SuperscriptHiddenValue = true },
            nary with { Body = new([nary]) }, nary with { HasUnsupportedProperties = true } })
            TestAssert.True(!DocxMathLayout.TryGetRuns(Equation(invalid), CancellationToken.None, out _),
                "Unqualified defaults, roles and sizing retain complete source fallback.");
        var font = OpenTypeFont.Load(FontBytes(0x2211));
        var resource = new PdfFontResource("Bounded", PdfEmbeddedFont.CreateForGlyphMappings(font,
            new[] { ((ushort)128, 0x2211), (font.MapCodePoint('B'), (int)'B'), (font.MapCodePoint('C'), (int)'C') }, CancellationToken.None));
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryGetRuns(Equation(nary), cancelled.Token, out _));
        TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryMeasure(Equation(nary), _ => resource, cancelled.Token, out _, out _));
        using var scope = OoxConversionBudget.BeginScope(new OoxConversionLimits { MaxSceneNodesPerConversion = 3 });
        TestAssert.Throws<OoxPdfLimitExceededException>(() => DocxMathLayout.TryMeasure(Equation(nary), _ => resource, CancellationToken.None, out _, out _));
    }

    public static void MathFontNarySideLimitsShareScriptArgumentsAndInlinePlacement()
    {
        var font = OpenTypeFont.Load(FontBytes(0x2211));
        var resource = new PdfFontResource("Side", PdfEmbeddedFont.CreateForGlyphMappings(font,
            new[] { ((ushort)128, 0x2211), ((ushort)130, 0x2211), (font.MapCodePoint('B'), (int)'B'), (font.MapCodePoint('C'), (int)'C') }, CancellationToken.None));
        DocxMathBox? firstInline = null;
        foreach (bool display in new[] { true, false })
        foreach (string location in new[] { "subSup", "undOvr" })
        {
            var equation = Equation(Nary("\u2211") with { LimitLocationValue = location }) with { Display = display };
            TestAssert.True(DocxMathLayout.TryMeasure(equation, _ => resource, CancellationToken.None, out var box, out var reason), reason ?? "Expected qualified limit placement.");
            var measured = TestAssert.NotNull(box);
            TestAssert.True(measured.Glyphs[0].OriginalGlyph == (display ? 128 : 130), "Display minimum and inline base selection retain separate roles.");
            if (!display || location == "subSup")
                TestAssert.True(measured.Glyphs[2].X >= measured.Glyphs[0].Advance && measured.Glyphs[3].X == measured.Glyphs[2].X &&
                    measured.Glyphs[1].X >= measured.Glyphs[3].X + measured.Glyphs[3].Advance,
                    "Both side limits occupy the operator's right edge before the literal body.");
            if (!display)
            {
                if (firstInline is null) firstInline = measured;
                else TestAssert.True(firstInline.Glyphs.SequenceEqual(measured.Glyphs), "Inline undOvr and subSup share Office's side-limit placement.");
            }
        }
        var invalidFont = OpenTypeFont.Load(FontBytes(0x2211, invalidScriptConstants: true));
        var invalid = new PdfFontResource("Invalid", PdfEmbeddedFont.CreateForGlyphMappings(invalidFont,
            new[] { ((ushort)128, 0x2211), (invalidFont.MapCodePoint('B'), (int)'B'), (invalidFont.MapCodePoint('C'), (int)'C') }, CancellationToken.None));
        TestAssert.True(!DocxMathLayout.TryMeasure(Equation(Nary("\u2211") with { LimitLocationValue = "subSup" }), _ => invalid,
            CancellationToken.None, out var absent, out _) && absent is null, "Malformed script constants retain complete fallback for n-ary side limits.");
    }

    private static DocxEquation Equation(DocxMathNary nary) => new(new([nary]), 0, 0, true, null, null, null) { ParentFontSize = 12d };
    public static void MathFontInlineIntegralUsesExtendedShapeItalicCorrectionAndIndependentDefaults()
    {
        var font = OpenTypeFont.Load(FontBytes(0x222B, extendedOperator: true));
        var resource = new PdfFontResource("Integral", PdfEmbeddedFont.CreateForGlyphMappings(font,
            new[] { ((ushort)130, 0x222B), (font.MapCodePoint('B'), (int)'B'), (font.MapCodePoint('C'), (int)'C') }, CancellationToken.None));
        var nary = Nary("\u222B") with { LimitLocationValue = null };
        var equation = Equation(nary) with { Display = false, LimitSettings = new("bad", "undOvr", true, false, []) };
        TestAssert.True(DocxMathLayout.TryMeasure(equation, _ => resource, CancellationToken.None, out var box, out var reason), reason ?? "Expected qualified inline integral.");
        var measured = TestAssert.NotNull(box); var op = measured.Glyphs[0]; var upper = measured.Glyphs[2]; var lower = measured.Glyphs[3];
        TestAssert.True(op.OriginalGlyph == 130 && Math.Abs(upper.X - lower.X - 2.4d) < 1e-9,
            "A font-owned italic correction lets the lower integral limit precede the upper limit without a point-size branch.");
        TestAssert.True(upper.BaselineY - op.BaselineY >= 8d,
            "The extended-shape flag activates the occupied-base drop rule for an inline integral.");
        TestAssert.True(!DocxMathLayout.TryGetRuns(equation with { Display = true }, CancellationToken.None, out _),
            "Display integral sizing retains its independently measured unresolved fallback.");
        var invalid = equation with { LimitSettings = new("undOvr", "bad", false, true, []) };
        TestAssert.True(!DocxMathLayout.TryMeasure(invalid, _ => resource, CancellationToken.None, out var absent, out _) && absent is null,
            "Integral defaults retain independent ownership rather than consulting an unrelated n-ary setting.");
        TestAssert.True(DocxMathLayout.TryMeasure(invalid with { Content = new([nary with { LimitLocationValue = "subSup" }]) },
            _ => resource, CancellationToken.None, out _, out _), "Explicit placement overrides unused invalid integral defaults.");
    }

    public static void MathFontNaryStructuredBodiesShareBoundedArgumentsRulesAndCompleteFallback()
    {
        var font = OpenTypeFont.Load(FontBytes(0x2211, withFractions: true));
        var resource = new PdfFontResource("Body", PdfEmbeddedFont.CreateForGlyphMappings(font,
            new[] { ((ushort)128, 0x2211), ((ushort)130, 0x2211), (font.MapCodePoint('B'), (int)'B'), (font.MapCodePoint('C'), (int)'C') }, CancellationToken.None));
        foreach (bool display in new[] { true, false })
        foreach (string kind in new[] { "bar", "noBar", "script" })
        {
            DocxMathNode body = kind == "script" ? new DocxMathScripts(Row("B"), Row("C"), Row("B")) :
                new DocxMathFraction(Row("B"), Row("C"), kind);
            var nary = Nary("\u2211") with { Body = new([body]) };
            var equation = Equation(nary) with { Display = display };
            TestAssert.True(DocxMathLayout.TryGetRuns(equation, CancellationToken.None, out var sources) &&
                sources.Count == (kind == "script" ? 5 : 4) && ReferenceEquals(sources[0].ControlOwner, nary),
                "Structured bodies retain bounded child roles under one operator owner.");
            if (kind != "script") TestAssert.True(sources.Take(2).All(s => s.ScriptLevel == (display ? 0 : 1)),
                "Fraction/stack body arguments keep their display or inline style.");
            TestAssert.True(DocxMathLayout.TryMeasure(equation, _ => resource, CancellationToken.None, out var box, out var reason), reason ?? "Expected bounded native body.");
            var measured = TestAssert.NotNull(box);
            TestAssert.Equal(kind == "script" ? "\u2211BCBCB" : "\u2211BCCB",
                string.Concat(measured.Glyphs.Select(g => char.ConvertFromUtf32(g.SourceCodePoint))));
            TestAssert.Equal(kind == "bar" ? 1 : 0, measured.Rules.Count);
            TestAssert.True(measured.Glyphs[0].Color.Red == 255 && measured.Glyphs.Skip(1).All(g => g.Color.Blue == 255) &&
                measured.Rules.All(r => r.Color.Blue == 255 && r.X > measured.Glyphs[0].X),
                "Inner fraction rules and child colors remain independent of operator color and placement.");
            var missing = nary with { Body = new([new DocxMathFraction(Row("D"), Row("C"), "bar")]) };
            TestAssert.True(!DocxMathLayout.TryMeasure(equation with { Content = new([missing]) }, _ => resource,
                CancellationToken.None, out var absent, out _) && absent is null,
                "A missing inner-body CID cannot publish a partial operator or outer limits.");
            foreach (var deeper in new DocxMathNode[] { new DocxMathFraction(new([body]), Row("C"), "bar"),
                new DocxMathScripts(new([new DocxMathScripts(Row("B"), Row("C"), null)]), Row("C"), null) })
                TestAssert.True(!DocxMathLayout.TryGetRuns(equation with { Content = new([nary with { Body = new([deeper]) }]) },
                    CancellationToken.None, out _), "The shared argument grammar cannot recursively admit arbitrary deeper bodies.");
        }
    }

    public static void MathFontNaryEmptyHiddenLimitsKeepVisibleRolesAndRejectStoredContent()
    {
        var font = OpenTypeFont.Load(FontBytes(0x2211));
        var resource = new PdfFontResource("Hidden", PdfEmbeddedFont.CreateForGlyphMappings(font,
            new[] { ((ushort)128, 0x2211), ((ushort)130, 0x2211), (font.MapCodePoint('B'), (int)'B'), (font.MapCodePoint('C'), (int)'C') }, CancellationToken.None));
        foreach (bool display in new[] { true, false })
        foreach (bool lowerHidden in new[] { true, false })
        foreach (bool upperHidden in new[] { true, false })
        {
            var nary = Nary("\u2211") with { SubscriptHiddenValue = lowerHidden, SuperscriptHiddenValue = upperHidden,
                Subscript = lowerHidden ? new([]) : Row("B"), Superscript = upperHidden ? new([]) : Row("C") };
            var equation = Equation(nary) with { Display = display };
            TestAssert.True(DocxMathLayout.TryGetRuns(equation, CancellationToken.None, out var sources) &&
                sources.Count == 3 - (lowerHidden ? 1 : 0) - (upperHidden ? 1 : 0),
                "Only genuinely empty hidden arguments are excluded from font discovery.");
            TestAssert.True(DocxMathLayout.TryMeasure(equation, _ => resource, CancellationToken.None, out var box, out var reason), reason ?? "Expected empty hidden limits.");
            var measured = TestAssert.NotNull(box);
            TestAssert.Equal("\u2211B" + (upperHidden ? "" : "C") + (lowerHidden ? "" : "B"),
                string.Concat(measured.Glyphs.Select(g => char.ConvertFromUtf32(g.SourceCodePoint))));
            foreach (var stored in new[] { nary with { SubscriptHiddenValue = true, Subscript = Row("B") },
                nary with { SuperscriptHiddenValue = true, Superscript = Row(" ") },
                nary with { SubscriptHiddenValue = true, Subscript = new([nary]) } })
                TestAssert.True(!DocxMathLayout.TryMeasure(equation with { Content = new([stored]) }, _ => resource,
                    CancellationToken.None, out var absent, out _) && absent is null,
                    "Stored nonempty, whitespace or structured hidden content cannot be silently discarded.");
        }
    }

    private static DocxMathNary Nary(string character) => new(Row("B"), Row("B"), Row("C"))
        { CharacterValue = character, LimitLocationValue = "undOvr", ControlText = new("", 12d, "FF0000", false, false, false, null, "TestFont") };
    private static DocxMathRow Row(string value) => new([new DocxMathRun(new(value, 12d, "0000FF", false, false, false, null, "TestFont"), "p", null)]);

    private static byte[] FontBytes(int point, bool invalidScriptConstants = false, bool withFractions = false, bool extendedOperator = false)
    {
        byte[] math = new byte[254];
        void Put(int at, ushort value) => BinaryPrimitives.WriteUInt16BigEndian(math.AsSpan(at), value);
        Put(0, 1); Put(4, 10); Put(10, 73); Put(12, 50); Put(16, 1500);
        foreach (var field in new[] { OpenTypeMathConstant.AxisHeight, OpenTypeMathConstant.UpperLimitGapMin,
            OpenTypeMathConstant.UpperLimitBaselineRiseMin, OpenTypeMathConstant.LowerLimitGapMin, OpenTypeMathConstant.LowerLimitBaselineDropMin })
            Put(18 + (int)field * 4, 250);
        if (invalidScriptConstants) BinaryPrimitives.WriteInt16BigEndian(math.AsSpan(18 + (int)OpenTypeMathConstant.SuperscriptShiftUp * 4), -1);
        if (withFractions)
        {
            foreach (var field in new[] { OpenTypeMathConstant.FractionNumeratorDisplayStyleShiftUp, OpenTypeMathConstant.StackTopDisplayStyleShiftUp }) Put(18 + (int)field * 4, 700);
            foreach (var field in new[] { OpenTypeMathConstant.FractionDenominatorDisplayStyleShiftDown, OpenTypeMathConstant.StackBottomDisplayStyleShiftDown,
                OpenTypeMathConstant.FractionNumeratorShiftUp, OpenTypeMathConstant.StackTopShiftUp }) Put(18 + (int)field * 4, 500);
            foreach (var field in new[] { OpenTypeMathConstant.FractionDenominatorShiftDown, OpenTypeMathConstant.StackBottomShiftDown }) Put(18 + (int)field * 4, 400);
            foreach (var field in new[] { OpenTypeMathConstant.FractionNumDisplayStyleGapMin, OpenTypeMathConstant.FractionDenomDisplayStyleGapMin,
                OpenTypeMathConstant.FractionNumeratorGapMin, OpenTypeMathConstant.FractionDenominatorGapMin, OpenTypeMathConstant.StackGapMin,
                OpenTypeMathConstant.StackDisplayStyleGapMin }) Put(18 + (int)field * 4, 100);
            Put(18 + (int)OpenTypeMathConstant.FractionRuleThickness * 4, 40);
        }
        Put(8, 224); Put(226, 12); Put(230, 1); Put(234, 18); Put(236, 1); Put(238, 1); Put(240, 130);
        Put(244, 2); Put(246, 130); Put(248, 1000); Put(250, 128); Put(252, 2000);
        if (extendedOperator)
        {
            Array.Resize(ref math, 282); Put(6, 254);
            Put(254, 14); Put(258, 8); Put(262, 1); Put(264, 1); Put(266, 130);
            Put(268, 8); Put(270, 1); Put(272, 200); Put(276, 1); Put(278, 1); Put(280, 130);
        }
        byte[] cmap = new byte[52];
        void U16(int at, ushort value) => BinaryPrimitives.WriteUInt16BigEndian(cmap.AsSpan(at), value);
        void U32(int at, uint value) => BinaryPrimitives.WriteUInt32BigEndian(cmap.AsSpan(at), value);
        U16(2, 1); U16(4, 3); U16(6, 10); U32(8, 12); U16(12, 12); U32(16, 40); U32(24, 2);
        U32(28, 0x20); U32(32, 0x9F); U32(36, 1); U32(40, (uint)point); U32(44, (uint)point); U32(48, 130);
        return TestFontBuilder.ReplaceTable(TestFontBuilder.CreateMathFont(math), "cmap", cmap);
    }
}
